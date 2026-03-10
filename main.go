package main

import (
	"bufio"
	"bytes"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	projectID     = 4
	parentSuiteID = 495
	defaultHost   = "https://tms.transtelematica.ru"

	// API endpoints (from swagger in testyapi.json)
	tokenObtainPath = "/api/token/obtain/"
	suitesPath      = "/api/v1/suites/"
	casesPath       = "/api/v1/cases/"
)

type TestCaseData struct {
	Name     string
	Setup    string
	Scenario string
	Expected string
	Section  string // suite delimiter name
}

type tokenObtainRequest struct {
	Username string `json:"username"`
	Password string `json:"password"`
}

type tokenObtainResponse struct {
	Token string `json:"token"`
	// Some deployments may return JWT pair
	Access  string `json:"access"`
	Refresh string `json:"refresh"`
}

type createSuiteRequest struct {
	Name        string `json:"name"`
	Parent      *int   `json:"parent,omitempty"`
	Project     int    `json:"project"`
	Description string `json:"description,omitempty"`
}

type suiteResponse struct {
	ID   int    `json:"id"`
	Name string `json:"name"`
}

type createTestCaseRequest struct {
	Name     string               `json:"name"`
	Project  int                  `json:"project"`
	Suite    int                  `json:"suite"`
	Setup    string               `json:"setup,omitempty"`
	Scenario string               `json:"scenario,omitempty"`
	Expected string               `json:"expected,omitempty"`
	IsSteps  bool                 `json:"is_steps"`
	Steps    []createTestCaseStep `json:"steps"`
}

type createTestCaseStep struct {
	Name     string `json:"name"`
	Scenario string `json:"scenario"`
	Expected string `json:"expected,omitempty"`
	// sort_order is supported by API, but optional
	SortOrder int `json:"sort_order,omitempty"`
}

type testCaseResponse struct {
	ID   int    `json:"id"`
	Name string `json:"name"`
}

func main() {
	fmt.Println("=== Testy Test Case Importer ===")

	var (
		excelFile = flag.String("file", "table-utmanualtc.xlsx", "Имя Excel-файла (.xlsx) в текущей директории")
		sheetName = flag.String("sheet", "", "Имя листа (если пусто — первый лист)")
		host      = flag.String("host", defaultHost, "Host Testy (например https://tms.transtelematica.ru)")
	)
	flag.Parse()

	login, password, err := getAuthCredentials()
	if err != nil {
		fmt.Printf("[ERROR] Не удалось прочитать логин/пароль: %v\n", err)
		os.Exit(1)
	}

	token, scheme, err := getToken(*host, login, password)
	if err != nil {
		fmt.Printf("[ERROR] Токен не получен: %v\n", err)
		os.Exit(1)
	}
	fmt.Println("[OK] Токен получен")

	absPath, err := filepath.Abs(*excelFile)
	if err != nil {
		fmt.Printf("[ERROR] Не удалось получить абсолютный путь к файлу: %v\n", err)
		os.Exit(1)
	}
	if _, err := os.Stat(absPath); err != nil {
		if os.IsNotExist(err) {
			fmt.Printf("[ERROR] Excel-файл не найден: %s\n", absPath)
			os.Exit(1)
		}
		fmt.Printf("[ERROR] Не удалось проверить файл: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("[INFO] Открыт файл: %s\n", absPath)

	testCases, err := readExcelFile(absPath, *sheetName)
	if err != nil {
		fmt.Printf("[ERROR] Ошибка чтения Excel: %v\n", err)
		os.Exit(1)
	}
	if len(testCases) == 0 {
		fmt.Println("[WARN] В Excel не найдено ни одного тест-кейса для импорта.")
		return
	}

	client := &http.Client{Timeout: 30 * time.Second}

	// Cache section name -> created child suite ID
	suiteIDsBySection := map[string]int{}
	currentSuiteID := parentSuiteID
	currentSection := ""

	createdCount := 0

	for i, tc := range testCases {
		// Handle section delimiter: create/reuse child suite under parentSuiteID.
		if tc.Section != "" && tc.Section != currentSection {
			currentSection = tc.Section
			fmt.Printf("[INFO] Найден раздел: %s\n", currentSection)

			if suiteID, ok := suiteIDsBySection[currentSection]; ok {
				currentSuiteID = suiteID
				fmt.Printf("[INFO] Используем существующий child suite ID: %d\n", currentSuiteID)
			} else {
				suiteID, err := createSuite(client, *host, parentSuiteID, currentSection, token, scheme)
				if err != nil {
					fmt.Printf("[ERROR] Не удалось создать child suite '%s': %v\n", currentSection, err)
					os.Exit(1)
				}
				suiteIDsBySection[currentSection] = suiteID
				currentSuiteID = suiteID
				fmt.Printf("[INFO] Создан child suite ID: %d\n", currentSuiteID)
			}
			continue
		}

		// If there was no section yet, we will place cases directly into parent suite.
		// (But in typical file structure, the first section appears right after header row.)
		targetSuiteID := currentSuiteID

		created, err := createTestCase(client, *host, targetSuiteID, tc, token, scheme)
		if err != nil {
			fmt.Printf("[ERROR] Не удалось создать тест-кейс (строка #%d, name=%q): %v\n", i+1, tc.Name, err)
			os.Exit(1)
		}

		fmt.Printf("[INFO] Создан тест-кейс: %q (ID: %d)\n", created.Name, created.ID)

		createdCount++
		// Подтверждение после первого и далее после каждого десятого созданного кейса.
		if createdCount == 1 || createdCount%10 == 0 {
			if !confirmContinuation() {
				fmt.Println("[INFO] Остановлено пользователем.")
				return
			}
			fmt.Println("[INFO] Продолжаем обработку...")
		}
	}

	fmt.Println("[OK] Импорт завершён.")
}

// getAuthCredentials запрашивает у пользователя логин/пароль через stdin.
func getAuthCredentials() (login, password string, err error) {
	in := bufio.NewReader(os.Stdin)
	fmt.Print("Введите логин: ")
	login, err = readLine(in)
	if err != nil {
		return "", "", err
	}
	fmt.Print("Введите пароль: ")
	password, err = readLine(in)
	if err != nil {
		return "", "", err
	}
	if strings.TrimSpace(login) == "" || strings.TrimSpace(password) == "" {
		return "", "", fmt.Errorf("логин/пароль не должны быть пустыми")
	}
	return strings.TrimSpace(login), password, nil
}

// getToken получает токен авторизации.
// По swagger: POST /api/token/obtain/ {username,password} => {token}.
// На некоторых инсталляциях может возвращаться JWT {access,refresh} — это тоже обрабатываем.
func getToken(host, login, password string) (token string, scheme string, err error) {
	client := &http.Client{Timeout: 20 * time.Second}

	reqBody, _ := json.Marshal(tokenObtainRequest{Username: login, Password: password})
	url := strings.TrimRight(host, "/") + tokenObtainPath

	req, err := http.NewRequest(http.MethodPost, url, bytes.NewReader(reqBody))
	if err != nil {
		return "", "", err
	}
	req.Header.Set("Content-Type", "application/json")

	resp, err := client.Do(req)
	if err != nil {
		return "", "", err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return "", "", fmt.Errorf("HTTP %d при получении токена: %s", resp.StatusCode, truncate(string(body), 800))
	}

	var out tokenObtainResponse
	if err := json.Unmarshal(body, &out); err != nil {
		return "", "", fmt.Errorf("не удалось распарсить ответ токена: %v; body=%s", err, truncate(string(body), 800))
	}

	// Prefer explicit token if present.
	if strings.TrimSpace(out.Token) != "" {
		return strings.TrimSpace(out.Token), "Token", nil
	}
	// Fallback to JWT access token.
	if strings.TrimSpace(out.Access) != "" {
		return strings.TrimSpace(out.Access), "Bearer", nil
	}

	return "", "", fmt.Errorf("в ответе нет поля token/access; body=%s", truncate(string(body), 800))
}

// readExcelFile парсит Excel и возвращает последовательность "событий":
// - строки-разделители (Section != "") для создания/переключения child suite
// - тест-кейсы (Name != "") под текущим разделом
func readExcelFile(path, sheetName string) ([]TestCaseData, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer func() { _ = f.Close() }()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("в файле нет листов")
	}
	if strings.TrimSpace(sheetName) == "" {
		sheetName = sheets[0]
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("лист %q пуст", sheetName)
	}

	// Build merged-row -> section name map
	mergedRowName := map[int]string{}
	mergeCells, _ := f.GetMergeCells(sheetName)
	for _, mc := range mergeCells {
		start := mc.GetStartAxis()
		end := mc.GetEndAxis()
		sCol, sRow, err1 := excelize.CellNameToCoordinates(start)
		eCol, eRow, err2 := excelize.CellNameToCoordinates(end)
		if err1 != nil || err2 != nil {
			continue
		}
		// We only care about single-row merged headers spanning multiple columns
		if sRow == eRow && eCol > sCol {
			val, _ := f.GetCellValue(sheetName, start)
			val = strings.TrimSpace(val)
			if val != "" {
				mergedRowName[sRow] = val
			}
		}
	}

	// Find header row by matching required names
	headerRow := -1
	colIndex := map[string]int{} // normalized header -> 1-based col index

	required := []string{"id", "название", "предусловие", "шаги", "ожидаемый результат"}
	for r := 1; r <= min(20, len(rows)); r++ {
		// Check row values for required headers
		headerMap := map[string]int{}
		for c := 1; c <= 20; c++ { // scan first 20 columns just in case
			cell, _ := f.GetCellValue(sheetName, mustCell(c, r))
			n := normalizeHeader(cell)
			if n != "" {
				headerMap[n] = c
			}
		}
		ok := true
		for _, req := range required {
			if _, exists := headerMap[req]; !exists {
				ok = false
				break
			}
		}
		if ok {
			headerRow = r
			colIndex = headerMap
			break
		}
	}
	if headerRow == -1 {
		return nil, fmt.Errorf("не найдена строка заголовков (ожидаются: %s)", strings.Join(required, " | "))
	}

	get := func(colName string, row int) string {
		c := colIndex[colName]
		if c == 0 {
			return ""
		}
		v, _ := f.GetCellValue(sheetName, mustCell(c, row))
		return strings.TrimSpace(v)
	}

	var out []TestCaseData

	currentSection := ""
	lastRow := len(rows)
	for r := headerRow + 1; r <= lastRow; r++ {
		// Section delimiter by merged row
		if sec, ok := mergedRowName[r]; ok {
			currentSection = sec
			out = append(out, TestCaseData{Section: sec})
			continue
		}

		// Read fields
		name := get("название", r)
		setup := get("предусловие", r)
		scenario := get("шаги", r)
		expected := get("ожидаемый результат", r)
		// id := get("id", r) // not used

		// If row looks like a "section" (sometimes it's not captured as merged):
		if name != "" && setup == "" && scenario == "" && expected == "" {
			// Heuristic: treat as delimiter only if ID is empty too (or absent).
			idVal := get("id", r)
			if strings.TrimSpace(idVal) == "" {
				currentSection = name
				out = append(out, TestCaseData{Section: name})
				continue
			}
		}

		// Skip empty rows
		if name == "" && setup == "" && scenario == "" && expected == "" {
			continue
		}

		if name == "" {
			return nil, fmt.Errorf("строка %d: отсутствует 'Название' у тест-кейса", r)
		}

		out = append(out, TestCaseData{
			Name:     name,
			Setup:    setup,
			Scenario: scenario,
			Expected: expected,
			Section:  currentSection,
		})
	}

	return out, nil
}

// createSuite создаёт child suite в проекте под parentID и возвращает его ID.
func createSuite(client *http.Client, host string, parentID int, name string, token string, scheme string) (int, error) {
	parent := parentID
	reqBody, _ := json.Marshal(createSuiteRequest{
		Name:        name,
		Parent:      &parent,
		Project:     projectID,
		Description: "",
	})

	url := strings.TrimRight(host, "/") + suitesPath
	req, err := http.NewRequest(http.MethodPost, url, bytes.NewReader(reqBody))
	if err != nil {
		return 0, err
	}
	req.Header.Set("Content-Type", "application/json")
	addAuth(req, token, scheme)

	resp, err := client.Do(req)
	if err != nil {
		return 0, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return 0, fmt.Errorf("HTTP %d при создании suite: %s", resp.StatusCode, truncate(string(body), 1200))
	}

	var out suiteResponse
	if err := json.Unmarshal(body, &out); err != nil {
		return 0, fmt.Errorf("не удалось распарсить ответ suite: %v; body=%s", err, truncate(string(body), 1200))
	}
	if out.ID == 0 {
		return 0, fmt.Errorf("API не вернул id suite; body=%s", truncate(string(body), 1200))
	}
	return out.ID, nil
}

// createTestCase создаёт тест-кейс в указанном suiteID.
func createTestCase(client *http.Client, host string, suiteID int, tc TestCaseData, token string, scheme string) (testCaseResponse, error) {
	// IMPORTANT: API requires "steps" array, even if we use plain text fields scenario/expected.
	// To keep mapping intact, we still fill top-level fields, and provide one minimal step
	// duplicating the scenario text for API validation.
	stepsText := tc.Scenario
	if strings.TrimSpace(stepsText) == "" {
		stepsText = "(нет шагов)"
	}

	reqObj := createTestCaseRequest{
		Name:     tc.Name,
		Project:  projectID,
		Suite:    suiteID,
		Setup:    tc.Setup,
		Scenario: tc.Scenario,
		Expected: tc.Expected,
		IsSteps:  false,
		Steps: []createTestCaseStep{
			{
				Name:      "1",
				Scenario:  stepsText,
				Expected:  tc.Expected,
				SortOrder: 1,
			},
		},
	}

	reqBody, _ := json.Marshal(reqObj)
	url := strings.TrimRight(host, "/") + casesPath
	req, err := http.NewRequest(http.MethodPost, url, bytes.NewReader(reqBody))
	if err != nil {
		return testCaseResponse{}, err
	}
	req.Header.Set("Content-Type", "application/json")
	addAuth(req, token, scheme)

	resp, err := client.Do(req)
	if err != nil {
		return testCaseResponse{}, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return testCaseResponse{}, fmt.Errorf("HTTP %d при создании тест-кейса: %s", resp.StatusCode, truncate(string(body), 2000))
	}

	var out testCaseResponse
	if err := json.Unmarshal(body, &out); err != nil {
		// Sometimes API returns extra fields; still should unmarshal. If it fails, print body.
		return testCaseResponse{}, fmt.Errorf("не удалось распарсить ответ тест-кейса: %v; body=%s", err, truncate(string(body), 2000))
	}
	if out.ID == 0 {
		// Not fatal, but surprising.
		out.Name = tc.Name
	}
	return out, nil
}

// confirmContinuation спрашивает пользователя, всё ли корректно, после каждого созданного тест-кейса.
func confirmContinuation() bool {
	in := bufio.NewReader(os.Stdin)
	fmt.Print("Тест-кейс создан. Всё корректно? (yes/no): ")
	ans, err := readLine(in)
	if err != nil {
		return false
	}
	ans = strings.ToLower(strings.TrimSpace(ans))
	return ans == "yes"
}

func addAuth(req *http.Request, token string, scheme string) {
	s := strings.TrimSpace(scheme)
	if s == "" {
		s = "Token"
	}
	req.Header.Set("Authorization", s+" "+strings.TrimSpace(token))
}

func readLine(r *bufio.Reader) (string, error) {
	s, err := r.ReadString('\n')
	if err != nil && err != io.EOF {
		return "", err
	}
	return strings.TrimRight(s, "\r\n"), nil
}

func normalizeHeader(s string) string {
	s = strings.ToLower(strings.TrimSpace(s))
	s = strings.ReplaceAll(s, "\u00a0", " ") // non-breaking space
	s = strings.Join(strings.Fields(s), " ")
	return s
}

func mustCell(col, row int) string {
	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		// This should never happen for valid coordinates
		return "A1"
	}
	return cell
}

func truncate(s string, max int) string {
	s = strings.TrimSpace(s)
	if len(s) <= max {
		return s
	}
	return s[:max] + "...(truncated)"
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
