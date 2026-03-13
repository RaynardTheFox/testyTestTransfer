package main

import (
	"bufio"
	"bytes"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	projectID     = 4
	parentSuiteID = 507
	// Имя тест-плана. Обычно совпадает с именем сьюта, в который загружаем тесты (пример: "test").
	defaultPlanName = "test"
	defaultHost     = "https://tms.transtelematica.ru"

	// API endpoints (from swagger in testyapi.json)
	tokenObtainPath = "/api/token/obtain/"
	suitesPath      = "/api/v1/suites/"
	casesPath       = "/api/v1/cases/"
	// Для тест-планов в твоём инстансе используется /api/v1/testplans/
	plansPath    = "/api/v1/testplans/"
	testsPath    = "/api/v1/tests/"
	resultsPath  = "/api/v1/results/"
	statusesPath = "/api/v1/statuses/"
)

type TestCaseData struct {
	Name     string
	Setup    string
	Scenario string
	Expected string
	Status   string // PASS / FAIL / BLOCK (по Excel)
	Section  string // suite delimiter name
}

// TestCaseStatus связывает созданный тест-кейс в Testy с его статусом из Excel.
type TestCaseStatus struct {
	CaseID int
	Status string // PASS / FAIL / BLOCK
	// Name — для логирования и возможной доп. диагностики.
	Name string
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

// Test Plan API structures
// Структура соответствует TestPlanInputV1 из swagger:
// обязательные поля: name, started_at, due_date, project.
type createTestPlanRequest struct {
	Name        string    `json:"name"`
	Project     int       `json:"project"`
	StartedAt   time.Time `json:"started_at"`
	DueDate     time.Time `json:"due_date"`
	Description string    `json:"description,omitempty"`
}

// Минимальный ответ по тест-плану (совместим с TestPlanMinV1, интересует только id и name).
type testPlanResponse struct {
	ID   int    `json:"id"`
	Name string `json:"name"`
}

// Tests API structures (связка план + кейс).
type createTestRequest struct {
	Project int `json:"project"`
	CaseID  int `json:"case"`
	PlanID  int `json:"plan"`
}

type testResponse struct {
	ID int `json:"id"`
}

// Results API structures (результаты прогона тестов).
type createResultRequest struct {
	StatusID int `json:"status"`
	TestID   int `json:"test"`
}

// ResultStatus описывает возможный статус результата (PASSED/FAILED/BLOCKED и т.п.).
type resultStatus struct {
	ID   int    `json:"id"`
	Name string `json:"name"`
}

type addCaseToPlanRequest struct {
	CaseID int `json:"case_id"`
}

type setCaseResultRequest struct {
	Status string `json:"status"`
}

func main() {
	fmt.Println("=== Testy Test Case Importer ===")

	var (
		excelFile = flag.String("file", "table-utmanualtc.xlsx", "Имя Excel-файла (.xlsx) в текущей директории")
		sheetName = flag.String("sheet", "", "Имя листа (если пусто — первый лист)")
		host      = flag.String("host", defaultHost, "Host Testy (например https://tms.transtelematica.ru)")
		planIDFlg = flag.Int("plan", 0, "ID существующего тест-плана, в который добавлять тесты и результаты (если 0 — этап плана пропускается)")
		suiteIDFg = flag.Int("suite", parentSuiteID, "ID родительского suite, под которым будут создаваться разделы/кейсы")
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
	rootSuiteID := *suiteIDFg
	currentSuiteID := rootSuiteID
	currentSection := ""

	createdCount := 0
	passCount := 0
	failCount := 0
	blockCount := 0
	var createdCases []TestCaseStatus

	for i, tc := range testCases {
		// Handle section delimiter: create/reuse child suite under parentSuiteID.
		if tc.Section != "" && tc.Section != currentSection {
			currentSection = tc.Section
			fmt.Printf("[INFO] Найден раздел: %s\n", currentSection)

			if suiteID, ok := suiteIDsBySection[currentSection]; ok {
				currentSuiteID = suiteID
				fmt.Printf("[INFO] Используем существующий child suite ID: %d\n", currentSuiteID)
			} else {
				suiteID, err := createSuite(client, *host, rootSuiteID, currentSection, token, scheme)
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

		// Сначала пробуем найти уже существующий тест-кейс с таким именем в текущем suite.
		existingID, err := findExistingTestCaseID(client, *host, projectID, targetSuiteID, tc.Name, token, scheme)
		var created testCaseResponse
		if err != nil {
			fmt.Printf("[ERROR] Ошибка при поиске существующего тест-кейса (строка #%d, name=%q): %v\n", i+1, tc.Name, err)
			os.Exit(1)
		}
		if existingID > 0 {
			// Кейс уже существует — не создаём новый, просто логируем и используем существующий ID.
			fmt.Printf("[INFO] Тест-кейс уже существует: %q (ID: %d) — пропускаем создание, используем существующий\n", tc.Name, existingID)
			created = testCaseResponse{ID: existingID, Name: tc.Name}
		} else {
			// Кейс не найден — создаём новый.
			created, err = createTestCase(client, *host, targetSuiteID, tc, token, scheme)
			if err != nil {
				fmt.Printf("[ERROR] Не удалось создать тест-кейс (строка #%d, name=%q): %v\n", i+1, tc.Name, err)
				os.Exit(1)
			}
			fmt.Printf("[INFO] Создан тест-кейс: %q (ID: %d) [Статус: %s]\n", created.Name, created.ID, tc.Status)
		}

		// Сохраняем соответствие ID тест-кейса и статуса для последующего создания Test Plan.
		createdCases = append(createdCases, TestCaseStatus{
			CaseID: created.ID,
			Status: tc.Status,
			Name:   created.Name,
		})

		// Считаем статистику по статусам.
		switch tc.Status {
		case "PASS":
			passCount++
		case "FAIL":
			failCount++
		case "BLOCK":
			blockCount++
		}

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

	if createdCount == 0 {
		fmt.Println("[INFO] Тест-кейсы не были созданы.")
		fmt.Println("[OK] Импорт завершён.")
		return
	}

	fmt.Printf("[INFO] Обработка Excel завершена. Создано тест-кейсов: %d (PASS=%d, FAIL=%d, BLOCK=%d)\n",
		createdCount, passCount, failCount, blockCount)

	// Если пользователь не передал ID плана — завершаем на создании кейсов.
	if *planIDFlg == 0 {
		fmt.Println("[INFO] ID тест-плана не передан (-plan), этап добавления тестов в план и проставления результатов пропускается.")
		fmt.Println("[OK] Импорт завершён.")
		return
	}

	planID := *planIDFlg
	fmt.Printf("[INFO] Используем существующий тест-план ID: %d\n", planID)

	// Подтверждение перед формированием тестов и результатов.
	if !confirmPlanCreation() {
		fmt.Println("[INFO] Пользователь отменил формирование тестов и результатов по плану.")
		fmt.Println("[OK] Импорт завершён.")
		return
	}

	// Загружаем возможные статусы результатов и строим мапу name -> id.
	statusIDs, err := fetchResultStatusIDs(client, *host, projectID, token, scheme)
	if err != nil {
		fmt.Printf("[ERROR] Не удалось получить список статусов результатов: %v\n", err)
		os.Exit(1)
	}

	fmt.Println("[INFO] Создаю тесты (Tests) для плана и устанавливаю результаты...")
	for _, cs := range createdCases {
		// 1. Создаём Test (связка план + кейс).
		testObj, err := createTestForCase(client, *host, projectID, planID, cs.CaseID, token, scheme)
		if err != nil {
			fmt.Printf("[ERROR] Не удалось создать Test для кейса %d в плане %d: %v\n", cs.CaseID, planID, err)
			continue
		}

		// 2. Находим ID статуса по имени.
		statusName := mapStatusToResult(cs.Status)
		if statusName == "" {
			fmt.Printf("[WARN] Невалидный статус %q для кейса %d — результат не устанавливается\n", cs.Status, cs.CaseID)
			continue
		}
		statusID, ok := statusIDs[statusName]
		if !ok {
			fmt.Printf("[WARN] В системе не найден статус %q (кейc %d) — результат не устанавливается\n", statusName, cs.CaseID)
			continue
		}

		// 3. Создаём результат для теста.
		if err := createResultForTest(client, *host, statusID, testObj.ID, token, scheme); err != nil {
			fmt.Printf("[ERROR] Не удалось установить результат для теста %d (case=%d, status=%s): %v\n", testObj.ID, cs.CaseID, statusName, err)
			continue
		}

		fmt.Printf("[OK] Результат для кейса %d (test=%d): %s\n", cs.CaseID, testObj.ID, statusName)
	}

	fmt.Printf("[INFO] Итоговая статистика: PASS=%d, FAIL=%d, BLOCK=%d\n", passCount, failCount, blockCount)
	fmt.Println("[OK] Импорт и формирование тест-плана завершены.")
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

	// Обязательные заголовки, включая новую колонку "Статус".
	required := []string{"id", "название", "предусловие", "шаги", "ожидаемый результат", "статус"}
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
	// Начальный статус по умолчанию — PASS.
	lastValidStatus := "PASS"
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
		rawStatus := get("статус", r)
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

		// Парсим статус с учётом наследования и значений по умолчанию.
		status, updatedLast := parseStatus(rawStatus, lastValidStatus)
		lastValidStatus = updatedLast

		out = append(out, TestCaseData{
			Name:     name,
			Setup:    setup,
			Scenario: scenario,
			Expected: expected,
			Status:   status,
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

// confirmPlanCreation запрашивает подтверждение перед созданием тест-плана
// и/или проставлением результатов выполнения.
func confirmPlanCreation() bool {
	in := bufio.NewReader(os.Stdin)
	fmt.Print("Создать тест-план и проставить результаты? (yes/no): ")
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

// parseStatus нормализует значение статуса из Excel (PASS/FAIL/BLOCK),
// поддерживает:
// - регистронезависимый ввод;
// - значение по умолчанию PASS;
// - наследование последнего валидного статуса (особенно BLOCK);
// - игнорирование невалидных значений.
func parseStatus(raw string, lastValid string) (normalized string, newLast string) {
	raw = strings.ToUpper(strings.TrimSpace(raw))

	switch raw {
	case "PASS", "FAIL", "BLOCK":
		// Явно указанный валидный статус — обновляем lastValid.
		return raw, raw
	case "":
		// Пустая ячейка — используем последний валидный или PASS по умолчанию.
		if strings.TrimSpace(lastValid) == "" {
			return "PASS", "PASS"
		}
		return lastValid, lastValid
	default:
		// Невалидное значение — игнорируем, используем последний валидный
		// или PASS по умолчанию.
		if strings.TrimSpace(lastValid) == "" {
			return "PASS", "PASS"
		}
		return lastValid, lastValid
	}
}

// mapStatusToResult конвертирует Excel-статус (PASS/FAIL/BLOCK)
// в статус результата Testy (PASSED/FAILED/BLOCKED).
func mapStatusToResult(status string) string {
	switch strings.ToUpper(strings.TrimSpace(status)) {
	case "PASS":
		return "PASSED"
	case "FAIL":
		return "FAILED"
	case "BLOCK":
		return "BLOCKED"
	default:
		return ""
	}
}

// fetchResultStatusIDs получает список возможных ResultStatus для проекта
// и строит мапу name -> id.
func fetchResultStatusIDs(client *http.Client, host string, projectID int, token string, scheme string) (map[string]int, error) {
	// В некоторых инсталляциях требуется указать project как query-параметр.
	url := fmt.Sprintf("%s%s?project=%d", strings.TrimRight(host, "/"), statusesPath, projectID)
	req, err := http.NewRequest(http.MethodGet, url, nil)
	if err != nil {
		return nil, err
	}
	addAuth(req, token, scheme)

	resp, err := client.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return nil, fmt.Errorf("HTTP %d при получении статусов результатов: %s", resp.StatusCode, truncate(string(body), 1200))
	}

	var statuses []resultStatus
	if err := json.Unmarshal(body, &statuses); err != nil {
		return nil, fmt.Errorf("не удалось распарсить ответ статусов результатов: %v; body=%s", err, truncate(string(body), 1200))
	}

	out := make(map[string]int, len(statuses))
	for _, st := range statuses {
		name := strings.ToUpper(strings.TrimSpace(st.Name))
		if name != "" {
			out[name] = st.ID
		}
	}
	return out, nil
}

// createTestForCase создаёт сущность Test (привязка case + plan).
func createTestForCase(client *http.Client, host string, projectID, planID, caseID int, token string, scheme string) (testResponse, error) {
	// Сначала пробуем найти уже существующий Test для пары (plan, case).
	if existing, err := findExistingTest(client, host, projectID, planID, caseID, token, scheme); err == nil && existing.ID > 0 {
		return existing, nil
	}

	reqBody, _ := json.Marshal(createTestRequest{
		Project: projectID,
		CaseID:  caseID,
		PlanID:  planID,
	})

	url := strings.TrimRight(host, "/") + testsPath
	req, err := http.NewRequest(http.MethodPost, url, bytes.NewReader(reqBody))
	if err != nil {
		return testResponse{}, err
	}
	req.Header.Set("Content-Type", "application/json")
	addAuth(req, token, scheme)

	resp, err := client.Do(req)
	if err != nil {
		return testResponse{}, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return testResponse{}, fmt.Errorf("HTTP %d при создании Test: %s", resp.StatusCode, truncate(string(body), 1200))
	}

	var out testResponse
	if err := json.Unmarshal(body, &out); err != nil {
		return testResponse{}, fmt.Errorf("не удалось распарсить ответ Test: %v; body=%s", err, truncate(string(body), 1200))
	}
	if out.ID == 0 {
		return testResponse{}, fmt.Errorf("API не вернул id Test; body=%s", truncate(string(body), 1200))
	}
	return out, nil
}

// findExistingTestCaseID ищет тест-кейс по имени в рамках проекта и suite.
// Возвращает ID найденного кейса или 0, если не найден.
func findExistingTestCaseID(client *http.Client, host string, projectID, suiteID int, name string, token string, scheme string) (int, error) {
	name = strings.TrimSpace(name)
	if name == "" {
		return 0, nil
	}

	base := strings.TrimRight(host, "/") + casesPath
	values := url.Values{}
	values.Set("project", fmt.Sprintf("%d", projectID))
	values.Set("suite", fmt.Sprintf("%d", suiteID))
	values.Set("name", name)
	fullURL := base + "?" + values.Encode()

	req, err := http.NewRequest(http.MethodGet, fullURL, nil)
	if err != nil {
		return 0, err
	}
	addAuth(req, token, scheme)

	resp, err := client.Do(req)
	if err != nil {
		return 0, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return 0, fmt.Errorf("HTTP %d при поиске кейса: %s", resp.StatusCode, truncate(string(body), 800))
	}

	// Ответ cases-list — пагинация с полем results.
	var list struct {
		Results []struct {
			ID   int    `json:"id"`
			Name string `json:"name"`
		} `json:"results"`
	}
	if err := json.Unmarshal(body, &list); err != nil {
		return 0, fmt.Errorf("не удалось распарсить ответ поиска кейса: %v; body=%s", err, truncate(string(body), 800))
	}
	if len(list.Results) == 0 {
		return 0, nil
	}
	// Берём первый совпавший ID.
	return list.Results[0].ID, nil
}

// findExistingTest ищет Test по проекту, плану и кейсу.
// Возвращает существующий Test или testResponse{ID:0}, если не найден.
func findExistingTest(client *http.Client, host string, projectID, planID, caseID int, token string, scheme string) (testResponse, error) {
	base := strings.TrimRight(host, "/") + testsPath
	values := url.Values{}
	values.Set("project", fmt.Sprintf("%d", projectID))
	values.Set("plan", fmt.Sprintf("%d", planID))
	values.Set("case", fmt.Sprintf("%d", caseID))
	fullURL := base + "?" + values.Encode()

	req, err := http.NewRequest(http.MethodGet, fullURL, nil)
	if err != nil {
		return testResponse{}, err
	}
	addAuth(req, token, scheme)

	resp, err := client.Do(req)
	if err != nil {
		return testResponse{}, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return testResponse{}, fmt.Errorf("HTTP %d при поиске Test: %s", resp.StatusCode, truncate(string(body), 800))
	}

	var list struct {
		Results []struct {
			ID int `json:"id"`
		} `json:"results"`
	}
	if err := json.Unmarshal(body, &list); err != nil {
		return testResponse{}, fmt.Errorf("не удалось распарсить ответ поиска Test: %v; body=%s", err, truncate(string(body), 800))
	}
	if len(list.Results) == 0 {
		return testResponse{}, nil
	}
	return testResponse{ID: list.Results[0].ID}, nil
}

// createResultForTest создаёт результат выполнения теста.
func createResultForTest(client *http.Client, host string, statusID, testID int, token string, scheme string) error {
	reqBody, _ := json.Marshal(createResultRequest{
		StatusID: statusID,
		TestID:   testID,
	})

	url := strings.TrimRight(host, "/") + resultsPath
	req, err := http.NewRequest(http.MethodPost, url, bytes.NewReader(reqBody))
	if err != nil {
		return err
	}
	req.Header.Set("Content-Type", "application/json")
	addAuth(req, token, scheme)

	resp, err := client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return fmt.Errorf("HTTP %d при создании результата: %s", resp.StatusCode, truncate(string(body), 1200))
	}

	return nil
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
