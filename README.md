## Testy Excel Importer (Go)

Импортирует мануальные тест-кейсы из Excel (`.xlsx`) в TMS Testy через API:
- создаёт **child suites** под родительским suite `485` по “разделителям” в Excel
- создаёт тест-кейсы в соответствующих child suite
- после каждого кейса спрашивает подтверждение `yes/no`

### Требования

- Go **1.20+**
- Excel-файл в текущей директории запуска

### Установка и запуск

```bash
go mod tidy
go run . -file table-utmanualtc.xlsx
```

Опционально:
- `-sheet "ИмяЛиста"` — если нужно указать лист явно
- `-host https://tms.transtelematica.ru` — если host отличается







