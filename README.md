# Генерация и рассылка сертификатов

Этот проект позволяет автоматически генерировать сертификаты в формате PDF на основе шаблона PowerPoint (`.pptx`) и отправлять их участникам по электронной почте. Поддерживаются два сервиса отправки: Gmail и SendGrid. Код также включает логирование и возможность гибкого управления процессом (генерация, отправка, ограничение количества сертификатов).

## Возможности
- Генерация персонализированных сертификатов из шаблона `.pptx` с подгонкой размера текста.
- Конвертация сертификатов в PDF с использованием Microsoft PowerPoint.
- Отправка сертификатов через Gmail или SendGrid.
- Ограничение количества обрабатываемых сертификатов.
- Разделение процесса на генерацию и отправку.
- Логирование всех операций в файл `cert_log.txt`.
- Предотвращение повторной отправки уже отправленных сертификатов.

## Требования
- **Операционная система**: Windows (для конвертации через PowerPoint).
- **Программное обеспечение**: Microsoft PowerPoint (для конвертации `.pptx` в `.pdf`).
- **Python**: 3.6 или выше.

### Зависимости
Установите необходимые библиотеки с помощью `pip`:
```bash
pip install pandas python-pptx yagmail comtypes sendgrid
```

## Установка
1. Склонируйте репозиторий или скачайте файлы проекта.
2. Установите зависимости (см. выше).
3. Подготовьте файлы:
   - Поместите список участников в `data/test_list.xlsx` с колонками `ФИО` и `Email`.
   - Поместите шаблон сертификата в `data/sert.pptx` с placeholder-ом `{ФИО}`.
4. Настройте почтовые сервисы:
   - **Gmail**: Убедитесь, что у вас есть пароль приложения (вставлен в `GMAIL_PASSWORD`).
   - **SendGrid**: Зарегистрируйтесь на [SendGrid](https://sendgrid.com/), получите API-ключ и вставьте его в `SENDGRID_API_KEY`.

## Использование
Скрипт запускается через командную строку с использованием аргументов.

### Основные команды
1. **Создание сертификатов**:
   ```bash
   python main.py --mode generate --limit 10
   ```
   - Создаёт 10 PDF-сертификатов в папке `certificates`.

2. **Отправка сертификатов через Gmail**:
   ```bash
   python main.py --mode send --limit 5 --email_service gmail --start_index 0
   ```
   - Отправляет 5 сертификатов через Gmail, начиная с первой записи.

3. **Отправка через SendGrid**:
   ```bash
   python main.py --mode send --limit 5 --email_service sendgrid --start_index 0
   ```
   - Отправляет 5 сертификатов через SendGrid.

4. **Создание и отправка разом**:
   ```bash
   python main.py --mode all --limit 50 --email_service gmail
   ```
   - Создаёт и отправляет 50 сертификатов через Gmail.

### Аргументы
- `--limit <число>`: Ограничивает количество обрабатываемых сертификатов (по умолчанию — все).
- `--mode [all|generate|send]`: Режим работы:
  - `all`: Создание и отправка.
  - `generate`: Только создание PDF.
  - `send`: Только отправка.
- `--email_service [gmail|sendgrid]`: Выбор сервиса отправки (по умолчанию `gmail`).
- `--start_index <число>`: Индекс строки, с которой начать отправку (по умолчанию 0).

### Пример полного процесса
1. Создать все 564 сертификата:
   ```bash
   python main.py --mode generate
   ```
2. Отправить первые 300 через Gmail:
   ```bash
   python main.py --mode send --limit 300 --start_index 0 --email_service gmail
   ```
3. Отправить оставшиеся 264:
   ```bash
   python main.py --mode send --limit 264 --start_index 300 --email_service gmail
   ```

## Логирование
Все операции записываются в файл `cert_log.txt`. Пример:
```
2025-04-09 12:00:00 - INFO - Сертификат создан: certificates/Сертификат_Иванов_Иван.pdf
2025-04-09 12:00:05 - INFO - Сертификат отправлен: Иванов Иван (ivanov@example.com)
2025-04-09 12:00:10 - ERROR - Ошибка при отправке письма для Петров Петр: ...
```

## Примечания
- **Gmail лимиты**: 500 писем в сутки. Используйте задержку 5 секунд между отправками (`time.sleep(5)` в коде).
- **SendGrid**: Требуется API-ключ и верификация отправителя. Бесплатный тариф — 100 писем/день, платный — до 40,000/месяц.
- **Повторная отправка**: Скрипт проверяет лог и пропускает уже отправленные сертификаты.
- **Дисковое пространство**: 564 PDF (≈300 КБ каждый) займут около 170 МБ.

## Возможные улучшения
- Поддержка LibreOffice для конвертации на Linux/Mac.
- Уменьшение задержки при использовании SendGrid.
- Интерактивный интерфейс вместо аргументов командной строки.

## Автор
Артем Зенов