Как протестировать и выполнить задачу
Установка SendGrid:
bash

Свернуть

Перенос

Копировать
pip install sendgrid
Тестирование:
python script.py --mode generate --limit 10 — создайте 10 PDF и проверьте их.
python script.py --mode send --limit 10 — отправьте 10 писем и проверьте логи.
Разделение задачи:
День 1: python script.py --mode generate — создайте все 564 PDF.
День 2: python script.py --mode send --limit 300 — отправьте первые 300.
День 3: python script.py --mode send --limit 264 — отправьте оставшиеся 264.
Полная автоматизация:
python script.py --limit 564 — всё за один раз (но лучше разбить из-за задержек).