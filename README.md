# HDE AI Analysis

Web tool around `hde_ai_analysis.py`. You give it an Excel file with HelpDeskEddy ticket numbers (column **Заявка**).
It downloads the tickets, masks personal data, groups tickets that share one conversation id, and writes the result columns
back into a copy of the workbook. Served at `http://trasherlw.fvds.ru/hde-analysis/`, protected by a password.

## Where things are
| Path | What |
|---|---|
| `/opt/hde-analysis/data/inbox/` | **put your .xlsx files here** (WinSCP) or upload them on the page |
| `/opt/hde-analysis/data/work/` | per-file working copy, result and ticket cache (safe to delete; tickets will be downloaded again) |
| `/opt/hde-analysis/.env` | HDE email and API key, page password. **Secret: keep `chmod 600`, never share** |
| `app/hde_ai_analysis.py` | your script. Only 3 lines changed (read `HDE_DOMAIN`, `DATA_DIR`, `HDE_BASE_URL` from the environment) |
| `nginx-hde-analysis.conf` | nginx rules for `/hde-analysis/` |
| `hub-index.html` | project hub page with the third card |

## Everyday commands (run in /opt/hde-analysis)
    sudo docker compose ps
    sudo docker compose logs --tail 100 hde-analysis
    sudo docker compose up -d --build        # after changing code or INSTALL_NATASHA
    sudo docker compose up -d                # after editing .env

## Changing the HDE API key later
    sudo nano /opt/hde-analysis/.env         # edit HDE_API_KEY
    cd /opt/hde-analysis && sudo docker compose up -d

## Personal names (natasha)
Without `natasha` the script masks emails, phones, links and long numbers, but **not personal names**.
To also mask names set `INSTALL_NATASHA=true` in `.env` and run `sudo docker compose up -d --build`.
It needs roughly 300 MB extra RAM while a job runs; the container is limited to 400 MB, so on a 1 GB server
check the load first (`free -h`) and stop other heavy jobs.

## Two kinds of input file
**1. Ready daily report** (sheet «Обращения» with the columns «Итог» and «Запрос клиента»; normally also «Сводка», «Эскалации»).
Nothing is downloaded from HelpDeskEddy, no HDE key is needed. The script only writes **Корректно** into the column
**«Комментарий КСЭ»** (found by its header, not by letter) and never overwrites a cell that already has your text.
Sheet «Обращения», any one of:
- **Подтверждение отмены:** «Итог» = «Запрос подтверждения отмены → ответ клиенту».
- **Отмена:** the client text (**Запрос клиента**) or the subject (**Тема**) contains a cancel/refusal word (отмена, отказался,
  возврат, вернуть, аннулировать, не актуален, передумал …) **and** «Итог» = «Возврат оформлен + ответ клиенту» (or
  «Отменена часть списка → остальное оператору», or the note contains one of the two old "return registered" phrases).
- **Перенос даты:** the client text or the subject asks to change/move the delivery date **and** «Итог» is one of
  «Дата доставки изменена + ответ клиенту», «Дата уже стоит в УПД → ответ клиенту по шаблону»,
  «Дата недоступна → клиенту отправлен список вариантов».

Sheet «Эскалации» («Комментарий КСЭ»): «Кого подключили» = «Управляющий офисом», the domain of the address in «Адрес» equals the domain of
the **«Кому:»** address in «Изначальные адресаты» (addresses after «Копия:» are ignored), and either «Итог» ≠ «Ждёт рабочего времени офиса»,
or it is equal and the time in «Время» (UTC+3) is outside 09:00–18:00 in the region of «Офис». Regions missing from the timezone table
(`REGION_UTC`) are listed in the log and skipped. Other roles («Региональный управляющий», «ОКОВРС») are not checked automatically.

At the end the log lists what is left for manual review, grouped by «Итог».
The phrases and keywords are constants at the top of `app/hde_rules.py`.

**2. Plain list of ticket numbers** (column «Заявка», no ready report): the tickets are downloaded from HDE (rate limit below) and
the file gets the columns «ID переписки», «Запрос клиента», «Ответы ИИ», «Диалог (размечено)». If the file name contains a date
(`queue_daily_20261025.xlsx`), «Запрос клиента» holds only that day's `Клиент: … / ИИ: …` pairs.

Tests: `python -m unittest discover -s tests`

## API limit (only when tickets are downloaded)
HelpDeskEddy allows 300 requests/min per account and blocks the account for 20 minutes when exceeded. The script keeps a
sliding 60-second window at 90% of the limit (`HDE_RATE_LIMIT_RPM`, `HDE_RATE_SAFETY`), slows down when the response
headers report an almost empty quota, and on HTTP 429 stops at once (no retries) and keeps the tickets already downloaded.
Run it again after about 20 minutes.

## Limits and safety
- One job at a time. The container is limited to 0.5 CPU and 400 MB RAM so it cannot starve the dashboard or tracker.
- Upload limit 25 MB, `.xlsx` only.
- The port (3003) is bound to 127.0.0.1; only nginx can reach it.
- Enable HTTPS before using the page from outside your home network, otherwise the password is sent unencrypted.
