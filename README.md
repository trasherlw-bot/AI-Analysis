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

## Input file and result columns
- Name the file with the conversation date, e.g. `queue_daily_20261025.xlsx` (YYYYMMDD). **Запрос клиента** then contains only
  that day's interactions, as `Клиент: … / ИИ: …` pairs. Without a date in the name, all messages are used.
- Result columns: **ID переписки**, **Запрос клиента**, **Ответы ИИ**, **Диалог (размечено)**. The old columns
  «Заявки в группе», «Кол-во ответов ИИ», «Статус обработки» are removed; ticket errors are written into «Запрос клиента».
- **Корректно** (`app/hde_rules.py`; phrases and column letters are constants at the top of that file):
  - main sheet, column **H** (only if empty): client asks for cancel/return and (AI reply is one of the two return phrases
    or column **F** = «Отменена часть списка → остальное оператору»); or client asks to change the delivery date and
    column **F** is one of the three date phrases;
  - sheet «Эскалации», column **K** (only if empty): G = «Управляющий офисом», same e-mail domain in E and H, and either
    I ≠ «Ждёт рабочего времени офиса», or I = it and the time in B (UTC+3) is outside 09:00–18:00 in the region from D.
    Regions missing from the timezone table are listed in the log and skipped.
- Tests: `python -m unittest discover -s tests`

## API limit
HelpDeskEddy allows 300 requests/min per account and blocks the account for 20 minutes when exceeded. The script keeps a
sliding 60-second window at 90% of the limit (`HDE_RATE_LIMIT_RPM`, `HDE_RATE_SAFETY`), slows down when the response
headers report an almost empty quota, and on HTTP 429 stops at once (no retries) and keeps the tickets already downloaded.
Run it again after about 20 minutes.

## Limits and safety
- One job at a time. The container is limited to 0.5 CPU and 400 MB RAM so it cannot starve the dashboard or tracker.
- Upload limit 25 MB, `.xlsx` only.
- The port (3003) is bound to 127.0.0.1; only nginx can reach it.
- Enable HTTPS before using the page from outside your home network, otherwise the password is sent unencrypted.
