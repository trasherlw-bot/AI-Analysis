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

## Limits and safety
- One job at a time. The container is limited to 0.5 CPU and 400 MB RAM so it cannot starve the dashboard or tracker.
- Upload limit 25 MB, `.xlsx` only.
- The port (3003) is bound to 127.0.0.1; only nginx can reach it.
- Enable HTTPS before using the page from outside your home network, otherwise the password is sent unencrypted.
