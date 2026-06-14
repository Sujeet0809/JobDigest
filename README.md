# 🎯 Product Leadership Job Digest

A daily email digest of **product-leadership roles** (VP / Director / Head /
CPO / GPM / Principal / Senior PM) scraped from 22 job boards, scored for
relevancy to Sujeet's profile (health-tech · BFSI · B2B SaaS · India & Remote),
and delivered as a styled HTML email **with a full Excel attachment**.

It runs **in the cloud via GitHub Actions every day at 7 PM IST** — your laptop
does **not** need to be on.

## How it works

1. `job_scraper.py` fetches roles from 22 sources (LinkedIn, Remotive, Built In,
   RemoteOK, TimesJobs, Naukri, iimjobs, Shine, Indeed India, Foundit, The
   Product Folks, Instahyre, Weekday, Cutshort, and several executive-search
   firms), each with API → embedded-JSON → JSON-LD → HTML fallbacks.
2. Titles are filtered, deduplicated, scored 1–10, and tagged with a country
   flag and work-mode.
3. The results are emailed (HTML + plaintext + `.xlsx`) via Gmail SMTP.
4. `.github/workflows/daily-digest.yml` runs the whole thing on a schedule.

## Secrets (required)

Credentials are read from the environment — **nothing sensitive is committed**.
Add these as repository secrets under
**Settings → Secrets and variables → Actions → New repository secret**:

| Secret name      | Value                                                        |
| ---------------- | ------------------------------------------------------------ |
| `GMAIL_APP_PASS` | 16-char Gmail App Password (https://myaccount.google.com/apppasswords) |
| `SENDER_EMAIL`   | `sujeetkumar0809@gmail.com` (optional — defaults to this)    |
| `RECEIVER_EMAIL` | where to receive the digest (optional — defaults to sender)  |

> A Gmail **App Password** (not your normal password) is required, and your
> Google account must have 2-Step Verification enabled.

## Run it now (manually)

- **In the cloud:** GitHub → **Actions** tab → *Daily Job Digest* → **Run workflow**.
- **Locally:**
  ```bash
  pip install -r requirements.txt
  export GMAIL_APP_PASS="abcd efgh ijkl mnop"
  python job_scraper.py
  ```

## Tests

```bash
pip install -r requirements.txt
python -m pytest tests/ -q
```

The suite covers the network-free logic (config loading, title filtering,
relevancy scoring, country/work-mode detection, deduplication, JSON-LD
parsing). Tests also run as a gate in CI before every email is sent.

## Schedule

Defined by the cron expression in `.github/workflows/daily-digest.yml`:
`30 13 * * *` (13:30 UTC = **19:00 IST**). Edit that line to change the time.
