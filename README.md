# 📋 Work Reporter

**Work Reporter** is a secure, single-user Python (Flask) web application designed to eliminate the friction of daily logging and recurring status reports. 

You write what you did each day in plain English notes. The app:
1. **Organizes notes into a structured Excel workbook** using Google Gemini AI (or an intelligent rule-based fallback parser).
2. **Automatically emails the weekly report** (`.xlsx`) to your manager every week, CC'ing you.
3. **Automatically emails your daily timetable screenshot** to your manager every weekday morning.
4. **Protects your workflow**:
   - If no notes were taken during the week, it skips emailing your manager and sends *you* a warning email instead.
   - If you haven't uploaded today's timetable screenshot, it prevents sending a stale image and emails *you* a reminder instead.
5. **Runs 100% on free-tier services**: Local SQLite / Neon PostgreSQL, Render Free Web Service, cron-job.org, Gmail SMTP, and Google Gemini API Free Tier.

---

## 🚀 Features

- **Single-User Security**: Protected by a single `APP_PASSWORD` stored in an environment variable, compared using constant-time cryptography (`secrets.compare_digest`), with session-based authentication and CSRF validation.
- **Daily Notes**:
  - Interactive calendar with date picker (defaults to today in your configured timezone).
  - Clean textarea for free-form English entries (e.g. *"Fixed login bug, attended sprint planning for 1 hr, reviewed 2 PRs"*).
  - Saving overwrites the note for that day; submitting empty text deletes it.
  - Interactive Mon–Fri cards showing recorded statuses, clickable to load directly into the editor.
- **Automated Weekly Excel Report**:
  - Sends notes to **Google Gemini API** (free tier REST call) requesting a structured JSON list: `day`, `task`, `category`, `hours`, `status`. Strict prompt prevents hallucination.
  - **Rule-Based Fallback**: If `GEMINI_API_KEY` is missing or fails, an intelligent regex parser extracts tasks, detects categories (*Meeting*, *Development*, *Review*, *DevOps*, *Documentation*), parses hours (*"1.5 hrs"*, *"30 mins"*), and infers status.
  - Generates a styled `.xlsx` workbook with **OpenPyXL**:
    - **"Weekly Report" sheet**: Styled navy title, headers, itemized tasks, category, hours, wrapped text, and custom column widths.
    - **"Summary" sheet**: Daily task counts and total hours for Monday–Friday, plus a total row.
  - Dispatched to `MANAGER_EMAIL` on `REPORT_DAY` at `REPORT_TIME` (default: Friday 17:00), CC'ing `OWNER_EMAIL`.
- **Semester Timetable, Excel API & Night-Before Schedule Dispatch**:
  - **Excel Table & API Integration**: Connect via Excel spreadsheet (`.xlsx`, `.csv`), web upload, or REST API (`POST /api/timetable/upload`, `GET /api/timetable/slots`). OpenPyXL parses day rows, normalized time slots (e.g. `8:00- 9:00`, `12:00-1:00`), and automatically expands course acronyms using the bottom legend (e.g. `DMW` $\rightarrow$ `Data Mining and Warehousing (DMW)`).
  - **Simple Email Message for Manager**: Sends a concise, direct message focused strictly on **Date**, **Class**, and **Timings** without requiring bulky image attachments. If no classes are scheduled, it sends a clear *"No classes scheduled for tomorrow"* notice.
  - **Dispatched One Day Before (Night)**: Automatically dispatched at `SCHEDULE_TIME` (default: `18:00` / 6:00 PM) the night before (Sunday night for Monday, Monday night for Tuesday, etc.).
  - **Weekend & Manual Entry (Saturday & Sunday)**: Full 7-day support (Mon–Sun). The Human Input area includes a date selector with quick-select buttons (*Tomorrow*, *This Saturday*, *This Sunday*) allowing you to add extra classes, remedial sessions, or announcements in the exact same format (`Date`, `Class`, `Timings`). If classes are scheduled for Saturday or Sunday, the system automatically sends the schedule update to your manager on Friday or Saturday evening!
- **Dynamic Configuration**:
  - Live `.env` reloading ensures updates to `MANAGER_EMAIL`, `OWNER_EMAIL`, or send times reflect immediately without restarting the server.
  - In-app **Settings Editor** lets you view and update configuration values directly from your browser.
- **External Cron Architecture**:
  - Free hosting services (like Render) sleep when idle. Instead of fragile background threads (e.g. APScheduler), scheduling is driven by an external free ping service calling `/cron/tick?key=CRON_SECRET` every 10 minutes.
  - Validates whether a report or schedule is due in your `TIMEZONE` within a 2-hour grace window.
  - Prevents duplicate sends using a database `Sent` audit table.
- **Modern Responsive UI**:
  - Single inline template with CSS variables.
  - **Light & Dark mode** with automatic system theme detection and manual toggle saved in `localStorage`.
  - Action buttons: *"Download preview Excel"*, *"Send weekly report now"*, and *"Send schedule now"*.

---

## 🛠️ Project Structure

```
Automation_Tool/
├── app.py              # Complete Flask application (routes, models, parsers, UI)
├── requirements.txt    # Pinned production dependencies
├── render.yaml         # Render Infrastructure-as-Code deployment blueprint
├── .env.example        # Reference environment variables with detailed comments
├── .gitignore          # Git ignore rules for Python, virtualenv, and secrets
└── README.md           # Documentation and deployment guide
```

---

## 💻 Local Setup & Development

### 1. Prerequisites
- Python 3.11 or higher
- Git

### 2. Clone and Create Virtual Environment
```bash
# Clone the repository
git clone <your-repo-url>
cd Automation_Tool

# Create virtual environment
python -m venv .venv

# Activate virtual environment
# Windows (PowerShell):
.\.venv\Scripts\Activate.ps1
# Windows (CMD):
.\.venv\Scripts\activate.bat
# Linux / macOS:
source .venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure Environment Variables
Copy `.env.example` to `.env`:
```bash
# Windows:
copy .env.example .env
# Linux / macOS:
cp .env.example .env
```

Edit `.env` and set your preferred settings:
- Set `SECRET_KEY` (e.g. `python -c "import secrets; print(secrets.token_hex(32))"`)
- Set `APP_PASSWORD` (password used to log in to the dashboard)
- Set `CRON_SECRET` (secret query key for `/cron/tick`)
- Set `MANAGER_EMAIL` and `OWNER_EMAIL`
- Leave `DATABASE_URL` empty to default to local SQLite (`work_reporter.db`)

### 5. Run the Application
```bash
python app.py
```
Open your browser and navigate to `http://127.0.0.1:5000`. Enter your `APP_PASSWORD` to log in.

---

## ☁️ Free-Tier Deployment Guide

Deploy Work Reporter 100% free using:
- **GitHub** (Code repository)
- **Neon** (Free Serverless PostgreSQL database)
- **Render** (Free Web Service hosting Flask + Gunicorn)
- **Gmail** (Free SMTP email sending via App Password)
- **cron-job.org** (Free recurring cron pinger)

---

### Step 1: Push Code to GitHub
1. Initialize a Git repository and commit your files:
   ```bash
   git init
   git add .
   git commit -m "Initial commit of Work Reporter"
   ```
2. Create a **Private** repository on [GitHub](https://github.com/new).
3. Push your repository to GitHub:
   ```bash
   git remote add origin https://github.com/<your-username>/work-reporter.git
   git branch -M main
   git push -u origin main
   ```

---

### Step 2: Set Up Neon Free PostgreSQL Database
1. Go to [Neon.tech](https://neon.tech) and create a free account.
2. Click **Create Project**, name it `work-reporter`, and select a region closest to you.
3. In the Neon Dashboard, copy the connection string under **Connection Details**:
   ```
   postgresql://alex:password@ep-sample-123456.us-east-2.aws.neon.tech/neondb?sslmode=require
   ```
4. Save this URL for your `DATABASE_URL` environment variable.

---

### Step 3: Generate a Gmail App Password
To send emails securely through Gmail's SMTP server without sharing your main Google account password:
1. Log into your Google Account and ensure **2-Step Verification** is turned ON.
2. Go to **[Google App Passwords](https://myaccount.google.com/apppasswords)**.
3. Under *App name*, type `Work Reporter` and click **Create**.
4. Copy the generated **16-character password** (e.g., `abcd efgh ijkl mnop`).
5. In your settings:
   - `SMTP_HOST`: `smtp.gmail.com`
   - `SMTP_PORT`: `587`
   - `SMTP_USER`: `your.email@gmail.com`
   - `SMTP_PASS`: `abcdefghijklmnop` (without spaces)
   - `SMTP_FROM`: `"Your Name via Work Reporter" <your.email@gmail.com>`

---

### Step 4: Get a Free Google Gemini API Key (Optional)
1. Visit **[Google AI Studio](https://aistudio.google.com/)**.
2. Click **Get API key** and create an API key in a free project.
3. Copy the key for `GEMINI_API_KEY`.
*(Note: If you omit this key, the application automatically uses its built-in rule-based fallback parser).*

---

### Step 5: Deploy to Render
1. Go to **[Render.com](https://render.com)** and create a free account.
2. Click **New +** > **Web Service**.
3. Connect your GitHub account and select your `work-reporter` repository.
4. Fill in the service configuration:
   - **Name**: `work-reporter`
   - **Region**: Oregon (or your preferred region)
   - **Branch**: `main`
   - **Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Instance Type**: `Free`
5. Expand **Environment Variables** and add the following keys:

| Key | Description | Example / Value |
|---|---|---|
| `SECRET_KEY` | Flask session secret | *(Click "Generate" or paste 32-byte hex)* |
| `APP_PASSWORD` | Dashboard login password | `YourSecurePassword123` |
| `CRON_SECRET` | Secret token for external cron tick | `a_very_secret_cron_token_99` |
| `DATABASE_URL` | Neon PostgreSQL Connection URI | `postgresql://user:pass@ep-xyz.aws.neon.tech/neondb?sslmode=require` |
| `TIMEZONE` | IANA Timezone | `Asia/Kolkata` *(or your local timezone)* |
| `MANAGER_EMAIL` | Manager's destination address | `manager@company.com` |
| `OWNER_EMAIL` | Your address (CC, warnings, reminders) | `you@gmail.com` |
| `REPORT_DAY` | Day to send weekly report (0=Mon..4=Fri) | `4` *(Friday)* |
| `REPORT_TIME` | Time to send weekly report (24h) | `17:00` |
| `SCHEDULE_TIME` | Time to send daily schedule (24h) | `09:00` |
| `GEMINI_API_KEY` | Google AI Studio Key (optional) | `AIzaSy...` |
| `GEMINI_MODEL` | Gemini Model identifier | `gemini-1.5-flash` |
| `SMTP_HOST` | SMTP server | `smtp.gmail.com` |
| `SMTP_PORT` | SMTP port | `587` |
| `SMTP_USER` | Gmail address | `you@gmail.com` |
| `SMTP_PASS` | 16-character Gmail App Password | `abcdefghijklmnop` |
| `SMTP_FROM` | From email header | `"Your Name" <you@gmail.com>` |

6. Click **Create Web Service**. Render will build and deploy your app. Once deployed, Render will provide your public URL:
   ```
   https://work-reporter.onrender.com
   ```

---

### Step 6: Set Up Recurring Scheduler on cron-job.org
Because free Render instances spin down after 15 minutes of inactivity, **cron-job.org** is used to ping the app every 10 minutes. This wakes up the service and triggers the scheduling checks.

1. Go to **[cron-job.org](https://cron-job.org)** and sign up for a free account.
2. In the console, click **Create Cronjob**.
3. Configure the job:
   - **Title**: `Work Reporter Tick`
   - **URL**:
     ```
     https://work-reporter.onrender.com/cron/tick?key=YOUR_CRON_SECRET
     ```
     *(Replace `work-reporter.onrender.com` with your actual Render URL and `YOUR_CRON_SECRET` with the exact value set in your Render environment variables).*
   - **Execution Schedule**: **Every 10 minutes** (`*/10 * * * *`).
   - **Request Method**: `GET`
4. Click **Create**.
5. Test the job by clicking **Run now**. In the execution history, you should see HTTP status code `200` with response body:
   ```json
   {"current_time":"...","message":"Nothing due at this time","status":"ok","weekday":"..."}
   ```

---

## 🔍 How the Cron Logic & Safeguards Work

Every 10 minutes, `cron-job.org` calls `/cron/tick?key=CRON_SECRET`:
1. **Authentication**: `secrets.compare_digest` validates the query parameter. If missing or invalid, an HTTP 403 Forbidden is returned.
2. **Timezone Calculation**: The server converts current UTC time to your configured `TIMEZONE` (e.g. `Asia/Kolkata`).
3. **Grace Window (2 Hours)**: An event is considered due if current time is within 2 hours after the scheduled time (e.g., if scheduled at `09:00`, it will trigger between `09:00` and `11:00`).
4. **Duplicate Prevention**: The `Sent` table is queried for today's date and the action kind. If an email has already succeeded today, it will **never be sent twice**.
5. **Smart Fallbacks**:
   - **Daily Schedule**: If today's timetable image was uploaded today, it is attached and sent to `MANAGER_EMAIL` (CC `OWNER_EMAIL`). If no image or an old image is detected, it does **not** send a stale image; instead, it sends a reminder to `OWNER_EMAIL`.
   - **Weekly Report**: If notes were recorded for Monday–Friday, an Excel report is created and emailed to `MANAGER_EMAIL` (CC `OWNER_EMAIL`). If no notes exist, it does **not** send an empty email to your manager; it sends a warning to `OWNER_EMAIL`.

---

## 🛠️ Troubleshooting

### 1. `/cron/tick` returns `403 Forbidden`
- Ensure the query parameter is `?key=YOUR_CRON_SECRET` and matches the `CRON_SECRET` environment variable character-for-character.
- Note that comparison is case-sensitive and constant-time.

### 2. Email Sending Fails (`SMTPAuthenticationError`)
- Verify that **2-Step Verification** is enabled on your Google Account.
- Ensure you are using an **App Password** (16 letters), NOT your standard Gmail account password.
- Check that `SMTP_HOST=smtp.gmail.com` and `SMTP_PORT=587`.

### 3. Database Connection Issues on Neon
- Neon URLs require `sslmode=require`. Ensure your `DATABASE_URL` ends with `?sslmode=require`.
- Work Reporter automatically converts `postgres://` to `postgresql://` to maintain compatibility with SQLAlchemy 2.x.
- Neon databases sleep when inactive; Work Reporter sets `pool_pre_ping: True` to automatically reconnect without errors.

### 4. Excel Download or Generation Errors
- If using Google Gemini, check that `GEMINI_API_KEY` is valid.
- If the Gemini API quota is exceeded or unavailable, Work Reporter automatically falls back to its built-in rule-based parser without failing.

### 5. Timezone Discrepancies
- Ensure `TIMEZONE` is set to a valid IANA timezone string (e.g. `Asia/Kolkata`, `America/New_York`, `Europe/London`).
- The current localized server time is displayed on the dashboard for verification.

---

## 📄 License

MIT License. Designed and crafted for private, single-user automated productivity.
