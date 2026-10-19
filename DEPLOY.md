# Deploying the web app

Two ways to run it: **locally** (`start_web.bat`, http://127.0.0.1:5000) or **on Vercel** (below).
The old Streamlit version (`streamlit_app.py`, `start_app.bat`) still works locally but cannot run on Vercel.

## Vercel
This app handles bank statements and sales, so keep the GitHub repo **private** and **always set APP_PASSWORD**.
Without it, a Vercel deployment refuses to run (it shows a "locked" message).

1. Push this folder to a private GitHub repo.
2. vercel.com -> Add New -> Project -> import the repo. Framework preset: **Flask** (auto-detected). No build command.
3. Settings -> Environment Variables (Production, Preview):
   - `APP_PASSWORD` = a strong password (required)
   - `SECRET_KEY` = any long random string (recommended; keeps logins valid across restarts)
4. Deploy. Open the URL, log in with the password.

Or with the CLI: `npx vercel login`, then `npx vercel` and later `npx vercel --prod`.

### Limits on Vercel
- Total upload per run is about 4.5 MB (bank PDFs + sales Excel + CSVs). Typical daily files are far below this.
- A run must finish within `maxDuration` (60 s, set in `vercel.json`). Typical runs take a few seconds.
- Nothing is stored: files are processed in a temp folder and discarded; the Excel report is returned in the response.
- If you need bigger uploads or longer runs, run it locally or on a container host (Dockerfile is for the Streamlit version).

## Local
```
start_web.bat        # Windows: installs Flask etc. and opens http://127.0.0.1:5000
python app.py        # any OS
```
Set `APP_PASSWORD` first if other people can reach the PC: `set APP_PASSWORD=yourpassword` (Windows cmd).
