# Networking Outreach Automation

You fill in a row in Excel. Every weekday at 11:00 AM your computer emails each new person a personalized
message from your university account. If someone hasn't replied after 5 workdays, they get one follow-up,
sent as a reply in the same thread. The sheet updates itself: Status, Date Sent, Delivery, Replied, Follow Up.

## One-time setup (about 10 minutes)

1. Install Python 3 (python.org). On Windows, tick **"Add Python to PATH"**.
2. Open a terminal in this folder and run:
   ```
   pip install -r requirements.txt
   python outreach.py init            # creates outreach.xlsx (your contact list)
   ```
3. Copy `config.example.ini` to `config.ini` and fill in your email settings (see the notes inside it).
4. Paste your drafts into `templates/initial_email.txt` and `templates/follow_up_email.txt`.
   Placeholders: `[First Name]`, `[Area]`, `[Commonality]`, `[Company]`, `[Role]`, `[Last Name]`.
5. Check that it works:
   ```
   python outreach.py check           # can it log in?
   python outreach.py test            # sends both emails to YOURSELF with a fake contact
   python outreach.py preview         # shows exactly what would go out to your real list (sends nothing)
   ```
6. Turn on the daily 11 AM weekday schedule:
   - **Windows:** `powershell -ExecutionPolicy Bypass -File scheduling\schedule_windows.ps1`
   - **Mac:** `bash scheduling/schedule_mac.sh`

## Daily use

- Add people on the **Contacts** tab. First Name, Email, Area and Commonality are required. Leave Status empty.
- **Close the Excel file before 11 AM** so the script can save its updates.
- Your computer has to be on (asleep is OK on Mac, and on Windows it catches up when you wake it).
- Set Status to `Do Not Contact` to stop emailing someone.
- Logs are in `logs/outreach.log`. The sheet is backed up to `backups/` before every run.

## Safety limits (change in config.ini)
- At most 25 emails a day, with a random 30–90 second pause between them, so your school account doesn't get flagged.
- Nothing is sent on weekends, even if you run it by hand (unless you use `--force`).
- If the inbox can't be checked for replies, that day's follow-ups are skipped so nobody who already replied gets a follow-up.
