# Networking Outreach Automation (Mac + university Outlook)

You fill in a row in Excel. Every weekday at 11:00 AM your Mac emails each new person a personalized
message from your university account, through the Mac **Mail** app. If someone hasn't replied within
7 days, they get one follow-up. The sheet updates itself: Status, Date Sent, Delivery, Replied, Follow Up.

## One-time setup (about 10 minutes)

1. **Add your university email to the Mail app.** Open Mail > Settings > Accounts > **+** > **Microsoft Exchange**,
   then choose **Sign In** and log in with your school account the same way you would on the Outlook website.
   (You can keep using Outlook too. Mail only has to be set up; you don't have to use it day to day.)
2. **Get the project onto your Mac.** Download this repo and open **Terminal** in the folder
   (Finder: right-click the folder > Services > New Terminal at Folder). Then:
   ```
   python3 -m pip install --user -r requirements.txt
   python3 outreach.py init           # creates outreach.xlsx (your contact list)
   cp config.example.ini config.ini
   open -e config.ini                 # put your school email + name in, save
   ```
3. Email templates live in one folder per commonality, e.g. `templates/SNHU/initial_email.txt` and
   `templates/SNHU/follow_up_email.txt`. A row whose Commonality is `SNHU` gets the SNHU emails.
   To add another commonality, copy the SNHU folder, rename it (e.g. `Soccer`), and edit the text.
   Placeholders: `[First Name]`, `[Area]`, `[Company]`, `[Role]`, `[Last Name]`.
   Templates so far: `SNHU` (for SNHU alumni) and `General` (everyone else).
   **Put your resume PDF inside `templates/General/`.** Any PDF/Word file in a template folder is attached
   to that template's first email (not the follow-up). The General email says "I've attached my resume",
   so General contacts wait until the PDF is there. Resumes are never uploaded to GitHub.
4. **Test it:**
   ```
   python3 outreach.py check          # macOS asks "Terminal wants to control Mail" -> click OK
   python3 outreach.py test           # sends both emails to YOURSELF with a fake contact
   python3 outreach.py preview        # shows exactly what would go out to your real list (sends nothing)
   ```
5. **Turn on the weekday 11 AM schedule:**
   ```
   bash scheduling/schedule_mac.sh    # if a popup asks to let python3 control Mail, click OK
   ```

## Daily use

- Add people on the **Contacts** tab. First Name, Email, Area and Commonality are required. Commonality must match a template folder (e.g. `SNHU`). Leave Status empty.
- **Close the Excel file before 11 AM** so the script's updates don't get overwritten.
- Your Mac has to be on. If it's asleep or the lid is closed at 11, the run happens as soon as you open it.
- Set Status to `Do Not Contact` to stop emailing someone.
- Logs are in `logs/outreach.log`. The sheet is backed up to `backups/` before every run.
- To turn the schedule off: `launchctl unload ~/Library/LaunchAgents/com.outreach.daily.plist`

## Safety limits (change in config.ini)
- At most 25 emails a day, with a random 30–90 second pause between them, so your school account doesn't get flagged.
- Nothing is sent on weekends, even if you run it by hand (unless you use `--force`).
- If the inbox can't be checked for replies, that day's follow-ups are skipped so nobody who already replied gets a follow-up.
