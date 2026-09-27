#!/bin/bash
# Installs a macOS launchd job: Mon-Fri at 11:00 AM. If the Mac was asleep at 11:00, it runs on wake.
# Run once from the project folder:   bash scheduling/schedule_mac.sh
set -e
PROJECT="$(cd "$(dirname "$0")/.." && pwd)"
PY="$(command -v python3)"
PLIST="$HOME/Library/LaunchAgents/com.outreach.daily.plist"
{
cat <<XML
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>com.outreach.daily</string>
  <key>ProgramArguments</key><array><string>$PY</string><string>$PROJECT/outreach.py</string><string>run</string></array>
  <key>WorkingDirectory</key><string>$PROJECT</string>
  <key>StandardOutPath</key><string>$PROJECT/logs/launchd.log</string>
  <key>StandardErrorPath</key><string>$PROJECT/logs/launchd.log</string>
  <key>StartCalendarInterval</key><array>
XML
for d in 1 2 3 4 5; do echo "    <dict><key>Weekday</key><integer>$d</integer><key>Hour</key><integer>11</integer><key>Minute</key><integer>0</integer></dict>"; done
cat <<XML
  </array>
</dict></plist>
XML
} > "$PLIST"
mkdir -p "$PROJECT/logs"
launchctl unload "$PLIST" 2>/dev/null || true
launchctl load "$PLIST"

# Run a harmless 'check' once through launchd so macOS shows the
# "python3 wants to control Mail" prompt now (click OK), not silently at 11 AM.
CHECK="$HOME/Library/LaunchAgents/com.outreach.permcheck.plist"
cat > "$CHECK" <<XML
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>com.outreach.permcheck</string>
  <key>ProgramArguments</key><array><string>$PY</string><string>$PROJECT/outreach.py</string><string>check</string></array>
  <key>WorkingDirectory</key><string>$PROJECT</string>
  <key>RunAtLoad</key><true/>
  <key>StandardOutPath</key><string>$PROJECT/logs/permcheck.log</string>
  <key>StandardErrorPath</key><string>$PROJECT/logs/permcheck.log</string>
</dict></plist>
XML
launchctl unload "$CHECK" 2>/dev/null || true
launchctl load "$CHECK"
echo "If a popup asks to let python3 control Mail, click OK. Waiting 20 seconds..."
sleep 20
launchctl unload "$CHECK" 2>/dev/null || true
rm -f "$CHECK"
cat "$PROJECT/logs/permcheck.log" 2>/dev/null | tail -3
echo "Scheduled: runs Mon-Fri at 11:00 AM.  To remove: launchctl unload $PLIST && rm $PLIST"
