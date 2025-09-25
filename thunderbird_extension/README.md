# Email Automation Thunderbird Bridge (Headless)

This folder contains a Thunderbird MailExtension and a Native Messaging host that enable fully headless sending (no SMTP credentials required). Your desktop app writes a job JSON into a queue directory. The native host forwards it to Thunderbird via native messaging, the extension composes/sends the email, then a result JSON is written back. Your app only moves the file to "sent" after confirmed success.

Prerequisites
- Thunderbird 91+ (recommended: Thunderbird 115 or later)
- Python 3 installed and available on PATH
- Windows: No admin rights required (installs manifests into HKCU)

Install Steps
1) Install native host and generate manifests/XPI
   - Run:
     - Windows/macOS/Linux:
       python setup_thunderbird.py --install

   This will:
   - Create platform-specific native messaging manifests in your user directory
   - On Windows, register HKCU registry keys for the native host
   - Build an XPI at: thunderbird_extension/email-automation-bridge.xpi

2) Load the MailExtension into Thunderbird
   - Open Thunderbird
   - Menu → Add-ons and Themes → Tools (gear) → Debug Add-ons
   - Click "Load Temporary Add-on"
   - Select thunderbird_extension/manifest.json
     - Alternatively, load the generated XPI: thunderbird_extension/email-automation-bridge.xpi

3) Keep Thunderbird running
   - The extension listens and connects to the native host while Thunderbird is running
   - The desktop app will now queue jobs and wait for results

Queue Directory
- Default locations (app and native host use the same defaults):
  - Windows: %APPDATA%\EmailAutomation\tb_queue
  - macOS:   ~/Library/Application Support/EmailAutomation/tb_queue
  - Linux:   ~/.local/share/email_automation/tb_queue
- Structure:
  - jobs/    where the app drops {jobId}.json
  - results/ where the native host writes {jobId}.json results
  - native_host.log host logging file (for troubleshooting)

Configure the Desktop App
- In your profile JSON (UI: Configuration → Current Profile):
  - Set email_client to "thunderbird"
  - SMTP fields are NOT required for this mode
  - Optional: tb_queue_dir to override the default queue directory
- The app will:
  - Write a job file when a matching file is detected
  - Wait for results/{jobId}.json
  - On success, log to database and move the file to the "sent" folder

End-to-End Test
1) Run the setup step above and load the extension (once)
2) Ensure Thunderbird is open
3) In the app, select email_client = thunderbird (no SMTP needed)
4) Start Monitoring
5) Drop a test file in the monitored folder whose name matches your regex key pattern
6) Observe:
   - jobs/{jobId}.json created
   - results/{jobId}.json appears with {"success": true}
   - App logs success and moves file into sent/

Troubleshooting
- No result file appears:
  - Verify Thunderbird is running and the extension is loaded (Debug Add-ons page)
  - Check thunderbird_extension/native_host.py is running via native messaging (no extra process needed)
  - See queue_dir/native_host.log for errors
  - Re-run python setup_thunderbird.py --install to refresh manifests/registry
- Attachment failed:
  - Ensure the attachment file exists and is accessible at an absolute path
  - Windows paths are converted to file:/// URLs automatically
- Extension ID mismatch:
  - Extension ID is "email-automation@local"
  - The native manifest must allow this ID. The setup script sets allowed_extensions accordingly.

Files
- Extension:
  - manifest.json: MailExtension manifest
  - background.js: Bridge logic; listens to native messages and sends emails
- Native Host:
  - native_host.py: Watches jobs/, forwards them to TB, writes results/
- Setup:
  - ../setup_thunderbird.py: Installs manifests, registry (Windows), and builds XPI

Security Notes
- Native messaging manifests are installed under your user profile only
- No SMTP credentials are stored or used with the Thunderbird mode