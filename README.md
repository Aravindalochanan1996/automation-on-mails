# Outlook attachment downloader (new Outlook)

This automation uses Microsoft Graph, not the legacy Outlook desktop COM API. It downloads matching file attachments from a dedicated Outlook folder (or Inbox) into a local Windows folder and prevents repeat downloads.

## 1. Requirements

- Python 3.10 or newer.
- A Microsoft Outlook.com/Hotmail account, or a Microsoft 365 work/school mailbox hosted in Exchange Online.
- Permission to register an app in Microsoft Entra. For a managed work account, ask IT to register/approve it if app registration is restricted.
- The computer must be on and signed in when Windows Task Scheduler runs the script.

## 2. Create the Microsoft app

1. Go to [Microsoft Entra admin center](https://entra.microsoft.com/), then **Identity > Applications > App registrations > New registration**.
2. Name it `Outlook Attachment Downloader`.
3. Choose **Accounts in this organizational directory only** for a work/school account, or **Accounts in any organizational directory and personal Microsoft accounts** for Outlook.com/Hotmail.
4. Select **Register** and copy **Application (client) ID**.
5. Open **Authentication**, under **Advanced settings** set **Allow public client flows** to **Yes**, then save. Do not create a client secret: this is a local, interactive sign-in script.
6. Open **API permissions > Add a permission > Microsoft Graph > Delegated permissions**. Add `Mail.Read` and `User.Read`. If your organisation requires it, ask IT to select **Grant admin consent**.

`Mail.Read` is limited to the signed-in user's mailbox. Do not grant broad application permissions for this local task.

## 3. Configure and test

In PowerShell, from this folder:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
Copy-Item .env.example .env
notepad .env
py .\download_outlook_attachments.py
```

Put the client ID in `.env`, choose your download folder, and optionally set sender, subject, and allowed-extension filters. The first run prints a device-login code. Open the shown Microsoft URL, enter the code, and sign in with the same account used in new Outlook. Send a test email with an attachment; its file should appear in `DOWNLOAD_FOLDER`.

## 4. New Outlook steps (recommended)

1. In new Outlook, create a folder such as **Automation** under your mailbox.
2. Go to **Settings > Mail > Rules > Add new rule**.
3. Give it a name, add conditions such as sender/subject/has attachment, then set the action to **Move to** the **Automation** folder. Save it.
4. Set `OUTLOOK_FOLDER=Automation` in `.env`. Keep the sender/subject/file filters in `.env` too; this is a useful second safeguard.

The script checks the selected folder. It does not need Outlook to be open.

## 5. Run automatically on Windows

After the manual test succeeds, open **Task Scheduler > Create Task**:

- **General:** select **Run whether user is logged on or not** (or **Run only when user is logged on** if device sign-in/cache access is problematic).
- **Triggers:** create a trigger such as **Daily**, then **Repeat task every 5 minutes** for a duration of **1 day**.
- **Actions > Start a program:** program/script: `C:\\Users\\aravinalochanan\\Projects\\Automation-on-mails\\.venv\\Scripts\\python.exe`; add arguments: `C:\\Users\\aravinalochanan\\Projects\\Automation-on-mails\\download_outlook_attachments.py`; start in: `C:\\Users\\aravinalochanan\\Projects\\Automation-on-mails`.

The first scheduled execution must happen after the one-time manual sign-in. The local token and processed-attachment state are stored inside the download folder; protect that folder and never commit `.env` or the token cache to source control.
