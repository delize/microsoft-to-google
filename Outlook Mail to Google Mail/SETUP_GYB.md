# Setting Up GYB (Got Your Back)

GYB is a command-line tool for backing up and restoring Gmail. We use it to upload converted email to Gmail.

## Step 1: Install GYB

### macOS / Linux

```bash
pip install got-your-back
```

Or install from source:
```bash
git clone https://github.com/GAM-team/got-your-back.git
cd got-your-back
pip install -r requirements.txt
```

### Windows

```powershell
pip install got-your-back
```

Or download the standalone executable from [GYB Releases](https://github.com/GAM-team/got-your-back/releases).

### Verify Installation

```bash
gyb --version
```

You should see output like: `Got Your Back 1.x.x`

## Step 2: Create Google Cloud Project

GYB requires OAuth credentials to access Gmail. You'll need to create a Google Cloud project.

### 2.1 Go to Google Cloud Console

1. Open [Google Cloud Console](https://console.cloud.google.com/)
2. Sign in with your Google account

### 2.2 Create a New Project

1. Click the project dropdown at the top of the page
2. Click **"New Project"**
3. Enter a project name (e.g., "Gmail Migration")
4. Click **"Create"**
5. Wait for the project to be created, then select it

### 2.3 Enable the Gmail API

1. Go to **APIs & Services** > **Library**
2. Search for **"Gmail API"**
3. Click on **Gmail API**
4. Click **"Enable"**

## Step 3: Configure OAuth Consent Screen

Before creating credentials, you must configure the OAuth consent screen.

### 3.1 Go to OAuth Consent Screen

1. Go to **APIs & Services** > **OAuth consent screen**
2. Select **"External"** user type (unless you have Workspace)
3. Click **"Create"**

### 3.2 Fill in App Information

1. **App name:** "Gmail Migration Tool" (or any name)
2. **User support email:** Your email address
3. **Developer contact:** Your email address
4. Click **"Save and Continue"**

### 3.3 Add Scopes

1. Click **"Add or Remove Scopes"**
2. Search for and select:
   - `https://mail.google.com/` (Full Gmail access)
3. Click **"Update"**
4. Click **"Save and Continue"**

### 3.4 Add Test Users

Since the app is in testing mode, you need to add yourself as a test user:

1. Click **"Add Users"**
2. Enter your Gmail address
3. Click **"Add"**
4. Click **"Save and Continue"**

### 3.5 Review and Complete

1. Review your settings
2. Click **"Back to Dashboard"**

## Step 4: Create OAuth Credentials

### 4.1 Go to Credentials

1. Go to **APIs & Services** > **Credentials**
2. Click **"Create Credentials"**
3. Select **"OAuth client ID"**

### 4.2 Configure the Client

1. **Application type:** Desktop app
2. **Name:** "GYB Migration" (or any name)
3. Click **"Create"**

### 4.3 Download Credentials

1. A dialog will show your Client ID and Client Secret
2. Click **"Download JSON"**
3. Save the file as `client_secrets.json`

## Step 5: Configure GYB with Credentials

### 5.1 Place Credentials File

Copy the downloaded `client_secrets.json` to GYB's config directory:

**macOS/Linux:**
```bash
mkdir -p ~/.gyb
cp client_secrets.json ~/.gyb/
```

**Windows:**
```powershell
mkdir $env:USERPROFILE\.gyb -Force
copy client_secrets.json $env:USERPROFILE\.gyb\
```

### 5.2 Authenticate GYB

Run the authentication command:

```bash
gyb --email your.email@gmail.com --action check
```

This will:
1. Open a browser window
2. Ask you to sign in to Google
3. Request permission to access Gmail
4. Save the authentication token for future use

### 5.3 Verify Authentication

```bash
gyb --email your.email@gmail.com --action count
```

This should display a count of messages in your Gmail account, confirming authentication works.

## Troubleshooting

### "Access blocked: This app's request is invalid"

This usually means the OAuth consent screen isn't configured correctly:
1. Go back to **OAuth consent screen**
2. Ensure your email is added as a test user
3. Ensure the Gmail API scope is added

### "client_secrets.json not found"

GYB looks for credentials in this order:
1. Current directory
2. `~/.gyb/client_secrets.json`
3. GYB installation directory

Make sure the file is in one of these locations.

### "Token has been expired or revoked"

Re-authenticate:
```bash
gyb --email your.email@gmail.com --action reauth
```

### "Quota exceeded"

Google has daily quotas for Gmail API:
- 250 quota units per user per second
- 1 billion quota units per day

If you hit limits:
1. Wait 24 hours
2. Reduce batch sizes in your migration

### Rate Limiting

GYB handles rate limiting automatically with exponential backoff. If you see rate limit messages, just let it continue - it will retry.

## Security Notes

### What GYB Can Access

With the `https://mail.google.com/` scope, GYB can:
- Read all emails
- Send emails
- Delete emails
- Manage labels

### Where Credentials Are Stored

| File | Location | Purpose |
|------|----------|---------|
| `client_secrets.json` | `~/.gyb/` | OAuth app credentials |
| `*.oauth2.txt` | `~/.gyb/` | Per-account auth tokens |

### Revoking Access

To revoke GYB's access to your Gmail:
1. Go to [Google Account Permissions](https://myaccount.google.com/permissions)
2. Find "Gmail Migration Tool" (or your app name)
3. Click **"Remove Access"**

## Quick Reference

| Command | Purpose |
|---------|---------|
| `gyb --email EMAIL --action check` | Verify authentication |
| `gyb --email EMAIL --action count` | Count messages |
| `gyb --email EMAIL --action restore --local-folder DIR` | Restore EML files |
| `gyb --email EMAIL --action restore-mbox --local-folder DIR` | Restore MBOX files |
| `gyb --email EMAIL --action reauth` | Re-authenticate |

## Next Steps

Once GYB is set up and authenticated:

1. **Analyze your PST:** `python pst_analyzer.py backup.pst`
2. **Test with dry run:** `python pst_to_gmail.py backup.pst --email EMAIL --dry-run`
3. **Run migration:** `python pst_to_gmail.py backup.pst --email EMAIL`
