# Setting Up GYB (Got Your Back)

GYB is a command-line tool for backing up and restoring Gmail. We use it to upload converted email to Gmail.

## Step 1: Install GYB

Run the install script which downloads the latest GYB standalone build:

```bash
zsh scripts/install_deps.sh
```

Or download manually from [GYB Releases](https://github.com/GAM-team/got-your-back/releases).

### Verify Installation

```bash
# If installed via script:
./gyb/gyb --version

# Or if in PATH:
gyb --version
```

You should see output like: `Got Your Back 1.x.x`

## Step 2: Create Google Cloud Project (Automated)

GYB can automatically create a Google Cloud project for you. This is the easiest method.

```bash
./gyb/gyb --action create-project --email your.email@gmail.com
```

This will:

1. **Open a browser** - Sign in with your Google account and authorize GYB to create a project
2. **Create the project** - GYB creates "Got Your Back Project" and enables required APIs
3. **Prompt you to create OAuth credentials** - You'll see a URL like:
   ```
   Please go to:
   https://console.cloud.google.com/apis/credentials/oauthclient?project=gyb-project-xxx-xxx-xxx
   ```

4. **In the browser**, follow these steps:
   - Enter **"GYB"** for "Application name"
   - Leave other fields blank, click **"Save"**
   - Choose **"Desktop app"**
   - Enter any name (e.g., "GYB"), click **"Create"**
   - Copy the **Client ID** and paste it into the terminal
   - Copy the **Client Secret** and paste it into the terminal

5. **Done!** You'll see: `That's it! Your GYB Project is created and ready to use.`

## Step 3: Authenticate with Gmail

Now authorize GYB to access your Gmail account:

```bash
./gyb/gyb --email your.email@gmail.com --action count
```

This will:
1. Open a browser window
2. Ask you to sign in to Google
3. Request permission to access Gmail
4. Save the authentication token for future use
5. Display a count of messages in your Gmail (confirming it works)

## Alternative: Manual Project Setup

If the automated setup doesn't work, you can create the project manually.

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
| `gyb --email EMAIL --action count` | Verify authentication (shows message count) |
| `gyb --email EMAIL --action count` | Count messages |
| `gyb --email EMAIL --action restore --local-folder DIR` | Restore EML files |
| `gyb --email EMAIL --action restore-mbox --local-folder DIR` | Restore MBOX files |
| `gyb --email EMAIL --action reauth` | Re-authenticate |

## Next Steps

Once GYB is set up and authenticated:

1. **Analyze your PST:** `python pst_analyzer.py backup.pst`
2. **Test with dry run:** `python pst_to_gmail.py backup.pst --email EMAIL --dry-run`
3. **Run migration:** `python pst_to_gmail.py backup.pst --email EMAIL`
