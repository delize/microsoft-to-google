# Enabling Google Calendar API

This guide walks you through setting up Google Cloud credentials to use the calendar migration tools.

## Prerequisites

- A Google account (the one you want to import calendars into)
- Access to [Google Cloud Console](https://console.cloud.google.com/)

## Overview

To use the Google Calendar API, you need to:

1. Create a Google Cloud Project
2. Enable the Google Calendar API
3. Configure the OAuth Consent Screen
4. Create OAuth 2.0 Credentials
5. Download the credentials file

This process is free and typically takes 10-15 minutes.

---

## Step 1: Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)

2. Click the project dropdown at the top of the page (it may say "Select a project" or show an existing project name)

   ![Project Dropdown](images/01-project-dropdown.png)

3. Click **"New Project"** in the top right of the modal

   ![New Project Button](images/02-new-project.png)

4. Enter a project name (e.g., "Calendar Migration") and click **"Create"**

   ![Create Project](images/03-create-project.png)

5. Wait for the project to be created (you'll see a notification), then make sure it's selected in the project dropdown

---

## Step 2: Enable the Google Calendar API

1. In the Google Cloud Console, go to **APIs & Services > Library**
   
   Or navigate directly to: https://console.cloud.google.com/apis/library

   ![API Library Navigation](images/04-api-library-nav.png)

2. Search for **"Google Calendar API"**

   ![Search Calendar API](images/05-search-calendar-api.png)

3. Click on **"Google Calendar API"** in the results

4. Click the **"Enable"** button

   ![Enable API](images/06-enable-api.png)

5. Wait for the API to be enabled (you'll be redirected to the API dashboard)

---

## Step 3: Configure OAuth Consent Screen

Before creating credentials, you must configure the OAuth consent screen. This defines what users see when authorizing your app.

1. Go to **APIs & Services > OAuth consent screen**

   Or navigate directly to: https://console.cloud.google.com/apis/credentials/consent

   ![OAuth Consent Navigation](images/07-oauth-consent-nav.png)

2. Select **"External"** as the User Type and click **"Create"**

   > **Note:** "Internal" is only available for Google Workspace accounts. For personal Gmail accounts, you must select "External".

   ![User Type Selection](images/08-user-type.png)

3. Fill in the **App Information**:
   - **App name:** Calendar Migration Tool (or any name you prefer)
   - **User support email:** Select your email from the dropdown
   - **Developer contact information:** Enter your email address

   ![App Information](images/09-app-information.png)

4. Click **"Save and Continue"**

5. On the **Scopes** page, click **"Add or Remove Scopes"**

6. Search for and select the following scope:
   - `https://www.googleapis.com/auth/calendar` (See, edit, share, and permanently delete all the calendars you can access using Google Calendar)

   ![Add Scopes](images/10-add-scopes.png)

7. Click **"Update"** then **"Save and Continue"**

8. On the **Test Users** page, click **"Add Users"**

9. Enter your email address (the Google account you'll use for the migration)

   ![Add Test Users](images/11-test-users.png)

   > **Important:** Since the app is in "Testing" mode, only test users can authorize it. Add the email of the account you want to import calendars into.

10. Click **"Add"** then **"Save and Continue"**

11. Review the summary and click **"Back to Dashboard"**

---

## Step 4: Create OAuth 2.0 Credentials

1. Go to **APIs & Services > Credentials**

   Or navigate directly to: https://console.cloud.google.com/apis/credentials

   ![Credentials Navigation](images/12-credentials-nav.png)

2. Click **"+ Create Credentials"** at the top of the page

3. Select **"OAuth client ID"**

   ![Create OAuth Client](images/13-create-oauth.png)

4. Configure the OAuth client:
   - **Application type:** Desktop app
   - **Name:** Calendar Migration (or any name you prefer)

   ![Configure OAuth Client](images/14-configure-oauth.png)

5. Click **"Create"**

6. A modal will appear with your credentials. Click **"Download JSON"**

   ![Download Credentials](images/15-download-credentials.png)

7. Rename the downloaded file to `credentials.json`

8. Move `credentials.json` to the same directory as the migration scripts

---

## Step 5: First Run Authentication

The first time you run any of the migration scripts, a browser window will open asking you to authorize the application.

1. Run the script:
   ```bash
   python ics_to_google_calendar.py your_calendar.ics --dry-run
   ```

2. A browser window will open. Select your Google account.

3. You may see a warning: **"Google hasn't verified this app"**
   
   ![Unverified App Warning](images/16-unverified-warning.png)

   This is normal for personal projects. Click **"Continue"** (you may need to click "Advanced" first).

4. Review the permissions and click **"Continue"** or **"Allow"**

   ![Grant Permissions](images/17-grant-permissions.png)

5. You'll see a success message. You can close the browser window.

6. A `token.json` file will be created in your script directory. This stores your authorization and you won't need to re-authenticate unless you delete it.

---

## Troubleshooting

### "Access blocked: This app's request is invalid"

This usually means the OAuth consent screen isn't configured correctly. Make sure:
- You've added yourself as a test user
- The Calendar API scope is added
- You're signing in with the email you added as a test user

### "Error 403: access_denied"

The user denied permission or isn't in the test users list. Add your email to the test users in the OAuth consent screen.

### "Error: redirect_uri_mismatch"

Make sure you selected **"Desktop app"** as the application type when creating credentials, not "Web application".

### Token expired or invalid

Delete the `token.json` file and run the script again to re-authenticate:
```bash
rm token.json
python ics_to_google_calendar.py your_calendar.ics --dry-run
```

---

## Security Notes

- **Keep `credentials.json` private** - Don't commit it to public repositories
- **Keep `token.json` private** - It contains your authorization token
- Add both files to your `.gitignore`:
  ```
  credentials.json
  token.json
  ```
- The credentials only have access to Google Calendar, nothing else
- You can revoke access anytime at https://myaccount.google.com/permissions

---

## Publishing Your App (Optional)

If you want to share these tools with others without them creating their own credentials:

1. Go to the OAuth consent screen
2. Click **"Publish App"**
3. Your app will need to go through Google's verification process
4. This requires a privacy policy, terms of service, and possibly a security assessment

For personal use or sharing with a few people, keeping it in "Testing" mode and adding users manually is simpler.

---

## Quick Reference

| Item | Location |
|------|----------|
| Google Cloud Console | https://console.cloud.google.com/ |
| API Library | APIs & Services > Library |
| OAuth Consent Screen | APIs & Services > OAuth consent screen |
| Credentials | APIs & Services > Credentials |
| Manage App Permissions | https://myaccount.google.com/permissions |

---

## Next Steps

Once you have `credentials.json` in your project directory, you're ready to use the migration tools:

```bash
# Analyze your calendar first
python ics_analyzer.py "Your Calendar.ics"

# Validate for edge cases
python ics_validator.py "Your Calendar.ics"

# Dry run to see what would be imported
python ics_to_google_calendar.py "Your Calendar.ics" --dry-run

# Actually import
python ics_to_google_calendar.py "Your Calendar.ics"
```

See the main [README.md](README.md) for full usage instructions.
