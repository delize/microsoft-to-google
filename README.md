# Microsoft to Google Migration Tools

Tools for migrating calendar data from Microsoft Outlook/Exchange to Google Calendar without sending notifications to attendees.

## The Problem

When you export a calendar from Outlook, you get a single ICS file that can be hundreds of megabytes with thousands of events. Google Calendar's web import has a 1 MB limit, and using the standard Google Calendar API sends email invitations to all attendees—potentially spamming thousands of people about meetings that happened years ago.

## The Solution

These tools use Google's `events().import_()` API, which is specifically designed for calendar migrations and **never sends notifications to attendees**.

## Tools Included

| Tool | Purpose |
|------|---------|
| `ics_analyzer.py` | Analyze ICS file statistics (event counts, date ranges, attendees, timezones) |
| `ics_validator.py` | Validate ICS files and detect edge cases before import |
| `ics_to_google_calendar.py` | Import events to Google Calendar without notifications |

## Quick Start

### 1. Set Up Google Cloud Credentials

Follow the detailed guide: **[ENABLE_GOOGLE_API.md](ENABLE_GOOGLE_API.md)**

This walks you through creating a Google Cloud project, enabling the Calendar API, and downloading your `credentials.json` file.

### 2. Install Dependencies

```bash
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client icalendar pytz
```

### 3. Analyze Your Calendar

```bash
python ics_analyzer.py "Your Calendar.ics"
```

This shows you what you're working with: event counts, date ranges, attendee statistics, and timezone information.

### 4. Validate for Edge Cases

```bash
python ics_validator.py "Your Calendar.ics"
```

This detects potential issues like:
- Invalid organizer emails
- Events with 50+ attendees
- Distribution lists
- Missing timezone mappings
- Duplicate event UIDs

### 5. Test with Dry Run

```bash
python ics_to_google_calendar.py "Your Calendar.ics" --dry-run
```

See exactly what would be imported without actually creating any events.

### 6. Import Your Calendar

```bash
python ics_to_google_calendar.py "Your Calendar.ics"
```

## Usage Examples

### Import with date filtering

Only import events from 2024:

```bash
python ics_to_google_calendar.py "calendar.ics" \
    --start-date 2024-01-01 \
    --end-date 2025-01-01
```

### Test with a small batch

Import just 10 events to verify everything works:

```bash
python ics_to_google_calendar.py "calendar.ics" --limit 10
```

### Import without attendees

If you don't need attendee information:

```bash
python ics_to_google_calendar.py "calendar.ics" --no-attendees
```

### Import to a specific calendar

```bash
python ics_to_google_calendar.py "calendar.ics" \
    --calendar "your-calendar-id@group.calendar.google.com"
```

### Add yourself as attendee

If you get "not organizer/attendee" errors, add yourself to all events:

```bash
python ics_to_google_calendar.py "calendar.ics" \
    --add-self your.email@gmail.com
```

### Combine options

```bash
python ics_to_google_calendar.py "calendar.ics" \
    --dry-run \
    --start-date 2024-01-01 \
    --end-date 2025-01-01 \
    --limit 20 \
    --add-self your.email@gmail.com
```

## Command Line Options

### ics_to_google_calendar.py

| Option | Description |
|--------|-------------|
| `--dry-run` | Preview import without creating events |
| `--calendar ID` | Target calendar ID (default: primary) |
| `--credentials FILE` | Path to Google OAuth credentials file (default: credentials.json) |
| `--no-attendees` | Skip importing attendee information |
| `--no-skip-duplicates` | Import all events even if they already exist |
| `--start-date YYYY-MM-DD` | Only import events on or after this date |
| `--end-date YYYY-MM-DD` | Only import events before this date |
| `--limit N` | Import only the first N events |
| `--add-self EMAIL` | Add this email as attendee to all events (prevents "not organizer/attendee" errors) |
| `--list-calendars` | List available Google Calendars and exit |

### ics_analyzer.py

| Option | Description |
|--------|-------------|
| `--json` | Output analysis as JSON |

### ics_validator.py

| Option | Description |
|--------|-------------|
| `--sample N` | Number of random events to sample (default: 10) |

## Features

### Outlook-Specific Handling

- **Windows Timezone Conversion** - Converts Windows timezone names ("Central Standard Time") to IANA format ("America/Chicago")
- **Invalid Organizer Handling** - Gracefully handles Outlook's `invalid:nomail` placeholder
- **Response Status Preservation** - Maintains accepted/declined/tentative status from Outlook

### Edge Case Handling

- Skips `@resource.calendar.google.com` addresses (conference rooms)
- Handles distribution list email addresses
- Manages events with 100+ attendees
- Handles missing end times (uses duration or defaults to 1 hour)

### Safety Features

- **Dry-run mode** - Test before committing
- **Date filtering** - Import specific time ranges
- **Batch limiting** - Import small batches for testing
- **Progress tracking** - See real-time import progress
- **Detailed summaries** - Know exactly what was imported, skipped, or failed

## Supported Timezones

The importer handles 40+ Windows timezone names, including:

- US timezones (Eastern, Central, Mountain, Pacific)
- European timezones (GMT, CET, EET)
- Asian timezones (China, Japan, India, Singapore)
- Australian timezones (AUS Eastern, Central, Western)
- And many more...

See the `WINDOWS_TO_IANA_TIMEZONE` dictionary in `ics_to_google_calendar.py` for the complete list.

## File Structure

```
microsoft-to-google/
├── README.md                    # This file
├── ENABLE_GOOGLE_API.md         # Google Cloud setup guide
├── ics_analyzer.py              # Calendar analysis tool
├── ics_validator.py             # Validation and edge case detection
├── ics_to_google_calendar.py    # Main import tool
├── credentials.json             # Your Google Cloud credentials (not committed)
└── token.json                   # OAuth token (not committed, auto-generated)
```

## Security

- `credentials.json` and `token.json` contain sensitive data
- Both are listed in `.gitignore` and should never be committed
- Credentials only have access to Google Calendar
- Revoke access anytime at https://myaccount.google.com/permissions

## Why Not Use Existing Tools?

### Google Calendar Web Import
- 1 MB file size limit
- No progress tracking
- Limited error handling

### gcalcli
- General-purpose CLI, not optimized for bulk migrations
- No dry-run mode
- No date filtering
- No pre-import analysis

### gcal-import-ics
- Similar limitations
- No Outlook-specific handling

These tools were built specifically for migrating large Outlook calendars with features like dry-run testing, date filtering, and comprehensive edge case handling.

## Contributing

Found an edge case that isn't handled? Have an improvement? Contributions are welcome!

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## License

MIT License - See [LICENSE](LICENSE) for details.

## Related Blog Post

For the full story behind these tools, see: [Microsoft to Google: A Retirement Migration Story](https://your-blog-url-here)

---

*Built during the 2024 holiday season to help a family member migrate 30 years of calendar data without spamming thousands of former colleagues.*
