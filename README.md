# Microsoft to Google Migration Tools

A collection of tools for migrating data from Microsoft/Outlook services to Google Workspace without triggering notifications or disrupting users.

## Available Migration Tools

| Tool | Description |
|------|-------------|
| [Outlook Cal to Google Cal](./Outlook%20Cal%20to%20Google%20Cal/) | Migrate calendar events from Outlook/Exchange ICS exports to Google Calendar without sending invitations to attendees |
| [Outlook Mail to Google Mail](./Outlook%20Mail%20to%20Google%20Mail/) | Migrate email from Outlook PST files, EML, or MBOX archives to Gmail using readpst and GYB |

## Why These Tools?

Standard import methods often have limitations:
- **File size limits** - Google Calendar web import has a 1 MB limit
- **Notification spam** - Default APIs send emails to all attendees for historical events
- **Missing features** - No dry-run, date filtering, or progress tracking

These tools are built specifically for bulk migrations with features like:
- No notification sending to attendees/contacts
- Dry-run mode for testing
- Date range filtering
- Progress tracking and detailed summaries
- Edge case handling for Outlook-specific formats

## Getting Started

Navigate to the specific migration tool directory for setup instructions and usage documentation.

## Contributing

Found an edge case that isn't handled? Have an improvement? Contributions are welcome!

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## License

MIT License - See [LICENSE](LICENSE) for details.
