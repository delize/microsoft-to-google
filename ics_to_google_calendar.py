#!/usr/bin/env python3
"""
ICS to Google Calendar Importer (Enhanced Version)

This script reads ICS calendar files and imports events to Google Calendar
using the Google Calendar API with OAuth 2.0 authentication.

Features:
    - Imports attendees WITHOUT sending email notifications (uses import API)
    - Preserves organizer information
    - Imports reminders/alarms
    - Handles timezones properly
    - Supports recurrence rules and exceptions
    - Handles all-day and timed events
    - Duplicate detection
    - Dry-run mode for testing
    - Date range filtering

Requirements:
    pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client icalendar python-dateutil pytz

Setup:
    1. Go to https://console.cloud.google.com/
    2. Create a new project (or select existing)
    3. Enable the Google Calendar API:
       - Go to "APIs & Services" > "Library"
       - Search for "Google Calendar API" and enable it
    4. Create OAuth 2.0 credentials:
       - Go to "APIs & Services" > "Credentials"
       - Click "Create Credentials" > "OAuth client ID"
       - Choose "Desktop app" as application type
       - Download the JSON file and save as "credentials.json" in the same directory as this script
    5. Run the script - it will open a browser for authentication on first run

Usage:
    python ics_to_google_calendar.py <ics_file_or_directory> [options]

Examples:
    python ics_to_google_calendar.py calendar.ics
    python ics_to_google_calendar.py ./ics_files/
    python ics_to_google_calendar.py calendar.ics --calendar "work@group.calendar.google.com"
    python ics_to_google_calendar.py calendar.ics --no-attendees
    python ics_to_google_calendar.py calendar.ics --dry-run
    python ics_to_google_calendar.py calendar.ics --dry-run --start-date 2024-01-01 --end-date 2024-12-31
    python ics_to_google_calendar.py calendar.ics --dry-run --limit 10
"""

import os
import sys
import argparse
import time
import pickle
import re
import hashlib
from pathlib import Path
from datetime import datetime, timedelta, date
from typing import Optional, List, Dict, Any, Set, Tuple
from collections import defaultdict

# Google API imports
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ICS parsing
from icalendar import Calendar, Event, vRecur
from dateutil import tz as dateutil_tz
from dateutil.rrule import rrulestr

try:
    import pytz
    HAS_PYTZ = True
except ImportError:
    HAS_PYTZ = False

# OAuth scopes
SCOPES = ['https://www.googleapis.com/auth/calendar']

# Rate limiting settings
REQUESTS_PER_SECOND = 5
BATCH_SIZE = 50

# Common timezone mappings (Windows/Outlook to IANA)
TIMEZONE_MAPPINGS = {
    'Eastern Standard Time': 'America/New_York',
    'Eastern Daylight Time': 'America/New_York',
    'Central Standard Time': 'America/Chicago',
    'Central Daylight Time': 'America/Chicago',
    'Mountain Standard Time': 'America/Denver',
    'Mountain Daylight Time': 'America/Denver',
    'US Mountain Standard Time': 'America/Phoenix',
    'Pacific Standard Time': 'America/Los_Angeles',
    'Pacific Daylight Time': 'America/Los_Angeles',
    'Alaska Standard Time': 'America/Anchorage',
    'Hawaiian Standard Time': 'Pacific/Honolulu',
    'GMT Standard Time': 'Europe/London',
    'W. Europe Standard Time': 'Europe/Berlin',
    'Romance Standard Time': 'Europe/Paris',
    'Central European Standard Time': 'Europe/Budapest',
    'E. Europe Standard Time': 'Europe/Bucharest',
    'FLE Standard Time': 'Europe/Kiev',
    'Russian Standard Time': 'Europe/Moscow',
    'Tokyo Standard Time': 'Asia/Tokyo',
    'China Standard Time': 'Asia/Shanghai',
    'Singapore Standard Time': 'Asia/Singapore',
    'India Standard Time': 'Asia/Kolkata',
    'Arabian Standard Time': 'Asia/Dubai',
    'Israel Standard Time': 'Asia/Jerusalem',
    'AUS Eastern Standard Time': 'Australia/Sydney',
    'E. Australia Standard Time': 'Australia/Brisbane',
    'AUS Central Standard Time': 'Australia/Darwin',
    'Cen. Australia Standard Time': 'Australia/Adelaide',
    'W. Australia Standard Time': 'Australia/Perth',
    'New Zealand Standard Time': 'Pacific/Auckland',
    'Venezuela Standard Time': 'America/Caracas',
    'SA Pacific Standard Time': 'America/Bogota',
    'Atlantic Standard Time': 'America/Halifax',
    'UTC': 'UTC',
    'Coordinated Universal Time': 'UTC',
}


def normalize_timezone(tz_str: str) -> str:
    """Convert Windows/Outlook timezone names to IANA timezone names"""
    if not tz_str:
        return 'UTC'
    
    # Check if it's already an IANA timezone
    if HAS_PYTZ:
        try:
            pytz.timezone(tz_str)
            return tz_str
        except pytz.UnknownTimeZoneError:
            pass
    
    # Try mapping
    if tz_str in TIMEZONE_MAPPINGS:
        return TIMEZONE_MAPPINGS[tz_str]
    
    # Try partial matches
    for windows_tz, iana_tz in TIMEZONE_MAPPINGS.items():
        if windows_tz.lower() in tz_str.lower() or tz_str.lower() in windows_tz.lower():
            return iana_tz
    
    # Default to UTC if unknown
    return 'UTC'


def filter_events_by_date(events: List[Dict[str, Any]], 
                          start_date: Optional[date] = None, 
                          end_date: Optional[date] = None) -> List[Dict[str, Any]]:
    """Filter events by date range based on their start time.
    
    Args:
        events: List of Google Calendar event dictionaries
        start_date: Only include events starting on or after this date
        end_date: Only include events starting before this date
    
    Returns:
        Filtered list of events
    """
    if not start_date and not end_date:
        return events
    
    filtered = []
    for event in events:
        start = event.get('start', {})
        
        # Get the start date/datetime
        start_str = start.get('dateTime') or start.get('date')
        if not start_str:
            continue
        
        try:
            # Parse the start date
            if 'T' in start_str:
                # DateTime format: 2024-01-15T10:00:00-06:00
                event_date = datetime.fromisoformat(start_str.replace('Z', '+00:00')).date()
            else:
                # Date format: 2024-01-15
                event_date = datetime.strptime(start_str, '%Y-%m-%d').date()
            
            # Apply filters
            if start_date and event_date < start_date:
                continue
            if end_date and event_date >= end_date:
                continue
            
            filtered.append(event)
            
        except (ValueError, TypeError):
            # If we can't parse the date, include the event
            filtered.append(event)
    
    return filtered


class GoogleCalendarImporter:
    def __init__(self, credentials_file: str = 'credentials.json', token_file: str = 'token.pickle'):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.service = None
        self.default_timezone = 'UTC'
        self.dry_run = False  # Track if we're in dry-run mode
        self.stats = {
            'total_events': 0,
            'imported': 0,
            'skipped': 0,
            'errors': 0,
            'attendees_imported': 0
        }
    
    def authenticate(self) -> None:
        """Authenticate with Google using OAuth 2.0"""
        creds = None
        
        # Load existing token if available
        if os.path.exists(self.token_file):
            with open(self.token_file, 'rb') as token:
                creds = pickle.load(token)
        
        # If no valid credentials, initiate OAuth flow
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                print("Refreshing expired credentials...")
                creds.refresh(Request())
            else:
                if not os.path.exists(self.credentials_file):
                    print(f"\nError: '{self.credentials_file}' not found!")
                    print("\nTo set up Google Calendar API credentials:")
                    print("1. Go to https://console.cloud.google.com/")
                    print("2. Create a new project or select an existing one")
                    print("3. Enable the Google Calendar API:")
                    print("   - Go to 'APIs & Services' > 'Library'")
                    print("   - Search for 'Google Calendar API' and enable it")
                    print("4. Create OAuth 2.0 credentials:")
                    print("   - Go to 'APIs & Services' > 'Credentials'")
                    print("   - Click 'Create Credentials' > 'OAuth client ID'")
                    print("   - Choose 'Desktop app' as application type")
                    print("   - Download the JSON file")
                    print(f"5. Save the downloaded file as '{self.credentials_file}' in this directory")
                    sys.exit(1)
                
                print("\nOpening browser for Google authentication...")
                print("(If browser doesn't open, check the URL in the terminal)\n")
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, SCOPES
                )
                creds = flow.run_local_server(port=0)
            
            # Save credentials for future runs
            with open(self.token_file, 'wb') as token:
                pickle.dump(creds, token)
            print("Authentication successful! Credentials saved.\n")
            sys.stdout.flush()
        
        # Build the Calendar API service
        print("Building Google Calendar API service...")
        sys.stdout.flush()
        self.service = build('calendar', 'v3', credentials=creds)
        
        # Get default timezone from primary calendar
        print("Getting calendar timezone...")
        sys.stdout.flush()
        try:
            calendar = self.service.calendars().get(calendarId='primary').execute()
            self.default_timezone = calendar.get('timeZone', 'UTC')
        except Exception:
            self.default_timezone = 'UTC'
    
    def list_calendars(self) -> List[Dict[str, str]]:
        """List all calendars accessible by the authenticated user"""
        calendars = []
        page_token = None
        
        while True:
            calendar_list = self.service.calendarList().list(pageToken=page_token).execute()
            for calendar in calendar_list['items']:
                calendars.append({
                    'id': calendar['id'],
                    'summary': calendar.get('summary', 'Untitled'),
                    'primary': calendar.get('primary', False)
                })
            page_token = calendar_list.get('nextPageToken')
            if not page_token:
                break
        
        return calendars
    
    def parse_ics_file(self, ics_path: str, include_attendees: bool = True) -> Tuple[List[Dict[str, Any]], str]:
        """Parse an ICS file and extract events. Returns (events, calendar_timezone)"""
        events = []
        calendar_tz = self.default_timezone
        
        print(f"\nParsing: {ics_path}")
        file_size = os.path.getsize(ics_path)
        print(f"File size: {file_size / 1024 / 1024:.2f} MB")
        print("Reading file (this may take a moment for large files)...")
        sys.stdout.flush()
        
        with open(ics_path, 'rb') as f:
            try:
                raw_data = f.read()
                print(f"Read {len(raw_data) / 1024 / 1024:.2f} MB, parsing ICS format...")
                sys.stdout.flush()
                cal = Calendar.from_ical(raw_data)
                print("ICS parsed successfully")
                sys.stdout.flush()
            except Exception as e:
                print(f"Error parsing ICS file: {e}")
                return events, calendar_tz
        
        # Extract calendar-level timezone
        for component in cal.walk():
            if component.name == "VCALENDAR":
                x_wr_timezone = component.get('x-wr-timezone')
                if x_wr_timezone:
                    calendar_tz = normalize_timezone(str(x_wr_timezone))
                    break
            elif component.name == "VTIMEZONE":
                tzid = component.get('tzid')
                if tzid:
                    calendar_tz = normalize_timezone(str(tzid))
                    break
        
        print(f"Calendar timezone: {calendar_tz}")
        print("Converting events to Google Calendar format...")
        sys.stdout.flush()
        
        # Parse events
        event_count = 0
        for component in cal.walk():
            if component.name == "VEVENT":
                try:
                    event = self._convert_vevent_to_google_event(component, calendar_tz, include_attendees)
                    if event:
                        events.append(event)
                        event_count += 1
                        if event_count % 1000 == 0:
                            print(f"  Converted {event_count} events...")
                            sys.stdout.flush()
                except Exception as e:
                    print(f"Warning: Could not parse event: {e}")
                    self.stats['errors'] += 1
        
        print(f"Found {len(events)} events in file")
        sys.stdout.flush()
        return events, calendar_tz
    
    def _convert_vevent_to_google_event(self, vevent, calendar_tz: str, include_attendees: bool = True) -> Optional[Dict[str, Any]]:
        """Convert an iCalendar VEVENT to Google Calendar event format"""
        event = {}
        
        # UID - required for import API
        uid = vevent.get('uid')
        if uid:
            # Google requires iCalUID for import
            event['iCalUID'] = str(uid)
        else:
            # Generate a UID if none exists
            event['iCalUID'] = hashlib.md5(str(vevent.to_ical()).encode()).hexdigest() + "@imported"
        
        # Summary (title)
        summary = vevent.get('summary')
        if summary and str(summary).strip():
            event['summary'] = str(summary)
        else:
            # Smart default title based on event content
            description = vevent.get('description')
            location = vevent.get('location')
            desc_str = str(description).lower() if description else ''
            loc_str = str(location).lower() if location else ''
            
            if 'zoom' in desc_str or 'zoom' in loc_str:
                event['summary'] = 'Zoom Meeting'
            elif 'teams' in desc_str or 'teams' in loc_str or 'microsoft teams' in desc_str:
                event['summary'] = 'Teams Meeting'
            elif 'webex' in desc_str or 'webex' in loc_str:
                event['summary'] = 'Webex Meeting'
            elif 'meet.google' in desc_str or 'meet.google' in loc_str:
                event['summary'] = 'Google Meet'
            else:
                # Check busy status
                busystatus = vevent.get('x-microsoft-cdo-busystatus')
                transp = vevent.get('transp')
                
                if busystatus:
                    status_str = str(busystatus).upper()
                    if status_str == 'FREE':
                        event['summary'] = 'Free'
                    elif status_str == 'TENTATIVE':
                        event['summary'] = 'Tentative'
                    elif status_str == 'OOF':
                        event['summary'] = 'Out of Office'
                    else:
                        event['summary'] = 'Busy'
                elif transp and str(transp).upper() == 'TRANSPARENT':
                    event['summary'] = 'Free'
                else:
                    event['summary'] = 'Busy'
        
        # Description
        description = vevent.get('description')
        if description:
            event['description'] = str(description)
        
        # Location
        location = vevent.get('location')
        if location:
            event['location'] = str(location)
        
        # Start time
        dtstart = vevent.get('dtstart')
        if not dtstart:
            return None
        
        start_dt = dtstart.dt
        start_tz = self._get_timezone(dtstart, calendar_tz)
        
        # End time
        dtend = vevent.get('dtend')
        if dtend:
            end_dt = dtend.dt
            end_tz = self._get_timezone(dtend, calendar_tz)
        else:
            # If no end time, check for duration or assume 1 hour/1 day
            duration = vevent.get('duration')
            if duration:
                end_dt = start_dt + duration.dt
            elif isinstance(start_dt, date) and not isinstance(start_dt, datetime):
                end_dt = start_dt + timedelta(days=1)
            else:
                end_dt = start_dt + timedelta(hours=1)
            end_tz = start_tz
        
        # Check if all-day event (date vs datetime)
        if isinstance(start_dt, date) and not isinstance(start_dt, datetime):
            # All-day event
            event['start'] = {'date': start_dt.isoformat()}
            if isinstance(end_dt, date) and not isinstance(end_dt, datetime):
                event['end'] = {'date': end_dt.isoformat()}
            else:
                event['end'] = {'date': end_dt.date().isoformat()}
        else:
            # Timed event
            event['start'] = {
                'dateTime': start_dt.isoformat(),
                'timeZone': start_tz
            }
            event['end'] = {
                'dateTime': end_dt.isoformat(),
                'timeZone': end_tz
            }
        
        # Recurrence rules
        rrule = vevent.get('rrule')
        if rrule:
            recurrence = []
            rrule_str = rrule.to_ical().decode('utf-8')
            recurrence.append(f'RRULE:{rrule_str}')
            
            # Handle EXDATE (exceptions to recurrence)
            exdates = vevent.get('exdate')
            if exdates:
                if not isinstance(exdates, list):
                    exdates = [exdates]
                for exdate in exdates:
                    if hasattr(exdate, 'dts'):
                        for dt in exdate.dts:
                            if isinstance(dt.dt, date) and not isinstance(dt.dt, datetime):
                                recurrence.append(f'EXDATE;VALUE=DATE:{dt.dt.strftime("%Y%m%d")}')
                            else:
                                recurrence.append(f'EXDATE:{dt.dt.strftime("%Y%m%dT%H%M%SZ")}')
            
            # Handle RDATE (additional dates)
            rdates = vevent.get('rdate')
            if rdates:
                if not isinstance(rdates, list):
                    rdates = [rdates]
                for rdate in rdates:
                    if hasattr(rdate, 'dts'):
                        for dt in rdate.dts:
                            if isinstance(dt.dt, date) and not isinstance(dt.dt, datetime):
                                recurrence.append(f'RDATE;VALUE=DATE:{dt.dt.strftime("%Y%m%d")}')
                            else:
                                recurrence.append(f'RDATE:{dt.dt.strftime("%Y%m%dT%H%M%SZ")}')
            
            event['recurrence'] = recurrence
        
        # Status
        status = vevent.get('status')
        if status:
            status_str = str(status).upper()
            if status_str == 'CANCELLED':
                event['status'] = 'cancelled'
            elif status_str == 'TENTATIVE':
                event['status'] = 'tentative'
            else:
                event['status'] = 'confirmed'
        
        # Transparency (busy/free)
        transp = vevent.get('transp')
        if transp:
            event['transparency'] = 'transparent' if str(transp).upper() == 'TRANSPARENT' else 'opaque'
        
        # Visibility/Class
        classification = vevent.get('class')
        if classification:
            class_str = str(classification).upper()
            if class_str == 'PRIVATE':
                event['visibility'] = 'private'
            elif class_str == 'CONFIDENTIAL':
                event['visibility'] = 'confidential'
            else:
                event['visibility'] = 'default'
        
        # Organizer
        organizer = vevent.get('organizer')
        if organizer:
            org_email = str(organizer).replace('mailto:', '').replace('MAILTO:', '')
            # Skip invalid organizer emails
            if org_email and '@' in org_email and not org_email.startswith('invalid:'):
                org_name = organizer.params.get('CN', '') if hasattr(organizer, 'params') else ''
                event['organizer'] = {'email': org_email}
                if org_name:
                    event['organizer']['displayName'] = org_name
        
        # Attendees
        if include_attendees:
            attendees = vevent.get('attendee')
            if attendees:
                if not isinstance(attendees, list):
                    attendees = [attendees]
                
                event_attendees = []
                for attendee in attendees:
                    try:
                        email = str(attendee).replace('mailto:', '').replace('MAILTO:', '')
                        
                        # Skip invalid email addresses
                        if not email or '@' not in email:
                            continue
                        if email.startswith('invalid:'):
                            continue
                        if email.endswith('@resource.calendar.google.com'):
                            # Google resource calendars - skip or mark as resource
                            continue
                        
                        att_data = {'email': email}
                        
                        # Get attendee parameters
                        if hasattr(attendee, 'params'):
                            params = attendee.params
                            
                            # Display name
                            cn = params.get('CN')
                            if cn:
                                att_data['displayName'] = str(cn)
                            
                            # Response status
                            partstat = params.get('PARTSTAT', '').upper()
                            if partstat == 'ACCEPTED':
                                att_data['responseStatus'] = 'accepted'
                            elif partstat == 'DECLINED':
                                att_data['responseStatus'] = 'declined'
                            elif partstat == 'TENTATIVE':
                                att_data['responseStatus'] = 'tentative'
                            else:
                                att_data['responseStatus'] = 'needsAction'
                            
                            # Role
                            role = params.get('ROLE', '').upper()
                            if role == 'OPT-PARTICIPANT':
                                att_data['optional'] = True
                            
                            # Resource (room, equipment)
                            cutype = params.get('CUTYPE', '').upper()
                            if cutype == 'RESOURCE' or cutype == 'ROOM':
                                att_data['resource'] = True
                        
                        event_attendees.append(att_data)
                        self.stats['attendees_imported'] += 1
                        
                    except Exception as e:
                        # Skip malformed attendees
                        continue
                
                if event_attendees:
                    event['attendees'] = event_attendees
        
        # Reminders/Alarms
        reminders = []
        for component in vevent.walk():
            if component.name == "VALARM":
                try:
                    action = component.get('action')
                    trigger = component.get('trigger')
                    
                    if trigger and hasattr(trigger, 'dt'):
                        # Convert trigger to minutes before event
                        if isinstance(trigger.dt, timedelta):
                            minutes = abs(int(trigger.dt.total_seconds() / 60))
                            
                            # Google Calendar max is 40320 minutes (4 weeks)
                            if minutes > 40320:
                                minutes = 40320
                            
                            method = 'popup'  # Default to popup
                            if action and str(action).upper() == 'EMAIL':
                                method = 'email'
                            
                            reminders.append({
                                'method': method,
                                'minutes': minutes
                            })
                except Exception:
                    continue
        
        if reminders:
            # Deduplicate and limit to 5 reminders (Google's limit)
            seen = set()
            unique_reminders = []
            for r in reminders:
                key = (r['method'], r['minutes'])
                if key not in seen:
                    seen.add(key)
                    unique_reminders.append(r)
            
            event['reminders'] = {
                'useDefault': False,
                'overrides': unique_reminders[:5]
            }
        
        # Sequence number
        sequence = vevent.get('sequence')
        if sequence:
            event['sequence'] = int(sequence)
        
        # Store original UID in extended properties for reference
        if uid:
            event['extendedProperties'] = {
                'private': {
                    'outlookUID': str(uid)
                }
            }
        
        return event
    
    def _get_timezone(self, dt_prop, default_tz: str) -> str:
        """Extract and normalize timezone from a datetime property"""
        if hasattr(dt_prop, 'params'):
            tzid = dt_prop.params.get('TZID')
            if tzid:
                return normalize_timezone(str(tzid))
        
        # Check if datetime has tzinfo
        if hasattr(dt_prop, 'dt') and hasattr(dt_prop.dt, 'tzinfo') and dt_prop.dt.tzinfo:
            tz_name = str(dt_prop.dt.tzinfo)
            if tz_name and tz_name != 'UTC':
                return normalize_timezone(tz_name)
        
        return default_tz
    
    def import_events(self, events: List[Dict[str, Any]], calendar_id: str = 'primary', 
                      skip_duplicates: bool = True, dry_run: bool = False) -> None:
        """Import events to Google Calendar using the import API (no notifications sent)"""
        self.dry_run = dry_run  # Track for summary
        total = len(events)
        self.stats['total_events'] += total
        
        mode_str = "DRY RUN - " if dry_run else ""
        print(f"\n{mode_str}Importing {total} events to Google Calendar...")
        print(f"Calendar ID: {calendar_id}")
        if dry_run:
            print("MODE: Dry run - no events will be created")
        else:
            print("Using import API - NO email notifications will be sent to attendees")
        print("-" * 60)
        
        # Get existing event UIDs if checking for duplicates
        existing_uids = set()
        if skip_duplicates and not dry_run:
            print("Checking for existing events...")
            existing_uids = self._get_existing_uids(calendar_id)
            print(f"Found {len(existing_uids)} existing events")
        
        imported = 0
        skipped = 0
        errors = 0
        last_progress_time = time.time()
        
        for i, event in enumerate(events):
            # Check for duplicate by iCalUID
            ical_uid = event.get('iCalUID', '')
            if skip_duplicates and ical_uid in existing_uids:
                skipped += 1
                continue
            
            if dry_run:
                # In dry-run mode, just count and optionally show details
                imported += 1
                if imported <= 10:
                    # Show first 10 events
                    summary = event.get('summary', '(no title)')
                    start = event.get('start', {})
                    start_str = start.get('dateTime', start.get('date', 'unknown'))
                    attendee_count = len(event.get('attendees', []))
                    att_str = f" ({attendee_count} attendees)" if attendee_count else ""
                    print(f"  Would import: {summary[:60]}{att_str}")
                    print(f"               Start: {start_str}")
                    sys.stdout.flush()
                elif imported == 11:
                    print(f"  ... and more events")
                    sys.stdout.flush()
                continue
            
            # Rate limiting
            if imported > 0 and imported % REQUESTS_PER_SECOND == 0:
                time.sleep(1)
            
            try:
                # Use import_ instead of insert - this doesn't send notifications
                self.service.events().import_(
                    calendarId=calendar_id,
                    body=event
                ).execute()
                imported += 1
                
                # Add to existing UIDs to prevent duplicates within same run
                if ical_uid:
                    existing_uids.add(ical_uid)
                
            except HttpError as e:
                if e.resp.status == 429:  # Rate limit exceeded
                    print("\nRate limit hit, waiting 60 seconds...")
                    time.sleep(60)
                    # Retry
                    try:
                        self.service.events().import_(
                            calendarId=calendar_id,
                            body=event
                        ).execute()
                        imported += 1
                        if ical_uid:
                            existing_uids.add(ical_uid)
                    except Exception as retry_error:
                        errors += 1
                        if errors <= 10:
                            print(f"Retry failed for '{event.get('summary', 'Unknown')}': {retry_error}")
                elif e.resp.status == 409:  # Conflict - event already exists
                    skipped += 1
                elif e.resp.status == 400 and 'participantIsNeitherOrganizerNorAttendee' in str(e):
                    # User is not organizer or attendee - retry without attendees/organizer
                    event_copy = event.copy()
                    event_copy.pop('organizer', None)
                    event_copy.pop('attendees', None)
                    try:
                        self.service.events().import_(
                            calendarId=calendar_id,
                            body=event_copy
                        ).execute()
                        imported += 1
                        self.stats['imported_without_attendees'] = self.stats.get('imported_without_attendees', 0) + 1
                        if ical_uid:
                            existing_uids.add(ical_uid)
                    except Exception as retry_error:
                        errors += 1
                        if errors <= 10:
                            print(f"Fallback failed for '{event.get('summary', 'Unknown')}': {retry_error}")
                else:
                    errors += 1
                    if errors <= 10:
                        print(f"Error importing '{event.get('summary', 'Unknown')}': {e}")
                    elif errors == 11:
                        print("(Suppressing further error messages...)")
            
            except Exception as e:
                errors += 1
                if errors <= 10:
                    print(f"Error importing '{event.get('summary', 'Unknown')}': {e}")
            
            # Progress update every 100 events or every 10 seconds (skip for dry-run, shown at end)
            if not dry_run:
                current_time = time.time()
                if (i + 1) % 100 == 0 or (i + 1) == total or (current_time - last_progress_time) > 10:
                    elapsed = current_time - last_progress_time
                    rate = 100 / elapsed if elapsed > 0 and (i + 1) % 100 == 0 else 0
                    print(f"Progress: {i + 1}/{total} ({imported} imported, {skipped} skipped, {errors} errors)" + 
                          (f" [{rate:.1f} events/sec]" if rate > 0 else ""))
                    last_progress_time = current_time
        
        self.stats['imported'] += imported
        self.stats['skipped'] += skipped
        self.stats['errors'] += errors
        
        print("-" * 60)
        if dry_run:
            print(f"DRY RUN COMPLETE: Would import {imported} events, skip {skipped} duplicates")
        else:
            print(f"File complete: {imported} imported, {skipped} skipped, {errors} errors")
    
    def _get_existing_uids(self, calendar_id: str) -> Set[str]:
        """Get iCalUIDs of events already in the calendar"""
        uids = set()
        page_token = None
        
        try:
            while True:
                events_result = self.service.events().list(
                    calendarId=calendar_id,
                    pageToken=page_token,
                    maxResults=2500,
                    singleEvents=False,
                    showDeleted=False
                ).execute()
                
                for event in events_result.get('items', []):
                    # Get iCalUID
                    ical_uid = event.get('iCalUID')
                    if ical_uid:
                        uids.add(ical_uid)
                    
                    # Also check extended properties for outlookUID
                    outlook_uid = event.get('extendedProperties', {}).get('private', {}).get('outlookUID')
                    if outlook_uid:
                        uids.add(outlook_uid)
                
                page_token = events_result.get('nextPageToken')
                if not page_token:
                    break
                    
        except Exception as e:
            print(f"Warning: Could not fully check for existing events: {e}")
        
        return uids
    
    def print_summary(self) -> None:
        """Print import summary"""
        print("\n" + "=" * 60)
        if self.dry_run:
            print("DRY RUN SUMMARY")
        else:
            print("IMPORT SUMMARY")
        print("=" * 60)
        print(f"Total events processed:  {self.stats['total_events']}")
        if self.dry_run:
            print(f"Would import:            {self.stats['imported']}")
            print(f"Would skip (duplicates): {self.stats['skipped']}")
        else:
            print(f"Successfully imported:   {self.stats['imported']}")
            if self.stats.get('imported_without_attendees', 0) > 0:
                print(f"  (of which {self.stats['imported_without_attendees']} imported without attendees - you weren't organizer/attendee)")
            print(f"Skipped (duplicates):    {self.stats['skipped']}")
        print(f"Errors:                  {self.stats['errors']}")
        if not self.dry_run:
            print(f"Attendees imported:      {self.stats['attendees_imported']}")
        print("=" * 60)
        if not self.dry_run:
            print("\nNOTE: No email notifications were sent to any attendees.")


def find_ics_files(path: str) -> List[str]:
    """Find all ICS files in a path (file or directory)"""
    path = Path(path)
    
    if path.is_file():
        if path.suffix.lower() == '.ics':
            return [str(path)]
        else:
            print(f"Error: {path} is not an ICS file")
            return []
    
    elif path.is_dir():
        ics_files = list(path.glob('*.ics')) + list(path.glob('*.ICS'))
        return [str(f) for f in sorted(ics_files)]
    
    else:
        print(f"Error: {path} not found")
        return []


def main():
    parser = argparse.ArgumentParser(
        description='Import ICS calendar files to Google Calendar (with attendees, no notifications)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s calendar.ics
  %(prog)s ./ics_files/
  %(prog)s calendar.ics --calendar "work@group.calendar.google.com"
  %(prog)s calendar.ics --no-attendees
  %(prog)s calendar.ics --no-skip-duplicates
  %(prog)s --list-calendars

Notes:
  - Uses Google Calendar's import API, so NO notifications are sent to attendees
  - Attendees are preserved with their response status (accepted/declined/tentative)
  - Reminders/alarms are imported (up to 5 per event)
  - Recurring events with exceptions are supported
        """
    )
    
    parser.add_argument('ics_path', nargs='?', help='Path to ICS file or directory containing ICS files')
    parser.add_argument('--calendar', '-c', default='primary',
                        help='Target Google Calendar ID (default: primary)')
    parser.add_argument('--credentials', default='credentials.json',
                        help='Path to Google OAuth credentials file (default: credentials.json)')
    parser.add_argument('--no-skip-duplicates', action='store_true',
                        help='Import all events even if they already exist')
    parser.add_argument('--no-attendees', action='store_true',
                        help='Do not import attendees')
    parser.add_argument('--list-calendars', action='store_true',
                        help='List available Google Calendars and exit')
    parser.add_argument('--dry-run', action='store_true',
                        help='Simulate import without creating events (shows what would be imported)')
    parser.add_argument('--start-date', type=str,
                        help='Only import events starting on or after this date (YYYY-MM-DD)')
    parser.add_argument('--end-date', type=str,
                        help='Only import events starting before this date (YYYY-MM-DD)')
    parser.add_argument('--limit', type=int,
                        help='Maximum number of events to import (useful for testing)')
    parser.add_argument('--add-self', type=str, metavar='EMAIL',
                        help='Add this email as attendee to all events (prevents "not organizer/attendee" errors)')
    
    args = parser.parse_args()
    
    # Initialize importer
    importer = GoogleCalendarImporter(credentials_file=args.credentials)
    
    print("=" * 60)
    print("ICS to Google Calendar Importer (Enhanced)")
    print("=" * 60)
    print("Features:")
    print("  ✓ Imports attendees WITHOUT sending email notifications")
    print("  ✓ Preserves organizer, reminders, recurrence rules")
    print("  ✓ Handles timezones (including Outlook/Windows formats)")
    print("  ✓ Duplicate detection and skipping")
    print("=" * 60)
    
    # Authenticate
    print("\nAuthenticating with Google...")
    sys.stdout.flush()
    importer.authenticate()
    print("Connected to Google Calendar API")
    sys.stdout.flush()
    
    # List calendars if requested
    if args.list_calendars:
        print("\nAvailable calendars:")
        print("-" * 60)
        calendars = importer.list_calendars()
        for cal in calendars:
            primary = " (PRIMARY)" if cal['primary'] else ""
            print(f"  {cal['summary']}{primary}")
            print(f"    ID: {cal['id']}")
        return
    
    # Check for ICS path
    if not args.ics_path:
        parser.print_help()
        print("\nError: Please provide an ICS file or directory path")
        sys.exit(1)
    
    # Parse date filters
    start_date = None
    end_date = None
    if args.start_date:
        try:
            start_date = datetime.strptime(args.start_date, '%Y-%m-%d').date()
            print(f"Filtering: events on or after {start_date}")
        except ValueError:
            print(f"Error: Invalid start date format '{args.start_date}'. Use YYYY-MM-DD")
            sys.exit(1)
    
    if args.end_date:
        try:
            end_date = datetime.strptime(args.end_date, '%Y-%m-%d').date()
            print(f"Filtering: events before {end_date}")
        except ValueError:
            print(f"Error: Invalid end date format '{args.end_date}'. Use YYYY-MM-DD")
            sys.exit(1)
    
    if args.dry_run:
        print("\n*** DRY RUN MODE - No events will be created ***")
    
    if args.limit:
        print(f"Limiting to first {args.limit} events")
    
    # Find ICS files
    ics_files = find_ics_files(args.ics_path)
    
    if not ics_files:
        print("No ICS files found")
        sys.exit(1)
    
    print(f"\nFound {len(ics_files)} ICS file(s) to process")
    if not args.no_attendees:
        print("Attendee import: ENABLED (no emails will be sent)")
    else:
        print("Attendee import: DISABLED")
    
    # Process each file
    for ics_file in ics_files:
        events, _ = importer.parse_ics_file(ics_file, include_attendees=not args.no_attendees)
        
        # Apply date filtering
        if start_date or end_date:
            original_count = len(events)
            events = filter_events_by_date(events, start_date, end_date)
            print(f"Date filter: {original_count} -> {len(events)} events")
        
        # Apply limit
        if args.limit and len(events) > args.limit:
            events = events[:args.limit]
            print(f"Limited to first {args.limit} events")
        
        # Add self as attendee if requested
        if args.add_self and not args.no_attendees:
            print(f"Adding {args.add_self} as attendee to all events...")
            for event in events:
                attendees = event.get('attendees', [])
                # Check if already in attendees list
                emails = [a.get('email', '').lower() for a in attendees]
                if args.add_self.lower() not in emails:
                    attendees.append({
                        'email': args.add_self,
                        'responseStatus': 'accepted'
                    })
                    event['attendees'] = attendees
        
        if events:
            importer.import_events(
                events,
                calendar_id=args.calendar,
                skip_duplicates=not args.no_skip_duplicates,
                dry_run=args.dry_run
            )
    
    # Print summary
    importer.print_summary()
    if args.dry_run:
        print("\n*** DRY RUN COMPLETE - No events were created ***")
    print("\nDone!")


if __name__ == '__main__':
    main()
