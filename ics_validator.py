#!/usr/bin/env python3
"""
ICS Pre-Import Validator

Randomly samples different categories of events to validate data quality
before importing to Google Calendar.

Usage:
    python ics_validator.py <ics_file> [--samples 5]
"""

import sys
import random
from datetime import datetime, date, timedelta
from typing import List, Dict, Any, Optional
from collections import defaultdict
from icalendar import Calendar

# Timezone mappings (same as importer)
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
    """Convert Windows timezone to IANA"""
    if not tz_str:
        return 'UTC'
    if tz_str in TIMEZONE_MAPPINGS:
        return TIMEZONE_MAPPINGS[tz_str]
    for win_tz, iana_tz in TIMEZONE_MAPPINGS.items():
        if win_tz.lower() in tz_str.lower():
            return iana_tz
    return tz_str  # Return as-is if unknown


def get_event_info(vevent) -> Dict[str, Any]:
    """Extract key info from a VEVENT"""
    info = {}
    
    # Basic fields
    info['summary'] = str(vevent.get('summary', '')) or '(no title)'
    info['description'] = str(vevent.get('description', ''))[:200] if vevent.get('description') else None
    info['location'] = str(vevent.get('location', ''))[:100] if vevent.get('location') else None
    info['uid'] = str(vevent.get('uid', ''))[:50]
    
    # Start/End
    dtstart = vevent.get('dtstart')
    if dtstart:
        info['start'] = dtstart.dt
        info['start_tz'] = dtstart.params.get('TZID') if hasattr(dtstart, 'params') else None
        info['is_all_day'] = isinstance(dtstart.dt, date) and not isinstance(dtstart.dt, datetime)
    
    dtend = vevent.get('dtend')
    if dtend:
        info['end'] = dtend.dt
    
    # Recurrence
    rrule = vevent.get('rrule')
    if rrule:
        info['rrule'] = rrule.to_ical().decode('utf-8')
    
    exdate = vevent.get('exdate')
    if exdate:
        info['has_exdate'] = True
        if isinstance(exdate, list):
            info['exdate_count'] = sum(len(e.dts) if hasattr(e, 'dts') else 1 for e in exdate)
        elif hasattr(exdate, 'dts'):
            info['exdate_count'] = len(exdate.dts)
        else:
            info['exdate_count'] = 1
    
    # Attendees
    attendees = vevent.get('attendee')
    if attendees:
        if not isinstance(attendees, list):
            attendees = [attendees]
        info['attendee_count'] = len(attendees)
        info['attendees'] = []
        for att in attendees[:5]:
            att_info = {
                'email': str(att).replace('mailto:', '').replace('MAILTO:', '')
            }
            if hasattr(att, 'params'):
                att_info['name'] = att.params.get('CN', '')
                att_info['status'] = att.params.get('PARTSTAT', 'UNKNOWN')
            info['attendees'].append(att_info)
    
    # Organizer
    organizer = vevent.get('organizer')
    if organizer:
        info['organizer'] = str(organizer).replace('mailto:', '').replace('MAILTO:', '')
        if hasattr(organizer, 'params') and organizer.params.get('CN'):
            info['organizer_name'] = organizer.params.get('CN')
    
    # Reminders
    reminders = []
    for component in vevent.walk():
        if component.name == "VALARM":
            trigger = component.get('trigger')
            action = component.get('action')
            if trigger and hasattr(trigger, 'dt'):
                if isinstance(trigger.dt, timedelta):
                    minutes = abs(int(trigger.dt.total_seconds() / 60))
                    reminders.append({
                        'minutes': minutes,
                        'action': str(action) if action else 'DISPLAY'
                    })
    if reminders:
        info['reminders'] = reminders
    
    # Status/Visibility
    info['status'] = str(vevent.get('status', '')) or None
    info['class'] = str(vevent.get('class', '')) or None
    info['transp'] = str(vevent.get('transp', '')) or None
    info['busystatus'] = str(vevent.get('x-microsoft-cdo-busystatus', '')) or None
    
    return info


def print_event(event: Dict[str, Any], index: int):
    """Pretty print an event"""
    print(f"\n  --- Sample {index} ---")
    print(f"    Title:      {event['summary'][:60]}{'...' if len(event['summary']) > 60 else ''}")
    
    if event.get('start'):
        if event.get('is_all_day'):
            print(f"    Start:      {event['start']} (all-day)")
        else:
            tz = event.get('start_tz', 'UTC')
            tz_mapped = normalize_timezone(tz) if tz else 'UTC'
            print(f"    Start:      {event['start']}")
            if tz:
                print(f"    Timezone:   {tz} → {tz_mapped}")
    
    if event.get('end'):
        print(f"    End:        {event['end']}")
    
    if event.get('location'):
        print(f"    Location:   {event['location'][:70]}{'...' if len(event['location']) > 70 else ''}")
    
    if event.get('rrule'):
        print(f"    Recurrence: {event['rrule'][:60]}{'...' if len(event['rrule']) > 60 else ''}")
        if event.get('has_exdate'):
            print(f"    Exceptions: {event['exdate_count']} dates excluded")
    
    if event.get('organizer'):
        org_name = f" ({event['organizer_name']})" if event.get('organizer_name') else ''
        print(f"    Organizer:  {event['organizer']}{org_name}")
    
    if event.get('attendee_count'):
        print(f"    Attendees:  {event['attendee_count']} total")
        for att in event.get('attendees', [])[:3]:
            name = f" ({att['name']})" if att.get('name') else ''
            status = att.get('status', 'UNKNOWN')
            print(f"                - {att['email']}{name} [{status}]")
        if event['attendee_count'] > 3:
            print(f"                ... and {event['attendee_count'] - 3} more")
    
    if event.get('reminders'):
        rem_strs = [f"{r['minutes']}min ({r['action']})" for r in event['reminders'][:3]]
        print(f"    Reminders:  {', '.join(rem_strs)}")
    
    if event.get('class') and event['class'] != 'PUBLIC':
        print(f"    Visibility: {event['class']}")
    
    if event.get('description'):
        desc = event['description'].replace('\n', ' ')[:100]
        print(f"    Description: {desc}{'...' if len(event['description']) > 100 else ''}")


def validate_ics(ics_path: str, samples_per_category: int = 5):
    """Main validation function"""
    
    print(f"Loading: {ics_path}")
    
    with open(ics_path, 'rb') as f:
        cal = Calendar.from_ical(f.read())
    
    # Categorize events
    all_events = []
    events_with_attendees = []
    events_recurring = []
    events_with_reminders = []
    events_all_day = []
    events_timed = []
    events_private = []
    events_with_location = []
    events_old = []  # Before 2020
    events_recent = []  # 2024 or later
    events_long_description = []
    
    # Additional edge case categories
    events_many_attendees = []  # Events with 50+ attendees
    events_invalid_organizer = []
    events_long_title = []
    events_no_end = []
    duplicate_uids = defaultdict(list)
    
    for component in cal.walk():
        if component.name != "VEVENT":
            continue
        
        info = get_event_info(component)
        all_events.append(info)
        
        # Track UIDs for duplicates
        if info.get('uid'):
            duplicate_uids[info['uid']].append(info)
        
        # Check for many attendees
        if info.get('attendee_count', 0) >= 50:
            events_many_attendees.append(info)
        
        # Check for invalid organizer
        if info.get('organizer') and info['organizer'].startswith('invalid:'):
            events_invalid_organizer.append(info)
        
        # Check for long titles (Google limit is ~1024 chars, but display issues start earlier)
        if len(info.get('summary', '')) > 200:
            events_long_title.append(info)
        
        # Check for missing end time
        if not info.get('end'):
            events_no_end.append(info)
        
        # Categorize
        if info.get('attendee_count'):
            events_with_attendees.append(info)
        
        if info.get('rrule'):
            events_recurring.append(info)
        
        if info.get('reminders'):
            events_with_reminders.append(info)
        
        if info.get('is_all_day'):
            events_all_day.append(info)
        else:
            events_timed.append(info)
        
        if info.get('class') in ['PRIVATE', 'CONFIDENTIAL']:
            events_private.append(info)
        
        if info.get('location'):
            events_with_location.append(info)
        
        # Date-based categorization
        start = info.get('start')
        if start:
            try:
                year = start.year if hasattr(start, 'year') else None
                if year and year < 2020:
                    events_old.append(info)
                elif year and year >= 2024:
                    events_recent.append(info)
            except:
                pass
        
        if info.get('description') and len(info['description']) > 150:
            events_long_description.append(info)
    
    # Print report
    print("\n" + "=" * 70)
    print("PRE-IMPORT VALIDATION REPORT")
    print("=" * 70)
    
    print(f"\nTotal events: {len(all_events)}")
    
    categories = [
        ("EVENTS WITH ATTENDEES", events_with_attendees, 
         "Verifying attendee emails, names, and response statuses"),
        ("RECURRING EVENTS", events_recurring,
         "Verifying recurrence rules and exceptions (EXDATE)"),
        ("EVENTS WITH REMINDERS", events_with_reminders,
         "Verifying reminder times and types"),
        ("ALL-DAY EVENTS", events_all_day,
         "Verifying all-day event handling"),
        ("TIMED EVENTS (with timezone)", events_timed,
         "Verifying timezone conversion"),
        ("PRIVATE/CONFIDENTIAL EVENTS", events_private,
         "Verifying visibility settings"),
        ("EVENTS WITH LOCATION", events_with_location,
         "Verifying location field (including Zoom/Teams links)"),
        ("OLD EVENTS (before 2020)", events_old,
         "Verifying historical events import correctly"),
        ("RECENT EVENTS (2024+)", events_recent,
         "Verifying recent events look correct"),
        ("EVENTS WITH LONG DESCRIPTIONS", events_long_description,
         "Verifying descriptions aren't truncated"),
    ]
    
    issues_found = []
    
    for category_name, events, description in categories:
        print(f"\n{'=' * 70}")
        print(f"{category_name} ({len(events)} total)")
        print(f"Purpose: {description}")
        print("-" * 70)
        
        if not events:
            print("  (none found)")
            continue
        
        # Random sample
        sample = random.sample(events, min(samples_per_category, len(events)))
        
        for i, event in enumerate(sample, 1):
            print_event(event, i)
            
            # Validation checks
            if event.get('start_tz'):
                mapped = normalize_timezone(event['start_tz'])
                if mapped == event['start_tz'] and 'Standard Time' in event['start_tz']:
                    issues_found.append(f"Unknown timezone: {event['start_tz']}")
            
            if event.get('attendees'):
                for att in event['attendees']:
                    email = att.get('email', '')
                    if email.startswith('invalid:'):
                        issues_found.append(f"Invalid attendee (will be skipped): {email}")
                    elif '@' not in email:
                        issues_found.append(f"Invalid attendee email (will be skipped): {email}")
    
    # Summary
    print("\n" + "=" * 70)
    print("VALIDATION SUMMARY")
    print("=" * 70)
    
    print(f"\n✓ Events with attendees:    {len(events_with_attendees):,} ({len(events_with_attendees)/len(all_events)*100:.1f}%)")
    print(f"✓ Recurring events:         {len(events_recurring):,}")
    print(f"✓ Events with reminders:    {len(events_with_reminders):,}")
    print(f"✓ All-day events:           {len(events_all_day):,}")
    print(f"✓ Timed events:             {len(events_timed):,}")
    print(f"✓ Private/Confidential:     {len(events_private):,}")
    print(f"✓ Events with location:     {len(events_with_location):,}")
    print(f"✓ Old events (pre-2020):    {len(events_old):,}")
    print(f"✓ Recent events (2024+):    {len(events_recent):,}")
    
    # Edge case summary
    actual_duplicates = {uid: evts for uid, evts in duplicate_uids.items() if len(evts) > 1}
    
    # Collect distribution lists
    distribution_lists = set()
    resource_calendars = set()
    for event in all_events:
        for att in event.get('attendees', []):
            email = att.get('email', '')
            if email.lower().startswith('dl-'):
                distribution_lists.add(email)
            if '@resource.calendar.google.com' in email.lower():
                resource_calendars.add(email)
    
    print("\n" + "-" * 70)
    print("EDGE CASES")
    print("-" * 70)
    
    edge_cases_found = False
    
    if events_invalid_organizer:
        edge_cases_found = True
        print(f"\n📋 Invalid Organizers: {len(events_invalid_organizer)} events")
        print("   These events have 'invalid:nomail' as organizer - will import without organizer")
        for evt in events_invalid_organizer[:3]:
            print(f"   • {evt.get('summary', '(no title)')[:60]}")
        if len(events_invalid_organizer) > 3:
            print(f"   ... and {len(events_invalid_organizer) - 3} more")
    
    if events_many_attendees:
        edge_cases_found = True
        print(f"\n👥 Large Attendee Lists (50+): {len(events_many_attendees)} events")
        print("   Google Calendar import API typically handles these fine")
        for evt in sorted(events_many_attendees, key=lambda x: -x.get('attendee_count', 0))[:5]:
            print(f"   • {evt.get('summary', '(no title)')[:50]} ({evt.get('attendee_count')} attendees)")
        if len(events_many_attendees) > 5:
            print(f"   ... and {len(events_many_attendees) - 5} more")
    
    if distribution_lists:
        edge_cases_found = True
        print(f"\n📧 Distribution Lists: {len(distribution_lists)} unique")
        print("   Will import as single attendees (members won't see event)")
        for dl in list(distribution_lists)[:5]:
            print(f"   • {dl}")
        if len(distribution_lists) > 5:
            print(f"   ... and {len(distribution_lists) - 5} more")
    
    if resource_calendars:
        edge_cases_found = True
        print(f"\n🏢 Google Resource Calendars: {len(resource_calendars)} unique")
        print("   Will be skipped (can't transfer between accounts)")
    
    if actual_duplicates:
        edge_cases_found = True
        print(f"\n🔄 Duplicate UIDs: {len(actual_duplicates)} UIDs with multiple events")
        print("   These are updates to the same event - only latest version will remain")
    
    if events_no_end:
        edge_cases_found = True
        print(f"\n⏱️  No End Time: {len(events_no_end)} events")
        print("   Will default to 1 hour duration")
    
    if events_long_title:
        edge_cases_found = True
        print(f"\n📝 Long Titles (>200 chars): {len(events_long_title)} events")
        print("   May be truncated in Google Calendar display")
    
    if not edge_cases_found:
        print("\n✅ No edge cases detected!")
    
    if issues_found:
        print(f"\n⚠️  POTENTIAL ISSUES FOUND:")
        for issue in set(issues_found):
            print(f"  - {issue}")
    else:
        print(f"\n✅ No issues detected - ready to import!")
    
    print("\n" + "=" * 70)
    print("Run the import with:")
    print(f"  python ics_to_google_calendar.py \"{ics_path}\"")
    print("=" * 70)


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Validate ICS file before importing')
    parser.add_argument('ics_path', help='Path to ICS file')
    parser.add_argument('--samples', '-s', type=int, default=5,
                        help='Number of samples per category (default: 5)')
    
    args = parser.parse_args()
    validate_ics(args.ics_path, args.samples)


if __name__ == '__main__':
    main()
