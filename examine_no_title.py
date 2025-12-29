#!/usr/bin/env python3
"""
Examine events that have no title (SUMMARY field) in an ICS file.
"""

import sys
from datetime import datetime, date
from icalendar import Calendar

def examine_no_title_events(ics_path: str, max_events: int = 20):
    """Find and display details of events without titles"""
    
    print(f"Examining events without titles in: {ics_path}\n")
    
    with open(ics_path, 'rb') as f:
        cal = Calendar.from_ical(f.read())
    
    no_title_events = []
    total_events = 0
    
    for component in cal.walk():
        if component.name == "VEVENT":
            total_events += 1
            summary = component.get('summary')
            
            # Check if no title or empty title
            if not summary or str(summary).strip() == '':
                event_info = {
                    'summary': summary,
                    'dtstart': component.get('dtstart'),
                    'dtend': component.get('dtend'),
                    'description': component.get('description'),
                    'location': component.get('location'),
                    'organizer': component.get('organizer'),
                    'attendee': component.get('attendee'),
                    'class': component.get('class'),
                    'transp': component.get('transp'),
                    'status': component.get('status'),
                    'busystatus': component.get('x-microsoft-cdo-busystatus'),
                    'uid': component.get('uid'),
                }
                no_title_events.append(event_info)
    
    print(f"Total events: {total_events}")
    print(f"Events without title: {len(no_title_events)} ({len(no_title_events)/total_events*100:.1f}%)")
    print("=" * 70)
    
    # Analyze patterns
    print("\nPATTERN ANALYSIS:")
    print("-" * 40)
    
    # Check class (public/private)
    class_counts = {}
    for e in no_title_events:
        cls = str(e['class']).upper() if e['class'] else 'NONE'
        class_counts[cls] = class_counts.get(cls, 0) + 1
    print(f"By visibility (CLASS):")
    for cls, count in sorted(class_counts.items(), key=lambda x: -x[1]):
        print(f"  - {cls}: {count}")
    
    # Check busy status
    busy_counts = {}
    for e in no_title_events:
        busy = str(e['busystatus']).upper() if e['busystatus'] else 'NONE'
        busy_counts[busy] = busy_counts.get(busy, 0) + 1
    print(f"\nBy busy status:")
    for busy, count in sorted(busy_counts.items(), key=lambda x: -x[1]):
        print(f"  - {busy}: {count}")
    
    # Check if they have descriptions
    with_desc = sum(1 for e in no_title_events if e['description'])
    print(f"\nWith description: {with_desc}")
    print(f"Without description: {len(no_title_events) - with_desc}")
    
    # Check if they have attendees
    with_attendees = sum(1 for e in no_title_events if e['attendee'])
    print(f"\nWith attendees: {with_attendees}")
    print(f"Without attendees: {len(no_title_events) - with_attendees}")
    
    # Check if they have organizer
    with_organizer = sum(1 for e in no_title_events if e['organizer'])
    print(f"\nWith organizer: {with_organizer}")
    print(f"Without organizer: {len(no_title_events) - with_organizer}")
    
    # Show samples
    print("\n" + "=" * 70)
    print(f"SAMPLE EVENTS (showing first {min(max_events, len(no_title_events))} of {len(no_title_events)}):")
    print("=" * 70)
    
    for i, event in enumerate(no_title_events[:max_events], 1):
        print(f"\n--- Event {i} ---")
        
        # Start/End
        if event['dtstart']:
            dt = event['dtstart'].dt
            if isinstance(dt, datetime):
                print(f"  Start:       {dt.strftime('%Y-%m-%d %H:%M %Z')}")
            else:
                print(f"  Start:       {dt} (all-day)")
        
        if event['dtend']:
            dt = event['dtend'].dt
            if isinstance(dt, datetime):
                print(f"  End:         {dt.strftime('%Y-%m-%d %H:%M %Z')}")
            else:
                print(f"  End:         {dt} (all-day)")
        
        # Other fields
        if event['class']:
            print(f"  Visibility:  {event['class']}")
        
        if event['busystatus']:
            print(f"  Busy status: {event['busystatus']}")
        
        if event['transp']:
            print(f"  Show as:     {event['transp']}")
        
        if event['location']:
            loc = str(event['location'])[:80]
            print(f"  Location:    {loc}{'...' if len(str(event['location'])) > 80 else ''}")
        
        if event['organizer']:
            org = str(event['organizer']).replace('mailto:', '').replace('MAILTO:', '')
            print(f"  Organizer:   {org}")
        
        if event['attendee']:
            attendees = event['attendee']
            if not isinstance(attendees, list):
                attendees = [attendees]
            print(f"  Attendees:   {len(attendees)}")
            for att in attendees[:3]:
                email = str(att).replace('mailto:', '').replace('MAILTO:', '')
                print(f"               - {email}")
            if len(attendees) > 3:
                print(f"               ... and {len(attendees) - 3} more")
        
        if event['description']:
            desc = str(event['description'])[:200]
            print(f"  Description: {desc}{'...' if len(str(event['description'])) > 200 else ''}")
        
        if event['uid']:
            print(f"  UID:         {str(event['uid'])[:60]}...")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python examine_no_title.py <ics_file> [max_samples]")
        sys.exit(1)
    
    ics_file = sys.argv[1]
    max_samples = int(sys.argv[2]) if len(sys.argv) > 2 else 20
    
    examine_no_title_events(ics_file, max_samples)
