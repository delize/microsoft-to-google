#!/usr/bin/env python3
"""
ICS File Analyzer

Analyzes ICS calendar files to show what data they contain before importing.
Useful for understanding the structure and contents of Outlook calendar exports.

Requirements:
    pip install icalendar python-dateutil

Usage:
    python ics_analyzer.py <ics_file_or_directory>
    python ics_analyzer.py calendar.ics --show-samples 5
    python ics_analyzer.py calendar.ics --output report.txt
"""

import os
import sys
import argparse
import json
from pathlib import Path
from datetime import datetime, timedelta, date
from typing import Optional, List, Dict, Any, Set, Tuple
from collections import defaultdict, Counter
from dataclasses import dataclass, field

from icalendar import Calendar, Event
from dateutil import tz as dateutil_tz


@dataclass
class FieldStats:
    """Statistics for a single field"""
    count: int = 0
    sample_values: List[str] = field(default_factory=list)
    unique_values: Set[str] = field(default_factory=set)
    
    def add(self, value: str, max_samples: int = 5):
        self.count += 1
        if len(self.sample_values) < max_samples:
            # Truncate long values for display
            display_val = str(value)[:100] + "..." if len(str(value)) > 100 else str(value)
            if display_val not in self.sample_values:
                self.sample_values.append(display_val)
        self.unique_values.add(str(value)[:200])


@dataclass 
class AnalysisReport:
    """Complete analysis report for ICS file(s)"""
    file_count: int = 0
    total_size_bytes: int = 0
    total_events: int = 0
    
    # Event types
    all_day_events: int = 0
    timed_events: int = 0
    recurring_events: int = 0
    cancelled_events: int = 0
    
    # Date range
    earliest_event: Optional[datetime] = None
    latest_event: Optional[datetime] = None
    
    # Field presence
    fields: Dict[str, FieldStats] = field(default_factory=dict)
    
    # Attendee stats
    events_with_attendees: int = 0
    total_attendees: int = 0
    unique_attendees: Set[str] = field(default_factory=set)
    attendee_response_stats: Counter = field(default_factory=Counter)
    
    # Organizer stats
    events_with_organizer: int = 0
    unique_organizers: Set[str] = field(default_factory=set)
    
    # Recurrence stats
    recurrence_types: Counter = field(default_factory=Counter)
    events_with_exceptions: int = 0
    
    # Reminder/Alarm stats
    events_with_reminders: int = 0
    reminder_types: Counter = field(default_factory=Counter)
    
    # Timezone stats
    timezones_found: Counter = field(default_factory=Counter)
    
    # Potential issues
    issues: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    
    # Sample events
    sample_events: List[Dict[str, Any]] = field(default_factory=list)


class ICSAnalyzer:
    """Analyzes ICS files and generates detailed reports"""
    
    # Fields we track
    TRACKED_FIELDS = [
        'summary', 'description', 'location', 'dtstart', 'dtend', 
        'rrule', 'exdate', 'rdate', 'attendee', 'organizer',
        'status', 'transp', 'class', 'priority', 'categories',
        'url', 'uid', 'sequence', 'created', 'last-modified',
        'x-microsoft-cdo-busystatus', 'x-microsoft-cdo-importance',
        'x-microsoft-disallow-counter', 'x-microsoft-skypeteamsmeetingurl',
        'x-alt-desc'
    ]
    
    # Known Windows/Outlook timezone names
    WINDOWS_TIMEZONES = {
        'Eastern Standard Time', 'Eastern Daylight Time',
        'Central Standard Time', 'Central Daylight Time', 
        'Mountain Standard Time', 'Mountain Daylight Time',
        'Pacific Standard Time', 'Pacific Daylight Time',
        'GMT Standard Time', 'W. Europe Standard Time',
        'Romance Standard Time', 'Central European Standard Time',
    }
    
    def __init__(self):
        self.report = AnalysisReport()
    
    def analyze_file(self, ics_path: str, max_samples: int = 5) -> None:
        """Analyze a single ICS file"""
        self.report.file_count += 1
        file_size = os.path.getsize(ics_path)
        self.report.total_size_bytes += file_size
        
        print(f"Analyzing: {ics_path} ({file_size / 1024 / 1024:.2f} MB)")
        
        with open(ics_path, 'rb') as f:
            try:
                cal = Calendar.from_ical(f.read())
            except Exception as e:
                self.report.issues.append(f"Failed to parse {ics_path}: {e}")
                return
        
        # Check calendar-level properties
        for component in cal.walk():
            if component.name == "VCALENDAR":
                self._analyze_calendar_props(component)
            elif component.name == "VTIMEZONE":
                self._analyze_timezone(component)
            elif component.name == "VEVENT":
                self._analyze_event(component, max_samples)
    
    def _analyze_calendar_props(self, cal) -> None:
        """Analyze calendar-level properties"""
        prodid = cal.get('prodid')
        if prodid:
            self._add_field('prodid', str(prodid))
        
        version = cal.get('version')
        if version:
            self._add_field('version', str(version))
        
        x_wr_timezone = cal.get('x-wr-timezone')
        if x_wr_timezone:
            self._add_field('x-wr-timezone', str(x_wr_timezone))
            self.report.timezones_found[str(x_wr_timezone)] += 1
    
    def _analyze_timezone(self, tz_component) -> None:
        """Analyze VTIMEZONE components"""
        tzid = tz_component.get('tzid')
        if tzid:
            self.report.timezones_found[str(tzid)] += 1
            
            # Check if it's a Windows timezone
            if str(tzid) in self.WINDOWS_TIMEZONES:
                if "Windows/Outlook timezone detected" not in self.report.warnings:
                    self.report.warnings.append(
                        "Windows/Outlook timezone names detected - will be converted to IANA format during import"
                    )
    
    def _analyze_event(self, vevent, max_samples: int) -> None:
        """Analyze a single VEVENT"""
        self.report.total_events += 1
        event_data = {}
        
        # Track all present fields
        for field_name in self.TRACKED_FIELDS:
            value = vevent.get(field_name)
            if value is not None:
                self._add_field(field_name, value)
                event_data[field_name] = str(value)[:100]
        
        # Also check for any X- (extended) properties we might have missed
        for key in vevent.keys():
            key_lower = key.lower()
            if key_lower.startswith('x-') and key_lower not in [f.lower() for f in self.TRACKED_FIELDS]:
                self._add_field(key_lower, vevent.get(key))
        
        # Analyze start/end times
        dtstart = vevent.get('dtstart')
        if dtstart:
            start_dt = dtstart.dt
            
            # Check if all-day or timed
            if isinstance(start_dt, date) and not isinstance(start_dt, datetime):
                self.report.all_day_events += 1
                event_data['_type'] = 'all-day'
            else:
                self.report.timed_events += 1
                event_data['_type'] = 'timed'
                
                # Track timezone
                if hasattr(dtstart, 'params'):
                    tzid = dtstart.params.get('TZID')
                    if tzid:
                        self.report.timezones_found[str(tzid)] += 1
            
            # Track date range
            try:
                if isinstance(start_dt, datetime):
                    compare_dt = start_dt
                else:
                    compare_dt = datetime.combine(start_dt, datetime.min.time())
                
                if self.report.earliest_event is None or compare_dt < self.report.earliest_event:
                    self.report.earliest_event = compare_dt
                if self.report.latest_event is None or compare_dt > self.report.latest_event:
                    self.report.latest_event = compare_dt
            except:
                pass
        
        # Check for recurrence
        rrule = vevent.get('rrule')
        if rrule:
            self.report.recurring_events += 1
            rrule_str = rrule.to_ical().decode('utf-8')
            
            # Parse recurrence frequency
            if 'FREQ=DAILY' in rrule_str:
                self.report.recurrence_types['DAILY'] += 1
            elif 'FREQ=WEEKLY' in rrule_str:
                self.report.recurrence_types['WEEKLY'] += 1
            elif 'FREQ=MONTHLY' in rrule_str:
                self.report.recurrence_types['MONTHLY'] += 1
            elif 'FREQ=YEARLY' in rrule_str:
                self.report.recurrence_types['YEARLY'] += 1
            
            # Check for exceptions
            if vevent.get('exdate'):
                self.report.events_with_exceptions += 1
        
        # Check status
        status = vevent.get('status')
        if status and str(status).upper() == 'CANCELLED':
            self.report.cancelled_events += 1
        
        # Analyze attendees
        attendees = vevent.get('attendee')
        if attendees:
            if not isinstance(attendees, list):
                attendees = [attendees]
            
            self.report.events_with_attendees += 1
            
            for attendee in attendees:
                self.report.total_attendees += 1
                email = str(attendee).replace('mailto:', '').replace('MAILTO:', '')
                self.report.unique_attendees.add(email.lower())
                
                # Track response status
                if hasattr(attendee, 'params'):
                    partstat = attendee.params.get('PARTSTAT', 'UNKNOWN')
                    self.report.attendee_response_stats[str(partstat).upper()] += 1
        
        # Analyze organizer
        organizer = vevent.get('organizer')
        if organizer:
            self.report.events_with_organizer += 1
            org_email = str(organizer).replace('mailto:', '').replace('MAILTO:', '')
            self.report.unique_organizers.add(org_email.lower())
        
        # Analyze reminders/alarms
        has_alarm = False
        for component in vevent.walk():
            if component.name == "VALARM":
                if not has_alarm:
                    has_alarm = True
                    self.report.events_with_reminders += 1
                
                action = component.get('action')
                if action:
                    self.report.reminder_types[str(action).upper()] += 1
        
        # Store sample event
        if len(self.report.sample_events) < max_samples:
            event_data['_has_attendees'] = len(attendees) if attendees else 0
            event_data['_has_reminders'] = has_alarm
            self.report.sample_events.append(event_data)
    
    def _add_field(self, field_name: str, value: Any) -> None:
        """Add a field value to statistics"""
        field_name = field_name.lower()
        if field_name not in self.report.fields:
            self.report.fields[field_name] = FieldStats()
        
        str_value = str(value) if value else "(empty)"
        self.report.fields[field_name].add(str_value)
    
    def generate_report(self) -> str:
        """Generate a formatted text report"""
        r = self.report
        lines = []
        
        # Header
        lines.append("=" * 70)
        lines.append("ICS FILE ANALYSIS REPORT")
        lines.append("=" * 70)
        lines.append("")
        
        # File summary
        lines.append("FILE SUMMARY")
        lines.append("-" * 40)
        lines.append(f"  Files analyzed:     {r.file_count}")
        lines.append(f"  Total size:         {r.total_size_bytes / 1024 / 1024:.2f} MB")
        lines.append(f"  Total events:       {r.total_events:,}")
        lines.append("")
        
        # Event breakdown
        lines.append("EVENT BREAKDOWN")
        lines.append("-" * 40)
        lines.append(f"  Timed events:       {r.timed_events:,}")
        lines.append(f"  All-day events:     {r.all_day_events:,}")
        lines.append(f"  Recurring events:   {r.recurring_events:,}")
        lines.append(f"  Cancelled events:   {r.cancelled_events:,}")
        lines.append("")
        
        # Date range
        if r.earliest_event and r.latest_event:
            lines.append("DATE RANGE")
            lines.append("-" * 40)
            lines.append(f"  Earliest event:     {r.earliest_event.strftime('%Y-%m-%d')}")
            lines.append(f"  Latest event:       {r.latest_event.strftime('%Y-%m-%d')}")
            span = (r.latest_event - r.earliest_event).days
            lines.append(f"  Span:               {span:,} days ({span // 365} years, {(span % 365) // 30} months)")
            lines.append("")
        
        # Attendees
        lines.append("ATTENDEE INFORMATION")
        lines.append("-" * 40)
        lines.append(f"  Events with attendees:  {r.events_with_attendees:,} ({r.events_with_attendees / r.total_events * 100:.1f}% of events)")
        lines.append(f"  Total attendee entries: {r.total_attendees:,}")
        lines.append(f"  Unique attendees:       {len(r.unique_attendees):,}")
        if r.attendee_response_stats:
            lines.append("  Response status breakdown:")
            for status, count in sorted(r.attendee_response_stats.items(), key=lambda x: -x[1]):
                lines.append(f"    - {status}: {count:,}")
        lines.append("")
        
        # Organizers
        lines.append("ORGANIZER INFORMATION")
        lines.append("-" * 40)
        lines.append(f"  Events with organizer:  {r.events_with_organizer:,}")
        lines.append(f"  Unique organizers:      {len(r.unique_organizers):,}")
        if r.unique_organizers and len(r.unique_organizers) <= 10:
            lines.append("  Organizers:")
            for org in sorted(r.unique_organizers):
                lines.append(f"    - {org}")
        lines.append("")
        
        # Recurrence
        if r.recurring_events > 0:
            lines.append("RECURRENCE INFORMATION")
            lines.append("-" * 40)
            lines.append(f"  Recurring events:       {r.recurring_events:,}")
            lines.append(f"  With exceptions:        {r.events_with_exceptions:,}")
            if r.recurrence_types:
                lines.append("  Recurrence types:")
                for rtype, count in sorted(r.recurrence_types.items(), key=lambda x: -x[1]):
                    lines.append(f"    - {rtype}: {count:,}")
            lines.append("")
        
        # Reminders
        lines.append("REMINDER/ALARM INFORMATION")
        lines.append("-" * 40)
        lines.append(f"  Events with reminders:  {r.events_with_reminders:,} ({r.events_with_reminders / r.total_events * 100:.1f}% of events)")
        if r.reminder_types:
            lines.append("  Reminder types:")
            for rtype, count in sorted(r.reminder_types.items(), key=lambda x: -x[1]):
                lines.append(f"    - {rtype}: {count:,}")
        lines.append("")
        
        # Timezones
        if r.timezones_found:
            lines.append("TIMEZONES DETECTED")
            lines.append("-" * 40)
            for tz_name, count in sorted(r.timezones_found.items(), key=lambda x: -x[1]):
                lines.append(f"  - {tz_name}: {count:,} events")
            lines.append("")
        
        # Field presence
        lines.append("FIELD PRESENCE (what data is available)")
        lines.append("-" * 40)
        
        # Sort fields by count
        sorted_fields = sorted(r.fields.items(), key=lambda x: -x[1].count)
        
        for field_name, stats in sorted_fields:
            pct = stats.count / r.total_events * 100
            bar = "█" * int(pct / 5) + "░" * (20 - int(pct / 5))
            lines.append(f"  {field_name:30s} {stats.count:>8,} ({pct:5.1f}%) {bar}")
        lines.append("")
        
        # Sample values for key fields
        lines.append("SAMPLE VALUES (first few unique values per field)")
        lines.append("-" * 40)
        
        key_fields = ['summary', 'location', 'categories', 'organizer', 'status', 'transp', 'class']
        for field_name in key_fields:
            if field_name in r.fields and r.fields[field_name].sample_values:
                lines.append(f"\n  {field_name.upper()}:")
                for val in r.fields[field_name].sample_values[:3]:
                    lines.append(f"    • {val}")
        lines.append("")
        
        # Microsoft/Outlook specific fields
        ms_fields = [f for f in r.fields.keys() if f.startswith('x-microsoft') or f.startswith('x-alt')]
        if ms_fields:
            lines.append("MICROSOFT/OUTLOOK SPECIFIC FIELDS")
            lines.append("-" * 40)
            for field_name in sorted(ms_fields):
                stats = r.fields[field_name]
                pct = stats.count / r.total_events * 100
                lines.append(f"  {field_name}: {stats.count:,} events ({pct:.1f}%)")
            lines.append("")
        
        # Warnings and issues
        if r.warnings:
            lines.append("⚠️  WARNINGS")
            lines.append("-" * 40)
            for warning in r.warnings:
                lines.append(f"  • {warning}")
            lines.append("")
        
        if r.issues:
            lines.append("❌ ISSUES")
            lines.append("-" * 40)
            for issue in r.issues:
                lines.append(f"  • {issue}")
            lines.append("")
        
        # Import recommendations
        lines.append("IMPORT RECOMMENDATIONS")
        lines.append("-" * 40)
        
        if r.events_with_attendees > 0:
            lines.append(f"  ✓ {r.events_with_attendees:,} events have attendees - they will be imported without sending notifications")
        
        if r.events_with_reminders > 0:
            lines.append(f"  ✓ {r.events_with_reminders:,} events have reminders - up to 5 per event will be imported")
        
        if r.recurring_events > 0:
            lines.append(f"  ✓ {r.recurring_events:,} recurring events detected - recurrence rules will be preserved")
        
        if r.events_with_exceptions > 0:
            lines.append(f"  ✓ {r.events_with_exceptions:,} recurring events have exceptions (EXDATE) - these will be imported")
        
        estimated_time = r.total_events / 5  # ~5 events per second
        if estimated_time > 60:
            lines.append(f"  ⏱ Estimated import time: ~{estimated_time / 60:.0f} minutes")
        else:
            lines.append(f"  ⏱ Estimated import time: ~{estimated_time:.0f} seconds")
        
        lines.append("")
        
        # Sample events
        if r.sample_events:
            lines.append("SAMPLE EVENTS (first few events)")
            lines.append("-" * 40)
            for i, event in enumerate(r.sample_events[:5], 1):
                lines.append(f"\n  Event {i}:")
                lines.append(f"    Title:     {event.get('summary', '(no title)')}")
                lines.append(f"    Type:      {event.get('_type', 'unknown')}")
                if 'dtstart' in event:
                    lines.append(f"    Start:     {event['dtstart']}")
                if 'location' in event:
                    lines.append(f"    Location:  {event['location']}")
                if event.get('_has_attendees'):
                    lines.append(f"    Attendees: {event['_has_attendees']}")
                if event.get('_has_reminders'):
                    lines.append(f"    Reminders: Yes")
                if 'rrule' in event:
                    lines.append(f"    Recurring: {event['rrule'][:50]}...")
        
        lines.append("")
        lines.append("=" * 70)
        lines.append("END OF REPORT")
        lines.append("=" * 70)
        
        return "\n".join(lines)
    
    def generate_json_report(self) -> Dict[str, Any]:
        """Generate a JSON-serializable report"""
        r = self.report
        
        return {
            "summary": {
                "files_analyzed": r.file_count,
                "total_size_mb": round(r.total_size_bytes / 1024 / 1024, 2),
                "total_events": r.total_events,
            },
            "event_types": {
                "timed": r.timed_events,
                "all_day": r.all_day_events,
                "recurring": r.recurring_events,
                "cancelled": r.cancelled_events,
            },
            "date_range": {
                "earliest": r.earliest_event.isoformat() if r.earliest_event else None,
                "latest": r.latest_event.isoformat() if r.latest_event else None,
            },
            "attendees": {
                "events_with_attendees": r.events_with_attendees,
                "total_attendees": r.total_attendees,
                "unique_attendees": len(r.unique_attendees),
                "response_stats": dict(r.attendee_response_stats),
            },
            "organizers": {
                "events_with_organizer": r.events_with_organizer,
                "unique_organizers": len(r.unique_organizers),
                "organizer_list": list(r.unique_organizers)[:20],
            },
            "recurrence": {
                "recurring_events": r.recurring_events,
                "with_exceptions": r.events_with_exceptions,
                "types": dict(r.recurrence_types),
            },
            "reminders": {
                "events_with_reminders": r.events_with_reminders,
                "types": dict(r.reminder_types),
            },
            "timezones": dict(r.timezones_found),
            "fields": {
                name: {
                    "count": stats.count,
                    "percentage": round(stats.count / r.total_events * 100, 1) if r.total_events > 0 else 0,
                    "sample_values": stats.sample_values[:3],
                }
                for name, stats in r.fields.items()
            },
            "warnings": r.warnings,
            "issues": r.issues,
        }


def find_ics_files(path: str) -> List[str]:
    """Find all ICS files in a path"""
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
        description='Analyze ICS calendar files before importing',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s calendar.ics
  %(prog)s ./ics_files/
  %(prog)s calendar.ics --show-samples 10
  %(prog)s calendar.ics --output report.txt
  %(prog)s calendar.ics --json
        """
    )
    
    parser.add_argument('ics_path', help='Path to ICS file or directory')
    parser.add_argument('--show-samples', '-s', type=int, default=5,
                        help='Number of sample events to show (default: 5)')
    parser.add_argument('--output', '-o', help='Save report to file')
    parser.add_argument('--json', '-j', action='store_true',
                        help='Output report as JSON')
    
    args = parser.parse_args()
    
    # Find ICS files
    ics_files = find_ics_files(args.ics_path)
    
    if not ics_files:
        print("No ICS files found")
        sys.exit(1)
    
    print(f"\nFound {len(ics_files)} ICS file(s) to analyze\n")
    
    # Analyze
    analyzer = ICSAnalyzer()
    
    for ics_file in ics_files:
        analyzer.analyze_file(ics_file, max_samples=args.show_samples)
    
    # Generate report
    if args.json:
        report = json.dumps(analyzer.generate_json_report(), indent=2)
    else:
        report = analyzer.generate_report()
    
    # Output
    if args.output:
        with open(args.output, 'w') as f:
            f.write(report)
        print(f"\nReport saved to: {args.output}")
    else:
        print("\n")
        print(report)


if __name__ == '__main__':
    main()
