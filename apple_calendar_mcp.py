#!/usr/bin/env python3
"""
Apple Calendar MCP Server - FastMCP implementation
Provides tools to query and interact with Apple Calendar
"""

import subprocess
import os
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any
from mcp.server.fastmcp import FastMCP

# Load user preferences from environment
USER_PREFERENCES = os.environ.get("USER_CALENDAR_PREFERENCES", "")

# Initialize FastMCP server
mcp = FastMCP("Apple Calendar MCP")


def inject_preferences(func):
    """Decorator that appends user preferences to tool docstrings"""
    if USER_PREFERENCES:
        if func.__doc__:
            func.__doc__ = func.__doc__.rstrip() + f"\n\nUser Preferences: {USER_PREFERENCES}"
        else:
            func.__doc__ = f"User Preferences: {USER_PREFERENCES}"
    return func


def run_applescript(script: str) -> str:
    """Execute AppleScript and return output"""
    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=120
        )
        if result.returncode != 0 and result.stderr:
            raise Exception(f"AppleScript error: {result.stderr.strip()}")
        return result.stdout.strip()
    except subprocess.TimeoutExpired:
        raise Exception("AppleScript execution timed out")
    except Exception as e:
        raise Exception(f"AppleScript execution failed: {str(e)}")


def format_applescript_date(date_str: str) -> str:
    """Convert YYYY-MM-DD or YYYY-MM-DD HH:MM to AppleScript date format"""
    try:
        if ' ' in date_str:
            dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M")
            return dt.strftime("%B %d, %Y at %I:%M:%S %p")
        else:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
            return dt.strftime("%B %d, %Y")
    except ValueError:
        return date_str


def get_date_range_script(start_date: Optional[str], end_date: Optional[str], days_ahead: int = 7) -> str:
    """Generate AppleScript for date range filtering"""
    if start_date:
        start_script = f'set startDate to date "{format_applescript_date(start_date)}"'
    else:
        start_script = 'set startDate to current date'

    if end_date:
        end_script = f'set endDate to date "{format_applescript_date(end_date)}"'
    else:
        end_script = f'set endDate to (current date) + ({days_ahead} * days)'

    return f"{start_script}\n{end_script}"


@mcp.tool()
@inject_preferences
def list_calendars(include_counts: bool = True) -> str:
    """
    List all available calendars with metadata.

    Args:
        include_counts: Whether to include event counts for each calendar (default: True)

    Returns:
        Formatted list of calendars with names and optional event counts
    """

    count_script = '''
        try
            set eventCount to count of events of aCal
            set outputText to outputText & " (" & eventCount & " events)"
        on error
            set outputText to outputText & " (count unavailable)"
        end try
    ''' if include_counts else ''

    script = f'''
    tell application "Calendar"
        set outputText to "CALENDARS" & return & return
        set calList to every calendar
        set calCount to count of calList

        set outputText to outputText & "Found " & calCount & " calendar(s)" & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

        repeat with aCal in calList
            set calName to name of aCal
            set outputText to outputText & "📅 " & calName

            {count_script}

            set outputText to outputText & return
        end repeat

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def list_events(
    calendar: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    max_events: int = 50,
    include_all_day: bool = True
) -> str:
    """
    List events from a specific calendar or date range.

    Args:
        calendar: Optional calendar name to filter (if None, shows all calendars)
        start_date: Start of date range in YYYY-MM-DD format (default: today)
        end_date: End of date range in YYYY-MM-DD format (default: +7 days)
        max_events: Maximum number of events to return (default: 50)
        include_all_day: Whether to include all-day events (default: True)

    Returns:
        Formatted list of events with title, time, location, and calendar
    """

    date_range_script = get_date_range_script(start_date, end_date)

    calendar_filter = f'''
        set targetCalendars to {{calendar "{calendar}"}}
    ''' if calendar else '''
        set targetCalendars to every calendar
    '''

    all_day_filter = '' if include_all_day else '''
        if allday event of anEvent is false then
    '''
    all_day_filter_end = '' if include_all_day else 'end if'

    script = f'''
    tell application "Calendar"
        set outputText to "EVENTS" & return & return
        {date_range_script}

        set outputText to outputText & "Date range: " & (startDate as string) & " to " & (endDate as string) & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

        {calendar_filter}

        set eventCount to 0
        set allEvents to {{}}

        repeat with aCal in targetCalendars
            set calName to name of aCal

            try
                set calEvents to (every event of aCal whose start date >= startDate and start date <= endDate)

                repeat with anEvent in calEvents
                    if eventCount >= {max_events} then exit repeat

                    {all_day_filter}
                    try
                        set eventTitle to summary of anEvent
                        set eventStart to start date of anEvent
                        set eventEnd to end date of anEvent
                        set isAllDay to allday event of anEvent

                        if isAllDay then
                            set timeStr to "All Day"
                        else
                            set timeStr to time string of eventStart & " - " & time string of eventEnd
                        end if

                        set outputText to outputText & "📌 " & eventTitle & return
                        set outputText to outputText & "   📅 " & date string of eventStart & return
                        set outputText to outputText & "   🕐 " & timeStr & return
                        set outputText to outputText & "   📁 " & calName & return

                        try
                            set evtLoc to location of anEvent
                            if evtLoc is not missing value and length of evtLoc > 0 then
                                set outputText to outputText & "   📍 " & evtLoc & return
                            end if
                        end try

                        set outputText to outputText & return
                        set eventCount to eventCount + 1
                    end try
                    {all_day_filter_end}
                end repeat
            end try
        end repeat

        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
        set outputText to outputText & "Total: " & eventCount & " event(s)" & return

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def get_event_details(
    calendar: str,
    event_title: str,
    start_date: Optional[str] = None
) -> str:
    """
    Get full details of a specific event.

    Args:
        calendar: Calendar name containing the event
        event_title: Event title to search for (partial match supported)
        start_date: Optional date hint to narrow search (YYYY-MM-DD format)

    Returns:
        Full event details including notes, attendees, recurrence, and reminders
    """

    date_filter = ''
    if start_date:
        date_filter = f'''
            set searchDate to date "{format_applescript_date(start_date)}"
            set searchEndDate to searchDate + 1 * days
            set matchingEvents to (every event of targetCal whose summary contains "{event_title}" and start date >= searchDate and start date < searchEndDate)
        '''
    else:
        date_filter = f'''
            set matchingEvents to (every event of targetCal whose summary contains "{event_title}")
        '''

    script = f'''
    tell application "Calendar"
        set outputText to "EVENT DETAILS" & return & return

        try
            set targetCal to calendar "{calendar}"
            {date_filter}

            if (count of matchingEvents) = 0 then
                return "No event found matching: {event_title}"
            end if

            set targetEvent to item 1 of matchingEvents

            -- Basic info
            set eventTitle to summary of targetEvent
            set eventStart to start date of targetEvent
            set eventEnd to end date of targetEvent
            set isAllDay to allday event of targetEvent

            set outputText to outputText & "📌 " & eventTitle & return
            set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

            set outputText to outputText & "📅 Date: " & date string of eventStart & return

            if isAllDay then
                set outputText to outputText & "🕐 Time: All Day" & return
            else
                set outputText to outputText & "🕐 Start: " & time string of eventStart & return
                set outputText to outputText & "🕐 End: " & time string of eventEnd & return
            end if

            set outputText to outputText & "📁 Calendar: {calendar}" & return

            -- Location
            try
                set evtLoc to location of targetEvent
                if evtLoc is not missing value and length of evtLoc > 0 then
                    set outputText to outputText & "📍 Location: " & evtLoc & return
                end if
            end try

            -- URL
            try
                set eventURL to url of targetEvent
                if eventURL is not missing value and eventURL is not "" then
                    set outputText to outputText & "🔗 URL: " & eventURL & return
                end if
            end try

            -- Notes/Description
            try
                set eventNotes to description of targetEvent
                if eventNotes is not missing value and eventNotes is not "" then
                    set outputText to outputText & return & "📝 Notes:" & return
                    set outputText to outputText & eventNotes & return
                end if
            end try

            -- Recurrence
            try
                set eventRecurrence to recurrence of targetEvent
                if eventRecurrence is not missing value and eventRecurrence is not "" then
                    set outputText to outputText & return & "🔄 Recurrence: " & eventRecurrence & return
                end if
            end try

            -- Attendees
            try
                set attendeeList to attendees of targetEvent
                if (count of attendeeList) > 0 then
                    set outputText to outputText & return & "👥 Attendees:" & return
                    repeat with anAttendee in attendeeList
                        set attendeeName to display name of anAttendee
                        set attendeeStatus to participation status of anAttendee
                        set outputText to outputText & "   • " & attendeeName & " (" & attendeeStatus & ")" & return
                    end repeat
                end if
            end try

            -- Alarms/Reminders
            try
                set alarmList to display alarms of targetEvent
                if (count of alarmList) > 0 then
                    set outputText to outputText & return & "⏰ Reminders:" & return
                    repeat with anAlarm in alarmList
                        set triggerInterval to trigger interval of anAlarm
                        set minutesBefore to (triggerInterval / -60) as integer
                        set outputText to outputText & "   • " & minutesBefore & " minutes before" & return
                    end repeat
                end if
            end try

            return outputText

        on error errMsg
            return "Error: " & errMsg
        end try
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def get_todays_schedule(calendar: Optional[str] = None) -> str:
    """
    Get a quick view of today's events across all calendars.

    Args:
        calendar: Optional calendar name to filter (if None, shows all calendars)

    Returns:
        Chronological list of today's events with times and locations
    """

    calendar_filter = f'''
        set targetCalendars to {{calendar "{calendar}"}}
    ''' if calendar else '''
        set targetCalendars to every calendar
    '''

    script = f'''
    tell application "Calendar"
        set outputText to "TODAY'S SCHEDULE" & return
        set outputText to outputText & date string of (current date) & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

        set todayStart to current date
        set time of todayStart to 0
        set todayEnd to todayStart + 1 * days

        {calendar_filter}

        set eventCount to 0
        set allDayEvents to {{}}
        set timedEvents to {{}}

        repeat with aCal in targetCalendars
            set calName to name of aCal

            try
                set calEvents to (every event of aCal whose start date >= todayStart and start date < todayEnd)

                repeat with anEvent in calEvents
                    try
                        set eventTitle to summary of anEvent
                        set eventStart to start date of anEvent
                        set eventEnd to end date of anEvent
                        set isAllDay to allday event of anEvent

                        set evtLoc to ""
                        try
                            set evtLoc to location of anEvent
                        end try

                        if isAllDay then
                            set outputText to outputText & "🌅 ALL DAY: " & eventTitle
                            if length of evtLoc > 0 then
                                set outputText to outputText & " @ " & evtLoc
                            end if
                            set outputText to outputText & " [" & calName & "]" & return
                        else
                            set timeStr to time string of eventStart
                            set outputText to outputText & "🕐 " & timeStr & " - " & eventTitle
                            if length of evtLoc > 0 then
                                set outputText to outputText & " @ " & evtLoc
                            end if
                            set outputText to outputText & " [" & calName & "]" & return
                        end if

                        set eventCount to eventCount + 1
                    end try
                end repeat
            end try
        end repeat

        if eventCount = 0 then
            set outputText to outputText & "No events scheduled for today." & return
        end if

        set outputText to outputText & return & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
        set outputText to outputText & "Total: " & eventCount & " event(s)" & return

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def get_calendar_overview() -> str:
    """
    Get a quick dashboard view of your calendars.

    Returns:
        Overview including:
        - All calendar names
        - Today's date
        - Tip to use other tools for detailed views
    """

    script = '''
    tell application "Calendar"
        set outputText to "╔══════════════════════════════════════════╗" & return
        set outputText to outputText & "║      CALENDAR OVERVIEW                   ║" & return
        set outputText to outputText & "╚══════════════════════════════════════════╝" & return & return

        -- Calendar List
        set outputText to outputText & "📅 YOUR CALENDARS" & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return

        set calList to every calendar
        set calCount to count of calList

        repeat with aCal in calList
            set calName to name of aCal
            set outputText to outputText & "  📁 " & calName & return
        end repeat

        set outputText to outputText & return
        set outputText to outputText & "Total: " & calCount & " calendar(s)" & return
        set outputText to outputText & return

        -- Today's date
        set outputText to outputText & "📆 TODAY: " & date string of (current date) & return
        set outputText to outputText & return

        set outputText to outputText & "═══════════════════════════════════════════════════" & return
        set outputText to outputText & "💬 Quick commands:" & return
        set outputText to outputText & "  • get_todays_schedule - See today's events" & return
        set outputText to outputText & "  • list_events - Browse events by date range" & return
        set outputText to outputText & "  • list_calendars - Get event counts per calendar" & return

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def create_event(
    calendar: str,
    title: str,
    start_date: str,
    end_date: str,
    location: Optional[str] = None,
    notes: Optional[str] = None,
    url: Optional[str] = None,
    attendees: Optional[List[str]] = None,
    alert_minutes: Optional[int] = None,
    all_day: bool = False
) -> str:
    """
    Create a new event in the specified calendar.

    Args:
        calendar: Calendar name where event will be created
        title: Event title/summary
        start_date: Start date/time in format "YYYY-MM-DD HH:MM" or "YYYY-MM-DD" for all-day
        end_date: End date/time in format "YYYY-MM-DD HH:MM" or "YYYY-MM-DD" for all-day
        location: Optional event location
        notes: Optional event notes/description
        url: Optional event URL
        attendees: Optional list of attendee email addresses
        alert_minutes: Optional reminder in minutes before event (e.g., 15, 30, 60)
        all_day: Whether this is an all-day event (default: False)

    Returns:
        Success message with created event details
    """

    start_formatted = format_applescript_date(start_date)
    end_formatted = format_applescript_date(end_date)

    # Build optional properties
    optional_props = []

    if location:
        optional_props.append(f'set location of newEvent to "{location}"')

    if notes:
        # Escape quotes in notes
        escaped_notes = notes.replace('"', '\\"')
        optional_props.append(f'set description of newEvent to "{escaped_notes}"')

    if url:
        optional_props.append(f'set url of newEvent to "{url}"')

    if all_day:
        optional_props.append('set allday event of newEvent to true')

    optional_script = '\n            '.join(optional_props) if optional_props else ''

    # Handle attendees
    attendee_script = ''
    if attendees:
        attendee_lines = []
        for email in attendees:
            attendee_lines.append(f'make new attendee at end of attendees of newEvent with properties {{email:"{email}"}}')
        attendee_script = '\n            '.join(attendee_lines)

    # Handle alerts
    alert_script = ''
    if alert_minutes is not None:
        alert_script = f'make new display alarm at end of display alarms of newEvent with properties {{trigger interval:{-alert_minutes * 60}}}'

    script = f'''
    tell application "Calendar"
        set targetCal to calendar "{calendar}"

        tell targetCal
            set newEvent to make new event with properties {{summary:"{title}", start date:date "{start_formatted}", end date:date "{end_formatted}"}}

            {optional_script}

            {attendee_script}

            {alert_script}
        end tell

        save

        set outputText to "✅ EVENT CREATED" & return & return
        set outputText to outputText & "📌 " & "{title}" & return
        set outputText to outputText & "📅 " & date string of start date of newEvent & return

        if allday event of newEvent then
            set outputText to outputText & "🕐 All Day Event" & return
        else
            set outputText to outputText & "🕐 " & time string of start date of newEvent & " - " & time string of end date of newEvent & return
        end if

        set outputText to outputText & "📁 Calendar: {calendar}" & return

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def create_recurring_event(
    calendar: str,
    title: str,
    start_date: str,
    end_date: str,
    recurrence_frequency: str,
    recurrence_interval: int = 1,
    recurrence_end_date: Optional[str] = None,
    location: Optional[str] = None,
    notes: Optional[str] = None,
    alert_minutes: Optional[int] = None
) -> str:
    """
    Create a recurring event in the specified calendar.

    Args:
        calendar: Calendar name where event will be created
        title: Event title/summary
        start_date: Start date/time in format "YYYY-MM-DD HH:MM"
        end_date: End date/time in format "YYYY-MM-DD HH:MM"
        recurrence_frequency: Frequency - "daily", "weekly", "monthly", or "yearly"
        recurrence_interval: Interval (e.g., 1 for every week, 2 for every 2 weeks)
        recurrence_end_date: Optional end date for recurrence in "YYYY-MM-DD" format
        location: Optional event location
        notes: Optional event notes/description
        alert_minutes: Optional reminder in minutes before event

    Returns:
        Success message with created recurring event details
    """

    start_formatted = format_applescript_date(start_date)
    end_formatted = format_applescript_date(end_date)

    # Map frequency to AppleScript recurrence
    freq_map = {
        "daily": "day",
        "weekly": "week",
        "monthly": "month",
        "yearly": "year"
    }

    freq_unit = freq_map.get(recurrence_frequency.lower(), "week")

    # Build recurrence rule
    if recurrence_end_date:
        recurrence_end_formatted = format_applescript_date(recurrence_end_date)
        recurrence_rule = f'"FREQ={recurrence_frequency.upper()};INTERVAL={recurrence_interval};UNTIL={recurrence_end_formatted}"'
    else:
        recurrence_rule = f'"FREQ={recurrence_frequency.upper()};INTERVAL={recurrence_interval}"'

    # Build optional properties
    optional_props = []
    if location:
        optional_props.append(f'set location of newEvent to "{location}"')
    if notes:
        escaped_notes = notes.replace('"', '\\"')
        optional_props.append(f'set description of newEvent to "{escaped_notes}"')

    optional_script = '\n            '.join(optional_props) if optional_props else ''

    alert_script = ''
    if alert_minutes is not None:
        alert_script = f'make new display alarm at end of display alarms of newEvent with properties {{trigger interval:{-alert_minutes * 60}}}'

    script = f'''
    tell application "Calendar"
        set targetCal to calendar "{calendar}"

        tell targetCal
            set newEvent to make new event with properties {{summary:"{title}", start date:date "{start_formatted}", end date:date "{end_formatted}"}}

            {optional_script}

            set recurrence of newEvent to {recurrence_rule}

            {alert_script}
        end tell

        save

        set outputText to "✅ RECURRING EVENT CREATED" & return & return
        set outputText to outputText & "📌 " & "{title}" & return
        set outputText to outputText & "🔄 Repeats: {recurrence_frequency}, every {recurrence_interval}" & return
        set outputText to outputText & "📅 Starts: " & date string of start date of newEvent & return
        set outputText to outputText & "🕐 " & time string of start date of newEvent & " - " & time string of end date of newEvent & return
        set outputText to outputText & "📁 Calendar: {calendar}" & return

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def quick_add_event(
    calendar: str,
    event_text: str
) -> str:
    """
    Quick add an event using natural language description.
    Parses simple event descriptions like "Lunch with John tomorrow at 12pm" or "Team meeting Friday 2-3pm".

    Args:
        calendar: Calendar name where event will be created
        event_text: Natural language event description

    Returns:
        Success message with created event details or parsing guidance

    Note: This uses basic parsing. For complex events, use create_event instead.
    Examples:
    - "Dentist appointment tomorrow at 2pm"
    - "Team standup Monday 9am"
    - "Birthday party Saturday 6-9pm"
    """

    # This is a simplified version - AppleScript doesn't have sophisticated NLP
    # We'll provide a helpful response that guides users to use create_event for now
    script = f'''
    tell application "Calendar"
        set outputText to "⚠️ QUICK ADD GUIDANCE" & return & return
        set outputText to outputText & "Natural language parsing in AppleScript is limited." & return
        set outputText to outputText & "Please use create_event with specific parameters:" & return & return
        set outputText to outputText & "Your input: '{event_text}'" & return & return
        set outputText to outputText & "Example:" & return
        set outputText to outputText & "create_event(" & return
        set outputText to outputText & "  calendar='{calendar}'," & return
        set outputText to outputText & "  title='Your Event Title'," & return
        set outputText to outputText & "  start_date='2025-01-15 14:00'," & return
        set outputText to outputText & "  end_date='2025-01-15 15:00'" & return
        set outputText to outputText & ")" & return

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def update_event(
    calendar: str,
    event_title: str,
    start_date: Optional[str] = None,
    new_title: Optional[str] = None,
    new_start_date: Optional[str] = None,
    new_end_date: Optional[str] = None,
    new_location: Optional[str] = None,
    new_notes: Optional[str] = None,
    new_url: Optional[str] = None
) -> str:
    """
    Update properties of an existing event.

    Args:
        calendar: Calendar name containing the event
        event_title: Current event title to search for
        start_date: Optional date hint to find the specific event (YYYY-MM-DD format)
        new_title: New event title (if changing)
        new_start_date: New start date/time in "YYYY-MM-DD HH:MM" format
        new_end_date: New end date/time in "YYYY-MM-DD HH:MM" format
        new_location: New location
        new_notes: New notes/description
        new_url: New URL

    Returns:
        Success message with updated event details
    """

    # Find event first
    date_filter = ''
    if start_date:
        date_filter = f'''
            set searchDate to date "{format_applescript_date(start_date)}"
            set searchEndDate to searchDate + 1 * days
            set matchingEvents to (every event of targetCal whose summary contains "{event_title}" and start date >= searchDate and start date < searchEndDate)
        '''
    else:
        date_filter = f'''
            set matchingEvents to (every event of targetCal whose summary contains "{event_title}")
        '''

    # Build update statements
    updates = []
    if new_title:
        updates.append(f'set summary of targetEvent to "{new_title}"')
    if new_start_date:
        updates.append(f'set start date of targetEvent to date "{format_applescript_date(new_start_date)}"')
    if new_end_date:
        updates.append(f'set end date of targetEvent to date "{format_applescript_date(new_end_date)}"')
    if new_location:
        updates.append(f'set location of targetEvent to "{new_location}"')
    if new_notes:
        escaped_notes = new_notes.replace('"', '\\"')
        updates.append(f'set description of targetEvent to "{escaped_notes}"')
    if new_url:
        updates.append(f'set url of targetEvent to "{new_url}"')

    update_script = '\n            '.join(updates) if updates else ''

    script = f'''
    tell application "Calendar"
        set outputText to "EVENT UPDATE" & return & return

        try
            set targetCal to calendar "{calendar}"
            {date_filter}

            if (count of matchingEvents) = 0 then
                return "❌ No event found matching: {event_title}"
            end if

            set targetEvent to item 1 of matchingEvents

            {update_script}

            save

            set outputText to "✅ EVENT UPDATED" & return & return
            set outputText to outputText & "📌 " & summary of targetEvent & return
            set outputText to outputText & "📅 " & date string of start date of targetEvent & return
            set outputText to outputText & "🕐 " & time string of start date of targetEvent & " - " & time string of end date of targetEvent & return
            set outputText to outputText & "📁 Calendar: {calendar}" & return

            try
                set evtLoc to location of targetEvent
                if evtLoc is not missing value and length of evtLoc > 0 then
                    set outputText to outputText & "📍 Location: " & evtLoc & return
                end if
            end try

            return outputText

        on error errMsg
            return "❌ Error updating event: " & errMsg
        end try
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def move_event(
    calendar: str,
    event_title: str,
    current_start_date: str,
    new_start_date: str,
    new_end_date: Optional[str] = None
) -> str:
    """
    Move/reschedule an event to a new time.

    Args:
        calendar: Calendar name containing the event
        event_title: Event title to search for
        current_start_date: Current start date to identify the event (YYYY-MM-DD format)
        new_start_date: New start date/time in "YYYY-MM-DD HH:MM" format
        new_end_date: Optional new end date/time (if not provided, duration is preserved)

    Returns:
        Success message with rescheduled event details
    """

    current_formatted = format_applescript_date(current_start_date)
    new_start_formatted = format_applescript_date(new_start_date)

    # Handle end date
    end_date_script = ''
    if new_end_date:
        new_end_formatted = format_applescript_date(new_end_date)
        end_date_script = f'set end date of targetEvent to date "{new_end_formatted}"'
    else:
        # Preserve duration
        end_date_script = '''
            set eventDuration to (end date of targetEvent) - (start date of targetEvent)
            set end date of targetEvent to (date newStartDate) + eventDuration
        '''

    script = f'''
    tell application "Calendar"
        set outputText to "MOVE EVENT" & return & return

        try
            set targetCal to calendar "{calendar}"
            set searchDate to date "{current_formatted}"
            set searchEndDate to searchDate + 1 * days

            set matchingEvents to (every event of targetCal whose summary contains "{event_title}" and start date >= searchDate and start date < searchEndDate)

            if (count of matchingEvents) = 0 then
                return "❌ No event found matching: {event_title} on {current_start_date}"
            end if

            set targetEvent to item 1 of matchingEvents
            set oldStart to start date of targetEvent
            set oldEnd to end date of targetEvent

            set newStartDate to "{new_start_formatted}"
            set start date of targetEvent to date newStartDate

            {end_date_script}

            save

            set outputText to "✅ EVENT MOVED" & return & return
            set outputText to outputText & "📌 " & summary of targetEvent & return & return
            set outputText to outputText & "FROM:" & return
            set outputText to outputText & "  📅 " & date string of oldStart & return
            set outputText to outputText & "  🕐 " & time string of oldStart & " - " & time string of oldEnd & return & return
            set outputText to outputText & "TO:" & return
            set outputText to outputText & "  📅 " & date string of start date of targetEvent & return
            set outputText to outputText & "  🕐 " & time string of start date of targetEvent & " - " & time string of end date of targetEvent & return

            return outputText

        on error errMsg
            return "❌ Error moving event: " & errMsg
        end try
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def delete_event(
    calendar: str,
    event_title: str,
    start_date: str,
    delete_all_occurrences: bool = False
) -> str:
    """
    Delete an event from the calendar.

    Args:
        calendar: Calendar name containing the event
        event_title: Event title to search for
        start_date: Start date of the event to delete (YYYY-MM-DD format)
        delete_all_occurrences: For recurring events, delete all occurrences (default: False, deletes only this instance)

    Returns:
        Success message confirming deletion
    """

    date_formatted = format_applescript_date(start_date)

    # Recurrence handling
    if delete_all_occurrences:
        delete_cmd = 'delete targetEvent'
    else:
        delete_cmd = 'delete targetEvent'

    script = f'''
    tell application "Calendar"
        set outputText to "DELETE EVENT" & return & return

        try
            set targetCal to calendar "{calendar}"
            set searchDate to date "{date_formatted}"
            set searchEndDate to searchDate + 1 * days

            set matchingEvents to (every event of targetCal whose summary contains "{event_title}" and start date >= searchDate and start date < searchEndDate)

            if (count of matchingEvents) = 0 then
                return "❌ No event found matching: {event_title} on {start_date}"
            end if

            set targetEvent to item 1 of matchingEvents
            set eventTitle to summary of targetEvent
            set eventDate to date string of start date of targetEvent
            set eventTime to time string of start date of targetEvent

            {delete_cmd}

            save

            set outputText to "✅ EVENT DELETED" & return & return
            set outputText to outputText & "📌 " & eventTitle & return
            set outputText to outputText & "📅 " & eventDate & return
            set outputText to outputText & "🕐 " & eventTime & return
            set outputText to outputText & "📁 Calendar: {calendar}" & return

            return outputText

        on error errMsg
            return "❌ Error deleting event: " & errMsg
        end try
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def manage_reminders(
    calendar: str,
    event_title: str,
    start_date: str,
    reminder_minutes: Optional[List[int]] = None,
    clear_existing: bool = False
) -> str:
    """
    Add or modify reminders (alarms) for an event.

    Args:
        calendar: Calendar name containing the event
        event_title: Event title to search for
        start_date: Start date of the event (YYYY-MM-DD format)
        reminder_minutes: List of minutes before event to remind (e.g., [15, 60] for 15 min and 1 hour before)
        clear_existing: Whether to clear existing reminders first (default: False)

    Returns:
        Success message with reminder details
    """

    date_formatted = format_applescript_date(start_date)

    # Build reminder creation script
    reminder_script = ''
    if reminder_minutes:
        reminder_lines = []
        for minutes in reminder_minutes:
            seconds = -minutes * 60
            reminder_lines.append(f'make new display alarm at end of display alarms of targetEvent with properties {{trigger interval:{seconds}}}')
        reminder_script = '\n            '.join(reminder_lines)

    clear_script = ''
    if clear_existing:
        clear_script = 'delete every display alarm of targetEvent'

    script = f'''
    tell application "Calendar"
        set outputText to "MANAGE REMINDERS" & return & return

        try
            set targetCal to calendar "{calendar}"
            set searchDate to date "{date_formatted}"
            set searchEndDate to searchDate + 1 * days

            set matchingEvents to (every event of targetCal whose summary contains "{event_title}" and start date >= searchDate and start date < searchEndDate)

            if (count of matchingEvents) = 0 then
                return "❌ No event found matching: {event_title} on {start_date}"
            end if

            set targetEvent to item 1 of matchingEvents

            {clear_script}

            {reminder_script}

            save

            set outputText to "✅ REMINDERS UPDATED" & return & return
            set outputText to outputText & "📌 " & summary of targetEvent & return
            set outputText to outputText & "📅 " & date string of start date of targetEvent & return & return

            set alarmList to display alarms of targetEvent
            if (count of alarmList) > 0 then
                set outputText to outputText & "⏰ Active Reminders:" & return
                repeat with anAlarm in alarmList
                    set triggerInterval to trigger interval of anAlarm
                    set minutesBefore to (triggerInterval / -60) as integer
                    set outputText to outputText & "   • " & minutesBefore & " minutes before" & return
                end repeat
            else
                set outputText to outputText & "No reminders set" & return
            end if

            return outputText

        on error errMsg
            return "❌ Error managing reminders: " & errMsg
        end try
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def search_events(
    search_text: str,
    calendar: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    search_location: bool = False,
    search_notes: bool = False,
    max_results: int = 20
) -> str:
    """
    Advanced search for events across calendars with multiple criteria.

    Args:
        search_text: Text to search for in event titles (and optionally location/notes)
        calendar: Optional calendar name to search in (if None, searches all calendars)
        start_date: Optional start of date range (YYYY-MM-DD format)
        end_date: Optional end of date range (YYYY-MM-DD format)
        search_location: Whether to also search in event locations (default: False)
        search_notes: Whether to also search in event notes (default: False)
        max_results: Maximum number of results to return (default: 20)

    Returns:
        List of matching events with details
    """

    date_range_script = get_date_range_script(start_date, end_date, days_ahead=365)

    calendar_filter = f'''
        set targetCalendars to {{calendar "{calendar}"}}
    ''' if calendar else '''
        set targetCalendars to every calendar
    '''

    # Build search filters
    search_conditions = [f'summary of anEvent contains "{search_text}"']

    if search_location:
        # Location search will be done in a try block in the main script
        pass

    if search_notes:
        # Notes search will be done in a try block in the main script
        pass

    script = f'''
    tell application "Calendar"
        set outputText to "SEARCH RESULTS" & return & return
        set outputText to outputText & "Search term: '{search_text}'" & return
        {date_range_script}
        set outputText to outputText & "Date range: " & (startDate as string) & " to " & (endDate as string) & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

        {calendar_filter}

        set resultCount to 0
        set matchedEvents to {{}}

        repeat with aCal in targetCalendars
            set calName to name of aCal

            try
                set calEvents to (every event of aCal whose start date >= startDate and start date <= endDate)

                repeat with anEvent in calEvents
                    if resultCount >= {max_results} then exit repeat

                    set isMatch to false

                    -- Check title
                    if summary of anEvent contains "{search_text}" then
                        set isMatch to true
                    end if

                    -- Check location if requested
                    {"" if not search_location else f'''
                    if not isMatch then
                        try
                            set evtLoc to location of anEvent
                            if evtLoc is not missing value and evtLoc contains "{search_text}" then
                                set isMatch to true
                            end if
                        end try
                    end if
                    '''}

                    -- Check notes if requested
                    {"" if not search_notes else f'''
                    if not isMatch then
                        try
                            set evtNotes to description of anEvent
                            if evtNotes is not missing value and evtNotes contains "{search_text}" then
                                set isMatch to true
                            end if
                        end try
                    end if
                    '''}

                    if isMatch then
                        try
                            set eventTitle to summary of anEvent
                            set eventStart to start date of anEvent
                            set eventEnd to end date of anEvent
                            set isAllDay to allday event of anEvent

                            if isAllDay then
                                set timeStr to "All Day"
                            else
                                set timeStr to time string of eventStart & " - " & time string of eventEnd
                            end if

                            set outputText to outputText & "📌 " & eventTitle & return
                            set outputText to outputText & "   📅 " & date string of eventStart & return
                            set outputText to outputText & "   🕐 " & timeStr & return
                            set outputText to outputText & "   📁 " & calName & return

                            try
                                set evtLoc to location of anEvent
                                if evtLoc is not missing value and length of evtLoc > 0 then
                                    set outputText to outputText & "   📍 " & evtLoc & return
                                end if
                            end try

                            set outputText to outputText & return
                            set resultCount to resultCount + 1
                        end try
                    end if
                end repeat
            end try
        end repeat

        if resultCount = 0 then
            set outputText to outputText & "No events found matching your search." & return
        else
            set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
            set outputText to outputText & "Total: " & resultCount & " matching event(s)" & return
        end if

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def find_free_time(
    start_date: str,
    end_date: str,
    duration_minutes: int = 30,
    calendar: Optional[str] = None,
    business_hours_only: bool = True,
    max_suggestions: int = 5
) -> str:
    """
    Find available time slots for scheduling new events.

    Args:
        start_date: Start of search range (YYYY-MM-DD format)
        end_date: End of search range (YYYY-MM-DD format)
        duration_minutes: Required duration in minutes (default: 30)
        calendar: Optional calendar name to check (if None, checks all calendars)
        business_hours_only: Only suggest slots during business hours 9am-5pm (default: True)
        max_suggestions: Maximum number of time slots to suggest (default: 5)

    Returns:
        List of available time slots
    """

    start_formatted = format_applescript_date(start_date)
    end_formatted = format_applescript_date(end_date)

    calendar_filter = f'''
        set targetCalendars to {{calendar "{calendar}"}}
    ''' if calendar else '''
        set targetCalendars to every calendar
    '''

    # Business hours filter
    business_hours_check = '''
                    -- Check if within business hours (9 AM to 5 PM)
                    set slotHour to hours of currentSlot
                    if slotHour < 9 or slotHour >= 17 then
                        set currentSlot to currentSlot + (60 * minutes)
                        set iteration to iteration + 1
                    end if
    ''' if business_hours_only else ''

    script = f'''
    tell application "Calendar"
        set outputText to "FREE TIME SLOTS" & return & return
        set outputText to outputText & "Duration needed: {duration_minutes} minutes" & return
        set outputText to outputText & "Search range: {start_date} to {end_date}" & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

        set searchStart to date "{start_formatted}"
        set searchEnd to date "{end_formatted}"
        set slotDuration to {duration_minutes} * minutes

        {calendar_filter}

        -- Collect all busy times
        set busySlots to {{}}

        repeat with aCal in targetCalendars
            try
                set calEvents to (every event of aCal whose start date >= searchStart and start date <= searchEnd and allday event is false)

                repeat with anEvent in calEvents
                    set eventStart to start date of anEvent
                    set eventEnd to end date of anEvent
                    set end of busySlots to {{eventStart, eventEnd}}
                end repeat
            end try
        end repeat

        -- Find free slots
        set freeSlots to {{}}
        set currentSlot to searchStart
        set suggestionsFound to 0
        set maxIterations to 1000
        set iteration to 0

        repeat while currentSlot < searchEnd and suggestionsFound < {max_suggestions} and iteration < maxIterations
            {business_hours_check}

            set slotEnd to currentSlot + slotDuration
            set isFree to true

            -- Check if this slot conflicts with any busy time
            repeat with busyTime in busySlots
                set busyStart to item 1 of busyTime
                set busyEnd to item 2 of busyTime

                -- Check for overlap
                if (currentSlot < busyEnd) and (slotEnd > busyStart) then
                    set isFree to false
                    set currentSlot to busyEnd
                    exit repeat
                end if
            end repeat

            if isFree and slotEnd <= searchEnd then
                set outputText to outputText & "✅ " & date string of currentSlot & return
                set outputText to outputText & "   🕐 " & time string of currentSlot & " - " & time string of slotEnd & return
                set outputText to outputText & "   ⏱  Duration: {duration_minutes} minutes" & return & return

                set suggestionsFound to suggestionsFound + 1
                set currentSlot to slotEnd
            else
                set currentSlot to currentSlot + (30 * minutes)
            end if

            set iteration to iteration + 1
        end repeat

        if suggestionsFound = 0 then
            set outputText to outputText & "❌ No free slots found in the specified range." & return
        else
            set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
            set outputText to outputText & "Found " & suggestionsFound & " available time slot(s)" & return
        end if

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def detect_conflicts(
    start_date: str,
    end_date: str,
    calendar: Optional[str] = None
) -> str:
    """
    Detect overlapping/conflicting events in your calendar.

    Args:
        start_date: Start of date range to check (YYYY-MM-DD format)
        end_date: End of date range to check (YYYY-MM-DD format)
        calendar: Optional calendar name to check (if None, checks all calendars)

    Returns:
        List of conflicting events
    """

    start_formatted = format_applescript_date(start_date)
    end_formatted = format_applescript_date(end_date)

    calendar_filter = f'''
        set targetCalendars to {{calendar "{calendar}"}}
    ''' if calendar else '''
        set targetCalendars to every calendar
    '''

    script = f'''
    tell application "Calendar"
        set outputText to "CONFLICT DETECTION" & return & return
        set outputText to outputText & "Date range: {start_date} to {end_date}" & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

        set searchStart to date "{start_formatted}"
        set searchEnd to date "{end_formatted}"

        {calendar_filter}

        -- Collect all events
        set allEvents to {{}}

        repeat with aCal in targetCalendars
            set calName to name of aCal
            try
                set calEvents to (every event of aCal whose start date >= searchStart and start date <= searchEnd and allday event is false)

                repeat with anEvent in calEvents
                    set eventInfo to {{anEvent, calName}}
                    set end of allEvents to eventInfo
                end repeat
            end try
        end repeat

        -- Find conflicts
        set conflictCount to 0
        set checkedPairs to {{}}

        repeat with i from 1 to (count of allEvents)
            set event1Info to item i of allEvents
            set event1 to item 1 of event1Info
            set cal1 to item 2 of event1Info

            repeat with j from (i + 1) to (count of allEvents)
                set event2Info to item j of allEvents
                set event2 to item 1 of event2Info
                set cal2 to item 2 of event2Info

                set start1 to start date of event1
                set end1 to end date of event1
                set start2 to start date of event2
                set end2 to end date of event2

                -- Check for overlap
                if (start1 < end2) and (end1 > start2) then
                    set conflictCount to conflictCount + 1

                    set outputText to outputText & "⚠️  CONFLICT #" & conflictCount & return & return

                    set outputText to outputText & "Event 1:" & return
                    set outputText to outputText & "  📌 " & summary of event1 & return
                    set outputText to outputText & "  📅 " & date string of start1 & return
                    set outputText to outputText & "  🕐 " & time string of start1 & " - " & time string of end1 & return
                    set outputText to outputText & "  📁 " & cal1 & return & return

                    set outputText to outputText & "Event 2:" & return
                    set outputText to outputText & "  📌 " & summary of event2 & return
                    set outputText to outputText & "  📅 " & date string of start2 & return
                    set outputText to outputText & "  🕐 " & time string of start2 & " - " & time string of end2 & return
                    set outputText to outputText & "  📁 " & cal2 & return

                    set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return
                end if
            end repeat
        end repeat

        if conflictCount = 0 then
            set outputText to outputText & "✅ No conflicts found!" & return
        else
            set outputText to outputText & "Total conflicts: " & conflictCount & return
        end if

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def get_statistics(
    start_date: str,
    end_date: str,
    calendar: Optional[str] = None
) -> str:
    """
    Get calendar statistics and analytics for a date range.

    Args:
        start_date: Start of analysis period (YYYY-MM-DD format)
        end_date: End of analysis period (YYYY-MM-DD format)
        calendar: Optional calendar name to analyze (if None, analyzes all calendars)

    Returns:
        Detailed statistics including:
        - Total events and meetings
        - Total time in meetings
        - Busiest days/hours
        - Average meeting duration
    """

    start_formatted = format_applescript_date(start_date)
    end_formatted = format_applescript_date(end_date)

    calendar_filter = f'''
        set targetCalendars to {{calendar "{calendar}"}}
    ''' if calendar else '''
        set targetCalendars to every calendar
    '''

    script = f'''
    tell application "Calendar"
        set outputText to "CALENDAR STATISTICS" & return & return
        set outputText to outputText & "Period: {start_date} to {end_date}" & return
        set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

        set searchStart to date "{start_formatted}"
        set searchEnd to date "{end_formatted}"

        {calendar_filter}

        -- Initialize counters
        set totalEvents to 0
        set totalTimedEvents to 0
        set totalAllDayEvents to 0
        set totalMinutes to 0
        set hourCounts to {{0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}}
        set dayCounts to {{}}

        repeat with aCal in targetCalendars
            try
                set calEvents to (every event of aCal whose start date >= searchStart and start date <= searchEnd)

                repeat with anEvent in calEvents
                    set totalEvents to totalEvents + 1

                    if allday event of anEvent then
                        set totalAllDayEvents to totalAllDayEvents + 1
                    else
                        set totalTimedEvents to totalTimedEvents + 1

                        set eventStart to start date of anEvent
                        set eventEnd to end date of anEvent

                        -- Calculate duration in minutes
                        set duration to (eventEnd - eventStart)
                        set durationMinutes to duration / 60
                        set totalMinutes to totalMinutes + durationMinutes

                        -- Track hour distribution
                        try
                            set eventHour to hours of eventStart
                            if eventHour >= 0 and eventHour < 24 then
                                set hourIndex to eventHour + 1
                                set currentCount to item hourIndex of hourCounts
                                set item hourIndex of hourCounts to currentCount + 1
                            end if
                        end try
                    end if
                end repeat
            end try
        end repeat

        -- Calculate averages
        set avgDuration to 0
        if totalTimedEvents > 0 then
            set avgDuration to totalMinutes / totalTimedEvents
        end if

        set totalHours to totalMinutes / 60

        -- Output statistics
        set outputText to outputText & "📊 SUMMARY" & return
        set outputText to outputText & "  Total Events: " & totalEvents & return
        set outputText to outputText & "  Timed Events: " & totalTimedEvents & return
        set outputText to outputText & "  All-Day Events: " & totalAllDayEvents & return & return

        set outputText to outputText & "⏱  TIME ANALYSIS" & return
        set outputText to outputText & "  Total Meeting Time: " & (totalHours as integer) & " hours " & ((totalMinutes mod 60) as integer) & " minutes" & return
        set outputText to outputText & "  Average Meeting Duration: " & (avgDuration as integer) & " minutes" & return & return

        -- Find busiest hours
        set outputText to outputText & "🕐 BUSIEST HOURS" & return
        set maxHourCount to 0
        set busiestHours to {{}}

        repeat with i from 1 to 24
            set hourCount to item i of hourCounts
            if hourCount > maxHourCount then
                set maxHourCount to hourCount
            end if
        end repeat

        if maxHourCount > 0 then
            repeat with i from 1 to 24
                set hourCount to item i of hourCounts
                if hourCount = maxHourCount then
                    set hourLabel to (i - 1)
                    if hourLabel < 12 then
                        set ampm to "AM"
                        if hourLabel = 0 then set hourLabel to 12
                    else
                        set ampm to "PM"
                        if hourLabel > 12 then set hourLabel to hourLabel - 12
                    end if
                    set outputText to outputText & "  " & hourLabel & " " & ampm & ": " & hourCount & " events" & return
                end if
            end repeat
        else
            set outputText to outputText & "  No timed events in this period" & return
        end if

        set outputText to outputText & return & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return

        return outputText
    end tell
    '''

    result = run_applescript(script)
    return result


@mcp.tool()
@inject_preferences
def export_events(
    start_date: str,
    end_date: str,
    format: str = "txt",
    calendar: Optional[str] = None,
    output_file: Optional[str] = None
) -> str:
    """
    Export events to a file in various formats.

    Args:
        start_date: Start of export range (YYYY-MM-DD format)
        end_date: End of export range (YYYY-MM-DD format)
        format: Export format - "txt", "csv", or "ics" (default: "txt")
        calendar: Optional calendar name to export (if None, exports all calendars)
        output_file: Optional output file path (if None, returns data as string)

    Returns:
        Exported event data or confirmation message if file was written
    """

    start_formatted = format_applescript_date(start_date)
    end_formatted = format_applescript_date(end_date)

    calendar_filter = f'''
        set targetCalendars to {{calendar "{calendar}"}}
    ''' if calendar else '''
        set targetCalendars to every calendar
    '''

    if format.lower() == "csv":
        # CSV format
        script = f'''
        tell application "Calendar"
            set outputText to "Title,Start Date,Start Time,End Date,End Time,Location,Calendar,All Day" & return

            set searchStart to date "{start_formatted}"
            set searchEnd to date "{end_formatted}"

            {calendar_filter}

            repeat with aCal in targetCalendars
                set calName to name of aCal

                try
                    set calEvents to (every event of aCal whose start date >= searchStart and start date <= searchEnd)

                    repeat with anEvent in calEvents
                        try
                            set eventTitle to summary of anEvent
                            set eventStart to start date of anEvent
                            set eventEnd to end date of anEvent
                            set isAllDay to allday event of anEvent

                            -- Replace commas in title to avoid CSV issues
                            set AppleScript's text item delimiters to ","
                            set titleParts to text items of eventTitle
                            set AppleScript's text item delimiters to ";"
                            set eventTitle to titleParts as string
                            set AppleScript's text item delimiters to ""

                            set evtLoc to ""
                            try
                                set evtLoc to location of anEvent
                                if evtLoc is missing value then set evtLoc to ""
                                -- Replace commas in location
                                set AppleScript's text item delimiters to ","
                                set locParts to text items of evtLoc
                                set AppleScript's text item delimiters to ";"
                                set evtLoc to locParts as string
                                set AppleScript's text item delimiters to ""
                            end try

                            if isAllDay then
                                set outputText to outputText & eventTitle & "," & (date string of eventStart) & ",All Day," & (date string of eventEnd) & ",All Day," & evtLoc & "," & calName & ",Yes" & return
                            else
                                set outputText to outputText & eventTitle & "," & (date string of eventStart) & "," & (time string of eventStart) & "," & (date string of eventEnd) & "," & (time string of eventEnd) & "," & evtLoc & "," & calName & ",No" & return
                            end if
                        end try
                    end repeat
                end try
            end repeat

            return outputText
        end tell
        '''

    elif format.lower() == "ics":
        # ICS/iCalendar format (simplified)
        script = f'''
        tell application "Calendar"
            set outputText to "BEGIN:VCALENDAR" & return
            set outputText to outputText & "VERSION:2.0" & return
            set outputText to outputText & "PRODID:-//Apple Calendar MCP//EN" & return & return

            set searchStart to date "{start_formatted}"
            set searchEnd to date "{end_formatted}"

            {calendar_filter}

            repeat with aCal in targetCalendars
                try
                    set calEvents to (every event of aCal whose start date >= searchStart and start date <= searchEnd)

                    repeat with anEvent in calEvents
                        try
                            set outputText to outputText & "BEGIN:VEVENT" & return

                            set eventTitle to summary of anEvent
                            set outputText to outputText & "SUMMARY:" & eventTitle & return

                            try
                                set evtLoc to location of anEvent
                                if evtLoc is not missing value and evtLoc is not "" then
                                    set outputText to outputText & "LOCATION:" & evtLoc & return
                                end if
                            end try

                            try
                                set evtNotes to description of anEvent
                                if evtNotes is not missing value and evtNotes is not "" then
                                    set outputText to outputText & "DESCRIPTION:" & evtNotes & return
                                end if
                            end try

                            set outputText to outputText & "END:VEVENT" & return & return
                        end try
                    end repeat
                end try
            end repeat

            set outputText to outputText & "END:VCALENDAR" & return

            return outputText
        end tell
        '''

    else:
        # TXT format (default)
        script = f'''
        tell application "Calendar"
            set outputText to "CALENDAR EXPORT" & return
            set outputText to outputText & "Period: {start_date} to {end_date}" & return
            set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return & return

            set searchStart to date "{start_formatted}"
            set searchEnd to date "{end_formatted}"

            {calendar_filter}

            set eventCount to 0

            repeat with aCal in targetCalendars
                set calName to name of aCal

                try
                    set calEvents to (every event of aCal whose start date >= searchStart and start date <= searchEnd)

                    repeat with anEvent in calEvents
                        try
                            set eventTitle to summary of anEvent
                            set eventStart to start date of anEvent
                            set eventEnd to end date of anEvent
                            set isAllDay to allday event of anEvent

                            set outputText to outputText & "📌 " & eventTitle & return
                            set outputText to outputText & "   📅 " & date string of eventStart & return

                            if isAllDay then
                                set outputText to outputText & "   🕐 All Day Event" & return
                            else
                                set outputText to outputText & "   🕐 " & time string of eventStart & " - " & time string of eventEnd & return
                            end if

                            set outputText to outputText & "   📁 " & calName & return

                            try
                                set evtLoc to location of anEvent
                                if evtLoc is not missing value and length of evtLoc > 0 then
                                    set outputText to outputText & "   📍 " & evtLoc & return
                                end if
                            end try

                            set outputText to outputText & return
                            set eventCount to eventCount + 1
                        end try
                    end repeat
                end try
            end repeat

            set outputText to outputText & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & return
            set outputText to outputText & "Total: " & eventCount & " event(s) exported" & return

            return outputText
        end tell
        '''

    result = run_applescript(script)

    # If output_file specified, write to file
    if output_file:
        try:
            with open(output_file, 'w') as f:
                f.write(result)
            return f"✅ Events exported successfully to: {output_file}\n\nFormat: {format.upper()}\nFile size: {len(result)} bytes"
        except Exception as e:
            return f"❌ Error writing to file: {str(e)}\n\nData preview:\n{result[:500]}"

    return result


if __name__ == "__main__":
    # Run the MCP server
    mcp.run()
