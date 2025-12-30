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


if __name__ == "__main__":
    # Run the MCP server
    mcp.run()
