# Apple Calendar MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io)

A comprehensive Model Context Protocol (MCP) server that provides AI assistants with natural language access to Apple Calendar. Built with [FastMCP](https://github.com/jlowin/fastmcp), this server enables reading, searching, creating, and managing calendar events directly through Claude Desktop or other MCP-compatible clients.

## Features

### Reading & Listing
- **Calendar Overview**: Dashboard view with all calendars, today's events, and upcoming week
- **List Calendars**: View all available calendars with event counts
- **List Events**: Browse events from specific calendars or date ranges
- **Today's Schedule**: Quick view of current day's agenda

### Search & Analysis
- **Advanced Search**: Multi-criteria search (title, location, date range, attendees)
- **Find Free Time**: Discover available time slots for scheduling
- **Conflict Detection**: Check for overlapping events

### Event Creation
- **Create Events**: Full event creation with all options (location, notes, attendees)
- **Recurring Events**: Create daily, weekly, monthly recurring events
- **Quick Add**: Natural language event creation

### Event Management
- **Update Events**: Modify any event property
- **Move Events**: Reschedule events to new times
- **Delete Events**: Remove events with recurrence handling
- **Manage Reminders**: Add or modify event alerts

### Analytics & Export
- **Statistics**: Meeting time analysis, busy hours breakdown
- **Export**: Export events to ICS, CSV, or TXT formats

## Installation

### Prerequisites
- macOS with Apple Calendar configured
- Python 3.7 or higher
- Claude Desktop or any MCP-compatible client

### Quick Start

1. Clone the repository:
```bash
git clone https://github.com/andrewbergsma/apple-calendar-mcp.git
cd apple-calendar-mcp
```

2. Create and activate virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

3. Configure Claude Desktop (`~/Library/Application Support/Claude/claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "apple-calendar": {
      "command": "/path/to/apple-calendar-mcp/venv/bin/python3",
      "args": ["/path/to/apple-calendar-mcp/apple_calendar_mcp.py"]
    }
  }
}
```

4. Restart Claude Desktop

## Usage Examples

```
Show me my schedule for today
What meetings do I have this week?
Find a free 30-minute slot tomorrow afternoon
Create a meeting with John at 2pm on Friday
Move my 3pm meeting to 4pm
Search for all events about "project review"
```

## License

MIT License - see [LICENSE](LICENSE) for details.

## Acknowledgments

- Built with [FastMCP](https://github.com/jlowin/fastmcp)
- Inspired by [Apple Mail MCP](https://github.com/patrickfreyer/apple-mail-mcp)
- Part of the [Model Context Protocol](https://modelcontextprotocol.io) ecosystem
