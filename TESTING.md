# Apple Calendar MCP - Testing Guide

## Quick Start Testing

### 1. Configure Claude Desktop

Add this to your Claude Desktop config file:
`~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "apple-calendar": {
      "command": "/Users/andrew/GitHub/apple-calendar-mcp/venv/bin/python3",
      "args": ["/Users/andrew/GitHub/apple-calendar-mcp/apple_calendar_mcp.py"]
    }
  }
}
```

### 2. Restart Claude Desktop

Completely quit and restart Claude Desktop to load the new MCP server.

### 3. Verify Connection

In Claude Desktop, try: "List my calendars"

You should see a response showing your Apple Calendar calendars.

---

## Test Scenarios

### Basic Reading
```
1. "Show me my calendar overview"
2. "What's on my schedule today?"
3. "List all my calendars"
4. "Show me events this week"
```

### Event Creation
```
5. "Create a meeting called 'Team Standup' tomorrow at 9am for 30 minutes in calendar Work"
6. "Create a recurring weekly meeting every Monday at 2pm called 'Project Review'"
```

### Event Management
```
7. "Move my 'Team Standup' meeting from tomorrow to the day after at 10am"
8. "Update the 'Project Review' meeting to add location 'Conference Room A'"
9. "Delete the event 'Old Meeting' scheduled for 2025-01-15"
10. "Add a 15 minute reminder to my 'Team Standup' event"
```

### Search & Analysis
```
11. "Search for all events with 'project' in the title from last month"
12. "Find me a free 1-hour time slot tomorrow between 9am and 5pm"
13. "Check for any conflicting events this week"
```

### Analytics & Export
```
14. "Show me statistics for my calendar this month"
15. "Export all my events from this week to CSV format"
```

---

## Troubleshooting

### MCP Server Not Showing Up

1. Check Claude Desktop logs:
   ```bash
   tail -f ~/Library/Logs/Claude/mcp*.log
   ```

2. Verify server runs standalone:
   ```bash
   cd /Users/andrew/GitHub/apple-calendar-mcp
   source venv/bin/activate
   python3 apple_calendar_mcp.py
   ```

3. Check permissions:
   - System Settings > Privacy & Security > Automation
   - Ensure Terminal/Claude has access to Calendar

### "Calendar not found" Errors

List your actual calendar names first:
```
"List my calendars"
```

Then use the exact calendar name in commands:
```
"Create an event in calendar 'Work' ..."
```

### Date Format Issues

Always use YYYY-MM-DD format:
- Correct: `2025-01-15`
- Correct: `2025-01-15 14:30`
- Wrong: `Jan 15, 2025`
- Wrong: `01/15/2025`

---

## Success Indicators

✅ **Server is working if:**
- Claude can list your calendars
- You can create and see events in Apple Calendar app
- Search returns results
- No error messages in responses

❌ **Server needs debugging if:**
- "Tool not found" errors
- AppleScript timeout errors
- "Calendar not found" for valid calendars
- Blank responses

---

## Performance Notes

- First request may be slow (1-2 seconds) while Calendar.app launches
- Subsequent requests should be faster (<1 second)
- Large date ranges (>1 year) may take longer
- Export operations with many events (>100) may timeout

---

## Next Steps

Once testing is complete:
1. Update README with your real-world usage examples
2. Report any bugs or issues
3. Consider adding more features based on your workflow
4. Star the repo and share with others!
