# Project Tracker Pro

A comprehensive desktop project management application built with Python and Tkinter. Designed for freelancers and small teams to track projects, tasks, finances, and more.

## Features

### Core
- **Dashboard** - Stats cards, overdue alerts, upcoming deadlines, workload pie chart, project progress table
- **Project Management** - Full CRUD with client info, status, priority, risk level, tags, budget, custom colors
- **Tasks** - Per-project tasks with status (todo/doing/done), priority, due dates, hours logged vs estimated
- **Subtasks** - Nested subtasks within tasks
- **Task Comments** - Timestamped notes on any task
- **Milestones** - Key dates/markers per project with toggle done
- **Task Dependencies** - "Blocked by" field to track blocking tasks

### Views
- **Kanban Board** - Visual columns (Todo → Doing → Done) with color-coded cards, double-click to edit
- **Gantt Timeline** - Visual timeline with bars for projects and tasks, today marker
- **Calendar View** - Monthly calendar with task due dates, overdue highlighting, double-click for details
- **Timesheet** - Weekly timesheet view with navigation and CSV export

### Financial
- **Budget Tracking** - Per-project budget amounts
- **Invoices** - Track invoices with paid/unpaid status
- **Expenses** - Log expenses per project
- **Financial Dashboard** - Summary cards (budget, invoiced, paid, outstanding, expenses, profit) and per-project breakdown

### Productivity
- **Activity Log** - Timestamped history of all changes
- **Batch Operations** - Update status or priority for all tasks at once
- **Project Archive** - Archive completed projects instead of deleting
- **Duplicate Project** - Clone a project with all its data
- **Global Search** - Search across all projects and tasks (Ctrl+F)
- **Sortable Columns** - Click any column header to sort
- **Dark Mode** - Toggle light/dark theme

### Data
- **Auto-save** - All changes saved automatically on every action
- **Backup/Restore** - Export/import full data as JSON
- **CSV Import/Export** - Import existing data, export reports
- **Print Report** - Generate printable project summary in a popup window

### Design Phrases
- Category management with editable phrase lists (preserved from original app)

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+S` | Save data manually |
| `Ctrl+F` | Focus global search bar |
| `Ctrl+N` | Quick-add new project |
| `F5` | Refresh all views |

## Installation

### Requirements
- Python 3.8 or higher
- No external dependencies (uses only standard library)

### Run
```bash
python project_tracker.py
```

## Project Structure

```
project_tracker.py    # Main application (single-file, no dependencies)
design_phrases_data.json  # Auto-generated data file (created on first run)
README.md
.gitignore
LICENSE
```

## Data Storage

All data is stored in `design_phrases_data.json` in the same directory as the script. This includes:
- Projects with tasks, subtasks, milestones, invoices, expenses
- Design phrases and categories
- Activity log (last 500 entries)

Use **File > Backup Data** to create external backups.

## License

MIT License - see [LICENSE](LICENSE) file.
