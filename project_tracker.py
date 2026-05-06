import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog, colorchooser
import json
import os
import csv
import copy
from datetime import datetime, date, timedelta
import math


DATA_FILE = "design_phrases_data.json"


class DesignAppData:
    def __init__(self):
        self.phrases = {
            "3d_modelling": [
                "3d modelling design", "Concept Design", "Features", "Sketching",
                "Constraints", "Mates in Assembly", "Interference Detection",
                "Assembly Visualization", "Ergonomic Analysis", "Mortise and Tenons",
                "Plate thickness changes", "assemblies", "Sheetmetal modelling",
                "hole additions for fasteners", "Concept visualization"
            ],
            "drawings": [
                "2D drawings", "dimensioning", "GD&T symbols",
                "section views", "detail views"
            ],
            "calculations": [
                "FEA analysis", "load calculations", "tolerance stackup",
                "material strength check", "thermal analysis"
            ],
            "quality_check": [
                "Inspection report", "CMM check", "visual inspection",
                "NDT testing", "first article inspection"
            ],
            "weldstudy": [
                "weld symbol review", "weld size check", "preheat requirement",
                "NDT for welds", "weld procedure spec"
            ],
            "bom_generation": [
                "BOM structure", "part numbers", "quantity check",
                "material list", "commercial BOM"
            ],
            "others": [
                "Delivery - logistics", "meetings - review", "client feedback",
                "document control", "project handover"
            ]
        }
        self.categories = [
            {"key": "3d_modelling", "display_name": "3D Modelling"},
            {"key": "drawings", "display_name": "Drawings"},
            {"key": "calculations", "display_name": "Calculations"},
            {"key": "quality_check", "display_name": "Quality Check"},
            {"key": "weldstudy", "display_name": "Weld Study"},
            {"key": "bom_generation", "display_name": "BOM Generation"},
            {"key": "others", "display_name": "Others"}
        ]
        self.projects = []
        self.activity_log = []
        self.load_data()

    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r") as f:
                    data = json.load(f)
                    self.phrases = data.get("phrases", self.phrases)
                    self.categories = data.get("categories", self.categories)
                    loaded_projects = data.get("projects", [])
                    self.activity_log = data.get("activity_log", [])
                    migrated = []
                    for p in loaded_projects:
                        migrated.append({
                            "name": p.get("name", ""),
                            "code": p.get("code", ""),
                            "remarks": p.get("remarks", ""),
                            "client": p.get("client", ""),
                            "status": p.get("status", "not_started"),
                            "due_date": p.get("due_date", ""),
                            "priority": p.get("priority", "medium"),
                            "tags": p.get("tags", []),
                            "color": p.get("color", "#3498db"),
                            "budget": float(p.get("budget", 0.0)),
                            "expenses": p.get("expenses", []),
                            "invoices": p.get("invoices", []),
                            "milestones": p.get("milestones", []),
                            "risk_level": p.get("risk_level", "low"),
                            "archived": p.get("archived", False),
                            "created_date": p.get("created_date",
                                datetime.now().strftime("%Y-%m-%d")),
                            "tasks": self._migrate_tasks(p.get("tasks", []))
                        })
                    self.projects = migrated
            except Exception:
                pass

    def _migrate_tasks(self, tasks):
        migrated = []
        for t in tasks:
            migrated.append({
                "title": t.get("title", ""),
                "status": t.get("status", "todo"),
                "priority": t.get("priority", "medium"),
                "due_date": t.get("due_date", ""),
                "hours": float(t.get("hours", 0)),
                "estimated_hours": float(t.get("estimated_hours", 0)),
                "subtasks": self._migrate_tasks(t.get("subtasks", [])),
                "comments": t.get("comments", []),
                "blocked_by": t.get("blocked_by", ""),
                "attachments": t.get("attachments", [])
            })
        return migrated

    def save_data(self):
        with open(DATA_FILE, "w") as f:
            json.dump({
                "phrases": self.phrases,
                "categories": self.categories,
                "projects": self.projects,
                "activity_log": self.activity_log[-500:]
            }, f, indent=4)

    def log_activity(self, action, details=""):
        self.activity_log.append({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "action": action,
            "details": details
        })


STATUS_COLORS = {
    "not_started": "#e74c3c",
    "in_progress": "#f39c12",
    "completed": "#2ecc71",
    "on_hold": "#95a5a6"
}
PRIORITY_COLORS = {
    "low": "#3498db",
    "medium": "#f39c12",
    "high": "#e74c3c",
    "urgent": "#8e44ad"
}
RISK_COLORS = {"low": "#2ecc71", "medium": "#f39c12", "high": "#e74c3c"}


class ProjectTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Project Tracker Pro")
        self.root.geometry("1350x850")
        self.data = DesignAppData()
        self.current_category = None
        self.category_checkboxes = {}
        self.theme = "light"
        self.sort_col = None
        self.sort_reverse = False
        self._selected_project = None

        menubar = tk.Menu(root)
        root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Backup Data (JSON)", command=self.backup_data)
        file_menu.add_command(label="Restore Data", command=self.restore_data)
        file_menu.add_command(label="Import CSV", command=self.import_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)

        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Toggle Dark Mode", command=self.toggle_theme)
        view_menu.add_command(label="Refresh All", command=self.refresh_all)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_shortcuts)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.dashboard_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.dashboard_frame, text="Dashboard")

        self.projects_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.projects_frame, text="Projects")

        self.kanban_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.kanban_frame, text="Kanban Board")

        self.gantt_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.gantt_frame, text="Gantt Timeline")

        self.calendar_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.calendar_frame, text="Calendar")

        self.timesheet_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.timesheet_frame, text="Timesheet")

        self.financial_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.financial_frame, text="Financials")

        self.activity_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.activity_frame, text="Activity Log")

        self.phrases_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.phrases_frame, text="Design Phrases")

        self.setup_global_search()
        self.setup_dashboard()
        self.setup_projects_tab()
        self.setup_kanban()
        self.setup_gantt()
        self.setup_calendar()
        self.setup_timesheet()
        self.setup_financials()
        self.setup_activity_log()
        self.setup_phrases_tab()

        self.root.bind("<Control-s>", lambda e: (self.data.save_data(),
            messagebox.showinfo("Saved", "Data saved.")))
        self.root.bind("<Control-f>", lambda e: (
            self.global_search_entry.focus_set(),
            self.global_search_entry.select_range(0, tk.END)))
        self.root.bind("<Control-n>", lambda e: self.quick_new_project())
        self.root.bind("<F5>", lambda e: self.refresh_all())

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def toggle_theme(self):
        self.theme = "dark" if self.theme == "light" else "light"
        t_bg = "#1e1e1e" if self.theme == "dark" else "#ffffff"
        t_fg = "#d4d4d4" if self.theme == "dark" else "#000000"
        card_bg = "#2d2d2d" if self.theme == "dark" else "#f8f9fa"
        header_bg = "#323232" if self.theme == "dark" else "#f0f0f0"
        header_fg = "#ffffff" if self.theme == "dark" else "#000000"
        accent = "#569cd6" if self.theme == "dark" else "#3498db"

        self.root.configure(bg=t_bg)
        style = ttk.Style()
        if self.theme == "dark":
            style.theme_use("clam")
        else:
            style.theme_use("default")
        style.configure(".", background=t_bg, foreground=t_fg, fieldbackground=card_bg)
        style.configure("TFrame", background=t_bg)
        style.configure("TLabelFrame", background=t_bg, foreground=t_fg)
        style.configure("TLabel", background=t_bg, foreground=t_fg)
        style.configure("TButton", background=accent, foreground="white")
        style.configure("Treeview", background=card_bg, foreground=t_fg, fieldbackground=card_bg)
        style.configure("Treeview.Heading", background=header_bg, foreground=header_fg)

    def refresh_all(self):
        self.refresh_dashboard()
        self.refresh_project_list()
        self.refresh_kanban()
        self.refresh_gantt()
        self.refresh_calendar_view()
        self.refresh_timesheet()
        self.refresh_financials()
        self.refresh_activity_log()

    # ---------------------------- GLOBAL SEARCH ----------------------------
    def setup_global_search(self):
        bar = ttk.Frame(self.root)
        bar.pack(fill=tk.X, padx=10, pady=(5, 0))
        ttk.Label(bar, text="Search:").pack(side=tk.LEFT)
        self.global_search_var = tk.StringVar()
        self.global_search_var.trace("w", lambda *_: self.global_search())
        self.global_search_entry = ttk.Entry(bar, textvariable=self.global_search_var, width=40)
        self.global_search_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(bar, text="Clear", width=8, command=lambda: self.global_search_var.set("")).pack(side=tk.LEFT)
        ttk.Label(bar, text="(Ctrl+F)", foreground="gray").pack(side=tk.LEFT)

    def global_search(self):
        q = self.global_search_var.get().lower()
        if not q:
            return
        results = []
        for p in self.data.projects:
            if p.get("archived"):
                continue
            match = any(q in str(v).lower() for v in [
                p.get("name", ""), p.get("code", ""), p.get("client", ""),
                p.get("remarks", ""), " ".join(p.get("tags", []))])
            matched_tasks = []
            for t in p.get("tasks", []):
                tm = any(q in str(v).lower() for v in [
                    t.get("title", ""), t.get("status", ""),
                    " ".join(c.get("text", "") for c in t.get("comments", []))])
                if tm:
                    matched_tasks.append(t)
            if match or matched_tasks:
                results.append((p, matched_tasks))

        if results:
            msg = f"Found {len(results)} project(s):\n\n"
            for p, tasks in results[:10]:
                msg += f"  - {p['name']} ({p.get('code','')}) - {p.get('status','')} "
                if tasks:
                    msg += f"[{len(tasks)} matching task(s)]"
                msg += "\n"
            if len(results) > 10:
                msg += f"\n... and {len(results)-10} more"
            messagebox.showinfo("Search Results", msg)

    def show_shortcuts(self):
        messagebox.showinfo("Keyboard Shortcuts",
            "Ctrl+S - Save data\nCtrl+F - Focus search bar\n"
            "Ctrl+N - New project quick-add\nF5 - Refresh all views")

    def quick_new_project(self):
        name = simpledialog.askstring("Quick New Project", "Project name:")
        if not name:
            return
        code = simpledialog.askstring("Project Code", "Code:") or name[:4].upper()
        self.data.projects.append({
            "name": name.strip(), "code": code.strip(),
            "client": "", "status": "not_started",
            "priority": "medium", "due_date": "",
            "tags": [], "color": "#3498db", "budget": 0.0,
            "expenses": [], "invoices": [], "milestones": [],
            "risk_level": "low", "archived": False,
            "created_date": datetime.now().strftime("%Y-%m-%d"),
            "remarks": "", "tasks": []
        })
        self.data.log_activity("PROJECT_CREATED", f"'{name}' ({code})")
        self.data.save_data()
        self.refresh_all()

    # ---------------------------- DASHBOARD ----------------------------
    def setup_dashboard(self):
        top = ttk.Frame(self.dashboard_frame)
        top.pack(fill=tk.X, padx=20, pady=15)
        ttk.Label(top, text="PROJECT TRACKER DASHBOARD", font=("Arial", 18, "bold")).pack(anchor=tk.W)
        ttk.Label(top, text=datetime.now().strftime("%B %d, %Y"), font=("Arial", 10), foreground="gray").pack(anchor=tk.W)

        self.stats_container = ttk.Frame(self.dashboard_frame)
        self.stats_container.pack(fill=tk.X, padx=20, pady=10)

        self.alerts_container = ttk.Frame(self.dashboard_frame)
        self.alerts_container.pack(fill=tk.X, padx=20, pady=5)

        bottom = ttk.Frame(self.dashboard_frame)
        bottom.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        proj_frame = ttk.LabelFrame(bottom, text="Project Progress", padding=10)
        proj_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        cols = ("Name", "Code", "Status", "Progress", "Due", "Priority", "Hours", "Risk")
        self.dash_tree = ttk.Treeview(proj_frame, columns=cols, show="headings", height=10)
        for c, w in [("Name", 140), ("Code", 70), ("Status", 80), ("Progress", 100),
                      ("Due", 90), ("Priority", 65), ("Hours", 55), ("Risk", 45)]:
            self.dash_tree.heading(c, text=c)
            self.dash_tree.column(c, width=w)
        self.dash_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(proj_frame, orient="vertical", command=self.dash_tree.yview).pack(side=tk.RIGHT, fill=tk.Y)

        workload_frame = ttk.LabelFrame(bottom, text="Workload Overview", padding=10)
        workload_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        self.workload_canvas = tk.Canvas(workload_frame, height=200, bg="white")
        self.workload_canvas.pack(fill=tk.BOTH, expand=True)

        ttk.Button(self.dashboard_frame, text="Refresh Dashboard", command=self.refresh_dashboard).pack(pady=10)
        self.refresh_dashboard()

    def refresh_dashboard(self):
        for w in self.stats_container.winfo_children():
            w.destroy()

        active = [p for p in self.data.projects if not p.get("archived")]
        archived = [p for p in self.data.projects if p.get("archived")]
        total = len(active)
        completed = sum(1 for p in active if p.get("status") == "completed")
        in_progress = sum(1 for p in active if p.get("status") == "in_progress")
        on_hold = sum(1 for p in active if p.get("status") == "on_hold")
        not_started = sum(1 for p in active if p.get("status") == "not_started")

        all_tasks = []
        for p in active:
            all_tasks.extend(p.get("tasks", []))
        total_tasks = len(all_tasks)
        done_tasks = sum(1 for t in all_tasks if t.get("status") == "done")

        today = date.today()
        overdue_projects = []
        overdue_tasks = []
        upcoming = []
        for p in active:
            if p.get("due_date") and p.get("status") != "completed":
                try:
                    dd = datetime.strptime(p["due_date"], "%Y-%m-%d").date()
                    if dd < today:
                        overdue_projects.append(p)
                    elif dd <= today + timedelta(days=7):
                        upcoming.append((p, "project"))
                except ValueError:
                    pass
            for t in p.get("tasks", []):
                if t.get("due_date") and t.get("status") != "done":
                    try:
                        td = datetime.strptime(t["due_date"], "%Y-%m-%d").date()
                        if td < today:
                            overdue_tasks.append((p, t))
                        elif td <= today + timedelta(days=7):
                            upcoming.append((t, "task"))
                    except ValueError:
                        pass

        total_hours = sum(t.get("hours", 0) for p in active for t in p.get("tasks", []))
        total_estimated = sum(t.get("estimated_hours", 0) for p in active for t in p.get("tasks", []))
        total_budget = sum(p.get("budget", 0) for p in active)
        total_expenses = sum(sum(e.get("amount", 0) for e in p.get("expenses", [])) for p in active)
        total_invoiced = sum(sum(i.get("amount", 0) for i in p.get("invoices", [])) for p in active)
        total_paid = sum(sum(i.get("amount", 0) for i in p.get("invoices", []) if i.get("paid")) for p in active)

        cards = [
            ("Active Projects", str(total), "#3498db"),
            ("Completed", str(completed), "#2ecc71"),
            ("In Progress", str(in_progress), "#f39c12"),
            ("On Hold", str(on_hold), "#95a5a6"),
            ("Not Started", str(not_started), "#e74c3c"),
            ("Archived", str(len(archived)), "#7f8c8d"),
            ("Tasks Done", f"{done_tasks}/{total_tasks}", "#9b59b6"),
            ("Overdue Projects", str(len(overdue_projects)), "#c0392b"),
            ("Overdue Tasks", str(len(overdue_tasks)), "#e67e22"),
            ("Due This Week", str(len(upcoming)), "#1abc9c"),
            ("Hours Logged", f"{total_hours:.1f}h", "#2980b9"),
            ("Hours Estimated", f"{total_estimated:.1f}h", "#8e44ad"),
            ("Total Budget", f"${total_budget:,.0f}", "#16a085"),
            ("Total Expenses", f"${total_expenses:,.0f}", "#d35400"),
            ("Invoiced", f"${total_invoiced:,.0f}", "#27ae60"),
            ("Paid", f"${total_paid:,.0f}", "#2ecc71"),
        ]

        for i, (label, value, color) in enumerate(cards):
            col = i % 4
            row = i // 4
            card = tk.Frame(self.stats_container, bg=color, height=70)
            card.grid(row=row, column=col, padx=6, pady=5, sticky="nsew")
            tk.Label(card, text=value, font=("Arial", 18, "bold"), bg=color, fg="white").pack(pady=(8, 0))
            tk.Label(card, text=label, font=("Arial", 8), bg=color, fg="white").pack()

        for c in range(4):
            self.stats_container.columnconfigure(c, weight=1)

        for w in self.alerts_container.winfo_children():
            w.destroy()

        if overdue_projects or overdue_tasks:
            warn = ttk.LabelFrame(self.alerts_container, text="ALERTS - OVERDUE", padding=10)
            warn.pack(fill=tk.X, pady=(0, 5))
            txt = tk.Text(warn, height=4, wrap=tk.WORD, font=("Consolas", 9))
            txt.pack(fill=tk.BOTH, expand=True)
            for p in overdue_projects:
                txt.insert(tk.END, f"  [PROJECT OVERDUE] {p['name']} - Due: {p.get('due_date','')}\n")
            for proj, task in overdue_tasks:
                txt.insert(tk.END,
                    f"  [TASK OVERDUE] '{task['title']}' in {proj['name']} - Due: {task.get('due_date','')}\n")

        if upcoming:
            up_frame = ttk.LabelFrame(self.alerts_container, text="COMING UP THIS WEEK", padding=10)
            up_frame.pack(fill=tk.X)
            txt = tk.Text(up_frame, height=3, wrap=tk.WORD, font=("Consolas", 9))
            txt.pack(fill=tk.BOTH, expand=True)
            for item, kind in upcoming:
                if kind == "project":
                    txt.insert(tk.END, f"  [PROJECT] {item['name']} - Due: {item.get('due_date','')}\n")
                else:
                    txt.insert(tk.END, f"  [TASK] '{item['title']}' - Due: {item.get('due_date','')}\n")

        for item in self.dash_tree.get_children():
            self.dash_tree.delete(item)

        for proj in sorted(active, key=lambda x: x.get("due_date", "9999")):
            tasks = proj.get("tasks", [])
            total_t = len(tasks)
            done_t = sum(1 for t in tasks if t.get("status") == "done")
            hours = sum(t.get("hours", 0) for t in tasks)
            est = sum(t.get("estimated_hours", 0) for t in tasks)

            progress = f"{int(done_t / total_t * 100)}% ({done_t}/{total_t})" if total_t > 0 else "No tasks"
            est_str = f"{hours:.0f}/{est:.0f}h" if est > 0 else f"{hours:.0f}h"

            self.dash_tree.insert("", tk.END, values=(
                proj["name"], proj.get("code", ""),
                proj.get("status", "not_started").replace("_", " ").title(),
                progress, proj.get("due_date", ""),
                proj.get("priority", "medium").title(), est_str,
                proj.get("risk_level", "low").title()
            ))

        self.draw_workload_chart()

    def draw_workload_chart(self):
        self.workload_canvas.delete("all")
        active = [p for p in self.data.projects if not p.get("archived")]
        status_counts = {}
        for p in active:
            s = p.get("status", "not_started")
            status_counts[s] = status_counts.get(s, 0) + 1

        categories = [("not_started", "#e74c3c"), ("in_progress", "#f39c12"),
                       ("completed", "#2ecc71"), ("on_hold", "#95a5a6")]
        total = sum(status_counts.values()) or 1
        w, h = self.workload_canvas.winfo_width(), self.workload_canvas.winfo_height()
        if w < 50:
            w = 300
        if h < 50:
            h = 200

        cx, cy = w // 2, h // 2
        r = min(w, h) // 2 - 20
        start_angle = 0
        for status, color in categories:
            count = status_counts.get(status, 0)
            if count > 0:
                angle = count / total * 360
                self.workload_canvas.create_arc(
                    (cx - r, cy - r), (cx + r, cy + r),
                    start=start_angle - 90, extent=angle, fill=color, outline="white",
                    width=2, style=tk.PIESLICE)
                mid = start_angle + angle / 2
                label_r = r // 2
                lx = cx + label_r * math.cos(math.radians(mid - 90))
                ly = cy + label_r * math.sin(math.radians(mid - 90))
                self.workload_canvas.create_text(lx, ly, text=f"{status.replace('_',' ').title()}\n({count})",
                                                  fill="white", font=("Arial", 8, "bold"))
                start_angle += angle

        y_offset = h - 30
        self.workload_canvas.create_text(cx, y_offset, text="Task Status:",
                                          font=("Arial", 9, "bold"), fill="#333")

    # ---------------------------- PROJECTS TAB (FIXED LAYOUT) ----------------------------
    def setup_projects_tab(self):
        top_bar = ttk.Frame(self.projects_frame)
        top_bar.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(top_bar, text="Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *_: self.filter_projects())
        ttk.Entry(top_bar, textvariable=self.search_var, width=20).pack(side=tk.LEFT, padx=5)

        ttk.Label(top_bar, text="Status:").pack(side=tk.LEFT, padx=(15, 0))
        self.filter_var = tk.StringVar(value="all")
        filter_cb = ttk.Combobox(top_bar, textvariable=self.filter_var,
                                  values=["all", "not_started", "in_progress", "completed", "on_hold"],
                                  width=12, state="readonly")
        filter_cb.pack(side=tk.LEFT, padx=5)

        ttk.Label(top_bar, text="Risk:").pack(side=tk.LEFT, padx=(10, 0))
        self.risk_filter_var = tk.StringVar(value="all")
        risk_cb = ttk.Combobox(top_bar, textvariable=self.risk_filter_var,
                                values=["all", "low", "medium", "high"], width=8, state="readonly")
        risk_cb.pack(side=tk.LEFT, padx=5)

        ttk.Label(top_bar, text="Tags:").pack(side=tk.LEFT, padx=(10, 0))
        self.tag_filter_var = tk.StringVar()
        ttk.Entry(top_bar, textvariable=self.tag_filter_var, width=12).pack(side=tk.LEFT, padx=5)

        for v in [self.filter_var, self.risk_filter_var, self.tag_filter_var]:
            v.trace("w", lambda *_: self.filter_projects())

        btn_bar = ttk.Frame(top_bar)
        btn_bar.pack(side=tk.RIGHT)
        ttk.Button(btn_bar, text="Export CSV", command=self.export_csv).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_bar, text="Print Report", command=self.print_report).pack(side=tk.LEFT, padx=3)

        paned = tk.PanedWindow(self.projects_frame, orient=tk.HORIZONTAL, sashrelief=tk.RAISED)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        left_panel = ttk.LabelFrame(paned, text="Project Details", padding=8)
        paned.add(left_panel, minsize=350)

        self._build_project_form(left_panel)

        right_panel = ttk.LabelFrame(paned, text="Project List", padding=8)
        paned.add(right_panel, minsize=300)

        proj_cols = ("Name", "Code", "Client", "Status", "Priority", "Due", "Risk", "Budget")
        self.project_tree = ttk.Treeview(right_panel, columns=proj_cols, show="headings", height=18)
        for c, w in [("Name", 140), ("Code", 70), ("Client", 100), ("Status", 80),
                      ("Priority", 60), ("Due", 90), ("Risk", 45), ("Budget", 65)]:
            self.project_tree.heading(c, text=c, command=lambda c=c: self.sort_projects(c))
            self.project_tree.column(c, width=w)
        self.project_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(right_panel, orient="vertical", command=self.project_tree.yview).pack(side=tk.RIGHT, fill=tk.Y)
        self.project_tree.bind("<<TreeviewSelect>>", self.on_project_select)

        self.refresh_project_list()

    def _build_project_form(self, parent):
        form_frame = ttk.Frame(parent)
        form_frame.pack(fill=tk.X, pady=2)

        fields = [
            ("Name *", "name"), ("Code *", "code"), ("Client", "client"),
            ("Status", "status"), ("Priority", "priority"),
            ("Due Date (YYYY-MM-DD)", "due_date"), ("Risk Level", "risk_level"),
            ("Tags (comma-sep)", "tags"), ("Budget ($)", "budget"), ("Color", "color")
        ]
        self.proj_fields = {}
        for i, (label, key) in enumerate(fields):
            ttk.Label(form_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=2, padx=2)
            if key == "status":
                cb = ttk.Combobox(form_frame, values=["not_started", "in_progress", "completed", "on_hold"],
                                  width=22, state="readonly")
                cb.grid(row=i, column=1, pady=2, padx=2)
                self.proj_fields[key] = cb
            elif key == "priority":
                cb = ttk.Combobox(form_frame, values=["low", "medium", "high", "urgent"],
                                  width=22, state="readonly")
                cb.grid(row=i, column=1, pady=2, padx=2)
                self.proj_fields[key] = cb
            elif key == "risk_level":
                cb = ttk.Combobox(form_frame, values=["low", "medium", "high"],
                                  width=22, state="readonly")
                cb.grid(row=i, column=1, pady=2, padx=2)
                self.proj_fields[key] = cb
            elif key == "color":
                btn = ttk.Button(form_frame, text="Pick Color", width=10,
                                 command=lambda k=key: self.pick_color(k))
                btn.grid(row=i, column=1, sticky=tk.W, pady=2, padx=2)
                self.proj_fields[key] = btn
            else:
                ent = ttk.Entry(form_frame, width=24)
                ent.grid(row=i, column=1, pady=2, padx=2)
                self.proj_fields[key] = ent

        ttk.Label(form_frame, text="Notes:").grid(row=len(fields), column=0, sticky=tk.NW, pady=2, padx=2)
        notes_txt = tk.Text(form_frame, height=3, wrap=tk.WORD, font=("Consolas", 9))
        notes_txt.grid(row=len(fields), column=1, pady=2, padx=2)
        self.proj_fields["remarks"] = notes_txt

        btn_row = ttk.Frame(parent)
        btn_row.pack(fill=tk.X, pady=5)
        ttk.Button(btn_row, text="Add", command=self.add_project).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Update", command=self.update_project).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Delete", command=self.delete_project).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Archive", command=self.archive_project).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Duplicate", command=self.duplicate_project).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="Clear", command=self.clear_project_form).pack(side=tk.LEFT, padx=2)

        scroll_frame = ttk.Frame(parent)
        scroll_frame.pack(fill=tk.BOTH, expand=True)

        task_canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        task_scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=task_canvas.yview)
        task_inner = ttk.Frame(task_canvas)
        task_canvas.configure(yscrollcommand=task_scrollbar.set)
        task_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        task_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        task_inner.bind("<Configure>", lambda e: task_canvas.configure(scrollregion=task_canvas.bbox("all")))
        task_canvas.create_window((0, 0), window=task_inner, anchor="nw")

        tasks_lbl = ttk.LabelFrame(task_inner, text="Tasks", padding=5)
        tasks_lbl.pack(fill=tk.X, pady=(0, 5))

        task_cols = ("Title", "Status", "Pri", "Due", "Hrs", "Est")
        self.task_tree = ttk.Treeview(tasks_lbl, columns=task_cols, show="headings", height=4)
        for c, w in [("Title", 130), ("Status", 50), ("Pri", 40), ("Due", 80), ("Hrs", 35), ("Est", 35)]:
            self.task_tree.heading(c, text=c)
            self.task_tree.column(c, width=w)
        self.task_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(tasks_lbl, orient="vertical", command=self.task_tree.yview).pack(side=tk.RIGHT, fill=tk.Y)

        task_btns = ttk.Frame(task_inner)
        task_btns.pack(fill=tk.X, pady=3)
        ttk.Button(task_btns, text="+ Task", command=self.add_task).pack(side=tk.LEFT, padx=2)
        ttk.Button(task_btns, text="Edit Task", command=self.edit_task_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(task_btns, text="Del Task", command=self.delete_task).pack(side=tk.LEFT, padx=2)
        ttk.Button(task_btns, text="Batch Tasks", command=self.batch_tasks).pack(side=tk.LEFT, padx=2)

        ms_frame = ttk.LabelFrame(task_inner, text="Milestones", padding=5)
        ms_frame.pack(fill=tk.X, pady=(5, 5))
        ms_cols = ("Milestone", "Date", "Done")
        self.ms_tree = ttk.Treeview(ms_frame, columns=ms_cols, show="headings", height=2)
        for c, w in [("Milestone", 180), ("Date", 90), ("Done", 40)]:
            self.ms_tree.heading(c, text=c)
            self.ms_tree.column(c, width=w)
        self.ms_tree.pack(fill=tk.X)
        ms_btns = ttk.Frame(ms_frame)
        ms_btns.pack(fill=tk.X, pady=2)
        ttk.Button(ms_btns, text="+ Milestone", command=self.add_milestone).pack(side=tk.LEFT, padx=2)
        ttk.Button(ms_btns, text="Toggle Done", command=self.toggle_milestone).pack(side=tk.LEFT, padx=2)
        ttk.Button(ms_btns, text="Del Milestone", command=self.delete_milestone).pack(side=tk.LEFT, padx=2)

        fin_frame = ttk.LabelFrame(task_inner, text="Invoices & Expenses", padding=5)
        fin_frame.pack(fill=tk.X, pady=(5, 5))
        fin_cols = ("Type", "Amount", "Date", "Paid")
        self.fin_tree = ttk.Treeview(fin_frame, columns=fin_cols, show="headings", height=2)
        for c, w in [("Type", 100), ("Amount", 70), ("Date", 90), ("Paid", 40)]:
            self.fin_tree.heading(c, text=c)
            self.fin_tree.column(c, width=w)
        self.fin_tree.pack(fill=tk.X)
        fin_btns = ttk.Frame(fin_frame)
        fin_btns.pack(fill=tk.X, pady=2)
        ttk.Button(fin_btns, text="+ Invoice", command=lambda: self.add_financial("invoice")).pack(side=tk.LEFT, padx=2)
        ttk.Button(fin_btns, text="+ Expense", command=lambda: self.add_financial("expense")).pack(side=tk.LEFT, padx=2)

    def pick_color(self, key):
        color = colorchooser.askcolor(title="Pick project color")[1]
        if color:
            btn = self.proj_fields[key]
            btn.config(text=f"#{color}")
            btn._color_value = color

    def sort_projects(self, col):
        if self.sort_col == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_col = col
            self.sort_reverse = False
        self.refresh_project_list()

    def _get_active_projects(self):
        return [p for p in self.data.projects if not p.get("archived")]

    def refresh_project_list(self):
        for item in self.project_tree.get_children():
            self.project_tree.delete(item)

        projects = self._get_active_projects()
        if self.sort_col:
            projects = sorted(projects, key=lambda x: str(x.get(self.sort_col, "")), reverse=self.sort_reverse)

        for proj in projects:
            status = proj.get("status", "not_started")
            self.project_tree.insert("", tk.END, values=(
                proj["name"], proj.get("code", ""), proj.get("client", ""),
                status.replace("_", " ").title(), proj.get("priority", "medium").title(),
                proj.get("due_date", ""), proj.get("risk_level", "low").title(),
                f"${proj.get('budget', 0):,.0f}"
            ), tags=(status,), iid=id(proj))

        for status, color in STATUS_COLORS.items():
            self.project_tree.tag_configure(status, foreground=color)

    def filter_projects(self):
        query = self.search_var.get().lower()
        status_f = self.filter_var.get()
        risk_f = self.risk_filter_var.get()
        tag_f = self.tag_filter_var.get().lower()

        for item in self.project_tree.get_children():
            self.project_tree.detach(item)

        for proj in self._get_active_projects():
            iid = id(proj)
            match_search = not query or any(
                query in str(v).lower() for v in [proj.get("name", ""), proj.get("code", ""),
                    proj.get("client", ""), proj.get("remarks", "")])
            match_status = status_f == "all" or proj.get("status", "") == status_f
            match_risk = risk_f == "all" or proj.get("risk_level", "") == risk_f
            match_tag = not tag_f or any(tag_f in t.lower() for t in proj.get("tags", []))

            if match_search and match_status and match_risk and match_tag:
                self.project_tree.reattach(iid, "", "end")

    def on_project_select(self, event):
        selected = self.project_tree.selection()
        if not selected:
            return
        iid = self.project_tree.item(selected[0])["iid"]
        try:
            proj = next(p for p in self._get_active_projects() if id(p) == int(iid))
        except (StopIteration, ValueError):
            return

        self._selected_project = proj

        for key, widget in self.proj_fields.items():
            val = proj.get(key, "")
            if isinstance(widget, ttk.Combobox):
                widget.set(str(val))
            elif key == "color" and isinstance(widget, ttk.Button):
                widget.config(text=f"#{val}")
                widget._color_value = val

        tags = proj.get("tags", [])
        tags_str = ", ".join(tags) if isinstance(tags, list) else str(tags)
        w = self.proj_fields.get("tags")
        if w:
            try:
                w.delete(0, tk.END)
                w.insert(0, tags_str)
            except Exception:
                pass

        budget = proj.get("budget", 0)
        w = self.proj_fields.get("budget")
        if w:
            try:
                w.delete(0, tk.END)
                w.insert(0, str(budget))
            except Exception:
                pass

        remarks = self.proj_fields.get("remarks")
        if remarks:
            try:
                remarks.delete(1.0, tk.END)
                remarks.insert(1.0, proj.get("remarks", ""))
            except Exception:
                pass

        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        for task in proj.get("tasks", []):
            self.task_tree.insert("", tk.END, values=(
                task.get("title", ""), task.get("status", "todo"),
                task.get("priority", "medium")[:3].upper(), task.get("due_date", ""),
                f"{task.get('hours', 0):.1f}", f"{task.get('estimated_hours', 0):.1f}"
            ))

        for item in self.ms_tree.get_children():
            self.ms_tree.delete(item)
        for ms in proj.get("milestones", []):
            self.ms_tree.insert("", tk.END, values=(
                ms.get("title", ""), ms.get("date", ""),
                "Yes" if ms.get("done") else "No"))

        for item in self.fin_tree.get_children():
            self.fin_tree.delete(item)
        for inv in proj.get("invoices", []):
            self.fin_tree.insert("", tk.END, values=(
                "Invoice", f"${inv.get('amount', 0):,.0f}",
                inv.get("date", ""), "Yes" if inv.get("paid") else "No"))
        for exp in proj.get("expenses", []):
            self.fin_tree.insert("", tk.END, values=(
                "Expense", f"${exp.get('amount', 0):,.0f}",
                exp.get("date", ""), ""))

    def add_project(self):
        name = self.proj_fields["name"].get().strip()
        code = self.proj_fields["code"].get().strip()
        if not name or not code:
            messagebox.showwarning("Missing", "Name and Code are required.")
            return
        budget_w = self.proj_fields.get("budget")
        try:
            budget = float(budget_w.get()) if budget_w else 0.0
        except ValueError:
            budget = 0.0

        tags_w = self.proj_fields.get("tags")
        tags = [t.strip() for t in (tags_w.get() if tags_w else "").split(",") if t.strip()]

        color_btn = self.proj_fields.get("color")
        color = getattr(color_btn, "_color_value", "#3498db") if color_btn else "#3498db"

        proj = {
            "name": name, "code": code,
            "client": self.proj_fields["client"].get().strip(),
            "status": self.proj_fields["status"].get() or "not_started",
            "priority": self.proj_fields["priority"].get() or "medium",
            "due_date": self.proj_fields["due_date"].get().strip(),
            "risk_level": self.proj_fields["risk_level"].get() or "low",
            "tags": tags, "color": color, "budget": budget,
            "remarks": self.proj_fields["remarks"].get(1.0, tk.END).strip(),
            "created_date": datetime.now().strftime("%Y-%m-%d"),
            "expenses": [], "invoices": [], "milestones": [],
            "archived": False, "tasks": []
        }
        self.data.projects.append(proj)
        self.data.log_activity("PROJECT_CREATED", f"'{name}' ({code})")
        self.data.save_data()
        self.refresh_project_list()
        self.clear_project_form()
        messagebox.showinfo("Done", f"Project '{name}' added.")

    def update_project(self):
        if not self._selected_project:
            messagebox.showwarning("Select", "Choose a project to update.")
            return
        proj = self._selected_project
        name = self.proj_fields["name"].get().strip()
        code = self.proj_fields["code"].get().strip()
        if not name or not code:
            messagebox.showwarning("Missing", "Name and Code are required.")
            return

        budget_w = self.proj_fields.get("budget")
        try:
            budget = float(budget_w.get()) if budget_w else 0.0
        except ValueError:
            budget = proj.get("budget", 0)

        tags_w = self.proj_fields.get("tags")
        tags = [t.strip() for t in (tags_w.get() if tags_w else "").split(",") if t.strip()]

        color_btn = self.proj_fields.get("color")
        color = getattr(color_btn, "_color_value", proj.get("color", "#3498db")) if color_btn else proj.get("color", "#3498db")

        old_name = proj["name"]
        proj.update({
            "name": name, "code": code,
            "client": self.proj_fields["client"].get().strip(),
            "status": self.proj_fields["status"].get() or "not_started",
            "priority": self.proj_fields["priority"].get() or "medium",
            "due_date": self.proj_fields["due_date"].get().strip(),
            "risk_level": self.proj_fields["risk_level"].get() or "low",
            "tags": tags, "color": color, "budget": budget,
            "remarks": self.proj_fields["remarks"].get(1.0, tk.END).strip()
        })
        self.data.log_activity("PROJECT_UPDATED", f"'{old_name}' -> '{name}'")
        self.data.save_data()
        self.refresh_project_list()
        self.clear_project_form()
        messagebox.showinfo("Done", "Project updated.")

    def delete_project(self):
        if not self._selected_project:
            messagebox.showwarning("Select", "Choose a project to delete.")
            return
        proj = self._selected_project
        if messagebox.askyesno("Delete", f"Permanently delete '{proj['name']}'?"):
            self.data.projects.remove(proj)
            self._selected_project = None
            self.data.log_activity("PROJECT_DELETED", f"'{proj['name']}'")
            self.data.save_data()
            self.refresh_project_list()
            self.clear_project_form()

    def archive_project(self):
        if not self._selected_project:
            messagebox.showwarning("Select", "Choose a project to archive.")
            return
        proj = self._selected_project
        proj["archived"] = True
        self._selected_project = None
        self.data.log_activity("PROJECT_ARCHIVED", f"'{proj['name']}'")
        self.data.save_data()
        self.refresh_project_list()
        self.clear_project_form()
        messagebox.showinfo("Archived", f"'{proj['name']}' archived.")

    def duplicate_project(self):
        if not self._selected_project:
            messagebox.showwarning("Select", "Choose a project to duplicate.")
            return
        orig = self._selected_project
        new_proj = copy.deepcopy(orig)
        new_proj["name"] = f"{orig['name']} (Copy)"
        new_proj["code"] = f"{orig['code']}_COPY"
        new_proj["created_date"] = datetime.now().strftime("%Y-%m-%d")
        new_proj["archived"] = False
        self.data.projects.append(new_proj)
        self.data.log_activity("PROJECT_DUPLICATED", f"'{new_proj['name']}' from '{orig['name']}'")
        self.data.save_data()
        self.refresh_project_list()
        messagebox.showinfo("Done", f"Duplicated as '{new_proj['name']}'")

    def clear_project_form(self):
        self._selected_project = None
        for key, widget in self.proj_fields.items():
            if isinstance(widget, ttk.Combobox):
                widget.set("")
            elif isinstance(widget, ttk.Button):
                if key == "color":
                    widget.config(text="#3498db")
                    widget._color_value = "#3498db"
                continue
            else:
                try:
                    widget.delete(1.0, tk.END)
                except Exception:
                    widget.delete(0, tk.END)
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        for item in self.ms_tree.get_children():
            self.ms_tree.delete(item)
        for item in self.fin_tree.get_children():
            self.fin_tree.delete(item)

    # ---------------------------- TASKS ----------------------------
    def add_task(self):
        if not self._selected_project:
            messagebox.showwarning("Select", "Choose a project first.")
            return
        proj = self._selected_project
        dlg = tk.Toplevel(self.root)
        dlg.title("Add Task")
        dlg.geometry("400x300")
        dlg.transient(self.root)
        dlg.grab_set()

        fields = {}
        labels_defaults = [
            ("Title:", ""), ("Status:", "todo", ["todo", "doing", "done"]),
            ("Priority:", "medium", ["low", "medium", "high", "urgent"]),
            ("Due Date (YYYY-MM-DD):", ""), ("Estimated Hours:", "0"),
            ("Blocked By (task title):", "")
        ]

        for i, ld in enumerate(labels_defaults):
            label = ld[0]
            default = ld[1]
            ttk.Label(dlg, text=label).grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
            key = label.replace(":", "").lower().replace(" ", " ")

            if len(ld) == 3:
                cb = ttk.Combobox(dlg, values=ld[2], width=20, state="readonly")
                cb.set(default)
                cb.grid(row=i, column=1, padx=10, pady=5)
                fields[key] = cb
            else:
                ent = ttk.Entry(dlg, width=22)
                ent.insert(0, default)
                ent.grid(row=i, column=1, padx=10, pady=5)
                fields[key] = ent

        def save():
            title = fields["title"].get().strip()
            if not title:
                messagebox.showwarning("Missing", "Title required.")
                return
            try:
                est = float(fields.get("estimated hours", ttk.Entry(dlg)).get() or "0")
            except ValueError:
                est = 0.0

            task = {
                "title": title,
                "status": fields["status"].get(),
                "priority": fields["priority"].get(),
                "due_date": fields.get("due date (yyyy-mm-dd)", ttk.Entry(dlg)).get().strip(),
                "estimated_hours": est, "hours": 0,
                "subtasks": [], "comments": [],
                "blocked_by": fields.get("blocked by (task title)", ttk.Entry(dlg)).get().strip(),
                "attachments": []
            }
            proj.setdefault("tasks", []).append(task)
            self.data.log_activity("TASK_ADDED", f"'{title}' in '{proj['name']}'")
            self.data.save_data()
            self.on_project_select(None)
            dlg.destroy()

        ttk.Button(dlg, text="Add Task", command=save).grid(row=len(labels_defaults), column=0,
                                                              columnspan=2, pady=15)

    def edit_task_dialog(self):
        task_sel = self.task_tree.selection()
        if not task_sel or not self._selected_project:
            messagebox.showwarning("Select", "Choose a project and task to edit.")
            return
        task_idx = self.task_tree.index(task_sel[0])
        proj = self._selected_project
        task = proj["tasks"][task_idx]

        dlg = tk.Toplevel(self.root)
        dlg.title("Edit Task")
        dlg.geometry("450x400")
        dlg.transient(self.root)
        dlg.grab_set()

        fields_data = {
            "title": task.get("title", ""), "status": task.get("status", "todo"),
            "priority": task.get("priority", "medium"), "due_date": task.get("due_date", ""),
            "hours": str(task.get("hours", 0)), "estimated_hours": str(task.get("estimated_hours", 0)),
            "blocked_by": task.get("blocked_by", "")
        }

        row = 0
        widgets = {}
        for label, key in [("Title:", "title"), ("Status:", "status"), ("Priority:", "priority"),
                            ("Due Date:", "due_date"), ("Hours Logged:", "hours"),
                            ("Est. Hours:", "estimated_hours"), ("Blocked By:", "blocked_by")]:
            ttk.Label(dlg, text=label).grid(row=row, column=0, sticky=tk.W, padx=10, pady=4)
            if key == "status":
                cb = ttk.Combobox(dlg, values=["todo", "doing", "done"], width=20, state="readonly")
                cb.set(fields_data[key])
                cb.grid(row=row, column=1, padx=10, pady=4)
                widgets[key] = cb
            elif key == "priority":
                cb = ttk.Combobox(dlg, values=["low", "medium", "high", "urgent"], width=20, state="readonly")
                cb.set(fields_data[key])
                cb.grid(row=row, column=1, padx=10, pady=4)
                widgets[key] = cb
            else:
                ent = ttk.Entry(dlg, width=22)
                ent.insert(0, fields_data[key])
                ent.grid(row=row, column=1, padx=10, pady=4)
                widgets[key] = ent
            row += 1

        ttk.Label(dlg, text="Comments:").grid(row=row, column=0, sticky=tk.NW, padx=10, pady=4)
        row += 1
        comments_txt = tk.Text(dlg, height=3, wrap=tk.WORD, font=("Consolas", 9))
        comments_txt.grid(row=row, column=0, columnspan=2, padx=10, pady=4)
        for c in task.get("comments", []):
            comments_txt.insert(tk.END, f"[{c.get('time','')}] {c.get('text','')}  \n")
        row += 1

        def add_comment():
            c = simpledialog.askstring("Comment", "Add a comment:")
            if c:
                task.setdefault("comments", []).append({
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M"), "text": c
                })

        ttk.Button(dlg, text="Add Comment", command=add_comment).grid(row=row, column=0, columnspan=2, pady=4)
        row += 1

        ttk.Label(dlg, text="Subtasks:").grid(row=row, column=0, sticky=tk.NW, padx=10, pady=4)
        row += 1
        sub_tree = ttk.Treeview(dlg, columns=("Title", "Status"), show="headings", height=2)
        for c, w in [("Title", 200), ("Status", 60)]:
            sub_tree.heading(c, text=c)
            sub_tree.column(c, width=w)
        sub_tree.grid(row=row, column=0, columnspan=2, padx=10, pady=4)
        for st in task.get("subtasks", []):
            sub_tree.insert("", tk.END, values=(st.get("title", ""), st.get("status", "")))
        row += 1

        def add_subtask():
            st_title = simpledialog.askstring("Subtask", "Subtask title:")
            if st_title:
                task.setdefault("subtasks", []).append({
                    "title": st_title.strip(), "status": "todo", "hours": 0,
                    "estimated_hours": 0, "due_date": "", "priority": "medium",
                    "comments": [], "subtasks": [], "blocked_by": "", "attachments": []
                })
                for item in sub_tree.get_children():
                    sub_tree.delete(item)
                for st in task["subtasks"]:
                    sub_tree.insert("", tk.END, values=(st["title"], st["status"]))

        ttk.Button(dlg, text="+ Subtask", command=add_subtask).grid(row=row, column=0, padx=10, pady=4)
        row += 1

        def save():
            title = widgets["title"].get().strip()
            if not title:
                messagebox.showwarning("Missing", "Title required.")
                return
            try:
                hours = float(widgets["hours"].get() or "0")
            except ValueError:
                hours = task.get("hours", 0)
            try:
                est = float(widgets["estimated_hours"].get() or "0")
            except ValueError:
                est = task.get("estimated_hours", 0)

            old_title = task.get("title", "")
            task.update({
                "title": title, "status": widgets["status"].get(),
                "priority": widgets["priority"].get(),
                "due_date": widgets["due_date"].get().strip(),
                "hours": hours, "estimated_hours": est,
                "blocked_by": widgets["blocked_by"].get().strip()
            })
            self.data.log_activity("TASK_UPDATED", f"'{title}' in '{proj['name']}'")
            self.data.save_data()
            self.on_project_select(None)
            dlg.destroy()

        ttk.Button(dlg, text="Save", command=save).grid(row=row, column=0, columnspan=2, pady=10)

    def delete_task(self):
        task_sel = self.task_tree.selection()
        if not task_sel or not self._selected_project:
            messagebox.showwarning("Select", "Choose a project and task.")
            return
        task_idx = self.task_tree.index(task_sel[0])
        proj = self._selected_project
        title = proj["tasks"][task_idx]["title"]
        if messagebox.askyesno("Delete", f"Delete task '{title}'?"):
            proj["tasks"].pop(task_idx)
            self.data.log_activity("TASK_DELETED", f"'{title}' from '{proj['name']}'")
            self.data.save_data()
            self.on_project_select(None)

    def batch_tasks(self):
        if not self._selected_project:
            messagebox.showwarning("Select", "Choose a project.")
            return
        proj = self._selected_project
        action = simpledialog.askstring("Batch Operations",
            "Enter action:\n  set_status=<todo/doing/done>\n"
            "  set_priority=<low/medium/high/urgent>\n  delete_all")
        if not action:
            return
        if action.startswith("set_status="):
            new_status = action.split("=", 1)[1].strip()
            for t in proj["tasks"]:
                t["status"] = new_status
            self.data.log_activity("BATCH_UPDATE", f"Set all tasks to '{new_status}' in '{proj['name']}'")
        elif action.startswith("set_priority="):
            new_pri = action.split("=", 1)[1].strip()
            for t in proj["tasks"]:
                t["priority"] = new_pri
            self.data.log_activity("BATCH_UPDATE", f"Set all priorities to '{new_pri}' in '{proj['name']}'")
        elif action == "delete_all":
            if messagebox.askyesno("Confirm", "Delete ALL tasks in this project?"):
                count = len(proj["tasks"])
                proj["tasks"] = []
                self.data.log_activity("BATCH_DELETE", f"Deleted {count} tasks from '{proj['name']}'")

        self.data.save_data()
        self.on_project_select(None)

    # ---------------------------- MILESTONES ----------------------------
    def add_milestone(self):
        if not self._selected_project:
            messagebox.showwarning("Select", "Choose a project.")
            return
        proj = self._selected_project
        title = simpledialog.askstring("Milestone", "Milestone name:")
        if not title:
            return
        ms_date = simpledialog.askstring("Date", "Target date (YYYY-MM-DD):")
        proj.setdefault("milestones", []).append({
            "title": title.strip(), "date": (ms_date or "").strip(), "done": False
        })
        self.data.log_activity("MILESTONE_ADDED", f"'{title}' on '{proj['name']}'")
        self.data.save_data()
        self.on_project_select(None)

    def toggle_milestone(self):
        ms_sel = self.ms_tree.selection()
        if not ms_sel or not self._selected_project:
            return
        idx_ms = self.ms_tree.index(ms_sel[0])
        proj = self._selected_project
        proj["milestones"][idx_ms]["done"] = not proj["milestones"][idx_ms].get("done", False)
        self.data.save_data()
        self.on_project_select(None)

    def delete_milestone(self):
        ms_sel = self.ms_tree.selection()
        if not ms_sel or not self._selected_project:
            return
        idx_ms = self.ms_tree.index(ms_sel[0])
        proj = self._selected_project
        title = proj["milestones"][idx_ms]["title"]
        proj["milestones"].pop(idx_ms)
        self.data.log_activity("MILESTONE_DELETED", f"'{title}' from '{proj['name']}'")
        self.data.save_data()
        self.on_project_select(None)

    # ---------------------------- FINANCIALS ----------------------------
    def add_financial(self, fin_type):
        if not self._selected_project:
            messagebox.showwarning("Select", "Choose a project.")
            return
        proj = self._selected_project
        amount = simpledialog.askstring("Amount", f"{fin_type.title()} amount ($):")
        if not amount:
            return
        try:
            amt = float(amount)
        except ValueError:
            messagebox.showwarning("Invalid", "Enter a valid number.")
            return
        fdate = simpledialog.askstring("Date", "Date (YYYY-MM-DD):") or datetime.now().strftime("%Y-%m-%d")
        paid = False
        if fin_type == "invoice":
            paid = messagebox.askyesno("Paid?", "Has this invoice been paid?")

        if fin_type == "invoice":
            proj.setdefault("invoices", []).append({"amount": amt, "date": fdate.strip(), "paid": paid})
        else:
            desc = simpledialog.askstring("Description", "Expense description:")
            proj.setdefault("expenses", []).append({
                "amount": amt, "date": fdate.strip(), "description": (desc or "").strip()
            })

        self.data.log_activity(f"{fin_type.upper()}_ADDED", f"${amt} on '{proj['name']}'")
        self.data.save_data()
        self.on_project_select(None)

    # ---------------------------- KANBAN BOARD ----------------------------
    def setup_kanban(self):
        top_bar = ttk.Frame(self.kanban_frame)
        top_bar.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(top_bar, text="Kanban Board", font=("Arial", 12, "bold")).pack(side=tk.LEFT)
        ttk.Button(top_bar, text="Refresh", command=self.refresh_kanban).pack(side=tk.RIGHT)
        ttk.Button(top_bar, text="Add Card", command=self.kanban_add_card).pack(side=tk.RIGHT, padx=5)

        self.kanban_columns = {
            "todo": {"title": "TO DO", "color": "#e74c3c"},
            "doing": {"title": "IN PROGRESS", "color": "#f39c12"},
            "done": {"title": "DONE", "color": "#2ecc71"}
        }
        self.kanban_containers = {}

        cols_frame = ttk.Frame(self.kanban_frame)
        cols_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        for i, (key, info) in enumerate(self.kanban_columns.items()):
            col_frame = ttk.LabelFrame(cols_frame, text=info["title"], padding=5)
            col_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

            bg_color = "#f0f0f0" if self.theme == "light" else "#1e1e1e"
            canvas = tk.Canvas(col_frame, height=600, bg=bg_color)
            scrollbar = ttk.Scrollbar(col_frame, orient="vertical", command=canvas.yview)
            inner = ttk.Frame(canvas)

            inner.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
            canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            self.kanban_containers[key] = inner

        self.refresh_kanban()

    def refresh_kanban(self):
        for key in self.kanban_containers:
            container = self.kanban_containers[key]
            for w in container.winfo_children():
                w.destroy()

            active = [p for p in self.data.projects if not p.get("archived")]
            for proj in active:
                for task in proj.get("tasks", []):
                    if task.get("status") == key:
                        card_bg = "white" if self.theme == "light" else "#2d2d2d"
                        card = tk.Frame(container, bg=card_bg, relief=tk.RAISED, bd=1)
                        card.pack(pady=3, padx=2, fill=tk.X)

                        color_block = tk.Frame(card, width=5, height=50, bg=proj.get("color", "#3498db"))
                        color_block.pack(side=tk.LEFT, fill=tk.Y)
                        color_block.pack_propagate(False)

                        content = ttk.Frame(card)
                        content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

                        ttk.Label(content, text=task.get("title", ""), font=("Arial", 10, "bold")).pack(anchor=tk.W)
                        ttk.Label(content, text=f"Project: {proj['name']}", font=("Arial", 8), foreground="gray").pack(anchor=tk.W)

                        tags_str = f"Pri: {task.get('priority', 'medium')} | "
                        if task.get("due_date"):
                            tags_str += f"Due: {task['due_date']} | "
                        est = task.get("estimated_hours", 0)
                        hrs = task.get("hours", 0)
                        tags_str += f"{hrs:.0f}/{est:.0f}h"
                        if task.get("blocked_by"):
                            tags_str += f" | Blocked: {task['blocked_by']}"
                        ttk.Label(content, text=tags_str, font=("Arial", 7), foreground="#666").pack(anchor=tk.W)

                        has_sub = len(task.get("subtasks", []))
                        if has_sub:
                            done_sub = sum(1 for s in task["subtasks"] if s.get("status") == "done")
                            ttk.Label(content, text=f"Subtasks: {done_sub}/{has_sub}",
                                      font=("Arial", 7), foreground="#2980b9").pack(anchor=tk.W)

                        card.bind("<Double-1>", lambda e, t=task, p=proj: self.edit_task_inline(t, p))

    def kanban_add_card(self):
        self.notebook.select(self.projects_frame)
        self.add_task()

    def edit_task_inline(self, task, proj):
        new_title = simpledialog.askstring("Edit Task", "Task title:", initialvalue=task.get("title", ""))
        if new_title:
            task["title"] = new_title.strip()
        new_status = simpledialog.askstring("Change Status", "Status (todo/doing/done):",
                                             initialvalue=task.get("status", ""))
        if new_status:
            task["status"] = new_status.strip()
        try:
            h = simpledialog.askstring("Hours", "Hours logged:", initialvalue=str(task.get("hours", 0)))
            if h:
                task["hours"] = float(h)
        except ValueError:
            pass

        self.data.log_activity("TASK_UPDATED", f"'{task['title']}' in '{proj['name']}'")
        self.data.save_data()
        self.refresh_kanban()
        self.refresh_project_list()

    # ---------------------------- GANTT TIMELINE ----------------------------
    def setup_gantt(self):
        top_bar = ttk.Frame(self.gantt_frame)
        top_bar.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(top_bar, text="Gantt Timeline", font=("Arial", 12, "bold")).pack(side=tk.LEFT)
        ttk.Button(top_bar, text="Refresh", command=self.refresh_gantt).pack(side=tk.RIGHT)

        self.gantt_canvas = tk.Canvas(self.gantt_frame, bg="white")
        self.gantt_canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.gantt_canvas.bind("<Configure>", lambda e: self.refresh_gantt())

    def refresh_gantt(self):
        self.gantt_canvas.delete("all")
        w = self.gantt_canvas.winfo_width()
        if w < 100:
            return

        active = [p for p in self.data.projects if not p.get("archived")]
        if not active:
            return

        all_dates = []
        for proj in active:
            if proj.get("due_date"):
                try:
                    all_dates.append(datetime.strptime(proj["due_date"], "%Y-%m-%d").date())
                except ValueError:
                    pass
            for t in proj.get("tasks", []):
                if t.get("due_date"):
                    try:
                        all_dates.append(datetime.strptime(t["due_date"], "%Y-%m-%d").date())
                    except ValueError:
                        pass

        if not all_dates:
            return

        min_date = min(all_dates) - timedelta(days=7)
        max_date = max(all_dates) + timedelta(days=7)
        total_days = (max_date - min_date).days or 1

        header_h = 35
        row_h = 28
        label_w = 200
        chart_w = w - label_w

        self.gantt_canvas.create_rectangle(0, 0, w, header_h, fill="#2c3e50", outline="")
        self.gantt_canvas.create_text(label_w // 2, header_h // 2, text="Project / Task",
                                       fill="white", font=("Arial", 10, "bold"))

        step = max(1, total_days // 30)
        step_date = min_date
        while step_date <= max_date:
            x = label_w + (step_date - min_date).days / total_days * chart_w
            self.gantt_canvas.create_line(x, header_h, x, header_h + len(active) * 3 * row_h, fill="#ddd", width=1)
            self.gantt_canvas.create_text(x, header_h // 2, text=step_date.strftime("%m/%d"),
                                           anchor="s", fill="white", font=("Arial", 7))
            step_date += timedelta(days=step)

        today_x = label_w + (date.today() - min_date).days / total_days * chart_w
        if label_w <= today_x <= w:
            self.gantt_canvas.create_line(today_x, header_h, today_x,
                                           header_h + len(active) * 3 * row_h, fill="red", width=2, dash=(4, 4))
            self.gantt_canvas.create_text(today_x, header_h - 2, text="TODAY",
                                           anchor="s", fill="red", font=("Arial", 7, "bold"))

        y = header_h + 5
        for proj in active:
            proj_color = proj.get("color", "#3498db")
            self.gantt_canvas.create_text(5, y + row_h // 2, text=f"> {proj['name']} ({proj.get('code','')})",
                                           anchor="w", font=("Arial", 9, "bold"), fill=proj_color)

            tasks = proj.get("tasks", [])
            if proj.get("due_date"):
                try:
                    pd = datetime.strptime(proj["due_date"], "%Y-%m-%d").date()
                    start_x = label_w + max(0, (pd - timedelta(days=14) - min_date).days) / total_days * chart_w
                    end_x = label_w + (pd - min_date).days / total_days * chart_w
                    self.gantt_canvas.create_rectangle(start_x, y + 2, end_x, y + row_h - 4,
                                                        fill=proj_color, outline="", width=0)
                except ValueError:
                    pass

            y += row_h
            for task in tasks[:5]:
                bar_color = "#ccc"
                if task.get("status") == "done":
                    bar_color = "#2ecc71"
                elif task.get("status") == "doing":
                    bar_color = "#f39c12"

                self.gantt_canvas.create_text(10, y + row_h // 2, text=f"  - {task.get('title','')}",
                                               anchor="w", font=("Arial", 8), fill="#555")

                if task.get("due_date"):
                    try:
                        td = datetime.strptime(task["due_date"], "%Y-%m-%d").date()
                        est = task.get("estimated_hours", 0) or 8
                        task_days = max(1, int(est / 8))
                        start_d = td - timedelta(days=task_days)
                        sx = label_w + max(0, (start_d - min_date).days) / total_days * chart_w
                        ex = label_w + (td - min_date).days / total_days * chart_w
                        self.gantt_canvas.create_rectangle(sx, y + 4, ex, y + row_h - 6,
                                                            fill=bar_color, outline="", width=0)
                    except ValueError:
                        pass
                y += row_h

    # ---------------------------- CALENDAR VIEW ----------------------------
    def setup_calendar(self):
        top_bar = ttk.Frame(self.calendar_frame)
        top_bar.pack(fill=tk.X, padx=10, pady=5)

        self.cal_year_var = tk.StringVar(value=str(datetime.now().year))
        self.cal_month_var = tk.StringVar(value=str(datetime.now().month))

        ttk.Label(top_bar, text="Year:").pack(side=tk.LEFT)
        yy = ttk.Combobox(top_bar, textvariable=self.cal_year_var, width=5, state="readonly",
                          values=list(range(2024, 2031)))
        yy.pack(side=tk.LEFT, padx=5)

        ttk.Label(top_bar, text="Month:").pack(side=tk.LEFT)
        mm = ttk.Combobox(top_bar, textvariable=self.cal_month_var, width=3, state="readonly",
                          values=list(range(1, 13)))
        mm.pack(side=tk.LEFT, padx=5)

        for v in [self.cal_year_var, self.cal_month_var]:
            v.trace("w", lambda *_: self.refresh_calendar_view())

        ttk.Button(top_bar, text="Today", command=lambda: (
            self.cal_year_var.set(str(datetime.now().year)),
            self.cal_month_var.set(str(datetime.now().month)),
            self.refresh_calendar_view())).pack(side=tk.RIGHT)

        self.cal_frame = ttk.Frame(self.calendar_frame)
        self.cal_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    def refresh_calendar_view(self):
        for w in self.cal_frame.winfo_children():
            w.destroy()

        try:
            year = int(self.cal_year_var.get())
            month = int(self.cal_month_var.get())
        except ValueError:
            return

        cal_grid = ttk.Frame(self.cal_frame)
        cal_grid.pack(fill=tk.BOTH, expand=True)

        if month == 12:
            next_month = date(year + 1, 1, 1)
        else:
            next_month = date(year, month + 1, 1)
        days_in_month = (next_month - timedelta(days=1)).day
        first_day = date(year, month, 1).weekday()

        headers = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for i, h in enumerate(headers):
            ttk.Label(cal_grid, text=h, font=("Arial", 10, "bold"), relief=tk.RIDGE).grid(
                row=0, column=i, padx=2, pady=2, sticky="nsew", ipadx=15, ipady=8)

        active = [p for p in self.data.projects if not p.get("archived")]
        task_map = {}
        for proj in active:
            if proj.get("due_date"):
                try:
                    d = datetime.strptime(proj["due_date"], "%Y-%m-%d").date()
                    if d.year == year and d.month == month:
                        task_map.setdefault(d.day, []).append(f"[PROJ] {proj['name']}")
                except ValueError:
                    pass
            for t in proj.get("tasks", []):
                if t.get("due_date"):
                    try:
                        d = datetime.strptime(t["due_date"], "%Y-%m-%d").date()
                        if d.year == year and d.month == month:
                            task_map.setdefault(d.day, []).append(f"[TASK] {t['title']} ({proj['name']})")
                    except ValueError:
                        pass

        today = date.today()
        day = 1
        for row in range(1, 8):
            for col in range(7):
                cell_idx = (row - 1) * 7 + col
                if cell_idx < first_day or day > days_in_month:
                    cell = tk.Frame(cal_grid, bg="#f0f0f0", relief=tk.RIDGE, bd=1)
                    cell.grid(row=row, column=col, padx=1, pady=1, sticky="nsew", ipadx=5, ipady=20)
                else:
                    is_today = (today.year == year and today.month == month and today.day == day)
                    bg = "#ffeb3b" if is_today else "white"
                    tasks = task_map.get(day, [])
                    is_overdue = False
                    if not is_today and tasks:
                        check_date = date(year, month, day)
                        if check_date < today:
                            bg = "#ffcdd2"
                            is_overdue = True

                    cell = tk.Frame(cal_grid, bg=bg, relief=tk.RIDGE, bd=1)
                    cell.grid(row=row, column=col, padx=1, pady=1, sticky="nsew", ipadx=5, ipady=20)

                    tk.Label(cell, text=str(day), font=("Arial", 9, "bold" if is_today else ""),
                              bg=bg, fg="#c62828" if is_overdue else "#333").pack(anchor=tk.N)

                    if tasks:
                        txt = tk.Text(cell, height=min(len(tasks), 4), width=10,
                                       font=("Arial", 6), bg=bg, bd=0, padx=1)
                        txt.pack(fill=tk.BOTH, expand=True)
                        for t in tasks[:4]:
                            txt.insert(tk.END, f"- {t}\n")
                        if len(tasks) > 4:
                            txt.insert(tk.END, f"...+{len(tasks)-4}")

                    cell.bind("<Double-1>", lambda e, d=day: self.show_day_tasks(d, year, month))
                    day += 1

        for c in range(7):
            cal_grid.columnconfigure(c, weight=1)
        for r in range(1, 8):
            cal_grid.rowconfigure(r, weight=1)

    def show_day_tasks(self, day, year, month):
        active = [p for p in self.data.projects if not p.get("archived")]
        tasks = []
        check_date = date(year, month, day)
        for proj in active:
            if proj.get("due_date"):
                try:
                    d = datetime.strptime(proj["due_date"], "%Y-%m-%d").date()
                    if d == check_date:
                        tasks.append(f"PROJECT: {proj['name']} ({proj.get('status','')})")
                except ValueError:
                    pass
            for t in proj.get("tasks", []):
                if t.get("due_date"):
                    try:
                        d = datetime.strptime(t["due_date"], "%Y-%m-%d").date()
                        if d == check_date:
                            tasks.append(f"TASK: {t['title']} [{t.get('status','')}] - {proj['name']}")
                    except ValueError:
                        pass
        if tasks:
            messagebox.showinfo(f"Tasks for {check_date}", "\n".join(tasks))
        else:
            messagebox.showinfo(f"No tasks for {check_date}", "Nothing due on this day.")

    # ---------------------------- TIMESHEET ----------------------------
    def setup_timesheet(self):
        top_bar = ttk.Frame(self.timesheet_frame)
        top_bar.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(top_bar, text="Weekly Timesheet", font=("Arial", 12, "bold")).pack(side=tk.LEFT)

        self.ts_week_var = tk.StringVar()
        ttk.Label(top_bar, textvariable=self.ts_week_var).pack(side=tk.LEFT, padx=10)

        ttk.Button(top_bar, text="Previous Week", command=lambda: self.nav_week(-1)).pack(side=tk.LEFT, padx=3)
        ttk.Button(top_bar, text="This Week", command=self.update_ts_week).pack(side=tk.LEFT, padx=3)
        ttk.Button(top_bar, text="Next Week", command=lambda: self.nav_week(1)).pack(side=tk.LEFT, padx=3)
        ttk.Button(top_bar, text="Export CSV", command=self.export_timesheet_csv).pack(side=tk.RIGHT)

        self.ts_frame = ttk.Frame(self.timesheet_frame)
        self.ts_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.update_ts_week()

    def update_ts_week(self):
        today = date.today()
        start = today - timedelta(days=today.weekday())
        end = start + timedelta(days=6)
        self._ts_start = start
        self.ts_week_var.set(f"Week: {start.strftime('%Y-%m-%d')} to {end.strftime('%Y-%m-%d')}")
        self.refresh_timesheet()

    def nav_week(self, direction):
        if not hasattr(self, "_ts_start"):
            self.update_ts_week()
            return
        self._ts_start += timedelta(weeks=direction)
        end = self._ts_start + timedelta(days=6)
        self.ts_week_var.set(f"Week: {self._ts_start.strftime('%Y-%m-%d')} to {end.strftime('%Y-%m-%d')}")
        self.refresh_timesheet()

    def refresh_timesheet(self):
        for w in self.ts_frame.winfo_children():
            w.destroy()

        if not hasattr(self, "_ts_start"):
            return

        start = self._ts_start
        active = [p for p in self.data.projects if not p.get("archived")]

        headers = ["Project", "Task"] + [(start + timedelta(days=i)).strftime("%a %m/%d") for i in range(7)] + ["Total"]
        col_ids = [str(i) for i in range(len(headers))]
        tree = ttk.Treeview(self.ts_frame, columns=col_ids, show="headings")
        for i, h in enumerate(headers):
            tree.heading(col_ids[i], text=h)
            tree.column(col_ids[i], width=80 if i > 1 else 150)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(self.ts_frame, orient="vertical", command=tree.yview).pack(side=tk.RIGHT, fill=tk.Y)

        for proj in active:
            for task in proj.get("tasks", []):
                daily = [0.0] * 7
                hours = task.get("hours", 0)
                if task.get("due_date"):
                    try:
                        dd = datetime.strptime(task["due_date"], "%Y-%m-%d").date()
                        diff = (dd - start).days
                        if 0 <= diff < 7:
                            daily[diff] = hours
                    except ValueError:
                        daily[-1] = hours
                else:
                    daily[-1] = hours

                vals = [proj["name"], task.get("title", "")] + [f"{d:.1f}" for d in daily] + [f"{sum(daily):.1f}"]
                tree.insert("", tk.END, values=vals)

    def export_timesheet_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                            initialfile="timesheet.csv")
        if not path:
            return
        start = getattr(self, "_ts_start", date.today() - timedelta(days=date.today().weekday()))
        with open(path, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Project", "Task"] + [(start + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(7)] + ["Total Hours"])
            for proj in [p for p in self.data.projects if not p.get("archived")]:
                for task in proj.get("tasks", []):
                    daily = [0.0] * 7
                    hours = task.get("hours", 0)
                    if task.get("due_date"):
                        try:
                            dd = datetime.strptime(task["due_date"], "%Y-%m-%d").date()
                            diff = (dd - start).days
                            if 0 <= diff < 7:
                                daily[diff] = hours
                        except ValueError:
                            daily[-1] = hours
                    else:
                        daily[-1] = hours
                    w.writerow([proj["name"], task.get("title", "")] + [f"{d:.1f}" for d in daily] + [f"{sum(daily):.1f}"])
        messagebox.showinfo("Exported", f"Timesheet saved to {path}")

    # ---------------------------- FINANCIALS TAB ----------------------------
    def setup_financials(self):
        top = ttk.Frame(self.financial_frame)
        top.pack(fill=tk.X, padx=20, pady=15)
        ttk.Label(top, text="Financial Overview", font=("Arial", 16, "bold")).pack(anchor=tk.W)
        ttk.Button(top, text="Refresh", command=self.refresh_financials).pack(side=tk.RIGHT)

        self.fin_summary = ttk.Frame(self.financial_frame)
        self.fin_summary.pack(fill=tk.X, padx=20, pady=10)

        self.fin_table_frame = ttk.LabelFrame(self.financial_frame, text="Project Financial Details", padding=10)
        self.fin_table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        fin_cols = ("Project", "Budget", "Invoiced", "Paid", "Outstanding", "Expenses", "Profit")
        self.fin_tree = ttk.Treeview(self.fin_table_frame, columns=fin_cols, show="headings", height=12)
        for c, w in [("Project", 180), ("Budget", 90), ("Invoiced", 90), ("Paid", 90),
                      ("Outstanding", 90), ("Expenses", 90), ("Profit", 90)]:
            self.fin_tree.heading(c, text=c)
            self.fin_tree.column(c, width=w)
        self.fin_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(self.fin_table_frame, orient="vertical", command=self.fin_tree.yview).pack(side=tk.RIGHT, fill=tk.Y)

        self.refresh_financials()

    def refresh_financials(self):
        for w in self.fin_summary.winfo_children():
            w.destroy()

        active = [p for p in self.data.projects if not p.get("archived")]
        total_budget = sum(p.get("budget", 0) for p in active)
        total_invoiced = sum(sum(i.get("amount", 0) for i in p.get("invoices", [])) for p in active)
        total_paid = sum(sum(i.get("amount", 0) for i in p.get("invoices", []) if i.get("paid")) for p in active)
        total_expenses = sum(sum(e.get("amount", 0) for e in p.get("expenses", [])) for p in active)
        profit = total_paid - total_expenses

        cards = [
            ("Total Budget", f"${total_budget:,.0f}", "#3498db"),
            ("Total Invoiced", f"${total_invoiced:,.0f}", "#27ae60"),
            ("Total Paid", f"${total_paid:,.0f}", "#2ecc71"),
            ("Outstanding", f"${total_invoiced - total_paid:,.0f}", "#e67e22"),
            ("Total Expenses", f"${total_expenses:,.0f}", "#c0392b"),
            ("Net Profit", f"${profit:,.0f}", "#8e44ad" if profit >= 0 else "#e74c3c"),
        ]

        for i, (label, value, color) in enumerate(cards):
            card = tk.Frame(self.fin_summary, bg=color, height=60)
            card.grid(row=0, column=i, padx=8, pady=5, sticky="nsew")
            tk.Label(card, text=value, font=("Arial", 16, "bold"), bg=color, fg="white").pack(pady=(8, 0))
            tk.Label(card, text=label, font=("Arial", 8), bg=color, fg="white").pack()

        for item in self.fin_tree.get_children():
            self.fin_tree.delete(item)

        for proj in active:
            budget = proj.get("budget", 0)
            invoiced = sum(i.get("amount", 0) for i in proj.get("invoices", []))
            paid = sum(i.get("amount", 0) for i in proj.get("invoices", []) if i.get("paid"))
            expenses = sum(e.get("amount", 0) for e in proj.get("expenses", []))
            outstanding = invoiced - paid
            profit = paid - expenses

            self.fin_tree.insert("", tk.END, values=(
                proj["name"], f"${budget:,.0f}", f"${invoiced:,.0f}",
                f"${paid:,.0f}", f"${outstanding:,.0f}",
                f"${expenses:,.0f}", f"${profit:,.0f}"
            ))

    # ---------------------------- ACTIVITY LOG ----------------------------
    def setup_activity_log(self):
        top = ttk.Frame(self.activity_frame)
        top.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(top, text="Activity Log", font=("Arial", 12, "bold")).pack(side=tk.LEFT)
        ttk.Button(top, text="Refresh", command=self.refresh_activity_log).pack(side=tk.RIGHT)
        ttk.Button(top, text="Clear Log", command=self.clear_activity_log).pack(side=tk.RIGHT, padx=5)

        log_frame = ttk.Frame(self.activity_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("Time", "Action", "Details")
        self.log_tree = ttk.Treeview(log_frame, columns=cols, show="headings", height=20)
        for c, w in [("Time", 170), ("Action", 150), ("Details", 500)]:
            self.log_tree.heading(c, text=c)
            self.log_tree.column(c, width=w)
        self.log_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Scrollbar(log_frame, orient="vertical", command=self.log_tree.yview).pack(side=tk.RIGHT, fill=tk.Y)

        self.refresh_activity_log()

    def refresh_activity_log(self):
        for item in self.log_tree.get_children():
            self.log_tree.delete(item)
        for entry in reversed(self.data.activity_log[-100:]):
            self.log_tree.insert("", tk.END, values=(
                entry.get("timestamp", ""), entry.get("action", ""), entry.get("details", "")
            ))

    def clear_activity_log(self):
        if messagebox.askyesno("Clear Log", "Clear all activity log entries?"):
            self.data.activity_log = []
            self.data.save_data()
            self.refresh_activity_log()

    # ---------------------------- PHRASES TAB ----------------------------
    def setup_phrases_tab(self):
        left_frame = ttk.Frame(self.phrases_frame, width=320)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        left_frame.pack_propagate(False)

        ttk.Label(left_frame, text="Categories:", font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=5)
        categories_frame = ttk.Frame(left_frame)
        categories_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        canvas = tk.Canvas(categories_frame, height=400)
        scrollbar = ttk.Scrollbar(categories_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.category_widgets = {}
        self.refresh_categories_list()

        ttk.Separator(left_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="Add Category", command=self.add_category, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Delete Category", command=self.delete_category, width=15).pack(side=tk.LEFT, padx=2)

        instructions = "Instructions:\nClick a category to edit phrases\nEach phrase on a new line with '- ' prefix\nClick SAVE to save changes"
        ttk.Label(left_frame, text=instructions, justify=tk.LEFT, foreground="gray", font=("Arial", 9)).pack(anchor=tk.W, pady=10)

        right_frame = ttk.Frame(self.phrases_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        ttk.Label(right_frame, text="Edit Phrases (one per line, start with '- '):", font=("Arial", 11, "bold")).pack(anchor=tk.W)
        self.phrases_text = tk.Text(right_frame, wrap=tk.WORD, height=30, font=("Consolas", 10))
        self.phrases_text.pack(fill=tk.BOTH, expand=True)
        scrollbar_text = ttk.Scrollbar(self.phrases_text, command=self.phrases_text.yview)
        self.phrases_text.configure(yscrollcommand=scrollbar_text.set)
        scrollbar_text.pack(side=tk.RIGHT, fill=tk.Y)
        btn_frame2 = ttk.Frame(right_frame)
        btn_frame2.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame2, text="SAVE Changes", command=self.save_current_phrases, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame2, text="Refresh", command=self.refresh_current_display, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame2, text="Clear All", command=self.clear_text_area, width=12).pack(side=tk.LEFT, padx=5)

    def refresh_categories_list(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.category_checkboxes.clear()
        self.category_widgets.clear()
        for cat in self.data.categories:
            frame = ttk.Frame(self.scrollable_frame)
            frame.pack(fill=tk.X, pady=2)
            var = tk.BooleanVar()
            cb = ttk.Checkbutton(frame, text=cat["display_name"], variable=var,
                                 command=lambda k=cat["key"]: self.load_category_phrases(k))
            cb.pack(side=tk.LEFT, padx=5)
            self.category_checkboxes[cat["key"]] = var
            self.category_widgets[cat["key"]] = {"frame": frame, "var": var}

    def add_category(self):
        name = simpledialog.askstring("Add Category", "Enter category name:")
        if name and name.strip():
            key = name.lower().replace(" ", "_")
            if any(c["key"] == key for c in self.data.categories):
                messagebox.showerror("Error", "Category already exists!")
                return
            self.data.categories.append({"key": key, "display_name": name.strip()})
            self.data.phrases[key] = []
            self.refresh_categories_list()

    def delete_category(self):
        selected_key = None
        for key, var in self.category_checkboxes.items():
            if var.get():
                selected_key = key
                break
        if not selected_key:
            messagebox.showwarning("No Selection", "Select a category to delete.")
            return
        cat_display = next((c["display_name"] for c in self.data.categories if c["key"] == selected_key), selected_key)
        if messagebox.askyesno("Confirm Delete", f"Delete '{cat_display}' and all its phrases?"):
            self.data.categories = [c for c in self.data.categories if c["key"] != selected_key]
            self.data.phrases.pop(selected_key, None)
            if self.current_category == selected_key:
                self.phrases_text.delete(1.0, tk.END)
                self.current_category = None
            self.refresh_categories_list()

    def load_category_phrases(self, category_key):
        for key, var in self.category_checkboxes.items():
            if key != category_key:
                var.set(False)
        self.current_category = category_key
        self.phrases_text.delete(1.0, tk.END)
        phrases = self.data.phrases.get(category_key, [])
        for phrase in phrases:
            if phrase:
                self.phrases_text.insert(tk.END, f"- {phrase}\n")
        if not phrases:
            self.phrases_text.insert(tk.END, "- ")

    def save_current_phrases(self):
        if not self.current_category:
            messagebox.showwarning("No Selection", "Select a category first.")
            return
        content = self.phrases_text.get(1.0, tk.END).strip()
        if not content:
            self.data.phrases[self.current_category] = []
            messagebox.showinfo("Success", "Phrases cleared and saved.")
            return
        lines = content.split('\n')
        phrases = []
        for line in lines:
            line = line.strip()
            if line:
                if line.startswith('- '):
                    line = line[2:]
                phrases.append(line)
        self.data.phrases[self.current_category] = phrases
        messagebox.showinfo("Success", f"Saved {len(phrases)} phrases.")

    def refresh_current_display(self):
        if not self.current_category:
            messagebox.showwarning("No Selection", "Select a category first.")
            return
        self.load_category_phrases(self.current_category)

    def clear_text_area(self):
        if messagebox.askyesno("Clear", "Clear all phrases for this category?"):
            self.phrases_text.delete(1.0, tk.END)
            self.phrases_text.insert(tk.END, "- ")

    # ---------------------------- EXPORT / IMPORT / BACKUP ----------------------------
    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")],
                                            initialfile=f"projects_{datetime.now().strftime('%Y%m%d')}.csv")
        if not path:
            return
        with open(path, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Project", "Code", "Client", "Status", "Priority", "Due Date", "Risk", "Budget",
                             "Tags", "Notes", "Task", "Task Status", "Task Priority", "Task Due", "Hours", "Est Hours"])
            for proj in self.data.projects:
                if proj.get("archived"):
                    continue
                tasks = proj.get("tasks", [])
                for i, task in enumerate(tasks):
                    writer.writerow([
                        proj["name"] if i == 0 else "", proj.get("code", "") if i == 0 else "",
                        proj.get("client", "") if i == 0 else "", proj.get("status", "") if i == 0 else "",
                        proj.get("priority", "") if i == 0 else "", proj.get("due_date", "") if i == 0 else "",
                        proj.get("risk_level", "") if i == 0 else "", proj.get("budget", 0) if i == 0 else "",
                        ", ".join(proj.get("tags", [])) if i == 0 else "", proj.get("remarks", "") if i == 0 else "",
                        task.get("title", ""), task.get("status", ""), task.get("priority", ""),
                        task.get("due_date", ""), task.get("hours", 0), task.get("estimated_hours", 0)
                    ])
                if not tasks:
                    writer.writerow([proj["name"], proj.get("code", ""), proj.get("client", ""),
                        proj.get("status", ""), proj.get("priority", ""), proj.get("due_date", ""),
                        proj.get("risk_level", ""), proj.get("budget", 0),
                        ", ".join(proj.get("tags", [])), proj.get("remarks", ""), "", "", "", "", "", ""])
        messagebox.showinfo("Export", f"Exported to {path}")

    def import_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            with open(path, "r") as f:
                reader = csv.DictReader(f)
                imported = 0
                for row in reader:
                    name = row.get("Project", "").strip()
                    if not name:
                        continue
                    code = row.get("Code", name[:4].upper()).strip()
                    if not any(p["name"] == name for p in self.data.projects):
                        try:
                            budget_val = float(row.get("Budget", 0) or 0)
                        except ValueError:
                            budget_val = 0.0
                        proj = {
                            "name": name, "code": code,
                            "client": row.get("Client", "").strip(),
                            "status": row.get("Status", "not_started").lower().replace(" ", "_"),
                            "priority": row.get("Priority", "medium").lower(),
                            "due_date": row.get("Due Date", "").strip(),
                            "risk_level": row.get("Risk", "low").lower(),
                            "budget": budget_val,
                            "tags": [t.strip() for t in row.get("Tags", "").split(",") if t.strip()],
                            "remarks": row.get("Notes", "").strip(),
                            "color": "#3498db", "expenses": [], "invoices": [],
                            "milestones": [], "archived": False,
                            "created_date": datetime.now().strftime("%Y-%m-%d"), "tasks": []
                        }
                        task_title = row.get("Task", "").strip()
                        if task_title:
                            try:
                                hours_val = float(row.get("Hours", 0) or 0)
                            except ValueError:
                                hours_val = 0.0
                            try:
                                est_val = float(row.get("Est Hours", 0) or 0)
                            except ValueError:
                                est_val = 0.0
                            proj["tasks"].append({
                                "title": task_title,
                                "status": row.get("Task Status", "todo").lower(),
                                "priority": row.get("Task Priority", "medium").lower(),
                                "due_date": row.get("Task Due", "").strip(),
                                "hours": hours_val, "estimated_hours": est_val,
                                "subtasks": [], "comments": [], "blocked_by": "", "attachments": []
                            })
                        self.data.projects.append(proj)
                        imported += 1
            self.data.log_activity("CSV_IMPORTED", f"{imported} projects from {os.path.basename(path)}")
            self.data.save_data()
            self.refresh_all()
            messagebox.showinfo("Import", f"Imported {imported} projects.")
        except Exception as e:
            messagebox.showerror("Import Error", str(e))

    def backup_data(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")],
            initialfile=f"tracker_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.json")
        if not path:
            return
        with open(path, "w") as f:
            json.dump({"phrases": self.data.phrases, "categories": self.data.categories,
                       "projects": self.data.projects, "activity_log": self.data.activity_log}, f, indent=4)
        messagebox.showinfo("Backup", f"Backed up to {path}")

    def restore_data(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "r") as f:
                data = json.load(f)
            self.data.phrases = data.get("phrases", self.data.phrases)
            self.data.categories = data.get("categories", self.data.categories)
            self.data.projects = data.get("projects", self.data.projects)
            self.data.activity_log = data.get("activity_log", [])
            self._selected_project = None
            self.data.save_data()
            self.refresh_all()
            self.clear_project_form()
            messagebox.showinfo("Restore", f"Restored from {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Restore Error", str(e))

    def print_report(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Print Report")
        dlg.geometry("700x600")
        dlg.transient(self.root)

        close_btn = ttk.Button(dlg, text="Close", command=dlg.destroy)
        close_btn.pack(anchor=tk.E, padx=10, pady=5)

        report = tk.Text(dlg, wrap=tk.WORD, font=("Consolas", 9), width=80, height=30)
        report.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        report.insert(tk.END, "=" * 70 + "\n")
        report.insert(tk.END, f"PROJECT TRACKER REPORT - {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
        report.insert(tk.END, "=" * 70 + "\n\n")

        active = [p for p in self.data.projects if not p.get("archived")]
        report.insert(tk.END, f"Total Projects: {len(active)}\n")
        report.insert(tk.END, f"Completed: {sum(1 for p in active if p['status'] == 'completed')}\n")
        report.insert(tk.END, f"In Progress: {sum(1 for p in active if p['status'] == 'in_progress')}\n")
        report.insert(tk.END, f"On Hold: {sum(1 for p in active if p['status'] == 'on_hold')}\n\n")

        report.insert(tk.END, "-" * 70 + "\n")
        report.insert(tk.END, f"{'Project':<20} {'Code':<10} {'Status':<15} {'Priority':<10} {'Due':<12} {'Budget':>10}\n")
        report.insert(tk.END, "-" * 70 + "\n")

        for proj in active:
            report.insert(tk.END, f"{proj['name']:<20} {proj.get('code',''):<10} "
                f"{proj.get('status',''):<15} {proj.get('priority',''):<10} "
                f"{proj.get('due_date',''):<12} ${proj.get('budget',0):>9,.0f}\n")

        report.insert(tk.END, "\n" + "-" * 70 + "\nTASKS SUMMARY\n" + "-" * 70 + "\n")
        for proj in active:
            tasks = proj.get("tasks", [])
            if tasks:
                report.insert(tk.END, f"\n  {proj['name']}:\n")
                for t in tasks:
                    report.insert(tk.END, f"    [{t.get('status','')}] {t.get('title','')} "
                        f"(Pri: {t.get('priority','')}, Due: {t.get('due_date','')}, "
                        f"Hrs: {t.get('hours',0)}/{t.get('estimated_hours',0)})\n")

        report.insert(tk.END, "\n" + "=" * 70 + "\nEND OF REPORT\n")

    # ---------------------------- SAVE ON CLOSE ----------------------------
    def on_close(self):
        if self.current_category and self.phrases_text.get(1.0, tk.END).strip():
            response = messagebox.askyesnocancel("Save Changes", "Save changes to current phrases before exiting?")
            if response is True:
                self.save_current_phrases()
            elif response is None:
                return
        self.data.save_data()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = ProjectTrackerApp(root)
    root.mainloop()
