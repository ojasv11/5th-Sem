"""
CLARA - Classroom Allocation Resource Assistant
Desktop GUI (Tkinter / ttk), connected to the clara_db MySQL database.

Folder layout this file expects:
    CLARA/GUI/clara_app.py          (this file)
    CLARA/backend/db.py             (database access)
    CLARA/backend/timetable_files.py
    CLARA/converter/timetable_converter.py

Install once:   pip install "PyMySQL[rsa]" openpyxl
Run:            python clara_app.py        (from CLARA/GUI)
"""

import os
import sys
import tkinter as tk
import tkinter.font as tkfont
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


def _project_root():
    """The CLARA folder (the one that contains 'backend')."""
    try:
        here = Path(__file__).resolve().parent
    except NameError:                       # e.g. pasted into a Jupyter cell
        here = Path.cwd()
    for folder in (here, here.parent):
        if (folder / "backend").is_dir():
            return folder
    return here.parent


sys.path.insert(0, str(_project_root()))
from backend import db, timetable_files      # noqa: E402  (needs the path above)

# ---------------------------------------------------------------------------
# Theme tokens: black and white, with greys only for lines and secondary text
# ---------------------------------------------------------------------------
WHITE = "#FFFFFF"
BLACK = "#000000"
INK_SOFT = "#5C5C5C"    # secondary text
INK_MUTED = "#8A8A8A"   # captions and hints
LINE = "#D6D6D6"        # thin borders and dividers
FILL = "#F4F4F4"        # hover state and alternating table rows
THUMB = "#BDBDBD"       # scrollbar thumb and chip borders

PAGE_PAD_X = 28
PAGE_PAD_Y = 20
GAP = 12
SIDEBAR_WIDTH = 210

# ---------------------------------------------------------------------------
# Data loaded from the database (refreshed every time a screen is opened)
# ---------------------------------------------------------------------------
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
SHORT_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
WEEKDAYS = DAYS + ["Sunday"]  # index matches datetime.weekday()

PERIOD_TIMES = ["08:55", "09:55", "10:50", "11:10", "12:05", "13:55", "14:50", "15:10", "16:05", "17:55"]
TIME_CHOICES = sorted(set([f"{h:02d}:{m:02d}" for h in range(8, 18) for m in (0, 30)]
                          + ["18:00"] + PERIOD_TIMES))

ROOMS = []            # (room_name, room_type)
TIMETABLE = []        # (day, start, end, room, subject, (section labels...))
SECTIONS = []         # 'B.Tech CSE 3-A', ...
TODAY_BOOKINGS = []   # (room, start, end, purpose, allocation_id) for today

NAV_ITEMS = [("dashboard", "Dashboard"), ("view", "Timetable"),
             ("allocate", "Find a room"), ("upload", "Upload")]
PAGE_TITLES = {"dashboard": "Dashboard", "view": "Timetable",
               "allocate": "Find a room", "upload": "Upload timetable"}


# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------
def to_minutes(hhmm):
    hours, minutes = hhmm.split(":")
    return int(hours) * 60 + int(minutes)


def natural_key(text):
    """Sort 'CR2' before 'CR10' and '3-B' before '3-C'."""
    parts, number = [], ""
    for ch in str(text):
        if ch.isdigit():
            number += ch
        else:
            if number:
                parts.append((0, int(number)))
                number = ""
            parts.append((1, ch.lower()))
    if number:
        parts.append((0, int(number)))
    return parts


def room_types():
    return dict(ROOMS)


def all_sections():
    return sorted(set(SECTIONS) | {s for row in TIMETABLE for s in row[5]}, key=natural_key)


def parse_section(label):
    """'B.Tech CSE 3-A' -> ('B.Tech CSE', '3', 'A')."""
    try:
        course, rest = label.rsplit(" ", 1)
        year, section = rest.split("-", 1)
        return course, year, section
    except ValueError:
        return label, "", ""


def short_section(label):
    """'B.Tech CSE 3-A' -> '3-A' (used while only one course is in the database)."""
    return label.split()[-1]


def sections_text(sections):
    """Section names for a table cell. With several courses in the database the
    course is kept ('BCA 3-A'), so sections A of different courses are not confused."""
    many_courses = len({parse_section(s)[0] for s in all_sections()}) > 1
    text = ", ".join(s if many_courses else short_section(s) for s in sections)
    return text + "   \u00b7 combined" if len(sections) > 1 else text


def overlaps(start_a, end_a, start_b, end_b):
    """Half-open slots: a class ending at 11:00 does not clash with one starting at 11:00."""
    return to_minutes(start_a) < to_minutes(end_b) and to_minutes(start_b) < to_minutes(end_a)


def sort_rows(rows):
    return sorted(rows, key=lambda r: (DAYS.index(r[0]) if r[0] in DAYS else 9,
                                       to_minutes(r[1]), natural_key(r[3])))


# ---------------------------------------------------------------------------
# Fonts and ttk styles
# ---------------------------------------------------------------------------
def build_fonts(root):
    families = set(tkfont.families(root))
    preferred = ("Segoe UI", "Inter", "Helvetica Neue", "Helvetica", "Arial", "DejaVu Sans")
    family = next((f for f in preferred if f in families), "TkDefaultFont")
    return {
        "logo": (family, 17, "bold"), "h1": (family, 17, "bold"), "title": (family, 13, "bold"),
        "h2": (family, 11, "bold"), "body": (family, 10), "body_bold": (family, 10, "bold"),
        "small": (family, 9), "metric": (family, 24, "bold"),
    }


def configure_styles(root, fonts):
    style = ttk.Style(root)
    style.theme_use("clam")  # clam lets us recolour borders, fields and headings

    style.configure(".", background=WHITE, foreground=BLACK, font=fonts["body"])
    style.configure("TFrame", background=WHITE)
    style.configure("TLabel", background=WHITE, foreground=BLACK)

    style.configure("Primary.TButton", background=BLACK, foreground=WHITE,
                    bordercolor=BLACK, lightcolor=BLACK, darkcolor=BLACK,
                    focuscolor=BLACK, padding=(16, 7), font=fonts["body_bold"])
    style.map("Primary.TButton",
              background=[("pressed", "#333333"), ("active", "#262626")],
              lightcolor=[("pressed", "#333333"), ("active", "#262626")],
              darkcolor=[("pressed", "#333333"), ("active", "#262626")])

    style.configure("Secondary.TButton", background=WHITE, foreground=BLACK,
                    bordercolor=INK_SOFT, lightcolor=WHITE, darkcolor=WHITE,
                    focuscolor=WHITE, padding=(16, 7), font=fonts["body"])
    style.map("Secondary.TButton",
              background=[("pressed", LINE), ("active", FILL)],
              lightcolor=[("pressed", LINE), ("active", FILL)],
              darkcolor=[("pressed", LINE), ("active", FILL)])

    style.configure("TEntry", fieldbackground=WHITE, foreground=BLACK, bordercolor=LINE,
                    lightcolor=WHITE, darkcolor=WHITE, insertcolor=BLACK, padding=5)
    style.map("TEntry", bordercolor=[("focus", BLACK)])

    style.configure("TCombobox", fieldbackground=WHITE, background=WHITE, foreground=BLACK,
                    arrowcolor=BLACK, bordercolor=LINE, lightcolor=WHITE, darkcolor=WHITE, padding=4)
    style.map("TCombobox",
              fieldbackground=[("readonly", WHITE)],
              background=[("active", FILL), ("readonly", WHITE)],
              selectbackground=[("readonly", WHITE)],
              selectforeground=[("readonly", BLACK)],
              bordercolor=[("focus", BLACK)])
    root.option_add("*TCombobox*Listbox.background", WHITE)
    root.option_add("*TCombobox*Listbox.foreground", BLACK)
    root.option_add("*TCombobox*Listbox.selectBackground", BLACK)
    root.option_add("*TCombobox*Listbox.selectForeground", WHITE)
    root.option_add("*TCombobox*Listbox.font", fonts["body"])

    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    style.configure("Treeview", background=WHITE, fieldbackground=WHITE, foreground=BLACK,
                    rowheight=28, font=fonts["body"], borderwidth=0)
    style.map("Treeview", background=[("selected", BLACK)], foreground=[("selected", WHITE)])
    style.configure("Treeview.Heading", background=WHITE, foreground=BLACK,
                    font=fonts["body_bold"], bordercolor=LINE, lightcolor=WHITE,
                    darkcolor=WHITE, relief="flat", padding=(6, 6))
    style.map("Treeview.Heading", background=[("active", FILL)])

    style.configure("Vertical.TScrollbar", background=THUMB, troughcolor=WHITE,
                    bordercolor=WHITE, lightcolor=THUMB, darkcolor=THUMB,
                    arrowcolor=INK_SOFT, gripcount=0)
    style.map("Vertical.TScrollbar", background=[("active", INK_SOFT)])


# ---------------------------------------------------------------------------
# Reusable widgets
# ---------------------------------------------------------------------------
def make_label(parent, text, font, fg=BLACK, bg=WHITE, **kw):
    return tk.Label(parent, text=text, font=font, fg=fg, bg=bg, **kw)


def make_card(parent, padding=14, border=LINE):
    """A white panel with a 1px border. Returns (outer, inner)."""
    outer = tk.Frame(parent, bg=WHITE, highlightthickness=1,
                     highlightbackground=border, highlightcolor=border)
    inner = tk.Frame(outer, bg=WHITE)
    inner.pack(fill="both", expand=True, padx=padding, pady=padding)
    return outer, inner


def make_table(parent, columns, height=10, empty_text="No records yet"):
    """Bordered Treeview with a vertical scrollbar. columns = [(id, title, width), ...]"""
    frame = tk.Frame(parent, bg=WHITE, highlightthickness=1,
                     highlightbackground=LINE, highlightcolor=LINE)
    tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings",
                        height=height, selectmode="browse")
    for col_id, title, width in columns:
        tree.heading(col_id, text=title, anchor="w")
        tree.column(col_id, width=width, minwidth=50, anchor="w", stretch=True)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    tree.pack(side="left", fill="both", expand=True)
    tree.tag_configure("odd", background=FILL)
    tree.empty_label = tk.Label(frame, text=empty_text, bg=WHITE, fg=INK_SOFT,
                                font=ttk.Style().lookup("Treeview", "font"))
    fill_table(tree, [])
    return frame, tree


def fill_table(tree, rows, empty_text=None, iids=None):
    """Replace the rows of a table. iids (optional) gives each row an id, e.g. 'b-12'."""
    tree.delete(*tree.get_children())
    for i, row in enumerate(rows):
        tree.insert("", "end", iid=iids[i] if iids else None, values=row,
                    tags=("odd",) if i % 2 else ())
    if empty_text is not None:
        tree.empty_label.configure(text=empty_text)
    if rows:
        tree.empty_label.place_forget()
    else:
        tree.empty_label.place(relx=0.5, rely=0.55, anchor="center")


def style_chip(chip, active):
    if active:
        chip.configure(bg=BLACK, fg=WHITE, highlightbackground=BLACK, highlightcolor=BLACK)
    else:
        chip.configure(bg=WHITE, fg=BLACK, highlightbackground=THUMB, highlightcolor=THUMB)


def make_chip(parent, text, fonts, active=False, command=None):
    """Small bordered label. Active chips are black with white text."""
    chip = tk.Label(parent, text=text, font=fonts["small"], padx=10, pady=3,
                    highlightthickness=1, cursor="hand2" if command else "")
    style_chip(chip, active)
    if command:
        chip.bind("<Button-1>", lambda _e: command())
    return chip


class ChipGroup(tk.Frame):
    """A row of chips where exactly one is selected (used for the day filter)."""

    def __init__(self, parent, fonts, options, on_change):
        super().__init__(parent, bg=WHITE)
        self.on_change = on_change
        self.chips = []
        self.selected = 0
        for i, text in enumerate(options):
            chip = make_chip(self, text, fonts, command=lambda i=i: self.select(i, notify=True))
            chip.pack(side="left", padx=(0, 6))
            self.chips.append(chip)
        self.select(0)

    def select(self, index, notify=False):
        self.selected = index
        for i, chip in enumerate(self.chips):
            style_chip(chip, i == index)
        if notify:
            self.on_change(index)


class StatusMessage(tk.Frame):
    """One-line status area. Success is outlined in black; errors are inverted."""

    KINDS = {
        "info": (WHITE, INK_SOFT, LINE, "\u2022"),
        "success": (WHITE, BLACK, BLACK, "\u2713"),
        "error": (BLACK, WHITE, BLACK, "\u2715"),
    }

    def __init__(self, parent, fonts, text="", kind="info"):
        super().__init__(parent, highlightthickness=1)
        self.icon = tk.Label(self, font=fonts["body_bold"], width=2)
        self.text = tk.Label(self, font=fonts["body"], anchor="w", justify="left", wraplength=760)
        self.icon.pack(side="left", padx=(10, 2), pady=7)
        self.text.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=7)
        self.set(kind, text)

    def set(self, kind, text):
        bg, fg, border, icon = self.KINDS[kind]
        self.configure(bg=bg, highlightbackground=border, highlightcolor=border)
        self.icon.configure(bg=bg, fg=fg, text=icon)
        self.text.configure(bg=bg, fg=fg, text=text)


# ---------------------------------------------------------------------------
# Login window (the password is only kept in memory while the app runs)
# ---------------------------------------------------------------------------
class LoginDialog(tk.Toplevel):
    def __init__(self, master, fonts, message, current):
        super().__init__(master, bg=WHITE)
        self.title("CLARA - connect to database")
        self.resizable(False, False)
        self.result = None

        body = tk.Frame(self, bg=WHITE)
        body.pack(padx=28, pady=24)
        make_label(body, "Connect to the database", fonts["h1"]).grid(row=0, column=0, columnspan=2, sticky="w")
        make_label(body, "Enter your MySQL login. The password isn't saved.", fonts["small"],
                   fg=INK_SOFT).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 14))

        self.vars = {}
        fields = [("host", "Host", current.get("host", "localhost"), ""),
                  ("port", "Port", str(current.get("port", 3306)), ""),
                  ("database", "Database", current.get("database", "clara_db"), ""),
                  ("user", "User", current.get("user") or "root", ""),
                  ("password", "Password", "", "\u2022")]
        for i, (key, label, value, show) in enumerate(fields, start=2):
            make_label(body, label, fonts["small"], fg=INK_SOFT).grid(row=i, column=0, sticky="w", padx=(0, 12), pady=4)
            var = tk.StringVar(value=value)
            entry = ttk.Entry(body, textvariable=var, width=28, show=show)
            entry.grid(row=i, column=1, sticky="w", pady=4)
            self.vars[key] = var
            if key == "password":
                entry.focus_set()

        row = len(fields) + 2
        if message:
            status = StatusMessage(body, fonts, message, "error")
            status.text.configure(wraplength=300)
            status.grid(row=row, column=0, columnspan=2, sticky="we", pady=(10, 0))
            row += 1
        buttons = tk.Frame(body, bg=WHITE)
        buttons.grid(row=row, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(buttons, text="Exit", style="Secondary.TButton", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Connect", style="Primary.TButton",
                   command=self._connect).pack(side="right", padx=(0, 8))
        self.bind("<Return>", lambda _e: self._connect())

        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_reqwidth()) // 2
        y = (self.winfo_screenheight() - self.winfo_reqheight()) // 3
        self.geometry(f"+{x}+{y}")
        self.lift()
        self.attributes("-topmost", True)          # appear in front of other windows,
        self.after(300, lambda: self.attributes("-topmost", False))   # then behave normally
        self.focus_force()
        self.grab_set()

    def _connect(self):
        self.result = {key: var.get().strip() for key, var in self.vars.items()}
        self.destroy()


# ---------------------------------------------------------------------------
# App chrome: sidebar and header
# ---------------------------------------------------------------------------
class Sidebar(tk.Frame):
    def __init__(self, parent, fonts, on_select):
        super().__init__(parent, bg=WHITE, width=SIDEBAR_WIDTH)
        self.pack_propagate(False)
        self.fonts = fonts
        self.active = None
        self.items = {}

        brand = tk.Frame(self, bg=WHITE)
        brand.pack(fill="x", padx=18, pady=(22, 4))
        mark = tk.Canvas(brand, width=30, height=30, bg=WHITE, highlightthickness=0)
        mark.create_rectangle(0, 0, 30, 30, fill=BLACK, outline=BLACK)
        mark.create_text(15, 15, text="C", fill=WHITE, font=fonts["h2"])
        mark.pack(side="left")
        make_label(brand, "CLARA", fonts["logo"]).pack(side="left", padx=(10, 0))
        make_label(self, "Classroom allocation", fonts["small"], fg=INK_MUTED).pack(
            anchor="w", padx=18, pady=(0, 18))

        for key, text in NAV_ITEMS:
            item = tk.Label(self, text=text, font=fonts["body"], bg=WHITE, fg=INK_SOFT,
                            anchor="w", padx=14, pady=9, cursor="hand2")
            item.pack(fill="x", padx=10, pady=1)
            item.bind("<Button-1>", lambda _e, k=key: on_select(k))
            item.bind("<Enter>", lambda _e, k=key: self._hover(k, True))
            item.bind("<Leave>", lambda _e, k=key: self._hover(k, False))
            self.items[key] = item

        footer = tk.Frame(self, bg=WHITE)
        footer.pack(side="bottom", fill="x", padx=18, pady=16)
        tk.Frame(footer, bg=LINE, height=1).pack(fill="x", pady=(0, 10))
        make_label(footer, f"{db.BUILDING} building", fonts["small"], fg=INK_SOFT).pack(anchor="w")
        self.rooms_label = make_label(footer, "", fonts["small"], fg=INK_SOFT)
        self.rooms_label.pack(anchor="w")
        self.conn_label = make_label(footer, "", fonts["small"], fg=INK_MUTED)
        self.conn_label.pack(anchor="w", pady=(6, 0))

    def _hover(self, key, entering):
        if key != self.active:
            self.items[key].configure(bg=FILL if entering else WHITE)

    def set_active(self, key):
        self.active = key
        for k, item in self.items.items():
            if k == key:
                item.configure(bg=BLACK, fg=WHITE, font=self.fonts["body_bold"])
            else:
                item.configure(bg=WHITE, fg=INK_SOFT, font=self.fonts["body"])
        self.rooms_label.configure(text=f"{len(ROOMS)} rooms" if ROOMS else "No rooms loaded")

    def set_connection(self, ok):
        name = db.settings()["database"]
        self.conn_label.configure(text=f"\u25cf Connected to {name}" if ok else "\u25cb Not connected",
                                  fg=INK_SOFT if ok else BLACK)


class Header(tk.Frame):
    def __init__(self, parent, fonts):
        super().__init__(parent, bg=WHITE, height=64)
        self.pack_propagate(False)

        left = tk.Frame(self, bg=WHITE)
        left.pack(side="left", padx=PAGE_PAD_X)
        self.title = make_label(left, "", fonts["title"])
        self.title.pack(anchor="w")
        make_label(left, datetime.now().strftime("%A, %d %B %Y"), fonts["small"],
                   fg=INK_MUTED).pack(anchor="w")

        profile = tk.Frame(self, bg=WHITE)
        profile.pack(side="right", padx=PAGE_PAD_X)
        avatar = tk.Canvas(profile, width=34, height=34, bg=WHITE, highlightthickness=0)
        avatar.create_oval(1, 1, 33, 33, fill=BLACK, outline=BLACK)
        avatar.create_text(17, 17, text="A", fill=WHITE, font=fonts["body_bold"])
        avatar.pack(side="right")
        names = tk.Frame(profile, bg=WHITE)
        names.pack(side="right", padx=(0, 10))
        make_label(names, "Admin", fonts["body_bold"]).pack(anchor="e")
        make_label(names, "Administrator", fonts["small"], fg=INK_MUTED).pack(anchor="e")

    def set_title(self, text):
        self.title.configure(text=text)


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------
class Page(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=WHITE)
        self.app = app
        self.fonts = app.fonts
        self.body = tk.Frame(self, bg=WHITE)
        self.body.pack(fill="both", expand=True, padx=PAGE_PAD_X, pady=PAGE_PAD_Y)
        self.build()

    def build(self):
        raise NotImplementedError

    def on_show(self):
        """Called every time the page becomes visible (after the data is reloaded)."""


class DashboardPage(Page):
    def build(self):
        make_label(self.body, "Welcome to CLARA Admin", self.fonts["h1"]).pack(anchor="w")
        make_label(self.body, f"Room usage across the {db.BUILDING} building at a glance.",
                   self.fonts["body"], fg=INK_SOFT).pack(anchor="w", pady=(2, 16))

        cards = tk.Frame(self.body, bg=WHITE)
        cards.pack(fill="x")
        self.metrics = {}
        specs = [("rooms", "Rooms"), ("classes", "Weekly classes"),
                 ("sections", "Sections"), ("free", "Free now")]
        for i, (key, label) in enumerate(specs):
            cards.grid_columnconfigure(i, weight=1, uniform="metric")
            outer, inner = make_card(cards, border=BLACK if key == "free" else LINE)
            outer.grid(row=0, column=i, sticky="nsew",
                       padx=(0 if i == 0 else GAP // 2, 0 if i == len(specs) - 1 else GAP // 2))
            make_label(inner, label, self.fonts["small"], fg=INK_SOFT).pack(anchor="w")
            value = make_label(inner, "0", self.fonts["metric"])
            value.pack(anchor="w", pady=(2, 0))
            caption = make_label(inner, "", self.fonts["small"], fg=INK_SOFT)
            caption.pack(anchor="w")
            self.metrics[key] = (value, caption)

        lower = tk.Frame(self.body, bg=WHITE)
        lower.pack(fill="both", expand=True, pady=(GAP, 0))
        lower.grid_columnconfigure(0, weight=1, uniform="lower")
        lower.grid_columnconfigure(1, weight=1, uniform="lower")
        lower.grid_rowconfigure(0, weight=1)

        busy_outer, busy = make_card(lower)
        busy_outer.grid(row=0, column=0, sticky="nsew", padx=(0, GAP // 2))
        make_label(busy, "Busy right now", self.fonts["h2"]).pack(anchor="w", pady=(0, 8))
        frame, self.busy_tree = make_table(
            busy, [("room", "Room", 80), ("what", "Subject or booking", 200), ("until", "Until", 90)],
            height=7, empty_text="No rooms are in use right now")
        frame.pack(fill="both", expand=True)

        # Every room in the building, grouped into classrooms and labs
        rooms_outer, rooms = make_card(lower)
        rooms_outer.grid(row=0, column=1, sticky="nsew", padx=(GAP // 2, 0))
        rooms_head = tk.Frame(rooms, bg=WHITE)
        rooms_head.pack(fill="x", pady=(0, 8))
        make_label(rooms_head, "Rooms in the building", self.fonts["h2"]).pack(side="left")
        ttk.Button(rooms_head, text="Edit rooms", style="Secondary.TButton",
                   command=lambda: RoomEditor(self.app)).pack(side="right")
        self.room_groups = {}
        for kind, title in [("Classroom", "Classrooms"), ("Lab", "Labs")]:
            make_label(rooms, title, self.fonts["small"], fg=INK_SOFT).pack(anchor="w", pady=(4, 4))
            grid = tk.Frame(rooms, bg=WHITE)
            grid.pack(fill="x", pady=(0, 8))
            self.room_groups[kind] = grid

    def on_show(self):
        now = datetime.now()
        day, now_hm = WEEKDAYS[now.weekday()], now.strftime("%H:%M")
        minute = to_minutes(now_hm)
        running = [(r[3], r[4], r[2]) for r in TIMETABLE
                   if r[0] == day and to_minutes(r[1]) <= minute < to_minutes(r[2])]
        running += [(b[0], b[3], b[2]) for b in TODAY_BOOKINGS
                    if to_minutes(b[1]) <= minute < to_minutes(b[2])]
        busy_rooms = {r[0] for r in running}
        labs = sum(1 for t in room_types().values() if t == "Lab")
        combined = sum(1 for r in TIMETABLE if len(r[5]) > 1)
        sections = all_sections()
        courses = sorted({s.rsplit("-", 1)[0] for s in sections})

        self._set("rooms", len(ROOMS),
                  f"{len(ROOMS) - labs} classrooms, {labs} labs" if ROOMS else "No rooms loaded")
        self._set("classes", len(TIMETABLE),
                  f"{combined} combined" if TIMETABLE else "No timetable imported yet")
        self._set("sections", len(sections), ", ".join(courses) if sections else "No sections yet")
        self._set("free", len(ROOMS) - len(busy_rooms) if ROOMS else 0,
                  f"at {now_hm}" if ROOMS else "Nothing to check yet")
        fill_table(self.busy_tree,
                   [(room, what, f"till {until}") for room, what, until
                    in sorted(running, key=lambda r: natural_key(r[0]))])
        self._draw_rooms()

    ROOMS_PER_ROW = 6

    def _draw_rooms(self):
        """Redraw one box per room in the building, grouped by room type."""
        for kind, grid in self.room_groups.items():
            for child in grid.winfo_children():
                child.destroy()
            names = [name for name, room_kind in ROOMS if room_kind == kind]
            for col in range(self.ROOMS_PER_ROW):
                grid.grid_columnconfigure(col, weight=1, uniform="room")
            for i, name in enumerate(names):
                box = tk.Label(grid, text=name, font=self.fonts["small"], pady=4, highlightthickness=1)
                style_chip(box, False)
                box.grid(row=i // self.ROOMS_PER_ROW, column=i % self.ROOMS_PER_ROW,
                         sticky="ew", padx=(0, 6), pady=(0, 6))

    def _set(self, key, value, caption):
        value_label, caption_label = self.metrics[key]
        value_label.configure(text=str(value))
        caption_label.configure(text=caption)


class TimetablePage(Page):
    COLUMNS = [("time", "Time", 130), ("room", "Room", 110),
               ("subject", "Subject", 300), ("sections", "Sections", 240)]

    def build(self):
        self._first_show = True

        toolbar = tk.Frame(self.body, bg=WHITE)
        toolbar.pack(fill="x", pady=(0, 10))
        self.course_var = tk.StringVar()
        self.year_var = tk.StringVar()
        self.section_var = tk.StringVar()
        boxes = [("Course", self.course_var, 20), ("Year", self.year_var, 10),
                 ("Section", self.section_var, 12)]
        self.course_box, self.year_box, self.section_box = [
            self._selector(toolbar, label, var, width, first=(i == 0))
            for i, (label, var, width) in enumerate(boxes)]
        self.course_box.bind("<<ComboboxSelected>>", lambda _e: self._on_course())
        self.year_box.bind("<<ComboboxSelected>>", lambda _e: self._on_year())
        self.section_box.bind("<<ComboboxSelected>>", lambda _e: self.populate())
        ttk.Button(toolbar, text="Refresh", style="Secondary.TButton",
                   command=self.refresh).pack(side="right")

        today = datetime.now().weekday()
        self.days = ChipGroup(self.body, self.fonts, SHORT_DAYS,
                              on_change=lambda _i: self.populate())
        self.days.select(today if today < 6 else 0)       # today, or Monday on Sunday
        self.days.pack(anchor="w", pady=(0, 10))

        frame, self.tree = make_table(self.body, self.COLUMNS, height=14)
        frame.pack(fill="both", expand=True)
        self.count = make_label(self.body, "", self.fonts["small"], fg=INK_SOFT)
        self.count.pack(anchor="w", pady=(8, 0))

    def _selector(self, parent, label, var, width, first=False):
        make_label(parent, label, self.fonts["body"], fg=INK_SOFT).pack(
            side="left", padx=(0 if first else 18, 8))
        box = ttk.Combobox(parent, textvariable=var, state="readonly", width=width)
        box.pack(side="left")
        return box

    # ---- the three lists depend on each other: course -> year -> section
    def _parsed(self):
        return [parse_section(s) for s in all_sections()]

    def _set_options(self, box, var, options):
        """Fill a list; keep the current choice if it is still valid, otherwise pick the first."""
        values = sorted(options, key=natural_key)
        box.configure(values=values)
        if var.get() not in values:
            var.set(values[0] if values else "")

    def _refresh_courses(self):
        self._set_options(self.course_box, self.course_var, {c for c, _, _ in self._parsed()})

    def _refresh_years(self):
        course = self.course_var.get()
        self._set_options(self.year_box, self.year_var,
                          {y for c, y, _ in self._parsed() if c == course})

    def _refresh_sections(self):
        course, year = self.course_var.get(), self.year_var.get()
        self._set_options(self.section_box, self.section_var,
                          {s for c, y, s in self._parsed() if c == course and y == year})

    def _on_course(self):
        self._refresh_years()
        self._refresh_sections()
        self.populate()

    def _on_year(self):
        self._refresh_sections()
        self.populate()

    def on_show(self):
        parsed = sorted(self._parsed(), key=lambda p: natural_key(" ".join(p)))
        if self._first_show and parsed:          # open on the first section
            course, year, section = parsed[0]
            self.course_var.set(course)
            self.year_var.set(year)
            self.section_var.set(section)
            self._first_show = False
        self._refresh_courses()
        self._refresh_years()
        self._refresh_sections()
        self.populate()

    def _selected_label(self):
        """The chosen section as stored in the timetable, e.g. 'B.Tech CSE 3-A'."""
        course, year, section = self.course_var.get(), self.year_var.get(), self.section_var.get()
        return f"{course} {year}-{section}" if course and year and section else ""

    def populate(self):
        day = DAYS[self.days.selected]
        label = self._selected_label()
        rows = sort_rows([r for r in TIMETABLE if label and label in r[5] and r[0] == day])
        fill_table(self.tree,
                   [(f"{r[1]}\u2013{r[2]}", r[3], r[4], sections_text(r[5])) for r in rows],
                   "No timetable imported yet. Use Upload to import one." if not TIMETABLE
                   else f"No classes on {day} for this section.")
        if label:
            noun = "class" if len(rows) == 1 else "classes"
            self.count.configure(text=f"Showing {len(rows)} {day} {noun} for {self.course_var.get()}, "
                                      f"Year {self.year_var.get()}, Section {self.section_var.get()}")
        else:
            self.count.configure(text="")

    def refresh(self):
        if self.app.reload_data():
            self.on_show()


class FindRoomPage(Page):
    CHIPS_PER_ROW = 7

    def build(self):
        self.last_search = None          # (date, start, end, room type) of the results shown
        self.selected_room = None
        self.room_chips = {}

        outer, form = make_card(self.body)
        outer.pack(fill="x")
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d-%m-%Y"))
        self.start_var = tk.StringVar(value="10:00")
        self.end_var = tk.StringVar(value="11:00")
        self.type_var = tk.StringVar(value="Any")
        fields = [
            ("Date (DD-MM-YYYY)", ttk.Entry(form, textvariable=self.date_var, width=14)),
            ("Start", ttk.Combobox(form, textvariable=self.start_var, values=TIME_CHOICES,
                                   state="readonly", width=8)),
            ("End", ttk.Combobox(form, textvariable=self.end_var, values=TIME_CHOICES,
                                 state="readonly", width=8)),
            ("Room type", ttk.Combobox(form, textvariable=self.type_var,
                                       values=["Any", "Classroom", "Lab"], state="readonly", width=11)),
        ]
        for col, (label, widget) in enumerate(fields):
            make_label(form, label, self.fonts["small"], fg=INK_SOFT).grid(
                row=0, column=col, sticky="w", padx=(0, GAP), pady=(0, 4))
            widget.grid(row=1, column=col, sticky="w", padx=(0, GAP))
        ttk.Button(form, text="Find", style="Primary.TButton",
                   command=self.find).grid(row=1, column=len(fields), sticky="w")
        form.grid_columnconfigure(len(fields) + 1, weight=1)

        self.status = StatusMessage(self.body, self.fonts)
        self.status.pack(fill="x", pady=GAP)

        results = tk.Frame(self.body, bg=WHITE)
        results.pack(fill="both", expand=True)
        results.grid_columnconfigure(0, weight=1, uniform="res")
        results.grid_columnconfigure(1, weight=1, uniform="res")
        results.grid_rowconfigure(0, weight=1)

        free_outer, free = make_card(results)
        free_outer.grid(row=0, column=0, sticky="nsew", padx=(0, GAP // 2))
        self.free_title = make_label(free, "Free rooms", self.fonts["h2"])
        self.free_title.pack(anchor="w", pady=(0, 8))
        make_label(free, "Classrooms", self.fonts["small"], fg=INK_SOFT).pack(anchor="w")
        self.classroom_chips = tk.Frame(free, bg=WHITE)
        self.classroom_chips.pack(anchor="w", fill="x", pady=(4, 10))
        make_label(free, "Labs", self.fonts["small"], fg=INK_SOFT).pack(anchor="w")
        self.lab_chips = tk.Frame(free, bg=WHITE)
        self.lab_chips.pack(anchor="w", fill="x", pady=(4, 0))
        self.hint = make_label(free, "", self.fonts["small"], fg=INK_MUTED)
        self.hint.pack(anchor="w", pady=(8, 0))

        # Booking panel: hidden until a free room is clicked
        self.book_panel = tk.Frame(free, bg=WHITE, highlightthickness=1,
                                   highlightbackground=BLACK, highlightcolor=BLACK)
        panel = tk.Frame(self.book_panel, bg=WHITE)
        panel.pack(fill="x", padx=12, pady=10)
        self.book_title = make_label(panel, "", self.fonts["h2"])
        self.book_title.pack(anchor="w")
        self.book_slot = make_label(panel, "", self.fonts["small"], fg=INK_SOFT)
        self.book_slot.pack(anchor="w")
        row = tk.Frame(panel, bg=WHITE)
        row.pack(anchor="w", pady=(10, 0))
        ttk.Button(row, text="Book", style="Primary.TButton", command=self.book).pack(side="left")
        ttk.Button(row, text="Cancel", style="Secondary.TButton",
                   command=self._clear_selection).pack(side="left", padx=(8, 0))

        busy_outer, busy = make_card(results)
        busy_outer.grid(row=0, column=1, sticky="nsew", padx=(GAP // 2, 0))
        self.busy_title = make_label(busy, "Busy rooms", self.fonts["h2"])
        self.busy_title.pack(anchor="w", pady=(0, 8))
        frame, self.busy_tree = make_table(
            busy, [("room", "Room", 70), ("what", "Subject or booking", 150),
                   ("sections", "Sections", 120), ("time", "Time", 100)],
            height=9, empty_text="Run a search to see busy rooms")
        frame.pack(fill="both", expand=True)
        self.busy_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_busy_select())
        self.busy_hint = make_label(busy, "", self.fonts["small"], fg=INK_MUTED)
        self.busy_hint.pack(anchor="w", pady=(8, 0))

        # Cancel panel: hidden until a booking row is clicked
        self.busy_rows = {}              # row id -> (room, what, start, end, allocation_id or None)
        self.cancel_target = None
        self.cancel_panel = tk.Frame(busy, bg=WHITE, highlightthickness=1,
                                     highlightbackground=BLACK, highlightcolor=BLACK)
        cpanel = tk.Frame(self.cancel_panel, bg=WHITE)
        cpanel.pack(fill="x", padx=12, pady=10)
        self.cancel_title = make_label(cpanel, "", self.fonts["h2"])
        self.cancel_title.pack(anchor="w")
        self.cancel_slot = make_label(cpanel, "", self.fonts["small"], fg=INK_SOFT)
        self.cancel_slot.pack(anchor="w")
        crow = tk.Frame(cpanel, bg=WHITE)
        crow.pack(anchor="w", pady=(10, 0))
        ttk.Button(crow, text="Cancel booking", style="Primary.TButton",
                   command=self.cancel_booking).pack(side="left")
        ttk.Button(crow, text="Keep", style="Secondary.TButton",
                   command=self._hide_cancel).pack(side="left", padx=(8, 0))

        self._show_chips(self.classroom_chips, [], "Run a search to see free rooms")
        self._show_chips(self.lab_chips, [], "")
        self.status.set("info", "Choose a date and time slot, then click Find.")

    # ------------------------------------------------------------ chips
    def _show_chips(self, holder, names, empty_text):
        for child in holder.winfo_children():
            child.destroy()
        if not names:
            if empty_text:
                make_label(holder, empty_text, self.fonts["small"], fg=INK_MUTED).grid(row=0, column=0, sticky="w")
            return
        for i, name in enumerate(names):
            chip = make_chip(holder, name, self.fonts, command=lambda n=name: self._select_room(n))
            chip.grid(row=i // self.CHIPS_PER_ROW, column=i % self.CHIPS_PER_ROW,
                      sticky="w", padx=(0, 6), pady=(0, 6))
            self.room_chips[name] = chip

    def _select_room(self, name):
        self.selected_room = name
        for room, chip in self.room_chips.items():
            style_chip(chip, room == name)
        when, start, end, _kind = self.last_search
        self.book_title.configure(text=f"Book {name}")
        self.book_slot.configure(text=f"{WEEKDAYS[when.weekday()]} {when:%d-%m-%Y}, {start}\u2013{end}")
        self.book_panel.pack(fill="x", pady=(12, 0))

    def _clear_selection(self):
        self.selected_room = None
        for chip in self.room_chips.values():
            style_chip(chip, False)
        self.book_panel.pack_forget()

    # ----------------------------------------------------------- search
    def find(self):
        try:
            when = datetime.strptime(self.date_var.get().strip(), "%d-%m-%Y")
        except ValueError:
            self.status.set("error", "Enter the date as DD-MM-YYYY, for example 19-10-2026.")
            return
        start, end = self.start_var.get(), self.end_var.get()
        if to_minutes(end) <= to_minutes(start):
            self.status.set("error", f"End time ({end}) must be later than start time ({start}).")
            return
        if not ROOMS:
            self.status.set("error", "No rooms are in the database yet, so availability can't be checked.")
            return
        self._search(when, start, end, self.type_var.get())

    def _search(self, when, start, end, wanted):
        """Show free and busy rooms for one slot. Returns False if the bookings couldn't be read."""
        try:
            bookings = db.fetch_bookings(when.date().isoformat())
        except db.ClaraDBError as exc:
            self.status.set("error", f"Couldn't read bookings: {exc}")
            return False
        self.last_search = (when, start, end, wanted)
        self.room_chips = {}
        self._clear_selection()

        day = WEEKDAYS[when.weekday()]
        types = room_types()
        in_scope = [name for name, kind in ROOMS if wanted == "Any" or kind == wanted]
        busy = [(r[3], r[4], sections_text(r[5]), r[1], r[2], None) for r in TIMETABLE
                if r[0] == day and r[3] in in_scope and overlaps(r[1], r[2], start, end)]
        busy += [(b[0], b[3], "Booking", b[1], b[2], b[4]) for b in bookings
                 if b[0] in in_scope and overlaps(b[1], b[2], start, end)]
        busy_rooms = {b[0] for b in busy}
        free_rooms = sorted((n for n in in_scope if n not in busy_rooms), key=natural_key)

        self._show_chips(self.classroom_chips, [n for n in free_rooms if types[n] == "Classroom"],
                         "Not included in this search" if wanted == "Lab" else "None free")
        self._show_chips(self.lab_chips, [n for n in free_rooms if types[n] == "Lab"],
                         "Not included in this search" if wanted == "Classroom" else "None free")
        self.hint.configure(text="Click a free room to book it." if free_rooms else "")
        busy.sort(key=lambda b: (natural_key(b[0]), b[3]))
        iids = [f"b-{b[5]}" if b[5] else f"t-{i}" for i, b in enumerate(busy)]
        self.busy_rows = {iid: (b[0], b[1], b[3], b[4], b[5]) for iid, b in zip(iids, busy)}
        self._hide_cancel()
        fill_table(self.busy_tree,
                   [(room, what, who, f"{s}\u2013{e}") for room, what, who, s, e, _id in busy],
                   "No rooms are busy in this slot", iids=iids)
        has_bookings = any(b[5] for b in busy)
        self.busy_hint.configure(text="Click a booking to cancel it." if has_bookings else "")
        self.free_title.configure(text=f"Free rooms ({len(free_rooms)})")
        self.busy_title.configure(text=f"Busy rooms ({len(busy_rooms)})")

        slot = f"{day} {when:%d-%m-%Y}, {start}\u2013{end}"
        if free_rooms:
            self.status.set("success", f"{len(free_rooms)} of {len(in_scope)} rooms are free on {slot}.")
        else:
            self.status.set("error", f"No rooms are free on {slot}. Try another time slot.")
        return True

    # ----------------------------------------------------- cancelling
    def _on_busy_select(self):
        chosen = self.busy_tree.selection()
        if not chosen or chosen[0] not in self.busy_rows:
            return
        room, what, start, end, allocation_id = self.busy_rows[chosen[0]]
        if allocation_id is None:
            self._hide_cancel(keep_selection=True)
            self.busy_hint.configure(text=f"{room} has a regular class ({what}). Timetable classes "
                                          "are changed by re-importing the timetable.")
            return
        when = self.last_search[0]
        self.busy_hint.configure(text="Click a booking to cancel it.")
        self.cancel_target = (allocation_id, room, start, end)
        self.cancel_title.configure(text=f"Cancel booking: {room}")
        self.cancel_slot.configure(text=f"{WEEKDAYS[when.weekday()]} {when:%d-%m-%Y}, {start}\u2013{end}")
        self.cancel_panel.pack(fill="x", pady=(10, 0))

    def _hide_cancel(self, keep_selection=False):
        self.cancel_target = None
        self.cancel_panel.pack_forget()
        if not keep_selection:
            self.busy_tree.selection_remove(self.busy_tree.selection())

    def cancel_booking(self):
        if not self.cancel_target or not self.last_search:
            return
        allocation_id, room, start, end = self.cancel_target
        when, s, e, wanted = self.last_search
        slot = f"{WEEKDAYS[when.weekday()]} {when:%d-%m-%Y}, {start}\u2013{end}"
        if not messagebox.askyesno("Cancel booking", f"Cancel the booking of {room} on {slot}?",
                                   parent=self):
            return
        try:
            db.cancel_booking(allocation_id, room)
        except db.ClaraDBError as exc:
            self.status.set("error", f"Couldn't cancel: {exc}")
            return
        self.app.reload_data()
        self._search(when, s, e, wanted)           # redraw: the room is free again
        self.status.set("success", f"The booking of {room} on {slot} was cancelled.")

    # ---------------------------------------------------------- booking
    def book(self):
        if not self.selected_room or not self.last_search:
            self.status.set("error", "Click a free room first.")
            return
        when, start, end, wanted = self.last_search
        room = self.selected_room
        now = datetime.now()
        if when.date() < now.date():
            self.status.set("error", "That date has already passed. Choose today or a later date.")
            return
        if when.date() == now.date() and to_minutes(end) <= now.hour * 60 + now.minute:
            self.status.set("error", f"{start}\u2013{end} today has already ended. Choose a later time.")
            return
        slot = f"{WEEKDAYS[when.weekday()]} {when:%d-%m-%Y}, {start}\u2013{end}"
        try:
            db.book_room(room, when.date().isoformat(), start, end)
        except db.ClaraDBError as exc:
            self.status.set("error", f"{room} couldn't be booked: {exc}")
            return
        self.app.reload_data()                     # the Dashboard sees the new booking too
        self._search(when, start, end, wanted)     # redraw: the room moves to Busy rooms
        self.status.set("success", f"{room} is booked for {slot}.")


class UploadPage(Page):
    def build(self):
        self.file_path = None
        self.rows = []
        self.check_ok = False

        # File area with a dashed border (drawn on a canvas)
        self.drop = tk.Canvas(self.body, height=104, bg=WHITE, highlightthickness=0)
        self.drop.pack(fill="x")
        inner = tk.Frame(self.drop, bg=WHITE)
        self.drop_window = self.drop.create_window(0, 0, window=inner, anchor="center")
        self.drop.bind("<Configure>", self._redraw_drop)
        self.file_name = make_label(inner, "No file chosen", self.fonts["h2"])
        self.file_name.pack()
        self.file_meta = make_label(inner, "", self.fonts["small"], fg=INK_SOFT)
        self.file_meta.pack(pady=(2, 8))
        ttk.Button(inner, text="Choose file", style="Secondary.TButton",
                   command=self.choose_file).pack()

        # Details the converter needs for an Excel file
        details_outer, details = make_card(self.body, padding=10)
        details_outer.pack(fill="x", pady=(GAP, 0))
        self.course_var = tk.StringVar(value="B.Tech CSE")
        self.year_var = tk.StringVar(value="3")
        self.sections_var = tk.StringVar(value="A,B,C,D,E,F,G,H")
        for col, (label, widget) in enumerate([
                ("Course", ttk.Entry(details, textvariable=self.course_var, width=16)),
                ("Year", ttk.Combobox(details, textvariable=self.year_var, values=["1", "2", "3", "4"],
                                      state="readonly", width=4)),
                ("Sections, in file order", ttk.Entry(details, textvariable=self.sections_var, width=26))]):
            make_label(details, label, self.fonts["small"], fg=INK_SOFT).grid(row=0, column=col, sticky="w",
                                                                                padx=(0, GAP))
            widget.grid(row=1, column=col, sticky="w", padx=(0, GAP))
        make_label(details, "Used for Excel files only. A converted CSV already contains them.",
                   self.fonts["small"], fg=INK_MUTED).grid(row=1, column=3, sticky="w")

        # Step cards
        steps = tk.Frame(self.body, bg=WHITE)
        steps.pack(fill="x", pady=GAP)
        self.steps = {}
        for i, (key, title) in enumerate([("convert", "1  Convert"), ("check", "2  Check"),
                                          ("skipped", "3  Skipped"), ("import", "4  Import")]):
            steps.grid_columnconfigure(i, weight=1, uniform="steps")
            outer, card = make_card(steps, padding=9, border=BLACK if key == "import" else LINE)
            outer.grid(row=0, column=i, sticky="nsew",
                       padx=(0 if i == 0 else GAP // 2, 0 if i == 3 else GAP // 2))
            make_label(card, title, self.fonts["body_bold"]).pack(anchor="w")
            value = make_label(card, "\u2014", self.fonts["small"], fg=INK_SOFT)
            value.pack(anchor="w")
            self.steps[key] = value

        actions = tk.Frame(self.body, bg=WHITE)
        actions.pack(fill="x")
        ttk.Button(actions, text="Import timetable", style="Primary.TButton",
                   command=self.import_timetable).pack(side="left")
        ttk.Button(actions, text="Check file", style="Secondary.TButton",
                   command=self.check_file).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Clear", style="Secondary.TButton",
                   command=self.clear).pack(side="left", padx=(8, 0))
        make_label(actions, "Replaces the current timetable of these sections in one step.",
                   self.fonts["small"], fg=INK_SOFT).pack(side="left", padx=(14, 0))

        self.status = StatusMessage(self.body, self.fonts)
        self.status.pack(fill="x", pady=(GAP, 8))

        frame, self.problems_tree = make_table(
            self.body, [("cell", "Cell", 110), ("problem", "Problem", 640)],
            height=4, empty_text="No problems")
        frame.pack(fill="both", expand=True)
        self.clear()

    def _redraw_drop(self, event):
        self.drop.delete("border")
        self.drop.create_rectangle(1, 1, event.width - 2, event.height - 2,
                                   outline=INK_SOFT, dash=(4, 3), tags="border")
        self.drop.coords(self.drop_window, event.width / 2, event.height / 2)

    def _set_steps(self, convert="\u2014", check="\u2014", skipped="\u2014", imported="\u2014"):
        for key, text in [("convert", convert), ("check", check),
                          ("skipped", skipped), ("import", imported)]:
            self.steps[key].configure(text=text)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Choose timetable file",
            filetypes=[("Timetable files", "*.xlsx *.csv"), ("Excel files", "*.xlsx"),
                       ("CSV files", "*.csv")])
        if not path:
            return
        self.file_path = path
        self.file_name.configure(text=os.path.basename(path))
        self.file_meta.configure(text=f"{os.path.getsize(path) / 1024:.1f} KB in {os.path.dirname(path)}")
        self.check_file()

    def check_file(self):
        """Convert (Excel) or read (CSV), then list any problems. Nothing is saved yet."""
        self.check_ok = False
        self.rows = []
        if not self.file_path:
            self.status.set("error", "Choose a file first.")
            return
        name = os.path.basename(self.file_path)
        try:
            if name.lower().endswith(".xlsx"):
                sections = [s.strip() for s in self.sections_var.get().split(",") if s.strip()]
                if not self.course_var.get().strip() or not sections:
                    self.status.set("error", "Enter the course and the section names before checking an Excel file.")
                    return
                self.status.set("info", f"Converting {name}\u2026")
                self.update_idletasks()
                rows, problems, skipped = timetable_files.read_excel(
                    self.file_path, self.course_var.get().strip(), int(self.year_var.get()), sections)
            elif name.lower().endswith(".csv"):
                rows, problems, skipped = timetable_files.read_csv(self.file_path)
            else:
                self.status.set("error", f"{name} is not an .xlsx or .csv file.")
                return
        except Exception as exc:          # unreadable or unexpected file layout
            self.status.set("error", f"Couldn't read {name}: {exc}")
            fill_table(self.problems_tree, [])
            return

        self.rows = rows
        fill_table(self.problems_tree, problems, "No problems")
        sections_found = len({(r["course"], r["year"], r["section"]) for r in rows})
        self._set_steps(convert=f"{len(rows)} rows, {sections_found} sections",
                        check=f"{len(problems)} problem(s)" if problems else "No problems",
                        skipped=f"{len(skipped)} outside {db.BUILDING}" if skipped else "None",
                        imported="Ready" if rows and not problems else "\u2014")
        if problems:
            self.status.set("error", f"{len(problems)} problem(s) found. Fix them in the Excel file "
                                     "(see the Cell column), then click Check file again.")
        elif not rows:
            self.status.set("error", f"No timetable rows were found in {name}.")
        else:
            self.check_ok = True
            self.status.set("success", f"{len(rows)} rows for {sections_found} sections are ready. "
                                       "Click Import timetable.")

    def import_timetable(self):
        if not self.check_ok:
            self.check_file()
            if not self.check_ok:
                return
        if not messagebox.askyesno(
                "Import timetable",
                "This replaces the current timetable of the sections in this file. Continue?",
                parent=self):
            return
        self.status.set("info", "Importing\u2026")
        self.update_idletasks()
        try:
            result = db.import_timetable(self.rows)
        except db.ClaraDBError as exc:
            self._set_steps(*[self.steps[k].cget("text") for k in ("convert", "check", "skipped")],
                            imported="Failed")
            self.status.set("error", f"Import stopped, nothing was changed: {exc}")
            return
        if result["problems"]:
            fill_table(self.problems_tree, result["problems"])
            self._set_steps(*[self.steps[k].cget("text") for k in ("convert", "check", "skipped")],
                            imported="Stopped")
            self.status.set("error", f"The database checks found {len(result['problems'])} problem(s), "
                                     "so nothing was imported. See the list below.")
            return
        self.check_ok = False
        self._set_steps(*[self.steps[k].cget("text") for k in ("convert", "check", "skipped")],
                        imported="Done")
        self.status.set("success", f"Timetable imported: {result['csv_rows']} rows became "
                                   f"{result['entries']} timetable entries ({result['links']} section links).")
        self.app.reload_data()

    def clear(self):
        self.file_path = None
        self.rows = []
        self.check_ok = False
        self.file_name.configure(text="No file chosen")
        self.file_meta.configure(text="Department timetable (.xlsx) or a converted timetable_rows.csv")
        self._set_steps()
        fill_table(self.problems_tree, [], "No problems")
        self.status.set("info", "Choose the timetable file to begin.")


# ---------------------------------------------------------------------------
# Room editor (opened from the Dashboard's "Edit rooms" button)
# ---------------------------------------------------------------------------
class RoomEditor(tk.Toplevel):
    """Add, rename or delete rooms. Every change goes straight to the database."""

    def __init__(self, app):
        super().__init__(app, bg=WHITE)
        self.app = app
        self.fonts = app.fonts
        self.selected_id = None
        self._quiet_select = False        # True: the next selection event keeps the status message
        self.title("CLARA - edit rooms")
        self.geometry("860x540")
        self.minsize(760, 480)
        self.transient(app)

        body = tk.Frame(self, bg=WHITE)
        body.pack(fill="both", expand=True, padx=24, pady=20)
        make_label(body, "Edit rooms", self.fonts["h1"]).pack(anchor="w")
        make_label(body, f"Rooms of the {db.BUILDING} building. Select a room to rename or delete it.",
                   self.fonts["small"], fg=INK_SOFT).pack(anchor="w", pady=(2, 12))

        content = tk.Frame(body, bg=WHITE)
        content.pack(fill="both", expand=True)
        content.grid_columnconfigure(0, weight=3)
        content.grid_columnconfigure(1, weight=2)
        content.grid_rowconfigure(0, weight=1)

        frame, self.tree = make_table(
            content, [("room", "Room", 110), ("type", "Type", 90),
                      ("classes", "Weekly classes", 110), ("bookings", "Bookings", 80)],
            height=13, empty_text="No rooms yet")
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, GAP))
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._on_select())

        form_outer, form = make_card(content)
        form_outer.grid(row=0, column=1, sticky="nsew")
        self.form_title = make_label(form, "New room", self.fonts["h2"])
        self.form_title.pack(anchor="w", pady=(0, 10))
        make_label(form, "Room name", self.fonts["small"], fg=INK_SOFT).pack(anchor="w")
        self.name_var = tk.StringVar()
        name_entry = ttk.Entry(form, textvariable=self.name_var, width=24)
        name_entry.pack(anchor="w", fill="x", pady=(2, 10))
        make_label(form, "Type", self.fonts["small"], fg=INK_SOFT).pack(anchor="w")
        self.type_var = tk.StringVar(value="Classroom")
        ttk.Combobox(form, textvariable=self.type_var, values=list(db.ROOM_TYPES),
                     state="readonly", width=14).pack(anchor="w", pady=(2, 16))

        self.add_btn = ttk.Button(form, text="Add as new room", style="Primary.TButton", command=self._add)
        self.add_btn.pack(fill="x", pady=(0, 6))
        self.save_btn = ttk.Button(form, text="Save changes", style="Secondary.TButton", command=self._save)
        self.save_btn.pack(fill="x", pady=(0, 6))
        self.delete_btn = ttk.Button(form, text="Delete room", style="Secondary.TButton", command=self._delete)
        self.delete_btn.pack(fill="x", pady=(0, 6))
        ttk.Button(form, text="Clear selection", style="Secondary.TButton",
                   command=self._clear).pack(fill="x")
        make_label(form, "Renaming keeps all classes and bookings, because they are linked to the "
                         "room, not to its name.", self.fonts["small"], fg=INK_MUTED,
                   wraplength=220, justify="left").pack(anchor="w", pady=(14, 0))

        self.status = StatusMessage(body, self.fonts)
        self.status.text.configure(wraplength=780)
        self.status.pack(fill="x", pady=(12, 0))

        self._load()
        self._clear()
        name_entry.focus_set()
        self.grab_set()                       # keep focus here until the editor is closed

    def _load(self, select_id=None):
        try:
            rows = db.fetch_room_usage()
        except db.ClaraDBError as exc:
            self.status.set("error", str(exc))
            return
        self.tree.delete(*self.tree.get_children())
        for i, (room_id, name, kind, classes, bookings) in enumerate(rows):
            self.tree.insert("", "end", iid=str(room_id), values=(name, kind, classes, bookings),
                             tags=("odd",) if i % 2 else ())
        self.tree.empty_label.place_forget() if rows else self.tree.empty_label.place(
            relx=0.5, rely=0.55, anchor="center")
        if select_id is not None and self.tree.exists(str(select_id)):
            self._quiet_select = True
            self.tree.selection_set(str(select_id))
            self.tree.see(str(select_id))

    def _on_select(self):
        chosen = self.tree.selection()
        if not chosen:
            return
        self.selected_id = int(chosen[0])
        name, kind, classes, bookings = self.tree.item(chosen[0], "values")
        self.name_var.set(name)
        self.type_var.set(kind)
        self.form_title.configure(text=f"Editing {name}")
        if self._quiet_select:            # selected by the program after a change: keep the success message
            self._quiet_select = False
            return
        self.status.set("info", f"{name}: {classes} weekly classes, {bookings} upcoming bookings. "
                                "Change the name or type and click Save changes, or Delete room.")

    def _clear(self):
        self.selected_id = None
        self.tree.selection_remove(self.tree.selection())
        self.name_var.set("")
        self.type_var.set("Classroom")
        self.form_title.configure(text="New room")
        self.status.set("info", "Type a name and choose a type to add a room, "
                                "or select a room in the list to edit it.")

    def _changed(self, message, select_id=None):
        """After any change: refresh this list and the whole app."""
        self._load(select_id)
        self.app.reload_data()
        self.app.pages["dashboard"].on_show()
        self.status.set("success", message)

    def _add(self):
        try:
            new_id = db.add_room(self.name_var.get(), self.type_var.get())
        except db.ClaraDBError as exc:
            self.status.set("error", str(exc))
            return
        name = self.name_var.get().strip()
        self._changed(f"{name} was added as a {self.type_var.get().lower()}.", new_id)

    def _save(self):
        if self.selected_id is None:
            self.status.set("error", "Select a room in the list first.")
            return
        old_name = self.tree.item(str(self.selected_id), "values")[0]
        try:
            db.update_room(self.selected_id, self.name_var.get(), self.type_var.get())
        except db.ClaraDBError as exc:
            self.status.set("error", str(exc))
            return
        new_name = self.name_var.get().strip()
        note = (f"{old_name} is now {new_name}." if new_name != old_name
                else f"{new_name} was updated.")
        if new_name != old_name:
            note += (" If future timetable files still say the old name, add a line to "
                     "converter\\room_aliases.csv.")
        self._changed(note, self.selected_id)

    def _delete(self):
        if self.selected_id is None:
            self.status.set("error", "Select a room in the list first.")
            return
        name = self.tree.item(str(self.selected_id), "values")[0]
        if not messagebox.askyesno("Delete room", f"Delete {name} permanently?", parent=self):
            return
        try:
            db.delete_room(self.selected_id)
        except db.ClaraDBError as exc:
            self.status.set("error", str(exc))
            return
        self._clear()
        self._changed(f"{name} was deleted.")


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------
class ClaraApp(tk.Tk):
    PAGES = {"dashboard": DashboardPage, "view": TimetablePage,
             "allocate": FindRoomPage, "upload": UploadPage}

    def __init__(self):
        super().__init__()
        self.withdraw()                       # stay hidden until the database login works
        self.title("CLARA - Classroom Allocation Resource Assistant")
        self.configure(bg=WHITE)
        self.fonts = build_fonts(self)
        configure_styles(self, self.fonts)
        self._last_error = None

        self.connected = self._login()
        if not self.connected:
            self.destroy()
            return

        self.geometry("1240x800")
        self.minsize(1080, 700)
        self.sidebar = Sidebar(self, self.fonts, self.show_page)
        self.sidebar.pack(side="left", fill="y")
        tk.Frame(self, bg=LINE, width=1).pack(side="left", fill="y")

        main = tk.Frame(self, bg=WHITE)
        main.pack(side="left", fill="both", expand=True)
        self.header = Header(main, self.fonts)
        self.header.pack(fill="x")
        tk.Frame(main, bg=LINE, height=1).pack(fill="x")

        container = tk.Frame(main, bg=WHITE)
        container.pack(fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.pages = {}
        for key, page_cls in self.PAGES.items():
            page = page_cls(container, self)
            page.grid(row=0, column=0, sticky="nsew")
            self.pages[key] = page

        for i, (key, _title) in enumerate(NAV_ITEMS, start=1):
            self.bind_all(f"<Control-Key-{i}>", lambda _e, k=key: self.show_page(k))

        self.deiconify()
        self.show_page("dashboard")

    def _login(self):
        """Use the environment variables if set, otherwise ask. Returns False if the user exits."""
        message = ""
        if db.has_credentials():
            try:
                db.test_connection()
                return True
            except db.ClaraDBError as exc:
                message = str(exc)
        while True:
            dialog = LoginDialog(self, self.fonts, message, db.settings())
            self.wait_window(dialog)
            if dialog.result is None:
                return False
            db.configure(**dialog.result)
            try:
                db.test_connection()
                return True
            except db.ClaraDBError as exc:
                message = str(exc)

    def reload_data(self):
        """Read rooms, timetable, sections and today's bookings from the database."""
        try:
            rooms = db.fetch_rooms()
            timetable = db.fetch_timetable()
            sections = db.fetch_sections()
            bookings = db.fetch_bookings(date.today().isoformat())
        except db.ClaraDBError as exc:
            self.sidebar.set_connection(False)
            if str(exc) != self._last_error:          # don't repeat the same pop-up on every click
                messagebox.showerror("Database error", str(exc), parent=self)
            self._last_error = str(exc)
            return False
        ROOMS[:], TIMETABLE[:], SECTIONS[:], TODAY_BOOKINGS[:] = rooms, timetable, sections, bookings
        self._last_error = None
        self.sidebar.set_connection(True)
        return True

    def show_page(self, key):
        self.reload_data()
        page = self.pages[key]
        page.tkraise()
        self.header.set_title(PAGE_TITLES[key])
        self.sidebar.set_active(key)
        page.on_show()


def main():
    try:  # sharper text on high-DPI Windows displays
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = ClaraApp()
    if app.connected:
        app.mainloop()


if __name__ == "__main__":
    main()