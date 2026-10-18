"""
backend/db.py - all MySQL access for the CLARA GUI.

Install once:   pip install "PyMySQL[rsa]"

Login settings, read in this order (later ones win):
  1. environment variables CLARA_DB_HOST / PORT / USER / PASSWORD / NAME
  2. backend/db_config.ini   (keep this file private: don't share or submit it)
When a user and password are found, the GUI connects without a login window.
"""

import configparser
import os
import threading
from datetime import timedelta
from pathlib import Path

try:
    import pymysql
except ImportError:          # reported to the user by connect()
    pymysql = None

BUILDING = "CSIT"
STAGING_COLUMNS = ["day", "start_time", "end_time", "room", "subject",
                   "course", "year", "section", "group_name", "source_cell"]

_settings = {
    "host": os.getenv("CLARA_DB_HOST", "localhost"),
    "port": int(os.getenv("CLARA_DB_PORT", "3306")),
    "user": os.getenv("CLARA_DB_USER", ""),
    "password": os.getenv("CLARA_DB_PASSWORD", ""),
    "database": os.getenv("CLARA_DB_NAME", "clara_db"),
}


CONFIG_FILE = Path(__file__).with_name("db_config.ini")


def _load_config_file():
    """Read backend/db_config.ini if it exists (section [mysql])."""
    if not CONFIG_FILE.exists():
        return
    parser = configparser.ConfigParser(interpolation=None)   # allow '%' in passwords
    parser.read(CONFIG_FILE, encoding="utf-8")
    if not parser.has_section("mysql"):
        return
    for key in ("host", "port", "user", "password", "database"):
        value = parser.get("mysql", key, fallback="").strip()
        if value:
            _settings[key] = int(value) if key == "port" else value


_load_config_file()


class ClaraDBError(Exception):
    """A readable message the GUI can show as it is."""


# ---------------------------------------------------------------- connection
def configure(**values):
    """Update connection settings, e.g. configure(user='root', password='...')."""
    for key, value in values.items():
        if key in _settings and value not in (None, ""):
            _settings[key] = int(value) if key == "port" else value


def settings():
    """Current settings without the password (to pre-fill the login window)."""
    return {k: v for k, v in _settings.items() if k != "password"}


def has_credentials():
    return bool(_settings["user"] and _settings["password"])


def _message(exc):
    code = exc.args[0] if exc.args else None
    text = exc.args[1] if len(exc.args) > 1 else str(exc)
    friendly = {
        1062: "That name is already used by another room in this building.",
        1451: "This room is still used by the timetable or by bookings, so it can't be deleted.",
        1045: "Wrong MySQL username or password.",
        1049: f"The database '{_settings['database']}' doesn't exist. Run 01_clara_schema.sql first.",
        2003: "Can't reach MySQL. Check that the MySQL service is running.",
    }
    return friendly.get(code, text)     # trigger and procedure messages pass through as written


def connect():
    if pymysql is None:
        raise ClaraDBError('PyMySQL isn\'t installed. In Command Prompt run: pip install "PyMySQL[rsa]"')
    try:
        return pymysql.connect(host=_settings["host"], port=_settings["port"],
                               user=_settings["user"], password=_settings["password"],
                               database=_settings["database"], charset="utf8mb4",
                               autocommit=True,          # every read sees the latest data
                               connect_timeout=5)
    except pymysql.MySQLError as exc:
        raise ClaraDBError(_message(exc)) from exc
    except RuntimeError as exc:                           # missing 'cryptography' for MySQL 8 login
        raise ClaraDBError(f'{exc}. Run: pip install "PyMySQL[rsa]"') from exc


def _query(sql, params=None):
    conn = connect()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params or None)     # None: no '%' substitution at all
            return cur.fetchall()
    except pymysql.MySQLError as exc:
        raise ClaraDBError(_message(exc)) from exc
    finally:
        conn.close()


def _execute(sql, params=None):
    """Run one INSERT / UPDATE / DELETE. Returns (rows affected, new id)."""
    conn = connect()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params or None)
            return cur.rowcount, cur.lastrowid
    except pymysql.MySQLError as exc:
        raise ClaraDBError(_message(exc)) from exc
    finally:
        conn.close()


def hhmm(value):
    """PyMySQL returns TIME columns as timedelta; show them as HH:MM."""
    if isinstance(value, timedelta):
        minutes = int(value.total_seconds()) // 60
        return f"{minutes // 60:02d}:{minutes % 60:02d}"
    return str(value)[:5]


def test_connection():
    _query("SELECT 1")


# ------------------------------------------------------------------- reading
def fetch_rooms():
    """[(room_name, room_type), ...] for the building, in room order."""
    rows = _query("""SELECT r.room_name, r.room_type
                     FROM Classrooms r
                     JOIN Buildings b ON b.building_id = r.building_id
                     WHERE b.building_name = %s
                     ORDER BY r.room_id""", (BUILDING,))
    return [tuple(r) for r in rows]


def fetch_sections():
    """['B.Tech CSE 3-A', ...] for every section in the database."""
    rows = _query("""SELECT CONCAT(c.course_name, ' ', c.course_year, '-', s.section_name)
                     FROM Sections s
                     JOIN Courses c ON c.course_id = s.course_id
                     ORDER BY c.course_name, c.course_year, s.section_name""")
    return [r[0] for r in rows]


def fetch_timetable():
    """[(day, start, end, room, subject, (section labels...)), ...]
    A combined class is one row with several section labels."""
    rows = _query("""SELECT t.day_of_week, t.start_time, t.end_time, r.room_name, sub.subject_code,
                            GROUP_CONCAT(DISTINCT CONCAT(c.course_name, ' ', c.course_year, '-', sec.section_name)
                                         ORDER BY c.course_name, c.course_year, sec.section_name
                                         SEPARATOR '|')
                     FROM Timetable t
                     JOIN Classrooms r   ON r.room_id     = t.room_id
                     JOIN Buildings  b   ON b.building_id = r.building_id
                     JOIN Subjects   sub ON sub.subject_id = t.subject_id
                     LEFT JOIN TimetableSections ts  ON ts.timetable_id = t.timetable_id
                     LEFT JOIN Sections          sec ON sec.section_id  = ts.section_id
                     LEFT JOIN Courses           c   ON c.course_id     = sec.course_id
                     WHERE b.building_name = %s
                     GROUP BY t.timetable_id, t.day_of_week, t.start_time, t.end_time,
                              r.room_name, sub.subject_code""", (BUILDING,))
    return [(day, hhmm(start), hhmm(end), room, subject, tuple(labels.split("|")) if labels else ())
            for day, start, end, room, subject, labels in rows]


def fetch_bookings(date_iso):
    """Confirmed temporary bookings on a date ('YYYY-MM-DD'):
    [(room, start, end, purpose, allocation_id), ...]"""
    rows = _query("""SELECT r.room_name, a.start_time, a.end_time,
                            COALESCE(ar.purpose, 'Temporary booking'), a.allocation_id
                     FROM Allocations a
                     JOIN Classrooms r ON r.room_id = a.room_id
                     LEFT JOIN AllocationRequests ar ON ar.request_id = a.request_id
                     WHERE a.allocation_status = 'Confirmed' AND a.allocation_date = %s""",
                  (date_iso,))
    return [(room, hhmm(start), hhmm(end), purpose, allocation_id)
            for room, start, end, purpose, allocation_id in rows]


# --------------------------------------------------------------- editing rooms
ROOM_TYPES = ("Classroom", "Lab")


def _clean_room(name, room_type):
    name = (name or "").strip()
    if not name:
        raise ClaraDBError("Enter a room name.")
    if len(name) > 50:
        raise ClaraDBError("Room names can be at most 50 characters.")
    if room_type not in ROOM_TYPES:
        raise ClaraDBError("Choose Classroom or Lab as the room type.")
    return name


def fetch_room_usage():
    """[(room_id, room_name, room_type, weekly_classes, upcoming_bookings), ...]"""
    return [tuple(r) for r in _query(
        """SELECT r.room_id, r.room_name, r.room_type,
                  (SELECT COUNT(*) FROM Timetable t WHERE t.room_id = r.room_id),
                  (SELECT COUNT(*) FROM Allocations a
                   WHERE a.room_id = r.room_id AND a.allocation_status = 'Confirmed'
                     AND a.allocation_date >= CURDATE())
           FROM Classrooms r
           JOIN Buildings b ON b.building_id = r.building_id
           WHERE b.building_name = %s
           ORDER BY r.room_id""", (BUILDING,))]


def add_room(name, room_type):
    """Add a room to the building. Returns its new room_id."""
    name = _clean_room(name, room_type)
    count, new_id = _execute(
        """INSERT INTO Classrooms (room_name, room_type, building_id)
           SELECT %s, %s, building_id FROM Buildings WHERE building_name = %s""",
        (name, room_type, BUILDING))
    if count != 1:
        raise ClaraDBError(f"The {BUILDING} building isn't in the database.")
    return new_id


def update_room(room_id, name, room_type):
    """Rename a room and/or change its type. Classes and bookings stay linked,
    because they point to room_id, not to the name."""
    name = _clean_room(name, room_type)
    _execute("UPDATE Classrooms SET room_name = %s, room_type = %s WHERE room_id = %s",
             (name, room_type, room_id))


def delete_room(room_id):
    """Delete a room that nothing refers to. The foreign keys (ON DELETE RESTRICT)
    would refuse anyway; this check just gives a clearer message first."""
    (name, classes, bookings, requests), = _query(
        """SELECT r.room_name,
                  (SELECT COUNT(*) FROM Timetable          WHERE room_id = r.room_id),
                  (SELECT COUNT(*) FROM Allocations        WHERE room_id = r.room_id),
                  (SELECT COUNT(*) FROM AllocationRequests WHERE requested_room_id = r.room_id)
           FROM Classrooms r WHERE r.room_id = %s""", (room_id,)) or [(None, 0, 0, 0)]
    if name is None:
        raise ClaraDBError("That room no longer exists.")
    reasons = [f"{n} {what}" for n, what in [(classes, "timetable classes"),
                                              (bookings, "bookings"),
                                              (requests, "booking requests")] if n]
    if reasons:
        raise ClaraDBError(f"{name} can't be deleted: it has {' and '.join(reasons)}. "
                           "Re-import the timetable without it, or remove those first.")
    _execute("DELETE FROM Classrooms WHERE room_id = %s", (room_id,))


# ------------------------------------------------------------- booking rooms
# One lock per room, inside this program (OS-level mutual exclusion): two bookings
# of the same room made by CLARA can never run their check-and-insert at the same time.
_room_locks = {}
_room_locks_guard = threading.Lock()      # protects the dictionary itself


def _room_lock(room_name):
    with _room_locks_guard:
        return _room_locks.setdefault(room_name, threading.Lock())


def book_room(room_name, date_iso, start, end, purpose=""):
    """Temporarily book a room. Returns the new allocation_id.
    Protection, from outside in:
      1. this room's threading.Lock         (mutual exclusion in CLARA)
      2. SELECT ... FOR UPDATE on the room  (mutual exclusion in MySQL, across programs)
      3. the Allocations trigger            (refuses any overlap with bookings or classes)
    Everything runs in one transaction, so a failure saves nothing."""
    purpose = (purpose or "").strip()[:255] or None
    with _room_lock(room_name):
        conn = connect()
        try:
            with conn.cursor() as cur:
                cur.execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")
                conn.begin()
                cur.execute("""SELECT r.room_id FROM Classrooms r
                               JOIN Buildings b ON b.building_id = r.building_id
                               WHERE b.building_name = %s AND r.room_name = %s
                               FOR UPDATE""", (BUILDING, room_name))
                row = cur.fetchone()
                if not row:
                    raise ClaraDBError(f"{room_name} no longer exists.")
                room_id = row[0]
                cur.execute("""SELECT user_id FROM Users WHERE role = 'Admin' AND is_active
                               ORDER BY user_id LIMIT 1""")
                admin = cur.fetchone()
                if not admin:
                    raise ClaraDBError("There is no admin account in the database.")
                # The request row keeps the purpose; it is approved at once.
                cur.execute("""INSERT INTO AllocationRequests
                                   (requested_room_id, request_date, start_time, end_time,
                                    priority, request_status, purpose, requested_by)
                               VALUES (%s, %s, %s, %s, 'Medium', 'Approved', %s, %s)""",
                            (room_id, date_iso, start, end, purpose, admin[0]))
                request_id = cur.lastrowid
                cur.execute("""INSERT INTO Allocations
                                   (room_id, allocation_date, start_time, end_time,
                                    allocation_status, request_id, allocated_by)
                               VALUES (%s, %s, %s, %s, 'Confirmed', %s, %s)""",
                            (room_id, date_iso, start, end, request_id, admin[0]))
                allocation_id = cur.lastrowid
            conn.commit()
            return allocation_id
        except pymysql.MySQLError as exc:
            conn.rollback()
            raise ClaraDBError(_message(exc)) from exc
        except ClaraDBError:
            conn.rollback()
            raise
        finally:
            conn.close()


def cancel_booking(allocation_id, room_name):
    """Cancel a confirmed booking. The row is kept with status 'Cancelled' (history),
    and no longer blocks the room. Uses the same room lock as book_room()."""
    with _room_lock(room_name):
        conn = connect()
        try:
            with conn.cursor() as cur:
                conn.begin()
                cur.execute("""SELECT request_id FROM Allocations
                               WHERE allocation_id = %s AND allocation_status = 'Confirmed'
                               FOR UPDATE""", (allocation_id,))
                row = cur.fetchone()
                if not row:
                    raise ClaraDBError("This booking was already cancelled or no longer exists.")
                cur.execute("UPDATE Allocations SET allocation_status = 'Cancelled' "
                            "WHERE allocation_id = %s", (allocation_id,))
                if row[0] is not None:
                    cur.execute("UPDATE AllocationRequests SET request_status = 'Cancelled' "
                                "WHERE request_id = %s", (row[0],))
            conn.commit()
        except pymysql.MySQLError as exc:
            conn.rollback()
            raise ClaraDBError(_message(exc)) from exc
        except ClaraDBError:
            conn.rollback()
            raise
        finally:
            conn.close()


# ------------------------------------------------------------------ importing
# The same checks as STEP B of 03_import_timetable.sql. Each returns (source_cell, problem).
_CHECKS = [
    ("""SELECT st.source_cell, CONCAT('Room ', COALESCE(st.room, '?'), ' is not a room in ', %s)
        FROM TimetableStaging st
        WHERE NOT EXISTS (SELECT 1 FROM Classrooms r JOIN Buildings b ON b.building_id = r.building_id
                          WHERE b.building_name = %s AND r.room_name = st.room)""", True),
    ("""SELECT st.source_cell, 'Invalid or missing day, time, year, course, section, subject or group'
        FROM TimetableStaging st
        WHERE COALESCE(st.day, '') NOT IN ('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')
           OR COALESCE(st.start_time, '') NOT REGEXP '^[0-2][0-9]:[0-5][0-9]$'
           OR COALESCE(st.end_time, '')   NOT REGEXP '^[0-2][0-9]:[0-5][0-9]$'
           OR st.end_time <= st.start_time
           OR COALESCE(st.year, '') NOT REGEXP '^[1-4]$'
           OR COALESCE(st.course, '') = '' OR COALESCE(st.section, '') = ''
           OR COALESCE(st.subject, '') = ''
           OR COALESCE(st.group_name, '') NOT IN ('All', 'G1', 'G2')""", False),
    ("""SELECT MIN(source_cell),
               CONCAT(room, ' ', day, ' ', start_time, ': different subjects (',
                      GROUP_CONCAT(DISTINCT subject SEPARATOR ' / '), ')')
        FROM TimetableStaging
        GROUP BY room, day, start_time, end_time
        HAVING COUNT(DISTINCT subject) > 1""", False),
    ("""SELECT a.source_cell,
               CONCAT('Room clash: ', a.room, ' ', a.day, ' ', a.start_time, '-', a.end_time,
                      ' overlaps ', b.start_time, '-', b.end_time, ' (', b.source_cell, ')')
        FROM TimetableStaging a
        JOIN TimetableStaging b
          ON b.room = a.room AND b.day = a.day AND a.staging_id < b.staging_id
         AND a.start_time < b.end_time AND b.start_time < a.end_time
         AND NOT (a.start_time = b.start_time AND a.end_time = b.end_time)""", False),
    ("""SELECT a.source_cell,
               CONCAT('Section clash: ', a.course, ' ', a.year, '-', a.section, ' is in ', a.room,
                      ' and ', b.room, ' on ', a.day, ' ', a.start_time, ' (', b.source_cell, ')')
        FROM TimetableStaging a
        JOIN TimetableStaging b
          ON b.course = a.course AND b.year = a.year AND b.section = a.section AND b.day = a.day
         AND a.staging_id < b.staging_id
         AND a.start_time < b.end_time AND b.start_time < a.end_time
         AND (a.group_name = 'All' OR b.group_name = 'All' OR a.group_name = b.group_name)""", False),
]


def import_timetable(rows):
    """Load converter rows (dicts with STAGING_COLUMNS) into the database.
    Returns {"problems": [(cell, text), ...]} if the checks fail, otherwise
    {"problems": [], "csv_rows": n, "entries": n, "links": n}.
    The real import is sp_import_timetable: one transaction, all or nothing."""
    conn = connect()
    try:
        with conn.cursor() as cur:
            cur.execute("TRUNCATE TABLE TimetableStaging")
            placeholders = ", ".join(["%s"] * len(STAGING_COLUMNS))
            cur.executemany(f"INSERT INTO TimetableStaging ({', '.join(STAGING_COLUMNS)}) "
                            f"VALUES ({placeholders})",
                            [tuple(str(r.get(c, "")) for c in STAGING_COLUMNS) for r in rows])

            problems = []
            for sql, needs_building in _CHECKS:
                cur.execute(sql, (BUILDING, BUILDING) if needs_building else None)
                problems += [(cell or "", text) for cell, text in cur.fetchall()]
            if problems:
                return {"problems": problems}

            cur.execute("CALL sp_import_timetable(%s)", (BUILDING,))
            csv_rows, entries, links = cur.fetchone()
            while cur.nextset():          # a CALL returns an extra empty result
                pass
            return {"problems": [], "csv_rows": csv_rows, "entries": entries, "links": links}
    except pymysql.MySQLError as exc:
        raise ClaraDBError(_message(exc)) from exc
    finally:
        conn.close()