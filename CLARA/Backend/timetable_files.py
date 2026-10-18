
import csv
import sys
from pathlib import Path

CONVERTER_DIR = Path(__file__).resolve().parent.parent / "converter"
COLUMNS = ["day", "start_time", "end_time", "room", "subject",
           "course", "year", "section", "group_name", "source_cell"]


def read_excel(path, course, year, sections):
    """Convert the department's Excel timetable. `sections` is the list of
    section names, one per timetable block in the file, in order."""
    if str(CONVERTER_DIR) not in sys.path:
        sys.path.insert(0, str(CONVERTER_DIR))
    try:
        import timetable_converter as converter
    except ImportError as exc:
        missing = "openpyxl" if "openpyxl" in str(exc) else "timetable_converter.py"
        if missing == "openpyxl":
            return [], [("", "openpyxl isn't installed. In Command Prompt run: pip install openpyxl")], []
        return [], [("", f"Can't find {CONVERTER_DIR / 'timetable_converter.py'}")], []

    rows, raw_problems, skipped, blocks = converter.convert_files(
        [path], course=course, year=year, sections=sections)
    problems = [(p.get("Source Cell", ""), f"{p.get('Issue', '')}  [{p.get('Text', '')}]".strip())
                for p in raw_problems]
    if sections and len(sections) != len(blocks):
        problems.insert(0, ("", f"The file has {len(blocks)} timetable blocks but "
                                f"{len(sections)} section names were given."))
    return rows, problems, skipped


def read_csv(path):
    """Read a timetable_rows.csv made by the converter."""
    with open(path, newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        header = reader.fieldnames or []
        rows = list(reader)
    missing = [c for c in COLUMNS if c not in header]
    if missing:
        return [], [("", f"Missing column(s): {', '.join(missing)}. "
                         "Use the CSV made by timetable_converter.py.")], []
    return rows, [], []
