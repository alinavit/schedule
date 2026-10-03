"""Convert a USOS/iCalendar (.ics) class calendar into a printable PDF schedule.

Double-click this file (or run the accompanying .bat file) to select a calendar.
It can also be used from a terminal:
    python study_schedule_converter.py input.ics output.pdf v1|v2
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

FONT_DIR = Path(r"C:\Windows\Fonts")
REGULAR_FONT = FONT_DIR / "arial.ttf"
BOLD_FONT = FONT_DIR / "arialbd.ttf"


def unfold_ics(text: str) -> list[str]:
    """Join RFC 5545 folded lines and return logical calendar lines."""
    return re.sub(r"\r?\n[ \t]", "", text).replace("\r\n", "\n").split("\n")


def ical_unescape(value: str) -> str:
    return (value.replace(r"\n", "\n").replace(r"\N", "\n")
                 .replace(r"\,", ",").replace(r"\;", ";").replace(r"\\", "\\"))


def parse_datetime(value: str) -> datetime:
    value = value.rstrip("Z")
    for pattern in ("%Y%m%dT%H%M%S", "%Y%m%dT%H%M", "%Y%m%d"):
        try:
            return datetime.strptime(value, pattern)
        except ValueError:
            pass
    raise ValueError(f"Unsupported calendar date: {value}")


def parse_ics(path: Path) -> list[dict]:
    events, event = [], None
    for line in unfold_ics(path.read_text(encoding="utf-8-sig")):
        if line == "BEGIN:VEVENT":
            event = {}
        elif line == "END:VEVENT" and event:
            if "DTSTART" in event and "DTEND" in event:
                start, end = parse_datetime(event["DTSTART"]), parse_datetime(event["DTEND"])
                description = event.get("DESCRIPTION", "")
                room = ""
                building = ""
                for detail in description.splitlines():
                    if detail.lower().startswith("sala:"):
                        room = detail.split(":", 1)[1].strip()
                    elif detail and not detail.startswith("http") and not building:
                        building = detail.strip()
                events.append({
                    "start": start, "end": end,
                    "title": event.get("SUMMARY", "Untitled class"),
                    "room": room, "building": building,
                })
            event = None
        elif event is not None and ":" in line:
            key, value = line.split(":", 1)
            event[key.split(";", 1)[0]] = ical_unescape(value)
    return sorted(events, key=lambda item: item["start"])


def register_fonts() -> tuple[str, str]:
    if REGULAR_FONT.exists() and BOLD_FONT.exists():
        pdfmetrics.registerFont(TTFont("Schedule", str(REGULAR_FONT)))
        pdfmetrics.registerFont(TTFont("Schedule-Bold", str(BOLD_FONT)))
        return "Schedule", "Schedule-Bold"
    return "Helvetica", "Helvetica-Bold"


def make_portrait_pdf(events: list[dict], destination: Path, calendar_name: str) -> None:
    regular, bold = register_fonts()
    destination.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(destination), pagesize=A4, rightMargin=15*mm,
                            leftMargin=15*mm, topMargin=14*mm, bottomMargin=14*mm)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("schedule-title", parent=styles["Title"], fontName=bold,
                           fontSize=22, leading=26, textColor=colors.HexColor("#17324D"), alignment=TA_CENTER)
    subtitle = ParagraphStyle("schedule-subtitle", parent=styles["Normal"], fontName=regular,
                              fontSize=9.5, leading=13, textColor=colors.HexColor("#586674"), alignment=TA_CENTER)
    week_style = ParagraphStyle("week", parent=styles["Heading2"], fontName=bold,
                                fontSize=14, leading=18, textColor=colors.HexColor("#17324D"), spaceBefore=7, spaceAfter=5)
    day_style = ParagraphStyle("day", parent=styles["Heading3"], fontName=bold,
                               fontSize=11.5, leading=15, textColor=colors.white)
    item_style = ParagraphStyle("item", parent=styles["BodyText"], fontName=regular,
                                fontSize=9.2, leading=12, textColor=colors.HexColor("#1F2933"))
    small_style = ParagraphStyle("small", parent=item_style, fontSize=8.3, leading=10.5,
                                 textColor=colors.HexColor("#52616B"))

    story = [Paragraph("Semester 7", title),
             Spacer(1, 2*mm),
             #Paragraph(f"Created from {calendar_name} - {len(events)} scheduled classes", subtitle),
             Spacer(1, 5*mm)]
    grouped = defaultdict(list)
    for event in events:
        grouped[event["start"].date()].append(event)
    dates = sorted(grouped)
    if not dates:
        story.append(Paragraph("No dated classes were found in this calendar.", item_style))

    current_week = None
    for day in dates:
        iso_week = day.isocalendar()[:2]
        if iso_week != current_week:
            if current_week is not None:
                story.append(Spacer(1, 4*mm))
            monday = datetime.fromisocalendar(iso_week[0], iso_week[1], 1).date()
            sunday = datetime.fromisocalendar(iso_week[0], iso_week[1], 7).date()
            story.append(Paragraph(f"Week of {monday:%d %b} - {sunday:%d %b %Y}", week_style))
            current_week = iso_week
        header = Paragraph(day.strftime("%A, %d %B %Y"), day_style)
        header_table = Table([[header]], colWidths=[180*mm])
        header_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#287A8A")),
            ("LEFTPADDING", (0, 0), (-1, -1), 7), ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        rows = []
        for event in grouped[day]:
            time = f"{event['start']:%H:%M} - {event['end']:%H:%M}"
            subject = event["title"]
            location = " - ".join(part for part in (event["room"], event["building"]) if part)
            detail = f"<b>{subject}</b>"
            if location:
                detail += f"<br/><font color='#52616B'>{location}</font>"
            rows.append([Paragraph(time, ParagraphStyle("time", parent=item_style, fontName=bold, textColor=colors.HexColor("#17324D"))),
                         Paragraph(detail, item_style)])
        items = Table(rows, colWidths=[34*mm, 146*mm], repeatRows=0)
        items.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F8FA")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D5E0E6")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 7), ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story += [KeepTogether([header_table, items]), Spacer(1, 4*mm)]

    def footer(canvas, document):
        canvas.saveState()
        canvas.setFont(regular, 8)
        canvas.setFillColor(colors.HexColor("#667785"))
        #canvas.drawString(15*mm, 9*mm, "Study class schedule")
        canvas.drawRightString(195*mm, 9*mm, f"Page {document.page}")
        canvas.restoreState()
    doc.build(story, onFirstPage=footer, onLaterPages=footer)


def make_landscape_pdf(events: list[dict], destination: Path, calendar_name: str) -> None:
    """Create v2: one landscape page per week, with Fri/Sat/Sun in columns."""
    regular, bold = register_fonts()
    destination.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(destination), pagesize=landscape(A4), rightMargin=12*mm,
                            leftMargin=12*mm, topMargin=12*mm, bottomMargin=13*mm)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("v2-title", parent=styles["Title"], fontName=bold,
                           fontSize=20, leading=24, textColor=colors.HexColor("#17324D"), alignment=TA_CENTER)
    subtitle = ParagraphStyle("v2-subtitle", parent=styles["Normal"], fontName=regular,
                              fontSize=9, leading=12, textColor=colors.HexColor("#586674"), alignment=TA_CENTER)
    week_style = ParagraphStyle("v2-week", parent=styles["Heading2"], fontName=bold,
                                fontSize=14, leading=18, textColor=colors.HexColor("#17324D"), alignment=TA_CENTER)
    day_style = ParagraphStyle("v2-day", parent=styles["Normal"], fontName=bold,
                               fontSize=11, leading=14, textColor=colors.white, alignment=TA_CENTER)
    class_style = ParagraphStyle("v2-class", parent=styles["BodyText"], fontName=regular,
                                 fontSize=8.4, leading=10.5, textColor=colors.HexColor("#1F2933"))
    empty_style = ParagraphStyle("v2-empty", parent=class_style, textColor=colors.HexColor("#667785"), alignment=TA_CENTER)

    by_week = defaultdict(lambda: defaultdict(list))
    for event in events:
        monday = event["start"].date().fromordinal(event["start"].date().toordinal() - event["start"].weekday())
        by_week[monday][event["start"].date()].append(event)
    story = [
        #Paragraph("Study class schedule - v2", title), Spacer(1, 1.5*mm),
        #Paragraph(f"Landscape weekly view | Created from {calendar_name} - {len(events)} scheduled classes", subtitle), Spacer(1, 4*mm)
        ]
    for index, monday in enumerate(sorted(by_week)):
        sunday = monday.fromordinal(monday.toordinal() + 6)
        if index:
            story.append(PageBreak())
            story += [
                #Paragraph("Study class schedule - v2", title), Spacer(1, 4*mm)
            ]
        story.append(Paragraph(f"Week of {monday:%d %b} - {sunday:%d %b %Y}", week_style))
        story.append(Spacer(1, 3*mm))
        dates = [monday.fromordinal(monday.toordinal() + offset) for offset in (4, 5, 6)]
        headers = [Paragraph(day.strftime("%A<br/>%d %B"), day_style) for day in dates]
        cells = []
        for day in dates:
            event_parts = []
            for event in by_week[monday].get(day, []):
                detail = f"<b>{event['start']:%H:%M} - {event['end']:%H:%M}</b><br/>{event['title']}"
                location = " - ".join(part for part in (event["room"], event["building"]) if part)
                if location:
                    detail += f"<br/><font color='#52616B'>{location}</font>"
                event_parts.extend([Paragraph(detail, class_style), Spacer(1, 2*mm)])
            cells.append(event_parts or [Spacer(1, 3*mm), Paragraph("No scheduled classes", empty_style)])
        table = Table([headers, cells], colWidths=[91*mm, 91*mm, 91*mm], rowHeights=[15*mm, 110*mm])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#287A8A")),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F5F8FA")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B8CBD4")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 7), ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(table)

    def footer(canvas, document):
        canvas.saveState()
        canvas.setFont(regular, 8)
        canvas.setFillColor(colors.HexColor("#667785"))
        #canvas.drawString(12*mm, 8*mm, "Study class schedule - layout v2")
        canvas.drawRightString(285*mm, 8*mm, f"Page {document.page}")
        canvas.restoreState()
    doc.build(story, onFirstPage=footer, onLaterPages=footer)


def make_pdf(events: list[dict], destination: Path, calendar_name: str, layout: str = "v1") -> None:
    if layout == "v1":
        make_portrait_pdf(events, destination, calendar_name)
    elif layout == "v2":
        make_landscape_pdf(events, destination, calendar_name)
    else:
        raise ValueError("Layout must be 'v1' (portrait list) or 'v2' (landscape weekly grid).")


def choose_files() -> tuple[Path, Path, str] | None:
    root = Tk(); root.withdraw(); root.attributes("-topmost", True)
    source = filedialog.askopenfilename(title="Choose your calendar file", filetypes=[("Calendar files", "*.ics")])
    if not source:
        root.destroy(); return None
    output = filedialog.asksaveasfilename(title="Save study schedule as", defaultextension=".pdf",
                                          initialfile=f"{Path(source).stem}_study_schedule.pdf",
                                          filetypes=[("PDF files", "*.pdf")])
    if not output:
        root.destroy(); return None
    use_v2 = messagebox.askyesno("Choose layout", "Use layout v2?\n\nYes: landscape weekly grid (Friday, Saturday, Sunday side by side)\nNo: layout v1 (portrait schedule list)", parent=root)
    root.destroy()
    return Path(source), Path(output), ("v2" if use_v2 else "v1")


def main() -> None:
    if len(sys.argv) >= 2:
        source = Path(sys.argv[1])
        output = Path(sys.argv[2]) if len(sys.argv) >= 3 else source.with_name(f"{source.stem}_study_schedule.pdf")
        layout = sys.argv[3].lower() if len(sys.argv) >= 4 else "v1"
    else:
        picked = choose_files()
        if not picked:
            return
        source, output, layout = picked
    try:
        make_pdf(parse_ics(source), output, source.name, layout)
        if len(sys.argv) == 1:
            root = Tk(); root.withdraw(); messagebox.showinfo("Schedule created", f"Your PDF is ready:\n{output}"); root.destroy()
        else:
            print(f"Created: {output}")
    except Exception as error:
        if len(sys.argv) == 1:
            root = Tk(); root.withdraw(); messagebox.showerror("Could not create schedule", str(error)); root.destroy()
        else:
            raise


if __name__ == "__main__":
    main()
