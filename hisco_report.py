"""
HISCO Confirmation team - daily report updater.

The report workbook contains charts, comments, threaded comments and pivot
tables, so it is NOT loaded with openpyxl (that would drop them). Instead the
.xlsx (a zip of XML files) is patched directly:

  * rows are appended to the four data sheets  CONF, Tracking, Certificate,
    Moved to FAB
  * cell B6 of the month sheet (e.g. Nov'26) is set to the loaded date
  * the workbook is flagged to recalculate when opened in Excel

Everything else in the file is copied byte for byte.
"""
import datetime as dt
import io
import re
import unicodedata
import zipfile
from xml.sax.saxutils import escape

import openpyxl

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------

# EOD "Task" value (lower case)  ->  report sheet
TASK_TO_SHEET = {
    "confirmation": "CONF",
    "tracking": "Tracking",
    "certificates": "Certificate",
    "certificate": "Certificate",
    "moved to fab conf": "Moved to FAB",
    "moved to fab confirmation": "Moved to FAB",
    "moved to fab confirmations": "Moved to FAB",
}

# report sheet -> (report column letter, EOD column index) pairs.
# EOD columns: 0 Date,1 Name,2 Task,3 PO,4 Item,5 Stage,6 Buyer,7 Promise,
# 8 Rev Promise,9 Disposition,10 Sub Disposition,11 Comments/Jing,12 CSD price,
# 13 OA price,14 Confirmation price changes,15 % (not copied - it is a formula)
_FULL = [("F", 3), ("G", 4), ("H", 5), ("I", 6), ("J", 7), ("K", 8), ("L", 9),
         ("M", 10), ("N", 11), ("O", 12), ("P", 13), ("Q", 14)]
COLUMN_MAP = {
    "CONF": _FULL,
    "Tracking": _FULL,
    "Certificate": _FULL[:9],          # F..N
    "Moved to FAB": _FULL[:4],         # F..I
}
# sheets whose columns R and S hold the % price formulas
FORMULA_SHEETS = {"CONF", "Tracking"}

# Names that differ between the EOD files and the report
NAME_ALIASES = {"marycris": "Marycris Senton"}

DATE_STYLE = 174        # number format 14 (date) - exists in the report
EOD_COLS = 16


class ReportError(Exception):
    """A problem the user should read and fix (shown in red)."""


# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------

def norm_name(s):
    s = unicodedata.normalize("NFKC", str(s)).replace("\u00a0", " ")
    return re.sub(r"\s+", " ", s).strip().casefold()


def excel_serial(d):
    if isinstance(d, dt.datetime):
        d = d.date()
    return (d - dt.date(1899, 12, 30)).days


def week_label(d):
    """Week 1 = the Mon-Sun week that holds the first weekday of the month."""
    first = d.replace(day=1)
    while first.weekday() >= 5:
        first += dt.timedelta(days=1)
    monday = first - dt.timedelta(days=first.weekday())
    return "Week %d" % ((d - monday).days // 7 + 1)


def col_letters(ref):
    return re.match(r"[A-Z]+", ref).group(0)


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _text(s):
    return escape(_ILLEGAL.sub("", str(s)))


# ----------------------------------------------------------------------------
# Reading the agents' EOD workbooks
# ----------------------------------------------------------------------------

def read_eod(file_obj, filename="file"):
    """Return list of 16-tuples (EOD columns) for every filled row."""
    try:
        wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    except Exception as exc:
        raise ReportError("%s: cannot be opened as an Excel file (%s)" % (filename, exc))
    if "EOD" not in wb.sheetnames:
        raise ReportError("%s: there is no sheet named 'EOD'." % filename)
    rows = []
    header_seen = False
    for r in wb["EOD"].iter_rows(values_only=True):
        r = tuple(r[:EOD_COLS]) + (None,) * max(0, EOD_COLS - len(r))
        if not header_seen:
            if r[0] is not None and str(r[0]).strip().lower() == "date":
                head = [norm_name(x) if x is not None else "" for x in r]
                if head[1] != "name" or head[2] != "task" or not head[3].startswith("purchase"):
                    raise ReportError("%s: the EOD header row is not in the expected order "
                                      "(Date, Name, Task, Purchase Order #...)." % filename)
                header_seen = True
            continue
        if r[0] is None or r[3] is None or str(r[3]).strip() == "":
            continue                       # blank / filler rows
        rows.append(r)
    wb.close()
    if not header_seen:
        raise ReportError("%s: could not find the header row starting with 'Date'." % filename)
    return rows


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def summarize_files(files):
    """files: list of (filename, bytes). Returns dict used by the UI."""
    out = []
    all_dates = []
    for fname, data in files:
        rows = read_eod(io.BytesIO(data), fname)
        dates = sorted({to_date(r[0]) for r in rows if to_date(r[0])})
        names = sorted({str(r[1]).strip() for r in rows if r[1] is not None})
        all_dates.extend(dates)
        out.append({"file": fname, "rows": rows, "dates": dates, "names": names})
    latest = max(all_dates) if all_dates else None
    return out, latest


# ----------------------------------------------------------------------------
# Reading the report package
# ----------------------------------------------------------------------------

def _sheet_paths(z):
    wbx = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid2t = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        rid = re.search(r'\bId="([^"]+)"', tag).group(1)
        tgt = re.search(r'\bTarget="([^"]+)"', tag).group(1)
        rid2t[rid] = tgt
    paths = {}
    for m in re.finditer(r"<sheet\b[^>]*>", wbx):
        tag = m.group(0)
        name = re.search(r'\bname="([^"]*)"', tag).group(1)
        name = (name.replace("&apos;", "'").replace("&quot;", '"')
                .replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&"))
        rid = re.search(r'\br:id="([^"]+)"', tag).group(1)
        tgt = rid2t[rid]
        paths[name] = tgt.lstrip("/") if tgt.startswith("/") else "xl/" + tgt
    return paths


def _shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    s = z.read("xl/sharedStrings.xml").decode("utf8")
    out = []
    for si in re.finditer(r"<si\b[^>]*>(.*?)</si>", s, re.S):
        parts = re.findall(r"<t\b[^>]*>(.*?)</t>", si.group(1), re.S)
        txt = "".join(parts)
        txt = (txt.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
               .replace("&apos;", "'").replace("&amp;", "&"))
        out.append(txt)
    return out


def month_sheet_name(sheet_names, d):
    key = d.strftime("%b").lower()
    yy = d.year % 100
    for n in sheet_names:
        m = re.match(r"^([A-Za-z]+)'(\d\d)$", n)
        if m and m.group(1)[:3].lower() == key and int(m.group(2)) == yy:
            return n
    return None


def _month_sheet_agents(sheet_xml, sst):
    """Agent names in row 7 of the month sheet (columns B... until 'Total')."""
    m = re.search(r'<row r="7"[^>]*>(.*?)</row>', sheet_xml, re.S)
    names = []
    if not m:
        return names
    for c in re.finditer(r'<c r="([A-Z]+)7"([^>]*)>(?:<v>(.*?)</v>)?', m.group(1)):
        col, attrs, v = c.group(1), c.group(2), c.group(3)
        if col == "A" or v is None:
            continue
        txt = sst[int(v)] if 't="s"' in attrs else v
        if norm_name(txt) == "total":
            break
        names.append(txt)
    return names


def read_report_info(report_bytes, load_date):
    """Quick look used by the UI: month sheet and the agent names."""
    z = zipfile.ZipFile(io.BytesIO(report_bytes))
    paths = _sheet_paths(z)
    msheet = month_sheet_name(paths.keys(), load_date)
    if not msheet:
        return None, []
    sst = _shared_strings(z)
    xml = z.read(paths[msheet]).decode("utf8")
    return msheet, _month_sheet_agents(xml, sst)


def resolve_name(file_name, report_names, overrides=None):
    """Map an EOD name to a report name. Returns None when unknown."""
    overrides = overrides or {}
    if file_name in overrides and overrides[file_name]:
        return overrides[file_name]
    n = norm_name(file_name)
    lookup = {norm_name(x): x for x in report_names}
    if n in lookup:
        return lookup[n]
    first = n.split(" ")[0]
    if first in NAME_ALIASES and NAME_ALIASES[first] in report_names:
        return NAME_ALIASES[first]
    cands = [x for x in report_names if norm_name(x).split(" ")[0] == first]
    if len(cands) == 1:
        return cands[0]
    return None


# ----------------------------------------------------------------------------
# Building the new rows
# ----------------------------------------------------------------------------

def _template_cells(sheet_xml):
    """Return (last_row_number, [(column_letters, style)]) from the last row."""
    i = sheet_xml.rfind("<row ")
    if i < 0:
        raise ReportError("A data sheet has no rows.")
    last = sheet_xml[i:sheet_xml.rfind("</sheetData>")]
    last_no = int(re.match(r'<row r="(\d+)"', last).group(1))
    cells = []
    for m in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>)', last):
        sm = re.search(r'\bs="(\d+)"', m.group(2))
        cells.append((m.group(1), int(sm.group(1)) if sm else 0))
    row_tag = re.match(r"<row [^>]*>", last).group(0)
    spans = re.search(r'spans="([^"]+)"', row_tag)
    ht = re.search(r'\bht="([^"]+)"', row_tag)
    return last_no, cells, (spans.group(1) if spans else None), (ht.group(1) if ht else "15")


def _value_cell(ref, style, v):
    """XML for one value cell. v may be None."""
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return '<c r="%s" s="%d"/>' % (ref, style)
    if isinstance(v, (dt.datetime, dt.date)):
        return '<c r="%s" s="%d"><v>%d</v></c>' % (ref, DATE_STYLE, excel_serial(v))
    if isinstance(v, bool):
        return '<c r="%s" s="%d" t="b"><v>%d</v></c>' % (ref, style, int(v))
    if isinstance(v, int):
        return '<c r="%s" s="%d"><v>%d</v></c>' % (ref, style, v)
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return '<c r="%s" s="%d"><v>%d</v></c>' % (ref, style, int(v))
        return '<c r="%s" s="%d"><v>%r</v></c>' % (ref, style, v)
    return ('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, style, _text(v)))


def _pct_cached(o, p):
    """Replicates IFERROR(1-O/P,"-") for the cached value."""
    def num(x):
        if x is None or (isinstance(x, str) and x.strip() == ""):
            return 0.0
        if isinstance(x, bool) or not isinstance(x, (int, float)):
            return None
        return float(x)
    on, pn = num(o), num(p)
    if on is None or pn is None or pn == 0:
        return "-"
    return 1 - on / pn


def _formula_cells(r, style_r, style_s, o, p):
    pct = _pct_cached(o, p)
    fr = 'IFERROR(1-O%d/P%d,"-")' % (r, r)
    fs = ('IF(R%d&lt;=5%%,"Below 5%%",IF(AND(R%d&gt;=5%%,R%d&lt;=60),"Above 5%%","-"))' % (r, r, r))
    if isinstance(pct, str):
        cr = '<c r="R%d" s="%d" t="str"><f>%s</f><v>-</v></c>' % (r, style_r, fr)
        sv = "-"
    else:
        cr = '<c r="R%d" s="%d"><f>%s</f><v>%r</v></c>' % (r, style_r, fr, pct)
        sv = "Below 5%" if pct <= 0.05 else ("Above 5%" if pct <= 60 else "-")
    cs = '<c r="S%d" s="%d" t="str"><f>%s</f><v>%s</v></c>' % (r, style_s, fs, sv)
    return cr, cs


def build_rows(sheet_name, sheet_xml, entries, month_label):
    """entries: list of (report_name, eod_row, date). Returns (new_xml_rows, last_no, count)."""
    last_no, tcells, spans, ht = _template_cells(sheet_xml)
    style = {c: s for c, s in tcells}
    max_col = max(col_index(c) for c, _ in tcells)
    letters_in_order = [c for c, _ in tcells]
    cmap = {c: i for c, i in COLUMN_MAP[sheet_name]}
    out = []
    r = last_no
    for name, row, d in entries:
        r += 1
        cells = []
        for col in letters_in_order:
            ref = "%s%d" % (col, r)
            st = style.get(col, 0)
            if col == "A":
                cells.append(_value_cell(ref, st, month_label))
            elif col == "B":
                cells.append('<c r="%s" s="%d"><v>%d</v></c>' % (ref, st, excel_serial(d)))
            elif col == "C":
                cells.append(_value_cell(ref, st, week_label(d)))
            elif col == "E":
                cells.append(_value_cell(ref, st, name))
            elif sheet_name in FORMULA_SHEETS and col == "R":
                cr, cs = _formula_cells(r, st, style.get("S", 0), row[12], row[13])
                cells.append(cr)
            elif sheet_name in FORMULA_SHEETS and col == "S":
                cells.append(cs)
            elif col in cmap:
                cells.append(_value_cell(ref, st, row[cmap[col]]))
            else:
                cells.append('<c r="%s" s="%d"/>' % (ref, st))
        attrs = 'r="%d"' % r
        if spans:
            attrs += ' spans="%s"' % spans
        attrs += ' ht="%s" customHeight="1"' % ht
        out.append("<row %s>%s</row>" % (attrs, "".join(cells)))
    return "".join(out), r, len(entries)


def existing_keys(sheet_xml, sst, serial):
    """Set of normalised names that already have rows for this date serial."""
    found = set()
    pat = re.compile(r'<c r="B\d+"[^>]*><v>%d</v></c>' % serial)
    for m in pat.finditer(sheet_xml):
        # find the row this cell belongs to and read its E cell
        start = sheet_xml.rfind("<row ", 0, m.start())
        end = sheet_xml.find("</row>", m.end())
        row = sheet_xml[start:end]
        e = re.search(r'<c r="E\d+"([^>]*)>(?:<v>(.*?)</v>|<is>.*?<t[^>]*>(.*?)</t>)', row, re.S)
        if not e:
            continue
        if 't="s"' in e.group(1) and e.group(2) is not None:
            txt = sst[int(e.group(2))]
        else:
            txt = e.group(3) or e.group(2) or ""
        found.add(norm_name(txt))
    return found


# ----------------------------------------------------------------------------
# The update
# ----------------------------------------------------------------------------

def update_report(report_bytes, summaries, load_date, name_overrides=None):
    """
    summaries: output of summarize_files()[0]
    Returns (new_report_bytes, messages) where messages = [(level, text)]
    level is 'ok', 'info', 'warn'. Raises ReportError for blocking problems.
    """
    msgs = []
    zin = zipfile.ZipFile(io.BytesIO(report_bytes))
    paths = _sheet_paths(zin)
    for needed in set(TASK_TO_SHEET.values()):
        if needed not in paths:
            raise ReportError("The report has no sheet named '%s'." % needed)

    msheet = month_sheet_name(paths.keys(), load_date)
    if not msheet:
        raise ReportError(
            "The report has no month sheet for %s. Copy the previous month's sheet, rename it "
            "(for example Dec'26) and set up the Daily and Weekly blocks, then upload that "
            "report first." % load_date.strftime("%B %Y"))
    sst = _shared_strings(zin)
    month_xml = zin.read(paths[msheet]).decode("utf8")
    report_names = _month_sheet_agents(month_xml, sst)
    if not report_names:
        raise ReportError("Could not read the agent names from row 7 of '%s'." % msheet)

    # ---- sort the agents' rows into sheets -------------------------------
    per_sheet = {s: [] for s in set(TASK_TO_SHEET.values())}
    skipped_tasks = {}
    unknown_names = set()
    seen = {}
    for sm in summaries:
        n_file = 0
        for row in sm["rows"]:
            d = to_date(row[0])
            if d != load_date:
                continue
            rname = resolve_name(str(row[1]).strip(), report_names, name_overrides)
            if rname is None:
                unknown_names.add(str(row[1]).strip())
                continue
            sheet = TASK_TO_SHEET.get(norm_name(row[2]))
            if sheet is None:
                key = (str(row[2]).strip(), rname)
                skipped_tasks[key] = skipped_tasks.get(key, 0) + 1
                continue
            per_sheet[sheet].append((rname, row, d))
            seen[(sheet, rname)] = seen.get((sheet, rname), 0) + 1
            n_file += 1
        if n_file == 0:
            msgs.append(("warn", "%s: no rows for %s." % (sm["file"], load_date.strftime("%d-%b-%Y"))))
    if unknown_names:
        raise ReportError("These names are not in the report's agent list (%s): %s. "
                          "Pick the right agent in the drop-down." %
                          (", ".join(report_names), ", ".join(sorted(unknown_names))))
    if not any(per_sheet.values()):
        raise ReportError("There are no rows for %s in the files you selected." %
                          load_date.strftime("%d-%b-%Y"))

    # ---- double-count guard ----------------------------------------------
    serial = excel_serial(load_date)
    sheet_xml = {}
    for sheet in per_sheet:
        if not per_sheet[sheet]:
            continue
        sheet_xml[sheet] = zin.read(paths[sheet]).decode("utf8")
        have = existing_keys(sheet_xml[sheet], sst, serial)
        for name in {n for n, _, _ in per_sheet[sheet]}:
            if norm_name(name) in have:
                raise ReportError("Already loaded: %s / %s / %s. Nothing was changed, so the day "
                                  "cannot be counted twice." %
                                  (sheet, name, load_date.strftime("%d-%b-%Y")))

    # ---- patch -----------------------------------------------------------
    patched = {}
    for sheet, entries in per_sheet.items():
        if not entries:
            continue
        xml = sheet_xml[sheet]
        new_rows, last_no, count = build_rows(sheet, xml, entries, msheet)
        i = xml.rfind("</sheetData>")
        xml = xml[:i] + new_rows + xml[i:]
        xml = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"/>',
                     lambda m: '<dimension ref="%s:%s%d"/>' % (m.group(1), m.group(2), last_no),
                     xml, count=1)
        patched[paths[sheet]] = xml
        msgs.append(("ok", "%s: added %d rows." % (sheet, count)))

    # month sheet date
    m = re.search(r'(<c r="B6"[^>]*>)<v>[^<]*</v></c>', month_xml)
    if not m:
        raise ReportError("Cell B6 of '%s' (the report date) was not found." % msheet)
    month_xml = month_xml[:m.start()] + m.group(1) + "<v>%d</v></c>" % serial + month_xml[m.end():]
    patched[paths[msheet]] = month_xml
    msgs.append(("ok", "%s: report date set to %s." % (msheet, load_date.strftime("%d-%b-%Y"))))

    for (task, who), n in sorted(skipped_tasks.items()):
        msgs.append(("warn", "%s / %s: %d rows skipped (task '%s' has no sheet in the report)."
                     % (who, task, n, task)))

    # workbook.xml: recalc on open ; drop calcChain
    wbx = zin.read("xl/workbook.xml").decode("utf8")
    if "<calcPr" in wbx:
        wbx = re.sub(r"<calcPr\b([^>]*?)(/?)>",
                     lambda mm: "<calcPr%s fullCalcOnLoad=\"1\"%s>" % (
                         re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", mm.group(1)), mm.group(2)),
                     wbx, count=1)
    patched["xl/workbook.xml"] = wbx
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rels = re.sub(r"<Relationship\b[^>]*calcChain[^>]*/>", "", rels)
    patched["xl/_rels/workbook.xml.rels"] = rels
    ctypes = zin.read("[Content_Types].xml").decode("utf8")
    ctypes = re.sub(r"<Override\b[^>]*calcChain[^>]*/>", "", ctypes)
    patched["[Content_Types].xml"] = ctypes

    # ---- write the new package -------------------------------------------
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
        for info in zin.infolist():
            if info.filename == "xl/calcChain.xml":
                continue
            if info.filename in patched:
                zout.writestr(info.filename, patched[info.filename].encode("utf8"))
            else:
                zi = zipfile.ZipInfo(info.filename, info.date_time)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = info.external_attr
                with zin.open(info) as src, zout.open(zi, "w", force_zip64=True) as dst:
                    while True:
                        chunk = src.read(1024 * 1024)
                        if not chunk:
                            break
                        dst.write(chunk)
    return buf.getvalue(), msgs
