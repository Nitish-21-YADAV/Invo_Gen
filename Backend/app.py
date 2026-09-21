"""
Backend: Flask + DIRECT XLSX (zip + XML) editing

⚠️ IMPORTANT:
  Hum template ko openpyxl se SAVE NAHI kar rahe, kyunki openpyxl
  images/logo/charts ko drop kar deta hai.

  Instead, hum template .xlsx ko zip ki tarah kholte hai aur SIRF un
  cells ka XML badalte hai jo instruction me diye gaye hai:
      E4, E8, E9, E10, E11, E12, E15, E16, B23, E23
  Baaki har file/byte (logo, image, style, font, formula, sheet
  properties, comments, etc.) template se EXACTLY same copy hota hai.

Address splitting:
  - Greedy word-wrap: line <= 40 chars, mid-word cut NAHI hota
  - Max 4 lines (E9..E12). Usse zyada ho to extra drop + warning.
"""

import re
import zipfile
from pathlib import Path

import openpyxl   # sirf 2nd Excel (data file) padhne ke liye
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS

# ---------------------------------------------------------------- paths
BASE_DIR   = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "OutputFolder"
TMP_DIR    = BASE_DIR / "_tmp"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
TMP_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------- config
ADDRESS_LINE_LEN  = 38   # E9..E12 — har line me max 40 chars
ADDRESS_MAX_LINES = 4    # E9, E10, E11, E12

# Cells jo hum template me overwrite karenge (aur kuch nahi)
TARGET_CELLS = ("E4", "E8", "E9", "E10", "E11", "E12",
                "E15", "E16", "B23", "E23")

app = Flask(__name__)
CORS(app)


# ================================================================ helpers
def to_str(v) -> str:
    """Excel value -> safe string  (1078.0 -> '1078')."""
    if v is None:
        return ""
    s = str(v).strip()
    if re.fullmatch(r"-?\d+\.0", s):
        s = s[:-2]
    return s


def split_address(addr, size=ADDRESS_LINE_LEN, max_lines=ADDRESS_MAX_LINES):
    """
    Greedy word-wrap. Har line <= size chars.
    Mid-word cut NAHI karta (jab tak ek word khud hi size se bada na ho).
    Max max_lines lines. Zyada ho to extra drop + warning.
    """
    text = re.sub(r"\s+", " ", to_str(addr)).strip()
    if not text:
        return [""]

    words = text.split(" ")
    lines = []
    current = ""

    for w in words:
        # Agar word khud hi `size` se bada hai -> usko chunks me todo
        while len(w) > size:
            if current:
                lines.append(current)
                current = ""
            lines.append(w[:size])
            w = w[size:]

        candidate = (current + " " + w) if current else w
        if len(candidate) <= size:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = w

    if current:
        lines.append(current)

    # Max lines enforce
    if len(lines) > max_lines:
        overflow = lines[max_lines:]
        print(f"[WARN] Address ne {len(lines)} lines maangi, "
              f"max {max_lines} hi allowed. Extra drop: {overflow!r}")
        lines = lines[:max_lines]

    return lines


def xml_escape(s: str) -> str:
    return (str(s).replace("&", "&amp;")
                  .replace("<", "&lt;")
                  .replace(">", "&gt;"))


def find_sheet_xml_path(zf: zipfile.ZipFile) -> str:
    """Pehli worksheet xml file ka naam lauta do."""
    for n in zf.namelist():
        if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            return n
    raise RuntimeError("Template me koi sheet XML nahi mili")


# ---------------------------------------------------------------- cell edit
_CELL_PATTERN_CACHE = {}

def edit_cell_in_xml(xml_str: str, cell_ref: str, value):
    """
    Sheet XML me sirf `<c r="E4" ...>...</c>` wale cell ka value badlo.
    Baaki XML ka ek byte bhi nahi chhua jaata.
    Return: (new_xml, found_bool)
    """
    pat = _CELL_PATTERN_CACHE.get(cell_ref)
    if pat is None:
        pat = re.compile(
            rf'<c\s+r="{re.escape(cell_ref)}"([^>]*?)(?:/>|>(.*?)</c>)',
            re.DOTALL,
        )
        _CELL_PATTERN_CACHE[cell_ref] = pat

    def repl(m: re.Match) -> str:
        attrs = (m.group(1) or "").rstrip()
        # purana type attribute hatao (naya khud set karenge)
        attrs = re.sub(r'\s+t="[^"]*"', "", attrs)

        if value is None or value == "":
            return f'<c r="{cell_ref}"{attrs}/>'

        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return f'<c r="{cell_ref}"{attrs}><v>{value}</v></c>'

        # Text -> inline string (sharedStrings ko touch nahi karte)
        esc = xml_escape(value)
        return (f'<c r="{cell_ref}"{attrs} t="inlineStr">'
                f'<is><t xml:space="preserve">{esc}</t></is></c>')

    new_xml, n = pat.subn(repl, xml_str)
    return new_xml, n > 0


# ---------------------------------------------------------------- build file
def build_one_file(template_path: Path, out_path: Path, row: dict):
    """
    Template ko zip ki tarah kholo -> sirf target cells ka XML replace karo
    -> naya xlsx likho. Baaki SAB (logo, images, styles, formulas) AS-IS.
    """
    # Pehle saare naye values ready karo
    edits = {
        "E4":  row["inv_no"] if row["inv_no"] is not None else "",
        "E8":  to_str(row["rs_name"]),
        "E15": to_str(row["contact_person"]),
        "E16": to_str(row["mobile"]),
        "B23": row["box"] if row["box"] is not None else "",
        "E23": row["qty"] if row["qty"] is not None else "",
    }

    addr_lines = split_address(row["address"])
    for i in range(ADDRESS_MAX_LINES):
        edits[f"E{9 + i}"] = addr_lines[i] if i < len(addr_lines) else ""

    with zipfile.ZipFile(template_path, "r") as zin:
        sheet_xml_path = find_sheet_xml_path(zin)
        sheet_xml = zin.read(sheet_xml_path).decode("utf-8")

        # sirf target cells replace karo
        for cell_ref, val in edits.items():
            sheet_xml, ok = edit_cell_in_xml(sheet_xml, cell_ref, val)
            if not ok:
                print(f"[WARN] cell {cell_ref} template me nahi mila — skip")

        sheet_bytes = sheet_xml.encode("utf-8")

        # Naya zip banao — sab kuch template se AS-IS, sirf sheet xml replace
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == sheet_xml_path:
                    zout.writestr(item, sheet_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))


# ---------------------------------------------------------------- data reader
def read_data_rows(data_path: Path):
    """
    2nd Excel (data file) padhta hai. Formula wale Inv No (=H2+1)
    ko bhi resolve karta hai.
    """
    wb_calc = openpyxl.load_workbook(data_path, data_only=True)
    ws_calc = wb_calc.active
    wb_raw  = openpyxl.load_workbook(data_path, data_only=False)
    ws_raw  = wb_raw.active

    max_row = ws_raw.max_row

    # ---- Inv No resolve (col H = 8)
    inv_map: dict[int, object] = {}
    for r in range(2, max_row + 1):
        cached = ws_calc.cell(r, 8).value
        if cached is not None and not (isinstance(cached, str) and cached.startswith("=")):
            inv_map[r] = cached
            continue

        raw = ws_raw.cell(r, 8).value
        if isinstance(raw, str) and raw.startswith("="):
            m = re.fullmatch(r"=H(\d+)\s*([+\-])\s*(\d+)", raw.replace(" ", ""))
            if m:
                ref_row, op, num = int(m.group(1)), m.group(2), int(m.group(3))
                base = inv_map.get(ref_row)
                if base is None:
                    base = (ws_calc.cell(ref_row, 8).value
                            or ws_raw.cell(ref_row, 8).value)
                try:
                    bf = float(base)
                    inv_map[r] = bf + num if op == "+" else bf - num
                except (TypeError, ValueError):
                    inv_map[r] = raw
            else:
                inv_map[r] = raw
        else:
            inv_map[r] = raw

    # ---- baaki columns collect
    rows = []
    for r in range(2, max_row + 1):
        rs_name = ws_calc.cell(r, 1).value
        if rs_name is None or str(rs_name).strip() == "":
            continue
        rows.append({
            "_row":           r,
            "rs_name":        ws_calc.cell(r, 1).value,   # A
            "address":        ws_calc.cell(r, 2).value,   # B
            "contact_person": ws_calc.cell(r, 3).value,   # C
            "mobile":         ws_calc.cell(r, 4).value,   # D
            "qty":            ws_calc.cell(r, 5).value,   # E
            "box":            ws_calc.cell(r, 6).value,   # F
            "wt":             ws_calc.cell(r, 7).value,   # G
            "inv_no":         inv_map.get(r),             # H
        })
    return rows


# ---------------------------------------------------------------- generate all
def generate_all(template_path: Path, data_path: Path, out_dir: Path):
    if out_dir.exists():
        for f in out_dir.iterdir():
            if f.is_file():
                f.unlink()
    else:
        out_dir.mkdir(parents=True, exist_ok=True)

    rows = read_data_rows(data_path)
    created = []
    for row in rows:
        inv_name = to_str(row["inv_no"]) or f"row{row['_row']}"
        inv_name = re.sub(r'[\\/:*?"<>|]', "_", inv_name)   # safe filename
        filename = f"{inv_name}.xlsx"
        out_path = out_dir / filename
        try:
            build_one_file(template_path, out_path, row)
            created.append(filename)
        except Exception as e:
            print(f"[ERROR] row {row['_row']}: {e}")
    return created


# ================================================================ routes
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/api/generate", methods=["POST"])
def api_generate():
    template = request.files.get("template")
    data     = request.files.get("data")

    if not template or not data:
        return jsonify({"error": "Dono files upload karo (template + data)."}), 400

    tpl_path  = TMP_DIR / "template.xlsx"
    data_path = TMP_DIR / "data.xlsx"
    template.save(tpl_path)
    data.save(data_path)

    try:
        files = generate_all(tpl_path, data_path, OUTPUT_DIR)
    except Exception as e:
        return jsonify({"error": f"Generation failed: {e}"}), 500

    return jsonify({"count": len(files), "files": files})


@app.route("/api/files", methods=["GET"])
def api_files():
    files = sorted([f.name for f in OUTPUT_DIR.glob("*.xlsx")])
    return jsonify({"count": len(files), "files": files})


@app.route("/api/download/<path:filename>", methods=["GET"])
def api_download(filename):
    safe = Path(filename).name
    if not (OUTPUT_DIR / safe).exists():
        return jsonify({"error": "File not found"}), 404
    return send_from_directory(OUTPUT_DIR, safe, as_attachment=True)


@app.route("/api/download-all", methods=["GET"])
def api_download_all():
    zip_path = TMP_DIR / "OutputFolder.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in OUTPUT_DIR.glob("*.xlsx"):
            z.write(f, arcname=f.name)
    return send_file(zip_path, as_attachment=True, download_name="OutputFolder.zip")


@app.route("/api/reset", methods=["POST"])
def api_reset():
    for f in OUTPUT_DIR.glob("*"):
        if f.is_file():
            f.unlink()
    return jsonify({"message": "OutputFolder clear ho gaya"})


# ================================================================ main
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)