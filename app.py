from __future__ import annotations

import io
import json
import os
import re
import sqlite3
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "tooling.db"
AUTO_LOAD_FOLDER = os.environ.get("AUTO_LOAD_FOLDER", r"C:\Users\Will_W\OneDrive - Dell Technologies\Working file_Will W\03 Project\07 AI agent\AI_tool tracker\auto_load")

# Admin user settings (change these as needed)
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "cte2024"


def check_admin():
    """Check if user is authenticated as admin"""
    auth = request.authorization
    if auth and auth.username == ADMIN_USERNAME and auth.password == ADMIN_PASSWORD:
        return True
    return False


def create_app() -> Flask:
    app = Flask(__name__)
    
    # Configuration for production
    app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
    
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    _init_db()
    
    # Auto-load from folder on startup if configured
    auto_load_path = Path(AUTO_LOAD_FOLDER)
    if auto_load_path.exists():
        try:
            result = _load_from_folder(AUTO_LOAD_FOLDER)
            if result["success"] and result["files_processed"] > 0:
                print(f"Auto-loaded {result['files_processed']} files ({result['rows_imported']} rows) from {AUTO_LOAD_FOLDER}")
        except Exception as e:
            print(f"Auto-load failed: {e}")

    @app.get("/")
    def index() -> str:
        return render_template("dashboard.html")

    @app.get("/tooling_dashboard.html")
    def dashboard_alias() -> str:
        return render_template("dashboard.html")

    @app.get("/tooling_table.html")
    def tooling_table() -> str:
        return render_template("tooling_table.html")

    @app.get("/api/records")
    def get_records():
        search = request.args.get("search", "").strip().lower()
        records = _load_records()

        if search:
            filtered: list[dict[str, Any]] = []
            for row in records:
                if any(search in str(value).lower() for value in row.values()):
                    filtered.append(row)
            records = filtered

        headers = sorted({key for row in records for key in row.keys()})
        return jsonify({"count": len(records), "headers": headers, "records": records})

    @app.get("/api/stats")
    def get_stats():
        records = _load_records()
        stats = _build_stats(records)
        return jsonify(stats)

    @app.post("/api/upload")
    def upload_files():
        # Require admin authentication for upload
        if not check_admin():
            return jsonify({"error": "Authentication required"}), 401
            
        files = request.files.getlist("files")

        if not files:
            return jsonify({"error": "No files uploaded."}), 400

        total_rows = 0
        uploaded_files = []
        imports = []

        for file in files:
            if not file.filename:
                continue

            try:
                rows = _parse_excel(file.read(), file.filename)
            except Exception as exc:
                return (
                    jsonify(
                        {
                            "error": f"Failed to parse file '{file.filename}': {exc}",
                        }
                    ),
                    400,
                )
            batch = _insert_rows(rows, file.filename)
            total_rows += len(rows)
            uploaded_files.append(file.filename)
            imports.append(batch)

        return jsonify(
            {
                "message": "Upload complete.",
                "uploaded_files": uploaded_files,
                "rows_inserted": total_rows,
                "imports": imports,
            }
        )

    @app.delete("/api/records")
    def clear_records():
        # Require admin authentication for clearing data
        if not check_admin():
            return jsonify({"error": "Authentication required"}), 401
        _clear_records()
        return jsonify({"message": "All records cleared."})

    @app.get("/api/imports")
    def get_imports():
        return jsonify({"imports": _load_import_history()})

    @app.delete("/api/imports/<int:import_id>")
    def delete_import(import_id: int):
        # Require admin authentication for deleting imports
        if not check_admin():
            return jsonify({"error": "Authentication required"}), 401
        deleted = _delete_import(import_id)
        if not deleted:
            return jsonify({"error": "Import batch not found."}), 404
        return jsonify({"message": "Import batch deleted."})

    @app.get("/api/reports/summary.pdf")
    def download_summary_report():
        records = _load_records()
        stats = _build_stats(records)
        pdf_bytes = _build_pdf(records, stats)

        return send_file(
            pdf_bytes,
            as_attachment=True,
            download_name=f"tooling_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            mimetype="application/pdf",
        )

    @app.post("/api/auto-load")
    def auto_load_from_folder():
        folder = request.json.get("folder") if request.is_json else AUTO_LOAD_FOLDER
        result = _load_from_folder(folder)
        return jsonify(result)

    @app.get("/api/auto-load/status")
    def get_auto_load_status():
        folder = AUTO_LOAD_FOLDER
        folder_path = Path(folder)
        if not folder_path.exists():
            return jsonify({"folder": folder, "exists": False, "files": []})
        
        excel_files = list(folder_path.glob("*.xlsx")) + list(folder_path.glob("*.xls"))
        imported_files = {row["source_file"] for row in _load_import_history()}
        
        files_info = []
        for file in excel_files:
            files_info.append({
                "filename": file.name,
                "path": str(file),
                "size": file.stat().st_size,
                "modified": datetime.fromtimestamp(file.stat().st_mtime).isoformat(),
                "imported": file.name in imported_files
            })
        
        return jsonify({
            "folder": folder,
            "exists": True,
            "files": files_info,
            "total_files": len(files_info),
            "imported_count": sum(1 for f in files_info if f["imported"]),
            "pending_count": sum(1 for f in files_info if not f["imported"])
        })

    return app


def _get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _init_db() -> None:
    with _get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS import_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_file TEXT NOT NULL,
                imported_at TEXT NOT NULL,
                row_count INTEGER NOT NULL
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS tooling_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_file TEXT NOT NULL,
                uploaded_at TEXT NOT NULL,
                row_json TEXT NOT NULL,
                import_id INTEGER
            )
            """
        )

        existing_columns = {
            row["name"]
            for row in conn.execute("PRAGMA table_info(tooling_records)").fetchall()
        }
        if "import_id" not in existing_columns:
            conn.execute("ALTER TABLE tooling_records ADD COLUMN import_id INTEGER")

        legacy_groups = conn.execute(
            """
            SELECT source_file, uploaded_at, COUNT(*) AS row_count
            FROM tooling_records
            WHERE import_id IS NULL
            GROUP BY source_file, uploaded_at
            """
        ).fetchall()
        for group in legacy_groups:
            cursor = conn.execute(
                "INSERT INTO import_batches (source_file, imported_at, row_count) VALUES (?, ?, ?)",
                (group["source_file"], group["uploaded_at"], group["row_count"]),
            )
            conn.execute(
                """
                UPDATE tooling_records
                SET import_id = ?
                WHERE import_id IS NULL AND source_file = ? AND uploaded_at = ?
                """,
                (cursor.lastrowid, group["source_file"], group["uploaded_at"]),
            )
        conn.commit()


def _insert_rows(rows: list[dict[str, Any]], source_file: str) -> dict[str, Any]:
    uploaded_at = datetime.utcnow().isoformat(timespec="seconds") + "Z"

    if not rows:
        return {
            "import_id": None,
            "source_file": source_file,
            "imported_at": uploaded_at,
            "row_count": 0,
        }

    with _get_connection() as conn:
        cursor = conn.execute(
            "INSERT INTO import_batches (source_file, imported_at, row_count) VALUES (?, ?, ?)",
            (source_file, uploaded_at, len(rows)),
        )
        import_id = cursor.lastrowid

        payload = [
            (source_file, uploaded_at, json.dumps(row, ensure_ascii=False, default=str), import_id)
            for row in rows
        ]

        conn.executemany(
            "INSERT INTO tooling_records (source_file, uploaded_at, row_json, import_id) VALUES (?, ?, ?, ?)",
            payload,
        )
        conn.commit()

    return {
        "import_id": import_id,
        "source_file": source_file,
        "imported_at": uploaded_at,
        "row_count": len(rows),
    }


def _load_records() -> list[dict[str, Any]]:
    with _get_connection() as conn:
        rows = conn.execute(
            "SELECT id, source_file, uploaded_at, row_json, import_id FROM tooling_records ORDER BY id"
        ).fetchall()

    records = []
    for row in rows:
        data = json.loads(row["row_json"])
        data["_record_id"] = row["id"]
        data["_source_file"] = row["source_file"]
        data["_uploaded_at"] = row["uploaded_at"]
        data["_import_id"] = row["import_id"]
        records.append(data)
    return records


def _load_from_folder(folder_path: str) -> dict[str, Any]:
    folder = Path(folder_path)
    if not folder.exists():
        return {
            "success": False,
            "error": f"Folder does not exist: {folder_path}",
            "files_processed": 0,
            "rows_imported": 0,
        }
    
    if not folder.is_dir():
        return {
            "success": False,
            "error": f"Path is not a directory: {folder_path}",
            "files_processed": 0,
            "rows_imported": 0,
        }
    
    # Get all Excel files
    excel_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))
    
    if not excel_files:
        return {
            "success": True,
            "message": "No Excel files found in folder",
            "files_processed": 0,
            "rows_imported": 0,
        }
    
    # Get already imported files to avoid duplicates
    imported_files = {row["source_file"] for row in _load_import_history()}
    
    total_rows = 0
    processed_files = []
    skipped_files = []
    errors = []
    
    for file_path in excel_files:
        filename = file_path.name
        
        # Skip if already imported
        if filename in imported_files:
            skipped_files.append(filename)
            continue
        
        try:
            with open(file_path, "rb") as f:
                file_bytes = f.read()
            
            rows = _parse_excel(file_bytes, filename)
            if not rows:
                errors.append(f"{filename}: No data found")
                continue
            
            batch = _insert_rows(rows, filename)
            total_rows += len(rows)
            processed_files.append({
                "filename": filename,
                "rows": len(rows),
                "import_id": batch["import_id"]
            })
            
        except Exception as e:
            errors.append(f"{filename}: {str(e)}")
    
    return {
        "success": True,
        "message": f"Processed {len(processed_files)} files, imported {total_rows} rows",
        "files_processed": len(processed_files),
        "rows_imported": total_rows,
        "processed_files": processed_files,
        "skipped_files": skipped_files,
        "errors": errors,
    }


def _clear_records() -> None:
    with _get_connection() as conn:
        conn.execute("DELETE FROM tooling_records")
        conn.execute("DELETE FROM import_batches")
        conn.commit()


def _load_import_history() -> list[dict[str, Any]]:
    with _get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, source_file, imported_at, row_count
            FROM import_batches
            ORDER BY id DESC
            """
        ).fetchall()

    return [
        {
            "id": row["id"],
            "source_file": row["source_file"],
            "imported_at": row["imported_at"],
            "row_count": row["row_count"],
        }
        for row in rows
    ]


def _delete_import(import_id: int) -> bool:
    with _get_connection() as conn:
        import_row = conn.execute(
            "SELECT id FROM import_batches WHERE id = ?",
            (import_id,),
        ).fetchone()

        if import_row is None:
            return False

        conn.execute("DELETE FROM tooling_records WHERE import_id = ?", (import_id,))
        conn.execute("DELETE FROM import_batches WHERE id = ?", (import_id,))
        conn.commit()

    return True


def _parse_excel(file_bytes: bytes, source_file: str) -> list[dict[str, Any]]:
    try:
        workbook = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=0,
            header=None,
            dtype=str,
            engine="openpyxl",
            encoding='utf-8'
        )
        workbook = workbook.where(pd.notnull(workbook), "")
        # Convert all values to strings and handle encoding
        matrix = []
        for row in workbook.values.tolist():
            matrix.append([str(cell).strip() if cell is not None else "" for cell in row])
    except Exception as primary_error:
        try:
            matrix = _read_xlsx_matrix_without_styles(file_bytes)
        except Exception as fallback_error:
            raise ValueError(
                "Could not parse Excel file with pandas/openpyxl or fallback parser. "
                f"Primary: {primary_error}; Fallback: {fallback_error}"
            )

    rows = _matrix_to_records(matrix, source_file)
    if not rows:
        raise ValueError("No usable rows found in uploaded workbook")
    return rows


def _read_xlsx_matrix_without_styles(file_bytes: bytes) -> list[list[str]]:
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        shared_strings = _read_shared_strings(zf)
        worksheet_path = "xl/worksheets/sheet1.xml"
        if worksheet_path not in zf.namelist():
            worksheet_paths = [name for name in zf.namelist() if name.startswith("xl/worksheets/")]
            if not worksheet_paths:
                raise ValueError("No worksheet XML found in workbook")
            worksheet_path = worksheet_paths[0]

        xml_bytes = zf.read(worksheet_path)

    root = ET.fromstring(xml_bytes)
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    row_nodes = root.findall(".//x:sheetData/x:row", ns)
    if not row_nodes:
        return []

    grid: list[dict[int, str]] = []
    max_col = -1

    for row_node in row_nodes:
        row_values: dict[int, str] = {}
        for cell in row_node.findall("x:c", ns):
            ref = cell.attrib.get("r", "")
            col_idx = _column_index_from_ref(ref)
            if col_idx < 0:
                continue
            value = _read_cell_value(cell, ns, shared_strings)
            row_values[col_idx] = value
            max_col = max(max_col, col_idx)
        grid.append(row_values)

    if max_col < 0:
        return []

    matrix: list[list[str]] = []
    for row in grid:
        row_values = []
        for col in range(max_col + 1):
            row_values.append(str(row.get(col, "")).strip())
        matrix.append(row_values)

    return matrix


def _matrix_to_records(matrix: list[list[str]], source_file: str) -> list[dict[str, Any]]:
    if not matrix:
        return []

    # Explicitly use row 3 (index 2) as header row for the specific Excel format
    # Structure: Row 1 empty, Row 2 section headers, Row 3 column headers, Row 4+ data
    header_row_index = 2 if len(matrix) >= 3 else 0
    
    raw_headers = matrix[header_row_index] if header_row_index < len(matrix) else []
    headers = [str(h).strip() for h in raw_headers]
    
    if not any(headers):
        return []

    rows = [
        row
        for row in matrix[header_row_index + 1 :]
        if any(str(cell).strip() != "" for cell in row)
    ]

    update_date = _extract_update_date_from_filename(source_file)
    records: list[dict[str, Any]] = []

    for row in rows:
        obj: dict[str, Any] = {}
        for idx, header in enumerate(headers):
            key = header if header else f"Column_{idx + 1}"
            value = row[idx] if idx < len(row) else ""
            obj[key] = str(value).strip()
        obj["Update Date"] = update_date
        records.append(obj)

    return records


def _extract_update_date_from_filename(filename: str) -> str:
    date_match = re.search(r"(\d{4})(\d{2})(\d{2})", filename)
    if date_match:
        return f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"

    date_match = re.search(r"(\d{2})(\d{2})(\d{4})", filename)
    if date_match:
        return f"{date_match.group(3)}-{date_match.group(2)}-{date_match.group(1)}"

    date_match = re.search(r"(\d{8})", filename)
    if date_match:
        digits = date_match.group(1)
        return f"{digits[0:4]}-{digits[4:6]}-{digits[6:8]}"

    return ""


def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    path = "xl/sharedStrings.xml"
    if path not in zf.namelist():
        return []

    root = ET.fromstring(zf.read(path))
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    strings: list[str] = []

    for si in root.findall("x:si", ns):
        parts = [t.text or "" for t in si.findall(".//x:t", ns)]
        strings.append("".join(parts))

    return strings


def _read_cell_value(cell: ET.Element, ns: dict[str, str], shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = cell.find("x:v", ns)
    inline_node = cell.find("x:is/x:t", ns)

    if cell_type == "inlineStr" and inline_node is not None:
        return inline_node.text or ""

    if value_node is None or value_node.text is None:
        return ""

    raw = value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except Exception:
            return raw

    return raw


def _column_index_from_ref(ref: str) -> int:
    letters = ""
    for ch in ref:
        if ch.isalpha():
            letters += ch.upper()
        else:
            break

    if not letters:
        return -1

    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _build_stats(records: list[dict[str, Any]]) -> dict[str, int]:
    status_columns = []
    if records:
        for header in records[0].keys():
            lower = header.lower()
            if "status" in lower and "build complete" not in lower and "aggregate" not in lower:
                status_columns.append(header)

    green = yellow = red = 0

    for row in records:
        values = [str(row.get(col, "")).strip().upper() for col in status_columns]
        if any(v in {"R", "RED", "3"} for v in values):
            red += 1
        elif any(v in {"Y", "YELLOW", "2"} for v in values):
            yellow += 1
        elif any(v in {"G", "GREEN", "1"} for v in values):
            green += 1

    return {
        "total": len(records),
        "green": green,
        "yellow": yellow,
        "red": red,
    }


def _build_pdf(records: list[dict[str, Any]], stats: dict[str, int]) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()

    content = [
        Paragraph("Tooling Tracker Summary Report", styles["Title"]),
        Spacer(1, 8),
        Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]),
        Spacer(1, 12),
    ]

    stats_table_data = [
        ["Metric", "Value"],
        ["Total Records", str(stats["total"])],
        ["Green", str(stats["green"])],
        ["Yellow", str(stats["yellow"])],
        ["Red", str(stats["red"])],
    ]
    stats_table = Table(stats_table_data, hAlign="LEFT")
    stats_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e3a8a")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ]
        )
    )
    content.extend([stats_table, Spacer(1, 16)])

    preview_rows = records[:20]
    if preview_rows:
        headers = [h for h in preview_rows[0].keys() if not h.startswith("_")][:6]
        table_data = [headers]

        for row in preview_rows:
            table_data.append([str(row.get(h, ""))[:40] for h in headers])

        detail_table = Table(table_data, repeatRows=1)
        detail_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3b82f6")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                ]
            )
        )
        content.extend([
            Paragraph("Record Preview (first 20 rows)", styles["Heading3"]),
            Spacer(1, 8),
            detail_table,
        ])

    doc.build(content)
    buffer.seek(0)
    return buffer


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
