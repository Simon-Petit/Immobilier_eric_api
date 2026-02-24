"""
Vercel serverless function: POST /api/write-comparables
Accepts multipart/form-data: workbook (xlsx file), data (JSON string of filled configs).
Returns the modified workbook as xlsx download.
"""
from http.server import BaseHTTPRequestHandler
import json
import cgi
from io import BytesIO
import os
import sys
import traceback


sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl


def apply_comparables(workbook_bytes: bytes, data_str: str) -> bytes:
    print("[apply_comparables] START")
    print(f"[apply_comparables] workbook_bytes size: {len(workbook_bytes)} bytes")
    print(f"[apply_comparables] data_str type: {type(data_str)}, length: {len(data_str)}")
    print(f"[apply_comparables] data_str preview (first 500 chars): {data_str[:500]}")

    if isinstance(data_str, bytes):
        print("[apply_comparables] data_str is bytes, decoding to utf-8")
        data_str = data_str.decode("utf-8")

    print("[apply_comparables] Parsing JSON payload...")
    payload = json.loads(data_str)
    print(f"[apply_comparables] payload type: {type(payload)}")
    print(f"[apply_comparables] number of comparables: {len(payload) if isinstance(payload, list) else 'NOT A LIST'}")

    if isinstance(payload, list):
        for idx, comp in enumerate(payload):
            print(f"[apply_comparables] comparable[{idx}] type: {type(comp)}, length: {len(comp) if isinstance(comp, list) else 'NOT A LIST'}")
            if isinstance(comp, list) and len(comp) > 0:
                print(f"[apply_comparables] comparable[{idx}] first field sample: {comp[0]}")

    print("[apply_comparables] Loading workbook with openpyxl...")
    wb = openpyxl.load_workbook(BytesIO(workbook_bytes))
    print(f"[apply_comparables] Workbook loaded. Sheet names: {wb.sheetnames}")

    for i, comparable in enumerate(payload):
        sheet_name = f"Comparable_{i + 1}"
        print(f"[apply_comparables] Processing comparable {i} → sheet '{sheet_name}'")
        if sheet_name not in wb.sheetnames:
            print(f"[apply_comparables] WARNING: Sheet '{sheet_name}' not found in workbook, skipping")
            continue
        ws = wb[sheet_name]
        ws["C1"] = "Oui"
        written = 0
        skipped = 0
        for field in comparable:
            if field.get("value") is not None:
                ws[field["cell"]] = field["value"]
                written += 1
            else:
                skipped += 1
        print(f"[apply_comparables] Sheet '{sheet_name}': {written} fields written, {skipped} skipped (null value)")

    print("[apply_comparables] Saving workbook to memory buffer...")
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    result = output.getvalue()
    print(f"[apply_comparables] Done. Output size: {len(result)} bytes")
    return result


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        print("=" * 60)
        print("[handler] Incoming POST to /api/write-comparables")
        print(f"[handler] Path check: {self.path!r}")

        if self.path != "/api/write-comparables" and not self.path.endswith("write-comparables"):
            print(f"[handler] 404 — path does not match expected route")
            self.send_error(404)
            return

        try:
            content_type = self.headers.get("Content-Type", "")
            content_length = int(self.headers.get("Content-Length", 0))

            print(f"[handler] Content-Type:   {content_type}")
            print(f"[handler] Content-Length: {content_length}")
            print(f"[handler] All headers:")
            for key, val in self.headers.items():
                print(f"[handler]   {key}: {val}")

            if content_length == 0:
                print("[handler] ERROR: Content-Length is 0 — empty body")
                self.send_error(400, "Missing body")
                return

            print(f"[handler] Reading {content_length} bytes from body...")
            body = self.rfile.read(content_length)
            print(f"[handler] Body read OK. Actual bytes read: {len(body)}")

            env = {
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": content_type,
                "CONTENT_LENGTH": str(content_length),
            }
            print("[handler] Parsing multipart form with cgi.FieldStorage...")
            form = cgi.FieldStorage(
                fp=BytesIO(body),
                environ=env,
                keep_blank_values=True,
            )

            try:
                form_keys = list(form.keys())
            except Exception as ke:
                form_keys = []
                print(f"[handler] WARNING: form.keys() raised: {ke}")
            print(f"[handler] Form field names found: {form_keys}")

            wb_field = form["workbook"] if "workbook" in form else None
            if "data" in form:
                data_field = form["data"]
                print("[handler] Using JSON field: 'data'")
            elif "comparables_array" in form:
                data_field = form["comparables_array"]
                print("[handler] Using JSON field: 'comparables_array'")
            else:
                data_field = None
                print(f"[handler] ERROR: Neither 'data' nor 'comparables_array' found in form. Got: {form_keys}")

            print(f"[handler] wb_field present: {wb_field is not None}")
            print(f"[handler] data_field present: {data_field is not None}")
            if wb_field:
                print(f"[handler] wb_field type: {type(wb_field)}, has .file: {hasattr(wb_field, 'file')}")
            if data_field:
                print(f"[handler] data_field type: {type(data_field)}, has .value: {hasattr(data_field, 'value')}")

            if not wb_field:
                print("[handler] ERROR: Missing 'workbook' field")
                self.send_error(400, "Missing workbook field 'workbook'")
                return

            if not data_field:
                print("[handler] ERROR: Missing JSON payload field")
                self.send_error(400, "Missing JSON field 'data' or 'comparables_array'")
                return

            if hasattr(wb_field, "file"):
                workbook_bytes = wb_field.file.read()
                print(f"[handler] Workbook file read: {len(workbook_bytes)} bytes, filename: {getattr(wb_field, 'filename', 'unknown')}")
            else:
                print(f"[handler] ERROR: wb_field has no .file attribute. Type: {type(wb_field)}, value: {repr(wb_field)[:200]}")
                self.send_error(400, "workbook must be a file upload")
                return

            data_str = data_field.value if hasattr(data_field, "value") else data_field
            if isinstance(data_str, bytes):
                data_str = data_str.decode("utf-8")
            print(f"[handler] JSON payload length: {len(data_str)} chars")
            print(f"[handler] JSON payload preview: {data_str[:300]}")

            print("[handler] Calling apply_comparables()...")
            result = apply_comparables(workbook_bytes, data_str)
            print(f"[handler] apply_comparables returned {len(result)} bytes — sending 200 response")

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="Evaluation_Immobiliere.xlsx"')
            self.send_header("Content-Length", str(len(result)))
            self.end_headers()
            self.wfile.write(result)
            print("[handler] Response sent successfully.")

        except json.JSONDecodeError as e:
            print(f"[handler] JSON decode error: {repr(e)}")
            traceback.print_exc()
            self.send_error(400, f"Invalid JSON in data/comparables_array: {e}")
        except Exception as e:
            print(f"[handler] Unhandled exception: {repr(e)}")
            traceback.print_exc()
            self.send_error(500, str(e))
