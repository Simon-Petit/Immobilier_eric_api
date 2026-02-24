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
    """
    Pure function: take workbook bytes + JSON string, return modified workbook bytes.

    This is used both by the Vercel handler and by local test scripts.
    """
    if isinstance(data_str, bytes):
        data_str = data_str.decode("utf-8")

    payload = json.loads(data_str)
    wb = openpyxl.load_workbook(BytesIO(workbook_bytes))

    for i, comparable in enumerate(payload):
        sheet_name = f"Comparable_{i + 1}"
        if sheet_name not in wb.sheetnames:
            print(f"Sheet {sheet_name} not found in workbook, skipping")
            continue
        ws = wb[sheet_name]
        ws["C1"] = "Oui"
        for field in comparable:
            if field.get("value") is not None:
                ws[field["cell"]] = field["value"]

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        if self.path != "/api/write-comparables" and not self.path.endswith("write-comparables"):
            self.send_error(404)
            return

        try:
            content_type = self.headers.get("Content-Type", "")
            content_length = int(self.headers.get("Content-Length", 0))

            # Basic request logging for debugging in Vercel logs
            print("---- /api/write-comparables request ----")
            print("Method:", self.command)
            print("Path:", self.path)
            print("Content-Type:", content_type)
            print("Content-Length:", content_length)

            if content_length == 0:
                self.send_error(400, "Missing body")
                return

            body = self.rfile.read(content_length)
            env = {
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": content_type,
                "CONTENT_LENGTH": str(content_length),
            }
            form = cgi.FieldStorage(
                fp=BytesIO(body),
                environ=env,
                keep_blank_values=True,
            )

            # Log parsed form fields
            try:
                form_keys = list(form.keys())
            except Exception:
                form_keys = []
            print("Form field names:", form_keys)

            # cgi.FieldStorage is dict-like but does NOT implement .get()
            wb_field = form["workbook"] if "workbook" in form else None
            # Accept both legacy "data" and newer "comparables_array" field names
            if "data" in form:
                data_field = form["data"]
            elif "comparables_array" in form:
                data_field = form["comparables_array"]
            else:
                data_field = None

            print("Has workbook field:", bool(wb_field))
            print("Has data/comparables_array field:", bool(data_field))

            if not wb_field:
                self.send_error(400, "Missing workbook field 'workbook'")
                return

            if not data_field:
                self.send_error(400, "Missing JSON field 'data' or 'comparables_array'")
                return

            if hasattr(wb_field, "file"):
                workbook_bytes = wb_field.file.read()
            else:
                self.send_error(400, "workbook must be a file upload")
                return

            data_str = data_field.value if hasattr(data_field, "value") else data_field

            # Delegate to pure function for easier local testing
            result = apply_comparables(workbook_bytes, data_str)

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="Evaluation_Immobiliere.xlsx"')
            self.send_header("Content-Length", str(len(result)))
            self.end_headers()
            self.wfile.write(result)

        except json.JSONDecodeError as e:
            print("JSON decode error in /api/write-comparables:", repr(e))
            traceback.print_exc()
            self.send_error(400, f"Invalid JSON in data/comparables_array: {e}")
        except Exception as e:
            print("Unhandled error in /api/write-comparables:", repr(e))
            traceback.print_exc()
            self.send_error(500, str(e))
