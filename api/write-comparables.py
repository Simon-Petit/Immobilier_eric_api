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

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        if self.path != "/api/write-comparables" and not self.path.endswith("write-comparables"):
            self.send_error(404)
            return

        try:
            content_type = self.headers.get("Content-Type", "")
            content_length = int(self.headers.get("Content-Length", 0))

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

            wb_field = form.get("workbook")
            data_field = form.get("data")
            if not wb_field or not data_field:
                self.send_error(400, "Missing workbook or data field")
                return

            if hasattr(wb_field, "file"):
                workbook_bytes = wb_field.file.read()
            else:
                self.send_error(400, "workbook must be a file upload")
                return

            data_str = data_field.value if hasattr(data_field, "value") else data_field
            if isinstance(data_str, bytes):
                data_str = data_str.decode("utf-8")

            payload = json.loads(data_str)
            wb = openpyxl.load_workbook(BytesIO(workbook_bytes))

            for i, comparable in enumerate(payload):
                sheet_name = f"Comparable_{i + 1}"
                if sheet_name not in wb.sheetnames:
                    continue
                ws = wb[sheet_name]
                ws["C1"] = "Oui"
                for field in comparable:
                    if field.get("value") is not None:
                        ws[field["cell"]] = field["value"]

            output = BytesIO()
            wb.save(output)
            output.seek(0)
            result = output.getvalue()

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="Evaluation_Immobiliere.xlsx"')
            self.send_header("Content-Length", str(len(result)))
            self.end_headers()
            self.wfile.write(result)

        except json.JSONDecodeError as e:
            self.send_error(400, f"Invalid JSON in data: {e}")
        except Exception as e:
            self.send_error(500, str(e))
