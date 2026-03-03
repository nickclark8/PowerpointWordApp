from __future__ import annotations

from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
import cgi

from services.ppt_summarizer import build_exec_summary, extract_slide_content, generate_exec_docx

HOST = "0.0.0.0"
PORT = 5000
MAX_FILE_SIZE = 30 * 1024 * 1024


def render_index(error: str = "") -> bytes:
    error_html = f"<p class='error'>{error}</p>" if error else ""
    return f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>PowerPoint Executive Summarizer</title>
    <link rel="stylesheet" href="/static/style.css" />
  </head>
  <body>
    <main class="card">
      <h1>PowerPoint to Executive Word Summary</h1>
      <p>Upload a <strong>.pptx</strong> file and download an executive-style <strong>.docx</strong> summary.</p>
      {error_html}
      <form action="/summarize" method="post" enctype="multipart/form-data">
        <label for="presentation">PowerPoint File</label>
        <input id="presentation" type="file" name="presentation" accept=".pptx" required />
        <button type="submit">Generate Executive Word Document</button>
      </form>
    </main>
  </body>
</html>""".encode("utf-8")


class AppHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/":
            page = render_index()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(page)))
            self.end_headers()
            self.wfile.write(page)
            return

        if self.path == "/static/style.css":
            css = Path("static/style.css").read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/css; charset=utf-8")
            self.send_header("Content-Length", str(len(css)))
            self.end_headers()
            self.wfile.write(css)
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        if self.path != "/summarize":
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length > MAX_FILE_SIZE:
            self._send_html_error("File is too large. Limit is 30MB.", status=HTTPStatus.REQUEST_ENTITY_TOO_LARGE)
            return

        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            self._send_html_error("Form must be multipart/form-data.", status=HTTPStatus.BAD_REQUEST)
            return

        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})
        file_item = form["presentation"] if "presentation" in form else None

        if not file_item or not getattr(file_item, "filename", ""):
            self._send_html_error("Please choose a PowerPoint file (.pptx).")
            return

        filename = Path(file_item.filename).name
        if filename.rsplit(".", 1)[-1].lower() != "pptx":
            self._send_html_error("Unsupported format. Please upload a .pptx file.")
            return

        file_bytes = file_item.file.read()
        try:
            slides = extract_slide_content(file_bytes)
        except Exception:
            self._send_html_error("Could not parse presentation. Ensure the file is a valid .pptx.")
            return

        if not slides:
            self._send_html_error("No readable content found in this presentation.")
            return

        summary = build_exec_summary(slides, source_name=filename)
        docx_bytes = generate_exec_docx(summary)

        output_name = f"{Path(filename).stem}_executive_summary.docx"
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        self.send_header("Content-Disposition", f'attachment; filename="{output_name}"')
        self.send_header("Content-Length", str(len(docx_bytes)))
        self.end_headers()
        self.wfile.write(docx_bytes)

    def _send_html_error(self, message: str, status: HTTPStatus = HTTPStatus.BAD_REQUEST):
        page = render_index(message)
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(page)))
        self.end_headers()
        self.wfile.write(page)


if __name__ == "__main__":
    server = ThreadingHTTPServer((HOST, PORT), AppHandler)
    print(f"Serving on http://{HOST}:{PORT}")
    server.serve_forever()
