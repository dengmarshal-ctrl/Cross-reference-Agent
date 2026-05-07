from __future__ import annotations

import base64
import json
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from app.ooxml_processor import analyze_docx, audit_to_json, create_sample_docx, process_docx


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8000


class DemoHandler(BaseHTTPRequestHandler):
    server_version = "CSRDocAgentDemo/0.1"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return
        if parsed.path.startswith("/static/"):
            relative = parsed.path.removeprefix("/static/")
            self._send_static(relative)
            return
        if parsed.path == "/api/sample":
            sample = create_sample_docx()
            self._send_bytes(
                sample,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                extra_headers={
                    "Content-Disposition": 'attachment; filename="csr-crossref-sample.docx"'
                },
            )
            return
        if parsed.path == "/api/health":
            self._send_json({"ok": True})
            return
        self._send_error(HTTPStatus.NOT_FOUND, "Not found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)
        try:
            docx_bytes = self._read_body()
            if parsed.path == "/api/analyze":
                plan = analyze_docx(docx_bytes)
                self._send_json(plan)
                return
            if parsed.path == "/api/process":
                output_bytes, audit = process_docx(docx_bytes)
                filename = query.get("filename", ["processed.docx"])[0]
                output_name = _processed_filename(filename)
                self._send_json(
                    {
                        "filename": output_name,
                        "docx_base64": base64.b64encode(output_bytes).decode("ascii"),
                        "audit_filename": output_name.replace(".docx", "-audit.json"),
                        "audit_base64": base64.b64encode(audit_to_json(audit)).decode("ascii"),
                        "audit": audit,
                    }
                )
                return
            self._send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:  # pragma: no cover - intentionally returned to demo UI.
            self._send_error(HTTPStatus.BAD_REQUEST, str(exc))

    def log_message(self, format: str, *args: object) -> None:
        # Keep local demo output concise; the UI shows structured logs.
        return

    def _read_body(self) -> bytes:
        content_length = int(self.headers.get("Content-Length", "0"))
        return self.rfile.read(content_length)

    def _send_static(self, relative: str) -> None:
        safe_relative = Path(relative)
        if safe_relative.is_absolute() or ".." in safe_relative.parts:
            self._send_error(HTTPStatus.BAD_REQUEST, "Invalid static path")
            return
        path = STATIC_DIR / safe_relative
        suffix = path.suffix.lower()
        content_type = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".svg": "image/svg+xml",
        }.get(suffix, "application/octet-stream")
        self._send_file(path, content_type)

    def _send_file(self, path: Path, content_type: str) -> None:
        if not path.exists() or not path.is_file():
            self._send_error(HTTPStatus.NOT_FOUND, "File not found")
            return
        self._send_bytes(path.read_bytes(), content_type)

    def _send_json(self, data: dict) -> None:
        body = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        self._send_bytes(body, "application/json; charset=utf-8")

    def _send_error(self, status: HTTPStatus, message: str) -> None:
        body = json.dumps({"error": message}, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_bytes(
        self,
        body: bytes,
        content_type: str,
        extra_headers: dict[str, str] | None = None,
    ) -> None:
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        if extra_headers:
            for key, value in extra_headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)


def _processed_filename(filename: str) -> str:
    stem = Path(filename).stem or "processed"
    return f"{stem}-csr-crossref-demo.docx"


def run(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT) -> None:
    server = ThreadingHTTPServer((host, port), DemoHandler)
    print(f"CSR 文档题注与交叉引用治理 Demo 已启动: http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    run()

