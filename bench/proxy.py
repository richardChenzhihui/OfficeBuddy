"""Metering proxy: forwards Anthropic-compatible requests to MiniMax and logs
per-call usage (tokens, latency) tagged with the current run id.

Both agents point their base_url at http://127.0.0.1:<port>/anthropic so token
counts and call counts are measured identically for the comparison.
"""
import json
import threading
import time
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

UPSTREAM = "https://api.minimaxi.com"
BENCH_DIR = Path(__file__).resolve().parent
RESULTS_DIR = BENCH_DIR / "results"
CURRENT_RUN_FILE = RESULTS_DIR / "current_run.txt"
LOG_FILE = RESULTS_DIR / "llm_calls.jsonl"
_log_lock = threading.Lock()

FORWARD_HEADERS = {
    "x-api-key",
    "authorization",
    "content-type",
    "anthropic-version",
    "anthropic-beta",
    "accept",
}


def current_run_id() -> str:
    try:
        return CURRENT_RUN_FILE.read_text().strip()
    except OSError:
        return "unknown"


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):  # silence default access log
        pass

    def do_POST(self):
        run_id = current_run_id()  # capture at request time, not response time
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length) if length else b""
        headers = {
            k: v for k, v in self.headers.items() if k.lower() in FORWARD_HEADERS
        }
        url = UPSTREAM + self.path
        t0 = time.time()
        status, resp_body, resp_ctype = 502, b"", "application/json"
        try:
            req = urllib.request.Request(url, data=body, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=600) as resp:
                status = resp.status
                resp_body = resp.read()
                resp_ctype = resp.headers.get("Content-Type", "application/json")
        except urllib.error.HTTPError as exc:
            status = exc.code
            resp_body = exc.read()
            resp_ctype = exc.headers.get("Content-Type", "application/json")
        except Exception as exc:  # connection errors -> 502 for client retry
            resp_body = json.dumps(
                {"type": "error", "error": {"type": "api_error", "message": str(exc)}}
            ).encode()

        if status == 200:
            try:
                json.loads(resp_body)
            except Exception:
                # 200 with an unparseable body would crash the SDK json parse;
                # rewrite to a retryable 5xx in Anthropic error shape.
                status = 502
                resp_body = json.dumps(
                    {"type": "error", "error": {"type": "api_error", "message": "upstream returned non-JSON 200"}}
                ).encode()
                resp_ctype = "application/json"

        latency = time.time() - t0
        entry = {
            "ts": time.time(),
            "run_id": run_id,
            "path": self.path,
            "status": status,
            "latency_s": round(latency, 2),
            "req_bytes": len(body),
            "resp_bytes": len(resp_body),
        }
        try:
            data = json.loads(resp_body)
            usage = data.get("usage") or {}
            entry["model"] = data.get("model")
            entry["stop_reason"] = data.get("stop_reason")
            entry["input_tokens"] = usage.get("input_tokens")
            entry["output_tokens"] = usage.get("output_tokens")
            entry["cache_read_tokens"] = usage.get("cache_read_input_tokens")
            entry["cache_creation_tokens"] = usage.get("cache_creation_input_tokens")
            if status != 200:
                entry["error"] = str(data)[:300]
        except Exception:
            entry["error"] = resp_body[:200].decode("utf-8", "replace")
        with _log_lock:
            RESULTS_DIR.mkdir(exist_ok=True)
            with LOG_FILE.open("a") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")

        self.send_response(status)
        self.send_header("Content-Type", resp_ctype)
        self.send_header("Content-Length", str(len(resp_body)))
        self.end_headers()
        self.wfile.write(resp_body)


def main(port: int = 8791):
    RESULTS_DIR.mkdir(exist_ok=True)
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"metering proxy on http://127.0.0.1:{port} -> {UPSTREAM}", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    import sys

    main(int(sys.argv[1]) if len(sys.argv) > 1 else 8791)
