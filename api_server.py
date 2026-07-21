from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import time
from datetime import datetime
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, Tuple
from urllib.parse import urlparse

from cmis_eeprom_validator import CmisDump, load_expected_parameters, validate_parameters


HOST = "127.0.0.1"
PORT = 8765
ROOT = Path(__file__).resolve().parent


def result_to_dict(result) -> Dict[str, object]:
    parameter = result.parameter
    return {
        "sourceSheet": parameter.source_sheet,
        "check": parameter.check,
        "parameter": parameter.parameter,
        "description": parameter.description,
        "page": result.page,
        "bank": parameter.bank,
        "address": result.address,
        "bits": parameter.bits,
        "length": parameter.length,
        "dataType": parameter.data_type,
        "cmisType": parameter.cmis_type,
        "simpleType": parameter.simple_type,
        "expected": result.expected,
        "actual": result.actual,
        "rawHex": result.raw_hex,
        "status": result.status,
        "message": result.message,
        "activeCheck": parameter.active_check,
        "sourceRow": parameter.source_row,
    }


def summarize(results) -> Dict[str, int]:
    return {
        "total": len(results),
        "passed": sum(1 for result in results if result.status == "PASS"),
        "failed": sum(1 for result in results if result.status == "FAIL"),
        "read": sum(1 for result in results if result.status == "READ"),
        "notChecked": sum(1 for result in results if result.status == "NOT CHECKED"),
        "errors": sum(1 for result in results if result.status == "ERROR"),
    }


def validate_files(dump_path: str, workbook_path: str) -> Dict[str, object]:
    started = datetime.now()
    start_time = time.perf_counter()
    dump = CmisDump.from_file(dump_path)
    return validate_dump(dump, workbook_path, started, start_time)


def validate_dump_text(dump_text: str, workbook_path: str) -> Dict[str, object]:
    started = datetime.now()
    start_time = time.perf_counter()
    dump = CmisDump.from_text(dump_text)
    return validate_dump(dump, workbook_path, started, start_time)


def validate_dump(dump: CmisDump, workbook_path: str, started: datetime, start_time: float) -> Dict[str, object]:
    parameters = load_expected_parameters(workbook_path)
    results = validate_parameters(dump, parameters)
    duration = time.perf_counter() - start_time
    return {
        "results": [result_to_dict(result) for result in results],
        "summary": summarize(results),
        "validationTime": started.strftime("%m/%d/%Y %I:%M:%S %p").replace(" 0", " "),
        "durationSeconds": round(duration, 3),
    }


def sonic_page_command(port: str, page: str) -> str:
    normalized = page.strip().lower()
    if normalized in {"0", "0x0", "0x00", "00", "00h"}:
        return f"sudo sfputil show eeprom-hexdump -p {port}"
    return f"sudo sfputil show eeprom-hexdump -p {port} -n {normalized}"


def read_sonic_eeprom(host: str, username: str, password: str, port: str, pages: str) -> str:
    page_list = [page.strip() for page in pages.replace(",", " ").split() if page.strip()]
    if not page_list:
        page_list = ["0x00", "0x01", "0x02", "0x03", "0x13", "0x2f"]

    commands = [sonic_page_command(port, page) for page in page_list]

    try:
        import paramiko  # type: ignore
    except ImportError:
        return read_sonic_eeprom_with_system_ssh(host, username, password, commands)

    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        client.connect(
            hostname=host,
            username=username,
            password=password or None,
            timeout=15,
            banner_timeout=15,
            auth_timeout=15,
            look_for_keys=not password,
            allow_agent=not password,
        )
        chunks = []
        for command in commands:
            stdin, stdout, stderr = client.exec_command(command, timeout=30)
            output = stdout.read().decode("utf-8", errors="replace")
            error = stderr.read().decode("utf-8", errors="replace")
            status = stdout.channel.recv_exit_status()
            if status != 0:
                raise ValueError(f"Switch command failed: {command}\n{error.strip() or output.strip()}")
            chunks.append(output)
        return "\n".join(chunks)
    finally:
        client.close()


def read_sonic_eeprom_with_system_ssh(host: str, username: str, password: str, commands: list[str]) -> str:
    if password:
        plink = shutil.which("plink") or shutil.which("plink.exe")
        if plink:
            chunks = []
            target = f"{username}@{host}" if username else host
            for command in commands:
                process = subprocess.run(
                    [plink, "-batch", "-ssh", "-pw", password, target, command],
                    capture_output=True,
                    text=True,
                    timeout=45,
                    check=False,
                )
                if process.returncode != 0:
                    raise ValueError(f"Switch command failed: {command}\n{process.stderr.strip() or process.stdout.strip()}")
                chunks.append(process.stdout)
            return "\n".join(chunks)

        raise ValueError(
            "Password SSH needs either the Python 'paramiko' package or plink.exe from PuTTY on PATH. "
            "Install one of those, or configure SSH key access and leave the password blank."
        )

    ssh = shutil.which("ssh")
    if not ssh:
        raise ValueError("Could not find ssh.exe. Install OpenSSH Client or use a backend Python with paramiko.")

    chunks = []
    target = f"{username}@{host}" if username else host
    for command in commands:
        process = subprocess.run(
            [ssh, "-o", "BatchMode=yes", "-o", "StrictHostKeyChecking=accept-new", target, command],
            capture_output=True,
            text=True,
            timeout=45,
            check=False,
        )
        if process.returncode != 0:
            raise ValueError(f"Switch command failed: {command}\n{process.stderr.strip() or process.stdout.strip()}")
        chunks.append(process.stdout)
    return "\n".join(chunks)


def parse_multipart(content_type: str, body: bytes) -> Dict[str, Tuple[str, bytes]]:
    raw_message = (
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
        + body
    )
    message = BytesParser(policy=default).parsebytes(raw_message)
    files: Dict[str, Tuple[str, bytes]] = {}
    if not message.is_multipart():
        return files
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        if not name or not filename:
            continue
        files[name] = (filename, part.get_payload(decode=True) or b"")
    return files


class Handler(BaseHTTPRequestHandler):
    def end_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self.end_headers()

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/api/health":
            self.write_json({"ok": True})
            return
        self.send_error(404, "Not found")

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        try:
            if path == "/api/validate":
                self.handle_validate_upload()
                return
            if path == "/api/validate-sample":
                self.handle_validate_sample()
                return
            if path == "/api/validate-switch":
                self.handle_validate_switch()
                return
            self.send_error(404, "Not found")
        except Exception as exc:
            self.write_json({"error": str(exc)}, status=400)

    def handle_validate_upload(self) -> None:
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(content_length)
        files = parse_multipart(content_type, body)
        if "dump" not in files or "workbook" not in files:
            raise ValueError("Upload both an EEPROM dump file and a workbook/spreadsheet file.")

        with tempfile.TemporaryDirectory(prefix="cmis_validator_") as temp_dir:
            dump_path = self.write_upload(temp_dir, files["dump"])
            workbook_path = self.write_upload(temp_dir, files["workbook"])
            self.write_json(validate_files(dump_path, workbook_path))

    def handle_validate_sample(self) -> None:
        content_length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(content_length)
        payload = json.loads(body.decode("utf-8") or "{}") if body else {}
        sample_name = "sample_wrong_first_6_digits_eeprom_dump.hex" if payload.get("case") == "bad" else "sample_first_6_digits_eeprom_dump.hex"
        dump_path = ROOT / sample_name
        workbook_path = ROOT / "samples" / "sample_expected_values.csv"
        self.write_json(validate_files(str(dump_path), str(workbook_path)))

    def handle_validate_switch(self) -> None:
        content_length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(content_length)
        payload = json.loads(body.decode("utf-8") or "{}") if body else {}
        switch_type = str(payload.get("switchType", "sonic")).strip().lower()
        if switch_type not in {"sonic", "sonic os", "sonic-os"}:
            raise ValueError("Only SONiC switch reading is implemented right now.")

        host = str(payload.get("host", "")).strip()
        username = str(payload.get("username", "")).strip()
        password = str(payload.get("password", ""))
        port = str(payload.get("port", "")).strip()
        pages = str(payload.get("pages", "")).strip()
        workbook_path = str(payload.get("workbookPath", "")).strip()

        if not host or not username or not port:
            raise ValueError("Enter switch host/IP, username, and port.")
        if not workbook_path or not Path(workbook_path).exists():
            raise ValueError("Load a valid expected-values workbook before reading from the switch.")

        dump_text = read_sonic_eeprom(host, username, password, port, pages)
        if not dump_text.strip():
            raise ValueError("Switch EEPROM command returned no output.")
        self.write_json(validate_dump_text(dump_text, workbook_path))

    def write_upload(self, temp_dir: str, upload: Tuple[str, bytes]) -> str:
        filename, content = upload
        suffix = Path(filename).suffix
        safe_name = f"upload{suffix}"
        path = Path(temp_dir) / safe_name
        path.write_bytes(content)
        return str(path)

    def write_json(self, payload: Dict[str, object], status: int = 200) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main() -> int:
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"CMIS validator API running at http://{HOST}:{PORT}")
    server.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
