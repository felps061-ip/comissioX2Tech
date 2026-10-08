from __future__ import annotations

import argparse
import cgi
import io
import os
import socket
import tempfile
import threading
import time
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from converter import convert_c6_refin


APP_TITLE = "BEVI para 2TECH"
HOST = "127.0.0.1"


HTML = r"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>BEVI para 2TECH</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #eef3f8;
      --ink: #172033;
      --muted: #657386;
      --line: #d8e0ea;
      --panel: #ffffff;
      --blue: #1f5eff;
      --blue-dark: #1547c4;
      --green: #0f8a5f;
      --red: #bd2d38;
      --soft-blue: #e8efff;
      --soft-green: #e8f6ef;
      --shadow: 0 18px 45px rgba(23, 32, 51, 0.12);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: "Segoe UI", Arial, sans-serif;
      color: var(--ink);
      background:
        linear-gradient(135deg, rgba(31, 94, 255, 0.10), rgba(15, 138, 95, 0.08)),
        var(--bg);
    }
    .shell {
      width: min(1040px, calc(100vw - 40px));
      margin: 0 auto;
      padding: 34px 0;
    }
    .topbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 18px;
      margin-bottom: 24px;
    }
    .brand {
      display: flex;
      align-items: center;
      gap: 14px;
    }
    .mark {
      width: 44px;
      height: 44px;
      border-radius: 8px;
      display: grid;
      place-items: center;
      color: #fff;
      font-weight: 800;
      background: linear-gradient(135deg, var(--blue), var(--green));
      box-shadow: 0 12px 24px rgba(31, 94, 255, 0.24);
    }
    h1 {
      margin: 0;
      font-size: 26px;
      line-height: 1.15;
      letter-spacing: 0;
    }
    .subtitle {
      margin-top: 4px;
      color: var(--muted);
      font-size: 14px;
    }
    .pill {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 8px 12px;
      border-radius: 999px;
      color: #17442f;
      background: var(--soft-green);
      font-size: 13px;
      font-weight: 600;
      white-space: nowrap;
    }
    .grid {
      display: grid;
      grid-template-columns: 1.2fr 0.8fr;
      gap: 20px;
      align-items: start;
    }
    .panel {
      background: var(--panel);
      border: 1px solid rgba(216, 224, 234, 0.92);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }
    .main-card { padding: 26px; }
    .side-card { padding: 22px; }
    .step {
      display: flex;
      gap: 12px;
      align-items: flex-start;
      padding: 14px 0;
      border-bottom: 1px solid var(--line);
    }
    .step:last-child { border-bottom: 0; }
    .num {
      width: 28px;
      height: 28px;
      flex: 0 0 28px;
      border-radius: 50%;
      display: grid;
      place-items: center;
      color: #fff;
      background: var(--blue);
      font-weight: 700;
      font-size: 13px;
    }
    .step strong { display: block; margin-bottom: 2px; }
    .step span { color: var(--muted); font-size: 13px; line-height: 1.4; }
    label {
      display: block;
      margin: 18px 0 8px;
      font-size: 13px;
      font-weight: 700;
      color: #2c3748;
    }
    .drop {
      position: relative;
      border: 1.5px dashed #a9b8cc;
      border-radius: 8px;
      background: #f9fbfe;
      padding: 28px;
      text-align: center;
      transition: 0.16s ease;
    }
    .drop.dragover {
      border-color: var(--blue);
      background: var(--soft-blue);
    }
    .file-name {
      margin-top: 12px;
      color: var(--muted);
      font-size: 13px;
      word-break: break-word;
    }
    input[type="file"] {
      position: absolute;
      inset: 0;
      opacity: 0;
      cursor: pointer;
    }
    input[type="text"] {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 13px 14px;
      font-size: 15px;
      outline: none;
      background: #fff;
    }
    input[type="text"]:focus {
      border-color: var(--blue);
      box-shadow: 0 0 0 3px rgba(31, 94, 255, 0.14);
    }
    .actions {
      display: flex;
      gap: 12px;
      align-items: center;
      margin-top: 22px;
      flex-wrap: wrap;
    }
    button, .download {
      border: 0;
      border-radius: 8px;
      padding: 13px 18px;
      font-weight: 700;
      font-size: 14px;
      cursor: pointer;
      text-decoration: none;
    }
    .primary {
      color: #fff;
      background: var(--blue);
    }
    .primary:hover { background: var(--blue-dark); }
    .secondary {
      color: #314054;
      background: #e9eef5;
    }
    .download {
      display: none;
      color: #fff;
      background: var(--green);
    }
    button:disabled {
      cursor: not-allowed;
      opacity: 0.62;
    }
    .status {
      margin-top: 18px;
      padding: 14px;
      min-height: 50px;
      border-radius: 8px;
      border: 1px solid var(--line);
      background: #f9fbfe;
      color: var(--muted);
      line-height: 1.4;
      font-size: 14px;
    }
    .status.ok {
      border-color: #b9e3cf;
      color: #0f5138;
      background: var(--soft-green);
    }
    .status.err {
      border-color: #f1bcc2;
      color: var(--red);
      background: #fff1f2;
    }
    .rules h2 {
      margin: 0 0 12px;
      font-size: 17px;
    }
    .rules ul {
      margin: 0;
      padding-left: 20px;
      color: var(--muted);
      line-height: 1.7;
      font-size: 14px;
    }
    .footer {
      margin-top: 18px;
      color: var(--muted);
      font-size: 12px;
      text-align: center;
    }
    @media (max-width: 820px) {
      .grid { grid-template-columns: 1fr; }
      .topbar { align-items: flex-start; flex-direction: column; }
      .shell { width: min(100vw - 24px, 1040px); padding-top: 22px; }
      .main-card, .side-card { padding: 18px; }
    }
  </style>
</head>
<body>
  <main class="shell">
    <header class="topbar">
      <div class="brand">
        <div class="mark">2T</div>
        <div>
          <h1>BEVI para 2TECH</h1>
          <div class="subtitle">Conversor visual para tabelas C6 BANK, DIGIO, DAYCOVAL e SAFRA</div>
        </div>
      </div>
      <div class="pill">C6, Digio, Daycoval e Safra ativos</div>
    </header>

    <section class="grid">
      <div class="panel main-card">
        <form id="form">
          <label>Tabela de comissão BEVI</label>
          <div id="drop" class="drop">
            <strong>Arraste o arquivo aqui ou clique para selecionar</strong>
            <div class="file-name" id="fileName">Aceita .xls, .xlsx, .html e .htm</div>
            <input id="file" name="file" type="file" accept=".xls,.xlsx,.html,.htm" required />
          </div>

          <label for="inicio">Início da vigência</label>
          <input id="inicio" name="inicio" type="text" placeholder="Opcional: dd/mm/aaaa. Para PAN, NEO e FINANTO, informe a vigência." />

          <div class="actions">
            <button id="submit" class="primary" type="submit">Converter para 2TECH</button>
            <button id="shutdown" class="secondary" type="button">Encerrar app</button>
            <a id="download" class="download" href="#">Baixar arquivo gerado</a>
          </div>
          <div id="status" class="status">Pronto para converter. O arquivo final sai no layout de 41 colunas da 2TECH.</div>
        </form>
      </div>

      <aside class="panel side-card rules">
        <h2>O que este app faz</h2>
        <div class="step">
          <div class="num">1</div>
          <div><strong>Lê a tabela BEVI</strong><span>Identifica C6, Digio, Daycoval e Safra nas operações de INSS e SIAPE.</span></div>
        </div>
        <div class="step">
          <div class="num">2</div>
          <div><strong>Padroniza para 2TECH</strong><span>Reconhece C6, Digio, Daycoval e Safra, preenchendo convênio, prazo, contrato, formalização e taxas.</span></div>
        </div>
        <div class="step">
          <div class="num">3</div>
          <div><strong>Gera sem fórmulas</strong><span>Repasse 1 é calculado como 45% e salvo como valor final.</span></div>
        </div>
        <ul>
          <li>WEB PLUS vira FÍSICO.</li>
          <li>WEB sem PLUS vira DIGITAL.</li>
          <li>Prazo único ou faixa como 1 a 108 são aceitos.</li>
          <li>Tickets como maior 20K e 6K a 20K preenchem valores mínimo e máximo.</li>
          <li>Trechos como TX 1,850~1,35 são removidos do nome final.</li>
          <li>O prazo final do nome da tabela é removido.</li>
          <li>Fim da vigência fica vazio.</li>
        </ul>
      </aside>
    </section>
    <div class="footer">Pode fechar esta aba depois de baixar o arquivo.</div>
  </main>

  <script>
    const form = document.getElementById('form');
    const file = document.getElementById('file');
    const fileName = document.getElementById('fileName');
    const drop = document.getElementById('drop');
    const statusBox = document.getElementById('status');
    const submit = document.getElementById('submit');
    const download = document.getElementById('download');
    const shutdown = document.getElementById('shutdown');

    function setStatus(text, type = '') {
      statusBox.className = `status ${type}`;
      statusBox.textContent = text;
    }

    file.addEventListener('change', () => {
      fileName.textContent = file.files[0] ? file.files[0].name : 'Aceita .xls, .xlsx, .html e .htm';
      download.style.display = 'none';
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      drop.addEventListener(eventName, event => {
        event.preventDefault();
        drop.classList.add('dragover');
      });
    });
    ['dragleave', 'drop'].forEach(eventName => {
      drop.addEventListener(eventName, event => {
        event.preventDefault();
        drop.classList.remove('dragover');
      });
    });
    drop.addEventListener('drop', event => {
      if (event.dataTransfer.files.length) {
        file.files = event.dataTransfer.files;
        fileName.textContent = file.files[0].name;
        download.style.display = 'none';
      }
    });

    form.addEventListener('submit', async event => {
      event.preventDefault();
      if (!file.files[0]) {
        setStatus('Selecione a tabela BEVI antes de converter.', 'err');
        return;
      }
      const data = new FormData(form);
      submit.disabled = true;
      download.style.display = 'none';
      setStatus('Convertendo. Aguarde alguns segundos...');
      try {
        const response = await fetch('/convert', { method: 'POST', body: data });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || 'Erro ao converter.');
        download.href = payload.downloadUrl;
        download.download = payload.fileName;
        download.style.display = 'inline-block';
        setStatus(`Arquivo pronto: ${payload.fileName}. Linhas geradas: ${payload.rows}.`, 'ok');
      } catch (error) {
        setStatus(error.message, 'err');
      } finally {
        submit.disabled = false;
      }
    });

    shutdown.addEventListener('click', async () => {
      await fetch('/shutdown', { method: 'POST' }).catch(() => {});
      setStatus('App encerrado. Você já pode fechar esta aba.', 'ok');
    });
  </script>
</body>
</html>
"""


class AppState:
    def __init__(self) -> None:
        self.temp_dir = Path(tempfile.mkdtemp(prefix="bevi_2tech_"))
        self.last_output: Path | None = None
        self.last_name: str | None = None
        self.server: ThreadingHTTPServer | None = None


class Handler(BaseHTTPRequestHandler):
    state: AppState

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_bytes(HTML.encode("utf-8"), "text/html; charset=utf-8")
            return
        if parsed.path == "/download":
            self._handle_download()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/convert":
            self._handle_convert()
            return
        if parsed.path == "/shutdown":
            self._send_json({"ok": True})
            threading.Thread(target=self._shutdown_soon, daemon=True).start()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def log_message(self, format: str, *args: object) -> None:
        return

    def _handle_convert(self) -> None:
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                    "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
                },
            )
            upload = form["file"] if "file" in form else None
            if upload is None or not getattr(upload, "filename", ""):
                raise ValueError("Selecione a tabela BEVI para converter.")

            start_date = form.getfirst("inicio", "").strip() or None
            if start_date:
                datetime.strptime(start_date, "%d/%m/%Y")

            suffix = Path(upload.filename).suffix or ".xls"
            source_path = self.state.temp_dir / f"entrada{suffix}"
            with source_path.open("wb") as file_out:
                file_out.write(upload.file.read())

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"C6_2TECH_padronizado_{timestamp}.xlsx"
            output_path = self.state.temp_dir / output_name
            result = convert_c6_refin(source_path, output_path, start_date=start_date)
            self.state.last_output = result.output_path
            self.state.last_name = output_name
            self._send_json(
                {
                    "ok": True,
                    "rows": result.row_count,
                    "fileName": output_name,
                    "downloadUrl": f"/download?name={output_name}",
                }
            )
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc)}, status=HTTPStatus.BAD_REQUEST)

    def _handle_download(self) -> None:
        if not self.state.last_output or not self.state.last_output.exists():
            self.send_error(HTTPStatus.NOT_FOUND, "Nenhum arquivo gerado ainda.")
            return
        data = self.state.last_output.read_bytes()
        file_name = self.state.last_name or "C6_2TECH_padronizado.xlsx"
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{file_name}"')
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, payload: dict[str, object], status: HTTPStatus = HTTPStatus.OK) -> None:
        import json

        self._send_bytes(
            json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            "application/json; charset=utf-8",
            status,
        )

    def _send_bytes(
        self,
        data: bytes,
        content_type: str,
        status: HTTPStatus = HTTPStatus.OK,
    ) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _shutdown_soon(self) -> None:
        time.sleep(0.2)
        if self.state.server:
            self.state.server.shutdown()


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((HOST, 0))
        return int(sock.getsockname()[1])


def run_app(open_browser: bool = True) -> None:
    state = AppState()
    port = find_free_port()

    class BoundHandler(Handler):
        pass

    BoundHandler.state = state
    server = ThreadingHTTPServer((HOST, port), BoundHandler)
    state.server = server
    url = f"http://{HOST}:{port}/"
    if open_browser:
        webbrowser.open(url)
    print(url, flush=True)
    server.serve_forever()
    server.server_close()


def main() -> int:
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    run_app(open_browser=not args.no_browser)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
