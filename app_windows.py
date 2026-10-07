from __future__ import annotations

import argparse
import subprocess
from datetime import datetime
from pathlib import Path

from converter import convert_c6_refin


APP_TITLE = "BEVI para 2TECH - C6"


def run_powershell(script: str) -> str:
    result = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "Falha ao abrir diálogo do Windows.")
    return result.stdout.strip()


def choose_input_file() -> Path | None:
    script = rf"""
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = '{APP_TITLE}: selecione a tabela BEVI'
$dialog.Filter = 'Planilhas BEVI (*.xls;*.xlsx;*.html;*.htm)|*.xls;*.xlsx;*.html;*.htm|Todos os arquivos (*.*)|*.*'
$dialog.Multiselect = $false
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  Write-Output $dialog.FileName
}}
"""
    selected = run_powershell(script)
    return Path(selected) if selected else None


def choose_output_folder() -> Path | None:
    script = rf"""
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = '{APP_TITLE}: escolha a pasta de saída'
$dialog.ShowNewFolderButton = $true
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  Write-Output $dialog.SelectedPath
}}
"""
    selected = run_powershell(script)
    return Path(selected) if selected else None


def ask_start_date() -> str | None:
    script = rf"""
Add-Type -AssemblyName Microsoft.VisualBasic
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$value = [Microsoft.VisualBasic.Interaction]::InputBox(
  'Início da vigência em dd/mm/aaaa. Deixe em branco para usar a data da BEVI.',
  '{APP_TITLE}',
  ''
)
Write-Output $value
"""
    value = run_powershell(script).strip()
    return value or None


def show_message(message: str, kind: str = "info") -> None:
    icon = "Error" if kind == "error" else "Information"
    safe_message = message.replace("'", "''")
    script = rf"""
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show('{safe_message}', '{APP_TITLE}', 'OK', '{icon}') | Out-Null
"""
    run_powershell(script)


def main() -> int:
    try:
        args = parse_args()
        if args.input and args.output:
            result = convert_c6_refin(args.input, args.output, start_date=args.inicio)
            return 0

        input_path = choose_input_file()
        if input_path is None:
            return 0

        output_folder = choose_output_folder()
        if output_folder is None:
            return 0

        start_date = ask_start_date()
        if start_date:
            datetime.strptime(start_date, "%d/%m/%Y")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_folder / f"C6_2TECH_padronizado_{timestamp}.xlsx"
        result = convert_c6_refin(input_path, output_path, start_date=start_date)
    except Exception as exc:
        show_message(f"Não foi possível gerar o arquivo:\n\n{exc}", kind="error")
        return 1

    show_message(f"Arquivo gerado com sucesso.\n\nLinhas: {result.row_count}\nArquivo: {result.output_path}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--input")
    parser.add_argument("--output")
    parser.add_argument("--inicio")
    return parser.parse_known_args()[0]


if __name__ == "__main__":
    raise SystemExit(main())
