from __future__ import annotations

import argparse

from converter import convert_c6_refin


def main() -> None:
    parser = argparse.ArgumentParser(description="Converte tabela BEVI C6 para padrão 2TECH.")
    parser.add_argument("input", help="Arquivo .xls/.html recebido da BEVI")
    parser.add_argument("output", help="Arquivo .xlsx de saída")
    parser.add_argument("--inicio", help="Data de início da vigência em dd/mm/aaaa")
    args = parser.parse_args()

    result = convert_c6_refin(args.input, args.output, start_date=args.inicio)
    print(f"Gerado: {result.output_path} ({result.row_count} linhas)")


if __name__ == "__main__":
    main()
