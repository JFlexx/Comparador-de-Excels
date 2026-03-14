from __future__ import annotations

import argparse
from pathlib import Path

from comparator import (
    CompareOptions,
    apply_decisions,
    compare_workbooks,
    decisions_from_excel,
    export_decision_template,
)


def cmd_compare(args: argparse.Namespace) -> int:
    options = CompareOptions(
        strip_strings=not args.keep_spaces,
        case_sensitive=not args.ignore_case,
        ignore_empty_string_vs_none=not args.empty_string_is_value,
    )
    diff = compare_workbooks(args.a, args.b, options=options)

    template_path = Path(args.template)
    export_decision_template(diff, template_path, default_action=args.default_action)

    print("Comparación completada.")
    print(f"- Hojas en común: {len(diff.common_sheets)}")
    print(f"- Hojas solo en A: {len(diff.only_in_a)}")
    print(f"- Hojas solo en B: {len(diff.only_in_b)}")
    print(f"- Diferencias: {len(diff.all_differences())}")
    print(f"Plantilla de decisiones creada: {template_path}")
    return 0


def cmd_merge(args: argparse.Namespace) -> int:
    decisions = decisions_from_excel(args.decisions)
    output_path = Path(args.output)

    apply_decisions(
        base_workbook=args.a,
        decisions=decisions,
        output_path=output_path,
        workbook_b=args.b,
        include_sheets_only_in_b=not args.no_copy_extra_sheets,
    )

    print(f"Libro combinado generado: {output_path}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel_tool",
        description="Comparador de libros Excel con flujo de decisiones editable en Excel",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    compare = sub.add_parser("compare", help="Compara A vs B y genera plantilla de decisiones")
    compare.add_argument("--a", required=True, help="Ruta de Excel A (base)")
    compare.add_argument("--b", required=True, help="Ruta de Excel B (comparar)")
    compare.add_argument(
        "--template",
        default="decisiones.xlsx",
        help="Ruta salida de plantilla de decisiones",
    )
    compare.add_argument(
        "--default-action",
        choices=["use_a", "use_b", "manual"],
        default="use_b",
        help="Acción por defecto para cada diferencia",
    )
    compare.add_argument("--ignore-case", action="store_true", help="No distinguir mayúsculas/minúsculas")
    compare.add_argument("--keep-spaces", action="store_true", help="No recortar espacios en strings")
    compare.add_argument(
        "--empty-string-is-value",
        action="store_true",
        help="Tratar '' distinto de None",
    )
    compare.set_defaults(func=cmd_compare)

    merge = sub.add_parser("merge", help="Aplica decisiones y genera el libro final")
    merge.add_argument("--a", required=True, help="Ruta de Excel A (base)")
    merge.add_argument("--b", required=True, help="Ruta de Excel B (comparar)")
    merge.add_argument("--decisions", required=True, help="Excel de decisiones editado")
    merge.add_argument("--output", default="resultado_combinado.xlsx", help="Ruta de salida del resultado")
    merge.add_argument(
        "--no-copy-extra-sheets",
        action="store_true",
        help="No copiar hojas que existen solo en B",
    )
    merge.set_defaults(func=cmd_merge)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
