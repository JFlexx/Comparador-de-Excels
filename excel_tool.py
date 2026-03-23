from __future__ import annotations

import argparse
from pathlib import Path

from comparator import VALID_ACTIONS, VALID_COMPARE_MODES
from interface_adapter import (
    build_compare_options,
    compare_files,
    default_action_for_base,
    load_decisions,
    merge_labels,
    merge_workbooks,
    parse_sheet_keys_args,
    export_template,
)


def cmd_compare(args: argparse.Namespace) -> int:
    options = build_compare_options(
        compare_mode=args.compare_mode,
        ignore_case=args.ignore_case,
        keep_spaces=args.keep_spaces,
        empty_string_is_value=args.empty_string_is_value,
        header_row=args.header_row,
        sheet_keys=parse_sheet_keys_args(args.sheet_key),
    )
    diff = compare_files(args.a, args.b, options)

    template_path = Path(args.template)
    default_action = args.default_action or default_action_for_base(args.base)
    export_template(diff, template_path, base=args.base, default_action=default_action)

    labels = merge_labels(args.base)
    print("Comparación completada.")
    print(f"- Modo: {args.compare_mode}")
    print(f"- Hojas en común: {len(diff.common_sheets)}")
    print(f"- Hojas solo en A: {len(diff.only_in_a)}")
    print(f"- Hojas solo en B: {len(diff.only_in_b)}")
    print(f"- Diferencias: {diff.total_differences}")
    print(f"- Merge objetivo: traer cambios de {labels.source} hacia {labels.base}")
    print(f"- Acción por defecto: {default_action}")
    if args.compare_mode == "row-based":
        print("- Merge row-based habilitado: la plantilla conserva metadata para aplicar decisiones por clave.")
    print(f"Plantilla de decisiones creada: {template_path}")
    return 0


def cmd_merge(args: argparse.Namespace) -> int:
    decisions = load_decisions(args.decisions)
    output_path = Path(args.output)

    merge_workbooks(
        workbook_a=args.a,
        workbook_b=args.b,
        decisions=decisions,
        output_path=output_path,
        base=args.apply_onto,
        include_sheets_from_source_only=not args.no_copy_extra_sheets,
    )

    labels = merge_labels(args.apply_onto)
    print(f"Libro combinado generado: {output_path}")
    print(f"- Base destino: {labels.base}")
    print(f"- Se aplicaron decisiones para traer cambios de {labels.source}")
    if decisions.attrs.get("compare_mode") == "row-based":
        print("- El merge se resolvió por identidad lógica del registro (key/header/diff_type).")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel_tool",
        description="Comparador de libros Excel con flujo de decisiones editable en Excel",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    compare = sub.add_parser("compare", help="Compara A vs B y genera plantilla de decisiones")
    compare.add_argument(
        "--base",
        choices=["a", "b"],
        default="a",
        help="Libro que se considera base/destino al revisar decisiones (por defecto: A)",
    )
    compare.add_argument("--a", required=True, help="Ruta de Excel A")
    compare.add_argument("--b", required=True, help="Ruta de Excel B")
    compare.add_argument(
        "--template",
        default="decisiones.xlsx",
        help="Ruta salida de plantilla de decisiones",
    )
    compare.add_argument(
        "--default-action",
        choices=sorted(VALID_ACTIONS),
        default=None,
        help="Acción por defecto para cada diferencia; si se omite, se usa la del libro origen según --base",
    )
    compare.add_argument(
        "--compare-mode",
        choices=sorted(VALID_COMPARE_MODES),
        default="coordinate",
        help="coordinate para comparar celda a celda; row-based para comparar registros por encabezados/clave",
    )
    compare.add_argument(
        "--sheet-key",
        action="append",
        help="Clave por hoja con formato Hoja=columna1,columna2. Repetible.",
    )
    compare.add_argument(
        "--header-row",
        type=int,
        default=1,
        help="Fila que contiene los encabezados para el modo row-based",
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
    merge.add_argument("--a", required=True, help="Ruta de Excel A")
    merge.add_argument("--b", required=True, help="Ruta de Excel B")
    merge.add_argument("--decisions", required=True, help="Excel de decisiones editado")
    merge.add_argument(
        "--apply-onto",
        choices=["a", "b"],
        default="a",
        help="Libro destino sobre el que se aplican las decisiones (por defecto: A)",
    )
    merge.add_argument("--output", default="resultado_combinado.xlsx", help="Ruta de salida del resultado")
    merge.add_argument(
        "--no-copy-extra-sheets",
        action="store_true",
        help="No copiar hojas que existen solo en el libro origen",
    )
    merge.set_defaults(func=cmd_merge)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
