from __future__ import annotations

import argparse
from pathlib import Path

from adapter_models import CliCompareReport, CliMergeReport
from excel_adapter import export_decisions_workbook, import_decisions_workbook, merge_from_decisions_workbook
from interface_adapter import build_compare_options, compare_files, describe_merge, parse_sheet_keys


def parse_sheet_keys_args(values: list[str] | None) -> dict[str, list[str]]:
    return parse_sheet_keys(values or [], item_separator="=")


def run_compare_command(args: argparse.Namespace) -> CliCompareReport:
    options = build_compare_options(
        compare_mode=args.compare_mode,
        ignore_case=args.ignore_case,
        keep_spaces=args.keep_spaces,
        empty_string_is_value=args.empty_string_is_value,
        header_row=args.header_row,
        sheet_keys=parse_sheet_keys_args(args.sheet_key),
    )
    diff = compare_files(args.a, args.b, options)
    merge = describe_merge(args.base)
    template_path = Path(args.template)
    default_action = args.default_action or merge.default_action
    export_decisions_workbook(diff, template_path, default_action=default_action)
    return CliCompareReport(
        template_path=template_path,
        compare_mode=args.compare_mode,
        common_sheets=len(diff.common_sheets),
        only_in_a=len(diff.only_in_a),
        only_in_b=len(diff.only_in_b),
        total_differences=diff.total_differences,
        default_action=default_action,
        merge=merge,
    )


def run_merge_command(args: argparse.Namespace) -> CliMergeReport:
    decisions = import_decisions_workbook(args.decisions)
    output_path = Path(args.output)
    merge = describe_merge(args.apply_onto)
    merge_from_decisions_workbook(
        workbook_a=args.a,
        workbook_b=args.b,
        decisions=decisions,
        output_path=output_path,
        base=args.apply_onto,
        include_sheets_from_source_only=not args.no_copy_extra_sheets,
    )
    return CliMergeReport(output_path=output_path, merge=merge)
