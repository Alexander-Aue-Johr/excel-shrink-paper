#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional


def number(value: Any) -> Optional[float]:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def nested(data: Dict[str, Any], *keys: str) -> Any:
    current: Any = data
    for key in keys:
        if not isinstance(current, dict):
            return None
        current = current.get(key)
    return current


def rows_from_report(report: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for item in report.get("performance_files") or []:
        timings = item.get("timings_sec") or {}
        memory = item.get("memory") or {}
        rows.append(
            {
                "filename": item.get("filename"),
                "status": item.get("status"),
                "open_original_sec": number(timings.get("open_original_a")),
                "open_shrunk_sec": number(timings.get("open_shrunk_b")),
                "recalc_original_sec": number(timings.get("recalc_original_a")),
                "recalc_shrunk_sec": number(timings.get("recalc_shrunk_b")),
                "original_working_set_after_recalc_mb": number(
                    nested(
                        memory,
                        "original_app_after_recalc_original_a",
                        "working_set_mb",
                    )
                ),
                "shrunk_working_set_after_recalc_mb": number(
                    nested(
                        memory,
                        "comparison_app_after_recalc_shrunk_b",
                        "working_set_mb",
                    )
                ),
                "original_private_after_recalc_mb": number(
                    nested(
                        memory,
                        "original_app_after_recalc_original_a",
                        "private_usage_mb",
                    )
                ),
                "shrunk_private_after_recalc_mb": number(
                    nested(
                        memory,
                        "comparison_app_after_recalc_shrunk_b",
                        "private_usage_mb",
                    )
                ),
            }
        )
    return rows


def write_csv(rows: List[Dict[str, Any]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "filename",
        "status",
        "open_original_sec",
        "open_shrunk_sec",
        "recalc_original_sec",
        "recalc_shrunk_sec",
        "original_working_set_after_recalc_mb",
        "shrunk_working_set_after_recalc_mb",
        "original_private_after_recalc_mb",
        "shrunk_private_after_recalc_mb",
    ]
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def values(rows: Iterable[Dict[str, Any]], key: str) -> List[float]:
    out: List[float] = []
    for row in rows:
        value = number(row.get(key))
        if value is not None:
            out.append(value)
    return out


def make_boxplots(rows: List[Dict[str, Any]], out_dir: Path, stem: str) -> None:
    try:
        import matplotlib.pyplot as plt
    except ImportError as e:
        raise SystemExit("matplotlib is required for plots: pip install matplotlib") from e

    out_dir.mkdir(parents=True, exist_ok=True)

    plot_specs = [
        (
            "open_times_sec",
            "Excel Open Time",
            "seconds",
            [
                ("Original", "open_original_sec"),
                ("Shrunk", "open_shrunk_sec"),
            ],
        ),
        (
            "recalc_times_sec",
            "Excel Recalc Time",
            "seconds",
            [
                ("Original", "recalc_original_sec"),
                ("Shrunk", "recalc_shrunk_sec"),
            ],
        ),
        (
            "working_set_after_recalc_mb",
            "Excel Working Set After Recalc",
            "MB",
            [
                ("Original", "original_working_set_after_recalc_mb"),
                ("Shrunk", "shrunk_working_set_after_recalc_mb"),
            ],
        ),
        (
            "private_memory_after_recalc_mb",
            "Excel Private Memory After Recalc",
            "MB",
            [
                ("Original", "original_private_after_recalc_mb"),
                ("Shrunk", "shrunk_private_after_recalc_mb"),
            ],
        ),
    ]

    for suffix, title, ylabel, series in plot_specs:
        labels = [label for label, _key in series]
        data = [values(rows, key) for _label, key in series]
        if not any(data):
            continue

        fig, ax = plt.subplots(figsize=(6.0, 4.0))
        ax.boxplot(data, labels=labels, showfliers=False)
        ax.set_title(title)
        ax.set_ylabel(ylabel)
        ax.grid(axis="y", alpha=0.25)
        fig.tight_layout()

        png = out_dir / f"{stem}_{suffix}.png"
        pdf = out_dir / f"{stem}_{suffix}.pdf"
        fig.savefig(png, dpi=200)
        fig.savefig(pdf)
        plt.close(fig)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate CSV and boxplots from verify report performance_files."
    )
    parser.add_argument("report", type=Path)
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="Output directory. Defaults to <report parent>/performance_plots_<stem>.",
    )
    args = parser.parse_args()

    with args.report.open("r", encoding="utf-8") as f:
        report = json.load(f)

    rows = rows_from_report(report)
    out_dir = args.out_dir or args.report.with_name(
        f"performance_plots_{args.report.stem}"
    )
    csv_path = out_dir / f"{args.report.stem}_performance.csv"
    write_csv(rows, csv_path)
    make_boxplots(rows, out_dir, args.report.stem)

    print(f"Wrote {len(rows)} rows to {csv_path}")
    print(f"Wrote plots to {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
