#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]

DEFAULT_REPO_URL = "https://github.com/Alexander-Aue-Johr/excel-shrink-rust"
DEFAULT_REPO_DIR = "external/excel-shrink-rust"
DEFAULT_CARGO_PACKAGE = "excel-shrink-cli"


def run_command(command: list[str], cwd: Path) -> None:
    print()
    print("+", " ".join(command))
    subprocess.run(command, cwd=str(cwd), check=True)


def resolve_project_path(path_text: str) -> Path:
    path = Path(path_text)
    if path.is_absolute():
        return path
    return PROJECT_ROOT / path


def resolve_rust_arg_paths(rust_args: list[str]) -> list[str]:
    path_options = {"--analysis-csv"}
    resolved_args: list[str] = []
    index = 0

    while index < len(rust_args):
        arg = rust_args[index]
        resolved_args.append(arg)

        if arg in path_options and index + 1 < len(rust_args):
            resolved_args.append(str(resolve_project_path(rust_args[index + 1])))
            index += 2
            continue

        index += 1

    return resolved_args


def ensure_rust_repo(repo_url: str, repo_dir: Path, update: bool) -> None:
    if not repo_dir.exists():
        repo_dir.parent.mkdir(parents=True, exist_ok=True)
        run_command(["git", "clone", repo_url, str(repo_dir)], cwd=PROJECT_ROOT)
        return

    if update:
        run_command(["git", "pull", "--ff-only"], cwd=repo_dir)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Clone/build/run ExcelShrinkrust on a directory of XLSX files."
    )

    parser.add_argument("input_path")
    parser.add_argument("output_path")

    parser.add_argument(
        "--repo-url",
        default=DEFAULT_REPO_URL,
        help="Git URL of the ExcelShrinkrust repository.",
    )
    parser.add_argument(
        "--repo-dir",
        default=DEFAULT_REPO_DIR,
        help="Local checkout directory for the ExcelShrinkrust repository.",
    )
    parser.add_argument(
        "--cargo-package",
        default=DEFAULT_CARGO_PACKAGE,
        help="Cargo package to run. Adjust this if Cargo.toml uses another package name.",
    )
    parser.add_argument(
        "--cargo-bin",
        default=None,
        help="Optional Cargo binary name. Usually not needed.",
    )
    parser.add_argument(
        "--update",
        action="store_true",
        help="Run git pull --ff-only before building/running.",
    )
    parser.add_argument(
        "--debug-build",
        action="store_true",
        help="Use cargo run without --release.",
    )

    args, rust_args = parser.parse_known_args()
    rust_args = resolve_rust_arg_paths(rust_args)

    repo_dir = resolve_project_path(args.repo_dir)
    input_path = resolve_project_path(args.input_path)
    output_path = resolve_project_path(args.output_path)

    ensure_rust_repo(args.repo_url, repo_dir, args.update)

    output_path.mkdir(parents=True, exist_ok=True)

    cargo_command = ["cargo", "run"]

    if not args.debug_build:
        cargo_command.append("--release")

    if args.cargo_package:
        cargo_command.extend(["--package", args.cargo_package])

    if args.cargo_bin:
        cargo_command.extend(["--bin", args.cargo_bin])

    cargo_command.extend(
        [
            "--",
            str(input_path),
            str(output_path),
            *rust_args,
        ]
    )

    run_command(cargo_command, cwd=repo_dir)


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as error:
        sys.exit(error.returncode)
