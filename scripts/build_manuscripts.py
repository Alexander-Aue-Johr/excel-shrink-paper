"""Build the current manuscripts with stable ICDE 2027 upload filenames."""
import argparse
import os
from pathlib import Path
import shutil
import subprocess
import tempfile


ROOT = Path(__file__).resolve().parents[1]
MANUSCRIPTS = {
    "sheetshrink": "SheetShrink-ICDE2027",
    "deltasieve": "DeltaSieve-ICDE2027",
}


def find_tex_bin():
    executable = shutil.which("pdflatex") or shutil.which("pdftex")
    if executable:
        return Path(executable).parent
    installations = sorted(Path("C:/texlive").glob("*/bin/windows"), reverse=True)
    for directory in installations:
        if (directory / "pdftex.exe").is_file():
            return directory
    raise RuntimeError("Install TeX Live or put pdflatex and bibtex on PATH.")


def run_tex(command, directory, environment):
    result = subprocess.run(command, cwd=directory, env=environment,
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            text=True, encoding="utf-8", errors="replace")
    if result.returncode:
        raise RuntimeError("Build failed.\n" + result.stdout[-8000:])


def build_manuscript(paper, tex_bin, build_root):
    source = ROOT / "papers" / paper
    output = build_root / paper
    output.mkdir(parents=True, exist_ok=True)
    temporary = build_root / "tmp"
    temporary.mkdir(exist_ok=True)
    environment = os.environ.copy()
    environment.update(TEMP=str(temporary), TMP=str(temporary), TMPDIR=str(temporary))
    environment["TEXINPUTS"] = os.pathsep.join([".", source.as_posix() + "//", ""])
    environment["BIBINPUTS"] = os.pathsep.join([source.as_posix(), ""])
    environment["BSTINPUTS"] = os.pathsep.join([source.as_posix(), ROOT.as_posix(), ""])
    suffix = ".exe" if os.name == "nt" else ""
    latex = tex_bin / ("pdflatex" + suffix)
    command = [str(latex)]
    if not latex.is_file():
        command = [str(tex_bin / ("pdftex" + suffix)), "-progname=pdflatex", "-fmt=pdflatex"]
    job_name = MANUSCRIPTS[paper]
    command += ["-interaction=nonstopmode", "-halt-on-error", "-file-line-error",
                "-jobname=" + job_name, "-output-directory=" + output.as_posix(), "main.tex"]
    run_tex(command, source, environment)
    run_tex([str(tex_bin / ("bibtex" + suffix)), job_name], output, environment)
    run_tex(command, source, environment)
    run_tex(command, source, environment)
    compiled = output / (job_name + ".pdf")
    for filename in (job_name + ".pdf", "main.pdf"):
        destination = source / filename
        pending = destination.with_suffix(".pdf.pending")
        shutil.copyfile(compiled, pending)
        pending.replace(destination)
    print(source / (job_name + ".pdf"))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--paper", choices=["all", *MANUSCRIPTS], default="all")
    parser.add_argument("--tex-bin", type=Path, help="Directory containing the TeX executables")
    parser.add_argument("--build-root", type=Path, help="External directory for temporary and build files")
    args = parser.parse_args()
    tex_bin = args.tex_bin or find_tex_bin()
    default_root = Path("C:/texlive/builds/icde2027") if Path("C:/texlive").is_dir() else Path(tempfile.gettempdir()) / "icde2027-manuscripts"
    build_root = (args.build_root or default_root).resolve()
    selected = MANUSCRIPTS if args.paper == "all" else [args.paper]
    for paper in selected:
        build_manuscript(paper, tex_bin, build_root)


if __name__ == "__main__":
    main()
