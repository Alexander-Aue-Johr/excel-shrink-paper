"""Export the method diagrams as standalone vector PDFs from their TikZ sources."""

import argparse
import os
from pathlib import Path
import re
import shutil
import subprocess


REPOSITORY = Path(__file__).resolve().parents[2]
FIGURES = (
    "papers/sheetshrink/figures/scheduling.tex",
    "papers/deltasieve/figures/artifact-lineage.tex",
    "papers/deltasieve/figures/capture-partition.tex",
    "papers/deltasieve/figures/communication.tex",
    "papers/deltasieve/figures/capture-ownership.tex",
    "papers/deltasieve/figures/run-states.tex",
)


def standalone_document(source: Path) -> str:
    """Keep the diagram's native dimensions and add a three-millimetre margin."""
    text = source.read_text(encoding="utf-8")
    start = text.index(r"\begin{tikzpicture}")
    end = text.index(r"\end{tikzpicture}") + len(r"\end{tikzpicture}")
    diagram = text[start:end]
    bounds = re.search(
        r"\\path\[use as bounding box\]\s*"
        r"\(([-.\d]+),([-.\d]+)\)\s*rectangle\s*"
        r"\(([-.\d]+),([-.\d]+)\)",
        diagram,
    )
    if bounds is None:
        raise ValueError(f"Missing explicit diagram bounds in {source}")
    left, top, right, bottom = map(float, bounds.groups())
    width = right - left + 0.6
    height = top - bottom + 0.6
    style = (REPOSITORY / "papers/shared/state-diagram-style.tex").as_posix()
    return "\n".join(
        [
            r"\documentclass[10pt]{article}",
            r"\usepackage[T1]{fontenc}",
            r"\renewcommand{\rmdefault}{ptm}",
            rf"\setlength{{\paperwidth}}{{{width}cm}}",
            rf"\setlength{{\paperheight}}{{{height}cm}}",
            r"\pdfpagewidth=\paperwidth",
            r"\pdfpageheight=\paperheight",
            r"\setlength{\textwidth}{\dimexpr\paperwidth-6mm\relax}",
            r"\setlength{\textheight}{\dimexpr\paperheight-3mm\relax}",
            r"\setlength{\oddsidemargin}{\dimexpr-1in+3mm\relax}",
            r"\setlength{\topmargin}{\dimexpr-1in+3mm\relax}",
            r"\setlength{\headheight}{0pt}",
            r"\setlength{\headsep}{0pt}",
            r"\usepackage{tikz}",
            r"\usetikzlibrary{arrows.meta}",
            rf"\input{{{style}}}",
            r"\pagestyle{empty}",
            r"\setlength{\parindent}{0pt}",
            r"\setlength{\topskip}{0pt}",
            r"\begin{document}",
            r"\noindent" + diagram,
            r"\end{document}",
        ]
    )


def export_figure(source: Path, texlive: Path, scratch: Path) -> None:
    """Compile one owned wrapper and publish its PDF only after TeX succeeds."""
    wrapper = scratch / source.name
    wrapper.write_text(standalone_document(source), encoding="utf-8")
    environment = os.environ.copy()
    for variable in ("TEMP", "TMP", "TMPDIR"):
        environment[variable] = str(scratch)
    for variable, directory in (
        ("TEXMFHOME", "texmf-local"),
        ("TEXMFVAR", "texmf-var"),
        ("TEXMFCONFIG", "texmf-config"),
        ("TEXMFCACHE", "texmf-var"),
    ):
        environment[variable] = str(texlive / directory)
    command = [
        str(texlive / "bin/windows/pdflatex.exe"),
        "-interaction=nonstopmode",
        "-halt-on-error",
        "-no-shell-escape",
        f"-output-directory={scratch}",
        str(wrapper),
    ]
    result = subprocess.run(
        command, cwd=scratch, env=environment, capture_output=True, text=True
    )
    if result.returncode:
        raise RuntimeError(f"Diagram export failed for {source}\n{result.stdout}\n{result.stderr}")
    destination = source.with_suffix(".pdf")
    shutil.copy2(wrapper.with_suffix(".pdf"), destination)
    print(f"Exported {destination}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--texlive", type=Path, required=True)
    parser.add_argument(
        "--figure", type=Path, action="append",
        help="Export only these figure sources instead of the full set",
    )
    arguments = parser.parse_args()
    texlive = arguments.texlive.resolve()
    scratch = texlive / "scratch/diagram-design"
    scratch.mkdir(parents=True, exist_ok=True)
    sources = arguments.figure or [REPOSITORY / relative for relative in FIGURES]
    for source in sources:
        export_figure(source.resolve(), texlive, scratch)


if __name__ == "__main__":
    main()
