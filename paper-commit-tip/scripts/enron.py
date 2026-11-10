"""Run from the checkout without installing the paper's plotting dependencies."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from excel_shrink_research.enron.cli import main

if __name__ == "__main__":
    sys.exit(main())
