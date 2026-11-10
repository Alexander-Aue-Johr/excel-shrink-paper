"""Run the frozen FUSE workflow without importing the active research checkout."""
import argparse
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from excel_shrink_research.fuse_workflow import run
from excel_shrink_research.enron.common import write_json, utc_now

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True)
    args = parser.parse_args()
    try:
        run(args.output)
    except Exception as exc:
        write_json(Path(args.output) / 'progress.json', {'state': 'failed', 'error': str(exc), 'updated_utc': utc_now()})
        raise
