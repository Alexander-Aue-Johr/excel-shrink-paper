"""Export primary Check Performance workflow pairs from the frozen local receipts."""
import csv
import gzip
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / "papers/sheetshrink/evidence"
SOURCE = EVIDENCE / "check-performance-figure10"


def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        return list(csv.DictReader(stream))


def main():
    rows = read_csv(SOURCE / "phase-comparison.csv")
    nx = read_csv(EVIDENCE / "top49-performance-pairs.csv")
    xl = read_csv(EVIDENCE / "top49-xlstyles-pairs.csv")
    cohort = {r["input_sha256"] for r in nx} & {r["input_sha256"] for r in xl}
    baselines = {r["input_sha256"]: r for r in nx}
    audit = json.loads((SOURCE / "sheetshrink-baseline-audit.json").read_text())
    assert len(rows) == len(cohort) == 110
    assert {r["sha256"] for r in rows} == cohort
    verification = json.loads((SOURCE / "verification.json").read_text())
    assert verification["sqlite_integrity"] == "ok"
    assert not verification["original_mismatches"] and not verification["output_mismatches"]
    pairs = []
    receipts = []
    for row in rows:
        baseline = baselines[row["sha256"]]
        assert audit[row["sha256"]]["median_matches"]
        assert abs(float(row["historical_sheetshrink_median_seconds"]) - float(baseline["sheetshrink_seconds"])) < 1e-8
        status = row["status"]
        receipts.append({
            "corpus": row["corpus"], "input_sha": row["sha256"],
            "conversion": {"successful": "converted", "partial": "partial", "failed": "failed"}[status],
            "eligible_for_formatting_comparison": status == "successful",
        })
        if status != "successful":
            continue
        assert row["sampled_process_matches_native_receipt"] == "True"
        assert row["cpu_covers_exit"] == "True"
        pairs.append({
            "host_id": "LAPTOP-H86KBIH3", "comparator": "Check Performance",
            "corpus": row["corpus"], "input_sha256": row["sha256"],
            "comparator_seconds": row["workflow_seconds"],
            "sheetshrink_seconds": baseline["sheetshrink_seconds"],
            "sheetshrink_repetitions": 9, "comparator_repetitions": 1,
            "comparator_endpoint": "Native wrapper startup through exit including control saves, optimization and artifact creation",
            "sheetshrink_endpoint": "Historical fresh CLI process creation through exit",
        })
    assert len(pairs) == 102
    with (EVIDENCE / "top49-check-performance-pairs.csv").open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(pairs[0]))
        writer.writeheader()
        writer.writerows(pairs)
    receipt = {"host_id": "LAPTOP-H86KBIH3", "tool_label": "Check Performance", "records": receipts}
    payload = json.dumps(receipt, indent=2).encode("utf-8")
    (EVIDENCE / "top49-check-performance-receipts.json.gz").write_bytes(gzip.compress(payload, mtime=0))
    print("Verified 110 primary outcomes and exported 102 Check Performance pairs")


if __name__ == "__main__":
    main()
