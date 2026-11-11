"""Join the frozen Inquire cohort to verified same-host Sheet Shrink trials."""
import csv
import gzip
import json
import statistics
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / "papers/sheetshrink/evidence"
SOURCE = EVIDENCE / "inquire-figure10"


def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        return list(csv.DictReader(stream))


def main():
    comparison = read_csv(SOURCE / "comparison.csv")
    summaries = {row["id"]: row for row in read_csv(SOURCE / "sheetshrink-summary.csv")}
    with gzip.open(SOURCE / "trials.csv.gz", "rt", encoding="utf-8-sig", newline="") as stream:
        trials = list(csv.DictReader(stream))
    expected = {r["input_sha256"] for r in read_csv(EVIDENCE / "top49-performance-pairs.csv")} & {
        r["input_sha256"] for r in read_csv(EVIDENCE / "top49-xlstyles-pairs.csv")}
    assert len(comparison) == len(expected) == 110
    assert {r["input_sha256"] for r in comparison} == expected
    pairs = []
    receipts = []
    for row in comparison:
        summary = summaries[row["id"]]
        measured = [t for t in trials if t["case_id"] == row["id"] and t["warmup"] == "0"]
        assert len(measured) == int(summary["measured_repetitions"])
        assert all(t["status"] == "success" for t in measured)
        median = statistics.median(float(t["elapsed_seconds"]) for t in measured)
        assert abs(median - float(row["sheetshrink_median_seconds"])) < 1e-8
        assert len({t["content_sha256"] for t in measured}) == 1
        assert summary["host"] == "LV-ANB-001P-BS"
        status = row["inquire_primary_status"]
        receipts.append({
            "corpus": row["corpus"], "input_sha": row["input_sha256"],
            "conversion": {"success": "converted", "success_with_warnings": "partial", "failed": "failed"}[status],
            "eligible_for_formatting_comparison": status == "success",
        })
        if status != "success":
            continue
        pairs.append({
            "host_id": "LV-ANB-001P-BS", "comparator": "Inquire",
            "corpus": row["corpus"], "input_sha256": row["input_sha256"],
            "comparator_seconds": row["inquire_single_observation_seconds"],
            "sheetshrink_seconds": median,
            "sheetshrink_repetitions": summary["measured_repetitions"],
            "comparator_repetitions": 1,
            "historical_storage": row["inquire_historical_storage"],
            "comparator_endpoint": "Historical Excel workflow including opening, cleanup, saves and shutdown; per-case boundaries retained",
            "sheetshrink_endpoint": "Fresh CLI process creation through exit on local NVMe",
        })
    assert len(pairs) == 90
    with (EVIDENCE / "top49-inquire-pairs.csv").open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(pairs[0]))
        writer.writeheader()
        writer.writerows(pairs)
    receipt = {"host_id": "LV-ANB-001P-BS", "tool_label": "Inquire", "records": receipts}
    payload = json.dumps(receipt, indent=2).encode("utf-8")
    (EVIDENCE / "top49-inquire-receipts.json.gz").write_bytes(gzip.compress(payload, mtime=0))
    print("Verified 110 same-host baselines and exported 90 complete Inquire pairs")


if __name__ == "__main__":
    main()
