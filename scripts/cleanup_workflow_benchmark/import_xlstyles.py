"""Retain hash-matched XLStylesTool observations with confirmed primary timing receipts."""
import argparse
import hashlib
import json
from pathlib import Path
import sqlite3
from zipfile import ZipFile


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--database", type=Path, required=True)
    parser.add_argument("--selection", type=Path, required=True)
    parser.add_argument("--host-id", required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    selection = json.loads(args.selection.read_text(encoding="utf-8-sig"))
    with sqlite3.connect(args.database.resolve().as_uri() + "?mode=ro", uri=True) as db:
        db.row_factory = sqlite3.Row
        db.execute("BEGIN")
        records = [retain_observation(db, item) for item in selection]
    result = {"host_id": args.host_id, "tool_label": "XLStylesTool",
              "source_database": str(args.database), "records": records}
    args.output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps({"inputs": len(records), "eligible": sum(
        row["eligible_for_formatting_comparison"] for row in records)}))


def digest(path):
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def protected_parts(path):
    with ZipFile(path) as package:
        return {name: hashlib.sha256(package.read(name)).hexdigest() for name in package.namelist()
                if name.startswith(("xl/media/", "xl/embeddings/", "xl/drawings/", "xl/charts/"))}


def retain_observation(db, item):
    matches = db.execute("SELECT * FROM jobs WHERE tool='xlstyles' AND input_sha=?",
                         (item["sha256"],)).fetchall()
    if len(matches) != 1:
        raise ValueError(f"Expected one XLStylesTool input identity for {item['file']}")
    job = dict(matches[0])
    attempts = [dict(row) for row in db.execute(
        "SELECT * FROM attempts WHERE job_id=? AND phase='conversion' ORDER BY started", (job["id"],))]
    evidence = json.loads(attempts[0]["evidence"] or "{}") if attempts else {}
    confirmed = bool(attempts and attempts[0]["state"] in ("converted", "no_change") and
                     evidence.get("resources") and evidence.get("last_gui_event", {}).get("phase") == "confirmed")
    record = {"corpus": item["corpus"], "rank": item["rank"], "file": item["file"],
              "input_sha": item["sha256"], "conversion": job["conversion"],
              "output_sha": job["output_sha"], "output_path": job["output_path"],
              "version": job["version"], "configuration": json.loads(job["config"]),
              "attempt_history": attempts, "eligible_for_formatting_comparison": False,
              "exclusion": "complete confirmed primary timing receipt unavailable"}
    if not confirmed:
        return record
    configuration = record["configuration"]
    if any(configuration[key] for key in ("force_normal", "unhide_all", "remove_all_external_names")):
        raise ValueError("The retained XLStylesTool configuration differs from the frozen profile")
    if digest(item["input_path"]) != item["sha256"] or digest(job["output_path"]) != job["output_sha"]:
        raise ValueError(f"Workbook identity mismatch: {item['file']}")
    before = protected_parts(item["input_path"])
    after = protected_parts(job["output_path"])
    changed = sorted(name for name in before.keys() | after.keys() if before.get(name) != after.get(name))
    record["payload_audit"] = {"protected_parts": len(before), "changed_parts": changed}
    record["eligible_for_formatting_comparison"] = not changed
    record["exclusion"] = "changed protected package parts" if changed else None
    return record


if __name__ == "__main__":
    main()
