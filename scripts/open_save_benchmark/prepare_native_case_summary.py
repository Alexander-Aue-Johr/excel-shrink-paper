"""Join retained native audit scopes to exact case-study outputs and verified controls."""
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
EVIDENCE = ROOT / "papers/sheetshrink/evidence"
CASES = EVIDENCE / "open-save-case-study"
TOOL_IDS = {"SheetShrink": "sheetshrink", "CheckPerformance": "check_performance",
            "Inquire": "inquire", "NXPowerLite": "nxpowerlite", "XLStylesTool": "xlstyles"}


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def matched_report(audit, original_sha, tool, output_sha, control=None):
    reports = [r for r in audit["pairs"]
               if r["original_sha256"] == original_sha and r["tool"] == TOOL_IDS[tool]]
    if not reports:
        return None
    report = reports[0]
    if control:
        native_manifest = report["member_manifests"]["candidate"]["canonical_manifest_sha256"]
        assert control["all_members_equal_native_candidate"]
        assert control["canonical_member_manifest_sha256"] == native_manifest
    elif report["prepared_variant_sha256"] != output_sha:
        return None
    return report


def make_visual(report, historical):
    if report is None:
        return {"status": historical["status"] if historical else "no_matched_report"}
    result = {"status": report["status"], "report_sha256": report["report_sha256"],
              "baseline_scope": report["baseline_scope"]}
    comparisons = [c for c in report["comparisons"] if c["left"] == "OR" and c["right"] == "SDR"]
    if comparisons:
        assert len(comparisons) == 1
        result.update(comparisons[0])
    return result


def main():
    cases = read_json(CASES / "sizes.json")
    audit = read_json(CASES / "opening-memory-native-case-audit-20261111.json")
    historical = read_json(CASES / "visual-summary.json")
    controls = read_json(EVIDENCE / "bounded-open-save-followup/fresh-native-joins.json")["cases"]
    controls += read_json(EVIDENCE / "enron-open-save-followup/fresh-native-join.json")["cases"]
    enron = read_json(EVIDENCE / "enron-open-save-followup/protocol.json")
    result = []
    for case in cases:
        if case["case_id"] == "finance":
            continue
        original = case["variants"]["Original"]
        for tool, tool_id in TOOL_IDS.items():
            control = None
            size = case["variants"].get(tool)
            if case["case_id"] == "enron":
                output_sha = enron["prepared_variant_sha256"][tool]
                report = matched_report(audit, original["sha256"], tool, output_sha)
                assert report is not None
                payload = report["member_manifests"]["candidate"]["all_members_uncompressed_bytes"]
            else:
                output_sha = size["sha256"]
                payload = size["uncompressed_bytes"]
            if tool == "SheetShrink":
                control = next(c for c in controls if c["case"] == case["case_id"])
                assert control["original_sha256"] == original["sha256"]
                output_sha = control["fresh_output_sha256"]
                payload = control["all_members_uncompressed_bytes"]
            report = matched_report(audit, original["sha256"], tool, output_sha, control)
            previous = next((r for r in historical if r["case_id"] == case["case_id"]
                             and r["tool"] == tool), None)
            result.append({"case_id": case["case_id"], "tool": tool,
                           "original_sha256": original["sha256"], "output_sha256": output_sha,
                           "uncompressed_bytes": payload, "representative_control": bool(control),
                           "visual": make_visual(report, previous)})
    assert len(result) == 15
    (CASES / "matched-native-summary.json").write_text(
        json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print("Verified 15 case/tool joins, retaining incomplete native scopes")


if __name__ == "__main__":
    main()
