import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest


@unittest.skipUnless(os.name == "nt" and shutil.which("powershell.exe"), "Windows PowerShell is required")
class PreparationJobTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.repo = self.root / "rust"
        modules = self.repo / "tools/corpus-launcher"
        modules.mkdir(parents=True)
        (modules / "Core.psm1").write_text('''
function Get-Value($Object,$Name,$Default=$null) {
    $property=$Object.PSObject.Properties[$Name]
    if($null -eq $property -or $null -eq $property.Value){return $Default}
    return $property.Value
}
function Import-StudyConfiguration($Repo,$ConfigurationPath) {
    return Get-Content -LiteralPath $ConfigurationPath -Raw | ConvertFrom-Json
}
function New-StudyLaunchPlan {
    param($Configuration,$Repo,$Executable,$RunDirectory,[switch]$PrepareOnly)
    if(-not $PrepareOnly){throw 'Preparation mode was not requested.'}
    return [pscustomobject]@{Arguments=@('--prepare-only');CommandLine='';Executable=$Executable;Stderr=(Join-Path $RunDirectory 'stderr.log')}
}
function ConvertTo-WindowsArgument($Argument){return '"'+$Argument+'"'}
function Save-StudyJson($Path,$Value){$Value|ConvertTo-Json -Depth 20|Set-Content -LiteralPath $Path -Encoding UTF8}
Export-ModuleMember -Function *
''', encoding="utf-8")
        (modules / "Launch.psm1").write_text('''
function Start-StudyLaunch($Configuration,$ConfigurationPath,$Repo,$Plan) {
    return [pscustomobject]@{Pid=42;Stderr=$Plan.Stderr;Arguments=$Plan.Arguments;CommandLine=$Plan.CommandLine;Executable=$Plan.Executable;Settings=$Configuration.settings;Corpora=$Configuration.corpora;ConfigurationPath=$ConfigurationPath}
}
Export-ModuleMember -Function Start-StudyLaunch
''', encoding="utf-8")
        self.base = self.root / "base.json"
        configuration = {"version": 1, "settings": {
            "corpus": "Enron", "output_root": "old-output", "control_room_port": 8771,
            "partition": 1, "comparison_groups": 0, "hide_excel": False, "build_release": True,
            "order": "size-desc", "import_report": "old-report.json", "reuse_prepared_from": ["b" * 64],
            "worker_memory_mib": 4096, "machine_label": "LOCAL", "export_results": True,
        }, "corpora": [{"name": "Enron", "id": "c" * 64}]}
        self.base.write_text(json.dumps(configuration), encoding="utf-8")
        self.original_base = self.base.read_bytes()
        executable = self.root / "reviewed.exe"
        executable.write_bytes(b"reviewed fixture")
        self.receipt = self.root / "receipt.json"
        self.job = {
            "rust_repo": str(self.repo), "base_configuration": str(self.base),
            "configuration_path": str(self.root / "job-config.json"),
            "run_directory": str(self.root / "run"), "executable": str(executable),
            "executable_sha256": hashlib.sha256(executable.read_bytes()).hexdigest(),
            "receipt_path": str(self.receipt), "corpus": "Enron", "output_root": str(self.root / "enron"),
            "port": 8771, "approved_builds": ["a" * 64], "retry_skipped": True,
        }
        self.script = Path(__file__).resolve().parents[1] / "scripts/run_preparation_job.ps1"

    def tearDown(self):
        self.temporary.cleanup()

    def run_job(self):
        path = self.root / "job.json"
        path.write_text(json.dumps(self.job), encoding="utf-8")
        return subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                               "-File", str(self.script), "-JobFile", str(path)],
                              capture_output=True, text=True, timeout=20)

    def test_enron_retry_uses_approved_builds_and_preserves_base_configuration(self):
        result = self.run_job()
        self.assertEqual(0, result.returncode, result.stderr)
        receipt = json.loads(self.receipt.read_text(encoding="utf-8-sig"))
        settings = receipt["Settings"]
        self.assertEqual(["--prepare-only", "--retry-skipped"], receipt["Arguments"])
        self.assertIn('"--retry-skipped"', receipt["CommandLine"])
        self.assertEqual(0, settings["partition"])
        self.assertEqual(1, settings["comparison_groups"])
        self.assertEqual(["a" * 64], settings["reuse_prepared_from"])
        self.assertEqual(4096, settings["worker_memory_mib"])
        self.assertEqual("LOCAL", settings["machine_label"])
        self.assertFalse(settings["build_release"])
        self.assertTrue(settings["hide_excel"])
        self.assertTrue(settings["export_results"])
        self.assertIsNone(settings["import_report"])
        self.assertEqual(self.original_base, self.base.read_bytes())
        self.assertEqual(self.job["configuration_path"], receipt["ConfigurationPath"])

    def test_euses_has_its_own_corpus_mapping_without_retry(self):
        self.job.update(corpus="EUSES", archive=str(self.root / "euses.zip"), corpus_id="d" * 64,
                        total_workbooks=4498, port=8772, retry_skipped=False, approved_builds=[])
        result = self.run_job()
        self.assertEqual(0, result.returncode, result.stderr)
        receipt = json.loads(self.receipt.read_text(encoding="utf-8-sig"))
        self.assertEqual(["--prepare-only"], receipt["Arguments"])
        self.assertEqual(8772, int(receipt["Settings"]["control_room_port"]))
        self.assertEqual([{"name": "EUSES", "id": "d" * 64, "local_archive": self.job["archive"],
                          "url": None, "total_workbooks": 4498}], receipt["Corpora"])
        self.assertEqual(self.original_base, self.base.read_bytes())

    def test_changed_executable_is_rejected_before_launch(self):
        self.job["executable_sha256"] = "0" * 64
        result = self.run_job()
        self.assertNotEqual(0, result.returncode)
        self.assertIn("differs from the reviewed build", result.stderr)
        self.assertFalse(self.receipt.exists())

    def test_job_cannot_overwrite_base_configuration(self):
        self.job["configuration_path"] = str(self.base)
        result = self.run_job()
        self.assertNotEqual(0, result.returncode)
        self.assertIn("separate from the base configuration", result.stderr)
        self.assertEqual(self.original_base, self.base.read_bytes())
        self.assertFalse(self.receipt.exists())


if __name__ == "__main__":
    unittest.main()
