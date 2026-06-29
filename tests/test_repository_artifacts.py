from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parents[1]


class RepositoryArtifactTests(unittest.TestCase):
    def test_conversion_page_has_no_accidental_invalid_token(self):
        content = (ROOT / "src" / "Conversion" / "Conversion-Page.psm1").read_text(encoding="utf-8")

        self.assertNotIn("else{~", content)

    def test_recovery_script_exposes_safe_recovery_workflow(self):
        script = ROOT / "tools" / "Invoke-OneNoteRecovery.ps1"
        content = script.read_text(encoding="utf-8")

        self.assertIn("param(", content)
        self.assertIn("ValidateSet(", content)
        self.assertIn("RetryV1", content)
        self.assertIn("RetryV2", content)
        self.assertIn("Merge", content)
        self.assertIn("FixV1Misclassified", content)
        self.assertIn("FixV1Flat", content)
        self.assertIn("ReleaseComObject", content)
        self.assertIn("Start-SubstDrive", content)
        self.assertRegex(content, r"function\s+Parse-FailedLog")
        self.assertRegex(content, r"function\s+Copy-FileAndHtmlResources")

    def test_readme_documents_2026_recovery_usage(self):
        readme = (ROOT / "README.md").read_text(encoding="utf-8")

        self.assertIn("2026 recovery workflow", readme.lower())
        self.assertIn("tools\\Invoke-OneNoteRecovery.ps1", readme)
        self.assertIn("OneNote-Export-failed-pages.log", readme)
        self.assertIn("PowerShell 5.1", readme)

    def test_gitignore_ignores_generated_export_artifacts(self):
        gitignore = (ROOT / ".gitignore").read_text(encoding="utf-8")

        for pattern in [
            "config.ps1",
            "notes/",
            "OneNote-Export/",
            "OneNote-Export-Recovered/",
            "ONR2/",
            "*.log",
            "*.tmp",
        ]:
            self.assertRegex(gitignore, rf"(?m)^{re.escape(pattern)}$")


if __name__ == "__main__":
    unittest.main()
