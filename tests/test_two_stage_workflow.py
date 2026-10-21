from __future__ import annotations

import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
EXAMPLE = ROOT / "assets" / "content_review.example.json"


class TwoStageWorkflowTest(unittest.TestCase):
    def test_content_review_validates_and_renders_markdown(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            digest = tmp_path / "source_digest.json"
            qa = tmp_path / "content_review_qa.json"
            proof = tmp_path / "slide_content_review.md"
            digest.write_text(
                json.dumps(
                    {
                        "facts": [
                            {"id": "F1", "statement": "应用和硬件共同驱动映射"},
                            {"id": "F2", "statement": "映射包含任务切分和配置生成"},
                        ]
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            result = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPTS / "validate_content_review.py"),
                    str(EXAMPLE),
                    "--source-digest",
                    str(digest),
                    "--json-out",
                    str(qa),
                    "--markdown-out",
                    str(proof),
                    "--strict",
                ],
                check=False,
                text=True,
                capture_output=True,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            report = json.loads(qa.read_text(encoding="utf-8"))
            self.assertTrue(report["passed"])
            self.assertEqual(report["coverage"]["critical_visible"], 1.0)
            self.assertIn("尚未生成PPTX", proof.read_text(encoding="utf-8"))

    def test_missing_required_visible_term_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            review_path = Path(tmp) / "review.json"
            review = json.loads(EXAMPLE.read_text(encoding="utf-8"))
            review["slides"][0]["must_show_terms"].append("不存在的关键机制")
            review_path.write_text(json.dumps(review, ensure_ascii=False), encoding="utf-8")
            result = subprocess.run(
                [sys.executable, str(SCRIPTS / "validate_content_review.py"), str(review_path), "--strict"],
                check=False,
                text=True,
                capture_output=True,
            )
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("不存在的关键机制", result.stdout)

    def test_approval_digest_locks_two_stage_build(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            review_path = tmp_path / "content_review.json"
            approval_path = tmp_path / "content_approval.json"
            spec_path = tmp_path / "deck_spec.json"
            shutil.copyfile(EXAMPLE, review_path)

            approval = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPTS / "record_content_approval.py"),
                    str(review_path),
                    "--instruction",
                    "内容通过，继续生成PPT",
                    "--out",
                    str(approval_path),
                ],
                check=False,
                text=True,
                capture_output=True,
            )
            self.assertEqual(approval.returncode, 0, approval.stdout + approval.stderr)

            spec_path.write_text(
                json.dumps(
                    {
                        "version": 1,
                        "title": "Two-stage build",
                        "mode": "native",
                        "workflow": {
                            "mode": "two_stage",
                            "stage": "generation",
                            "content_review": "content_review.json",
                            "approval_record": "content_approval.json",
                        },
                        "slides": [
                            {
                                "type": "title_body",
                                "title": "应用语义与硬件约束共同驱动架构感知映射",
                                "takeaway": "映射器依据应用算子需求画像执行PE阵列映射与配置生成。",
                                "bullets": ["应用算子需求画像", "PE阵列映射", "配置生成"],
                                "fact_ids": ["F1", "F2"],
                                "must_show_terms": ["应用算子需求画像", "PE阵列映射", "配置生成"],
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            valid = subprocess.run(
                [sys.executable, str(SCRIPTS / "validate_deck_spec.py"), str(spec_path), "--json"],
                check=False,
                text=True,
                capture_output=True,
            )
            self.assertEqual(valid.returncode, 0, valid.stdout + valid.stderr)

            review_path.write_text(review_path.read_text(encoding="utf-8") + "\n", encoding="utf-8")
            invalid = subprocess.run(
                [sys.executable, str(SCRIPTS / "validate_deck_spec.py"), str(spec_path), "--json"],
                check=False,
                text=True,
                capture_output=True,
            )
            self.assertNotEqual(invalid.returncode, 0)
            self.assertIn("changed after approval", invalid.stdout)


if __name__ == "__main__":
    unittest.main()
