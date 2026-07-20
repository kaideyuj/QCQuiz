import json
import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
INDEX_PATH = ROOT / "index.html"
QUESTIONS_PATH = ROOT / "questions_v3.json"


class RepositoryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.index = INDEX_PATH.read_text(encoding="utf-8")
        cls.questions = json.loads(QUESTIONS_PATH.read_text(encoding="utf-8-sig"))

    def test_formal_question_bank_is_valid(self):
        self.assertEqual(len(self.questions), 2648)
        self.assertEqual(
            [question["id"] for question in self.questions],
            list(range(1, 2649))
        )

        required_keys = {
            "id",
            "unit",
            "page",
            "question",
            "options",
            "answer",
            "choice_order"
        }

        for question in self.questions:
            self.assertTrue(required_keys.issubset(question))
            self.assertTrue(question["question"].strip())
            self.assertIn(question["answer"], question["options"])
            self.assertEqual(
                set(question["choice_order"]),
                set(question["options"])
            )
            self.assertEqual(
                len(question["choice_order"]),
                len(question["options"])
            )

    def test_html_uses_only_formal_question_bank(self):
        self.assertIn('const DATA_FILE = "questions_v3.json";', self.index)
        self.assertNotRegex(
            self.index,
            r"fetch\([^\n]*(?<!_v3)questions\.json"
        )

    def test_storage_keys_remain_v4_compatible(self):
        self.assertIn('const PROFILE_STORAGE_KEY = "qcQuizProfilesV4";', self.index)
        self.assertIn('const ACTIVE_USER_KEY = "qcQuizActiveUserV4";', self.index)
        self.assertIn("function normalizeStore(parsed)", self.index)

    def test_versions_are_rendered_from_constants(self):
        self.assertRegex(self.index, r'const HTML_VERSION = "v\d+";')
        self.assertIn('const QUESTIONS_VERSION = "v3";', self.index)
        self.assertIn('id="htmlVersion"', self.index)
        self.assertIn('id="questionsVersion"', self.index)
        self.assertIn(
            '$("htmlVersion").textContent = `HTML ${HTML_VERSION}｜${HTML_BUILD_TIME}`;',
            self.index
        )
        self.assertIn(
            '$("questionsVersion").textContent = `Questions ${QUESTIONS_VERSION}｜${QUESTIONS_BUILD_TIME}`;',
            self.index
        )

    def test_html_is_formatted_across_lines(self):
        self.assertGreater(len(self.index.splitlines()), 1000)
        self.assertIsNone(re.search(r"<style>.*</style>", self.index))
        self.assertIsNone(re.search(r"<script>.*</script>", self.index))


if __name__ == "__main__":
    unittest.main()
