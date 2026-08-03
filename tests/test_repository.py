import json
import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
INDEX_PATH = ROOT / "index.html"
QUESTIONS_PATH = ROOT / "questions_v3.json"
XLSX_EXPORT_PATH = ROOT / "xlsx-export.js"


class RepositoryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.index = INDEX_PATH.read_text(encoding="utf-8")
        cls.questions = json.loads(QUESTIONS_PATH.read_text(encoding="utf-8-sig"))
        cls.xlsx_export = XLSX_EXPORT_PATH.read_text(encoding="utf-8")

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

    def test_yuj_profiles_use_secure_cloud_sync_endpoint(self):
        self.assertIn(
            'const SYNC_API_URL = "https://qcquiz-sync.kaideyuj.workers.dev";',
            self.index
        )
        self.assertIn('const SYNC_PREFIX = "yuj-";', self.index)
        self.assertIn('id="syncPassword"', self.index)
        self.assertIn('"X-Sync-Password": password', self.index)
        self.assertNotIn("GITHUB_TOKEN", self.index)
        self.assertIn("是否將目前代號", self.index)
        self.assertIn("migratedFrom: previousCode", self.index)

    def test_records_view_and_excel_export_are_available(self):
        self.assertIn('id="records"', self.index)
        self.assertIn('onclick="openRecords()"', self.index)
        self.assertIn('onclick="exportRecordsToExcel()"', self.index)
        self.assertIn('src="xlsx-export.js"', self.index)
        self.assertIn('name: "測驗紀錄"', self.index)
        self.assertIn('name: "作答明細"', self.index)
        self.assertIn('["使用者", "完成時間", "模式"', self.index)
        self.assertIn("function buildWorkbook(inputSheets)", self.xlsx_export)
        self.assertIn("application/vnd.openxmlformats-officedocument", self.xlsx_export)

    def test_clear_all_profiles_is_scoped_and_password_protected(self):
        self.assertIn('const CLEAR_ALL_PASSWORD = "1234";', self.index)
        self.assertIn('id="clearAllPassword"', self.index)
        self.assertIn('type="password"', self.index)
        self.assertIn("function confirmClearAllRecords()", self.index)
        self.assertIn("localStorage.removeItem(PROFILE_STORAGE_KEY);", self.index)
        self.assertIn("localStorage.removeItem(ACTIVE_USER_KEY);", self.index)
        self.assertNotIn("localStorage.clear()", self.index)
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

    def test_quiz_navigation_and_feedback_layout(self):
        self.assertIn('"previous next"', self.index)
        self.assertIn('"home home"', self.index)
        self.assertIn("#nextBtn,\n    #submitBtn", self.index)
        self.assertNotIn("答錯。正確答案：", self.index)
        self.assertNotIn("答對：", self.index)
        self.assertNotIn('id="feedback"', self.index)
        self.assertNotIn("renderFeedback", self.index)
        self.assertIn('const quizUserLabel = quizContext?.userCode || "訪客";', self.index)
        self.assertIn('`使用者：${quizUserLabel}｜原題號', self.index)

    def test_answers_are_graded_only_after_confirmation(self):
        self.assertIn("function selectAnswer(question, selectedKey)", self.index)
        self.assertIn("function confirmCurrentAnswer()", self.index)
        self.assertIn("confirmed: false", self.index)
        self.assertIn("response.confirmed = true;", self.index)
        self.assertIn('hasPendingAnswer ? "提送" : "下一題"', self.index)
        self.assertIn('hasPendingAnswer ? "提送" : "交卷"', self.index)
        self.assertNotIn("function answerQuestion(question, selectedKey)", self.index)

if __name__ == "__main__":
    unittest.main()
