import unittest

from scripts.comment_label_utils import extract_comment_label, validate_session_labels


class CommentLabelTests(unittest.TestCase):
    def test_extracts_two_and_three_digit_session_sequences(self):
        for label in (
            "S-20261106-99",
            "S-20261106-100",
            "S-20261106-101",
            "S-20261106-104",
        ):
            with self.subTest(label=label):
                self.assertEqual(extract_comment_label(f"{label} | חומרה: גבוהה"), label)

    def test_extracts_legacy_editor_label(self):
        self.assertEqual(extract_comment_label("E-0105 | חומרה: קריטית"), "E-0105")

    def test_does_not_extract_a_label_embedded_in_comment_body(self):
        self.assertIsNone(extract_comment_label("ההערה מזכירה S-20261106-100 בגוף הטקסט"))

    def test_validates_variable_length_session_labels(self):
        validate_session_labels(["S-20261106-99", "S-20261106-100", "S-20261106-104"])

    def test_rejects_duplicate_labels(self):
        with self.assertRaisesRegex(ValueError, "Duplicate"):
            validate_session_labels(["S-20261106-100", "S-20261106-100"])

    def test_rejects_malformed_labels(self):
        with self.assertRaisesRegex(ValueError, "Invalid"):
            validate_session_labels(["S-20261106-1"])


if __name__ == "__main__":
    unittest.main()
