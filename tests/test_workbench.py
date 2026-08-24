from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from system_core.ui_nicegui.workbench import (
    CANONICAL_WORKBENCH_TEXT,
    WorkbenchConfig,
    WorkbenchHistory,
    canonical_role,
    canonical_workbench_text,
)


class WorkbenchHistoryTests(unittest.TestCase):
    def make_history(self, root: Path, *, limit: int = 4) -> WorkbenchHistory:
        config = WorkbenchConfig(
            root=root,
            input_path=root / "input",
            output_path=root / "output",
            history_path=root / "config" / "path_history.json",
            history_limit=limit,
        )
        return WorkbenchHistory(config)

    def test_roles_are_canonical(self) -> None:
        self.assertEqual(canonical_role("source"), "source")
        self.assertEqual(canonical_role("destination"), "target")
        self.assertEqual(canonical_role("dst"), "target")

    def test_public_vocabulary_matches_image_tools_contract(self) -> None:
        self.assertEqual(
            [CANONICAL_WORKBENCH_TEXT["ru"][key] for key in (
                "source_folder",
                "add_file_short",
                "target_folder",
                "clear_io_short",
                "delete_io_short",
                "file_list_button",
            )],
            ["Источник", "Добавить файл...", "Назначение", "Сбросить", "Удалить", "Список"],
        )
        self.assertEqual(canonical_workbench_text("en", "clear_io_short"), "Reset")

    def test_history_remember_pin_clear_and_delete(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            history = self.make_history(root)
            history.ensure_initial()
            source = root / "outside source"
            target = root / "outside target"
            history.remember("source", str(source))
            history.remember("target", str(target))
            history.set_pinned("source", str(source), True, required_message="required")

            result = history.clear_cache_keep_pins()
            self.assertEqual(result["kept_pins"], 1)
            self.assertEqual([item["path"] for item in history.entries("source")], [str(source)])
            self.assertEqual(history.entries("target"), [])

            deleted = history.delete("source", str(source), required_message="required")
            self.assertEqual(deleted["removed"], 1)
            self.assertEqual(deleted["next_path"], str(root / "input"))

    def test_history_limit_preserves_most_recent_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            history = self.make_history(root, limit=2)
            for name in ("one", "two", "three"):
                history.remember("source", str(root / name))
            self.assertEqual(len(history.entries("source")), 2)


if __name__ == "__main__":
    unittest.main()
