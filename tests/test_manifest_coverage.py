from __future__ import annotations

import importlib
import inspect
from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from system_core.core.manifest import CommandNode, load_manifest
from system_core.services import office_service


def leaf_nodes(nodes: tuple[CommandNode, ...] | list[CommandNode]):
    for node in nodes:
        if node.children:
            yield from leaf_nodes(node.children)
        elif node.service:
            yield node


def public_office_services() -> set[str]:
    result: set[str] = set()
    for name, value in inspect.getmembers(office_service, inspect.isfunction):
        if name.startswith("_") or value.__module__ != office_service.__name__:
            continue
        parameters = list(inspect.signature(value).parameters.values())
        if parameters and parameters[0].name == "context":
            result.add(f"{office_service.__name__}:{name}")
    return result


class ManifestCoverageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest = load_manifest(ROOT / "config" / "tool_manifest.yaml")

    def manifest_services(self) -> set[str]:
        operation_services = {
            operation.service
            for operation in (*self.manifest.operations, *self.manifest.maintenance_operations)
            if operation.service
        }
        tree_services = {node.service for node in leaf_nodes(self.manifest.operation_groups)}
        return operation_services | tree_services

    def test_every_public_office_backend_service_has_a_manifest_route(self) -> None:
        self.assertEqual(public_office_services() - self.manifest_services(), set())

    def test_every_manifest_service_resolves_to_a_callable(self) -> None:
        for service in sorted(self.manifest_services()):
            module_name, function_name = service.split(":", 1)
            function = getattr(importlib.import_module(module_name), function_name, None)
            self.assertTrue(callable(function), service)

    def test_restored_table_backends_are_visible_in_gui_tree(self) -> None:
        ids = {node.id for node in leaf_nodes(self.manifest.operation_groups)}
        self.assertIn("docx_table_stitcher_running_header", ids)
        self.assertIn("docx_table_unify_width_only", ids)


if __name__ == "__main__":
    unittest.main()
