from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import contextlib
import json
import os
import shutil
import time

from system_core.core.jobs import JobContext
from tqdm import tqdm


class _LogStream:
    def __init__(self, context: JobContext) -> None:
        self.context = context
        self.buffer = ""
        self.last_progress_emit = 0.0

    def writable(self) -> bool:
        return True

    def write(self, text: str) -> int:
        if not text:
            return 0

        for char in text:
            if char == "\n":
                self._emit(self.buffer)
                self.buffer = ""
            elif char == "\r":
                self._emit_progress_snapshot()
                self.buffer = ""
            else:
                self.buffer += char
        return len(text)

    def flush(self) -> None:
        self._emit(self.buffer)
        self.buffer = ""

    def _emit(self, text: str) -> None:
        line = text.rstrip()
        if line:
            self.context.log(line)

    def _emit_progress_snapshot(self) -> None:
        now = time.monotonic()
        if now - self.last_progress_emit < 0.5:
            return
        self.last_progress_emit = now
        self._emit(self.buffer)


@contextlib.contextmanager
def relay_output(context: JobContext):
    stream = _LogStream(context)
    with contextlib.redirect_stdout(stream), contextlib.redirect_stderr(stream):
        yield
    stream.flush()


def _inventory(folder: Path) -> dict[str, int]:
    if not folder.exists():
        return {"files": 0, "dirs": 0}

    files = 0
    dirs = 0
    for item in folder.rglob("*"):
        if item.is_file():
            files += 1
        elif item.is_dir():
            dirs += 1
    return {"files": files, "dirs": dirs}


def validate_input(context: JobContext) -> dict[str, object]:
    context.paths.input.mkdir(parents=True, exist_ok=True)
    inventory = _inventory(context.paths.input)
    context.log(f"Input folder: {context.paths.input}")
    context.log(f"Input inventory: {inventory['files']} files, {inventory['dirs']} directories")
    context.progress(1.0)
    return {"input": str(context.paths.input), "inventory": inventory}


def run_sample_job(context: JobContext) -> dict[str, object]:
    context.paths.output.mkdir(parents=True, exist_ok=True)

    with relay_output(context):
        steps = range(1, 6)
        for step in tqdm(steps, desc="Sample job", unit="step"):
            if context.cancelled():
                context.log("Operation cancelled by user.")
                return {"cancelled": True}
            print(f"Processing sample step {step} of 5")
            context.progress(step / 5)
            time.sleep(0.2)

    report = {
        "project_root": str(context.paths.root),
        "input": _inventory(context.paths.input),
        "parameters": context.operation.parameters,
        "message": "Replace sample_service.py with real project services.",
    }
    report_path = context.paths.output / "sample_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    context.log(f"Created: {report_path}")
    return {"report": str(report_path)}


def package_output(context: JobContext) -> dict[str, object]:
    context.paths.output.mkdir(parents=True, exist_ok=True)
    context.paths.release.mkdir(parents=True, exist_ok=True)

    files = [p for p in context.paths.output.rglob("*") if p.is_file()]
    zip_path = context.paths.release / "output_package.zip"

    with ZipFile(zip_path, "w", ZIP_DEFLATED) as archive:
        total = max(1, len(files))
        for index, file_path in enumerate(files, start=1):
            if context.cancelled():
                context.log("Packaging cancelled by user.")
                return {"cancelled": True}
            archive.write(file_path, file_path.relative_to(context.paths.output).as_posix())
            context.progress(index / total)

    context.log(f"Created: {zip_path}")
    return {"zip": str(zip_path), "files": len(files)}


def _is_inside(child: Path, parent: Path) -> bool:
    """True iff `child` (after resolve) is inside `parent` (after resolve).

    Uses os.path.commonpath which is stricter than `parent in child.parents`
    when `child == parent` and works correctly across symlinks because both
    sides are already resolved.
    """
    try:
        child_resolved = str(child.resolve())
        parent_resolved = str(parent.resolve())
    except OSError:
        return False
    try:
        common = os.path.commonpath([child_resolved, parent_resolved])
    except ValueError:
        # Different drives on Windows, etc.
        return False
    return common == parent_resolved


def _clean_managed_folder(context: JobContext, folder: Path, label: str) -> dict[str, object]:
    root = context.paths.root.resolve()
    if folder.is_symlink():
        raise RuntimeError(f"{label} is a symbolic link. Cleanup blocked for safety.")

    folder.mkdir(parents=True, exist_ok=True)
    folder_resolved = folder.resolve()
    if not _is_inside(folder_resolved, root):
        raise RuntimeError(f"{label} path is outside project root. Cleanup blocked.")

    removed = 0
    skipped: list[str] = []
    for item in folder.iterdir():
        if item.name == ".gitkeep":
            continue
        try:
            if item.is_symlink():
                item.unlink()
            elif item.is_dir():
                if not _is_inside(item, folder_resolved):
                    skipped.append(f"{item.name} (escapes {label})")
                    continue
                shutil.rmtree(item)
            else:
                item.unlink()
            removed += 1
            context.log(f"Removed from {label}: {item.name}")
        except OSError as exc:
            skipped.append(f"{item.name} ({exc})")

    return {"folder": label, "removed_items": removed, "skipped_items": skipped}


def cleanup_input_output(context: JobContext) -> dict[str, object]:
    context.log("Cleaning managed input/output folders.")
    input_result = _clean_managed_folder(context, context.paths.input, "input")
    context.progress(0.5)
    if context.cancelled():
        context.log("Input/output cleanup cancelled by user after input cleanup.")
        return {"cancelled": True, "input": input_result}

    output_result = _clean_managed_folder(context, context.paths.output, "output")
    context.progress(1.0)
    total_removed = int(input_result["removed_items"]) + int(output_result["removed_items"])
    total_skipped = len(input_result["skipped_items"]) + len(output_result["skipped_items"])
    context.log(f"Input/output cleanup complete. Removed: {total_removed}, skipped: {total_skipped}")
    return {"input": input_result, "output": output_result, "removed_items": total_removed, "skipped_items": total_skipped}


def cleanup_workspace(context: JobContext) -> dict[str, object]:
    """Delete only files inside the managed workspace folder.

    Safety rules:
    - workspace must resolve to a path inside the project root.
    - Symlinks inside workspace are unlinked, never traversed. shutil.rmtree
      with follow_symlinks=False is not enough on Windows because Path.rmtree
      via shutil treats junctions inconsistently — we handle symlinks first
      and explicitly.
    - Non-symlink subdirectories are checked again before removal: their
      resolved path must still be inside workspace. This catches the case
      where workspace itself contains a directory whose resolve() escapes
      (e.g. a junction or reparse point set up between the outer check and
      iteration).
    """
    workspace = context.paths.workspace
    root = context.paths.root.resolve()

    # Reject workspaces that are themselves symlinks pointing outside root.
    if workspace.is_symlink():
        raise RuntimeError("Workspace is a symbolic link. Cleanup blocked for safety.")

    workspace_resolved = workspace.resolve()
    if not _is_inside(workspace_resolved, root):
        raise RuntimeError("Workspace path is outside project root. Cleanup blocked.")

    workspace.mkdir(parents=True, exist_ok=True)

    removed = 0
    skipped: list[str] = []

    for item in workspace.iterdir():
        if item.name == ".gitkeep":
            continue

        # Symlinks: unlink the link itself, do NOT follow.
        if item.is_symlink():
            try:
                item.unlink()
                removed += 1
                context.log(f"Removed symlink: {item.name}")
            except OSError as exc:
                skipped.append(f"{item.name} (symlink, {exc})")
            continue

        # Re-verify directory containment before recursive delete.
        if item.is_dir():
            if not _is_inside(item, workspace_resolved):
                skipped.append(f"{item.name} (escapes workspace after resolve)")
                context.log(f"Skipped (escapes workspace): {item.name}")
                continue
            shutil.rmtree(item)
            removed += 1
            continue

        # Regular file.
        item.unlink()
        removed += 1

    context.log(f"Workspace cleanup complete. Removed: {removed}, skipped: {len(skipped)}")
    if skipped:
        context.log("Skipped items: " + ", ".join(skipped))
    context.progress(1.0)
    return {"removed_items": removed, "skipped_items": skipped}
