"""WPS CLI - 会话管理（撤销/重做）。"""

import json
import os
import copy
from typing import Dict, Any, Optional, List
from datetime import datetime


def _locked_save_json(path, data, **dump_kwargs) -> None:
    """原子写入 JSON（带文件锁，Windows 使用 msvcrt）。"""
    try:
        f = open(path, "r+")
    except FileNotFoundError:
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        f = open(path, "w")
    with f:
        _locked = False
        try:
            import msvcrt
            msvcrt.locking(f.fileno(), msvcrt.LK_LOCK, 1)
            _locked = True
        except (ImportError, OSError):
            pass
        try:
            f.seek(0)
            f.truncate()
            json.dump(data, f, **dump_kwargs)
            f.flush()
        finally:
            if _locked:
                try:
                    f.seek(0)
                    import msvcrt
                    msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                except (ImportError, OSError):
                    pass


class Session:
    """管理文档项目状态及撤销/重做历史。"""

    MAX_UNDO = 50

    def __init__(self):
        self.project: Optional[Dict[str, Any]] = None
        self.project_path: Optional[str] = None
        self._undo_stack: List[Dict[str, Any]] = []
        self._redo_stack: List[Dict[str, Any]] = []
        self._modified: bool = False

    def has_project(self) -> bool:
        return self.project is not None

    def is_modified(self) -> bool:
        return self._modified

    def get_project(self) -> Dict[str, Any]:
        if self.project is None:
            raise RuntimeError(
                "没有加载任何文档。请先使用 'document new' 或 'document open'。"
            )
        return self.project

    def set_project(self, project: Dict[str, Any], path: Optional[str] = None) -> None:
        self.project = project
        self.project_path = path
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._modified = False

    def snapshot(self, description: str = "") -> None:
        """在执行变更前保存当前状态到撤销栈。"""
        if self.project is None:
            return
        state = {
            "project": copy.deepcopy(self.project),
            "description": description,
            "timestamp": datetime.now().isoformat(),
        }
        self._undo_stack.append(state)
        if len(self._undo_stack) > self.MAX_UNDO:
            self._undo_stack.pop(0)
        self._redo_stack.clear()
        self._modified = True

    def undo(self) -> Optional[str]:
        """撤销上一步操作。"""
        if not self._undo_stack:
            raise RuntimeError("没有可撤销的操作。")
        if self.project is None:
            raise RuntimeError("没有加载任何文档。")

        self._redo_stack.append({
            "project": copy.deepcopy(self.project),
            "description": "redo point",
            "timestamp": datetime.now().isoformat(),
        })

        state = self._undo_stack.pop()
        self.project = state["project"]
        self._modified = True
        return state.get("description", "")

    def redo(self) -> Optional[str]:
        """重做上一步撤销的操作。"""
        if not self._redo_stack:
            raise RuntimeError("没有可重做的操作。")
        if self.project is None:
            raise RuntimeError("没有加载任何文档。")

        self._undo_stack.append({
            "project": copy.deepcopy(self.project),
            "description": "undo point",
            "timestamp": datetime.now().isoformat(),
        })

        state = self._redo_stack.pop()
        self.project = state["project"]
        self._modified = True
        return state.get("description", "")

    def status(self) -> Dict[str, Any]:
        """获取会话状态。"""
        return {
            "has_project": self.project is not None,
            "project_path": self.project_path,
            "modified": self._modified,
            "undo_count": len(self._undo_stack),
            "redo_count": len(self._redo_stack),
            "document_name": (
                self.project.get("name", "untitled") if self.project else None
            ),
            "document_type": (
                self.project.get("type", "unknown") if self.project else None
            ),
        }

    def save_session(self, path: Optional[str] = None) -> str:
        """保存会话状态（项目）到磁盘。"""
        if self.project is None:
            raise RuntimeError("没有可保存的文档。")

        save_path = path or self.project_path
        if not save_path:
            raise ValueError("未指定保存路径。")

        self.project["metadata"]["modified"] = datetime.now().isoformat()
        _locked_save_json(save_path, self.project, indent=2, sort_keys=True, default=str)

        self.project_path = save_path
        self._modified = False
        return save_path

    def list_history(self) -> List[Dict[str, str]]:
        """列出撤销历史。"""
        result = []
        for i, state in enumerate(reversed(self._undo_stack)):
            result.append({
                "index": i,
                "description": state.get("description", ""),
                "timestamp": state.get("timestamp", ""),
            })
        return result
