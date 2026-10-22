"""跨检索模式共享的统计、进度快照和停止原因。"""

from dataclasses import dataclass
from pathlib import Path
from time import monotonic
from typing import Callable


@dataclass(frozen=True)
class SearchIssue:
    path: Path
    message: str


@dataclass(frozen=True)
class SearchSnapshot:
    scanned_count: int
    result_count: int
    current_path: str
    elapsed_seconds: float
    stop_reason: str
    issue_count: int


def describe_error(error: BaseException) -> str:
    """保留底层错误，帮助定位损坏、权限和格式兼容问题。"""

    messages = []
    seen = set()
    while error is not None and id(error) not in seen:
        seen.add(id(error))
        message = str(error) or type(error).__name__
        if message not in messages:
            messages.append(message)
        error = error.__cause__
    return "：".join(messages)


class SearchTracker:
    """由一个后台任务写入；界面只接收不可变快照，最多每秒更新五次。"""

    def __init__(self, callback: Callable[[SearchSnapshot], None] | None = None) -> None:
        self.callback = callback
        self.started = monotonic()
        self.last_emit = float("-inf")
        self.scanned_count = 0
        self.result_count = 0
        self.current_path = ""
        self.stop_reason = "running"
        self.issues: list[SearchIssue] = []

    def scan(self, path: Path) -> None:
        self.scanned_count += 1
        self.current_path = str(path)
        self.emit()

    def issue(self, path: Path, error: BaseException) -> None:
        self.issues.append(SearchIssue(path, describe_error(error)))
        self.emit()

    def snapshot(self) -> SearchSnapshot:
        return SearchSnapshot(
            self.scanned_count, self.result_count, self.current_path,
            monotonic() - self.started, self.stop_reason, len(self.issues),
        )

    def emit(self, force: bool = False) -> None:
        now = monotonic()
        if self.callback is not None and (force or now - self.last_emit >= 0.2):
            self.last_emit = now
            self.callback(self.snapshot())

    def finish(self, reason: str = "completed") -> None:
        self.stop_reason = reason
        self.emit(force=True)
