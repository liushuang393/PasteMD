"""Base workflow class."""

from abc import ABC, abstractmethod
from ...core.state import app_state
from ...domains.notification.manager import NotificationManager
from ...domains.document import DocumentGenerator
from ...domains.spreadsheet import SpreadsheetGenerator
from ...domains.preprocessor import MarkdownPreprocessor
from ...utils.logging import log


class BaseWorkflow(ABC):
    """工作流基类"""
    
    def __init__(self):
        # 轻量级初始化
        self.config = app_state.config
        self.notification_manager = NotificationManager()  # 复用单例
        
        # 延迟初始化(避免 Pandoc 开销)
        self._doc_generator = None
        self._sheet_generator = None
        self._markdown_preprocessor = None
    
    @property
    def doc_generator(self):
        """懒加载 DocumentGenerator"""
        if self._doc_generator is None:
            self._doc_generator = DocumentGenerator()
        return self._doc_generator
    
    @property
    def sheet_generator(self):
        """懒加载 SpreadsheetGenerator"""
        if self._sheet_generator is None:
            self._sheet_generator = SpreadsheetGenerator()
        return self._sheet_generator
    
    @property
    def markdown_preprocessor(self):
        """懒加载 Markdown Preprocessor"""
        if self._markdown_preprocessor is None:
            self._markdown_preprocessor = MarkdownPreprocessor(self.config)
        return self._markdown_preprocessor
    
    @abstractmethod
    def execute(self) -> None:
        """执行工作流(子类实现)"""
        pass
    
    # 公共辅助方法
    def _notify_success(self, msg: str):
        """通知成功"""
        self.notification_manager.notify("PasteMD", msg, ok=True)
    
    def _notify_error(self, msg: str):
        """通知错误"""
        self.notification_manager.notify("PasteMD", msg, ok=False)
    
    def _log(self, msg: str):
        """记录日志"""
        log(msg)
