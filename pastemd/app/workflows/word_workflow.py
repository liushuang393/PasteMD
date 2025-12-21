"""Word document workflow."""

from .base import BaseWorkflow
from ...domains.document import WordPlacer
from ...utils.clipboard import (
    get_clipboard_text, get_clipboard_html, is_clipboard_empty,
    read_markdown_files_from_clipboard
)
from ...utils.html_analyzer import is_plain_html_fragment
from ...utils.markdown_utils import merge_markdown_contents
from ...utils.fs import generate_output_path
from ...core.errors import ClipboardError, PandocError


class WordWorkflow(BaseWorkflow):
    """Word 文档工作流"""
    
    def __init__(self):
        super().__init__()
        self.placer = WordPlacer()  # 无需工厂,直接实例化
    
    def execute(self) -> None:
        """执行 Word 工作流"""
        try:
            # 1. 读取剪贴板
            content_type, content = self._read_clipboard()
            self._log(f"Clipboard content type: {content_type}")
            
            # 2. 预处理
            if content_type == "markdown":
                content = self.markdown_preprocessor.process(content)
            
            # 3. 生成 DOCX
            if content_type == "html":
                docx_bytes = self.doc_generator.convert_html_to_docx_bytes(
                    content, self.config
                )
            else:
                docx_bytes = self.doc_generator.convert_markdown_to_docx_bytes(
                    content, self.config
                )
            
            # 4. 落地内容(不做降级,失败即报错)
            result = self.placer.place(docx_bytes, self.config)
            
            # 5. 通知结果
            if result.success:
                method_str = result.method or "unknown"
                self._notify_success(f"成功插入到 Word (方式: {method_str})")
            else:
                self._notify_error(result.error or "Word 插入失败")
            
            # 6. 可选保存
            if result.success and self.config.get("keep_file", False):
                self._save_docx(docx_bytes)
        
        except ClipboardError as e:
            self._log(f"Clipboard error: {e}")
            self._notify_error("剪贴板读取失败")
        except PandocError as e:
            self._log(f"Pandoc error: {e}")
            self._notify_error("文档转换失败")
        except Exception as e:
            self._log(f"Word workflow failed: {e}")
            import traceback
            traceback.print_exc()
            self._notify_error("操作失败")
    
    def _read_clipboard(self) -> tuple[str, str]:
        """
        读取剪贴板,返回 (类型, 内容)
        
        Returns:
            ("html" | "markdown", content)
        """
        # 优先 HTML
        try:
            html = get_clipboard_html(self.config)
            if not is_plain_html_fragment(html):
                return ("html", html)
        except ClipboardError:
            pass
        
        # 降级 Markdown
        if not is_clipboard_empty():
            return ("markdown", get_clipboard_text())
        
        # 尝试 MD 文件
        found, files_data, _ = read_markdown_files_from_clipboard()
        if found:
            merged = merge_markdown_contents(files_data)
            return ("markdown", merged)
        
        raise ClipboardError("剪贴板为空或无有效内容")
    
    def _save_docx(self, docx_bytes: bytes):
        """保存 DOCX 到磁盘"""
        try:
            output_path = generate_output_path(
                keep_file=True,
                save_dir=self.config.get("save_dir", ""),
                md_text=""
            )
            with open(output_path, "wb") as f:
                f.write(docx_bytes)
            self._log(f"Saved DOCX to: {output_path}")
        except Exception as e:
            self._log(f"Failed to save DOCX: {e}")
