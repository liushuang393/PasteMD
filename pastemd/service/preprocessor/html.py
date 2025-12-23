"""HTML content preprocessor."""

from .base import BasePreprocessor
from ...utils.html_formatter import clean_html_content
from ...utils.logging import log


class HtmlPreprocessor(BasePreprocessor):
    """HTML 内容预处理器"""

    def process(self, html: str) -> str:
        """
        预处理 HTML 内容

        处理步骤:
        1. 清理无效元素（SVG等）
        2. 转换删除线标记
        3. 清理 LaTeX 公式块中的 br 标签
        4. 其他自定义处理...

        Args:
            html: 原始 HTML 内容

        Returns:
            预处理后的 HTML 内容
        """
        log("Preprocessing HTML content")

        # 使用 html_formatter 进行清理
        html = clean_html_content(html, self.config)

        # 未来可扩展其他处理...

        return html
