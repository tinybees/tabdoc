#!/usr/bin/env python3
# coding=utf-8

"""
@author: guoyanfeng
@software: PyCharm
@time: 19-2-11 下午6:14
"""
from collections import MutableMapping, Sequence

from path import Path
from reportlab.lib import colors
from reportlab.lib.fonts import addMapping
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

__all__ = ("PDFWriter",)

pdfmetrics.registerFont(
    TTFont('simhei', Path(__file__).dirname().joinpath("templates/SimHei.ttf").abspath()))
addMapping('simhei', 0, 0, 'simhei')  # normal
addMapping('simhei', 0, 1, 'simhei_italic')  # italic
addMapping('simhei', 1, 0, 'simhei_bold')  # bold
addMapping('simhei', 1, 1, 'simhei_boldItalic')  # italic and bold


class RotateTable(Table):  # Table Rotate

    def draw(self):
        #  获取当前的基准Y点
        current_y = vars(self.canv)["_currentMatrix"][-1]
        #  这里移动x点以左边界为基准向右移动, 文档的宽度-表格的高度-边界距离
        #  这里移动y点以上边界为基准向上移动,文档的高度-当前y点-边界距离，就是要向上移动的距离
        self.canv.translate(letter[0] - sum(self._rowHeights) - inch, letter[1] - current_y - inch)
        self.canv.rotate(-90)
        super().draw()


class RotateParagraph(Paragraph):  # Table Rotate

    def draw(self):
        self.canv.translate(letter[0] - 2 * inch, 0)
        self.canv.rotate(-90)
        super().draw()


class PDFWriter(object):
    """
    pdf book writer
    """

    def __init__(self, pdf_name, pdf_path=None, water_mark="", title=None):
        """
            excel book writer
        Args:
            pdf_name: pdf 名称
            title: 文件title
            pdf_path: pdf path
            water_mark: pdf 水印文字
        """
        self.story = []
        self.pdf_name = f"{pdf_name}.pdf"
        self.pdf_path = pdf_path
        self.document = SimpleDocTemplate(self.get_full_name(), pagesize=letter)
        self.document.water_mark = water_mark
        self.alignment_map = {"left": 0, "center": 1, "right": 2, "justify": 4}
        if title:
            self.add_heading(title, alignment="center")

    @property
    def styles(self, ):
        """
        获取样式，这里会更改样式的字体，以便于支持中文
        Args:
        """

        styles = getSampleStyleSheet()
        for key, value in styles.byName.items():
            value.fontName = "simhei"
            styles.byName[key] = value
        return styles

    def get_full_name(self, ):
        """
        获取全路径文件名
        Args:

        Returns:

        """
        if self.pdf_path is None:
            full_name = self.pdf_name
        else:
            full_name = Path(self.pdf_path).joinpath(self.pdf_name).abspath()
        return full_name

    def __enter__(self):
        """

        Args:

        Returns:

        """
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """

        Args:

        Returns:

        """
        self.save()

    @staticmethod
    def _reduce_datetimes(row):
        """Receives a row, converts datetimes to strings."""

        row = list(row)

        for i, val in enumerate(row):
            if hasattr(val, "strftime"):
                row[i] = val.strftime("%Y-%m-%d %H:%M:%S")
            elif hasattr(val, 'isoformat'):
                row[i] = val.isoformat()
        return row

    def add_heading(self, head_text: str = None, *, level: int = 1, alignment="left"):
        """
        为PDF文档中添加标题
        Args:
            head_text: 标题的内容
            level: 标题的级别， 共一到六级别
            alignment: 标题的对其方式,left,center,right,justify
        Returns:

        """
        if alignment not in self.alignment_map:
            raise ValueError("alignment必须是left,center,right,justify")
        if level < 1 or level > 6:
            raise ValueError("level必须在1和6之间。")
        styles = self.styles.get(f'Heading{level}')
        styles.alignment = self.alignment_map[alignment]
        self.story.append(Paragraph(head_text, styles))
        self.story.append(Spacer(1, 0.25 * inch))

    def add_paragraph(self, paragraph_text, alignment="left"):
        """
        PDF文档中添加段落
        Args:
            paragraph_text: 段落内容
            alignment: 标题的对其方式,left,center,right,justify
        Returns:

        """

        if alignment not in self.alignment_map:
            raise ValueError("alignment必须是left,center,right,justify")
        styles = self.styles.get('Normal')
        styles.alignment = self.alignment_map[alignment]
        self.story.append(Paragraph(paragraph_text, styles))
        self.story.append(Spacer(1, 0.15 * inch))

    def add_table(self, table_data: list, table_name=None, data_align='CENTER', table_halign='CENTER',
                  is_landscape=False):
        """
        为pdf添加表格数据
        Args:
            table_name: 表格的名称
            table_data: 表格的数据， 必须是列表中嵌套元祖、列表或者字典（从records查询出来的数据库的数据）
            data_align: The alignment of the data inside the table (eg.
                'LEFT', 'CENTER', 'RIGHT')
            table_halign: Horizontal alignment of the table on the page
                (eg. 'LEFT', 'CENTER', 'RIGHT')
            is_landscape: 是否横向展示，默认false
        Returns:

        """
        table_data = table_data if table_data else [[""]]
        self.story.append(PageBreak())
        if table_name:
            styles = self.styles.get('Heading4')
            styles.alignment = self.alignment_map.get(table_halign.lower(), "center")
            if is_landscape:
                self.story.append(RotateParagraph(table_name, styles))
            else:
                self.story.append(Paragraph(table_name, styles))
            # self.story.append(Spacer(1, 0.15 * inch)) # 这里是增加间距，测试后发现去掉更美观点

        for index, row in enumerate(table_data):
            if not isinstance(row, (MutableMapping, Sequence)):
                raise ValueError("table_data值数据类型错误,请检查")
            table_data[index] = row[:36]  # 解决超过36列行高大于一页而报错的问题

        # 处理list或者tuple个别长度不一致的情况
        first = table_data[0]
        if isinstance(first, Sequence):
            for index, row in enumerate(table_data[1:], 1):
                diff = len(row) - len(first)
                if abs(diff) > 0:
                    if isinstance(row, list):
                        row.extend(["" for _ in range(diff)])
                    else:
                        table_data[index] = [*row, *["" for _ in range(diff)]]
                table_data[index] = self._reduce_datetimes(row)

        else:
            for index, row in enumerate(table_data[1:], 1):
                diff = len(row) - len(first)
                if abs(diff) > 0:
                    row_ = [*list(row.values()), *["" for _ in range(diff)]]
                    table_data[index] = self._reduce_datetimes(row_)
                else:
                    table_data[index] = self._reduce_datetimes(row.values())
        cell_styles = self.styles["Normal"]
        for row_index, row in enumerate(table_data):
            for column_index, one_value in enumerate(row):
                table_data[row_index][column_index] = Paragraph(str(one_value) if one_value else "", cell_styles)

        # 第一列的宽度是其他列的两倍,第二列的宽度是其他列的1.5倍
        if is_landscape:
            column_len, column_width_per = len(table_data[-1]), (letter[1] - 2 * inch) / (len(table_data[-1]) + 2)
        else:
            column_len, column_width_per = len(table_data[-1]), self.document.width / (len(table_data[-1]) + 2)

        column_width = [column_width_per * 2, column_width_per * 1.5,
                        *[column_width_per for _ in range(column_len - 2)]]
        if is_landscape:
            table = RotateTable(table_data, hAlign=table_halign, colWidths=column_width)
        else:
            table = Table(table_data, hAlign=table_halign, colWidths=column_width)
        # (列,行) (0, 0)(-1, -1)代表0列0行到所有的单元格
        table.setStyle(TableStyle([('FONT', (0, 0), (-1, -1), 'simhei'),  # 所有单元格设置雅黑字体
                                   ('ALIGN', (0, 0), (-1, 0), 'LEFT'),  # 第一列左对齐
                                   ('ALIGN', (0, 0), (0, 0), data_align),  # 第一个单元格
                                   ('ALIGN', (1, 0), (-1, -1), data_align),  # 第一列到剩下的所有数据
                                   ('INNERGRID', (0, 0), (-1, -1), 0.50, colors.black),
                                   ('BOX', (0, 0), (-1, -1), 0.25, colors.black)]))
        self.story.append(table)

    @staticmethod
    def on_pages_setup(canvas, doc):
        """
        为每页增加水印，或者其他的logo等
        Args:

        Returns:

        """
        canvas.saveState()

        canvas.setFont("simhei", 60)
        canvas.rotate(30)  # 旋转30度
        canvas.setFillAlpha(0.1)  # 设置透明度
        canvas.setFillGray(0.50)  # 设置灰度
        canvas.drawCentredString(6.5 * inch, 3.75 * inch, doc.water_mark)

        canvas.restoreState()

    def save(self, ):
        """
        保存PDF
        Args:
        Returns:

        """
        self.document.build(self.story, onLaterPages=self.on_pages_setup)
