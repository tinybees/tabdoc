#!/usr/bin/env python3
# coding=utf-8

"""
@author: guoyanfeng
@software: PyCharm
@time: 19-2-11 下午6:14
"""
from collections import MutableMapping, Sequence

import tablib
from path import Path

__all__ = ("ExcelWriter",)


class ExcelWriter(object):
    """
    excel book writer
    """

    def __init__(self, excel_name, excel_path=None):
        """
            excel book writer
        Args:
            excel_path: excel path
            excel_name: excel 名称
        """
        self.excel_path = excel_path
        self.excel_name = f"{excel_name}.xls"
        self.excel_book = tablib.Databook()

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
        return tuple(row)

    def add_sheet(self, sheet_name, sheet_data: list):
        """
        为excel添加工作表
        Args:
            sheet_name: 工作表的名称
            sheet_data: 工作表的数据， 必须是列表中嵌套元祖、列表或者字典（从records查询出来的数据库的数据）
        Returns:

        """
        excel_sheet = tablib.Dataset(title=sheet_name)

        for row in sheet_data:
            if not isinstance(row, (MutableMapping, Sequence)):
                raise ValueError("sheet_data值数据类型错误,请检查")

        # 处理list或者tuple个别长度不一致的情况
        first = sheet_data[0]
        if isinstance(first, Sequence):
            for index, row in enumerate(sheet_data[1:], 1):
                diff = len(row) - len(first)
                if abs(diff) > 0:
                    if isinstance(row, list):
                        row.extend(["" for _ in range(diff)])
                    else:
                        sheet_data[index] = (*row, *["" for _ in range(diff)])

        if isinstance(first, (MutableMapping,)):
            excel_sheet.headers = list(first.keys())
            for row in sheet_data:
                row = self._reduce_datetimes(row.values())
                excel_sheet.append(row)
        else:
            excel_sheet.headers = first
            for row in sheet_data[1:]:
                row = self._reduce_datetimes(row)
                excel_sheet.append(row)

        self.excel_book.add_sheet(excel_sheet)

    def save(self, ):
        """
        保存工作簿
        Args:
        Returns:

        """
        if self.excel_path is None:
            file_path = self.excel_name
        else:
            file_path = Path(self.excel_path).joinpath(self.excel_name).abspath()

        with open(file_path, "wb") as f:
            f.write(self.excel_book.export("xls"))
