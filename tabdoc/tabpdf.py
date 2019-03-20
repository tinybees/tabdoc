#!/usr/bin/env python3
# coding=utf-8

"""
@author: guoyanfeng
@software: PyCharm
@time: 19-2-11 下午6:14
"""

__all__ = ("PDFWriter",)


class PDFWriter(object):
    """
    pdf book writer
    """

    def __init__(self, pdf_name):
        """
            excel book writer
        Args:
            pdf_name: pdf 名称
        """
        self.excel_name = f"{pdf_name}.xls"
