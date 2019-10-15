#!/usr/bin/env python3
# coding=utf-8

"""
@author: guoyanfeng
@software: PyCharm
@time: 19-7-31 下午2:16
"""
from tabdoc import PDFWriter

if __name__ == '__main__':
    data1 = [['基础基213中学教学班数、班额情况 ', '', '', '', '', '', '', '', '', '', '', '', ' 单位：个'],
             ['', '', '编号', '合计', '初中', '', '', '', '', '高中', '', '', ''],
             ['', '', '', '', '计', '一年级', '二年级', '三年级', '四年级', '计', '一年级', '二年级', '三年级'],
             ['甲', '', '乙', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'],
             ['总计', '', '01', '', '', '', '', '', '', '', '', '', ''],
             ['其中：四年制初中', '', '02', '', '', None, None, None, None, '', '', '', ''],
             ['班\n额\n', '25人及以下', '03', '', '', None, None, None, None, '', None, None, None],
             ['', '26-30人', '04', '', '', None, None, None, None, '', None, None, None],
             ['', '31-35人', '05', '', '', None, None, None, None, '', None, None, None],
             ['', '36-40人', '06', '', '', None, None, None, None, '', None, None, None],
             ['', '41-45人', '07', '', '', None, None, None, None, '', None, None, None],
             ['', '46-50人', '08', '', '', None, None, None, None, '', None, None, None],
             ['', '51-55人', '09', '', '', None, None, None, None, '', None, None, None],
             ['', '56-60人', '10', '', '', None, None, None, None, '', None, None, None],
             ['', '61-65人', '11', '', '', None, None, None, None, '', None, None, None],
             ['', '66人及以上', '12', '', '', None, None, None, None, '', None, None, None]]
    data2 = [
        ["6", None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
         None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
         None, None, None, None],
        ['学校（机构）标识码', None, None, None, None, None, None, None, None, None, '学校（机构）名称（章）', None, None, None,
         None, None, None, None, None, None, None, None, None, None, '学校（机构）英文名称', None, None, None, None, None,
         None, None, None, None, None, None, None]]

    with PDFWriter("test", title="hahha", water_mark="测试水印") as pdf:
        pdf.add_heading("第一个标题", level=3)
        pdf.add_paragraph("这是一个测试文档")
        pdf.add_paragraph("这是一个测试文档", alignment="right")
        for data in [data1, data2]:
            pdf.add_table(data, table_name="测试表格标题 001")
