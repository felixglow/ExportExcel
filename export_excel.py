# -*-coding:utf-8 -*-
# 
# Created on 2016-03-15, by felix
# 

__author__ = 'felix'

import datetime
import xlwt
import cStringIO

from django.db.models.query import QuerySet, ValuesQuerySet
from django.http import HttpResponse


class ExcelResponse(HttpResponse):
    """
    excel文件导出
    支持xls和csv格式文件
    支持多sheet页导出
    """

    def __init__(self, origin_data, output_name='excel_data', headers=None, encoding='utf8', is_template=False,
                 sheet_name=None):
        valid_data = False
        if is_template:     # 如果是下载excel模板，传入表格头
            sheet_data = [headers]
        else:
            sheet_data = []
            for n, obj in enumerate(origin_data):
                tmp_data = []
                if obj:
                    if isinstance(obj, ValuesQuerySet):
                        tmp_data = list(obj)
                    elif isinstance(obj, QuerySet):
                        tmp_data = list(obj.values())
                    if hasattr(obj, '__getitem__'):
                        if isinstance(obj[0], dict):
                            if headers[n] is None:
                                headers = obj[0].keys()
                            tmp_data = [[row.get(col, '') for col in headers[n]] for row in obj]
                            tmp_data.insert(0, headers[n])
                        if hasattr(obj[0], '__getitem__'):
                            valid_data = True
                    assert valid_data is True, "ExcelResponse requires a sequence of sequences"
                else:
                    tmp_data.insert(0, headers[n])
                sheet_data.append(tmp_data)

        output = cStringIO.StringIO()
        for n, data in enumerate(sheet_data):
            if len(data) < 65536:
                book = xlwt.Workbook(encoding=encoding)
                sheet = book.add_sheet(sheet_name[n] if sheet_name else 'Sheet ' + str(n+1))
                styles = {'datetime': xlwt.easyxf(num_format_str='yyyy-mm-dd hh:mm:ss'),
                          'date': xlwt.easyxf(num_format_str='yyyy-mm-dd'),
                          'time': xlwt.easyxf(num_format_str='hh:mm:ss'),
                          'default': xlwt.XFStyle}
                for rowx, row in enumerate(data):
                    for colx, value in enumerate(row):
                        if isinstance(value, datetime.datetime):
                            cell_style = styles['datetime']
                        elif isinstance(value, datetime.date):
                            cell_style = styles['date']
                        elif isinstance(value, datetime.time):
                            cell_style = styles['time']
                        else:
                            cell_style = styles['default']
                        sheet.write(rowx, colx, value, style=cell_style)
                book.save(output)
                content_type = 'application/vnd.ms-excel'
                file_ext = 'xls'
            else:
                for row in data:
                    out_row = []
                    for value in row:
                        if not isinstance(value, basestring):
                            value = unicode(value)
                        value = value.encode('gbk', 'ignore')   # 转为gbk便于execl展示
                        out_row.append(value.replace('"', '""'))
                    output.write('"%s"\n' % '","'.join(out_row))
                content_type = 'text/csv'
                file_ext = 'csv'
        output.seek(0)

        super(ExcelResponse, self).__init__(content=output.getvalue(), content_type=content_type)
        self['Content-Disposition'] = 'attachment;filename="%s.%s"' % (output_name.replace('"', '\"'), file_ext)
