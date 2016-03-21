# Export Execl
导出execl或csv文件
====================

用途:
----
    Django中Http Response返回execl文件或csv文件
    支持自定义文件名称
    支持生成多sheet页 自定义sheet页名称
    支持导入模板生成

支持:
----
    需要安装xlwt库

使用:
----

```python
from excel_response import ExcelResponse

def excelview(request):
    objs = SomeModel.objects.all()
    return ExcelResponse([objs])
```
or:

```python
from excel_response import ExcelResponse

def excelview(request):
    data = [
    [{u'姓名': 'Tom', u'年龄': 18, u'性别': u'男', u'身高': 175, u'体重': 67},
     {u'姓名': 'Lily', u'年龄': 22, u'性别': u'女', u'身高': 163, u'体重': 41}],
    [{u'姓名': 'Tom', u'身高': 175, u'体重': 67}],
    [{u'姓名': 'Lily', u'身高': 163, u'体重': 41}]
    ]
    headers = [(u'姓名', u'年龄', u'性别', u'身高', u'体重'),
               (u'姓名', u'身高', u'体重'),
               (u'姓名', u'身高', u'体重')]
    sheet_name=[u'总览', u'男生统计', u'女生统计']
    return ExcelResponse(data, output_name=u'班级体检统计', headers=headers, is_template=False, sheet_name=sheet_name)
```