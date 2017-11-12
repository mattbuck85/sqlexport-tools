# sqlexport-tools
This module provides functionality to export data to excel by passing a cursor into a DatabaseExport class, like so:

```python
  import MySQLdb
  from sqlexport_tools.export_tools import DatabaseExport,XlsxWriterTool
  mysql_connection = MySQLdb.connect(host='localhost',user='myuser',passwd='mypass',db='mydb',use_unicode=True)
  exporter1 = DatabaseExport(mysql_connection.cursor(),'mytable')
  exporter2 = DatabaseExport(mysql_connection.cursor(),'myothertable')
  writer = XlsxWriterTool(filename='export.xlsx',default_date_format='mm/dd/yy')
  exporter.export(writer,sheet_name='MyTable')
  exporter2.export(writer2,sheet_name='MyOtherTable')
    
```
This will dump 'mytable' and 'myothertable' to an xlsx file with a mm/dd/yy date format.  The DatabaseExporter constructor also takes a paramter custom_sql where you can use a custom sql query.

I think the biggest feature here is that this is integrated with the Django Admin Site as a Mixin.  Behold:

```python
from sqlexport_tools.django_admin import AdminExportMixin

class QuestionAdmin(admin.ModelAdmin,AdminExportMixin):
    model = models.Question
    list_display = ('question_text','pub_date')
    export_date_format = 'mm/dd/yy'
    actions = ['export_excel','export_csv']

admin.site.register(models.Question,QuestionAdmin)

```

This enables two actions: export_excel and export_csv.  For excel, the export_date_format can be specified as a parameter of the class.
