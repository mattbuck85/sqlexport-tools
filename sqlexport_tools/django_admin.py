from sqlexport_tools.export_tools import XlsxWriterTool, CsvWriterTool
from django.http import HttpResponse

def extract_qs_fields(qs, fields, caller=None):
    resolved_rows = []
    for row in qs:
        resolved_row = []
        for field in fields:
            try:
                row_field = getattr(row, field)
                if callable(row_field):
                    resolved_row.append(row_field())
                elif hasattr(row_field,'__unicode__'):
                    resolved_row.append(row_field.__unicode__())
                elif hasattr(row_field,'__str__'):
                    resolved_row.append(row_field.__str__())
                else:
                    resolved_row.append(row_field)
            except AttributeError:#Check the caller for an attribute, IE modeladmin
                resolved_row.append(str(getattr(caller, field)(row)))
        resolved_rows.append(resolved_row)
    return (fields, resolved_rows)
    

class AdminExportMixin(object):

    export_method = 'export_excel'
    export_date_format = 'mm/dd/yyyy'
    lookup_final_keywords = ('gte','gt','exact','lte','lt','contains',
                            'icontains','search')


    def __resolve_filters_to_model(self, obj, filter_chain, key, value):
        related_class = obj.__class__
        final_lookup = filter_chain.pop(-1)
        if final_lookup in self.lookup_final_keywords:
            final_keyword = filter_chain.pop(-1)
            final_lookup = "%s__%s" % (final_keyword, final_lookup)
            filter_chain.append(final_keyword)
        else:
            filter_chain.append(final_lookup)
        for i,chain in enumerate(filter_chain):
            obj = getattr(obj, chain)
            if (hasattr(obj, 'pk')):
                related_class = obj.__class__
                if (i == len(filter_chain)-1):
                    return unicode(obj)
            elif (i == len(filter_chain)-1):
                try:
                    related_obj = related_class.objects.get(**{ final_lookup:value })
                    return unicode(related_obj)
                except (related_class.DoesNotExist, related_class.MultipleObjectsReturned):
                    return '%s %s' % (final_lookup, value)

    #Give the user a friendly spreadsheet name by resolving the filters
    def get_spreadsheet_name(self, request, queryset, extension='.xlsx'):
        obj_name = ""
        for _filter in self.list_filter:
            for key,value in request.GET.iteritems():                     
                if _filter in key:
                    instance = queryset[0]
                    filter_chain = key.split('__')
                    if len(filter_chain) == 1:
                        obj_str = filter_chain[0] + '_' + str(getattr(instance, filter_chain[0]))
                    else:
                        obj_str = self.__resolve_filters_to_model(instance, filter_chain, key, value)
                    _obj_name = '_'.join([ name for name in obj_str.lower().split() ])
                    obj_name = "%s%s%s" % (obj_name,"__",_obj_name)
        return "%s%s%s" % (self.model.__name__, obj_name, extension)

    def export(self, request, queryset, tool, content_type, extension='.xlsx'):
        try:
            from StringIO import BytesIO
        except ImportError:
            from io import BytesIO #Python3 support
        output = BytesIO()
        writer = tool(output, default_date_format=self.export_date_format, in_memory=True)
        writer.perform(*self.resolve_list_fields(queryset), sheet_name='default')
        response = HttpResponse(output.getvalue(), content_type)
        filename = self.get_spreadsheet_name(request, queryset, extension)
        response['Content-Disposition'] = 'attachment; filename=%s' % filename
        writer.close()
        return response

    def export_excel(self, request, queryset):
        tool = XlsxWriterTool
        return self.export(request, queryset, tool, 'application/vnd.ms_excel')
    export_excel.short_description = "Export selected to Excel"

    def export_csv(self, request, queryset):
        tool = CsvWriterTool
        return self.export(request, queryset, tool, 'text/csv', '.csv')
    export_csv.short_description = "Export selected to CSV"

    def resolve_list_fields(self,qs):
        fields = list(self.list_display)
        return extract_qs_fields(qs, fields, self)
