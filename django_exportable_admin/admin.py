from django.contrib import admin
from django.conf.urls.defaults import patterns, url
from django.template.defaultfilters import slugify
from django.core.urlresolvers import reverse
from django.http import HttpResponse

import xlwt


class ExportableAdmin(admin.ModelAdmin):
    """
    Base class which contains all of the functionality to setup admin
    changelist export. Subclassing this class itself will do nothing unless you
    set export_formats on your ModelAdmin instance. See the other provided
    subclasses which are already setup for CSV, Pipe, and both.
    
    Note: do not override change_list_template or you will not get the
    "Export ..." button on your changelist page.
    """
    # use a custom changelist template which adds "Export ..." button(s)
    change_list_template = 'django_exportable_admin/change_list_exportable.html'

    # export 10,000 results by default
    export_queryset_limit = 10000
    
    # an iterable of 2-tuples of (format-name, format-delimiter), such as:
    #  ((u'CSV', u','), (u'Pipe', u'|'),)
    export_formats = tuple()

    # an iterable of strings defining export types, such as:
    #  ('xls',)
    export_types = tuple()

    def get_paginator(self, request, queryset, per_page, orphans=0, allow_empty_first_page=True):
        """
        When we are exporting, modify the paginator to set the result limit to
        'export_queryset_limit'.
        """
        if hasattr(request, 'is_export_request'):
            return self.paginator(queryset, self.export_queryset_limit, 0, True)
        return self.paginator(queryset, per_page, orphans, allow_empty_first_page)

    def get_export_buttons(self, request):
        """
        Returns a iterable of 2-tuples which contain:
            (button text, link URL)

        These will be used in the customized changelist template to output a
        button for each export format.
        """
        app, mod = self.model._meta.app_label, self.model._meta.module_name
        buttons = [
                ('Export %s' % format_name,
                    reverse("admin:%s_%s_export_%s" % (app, mod,
                        format_name.lower())))
                for format_name, delimiter in self.export_formats
                ]
        buttons += [
                ('Export %s' % type_name.upper(),
                    reverse("admin:%s_%s_export_%s" % (app, mod,
                        type_name.lower())))
                for type_name in self.export_types
                ]
        return buttons

    def changelist_view(self, request, extra_context=None):
        """
        After 1.3, the changelist view returns a TemplateResponse, which we can
        use to greatly simplify this class. Instead of having to redefine a
        copy of the changelist_view to alter the template, we can simple change
        it after we get the TemplateResponse back.
        """
        if extra_context and 'export_delimiter' in extra_context:
            # set this attr for get_paginator()
            request.is_export_request = True
            response = super(ExportableAdmin, self).changelist_view(request, extra_context)
            # response is a TemplateResponse so we can change the template
            response.template_name = 'django_exportable_admin/change_list_csv.html'
            response['Content-Type'] = 'text/csv'
            response['Content-Disposition'] = 'attachment; filename=%s.csv' % slugify(self.model._meta.verbose_name)
            return response

        elif extra_context and 'export_type' in extra_context:
            # Generate response from super
            request.is_export_request = True
            response = super(ExportableAdmin, self).changelist_view(request, extra_context)

            # Call appropriate method to handle export_type
            method_name = "export_type_%s" % extra_context['export_type']
            method = getattr(self, method_name)
            return method(request, response, extra_context)

        extra_context = extra_context or {}
        extra_context.update({
            'export_buttons' : self.get_export_buttons(request),
        })
        return super(ExportableAdmin, self).changelist_view(request, extra_context)

    # Adapted from https://gist.github.com/1560240
    def export_type_xls(self, request, response, extra_context=None):
        meta = self.model._meta
        modeladmin = self
        filename = '%s.xls' % meta.verbose_name_plural.lower()
        # Get filtered queryset from changelist view
        queryset = response.context_data['cl'].query_set

        print "Got queryset for export.. total count: ", queryset.count()

        def get_verbose_name(fieldname):
            name = filter(lambda x: x.name == fieldname, meta.fields)
            if name:
                return (name[0].verbose_name or name[0].name).upper()
            return fieldname.upper()

        response = HttpResponse(mimetype='application/ms-excel')
        response['Content-Disposition'] = 'attachment;filename=%s' % filename

        wbk = xlwt.Workbook()
        sht = wbk.add_sheet(meta.verbose_name_plural)

        for j, fieldname in enumerate(modeladmin.list_display[1:]):
            sht.write(0, j, get_verbose_name(fieldname))

        for i, row in enumerate(queryset):
            for j, fieldname in enumerate(modeladmin.list_display[1:]):
                sht.write(i + 1, j, unicode(getattr(row, fieldname, '')))

        wbk.save(response)
        return response

    def get_urls(self):
        """
        Add URL patterns for the export formats. Really all these URLs do are 
        set extra_context to contain the export_delimiter for the template
        which actually generates the "CSV".
        """
        urls = super(ExportableAdmin, self).get_urls()
        app, mod = self.model._meta.app_label, self.model._meta.module_name
        # make a URL pattern for each export format
        new_urls = [
            url(
                r'^export/%s$' % format_name.lower(),
                self.admin_site.admin_view(self.changelist_view),
                name="%s_%s_export_%s" % (app, mod, format_name.lower()),
                kwargs={'extra_context':{'export_delimiter':delimiter}},
            )
            for format_name, delimiter in self.export_formats
        ]
        for export_type in self.export_types:
            new_urls.append(url(
                r'^export/%s$' % export_type.lower(),
                self.admin_site.admin_view(self.changelist_view),
                name="%s_%s_export_%s" % (app, mod, export_type.lower()),
                kwargs={'extra_context':{'export_type':export_type.lower()}},
                ))
        my_urls = patterns('', *new_urls)
        return my_urls + urls

class XLSExportableAdmin(ExportableAdmin):
    """
    ExportableAdmin subclass which adds export to XLS functionality.
    """
    export_types = (
            'xls',
    )

class CSVExportableAdmin(ExportableAdmin):
    """
    ExportableAdmin subclass which adds export to CSV functionality.
    
    Note: do not override change_list_template or you will not get the
    "Export ..." button on your changelist page.
    """
    export_formats = (
        (u'CSV', u','),
    )

class PipeExportableAdmin(ExportableAdmin):
    """
    ExportableAdmin subclass which adds export to Pipe functionality.
    
    Note: do not override change_list_template or you will not get the
    "Export ..." button on your changelist page.
    """
    export_formats = (
        (u'Pipe', u'|'),
    )

class MultiExportableAdmin(ExportableAdmin):
    """
    ExportableAdmin subclass which adds export to CSV and Pipe
    functionality.
    
    Note: do not override change_list_template or you will not get the
    "Export ..." buttons on your changelist page.
    """
    export_formats = (
        (u'CSV', u','),
        (u'Pipe', u'|'),
    )
