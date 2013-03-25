import formencode
from pyramid.view import view_config
from pyramid_simpleform import Form
from pyramid_simpleform.renderers import FormRenderer

from tryp.lib.engine.ppandas import *

class ReportSchema(formencode.Schema):
    allow_extra_fields = True
    filename = formencode.validators.String(not_empty=True)

@view_config(route_name='report-new', renderer='tryp:templates/report.mako')
def report_new(request):

    form = Form(request, schema=ReportSchema)
    if form.validate():
        filename = request.POST['filename']
        f = open(filename + '.tryp', 'w')
        f.close()
    return {'form':FormRenderer(form)}
