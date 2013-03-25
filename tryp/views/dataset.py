#from pyramid.view import view_config
#from tryp.lib.engine.ppandas import *

#@view_config(route_name='dataset-add', renderer='tryp:templates/dataset.mako')
#def dataset_add(request):
#    if request.method == 'POST':
#        datasource = request.POST['datasource']
#        query = request.POST['query']
#        return {'df_html': ReportEngine().pandas_df(datasource, query)}
#    return {'df_html':'FUCK YOU LAH'}
