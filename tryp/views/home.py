
from pyramid.view import view_config
from tryp.lib.engine.ppandas import *

@view_config(route_name='home', renderer='tryp:templates/home.mako')
def home(request):
    return {}
