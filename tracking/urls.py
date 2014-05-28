
from django.conf.urls import patterns, include, url
from django.core.urlresolvers import reverse
from tracking.views import ImportaFactura, load_factura

urlpatterns = patterns('',
    url(r'^importa_factura/$',ImportaFactura.as_view(), name='utilerias_importa_factura'),
    url(r'^carga_factura/$','load_factura', name = 'utilerias_carga_factura'), 
)