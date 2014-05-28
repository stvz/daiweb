from django.conf.urls import patterns, include, url
from django.core.urlresolvers import reverse, reverse_lazy
from .views import PrevioAgregar, PrevioReporte, PrevioBuscar, PrevioRevisar 

urlpatterns = patterns('',
	url(r'^nuevo/$',PrevioAgregar.as_view(), name='previo_agregar'),
	url(r'^inicio/$',PrevioReporte.as_view(), name='previo_reporte'),
	url(r'^buscar/$',PrevioBuscar.as_view(), name='previo_buscar'),
	url(r'^revisar/$',PrevioRevisar.as_view(), name='previo_revisar'),
	#url(r'^carga_factura','load_factura'), 
)