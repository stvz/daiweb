# -*- coding: utf-8 -*-
from django.conf.urls import patterns, include, url
from django.core.urlresolvers import reverse
from tracking.views import *
from django.contrib.auth.decorators import login_required

urlpatterns = patterns('',
    url(r'^importa_factura/$',ImportaFactura.as_view(), name='utilerias_importa_factura'),
    url(r'^carga_factura/$',load_factura, name = 'utilerias_carga_factura'),
    
    #Calculo de impuestos
    url(r'^calculo_impuestos/$',login_required(Calculo_impuestos.as_view()),name='utilerias_calculo_impuestos'),
    #url(r'^procesa_archivo/$',login_required(ProcesaArchivos.as_view()), name='utilerias_procesa_archivos'),
    url(r'^procesa_archivo/$',procesa_archivo , name='utilerias_procesa_archivo'),
    
    
    # Reportes
    url(r'^reporte_vivas/$',ReporteVivas.as_view(), name='reporte_vivas'),
    url(r'^get_reporte_vivas/$',reporte_vivas, name='reportes_get_reporte_vivas'),
)