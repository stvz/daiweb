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
    
    # importa archivos
    url(r'^importa_archivos/$',ImportaArchivo.as_view(), name='utilerias_importa_archivo'),
    url(r'^carga_archivo/$',carga_archivo, name='utilerias_carga_archivo'),
    
    
    # Reportes
    # vivas
    url(r'^reporte_vivas/$',ReporteVivas.as_view(), name='reporte_vivas'),
    url(r'^get_reporte_vivas/$',reporte_vivas, name='reportes_get_reporte_vivas'),
    #pagos hechos por referencia
    url(r'^reporte_pagos_hechos_referencia/$',PagosHechosReferencia.as_view(), name='reporte_pagos_hechos_referencia'),
    url(r'^get_reporte_pagos_hechos_referencia/$',reporte_pagos_hechos_referencia, name='reportes_get_pagos_hechos_referencia'),
    # Layout ABB CFDIS
    url(r'^reporte_layout_abb_cdfis/$',LayoutAbbCfdi.as_view() ,name='reporte_layout_abb_cfdis'),
    url(r'^get_reporte_layout_abb_cdfis/$',reporte_layout_abb_cdfis ,name='reporte_get_layout_abb_cfdis'),
    
    #Utilerias Varias
    url(r'getReferencia/$',getReferencia, name='getReferencia_zego'),
    url(r'getProveedor/$',getProveedor, name='getProveedor_zego'),
    url(r'getCliente/$',getCliente, name='getCliente_zego'),
    
    url(r'auditoriaPedimento/$',AuditoriaFactura.as_view(), name='auditoria_pedimento'),
    url(r'auditaPedimento/$',audita_pedimento,name='audita_pedimento'),
    
)