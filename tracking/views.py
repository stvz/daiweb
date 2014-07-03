# -*- coding: utf-8 -*-
# Create your views here.

from django.http import HttpResponse, StreamingHttpResponse
from django.shortcuts import render
from django.template import RequestContext, loader
from utils.samsung_refacciones import Importa_Factura
from utils.documentacion_samsung import GLP, Mercancias
from utils.utilerias import nombre_aleatorio
from django.views.decorators.csrf import csrf_protect, csrf_exempt, requires_csrf_token
from django.contrib.auth.decorators import login_required
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
import json, os
from django.views.generic import CreateView, TemplateView
from django.views.generic.detail import BaseDetailView, SingleObjectTemplateResponseMixin
from django.template.loader import render_to_string
from utils.reportes_extranet import referencias_vivas, pagos_hechos_referencia
import  daiweb.settings as conf
import datetime
import re


class JSONResponseMixin(object):
    """
    A mixin that can be used to render a JSON response.
    """
    def render_to_json_response(self, context, **response_kwargs):
        """
        Returns a JSON response, transforming 'context' to make the payload.
        """
        return HttpResponse(
            self.convert_context_to_json(context),
            content_type='application/json',
            **response_kwargs
        )

    def convert_context_to_json(self, context):
        "Convert the context dictionary into a JSON object"
        # Note: This is *EXTREMELY* naive; in reality, you'll need
        # to do much more complex handling to ensure that arbitrary
        # objects -- such as Django model instances or querysets
        # -- can be serialized as JSON.
        return json.dumps(context)

class JSONView(JSONResponseMixin, TemplateView):
    def render_to_response(self, context, **response_kwargs):
        return self.render_to_json_response(context, **response_kwargs)
    
class JSONDetailView(JSONResponseMixin, BaseDetailView):
    def render_to_response(self, context, **response_kwargs):
        return self.render_to_json_response(context, **response_kwargs)

class ProcesaArchivos(TemplateView):
    template_name= None
    def get(self, request):
        print 'En el request'
        return render_to_response(self.template_name, data,
            context_instance=RequestContext(request))
    def post(self,request):
        print 'En el post'
        return render_to_response(self.template_name, data,
            context_instance=RequestContext(request))
    
    def dispatch(self, *args, **kwargs):
        print 'Estoy despachando la peticion'
        return super(AjaxGeneral, self).dispatch(*args, **kwargs)

class ImportaFactura(TemplateView):
    template_name= 'tracking/importa_factura.html'

    
#    
#    
#    return render(request, )

class Calculo_impuestos(TemplateView):
    template_name= 'tracking/calculo_impuestos.html'

class ReporteVivas (TemplateView):
    template_name='reportes/reporte_vivas.html'

class PagosHechosReferencia(TemplateView):
    template_name= 'reportes/reporte_pagos_hechos_referencia.html'

class LayoutAbbCfdi(TemplateView):
    template_name= 'reportes/reporte_layout_abb_cfdi.html'
    


@csrf_protect
@login_required
def reporte_vivas(request):
    reporte_ = referencias_vivas.Reporte_vivas('%s'%os.path.join(conf.BASE_DIR,conf.MEDIA_ROOT))
    xlsx_= reporte_.genera_xlsx()
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    xlsx_.save(response)
    response['Content-Disposition'] = 'attachment; filename="Referencias Vivas al %s.xlsx"'%datetime.date.today().isoformat()
    return response

# Reporte de Pagos Hechos por Referencia
# por Manuel Alejandro Estevez Fernandez
#   Junio 2014
@login_required
@csrf_protect
def reporte_pagos_hechos_referencia(request):
    
    if request.method =='POST':
        refs_ = re.compile('DAI[0-9]{2}-[0-9]{4,5}[A-Z]*')
        
        referencias_ = refs_.findall(request.POST.get('referencias').upper())
        honorarios_ = request.POST.get('honorarios')
        ruta_ = os.path.join(conf.MEDIA_ROOT,'temp','pagos_hechos')
        reporte_ = pagos_hechos_referencia.Pagos_hechos(ruta_,referencias_, honorarios_)
        ruta_xlsx_= '%stemp/pagos_hechos/%s'%(conf.MEDIA_URL,reporte_.genera_xlsx())
        return HttpResponse(json.dumps({'archivo':ruta_xlsx_}),content_type="application/json")
    
        
    return response

@login_required
@csrf_protect
def procesa_archivo(request):
    
    archivos_ = []
    respuesta_ = {'success':False}
    
    if request.method == 'POST':
        if len(request.FILES) >0 :
            for archivo_ in request.FILES.keys():
                #print request.FILES[archivo_]
                #print dir(request.FILES[archivo_])
                #print request.FILES[archivo_].name
                try:
                    archivos_.append(default_storage.save(os.path.join('calculo_impuestos',request.FILES[archivo_].name),ContentFile(request.FILES[archivo_].read())))
                except Exception as e:
                    print e
        # En caso que el archivo recibido sea el que contenga las mercancias
        
        if request.POST.get('tipo') == 'mercancia':
            
            # se procede a extraer la informacion.
            respuesta_['analizando_info'] = 'mercancia'
            respuesta_['informacion'] = []
            for archivo_ in archivos_:
                mercancia_ = Mercancias(os.path.join(default_storage.location,archivo_))
                #print mercancia_.carga_archivo()
                mensaje_, valores_ = mercancia_.carga_archivo()
                #print mensaje_
                #print
                #if not mensaje_:
                respuesta_['informacion'] = respuesta_['informacion']  + valores_
                #else:
                #    respuesta_['mensaje']= mensaje_
                #    break
            
            #if not respuesta_.has_key('mensaje'):
            respuesta_['success'] = True
            #else:
            #    respuesta_['success'] = False
            
            
            # Preparamos la respuesta para que arme una tabla con los datos resultantes.
        elif request.POST.get('tipo') == 'glp':
            respuesta_['analizando_info'] = 'glp'
            respuesta_['informacion'] = []
            
            for archivo_ in archivos_:
                glp_ = GLP(os.path.join(default_storage.location,archivo_))
                #print glp_.carga_archivo()
                respuesta_['informacion'] = respuesta_['informacion'] +glp_.carga_archivo()
                
            respuesta_['success'] = True
        
        # Una vez terminado el proceso se borran los archivos temporales
            # para no ocupar demasiado espacio.
        #for elemento_ in archivos_:
        #    print elemento_
        #    os.remove(os.path.join(default_storage.location,elemento_))
        
        
    
    return HttpResponse(json.dumps(respuesta_),content_type="application/json")
    
@login_required
@csrf_protect 
def load_factura(request):
    archivo_ = request.FILES['layout_factura']
    try:
        destino_ = default_storage.save('%s_%s'%(nombre_aleatorio(),archivo_._get_name()), ContentFile(archivo_.read()))
    except Exception as e:
        print e
    
    print destino_
    
    factura_ = Importa_Factura(os.path.join(default_storage.location,destino_),
                _clave_proveedor = request.POST['clave_proveedor'],_clave_cliente = request.POST['clave_cliente']
                , _patente = request.POST['patente'])
    factura_.serializa_xls()
    #print factura_.verifica_factura()
    factura_.verifica_mercancia()
    #print factura_.get_estructura()
    
    os.remove(os.path.join(default_storage.location,destino_))
    return HttpResponse(json.dumps(factura_.get_estructura()),content_type="application/json")