# Create your views here.

from django.http import HttpResponse
from django.shortcuts import render
from django.template import RequestContext, loader
from django.views.generic import TemplateView
from utils.samsung_refacciones import Importa_Factura
from utils.utilerias import nombre_aleatorio
from django.views.decorators.csrf import csrf_protect
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
import json, os
from django.views.generic import CreateView, TemplateView


class ImportaFactura(TemplateView):
    template_name= 'tracking/importa_factura.html'
#    
#    
#    return render(request, )



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