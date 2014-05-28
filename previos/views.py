from django.views.generic import CreateView, ListView, TemplateView, FormView
from .models import enc004previo, det004previo
from django.core.urlresolvers import reverse, reverse_lazy
from django.shortcuts import render, HttpResponse
from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
from django.core import serializers
from django.core.serializers.json import DjangoJSONEncoder
from types import MethodType
import json, ast

# Create your views here.
class PrevioAgregar(CreateView):
	template_name  = 'previos/registrar_previo.html'
	model = enc004previo
	success_url = reverse_lazy()



class PrevioReporte(TemplateView):
	template_name = 'previos/reporte_previos.html'
	model = enc004previo

class PrevioBuscar(TemplateView):
	
	#def post (self, request, *args, **kwargs ):
	#	buscar = request.POST['referencia']
	#	lista_previos_ = enc004previo.objects.filter(referencia__contains=buscar)
	#	paginacion_ = Paginator(lista_previos_,20) # mostrando 20 a la vez
	#	hoja_ = 1
	#	try:
	#		previos_ = paginacion_.page(hoja_)
	#	except PageNotAnInteger:
	#		previos_ = paginacion_.page(1)
	#	except EmptyPage:
	#		previos_ = paginacion_.page(paginacion_.num_pages)
	#		
	#	return render(request,'previos/resultado_busqueda.html'
	#				  , {'referencias': previos_})
	def get(self, request, *args, **kwargs ):
		buscar = request.GET.get('referencia')
		lista_previos_ = enc004previo.objects.filter(referencia__contains=buscar)
		paginacion_ = Paginator(lista_previos_,20) # mostrando 20 a la vez
		hoja_ = request.GET.get('hoja')
		try:
			previos_ = paginacion_.page(hoja_)
		except PageNotAnInteger:
			previos_ = paginacion_.page(0)
		except EmptyPage:
			previos_ = paginacion_.page(paginacion_.num_pages)
		# En esta seccion se buscan los datos necesarios para la paginacion
		keys_ = ("end_index", "has_next", "has_other_pages", "has_previous",
				 "next_page_number", "number", "start_index", "previous_page_number")
		# se crea un diccionario a partir de las llaves de la tupla
		datos_paginacion_ = {}
		# se recorre el objeto de la paginacion para extraer
		# los valores necesarios
		for attr in keys_:
			v = getattr(previos_, attr)
			if isinstance(v, MethodType):
				try:
					datos_paginacion_[attr] = v()
				except EmptyPage:
					datos_paginacion_[attr] = None
			elif isinstance(v, (str, int)):
				datos_paginacion_[attr] = v
		#definimos un serializador para crear un objeto completo
		pythonserializer = serializers.get_serializer("python")()
		# se utiliza el serializador anterior y se agregan los elementos
		# a mostrar en la pagina
		datos_paginacion_["object_list"] = pythonserializer.serialize(previos_.object_list
											,fields=('id', 'fecha', 'referencia'))
		# se realiza la serializacion a json para su envio
		pag_ = json.dumps(datos_paginacion_, cls=DjangoJSONEncoder)
		# se envia la respuesta a la consulta web
		return HttpResponse(content=pag_, content_type = 'application/json' )

class PrevioRevisar(TemplateView):
	
	def get(self, request, *args, **kwargs ):
		previo_ = enc004previo.objects.get(pk=request.GET.get('id'))
		archivos_ = det004previo.objects.get(previo=previo_)
		return render(request,'previos/revisar_previo.html'
					, {'previo':previo_ ,'archivos':archivos_, 'n_img':range(len(archivos_))})