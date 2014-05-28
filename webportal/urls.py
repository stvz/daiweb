from django.conf.urls import patterns, include, url
from .views import Registro

urlpatterns = patterns('',
	#url para iniciar sesion, tambien funciona como index
	url(r'^$', 'django.contrib.auth.views.login' , {'template_name':'index.html'}, name='entrar'),
	# url para cerrar sesion, redirecciona al index
	url(r'^cerrar/$', 'django.contrib.auth.views.logout'  , {'template_name':'index.html'}, name='salir'),
	# url para el registro de nuevos usuarios
	url(r'^registro/$',Registro.as_view(), name='registro')
)