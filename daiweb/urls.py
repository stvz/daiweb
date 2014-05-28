from django.conf.urls import patterns, include, url
from django.core.urlresolvers import reverse
import settings

from django.contrib import admin
admin.autodiscover()

urlpatterns = patterns('',
    
    

    url(r'^',include('webportal.urls')),
    url(r'^tracking/',include('tracking.urls')),
    url(r'^previos/',include('previos.urls')),
    url(r'^admin/', include(admin.site.urls)),
    #url(r'^tracking/importa_factura','tracking.views.importa_factura'),
    #url(r'^tracking/carga_factura','tracking.views.load_factura'),
    # url(r'^blog/', include('blog.urls')),
    
    (r'^media/(?P<path>.*)$', 'django.views.static.serve', {
        'document_root': settings.MEDIA_ROOT}),
    
)