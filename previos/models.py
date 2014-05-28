from django.db import models
from webportal.models import cat000usuario
from utils.utilerias import nombre_aleatorio

# Create your models here.


class enc004previo(models.Model):
	referencia = models.CharField(max_length=20, db_index=True)
	fecha = models.DateTimeField(auto_now= True, db_index=True)
	modificacion = models.DateTimeField(auto_now=True)
	usuario = models.ForeignKey(cat000usuario,null=True, blank=True)
	
	def __unicode__(self):
		return '%s'%self.referencia
	
class det004previo(models.Model):
	directorio_ref_ = 'previos/%s/%s'
	
	def _get_directorio(self,referencia):
		return self.directorio_ref_ % (nombre_aleatorio(),referencia)
	
	previo = models.ForeignKey(enc004previo)
	archivo = models.FileField(upload_to=_get_directorio)
	nombre_original = models.CharField(max_length=50)
	tipo = models.CharField(max_length = 6 )
	thumbnail = models.FileField(upload_to=_get_directorio, null=True)
	nuevo_nombre = models.CharField(max_length=50)
	fecha_creacion = models.DateTimeField(auto_now=True, null=True, blank = True)
	fecha_carga = models.DateTimeField(auto_now=True, db_index=True)
	
	def __unicode__(self):
		return u'%(referencia)s-%(nombre)s'%{'referencia':self.previo.referencia,'nombre':self.nombre_original}

