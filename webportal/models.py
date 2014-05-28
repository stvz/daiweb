from django.db import models
from django.contrib.auth.models import User

class cat000usuario(models.Model):
	usuario = models.OneToOneField(User)
	departamento = models.CharField(max_length=20)
	
	def __unicode__(self):
		return self.usuario.username