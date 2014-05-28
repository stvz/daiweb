from django.db import models
from django.contrib.auth.models import User

# Create your models here.



class actividades(models.Model):
    '''
     Clase que almace las diferentes actividades que se realizan para llevar a cabo
     un despacho
    '''
    actividad_id = models.AutoField(primary_key=True)
    nombre = models.CharField(max_length=50)
    descripcion = models.TextField(null=True, blank=True)

#Agregando caracteristicas a los usuarios para llevar el control de actividades

#User.add_to_class('iniciales', models.CharField(max_length=8))
#User.add_to_class('actividad',models.ForeignKey(actividades, null=True, blank = True))

class kardex(models.Model):
    '''
    Relacion de los 
    '''
    kardex_id = models.AutoField(primary_key=True)
    referencia = models

    