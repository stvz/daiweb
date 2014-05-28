from django.db import models
from django.contrib.auth.models import User

# Create your models here.

class estatus(models.Model):
    """
    Posibles estatus
    """
    estatus_id = models.AutoField(primary_key=True)
    descripcion = models.CharField(max_length=15)
    
    

class tipo(models.Model):
    """
    Catalogo de Tipos o clasificaciones de Tickets
    """
    tipo_id = models.AutoField(primary_key=True)
    descripcion = models.CharField(max_length=25)


class tickets(models.Model):
    """
    Clase para el almacenamiento de los tickets
    """
    ticket_id = models.AutoField(primary_key=True)
    titulo = models.CharField(max_length=50,db_index=True)
    estatus = models.ForeignKey(estatus)
    prioridad = models.SmallIntegerField(max_length=1,db_index=True)
    descripcion = models.TextField()
    tipo = models.ForeignKey(tipo)
    fecha_alta = models.DateField(db_index=True)
    fecha_cierre = models.DateField(db_index=True)
    usuario_alta = models.ForeignKey(User)

class asignacion(models.Model):
    """
    Relacion de los tickets Asignados.
    """
    asignacion_id = models.AutoField(primary_key=True)
    usuario_asigna = models.ForeignKey(User)
    usuario_asignado = models.ForeignKey(User, related_name='+')
    fecha_asignacion = models.DateTimeField(auto_now_add=True)

