# -*- coding: utf-8 -*-
"""
	Por: Manuel Alejandro Estevez Fernandez
		Noviembre 2013
	
	Modelos con catalogos para homologación de operaciones
	Aduanales
"""

from django.db import models
from django.contrib.auth.models import User


class cat001continentes(models.Model):
	"""
	Catalogo de Continentes
	"""
	continente_id = models.AutoField(primary_key=True)
	nombre_continente = models.CharField(max_length=20)
	borrado = models.BooleanField(default=False, db_index=True)
	
	#class Meta:
	#	verbose_name =_('Continente')
	#	verbose_name_plural =_('Continentes')
	
	class Meta:
		permissions = (
			('can_view_continentes','Can View Continentes'),
			)
	
	def __unicode__(self):
		return '%s'.capitalize()%nombre_continente

class cat001tipo_contenedores(models.Model):
	"""
	Clase que contiene el catalogo de Contenedores
	"""
	tipo_contenedor_id = models.AutoField(primary_key=True)
	descripcion = models.CharField(max_length = 150)
	categoria = models.CharField(max_length=20, null=True, blank = True)
	siglas = models.CharField(max_length=4, null=True, blank = True)
	dimensiones = models.CharField(max_length=4, null=True, blank = True)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_tipos_contenedores','Can View Tipos de Contenedores'),
			)
		
	def __unicode__(self):
		return '%s %s'%(self.descripcion, self.siglas)

class cat001monedas(models.Model):
	"""
	Catalogo de Monedas
	"""
	moneda_id = models.AutoField(primary_key=True)
	nombre_moneda = models.CharField(max_length=20)
	clave = models.CharField(max_length=10,db_index=True)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_monedas','Can View monedas'),
			)
	
	def __unicode__(self):
		return '%s'.capitalize()%self.nombre_moneda

class cat001paises(models.Model):
	"""
	Catalogo de Paises
	"""
	pais_id = models.AutoField(primary_key = True)
	nombre_pais = models.CharField(max_length=100)
	nombre_pais_ingles = models.CharField(max_length=100, null=True)
	codigo_numerico = models.CharField(max_length=5,db_index=True, null=True)
	codigo_2_letras = models.CharField(max_length=2,db_index=True, null=True)
	codigo_3_letras = models.CharField(max_length=3,db_index=True)
	continente = models.ForeignKey(cat001continentes, null=True)
	moneda = models.ForeignKey(cat001monedas, null=True)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_paises','Can View Paises'),
			)
	
	def __unicode__(self):
		return '%s %s'.capitalize()%(self.codigo_3_letras,self.nombre_pais)

class cat001puertos(models.Model):
	"""
	Catalogo de Puertos del Mundo
	"""
	puerto_id = models.AutoField(primary_key = True)
	nombre_puerto = models.CharField(max_length = 150,db_index=True )
	pais = models.ForeignKey(cat001paises)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_puertos','Can View Puertos'),
			)
	
	def __unicode__(self):
		return '%s, %s'%(self.nombre_puerto,self.pais.nombre_pais)

class cat001aduanas(models.Model):
	aduana_id = models.AutoField(primary_key=True)
	clave = models.IntegerField(unique=True)
	descripcion = models.CharField(max_length=150)
	borrado = models.BooleanField(default=False, db_index=True)

	class Meta:
		permissions = (
			('can_view_aduanas','Can View aduanas'),
			)
	
	def __unicode__(self):
		return '%s: %s'%(self.clave,self.descripcion)

class cat001direcciones(models.Model):
	"""
	Catalogo de Direcciones
	"""
	direccion_id = models.AutoField(primary_key=True)
	calle = models.TextField()
	numero_exterior = models.CharField(max_length=15)
	numero_interior = models.CharField(max_length=15)
	colonia = models.CharField(max_length=50)
	municipio = models.CharField(max_length=50)
	estado = models.CharField(max_length= 50)
	codigo_postal = models.CharField(max_length=20)
	pais = models.ForeignKey(cat001paises)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_direcciones','Can View direcciones'),
			)
	
	def __unicode__(self ):
		return '%s %s %s %s'%(self.calle,self.numero_exterior,self.numero_interior,self.colonia)

class cat001patentes(models.Model):
	"""
	Catalogo de Patentes de Agentes Aduanales
	"""
	patente_id = models.AutoField(primary_key=True)
	clave = models.CharField(max_length=4,unique=True)
	agente_aduanal = models.CharField(max_length=150)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_patentes','Can View Patentes'),
			)
	
	def __unicode_(self):
		return '%s %s'%(self.clave,self.agente_aduanal)

class cat001oficinas(models.Model):
	"""
	Catalogo de Oficinas Aduanales
	"""
	oficina_id = models.AutoField(primary_key=True)
	nombre = models.CharField(max_length=150)
	razon_social = models.CharField(max_length= 150)
	clave_oficina = models.CharField(max_length=10)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_oficinas','Can View oficinas'),
			)
	
	def __unicode__(self):
		return  self.nombre

class cat001razones_sociales(models.Model):
	"""
	Catalogo de Razones Sociales (Clientes)
	"""
	razon_social_id = models.AutoField(primary_key=True)
	razon_social = models.CharField(max_length = 255,db_index=True)
	descripcion = models.TextField(null=True, blank = True)
	rfc = models.CharField(max_length=20,db_index=True)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_razones_sociales','Can View Razones Sociales'),
			)
	
	def __unicode__(self):
		return '%s %s'%(self.rfc,self.descripcion)

class cat001navieras(models.Model):
	naviera_id= models.AutoField(primary_key=True)
	nombre_naviera = models.CharField(max_length=150)
	sitio_web = models.TextField(null = True, blank=True)
	borrado = models.BooleanField(default=False, db_index=True)
	
	def __unicode__(self):
		return '%s'.capitalize()%self.nombre_naviera

	class Meta:
		permissions = (
			('can_view_navieras','Can View Navieras'),
			)

class cat001buques(models.Model):
	buque = models.AutoField(primary_key = True)
	nombre_buque = models.CharField(max_length = 150)
	naviera = models.ForeignKey(cat001navieras)
	borrado = models.BooleanField(default=False, db_index=True)
	
	def __unicode__(self):
		return '%s'.capitalize()%self.nombre_buque
	
	class Meta:
		permissions = (
			('can_view_buques','Can View Buques'),
			)

class cat001linea_aerea(models.Model):
	linea_aerea = models.AutoField(primary_key = True)
	nombre = models.CharField(max_length = 150 )
	sitio_web = models.TextField(null = True, blank=True)
	borrado = models.BooleanField(default=False, db_index=True)
	
	def __unicode__(self):
		return '%s'.capitalize()%self.nombre
	
	class Meta:
		permissions = (
			('can_view_linea_aerea','Can View Linea Aerea'),
			)

class cat001impuestos(models.Model):
	"""
	Clase para el almacenamiento de catalogo de contribuciones seg?n anexo 22
	"""
	impuesto_id = models.AutoField(primary_key=True)
	descripcion = models.CharField(max_length=150)
	clave = models.IntegerField()
	abreviacion = models.CharField(max_length=10)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_impuestos','Can View Impuestos'),
			)
	
	def __unicode__(self):
		return '%s'%self.descripcion

class cat001proveedores(models.Model):
	proveedor_id= models.AutoField(primary_key=True)
	nombre = models.CharField(max_length=150)
	identificador_fiscal = models.CharField(max_length=50, null=True, blank = True, db_index=True)
	vinculado = models.SmallIntegerField(null=True, blank = True)
	codigo_samsung = models.CharField(max_length=50, blank= True, null=True)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_proveedores','Can View Proveedores'),
			)

class cat001unidad_medida(models.Model):
	unidad_id = models.AutoField(primary_key=True)
	descripcion= models.CharField(max_length=30)
	clave = models.CharField(max_length=5,null = True, blank=True)
	abreviacion = models.CharField(max_length=10, null=True, blank = True)
	breviacion_ingles = models.CharField(max_length=10, null=True, blank = True)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_unidad_medida','Can View Unidad de Medida'),
			)

class cat001articulos(models.Model):
	"""
	Contiene el catalogo de articulos
	"""
	tipos = (
		('PT','Producto Terminado'),
		('R','Refacciones'),
		('MA','Materia Prima'),
		('ME','Muestras')
	)
	articulo_id = models.AutoField(primary_key=True)
	descripcion = models.CharField(max_length=150)
	fraccion = models.CharField(max_length=8,null = True, blank=True)
	fraccion_8 = models.CharField(max_length=8,null = True, blank=True)
	observaciones= models.TextField(null=True, blank = True)
	tipo_mercancia= models.CharField(max_length=4,choices = tipos, null=True, blank = True)
	unidad = models.ForeignKey(cat001unidad_medida, null = True, blank=True)
	codigo_producto = models.CharField(max_length=150,db_index=True)
	borrado = models.BooleanField(default=False, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_articulos','Can View Articulos'),
			)


#
#   Relaciones
#
class rel001claves_cliente(models.Model):
	clave_id = models.AutoField(primary_key=True)
	numero_clave = models.CharField(max_length=15,db_index=True)
	descripcion = models.CharField(max_length=255)
	razon_social = models.ForeignKey(cat001razones_sociales)
	oficina = models.ForeignKey(cat001oficinas)
	activo = models.BooleanField(default=False,db_index=True)
	borrado = models.BooleanField(default=False, db_index=True)

class rel001direcciones_clientes(models.Model):
	direccion_cliente_id = models.AutoField(primary_key=True)
	direccion = models.ForeignKey(cat001direcciones)
	cliente = models.ForeignKey(rel001claves_cliente)
	activo = models.BooleanField(default=False,db_index=True)
	borrado = models.BooleanField(default=False, db_index=True)

class rel001direcciones_proveedor(models.Model):
	direccion_proveedor_id = models.AutoField(primary_key=True)
	direccion = models.ForeignKey(cat001direcciones)
	proveedor = models.ForeignKey(cat001proveedores)
	activo = models.BooleanField(default=False,db_index=True)
	borrado = models.BooleanField(default=False, db_index=True)
#
#   Registros
#
class reg001operacion(models.Model):
	
	estatus = (
		('C','Cancelado'),
		('A','Abierto'),
		('F','Finalizado')
	)
	
	operacion_id = models.AutoField(primary_key = True )
	oficina = models.ForeignKey(cat001oficinas)
	referencia = models.CharField(max_length=20, null=False,db_index=True)
	tipo = models.CharField(max_length=4)
	clave_pedimento = models.CharField(max_length=2,db_index=True)
	aduana = models.CharField(max_length=2)
	patente = models.ForeignKey(cat001patentes)
	seccion = models.CharField(max_length=1)
	pedimento = models.CharField(max_length=10)
	firmado = models.BooleanField(default=False)
	firma = models.CharField(max_length=50, null=True, blank = True)
	puerto_embarque = models.ForeignKey(cat001puertos)
	clave_embarque = models.IntegerField(null=True, blank = True)
	pais_origen = models.ForeignKey(cat001paises)
	pais_procedencia = models.ForeignKey(cat001paises, related_name='+')
	tipo_cambio = models.DecimalField(max_digits=12, decimal_places=6)
	buque = models.ForeignKey(cat001buques)
	linea_aerea = models.ForeignKey(cat001linea_aerea)
	fecha_pago = models.DateField(null=True, blank = True)
	fecha_entrada = models.DateField(null=True, blank = True)
	fecha_revalidacion = models.DateField(null=True, blank = True)
	fecha_despacho = models.DateField(null=True, blank = True)
	fecha_bl = models.DateField(null=True, blank = True)
	fecha_arribo = models.DateField(null=True, blank = True)
	peso_bruto  = models.DecimalField(max_digits=32, decimal_places=4, null=True, blank = True)
	total_bultos = models.CharField(max_length=150, null=True, blank = True)
	contenedores_pedimento = models.IntegerField(max_length = 4)
	contenedores_embarque = models.IntegerField(max_length= 4)
	fletes = models.DecimalField(max_digits=32, decimal_places=4, null=True, blank = True)
	seguros = models.DecimalField(max_digits=32, decimal_places=4, null=True, blank = True)
	otros_incrementables = models.DecimalField(max_digits=32, decimal_places=4, null=True, blank = True)
	valor_dls_factura = models.DecimalField(max_digits=32, decimal_places=4, null=True, blank = True)
	valor_monex_factura = models.DecimalField(max_digits=32, decimal_places=4, null=True, blank = True)
	valor_aduana = models.DecimalField(max_digits=32, decimal_places=4, null=True, blank = True)
	clave_cliente = models.ForeignKey(rel001claves_cliente)
	direccion_cliente = models.ForeignKey(cat001direcciones)
	pais_cliente = models.ForeignKey(cat001paises, related_name='++')
	borrado = models.BooleanField(default=False, db_index=True)
	estatus = models.CharField(max_length=1, choices = estatus, default='A', blank=True, db_index=True)
	
	class Meta:
		permissions = (
			('can_view_operaciones','Can View Operaciones'),
			)
	
	def __unicode__(self):
		return self.referencia

class det001impuestos_pedimento(models.Model):
	impuestos_pedimento_id = models.AutoField(primary_key=True)
	operacion = models.ForeignKey(reg001operacion)
	impuesto = models.ForeignKey(cat001impuestos)
	borrado = models.BooleanField(default=False, db_index=True)

class det001contenedores(models.Model):
	contenedor_id = models.AutoField(primary_key=True)
	numero_cotenedor = models.CharField(max_length=50, null =True)
	tipo_contenedor = models.ForeignKey(cat001tipo_contenedores)
	borrado = models.BooleanField(default=False, db_index=True)


class det001candados(models.Model):
	candado_id= models.AutoField(primary_key=True)
	numero = models.CharField(max_length=150)
	contenedor = models.ForeignKey(det001contenedores)
	borrado = models.BooleanField(default=False, db_index=True)

class det001guias(models.Model):
	guias_choices = (('ma','Master'),
					('ho','House'))
	guia_id = models.AutoField(primary_key=True)
	operacion = models.ForeignKey(reg001operacion)
	numero_guia = models.CharField(max_length=50)
	tipo = models.CharField(max_length=6,choices = guias_choices, default='ma')
	borrado = models.BooleanField(default=False, db_index=True)
	
	def __unicode__(self):
		return 'Ref: %s Guia: %s %s'%(self.operacion.referencia,self.numero_guia,self.tipo)

class det001identificadores_pedimento(models.Model):
	"""
	Catalogo de Identificadores 
	"""
	identificador_id = models.AutoField(primary_key=True)
	operacion = models.ForeignKey(reg001operacion)
	clave_identificador = models.CharField(max_length=5)
	descripcion = models.TextField(null=True, blank = True)
	permiso = models.CharField(max_length=50, null=True, blank = True)
	borrado = models.BooleanField(default=False, db_index=True)


class reg001facturas(models.Model):
	"""
	Registros de las facturas por referencia
	"""
	factura_id = models.AutoField(primary_key =True)
	operacion = models.ForeignKey(reg001operacion, null=True, blank = True)
	proveedor = models.ForeignKey(cat001proveedores)
	pais_factura = models.ForeignKey(cat001paises)
	moneda_factura = models.ForeignKey(cat001monedas)
	fecha_factura = models.DateField()
	incoterm = models.CharField(max_length=3, null=True, blank = True)
	factor_moneda = models.DecimalField(max_digits=12, decimal_places=6, null=True, blank = True)
	valor_dls = models.DecimalField(max_digits=32, decimal_places=6, null=True, blank = True)
	valor_monex = models.DecimalField(max_digits=32, decimal_places=6, null=True, blank = True)
	direccion = models.ForeignKey(cat001direcciones)
	edocument = models.CharField(max_length=25, null=True, blank = True)
	numero_operacion = models.CharField(max_length=25, null=True, blank = True)
	folio = models.CharField(max_length=30,null = True, blank=True)
	serie = models.CharField(max_length=30,null = True, blank=True)
	borrado = models.BooleanField(default=False, db_index=True)
	
class det001facturas(models.Model):
    """
        Clase para almacenar el detalle de la factura
    """
    detalle_id = models.AutoField(primary_key=True)
    factura = models.ForeignKey(reg001facturas)

class det001partidas(models.Model):
	partida_id = models.AutoField(primary_key=True)
	operacion = models.ForeignKey(reg001operacion)
	numero_partida = models.SmallIntegerField()
	detalle_mercancia = models.TextField()
	fraccion = models.CharField(max_length=8)
	cantidad_comercial = models.DecimalField(max_digits=32, decimal_places=6)
	um_comercial = models.SmallIntegerField()
	cantidad_tarifa = models.DecimalField(max_digits=32, decimal_places=6)
	um_tarifa = models.SmallIntegerField()
	valor_aduana = models.DecimalField(max_digits=32, decimal_places=6)
	valor_mercancia = models.DecimalField(max_digits=32, decimal_places=6)
	valor_dls = models.DecimalField(max_digits=32, decimal_places=6)
	numeros_serie = models.TextField(null=True, blank = True)
	observaciones = models.TextField(null=True, blank = True)
	metodo_valoracion = models.SmallIntegerField()
	vinculacion = models.SmallIntegerField()
	cuota_operacion = models.DecimalField(max_digits=32, decimal_places=6)
	tasa_igie = models.DecimalField(max_digits=6, decimal_places=3)
	igie = models.DecimalField(max_digits=32, decimal_places=6)
	tasa_iva = models.DecimalField(max_digits=6, decimal_places=3)
	iva = models.DecimalField(max_digits=32, decimal_places=6)
	tasa_isan = models.DecimalField(max_digits=6, decimal_places=3)
	isan = models.DecimalField(max_digits=32, decimal_places=6)
	tasa_ieps = models.DecimalField(max_digits=6, decimal_places=3)
	ieps = models.DecimalField(max_digits=32, decimal_places=6)
	tasa_max = models.DecimalField(max_digits=6, decimal_places=3)
	tasa_cc = models.DecimalField(max_digits=6, decimal_places=3)
	cc = models.DecimalField(max_digits=32, decimal_places=6)
	recargos = models.DecimalField(max_digits=32, decimal_places=6)
	factor_actualizacion = models.DecimalField(max_digits=6, decimal_places=4)
	precio_unitario = models.DecimalField(max_digits=32, decimal_places=6)
	dta_partida = models.DecimalField(max_digits=32, decimal_places=6)
	
class det001identificadores_partida(models.Model):
    """
        Detalle de identificadores por partida de pedimento
    """
    identificador_id = models.AutoField(primary_key=True)
    partida = models.ForeignKey(det001partidas)
    identificador = models.CharField(max_length=10, db_index=True)
    descripcion = models.CharField(max_length=100)
    complemento_1 = models.CharField(max_length=50, null=True,blank=True)
    complemento_2 = models.CharField(max_length=50, null=True,blank=True)
    complemento_3 = models.CharField(max_length=50, null=True,blank=True)

class bit001archivo_mercancias(models.Model):
	estados_ = (
		('B','Borrado'),
		('D','Disponible'),
		('I','Inaccesible')
	)
	archivo_mercancia_id = models.AutoField(primary_key=True)
	archivo = models.FileField(upload_to='archivos_mercancias_samsung/%Y/%m/%d')
	nombre_original = models.CharField(max_length=50, blank = True, null=True)
	fecha_carga = models.DateTimeField(auto_now =True, auto_now_add=True, db_index=True)
	estatus = models.CharField(max_length=2, choices=estados_,db_index=True, default='D')
	usuario = models.ForeignKey(User)
	bl = models.CharField(max_length=30,db_index=True)
	numero_factura = models.CharField(max_length=254, db_index=True)
	moneda = models.CharField(max_length=15)
	proveedor = models.CharField(max_length=255, null=True, blank=True)
	
	

#
#   Bitacoras
#
class bit001facturas(models.Model):
	patente = models.ForeignKey(cat001patentes)
	proveedor = models.ForeignKey(cat001proveedores)
	cliente = models.ForeignKey(rel001claves_cliente)
	archivo = models.FileField(upload_to='facturas_importadas')
	fecha_carga = models.DateTimeField(auto_now=True)
	ultima_actualizacion = models.DateTimeField(auto_now=True)
	usuario = models.ForeignKey(User)
	numero_factura = models.CharField(max_length=255, db_index=True)
	nombre_proveedor = models.CharField(max_length=255, null=True, blank=True)
	fecha_factura = models.DateField(null=True,blank=True)
	



'''
class ref001documentos_digitales(models.Model):
	"""Modelo que almacena la relacion de las imagenes digitalizadas
	en formato pdf"""

	documento_id = models.Autofield(primary_key = True)
	nombre = models.textField()
	referencia = models.ForeignKey(reg001operacion)
	

	class Meta:
		verbose_name = _('Documento')
		verbose_name_plural = _('Documentos')

	def __unicode__(self):
		return
'''