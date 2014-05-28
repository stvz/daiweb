from django.db import models
# Create your models here.

#   Modelos Para el almacenamiento de Comprobante Factura Digital.
#       por: Manuel Alejandro Estevez Fernandez
#               Abril , 2014

class Contribuyente(models.Model):
    contribuyente_id = models.AutoField(primary_key=True)
    nombre = models.CharField(max_length=250)
    rfc = models.CharField(max_length=20)
    direccion = models.CharField(max_length=100)
    colonia = models.CharField(max_length=100)
    codigo_postal = models.CharField(max_length=15)
    ciudad = models.CharField(max_length=100)
    estado = models.CharField(max_length=100)
    pais = models.CharField(max_length=100)
    fecha_alta = models.DateTimeField(null=True)
    
    def alta(self):
        import datetime
        return Entes.fecha_alta.date() == datetime.date.today()
    
    def __unicode__(self):
        return u"%s RFC: %s"%(self.nombre,self.rfc)
    
class Lugar_expedicion(models.Model):
    expedicion_id = models.AutoField(primary_key=True)
    calle = models.TextField(null=True)
    noExterior = models.TextField(null=True)
    noInterior = models.TextField(null=True)
    colonia = models.TextField(null=True)
    localidad = models.TextField(null=True)
    referencia = models.TextField(null=True)
    municipio = models.TextField(null=True)
    estado = models.TextField(null=True)
    pais = models.TextField(null=True)
    codigoPostal = models.TextField(null=True)
    
class Comprobantes(models.Model):
    opciones_tipoDeComprobante = (
        ('in','ingreso'),
        ('eg','egreso'),
        ('tr','traslado'),
    )
    opciones_estado = (
        ('c','Cancelada'),
        )
    id_comprobante = models.AutoField(primary_key=True)
    version = models.CharField(max_length=3)
    serie = models.CharField(max_length=10)
    folio= models.CharField(max_length=20)
    fecha = models.DateTimeField()
    sello = models.TextField()
    noAprobacion = models.IntegerField()
    anoAprobacion = models.SmallIntegerField()
    formaDePago = models.TextField()
    noCertificado = models.CharField(max_length=20)
    certificado = models.TextField()
    condicionesDePago = models.TextField()
    subTotal = models.DecimalField(max_digits=32, decimal_places=6)
    descuento = models.DecimalField(max_digits=32, decimal_places=6, null=True, blank =True)
    motivoDescuento = models.TextField(null=True, blank = True)
    tipoCambio = models.DecimalField(max_digits=12, decimal_places=6,null=True, blank =True)
    moneda = models.CharField(max_length = 25,null=True, blank =True)
    total = models.DecimalField(max_digits=12, decimal_places=6,null=True, blank =True)
    tipoDeComprobante = models.CharField(max_length=2, choices = opciones_tipoDeComprobante, null= True, blank = True)
    metodoDePago = models.CharField(max_length = 50 , null = True)
    lugarDeExpedicion = models.CharField(max_length = 255, null=True)
    numCtaPago = models.CharField(max_length = 50, null=True, blank = True)
    folioFiscalOrig = models.CharField(max_length=255,null=True, blank=True)
    serieFolioFiscalOrig = models.CharField(max_length=255,null=True, blank=True)
    fechaFolioFiscalOrig = models.DateTimeField(null=True, blank = True)
    montoFolioFiscalOrig = models.DecimalField(max_digits=12, decimal_places=6,null=True, blank =True)
    cadena_original = models.CharField(max_length=254, null=True, blank=True) # Propiedad Agregada, no oficial
    valido = models.BooleanField(default=False, blank =True) # Propiedad de control, no oficial
    fechaRegistro = models.DateTimeField()
    emisor = models.ForeignKey(Contribuyente)
    receptor = models.ForeignKey(Contribuyente, related_name = '+')
    lugar_expedicion = models.ForeignKey(Lugar_expedicion)
    estado = models.CharField(max_length=2, choices=opciones_estado,null=True)
    fecha_cancelacion = models.DateTimeField(null=True,blank = True)
    comentarios = models.TextField(null=True, blank =True)
    

    
class Archivos(models.Model):
    archivo_id = models.AutoField(primary_key=True)
    comprobante = models.ForeignKey(Comprobantes)
    nombre_archivo = models.CharField(max_length=50, unique=True)
    fecha_carga = models.DateTimeField(auto_now_add=True)
    archivo = models.FileField(upload_to='cfdi/%Y/%m/%d')
    #xml = models.XMLField()
    #pdf = models.BooleanField(default=False)

class Detalle(models.Model):
    id_detalle = models.AutoField(primary_key=True)
    comprobante = models.ForeignKey(Comprobantes)
    cantidad = models.DecimalField(max_digits=12,decimal_places=3)
    unidad = models.CharField(max_length=20)
    descripcion = models.TextField()
    valor_unitario = models.DecimalField(max_digits=12,decimal_places=3)
    importe = models.DecimalField(max_digits=12,decimal_places=3)
    
class Impuestos(models.Model):
    impuesto_id = models.AutoField(primary_key=True)
    descripcion = models.TextField(unique=True)
    tasa = models.DecimalField(max_digits=12,decimal_places=3)
    
class Detalle_impuestos(models.Model):
    detalle_impuesto_id = models.AutoField(primary_key=True)
    comprobante = models.ForeignKey(Comprobantes)
    impuesto = models.ForeignKey(Impuestos)
    importe = models.DecimalField(max_digits=12,decimal_places=3)
    traslado = models.BooleanField(default=False)
    