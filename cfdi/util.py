from xml.dom.minidom import getDOMImplementation
from xml.dom import minidom
import sys, os , re


#   Utilerias varias par ael modulo de 
#
#

class domicilio:
        calle=''
        noExterior=''
        noInterior=''
        colonia=''
        localidad=''
        referencia=''
        municipio=''
        estado=''
        pais=''
        codigoPostal=''
        
        def __init__(self):
            return
class Cfdi:
    '''
    Clase para el manejo de CFDI, segpun el anexo 20 del SAT
    
    '''
    
    cfdi_ = ''
    
      
    comprobante_ = {'version':''
                    ,'serie':''
                    ,'folio':''
                    ,'fecha':''
                    ,'sello':''
                    ,'formaDePago':''
                    ,'noCertificado':''
                    ,'certificado':''
                    ,'condicionesDePago':''
                    ,'subTotal':''
                    ,'descuento':''
                    ,'motivoDescuento':''
                    ,'tipoCambio':''
                    ,'moneda':''
                    ,'total':''
                    ,'tipoComprobante':''
                    ,'metodoDePago':''
                    ,'lugarExpedicion':''
                    ,'numCtaPago':''
                    ,'folioFiscalOrig':''
                    ,'serieFolioFiscalOrig':''
                    ,'fechaFolioFiscalOrig':''
                    ,'montoFolioFiscalOrig':''
                    }
    emisor_ = {'rfc':''
               ,'nombre':''
               ,'domicilioFiscal': domicilio()
               ,'expedidoEn': domicilio()
            }
    
    
    def __init__(self):
        
        return
    
    def carga_archivo(self,_path ='' ):
        try:
            ruta_,extension_  = os.path.splitext(_path)
            if extension_.lower() == '.xml':
                self.path_ = _path
            else:
                print 'El archivo no tiene extesion XML, favor de verificar.'
        except Exception as e:
            print e
        
        return
    
    def lee_archivo(self):
        """
        Abre el archivo cargado previamente y
        utiliza minidom para parsearlo como objeto xml
        """
        try:
            self.xml_ = minidom.parse(open(self.path_,'r'))
        except Exception as e :
            print e
        
        return
    
    def obten_emisor(self ):
        pass
    
    def obten_receptor(self):
        pass
    
    def obten_comprobante(self):
        pass
    
    def obten_version(self ):
        pass
    
    def obten_lugar_expedicion(self):
        pass
    
    
    
    
    
    
    
