#!/usr/bin/env python

# Librerias para explotacion del WEBService de Vucem
#    Por Manuel Alejandro Estevez Fernandez
#


from suds.client import Client
from suds.wsse import *
import ConfigParser


class Ws_Vucem:
    
    _metodo = None #posibles valores 
    
    def __init__(self, **kwargs):
        if kwargs is not None:
            for key, value in kwargs.iteritems():
                setattr(self,key,value)
        self._config = ConfigParser.ConfigParser()
        try:
            self._config.read([os.path.join(os.path.dirname(os.path.dirname(__file__)),'utils','vucem.conf')])
        except IOError:
            raise 'No se encuentra el archivo de configuracion vucem.conf'
        return
    
    def conexion(self,**kwargs):
        if kwargs is not None:
            for key, value in kwargs.iteritems():
                setattr(self,key,value)
        else:
            if hasattr(self,'_usuario','_pass'):
                return
            else:
               pass
    
    
    def get_Pedimentos(self, **kwargs ):
        if kwargs is not None:
            for key, value in kwargs.iteritems():
                setattr(self,key,value)
            self._ws_conexion = Client(self._config.get('vucem','ws_url_lista'))
            self
        else:
            raise "Debes ingresar alguno de los parametros para realizar la consulta."
        return
    
    def build_consulta(self):
        
    
    #code

class Peticion:
    

token = UsernameToken('myusername', 'mypassword')
security.tokens.append(token)
client.set_options(wsse=security)