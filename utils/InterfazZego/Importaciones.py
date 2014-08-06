# -*- coding: utf-8 -*-
#
#       Liberia de Metodos y Clases para la interaccion con las operaciones relacionadas con referencias
# en los archivos dbfs, donde se encuentra la información de saai, zego y contabilidad.
#           por: Manuel Alejandro Estevez Fernandez
#               Julio , 2014
#

from Base import Base
import pyodbc
class Importaciones(Base):
    
    
    def putReferencia(self, _referencia, _datos = {}):
        
        return
    
    
    def getReferencia(self,_referencia):
        """
        Realiza la consulta en la base de datos del zego para buscar la referencia dada.
        Devuelve una lista de diccionarios con los datos encontrados.
        """
        
        self.conexionODBC('dbf_saai')
        referencia_ = self.dictResult(_consulta="select * from ssdagi01 where refcia01 = '{0}'".format(_referencia))
        referencia_
        self.cerrarODBC()
        return referencia_
    
    def getReferencias(self, _referencias ):
        
        if isinstance(_referencias,list):
            consulta_ = "select * from ssdagi01 where refcia01 in ('{0}')".format("','".join(_referencias))
        elif isinstance(_referencias,str):
            consulta_ = "select * from ssdagi01 where refcia01 in ('{0}')".format("','".join(_referencias.split(',')))
        else:
            raise "Debe ingresar una lista o una cadena separada por comas"
            return
        
        self.conexionODBC()
        referencias_ = self.dictResult(_consulta=consulta_)
        self.cerrarODBC()
        return referencias_
