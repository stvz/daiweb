# -*- coding: utf-8 -*-
#
#
#


from Base import Base
import pyodbc

class Facturas(Base):
    
    def getFacturaSaai(self,_factura):
        """
        Realiza una consulta sobre la tabla del Saai donde se almacenan las facturas
        utilizando como condicion el numero de la factura
        """
        consulta_ = "select * from ssfact39 where alltrim(numfac39)='{0}' ".format(_factura)
        self.conexionODBC('dbf_saai')
        factura_ = self.dictResult(_consulta=consulta_)
        self.cerrarODBC()
        
        return factura_
    
    def getFacturasSaai(self,_referencia):
        """
        Realiza una consulta sobre la tabla del Saai donde se almacenan las facturas
        utilizando como condicion la referencia
        """
        consulta_ = "select * from ssfact39 where refcia39='{0}'".format(_referencia)
        self.conexionODBC('dbf_saai')
        facturas_ = self.dictResult(_consulta=consulta_)
        self.cerrarODBC()
        
        return facturas_
    
    def getFacturaZego(self,_factura):
        """
        Realiza una consulta sobre la tabla del Saai donde se almacenan las facturas
        utilizando como condicion el numero de la factura
        """
        consulta_ = "select * from d01factu where fact01='{0}' ".format(_factura)
        self.conexionODBC('dbf_zego')
        factura_ = self.dictResult(_consulta=consulta_)
        self.cerrarODBC()
        
        return factura_
    
    def getFacturasZego(self,_referencia):
        """
        Realiza una consulta sobre la tabla del Saai donde se almacenan las facturas
        utilizando como condicion la referencia
        """
        consulta_ = "select * from d01factu where refe01='{0}'".format(_referencia)
        self.conexionODBC('dbf_zego')
        facturas_ = self.dictResult(_consulta=consulta_)
        self.cerrarODBC()
        
        return facturas_    
    