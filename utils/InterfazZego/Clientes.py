# -*- coding: utf-8 -*-
#
#       Liberia de Metodos y Clases para la interaccion con las operaciones relacionadas con referencias
# en los archivos dbfs, donde se encuentra la información de saai, zego y contabilidad.
#           por: Manuel Alejandro Estevez Fernandez
#               Julio , 2014
#

from Base import Base
import pyodbc

class Clientes(Base):
    
    def getCliente(self,_cve = 0, _rfc ='' , _nombre = '', _exac=True):
        condicion_ = []
        if _cve != 0 :
            condicion_.append(' cvecli18 = {0} '.format(_cve))
        elif _rfc !=  '' :
            condicion_.append(" rfccli18 = '{0}' ".format(_rfc))
        elif _nombre != '':
            if _exac:
                condicion_.append("  nomcli18 = '{0}' ".format(_nombre))
            else:
                condicion_.append(" nomcli18 like '%{0}%' ".format(_nombre))
        
        if len(condicion_) >0:
            consulta_ = "select * from ssclie18 where {0}".format(' and '.join(condicion_))
        else:
            raise "Debe ingresar al menos uno de las siguientes condiciones: clave, rfc o nombre"
            return
        
        self.conexionODBC('dbf_saai')
        cliente_ = self.dictResult(_consulta=consulta_)
        self.cerrarODBC()
        
        return cliente_
    