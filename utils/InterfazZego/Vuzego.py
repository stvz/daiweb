# -*- coding: utf-8 -*-
#
#       Liberia para la interaccion/extraccion de informacion de la base de datos de ventanilla unica
#           por: Manuel Alejandro Estevez Fernandez
#               Julio , 2014
#

from utils.conector_mysql import Conexion

class Vuzego:
    
    _conexion = None
    
    def __init__(self):
        self._conexion = Conexion()
        return
    
    def getFacturas(self, _facturas = '', _proveedor = 0, _cliente=0, _fecha = '', _proveedores= '', _cove=''):
        condicion_ = []
        
        if _facturas!= '':
            condicion_.append(" fact_.t_numfac in ( '{0}') ".format("','".join(frozenset(_facturas.split(',')))))
        
        if _proveedor != 0 :
            condicion_.append(" fact_.i_cve_pro = {0} ".format(_proveedor) )
        
        if _proveedores != '':
            condicion_.append(" fact_.i_cve_pro in ({0}) ".format(_proveedores) )
        
        if _cliente != 0:
            condicion_.append(" fact_.i_cve_cli = {0} ".format(_cliente) )
        
        if _fecha != '':
            condicion_.append(" fact_.d_fecfac = '{0}' ".format(_fecha))
        
        if _cove!= '':
            condicion_.append(" fact_.t_e_document = '{0}' ".format(_cove) )
        
        if len(condicion_) >0:
            consulta_ = """
                        select fact_.t_numfac as 'factura'
                        , fact_.d_fecfac as fecha_factura
                        , fact_.i_cve_cli as clave_cliente
                        , fact_.t_rfccli as rfc_cliente
                        , fact_.t_nomcli as nombre_cliente
                        , fact_.i_cve_pro as clave_proveedor
                        , fact_.t_idfiscpro as irs_proveedor
                        , fact_.t_nompro as nombre_proveedor
                        , case fact_.i_tipo_operacion when 1 then 'Importacion' else 'Exportacion'end as operacion
                        , case fact_.n_no_operacion when 0 then '' else fact_.n_no_operacion end no_operacion
                        , ifnull(fact_.t_e_document,'') as cove
                        , fact_.t_monfac as moneda_factura
                        , sum(dmer_.n_valtotal) as total_factura
                    from enc001_facturas fact_
                    join det001_facturas dmer_ on fact_.n_cve_fact = dmer_.n_fk_fact
                    where {0}
                    group by fact_.n_cve_fact
                        """.format(' and '.join(condicion_))
        else:
            raise "No se tienen condiciones para realizar la consulta"
        resultados_ = self._conexion.get_resultados(consulta_)
        self._conexion.__close__()
        return resultados_