#!/usr/bin/env python
# -*- coding: utf-8 -*-

from xml.dom import minidom
import os

#  Liberia para la generacion del reporte de facturacion CFDI
# que se envia mediante el portal de ABB

class Abb:
    _padres = []
    
    _reporte = []
    _codigo_proveedor = 806550
    _factura = {
        'id':0
        ,'factura': None
        ,'uuid' : None
        ,'factura_padre': ''
        ,'codigo_proveedor': ''
        ,'importe_total': None
        ,'divisa':None
        ,'archivo_xml': None
    }
    _orden = ['factura','uuid','factura_padre','codigo_proveedor','importe_total','divisa','archivo_xml']
    def __init__(self):
        return
    
    def add_xml(self,_archivo, _id, _factura_padre = None):
        """
        Recibe la direccion del archivo xml y el numero de la factura padre
        correspondiente, agrega ese archivo xml al objeto para generar el reporte final
        """
        xml_ = None
        xml_ = minidom.parseString(_archivo.read())
        self._reporte.append(self.procesa_xml(xml_, _archivo, _id, _factura_padre))
        return
    
    def procesa_xml(self, _xml, _archivo,_id, _factura_padre):
        """
        Recibe el objeto xml, la ruta del archivo y el numero de la factura padre,
        extrae la informacion necesaria del objeto xml  y devuelve un diccionario.
        """
        factura_ = self._factura.copy()
        
        if _xml.hasChildNodes():
            for child_ in _xml.childNodes :
                if child_.nodeName == 'cfdi:Comprobante':
                    factura_['id'] = _id
                    factura_['importe_total'] = child_.getAttribute('total')
                    factura_['factura'] = '%s%s'%(child_.getAttribute('serie'),child_.getAttribute('folio'))
                    factura_['divisa'] = child_.getAttribute('Moneda')
                    if _factura_padre :
                        factura_['factura_padre'] = _factura_padre
                    else:
                        factura_['codigo_proveedor'] = self._codigo_proveedor
                        self._padres.append({'id':_id, 'factura':'%s%s'%(child_.getAttribute('serie'),child_.getAttribute('folio'))})
                    factura_['uuid'] = child_.getElementsByTagName('tfd:TimbreFiscalDigital')[0].getAttribute('UUID')
                    factura_['archivo_xml'] = _archivo.name
        return factura_
    
    def get_padre(self, _parent_id):
        
        for padre_ in self._padres:
            if padre_['id'] == _parent_id:
                return padre_['factura']
        
        return ''
    
    def genera_reporte(self):

        reporte_ = []
        for factura_ in self._reporte:
            cadena_ = ''
            for campo_ in self._orden:
                if campo_ == 'factura_padre':
                    cadena_ = cadena_ + '%s|'%self.get_padre(factura_['factura_padre'])
                else:
                    cadena_ = cadena_ + '%s|'%factura_[campo_]
            reporte_.append(cadena_)
        
        return '%0D%0A'.join(reporte_)
