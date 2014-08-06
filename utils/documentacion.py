# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import xlrd, operator
from os import path
import MySQLdb as my_
from tracking.models import cat001FormatosXls, cat001CamposXls, det001FormatoCampos

class ImportaDocumentacion:
    _archivo = None
    _formato = None
    _cliente = None
    _proveedor = None
    _tipo = None
    
    def __init__(self,_ruta_archivo,_formato , _cliente = 0, _proveedor=0):
        
        self._archivo = _ruta_archivo
        self._formato = self.getFormato(_formato)
        self._cliente = _cliente
        self._proveedor = _proveedor
        self._tipo = self.tipo_archivo(_ruta_archivo)
        if self._tipo == 0 :
            raise "El archivo no es de tipo xls o xlsx"
        
        return
    
    def procesaArchivo(self):
        """
        Funcion principal para la extraccion de la informacion.
        """
        
        formato_ = None
        registros_ = None
        
        if self._tipo == 1:
            registros_, formato_ = self.procesaXLS(self._archivo,self._formato)
        elif self._tipo == 2:
            registros_, formato_ = self.procesaXLSX(self._archivo,self._formato)
        
        
        return
    
    def verificaFormato(self,_encabezado,_formato):
        
        for indice_ in range(len(_encabezado)):
            _formato['campos'][_encabezado[indice_]] = indice_
        
        completo_ = True
        for campo_ in _formato['campos']:
            if campo_['indice'] == None:
                completo_ = False
        
        _formato['completo'] = completo_
        
        return _formato
    
    def procesaXLS(self,_archivo, _formato):
        """
        Metodo utilizado para extraer la informacion de un archivo de excel 97 - 2003
        """
        
        wb_ = xlrd.open_workbook(_archivo)
        ws_ = wb_.sheet_by_index(0)
        
        encabezado_ = [ws_.cell_value(0,y_) for y_ in range(ws_.ncols)] 
        registros_ = [[ws_.cell_value(x_,y_) for y_ in range(ws_.ncols)] for x_ in range(1,ws_.nrows)]
        resultado_ = [dict(zip(encabezado_,registro_)) for registro_ in registros_]
        
        formato_ = self.verificaFormato(encabezado_,_formato)
        
        if not formato_['completo']:
            raise "El formato seleccionado no coincide con el archivo proporcionado"
        
        return resultado_, formato_
    
    def procesaXLSX(self,_archivo, _formato):
        """
        Metodo utilizado para extraer informacion de un archivo de excel 2007<
        """
        
        wb_ = load_workbook(filename=archivo_)
        ws_ = wb_.worksheets[0]
        
        encabezado_ = [ws_.rows[0][y_].value for y_ in range(len(ws_.rows[0])) ]
        registros_ = [[ws_.rows[x_][y_].value for y_ in range(len(ws_.rows[0]))] for x_ in range(1,len(ws_.rows))]
        resultado_ = [dict(zip(encabezado_,registro_)) for registro_ in registros_]
        
        formato_ = self.verificaFormato(encabezado_,_formato)
        
        if not formato_['completo']:
            raise "El formato seleccionado no coincide con el archivo proporcionado"
        
        return resultado_, formato_
    
    def tipo_archivo(self, _archivo):
        """
        Analiza el nombre del archivo y devuelve el tipo de archivo:
        1 para excel <2003
        2 para excel >2007
        0 para cualquier otro
        """
        extension_ = path.splitext(_archivo)[1]
        
        if extension_ == '.xls':
            tipo_ = 1
        elif extension_ == '.xlsx':
            tipo_ = 2
        else:
            tipo_ = 0
        
        return tipo_
    
    def getFormato(self, _id_formato ):
        # Extraemos la informacin del formato de acuerdo a su ID
        formato_ = cat001FormatosXls.objects.get(pk=_id_formato)
        # Asi mismo el detalle de columnas del formato y los tipos de campos de estos
        campos_ = det001FormatoCampos.objects.select_related().get(formato = formato_)
        # Definimos un diccionario que contendra la informacion de ese tipo de formato
        # el cual es un diccionario documento
        doc_ = {'nombre':formato_.nombre,'completo':False,'campos':[]}
        # se llena el detalle de campos (lista) del documento con el nombre de la columna, su campo
        # correspondiente y el indice donde se encontrara
        doc_['campos'].append([{'columna':campo_.nombre_columna,'campo':campo_.campo.nombre,'indice':None} for campo_ in campos_ ])

        return doc_