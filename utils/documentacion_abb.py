# -*- coding: utf-8 -*-


from openpyxl import load_workbook
import xlrd, operator
from conector_mysql import Conexion
from os import path
from documentacion_samsung import tipo_archivo


class ABB:
    _tipo = 0
    _formato = {
        
    }
    _archivo = ''
    _campos = {
        'Billing Date':None ,
        'DChl': None ,
        'Billing Doc.': None ,
        'Item': None,
        'Customer Order': None,
        'Customer Item': None ,
        'Material': None,
        'Customer Material': None ,
        'Description' : None,
        'Bill. qty' : None,
        'SU': None,
        'Net Value': None ,
        'Curr.' :None
    }
    
    def __init__(self, _archivo ):
        self._tipo = tipo_archivo(_archivo)
        
        if self._tipo != 0:
            self._archivo = _archivo
            self.carga_archivo()
        else:
            raise "El tipo de archivo o extension, no coinciden con los de excel."
    
    def carga_archivo(self ):
        if self._tipo == 1:
            estatus_, mercancia_ = self.importa_xls()
        elif self._tipo == 2:
            estatus_, mercancia_ = self.importa_xlsx()
        else:
            estatus_ = ''
        
        return estatus_, mercancia_
    
    def verifica_formato(self):
        """
        Funcion que revisa que los campos necesarios se encuentren en el formato, esto debido
        a la variedad de proveedores que maneja el importador.
        """
        
        return
        
    
    def get_campos(self ):
        for indice_ in range(ws_.nrows):
            for col_ in range(ws_.ncols):
                _campos.append(ws_.row(indice_)[col_].value)
            break
        
        return
    
    def get_fracciones(self, ):
        for partida_ in range:
            code
        
    
    
    def importa_xls(self):
        wb_ = xlrd.open_workbook(self._archivo)
        ws_ = wb_.sheet_by_index(0)
        valores_ = []
        for indice_ in range(ws_.nrows):
            if ws_.row(indice_)[0] == self._campos[0]:
                pass
            else:
                mercancia_ = dict()
                for campo_ in range(len(self._campos)):
                    mercancia_.update({self._campos[campo_]:ws_.row(campo_)})
                    
                valores_.append(mercancia_)
        return valores_
    
    ##########################
    
from openpyxl import load_workbook
import xlrd, operator
from os import path
import MySQLdb as my_

archivo_ = r"C:\Users\AlfredoVG.DAIMEX\Documents\GitHub\daiweb\media\facturas\abb\FACTS EN EXCEL MHG42112297.xls"
wb_ = xlrd.open_workbook(archivo_)
ws_ = wb_.sheet_by_index(0)
valores_ = []
_campos = {
    'Billing Date':None ,
    'DChl': None ,
    'Billing Doc.': None ,
    'Item': None,
    'Customer Order': None,
    'Customer Item': None ,
    'Material': None,
    'Customer Material': None ,
    'Description' : None,
    'Bill. qty' : None,
    'SU': None,
    'Net Value': None ,
    'Curr.' :None
}
_campos = []    
for indice_ in range(ws_.nrows):
	for col_ in range(ws_.ncols):
		_campos.append(ws_.row(indice_)[col_].value)
	break

for indice_ in range(1,ws_.nrows):
	item_ = dict()
	for col_ in range(ws_.ncols):
		item_[_campos[col_]] = ws_.row(indice_)[col_].value
	_mercancia.append(item_)

_cxn = my_.connect(host='10.66.10.5'
                    , user= 'root'
                    , passwd = 'Grk8520Extranet'
                    , db = 'ventanilla_unica'
                    , charset='utf8'
                    , cursorclass=my_.cursors.DictCursor)
_cursor = _cxn.cursor()


#
# Opciones para el grabado de informacion en DBFS
# utilizando pyodbc
#

con_ = pyodbc.connect('DSN=dbf_zego',autocommit=False)
cursor_ = con_.cursor()

try:
    campo_insertado_ = cursor_.execute("""INSERT INTO d05artic(
                nuevocpo, ffactp05, tpmerc05, refe05, pedi05, fact05, prov05,
                cpro05, item05, frac05, desc05, obse05, agru05, pped05, pfac05,
                caco05, umco05, cata05, umta05, vafa05, grupo05, candef05, remesa05,
                agrub05, descod05, user05, dati05, comp05, cvpaor05, noreq05,
                nidmercv, nlote05)
        VALUES ('', {}, 'R', 'DAI14TEST', '', '', 0123,
                '', '', '', '', '', 0, 0, 0,
                0, 0, 0, 0, 0, 0, 0, '',
                0, '', '', {}, '', '', '',
                '', '')""").rowcount
    if campo_insertado_ != -1 :
        con_.commit()
    else:
        con_.rollback()
except IntegrityError as e:
    print e
    con_.rollback()
except DataError as e:
    print e
    con_.rollback()