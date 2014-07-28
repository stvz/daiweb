# -*- coding: utf-8 -*-


import pyodbc


class Base():
    
    _cxn_odbc = None
    _odbc_cursor = None
    
    
    def __init__(self):
        
        return
    
    def dictResult(self,_consulta = "", _cursor=None ):
        if _consulta != '':
            
            if self._odbc_cursor:
                print _consulta,'\n'
                cursor_ = self._odbc_cursor.execute("{0}".format(_consulta))
                campos_ = [columna[0] for columna in cursor_.description]
                resultado_ = []
                for fila_ in cursor_.fetchall():
                    resultado_.append(dict(zip(campos_,fila_)))
                
                return resultado_
        elif _cursor:
            if isinstance(_cursor,pyodbc.Cursor):
                campos_ = [columna[0] for columna in _cursor.description]
                resultado_ = []
                for fila_ in _cursor.fetchall():
                    resultado_.append(dict(zip(campos_,fila_)))
                
                return resultado_
            else:
                raise "El objeto no es un cursor"
        else:
            raise "Debe pasar una consulta o cursor"
        
        return

    
    def conexionODBC(self,_dsn):
        try:
            self._cxn_odbc = pyodbc.connect('DSN={0};unicode_results=True;Charset=cp1252'.format(_dsn),autocommit=False)
            self._odbc_cursor = self._cxn_odbc.cursor()
        except Exception as e:
            print 'Error Conexion: \n%s'%e
        return
    
    def cerrarODBC(self):
        '''
        Metodo para cerrar la conexion abierta a traves del odbc
        '''
        self._odbc_cursor.close()
        self._cxn_odbc.close()
