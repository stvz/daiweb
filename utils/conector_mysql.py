import MySQLdb as my_
import MySQLdb.cursors, ConfigParser
import os

class Conexion():
    
    _cxn = None
    _cursor = None

    def __init__(self):
        """
            Realiza la conexion con la db establecida en la configuracion. 
        """
        config = ConfigParser.ConfigParser()
        config.read([os.path.join(os.path.dirname(os.path.dirname(__file__)),'utils','vucem.conf')])
        self._cxn = my_.connect(host='%s'.strip()%config.get('db_mysql','servidor')
                    , user= '%s'.strip()%config.get('db_mysql','usuario')
                    , passwd = '%s'.strip()%config.get('db_mysql','pass')
                    , db = '%s'.strip()%config.get('db_mysql','db')
                    , charset='utf8'
                    , cursorclass=my_.cursors.DictCursor)
        self._cursor = self._cxn.cursor()
        return
    
    def __close__(self):
        """
            Finaliza los recursos abiertos
        """
        self._cursor.close()
        self._cxn.close()
        return
    
    def exe(self,_sql):
        """
            Ejecuta un comando sql
        """
        self._cursor.execute(_sql)
        return
    
    def get_resultados(self,_consulta):
        """
            Realiza la consulta y devuelve una tupla de diccionarios
        """
        self._cursor.execute(_consulta)
        resultados_ = self._cursor.fetchall()
        return resultados_