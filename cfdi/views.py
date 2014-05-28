# Create your views here.
import urllib

import os
import daiweb.settings as cnf_

class utilerias():
    _base_xsd = 'http://www.sat.gob.mx/cfd/3/cfdv3.xsd'
    
    
    def get_xsd(self ):
        try:
            urllib.urlretrieve(self._base_xsd, os.path.join(cnf_.BASE_DIR,self._base_xsd.split('/')[-1]))
            print 'Descargo el archivo correctamente.'
        except Exception as e:
            print 'Ocurrio la excepcion: %s'%e
        return
    