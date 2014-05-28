import os, sys, re
from django.core.files.storage import default_storage
from django.contrib.auth.models import User
from previos.models import enc004previo, det004previo
from django.core.files import File
from utils.utilerias import nombre_aleatorio

class CargaPrevios:
	_raiz = None	
	def __init__(self, _raiz):
		self._raiz = _raiz
		return
	
	def get_directorios(self):
		mask_previos_ = re.compile('^DAI*')
		directorios_ = []
		ban_ = 50
		cont_ = 0
		for (dirpath, dirnames, filenames) in os.walk(self._raiz):
			
			
			for dirname in dirnames:
				
				if mask_previos_.match(dirname):
					archivos_ = self.get_archivos(os.path.join(dirpath,dirname))
					if len(archivos_)>0:
						cont_ += 1
						previo_ = enc004previo()
						previo_.referencia = dirname
						#previo_.usuario = User.objects.get(pk=1)
						previo_.save()
						for archivo_ in archivos_:
							try:
								det_ = det004previo()
								det_.previo = previo_
								#file_f_ = File(file_)
								#file_f_.name = '%s%s'%(nuevo_nombre_,os.path.splitext(archivo_)[1])
								f_ = File(open(archivo_))
								f_.name = os.path.splitext(archivo_)[0].split('\\')[-1]
								det_.archivo= f_
								det_.nombre_original = os.path.splitext(archivo_)[0].split('\\')[-1]
								det_.tipo = os.path.splitext(archivo_)[1]
								det_.nuevo_nombre = nombre_aleatorio()
								det_.save()
							except IOError:
								pass
		return 'Directorios Importados: %s '%cont_

	def get_archivos(self,_directorio):
		archivos_ = []
		for (dirpath, dirnames, filenames) in os.walk(_directorio):
			for file_ in filenames:
				if file_.lower().find('jpg') != -1  or file_.lower().find('pdf') != -1:
					archivos_.append(os.path.join(dirpath,file_))
		return archivos_
if __name__ == '__main__':	
	c_ = CargaPrevios('t:\\')
	c_.get_directorios()