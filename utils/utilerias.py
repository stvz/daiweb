from django.core.paginator import Paginator, EmptyPage, PageNotAnInteger
import string , random

#
#   Liberia de funciones varias
#           por: Manuel Alejandro Estevez Fernandez
#

def nombre_aleatorio(size=6, chars=string.ascii_uppercase + string.digits):
    """
    Genera una cadena de manera aleatoria.
    Se puede definir la longitud y los caracteres.
    """
    return ''.join(random.choice(chars) for _ in range(size))


#def paginacion_json(_paginacion):
#    if isinstance(_paginacion, ):
#        pass