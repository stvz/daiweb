

#
#   Libreria para consumir archivos M3
#
#       por: Manuel Alejandro Estevez Fernandez
#           
#

from os import path

class M3:
    _ruta = ''
    def __init__(self,_ruta):
        self._ruta = _ruta
        return
    

class Pedimento:
    patente = None
    aduana = None
    pedimento = None
    tipo = None
    transporte = []
    guias = []
    contenedores = []
    facturas = []
    fechas = []
    identificadores = []
    cuentas_garantia = []
    tasa_pedimento = []
    contribuciones = []
    observaciones = []
    descargos = []
    partidas = []
    previos=[]
    rectificacion = []
    

