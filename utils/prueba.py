import sys
sys.path.append('C:\\Users\\AlfredoVG.DAIMEX\\Documents\\GitHub\\daiweb\\')
sys.path.append('C:\\Users\\AlfredoVG.DAIMEX\\Documents\\GitHub\\daiweb\\utils\\')
sys.path.append('C:\\Users\\AlfredoVG.DAIMEX\\Documents\\GitHub\\daiweb\\utils\\reportes_extranet\\')
import referencias_vivas
c_ = referencias_vivas.Reporte_vivas('C:\\Users\\AlfredoVG.DAIMEX\\Documents\\GitHub\\daiweb\\utils\\reportes_extranet\\')
c_.genera_xlsx()
