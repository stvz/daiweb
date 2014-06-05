

from openpyxl import load_workbook
import xlrd
from conector_mysql import Conexion
from os import path

def tipo_archivo(_archivo):
    """
    Analiza el nombre del archivo y devuelve el tipo de archivo:
    1 para excel <2003
    2 para excel >2007
    0 para cualquier otro
    """
    extension_ = path.splitext(_archivo)[1]
    
    if extension_ == 'xls':
        tipo_ = 1
    elif extension_ == 'xlsx':
        tipo_ = 2
    else:
        tipo_ = 0
    
    return tipo_

class GLP:
    """
    Clase para el manejo de los archivos en excel del reporte GLP,
    obtenido del sistema GLP de Samsung, disponible mediante acceso web.
    Este reporte debe encontrarse en la primera hoja del libro.
    El cual contiene los siguientes campos, en ese orden:
        S/R No
        House B/L
        Master B/L
        Shipper
        Consignee
        Manufacturer
        Buyer
        Branch
        Ship To Party
        Partner	Partner(Carrier)
        Partner(3PL)
        Partner(Fwder)
        Partner(Trucker)
        P/O No
        Container No
        D/O No
        Sales Org.
        Division
        ShippingType
        Incoterms
        G/I Date
        ETD
        Loading Location(SA)
        S/R Create Date
        ETA(SA)
        ETA(Updated)
        Discharging Location
        Final Destination
        VSL/FLT
        Customs Broker
        Cleared Date
        Gross Weight
        Total Quantity
        Invoice Amount
        Invoice Currency
        Material Number
        Seller Plant
        Buyer Plant
        Consignee Storage Location
        B/L uploaded Date
        Billing
        Customs Status
        FTA Flag
    """
    
    _tipo = 0
    _archivo = ''
    _glp = []
    _campos = ['S/R No', 'House B/L', 'Master B/L', 'Shipper'
               , 'Consignee', 'Manufacturer', 'Buyer', 'Branch'
               , 'Ship To Party', 'Partner\tPartner(Carrier)'
               , 'Partner(3PL)', 'Partner(Fwder)', 'Partner(Trucker)'
               , 'P/O No', 'Container No', 'D/O No', 'Sales Org.'
               , 'Division', 'ShippingType', 'Incoterms', 'G/I Date'
               , 'ETD', 'Loading Location(SA)', 'S/R Create Date'
               , 'ETA(SA)', 'ETA(Updated)', 'Discharging Location'
               , 'Final Destination', 'VSL/FLT', 'Customs Broker'
               , 'Cleared Date', 'Gross Weight', 'Total Quantity'
               , 'Invoice Amount', 'Invoice Currency', 'Material Number'
               , 'Seller Plant', 'Buyer Plant', 'Consignee Storage Location'
               , 'B/L uploaded Date', 'Billing', 'Customs Status', 'FTA Flag']
    
    def __init__(self,_archivo):
        
        self._tipo = tipo_archivo(_archivo)
        if self._tipo != 0:
            self._archivo = _archivo
            self.carga_archivo()
        else:
            raise "El tipo de archivo o extesión, no coinciden con los de excel"
        return
    
    def carga_archivo(self):
        
        return
    
    def importa_xls(self):
        wb_ = xlrd.open_workbook(self._archivo)
        
        
        return
    
    def importa_xlsx(self):
        """
        Metodo para importar el contenido del reporte a un objeto.
        """
        wb_ = load_workbook(filename=self._archivo, use_iterators= True)
        ws_ = wb_.worksheets[0]
        valores_ = []
        for fila_ in ws_.iter_rows():
            if fila_[0].internal_value == self._campos[0]:
                pass
            else:
                # defino un diccionario nuevo
                guia_ = dict()
                for indice_ in range(len(self._campos)):
                    #el cual se va llenando con los campos declarados y los valores
                    #obtenidos de la hoja en excel.
                    guia_.update({self._campos[indice_]:row[indice_].internal_value})
                valores_.append(guia_)
        wb_.close()
        return valores_

    