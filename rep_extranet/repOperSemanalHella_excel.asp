
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<%
' ESTE ASP ES EL SEGUNDO Y ES PARA ADMINISTRADORES
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
Response.Buffer = TRUE
Response.Addheader "Content-Disposition", "attachment;filename=Hella_Semanal.xls"
Response.ContentType = "application/vnd.ms-excel"
Server.ScriptTimeOut=100000


strUsuario = request.Form("user")
strTipoUsuario = request.Form("TipoUser")


strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
permi2 = PermisoClientesTabla("B",Session("GAduana") ,strPermisos,"clie31")


if not permi2 = "" then
  permi2 = "  and (" & permi2 & ") "
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")
if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if
if blnAplicaFiltro then
   permi2 = " AND B.clie31 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi2 = ""
end if


if not permi = "" then
  permi = "  and (" & permi & ") "
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")
if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if
if blnAplicaFiltro then
   permi = " AND cvecli01 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
end if



strDateIni=""
strDateFin=""
strTipoPedimento= ""
strCodError = "0"

strDateIni=trim(request.Form("txtDateIni"))
strDateFin=trim(request.Form("txtDateFin"))
'*******************************************************
' Si es Impo o Expo
strTipoPedimento=trim(request.Form("rbnTipoDate"))
'*******************************************************

'***************************************************************************************************************
strDescripcion=trim(request.Form("txtDescripcion"))
strDateIni2=trim(request.Form("txtDateIni2"))
strDateFin2=trim(request.Form("txtDateFin2"))
strTipoPedimento2=trim(request.Form("rbnTipoDate2"))
strTipoFiltro=trim(request.Form("TipoFiltro"))


rbnTipoReporte=trim(request.Form("rbnTipoReporte"))
rbnTipoReporte = 2

'strDateIni=trim(request.Form("txtDateIni"))
'strDateFin=trim(request.Form("txtDateFin"))
'strTipoPedimento=trim(request.Form("rbnTipoDate"))

if not isdate(strDateIni) then
	strCodError = "5"
end if
if not isdate(strDateFin) then
	strCodError = "6"
end if
if strDateIni="" or strDateFin="" then
	strCodError = "1"
end if


if strCodError = "0" then

strHTML = ""
tmpTipo = ""
strSQL = ""


  if rbnTipoReporte  = "1" then 'Si es el encabezado

     '******************************************************************************************************************************
  else 'El detalle

     if strTipoPedimento  = "1" then
         tmpTipo = "IMPORTACION"
         strSQL =   " SELECT concat(concat(concat(concat(concat(concat(ltrim(substring(year(FECPAG01),3,2)),'-'),CVEADU01),'-'),PATENT01),'-'),NUMPED01) as IMPORTA, " & _
                    "        ltrim(refcia01) as Referencia, " & _
                    "        TIPCAM01 as TipoCambio, " & _
                    "        FACTMO01, " & _
                    "        DESDOC01 " & _
                    "  from ssdagi01  " & _
                    "  where cveped01 <> 'R1' and  " & _
                    "        firmae01 <> ''   and  " & _
                    "       fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                    "       fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND " & _
                    "       LTRIM(refcia01) <> 'GABBY' " & _
                    Permi & _
                  " GROUP BY REFCIA01"
     end if
     if strTipoPedimento  = "2" then
        tmpTipo = "EXPORTACION"
         strSQL =   " SELECT concat(concat(concat(concat(concat(concat(ltrim(substring(year(FECPAG01),3,2)),'-'),CVEADU01),'-'),PATENT01),'-'),NUMPED01) as IMPORTA, " & _
                    "        ltrim(refcia01) as Referencia, " & _
                    "        TIPCAM01 as TipoCambio, " & _
                    "        FACTMO01, " & _
                    "        DESDOC01 " & _
                    "  from ssdage01  " & _
                    "  where cveped01 <> 'R1' and  " & _
                    "        firmae01 <> ''   and  " & _
                    "       fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                    "       fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND " & _
                    "       LTRIM(refcia01) <> 'GABBY' " & _
                    Permi & _
                  " GROUP BY REFCIA01"
     end if
     'response.write(strSQL)
     'response.end


     if not trim(strSQL)="" then
          Set RsRep = Server.CreateObject("ADODB.Recordset")
		      RsRep.ActiveConnection = MM_EXTRANET_STRING
		      RsRep.Source = strSQL
  		    RsRep.CursorType = 0
	  	    RsRep.CursorLocation = 2
		      RsRep.LockType = 1
		      RsRep.Open()

             ' Comienza el HTML, se pintan los titulos de las columnas
             'strHTML = strHTML & " <p> <img src='../../ext-Images/Gifs/abbot.gif'> </p>"
             'strHTML = strHTML & " <p> <img width='181' eight='38'  src='http://10.66.1.4/PortalMySQL/Extranet/ext-Images/Gifs/abbot.gif'> </p> <P>&nbsp;</P>"
             strHTML = strHTML & " <p> &nbsp; </p>"
             strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">GRUPO REYES KURI, S.C. </font></strong> <br> "
             if tmpTipo = "EXPORTACION" then
	  	         strHTML = strHTML & "<strong><font color=""#969696"" size=""3"" face=""Arial, Helvetica, sans-serif""> Reporte Semanal de Operaciones de Exportación del " & strDateIni & " al " & strDateFin & " </font></strong>"
             else
               strHTML = strHTML & "<strong><font color=""#969696"" size=""3"" face=""Arial, Helvetica, sans-serif""> Reporte Semanal de Operaciones de Importación del " & strDateIni & " al " & strDateFin & " </font></strong>"
             end if
             strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	           strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)
             strHTML = strHTML & "<td width=""80""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> NUM PARTE           </font></strong></td>" & chr(13) & chr(10) ' Codigo del Numero de Parte
             strHTML = strHTML & "<td width=""80""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESCRIPCION         </font></strong></td>" & chr(13) & chr(10) ' Descripcion del numero de parte
             strHTML = strHTML & "<td width=""65""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FRACCION            </font></strong></td>" & chr(13) & chr(10) ' Fracción Arancelaria
             strHTML = strHTML & "<td width=""50""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TIPO                </font></strong></td>" & chr(13) & chr(10) ' Tipo de Bien
             strHTML = strHTML & "<td width=""80""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CONVERSION          </font></strong></td>" & chr(13) & chr(10) ' Factor moneda
             strHTML = strHTML & "<td width=""60""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> UNIDAD              </font></strong></td>" & chr(13) & chr(10) ' Unidad del numero de parte
             strHTML = strHTML & "<td width=""70""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PROYECTO            </font></strong></td>" & chr(13) & chr(10) ' Proyecto
             strHTML = strHTML & "<td width=""100"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> NUMSERIE            </font></strong></td>" & chr(13) & chr(10) ' Numero de serie
             strHTML = strHTML & "<td width=""100"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DOCENTRADA          </font></strong></td>" & chr(13) & chr(10) ' Numero de pedimento
             strHTML = strHTML & "<td width=""95""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA DE ALTA       </font></strong></td>" & chr(13) & chr(10) ' Fecha de alta empresa
             strHTML = strHTML & "<td width=""140"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PAIS DE PROCEDENCIA </font></strong></td>" & chr(13) & chr(10) ' Pais de procedencia
             strHTML = strHTML & "<td width=""120"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TIPO DE ENTRADA     </font></strong></td>" & chr(13) & chr(10) ' Tipo de entrada: Temporal, definitiva
             strHTML = strHTML & "</tr>"& chr(13) & chr(10)

             'if tmpTipo = "EXPORTACION" then
		         '  strHTML = strHTML & "<td width=""95""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> EXPORTA         </font></strong></td>" & chr(13) & chr(10) 'Folio del pedimento
             'else
             '  strHTML = strHTML & "<td width=""95""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IMPORTA         </font></strong></td>" & chr(13) & chr(10) 'Folio del pedimento
             'end if
             'strHTML = strHTML & "<td width=""70""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FACTURA       </font></strong></td>" & chr(13) & chr(10) ' Folio de la factura
             'strHTML = strHTML & "<td width=""60""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CODIGO        </font></strong></td>" & chr(13) & chr(10) ' Codigo del proveedor
             'strHTML = strHTML & "<td width=""80""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA FAC     </font></strong></td>" & chr(13) & chr(10) ' Fecha de la factura
             'strHTML = strHTML & "<td width=""40""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TASA          </font></strong></td>" & chr(13) & chr(10) ' Tasa Arancelaria
             'strHTML = strHTML & "<td width=""95""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TIPO TASA     </font></strong></td>" & chr(13) & chr(10) ' Tipo de Tasa Arancelaria
	  	       'strHTML = strHTML & "<td width=""90""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PRECIO        </font></strong></td>" & chr(13) & chr(10) ' Precio Unitario
             'strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CANTIDAD      </font></strong></td>" & chr(13) & chr(10) ' Cantidad Facturada
		         'strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CONVERSION    </font></strong></td>" & chr(13) & chr(10) ' Factor de Conversión
             'strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ORIGEN        </font></strong></td>" & chr(13) & chr(10) ' Pais de Origen
             'strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VENDEDOR      </font></strong></td>" & chr(13) & chr(10) ' Pais Vendedor
             'strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> F PAGO        </font></strong></td>" & chr(13) & chr(10) ' Forma de Pago
             'strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> INCOTERM      </font></strong></td>" & chr(13) & chr(10) ' Termino Internacional de Comercio

	        if not RsRep.eof then

           While NOT RsRep.EOF 'recorremos el cursor de operaciones
                  xRefer     =  RsRep.Fields.Item("Referencia").Value
                  strRefer   =  RsRep.Fields.Item("Referencia").Value
                  xImporta   =  RsRep.Fields.Item("IMPORTA").Value
                  'xFactu     =  RsRep.Fields.Item("facturas").Value
                  'xProv      =  RsRep.Fields.Item("proveedor").Value
                  xFactmo    =  RsRep.Fields.Item("FACTMO01").Value
                  xTipCam    =  RsRep.Fields.Item("TipoCambio").Value
                  xRegimen   =  RsRep.Fields.Item("DESDOC01").Value

                  '*************************************************************************************************
                  'traemos las Fracciones
                  strDetFracciones= " SELECT fraarn02,                      " & _
                                    "        cancom02,                      " & _
                                    "        u_medc02,                      " & _
                                    "        paiOri02 as PaisOrigenDestino, " & _
                                    "        paiscv02 as PaisVendedor,      " & _
                                    "        PREUNI02,                      " & _
                                    "        ordfra02                       " & _
                                    " FROM SSFRAC02                         " & _
                                    " WHERE  refcia02='" &  strRefer & "'   "
                  'response.write(strDetFracciones)
                  'response.end

                  Set RsDetFracciones = Server.CreateObject("ADODB.Recordset")
			            RsDetFracciones.ActiveConnection = MM_EXTRANET_STRING
			            RsDetFracciones.Source = strDetFracciones
     	            RsDetFracciones.CursorType = 0
			            RsDetFracciones.CursorLocation = 2
			            RsDetFracciones.LockType = 1
			            RsDetFracciones.Open()

			            if not RsDetFracciones.eof then ' si tiene fracciones
                     While not RsDetFracciones.eof ' recorremos las fracciones
                     '*************************************************************************************************
                         'dblDTA = RsDetFracciones.Fields.Item("DTA").Value

                         xunicom     = RsDetFracciones.Fields.Item("u_medc02").Value
                         xCantUniMed = 1

                         SELECT CASE xunicom
                             CASE 9
                                      xCantUniMed = 2
                             CASE 11
                                      xCantUniMed = 1000
                             CASE 14
                                      xCantUniMed = 1000
                             CASE 17
                                      xCantUniMed = 10
                             CASE 18
                                      xCantUniMed = 100
                             CASE 19
                                      xCantUniMed = 12
                             CASE ELSE
                                      xCantUniMed = 1
                         END SELECT

                         xFrac       = RsDetFracciones.Fields.Item("fraarn02").Value
                         xCant       = RsDetFracciones.Fields.Item("cancom02").Value * (xCantUniMed)
                         'xUniMed     = RsDetFracciones.Fields.Item("u_medc02").Value
                         xpaiOri     = RsDetFracciones.Fields.Item("PaisOrigenDestino").Value
                         xpaiscv     = RsDetFracciones.Fields.Item("PaisVendedor").Value
                         'xPreUni     = RsDetFracciones.Fields.Item("valorcomercial").Value/RsDetFracciones.Fields.Item("cancom02").Value
                         'xMercan     = RsDetFracciones.Fields.Item("mercancia").Value
                         xOrdfra     = RsDetFracciones.Fields.Item("ordfra02").Value
                         'xtasaIgi    = RsDetFracciones.Fields.Item("tasaIgi").Value
                         'xTipoTasa   = "IGI"

                         '*************************************************************************************************
                         'traemos las Mercancias
                         '*************************************************************************************************

                         strDetMercancias= " SELECT caco05,    " & _
                                           "         umco05,   " & _
                                           "         item05,   " & _
                                           "         desc05,   " & _
                                           "         tpmerc05, " & _
                                           "         nompro22, " & _
                                           "         ltrim(descri31) as unimed  " & _
                                           " FROM D05ARTIC  LEFT JOIN SSPROV22  ON CVEPRO22 = PROV05 " & _
                                           "                LEFT JOIN SSUMED31 ON CLAVEM31 = umco05 " & _
                                           "                LEFT JOIN SSFACT39   ON REFCIA39  = REFE05  AND  LTRIM(NUMFAC39) =LTRIM(fACT05) " & _
                                           " WHERE  refe05='" &  strRefer & "'  AND " & _
                                           "                  Agru05=" & xOrdfra
                         'response.write(strDetMercancias)
                         'response.end

                         Set RsDetMercancias = Server.CreateObject("ADODB.Recordset")
			                   RsDetMercancias.ActiveConnection = MM_EXTRANET_STRING
			                   RsDetMercancias.Source = strDetMercancias
            	           RsDetMercancias.CursorType = 0
                         RsDetMercancias.CursorLocation = 2
			                   RsDetMercancias.LockType = 1
			                   RsDetMercancias.Open()

                         if not RsDetMercancias.eof then ' si tiene fracciones
                            While not RsDetMercancias.eof ' recorremos las fracciones
                                '*************************************************************************************************
                                'dblDTA = RsDetMercancias.Fields.Item("DTA").Value

                                 xunicom     = RsDetMercancias.Fields.Item("umco05").Value
                                 xStrUniMed  = ""
                                 xCantUniMed = 1
                                 SELECT CASE xunicom
                                         CASE 15  'BARRIL
                                             xStrUniMed  = "BAR"
                                         CASE 7   'CABEZA
                                             xStrUniMed  = "CAB"
                                         CASE 16  'GRAMO NETO
                                             xStrUniMed  = "GRN"
                                        CASE 2    'GRAMOS
                                             xStrUniMed  = "GRS"
                                        CASE 12   'JUEGO
                                             xStrUniMed  = "JUE"
                                        CASE 1    'KILOGRAMOS
                                             xStrUniMed  = "KGS"
                                        CASE 13   'KILO WATT POR HORA
                                             xStrUniMed  = "KWH"
                                        CASE 10   'KILO WATTS
                                             xStrUniMed  = "KWS"
                                        CASE 8    'LITROS
                                             xStrUniMed  = "LTS"
                                        CASE 3    'METRO LINEAL
                                             xStrUniMed  = "M"
                                        CASE 4    'METRO CUADRADO
                                             xStrUniMed  = "M2"
                                        CASE 5    'METRO CUBICO
                                             xStrUniMed  = "M3"
                                        CASE 11   'MILLAR
                                             xStrUniMed  = "MIL"
                                        CASE 9    'PAR
                                             xStrUniMed  = "PAR"
                                        CASE 6    'PIEZA
                                             xStrUniMed  = "PZA"
                                        CASE 14   'TONELADA
                                             xStrUniMed  = "TON"
                                        CASE ELSE
                                             xStrUniMed  = RsDetMercancias.Fields.Item("unimed").Value
                                 END SELECT

                                 xmodelo  = RsDetMercancias.Fields.Item("item05").Value
                                 xProv    = RsDetMercancias.Fields.Item("nompro22").Value
                                 xCant    = RsDetMercancias.Fields.Item("caco05").Value * (xCantUniMed)
                                 'xPreUni  = RsDetMercancias.Fields.Item("vafa05").Value/RsDetMercancias.Fields.Item("caco05").Value
                                 'xValor   = RsDetMercancias.Fields.Item("vafa05").Value
                                 'xFactu   = RsDetMercancias.Fields.Item("fACT05").Value
                                 xMercan  = RsDetMercancias.Fields.Item("desc05").Value
                                 xTipoMer = RsDetMercancias.Fields.Item("tpmerc05").Value
                                 'xIncoterm= RsDetMercancias.Fields.Item("INCOTERM").Value
                                 'xFecpag  = RsDetMercancias.Fields.Item("FECHAFACTURA").Value
                                 'if isnull(xFecpag) then
                                 '   xFecpag  = RsDetMercancias.Fields.Item("ffactp05").Value
                                 'end if
                                 'xTipoTasa   = "IGI"
                                 'xUniMed  = RsDetMercancias.Fields.Item("unimed").Value

                                 '*************************************************************************************************
                                 '*************************************************************************************************

                                 'strRefer
                                 'xImporta
                                 'xFactmo
                                 'xTipCam
                                 'xFrac
                                 'xpaiOri
                                 'xpaiscv
                                 'xOrdfra
                                 'xtasaIgi
                                 'xunicom
                                 'xCantUniMed
                                 'xmodelo
                                 'xProv
                                 'xCant
                                 'xPreUni
                                 'xValor
                                 'xFactu
                                 'xUniMed
                                 'xMercan
                                 'xTipoMer
                                 'xIncoterm
                                 'xFecpag
                                 'xTipoTasa

                               strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xmodelo     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'NUM PARTE    Codigo del Numero de Parte
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xMercan     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'DESCRIPCION  Descripcion del numero de parte
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFrac       & "&nbsp;    </font></td>" & chr(13) & chr(10) 'CLAVE F      Fracción Arancelaria
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xTipoMer    & "&nbsp;    </font></td>" & chr(13) & chr(10) 'TIPO         Tipo de Bien
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactmo     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'F MONEDA     Factor moneda
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xStrUniMed  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'UNIDAD       Unidad del numero de parte
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &               "&nbsp;    </font></td>" & chr(13) & chr(10) 'Proyecto
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &               "&nbsp;    </font></td>" & chr(13) & chr(10) 'Numero de serie
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xImporta    & "&nbsp;    </font></td>" & chr(13) & chr(10) 'DocEntrada   Numero de pedimento
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &               "&nbsp;    </font></td>" & chr(13) & chr(10) 'fecha de alta empresa
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xpaiOri     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'ORIGEN       Pais de procedencia
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xRegimen    & "&nbsp;    </font></td>" & chr(13) & chr(10) 'TIpo de entrada Tipo de Entrada; Temporal T, Definitiva D.


                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactu      & "&nbsp;    </font></td>" & chr(13) & chr(10) 'FACTURA      Folio de la factura
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xProv       & "&nbsp;    </font></td>" & chr(13) & chr(10) 'CODIGO       Codigo del proveedor
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFecpag     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'FECHA FAC    Fecha de la factura
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xtasaIgi    & "&nbsp;    </font></td>" & chr(13) & chr(10) 'TASA         Tasa Arancelaria
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xTipoTasa   & "&nbsp;    </font></td>" & chr(13) & chr(10) 'TIPO TASA   Tipo de Tasa Arancelaria
	  	                         'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xPreUni     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'PRECIO       Precio Unitario
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xCant       & "&nbsp;    </font></td>" & chr(13) & chr(10) 'CANTIDAD     Cantidad Facturada
		                           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">       1                      </font></td>" & chr(13) & chr(10) 'CONVERSION   Factor de Conversión
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xpaiscv     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'VENDEDOR     Pais Vendedor
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " 0 "       & "&nbsp;    </font></td>" & chr(13) & chr(10) 'F PAGO       Forma de Pago
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xIncoterm   & "&nbsp;    </font></td>" & chr(13) & chr(10) 'INCOTERM     Termino Internacional de Comercio

                               'NUM PARTE
                               'DESCRIPCION
                               'CLAVE F
                               'TIPO
                               'F MONEDA
                               'UNIDAD
                               'PROYECTO
                               'NUMERO DE SERIE
                               'DocEntrada
                               'Fecha de alta
                               'Pais de Procedencia
                               'Tipo de Entrada





                               strHTML = strHTML&"</tr>"& chr(13) & chr(10)

                                RsDetMercancias.movenext
                            wend
                         else
                                '
                                'Fecha de la factura
                                'Unidad del numero de parte
                                'INCOTERM

                               ' si no tiene Mercancias
                               'strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xImporta    & "&nbsp;    </font></td>" & chr(13) & chr(10) 'Folio del pedimento
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactu      & "&nbsp;    </font></td>" & chr(13) & chr(10) 'FACTURA      Folio de la factura
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xProv       & "&nbsp;    </font></td>" & chr(13) & chr(10) 'CODIGO       Codigo del proveedor
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'FECHA FAC    Fecha de la factura
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactmo     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'F MONEDA     Factor moneda
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'NUM PARTE    Codigo del Numero de Parte
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xMercan     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'DESCRIPCION  Descripcion del numero de parte
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'TIPO         Tipo de Bien
      		                     'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFrac       & "&nbsp;    </font></td>" & chr(13) & chr(10) 'CLAVE F      Fracción Arancelaria
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xtasaIgi    & "&nbsp;    </font></td>" & chr(13) & chr(10) 'TASA         Tasa Arancelaria
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xTipoTasa  & "&nbsp;    </font></td>" & chr(13) & chr(10)  'TIPO TASA   Tipo de Tasa Arancelaria
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'UNIDAD       Unidad del numero de parte
	  	                         'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xPreUni     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'PRECIO       Precio Unitario
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xCant       & "&nbsp;    </font></td>" & chr(13) & chr(10) 'CANTIDAD     Cantidad Facturada
		                           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">       1                      </font></td>" & chr(13) & chr(10) 'CONVERSION   Factor de Conversión
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xpaiOri     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'ORIGEN       Pais de Origen
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xpaiscv     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'VENDEDOR     Pais Vendedor
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " 0 "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'F PAGO       Forma de Pago
                               'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'INCOTERM     Termino Internacional de Comercio
                               'strHTML = strHTML&"</tr>"& chr(13) & chr(10)

                         end if
                         RsDetMercancias.close
                         set RsDetMercancias = Nothing
                         '**********************************************************************************
                         '***** Final de mercancias
                         '**********************************************************************************

                     RsDetFracciones.movenext
                   wend
                  else  ' si no tiene fracciones
                      strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xImporta    & "&nbsp;    </font></td>" & chr(13) & chr(10) 'Folio del pedimento
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactu      & "&nbsp;    </font></td>" & chr(13) & chr(10) 'FACTURA      Folio de la factura
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xProv       & "&nbsp;    </font></td>" & chr(13) & chr(10) 'CODIGO       Codigo del proveedor
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'FECHA FAC    Fecha de la factura
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactmo     & "&nbsp;    </font></td>" & chr(13) & chr(10) 'F MONEDA     Factor moneda
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'NUM PARTE    Codigo del Numero de Parte
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'DESCRIPCION  Descripcion del numero de parte
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'TIPO         Tipo de Bien
      		            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'CLAVE F      Fracción Arancelaria
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'TASA         Tasa Arancelaria
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10)  'TIPO TASA   Tipo de Tasa Arancelaria
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'UNIDAD       Unidad del numero de parte
	  	                strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'PRECIO       Precio Unitario
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'CANTIDAD     Cantidad Facturada
		                  strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">       1                      </font></td>" & chr(13) & chr(10) 'CONVERSION   Factor de Conversión
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'ORIGEN       Pais de Origen
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'VENDEDOR     Pais Vendedor
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'F PAGO       Forma de Pago
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "&nbsp;    </font></td>" & chr(13) & chr(10) 'INCOTERM     Termino Internacional de Comercio
                      strHTML = strHTML&"</tr>"& chr(13) & chr(10)

                  end if
                  RsDetFracciones.close
                  set RsDetFracciones = Nothing
                  '**********************************************************************************
                  '*****   Final de fracciones
                  '**********************************************************************************

             RsRep.movenext
           Wend

             strHTML = strHTML & "</table>"& chr(13) & chr(10)
          else

          strHTML = strHTML & "<tr> <td colspan=12> NO EXISTEN REGISTROS </td></tr>"
          strHTML = strHTML & "</table>"& chr(13) & chr(10)
     '   response.Write(strHTML)


          end if





          RsRep.close
          Set RsRep = Nothing
          'Se pinta todo el HTML formado
          response.Write(strHTML)
          if strHTML = "" then
             strHTML = "NO EXISTEN REGISTROS"
             response.Write(strHTML)
          end if

     'else
     '   strHTML = "NO EXISTEN REGISTROS"
     '   response.Write(strHTML)
     end if

  end if ' FIN ENCABEZADO/DETALLE

else
	select case strCodError
    case "1"
	   strMenjError = "Campo en Blanco Especifique!.."
	case "5","6"
	   strMenjError = "Fechas Erroneas, Verifique!"
	case "7"
	   strMenjError = "Registros No Encontrados!"
	end select
	%>


	<table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
	  <tr>
		<td><%=strMenjError%></td>
	  </tr>
	</table>
<br>
<%end if%>
