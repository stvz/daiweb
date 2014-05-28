<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->


<%
' ESTE ASP ES EL SEGUNDO Y ES PARA ADMINISTRADORES
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

Response.Buffer = TRUE

strUsuario = request.Form("user")
strTipoUsuario = request.Form("TipoUser")

strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
permi2 = PermisoClientesTabla("B",Session("GAduana") ,strPermisos,"clie31")

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
rbnTipoReporte=trim(request.Form("rbnTipoReporte"))
'*****************************************************
strDescripcion=trim(request.Form("txtDescripcion"))
strDateIni2=trim(request.Form("txtDateIni2"))
strDateFin2=trim(request.Form("txtDateFin2"))
strTipoPedimento2=trim(request.Form("rbnTipoDate2"))
strTipoFiltro=trim(request.Form("TipoFiltro"))


'rbnTipoReporte = 1
'strTipoPedimento = 1

'strDateIni=trim(request.Form("txtDateIni"))
'strDateFin=trim(request.Form("txtDateFin"))
'strTipoPedimento=trim(request.Form("rbnTipoDate"))


  if rbnTipoReporte  = "1" then 'Si es el encabezado
     if strTipoPedimento  = "1" then
        Response.Addheader "Content-Disposition", "attachment;filename=Reporte_Encabezado_Pedimentos_importacion.xls"
     else
        if strTipoPedimento  = "2" then
           Response.Addheader "Content-Disposition", "attachment;filename=Reporte_Encabezado_Pedimentos_exportacion.xls"
        end if
     end if
  else 'El detalle
     if strTipoPedimento  = "1" then
        Response.Addheader "Content-Disposition", "attachment;filename=Reporte_Facturas_Pedimentos_importacion.xls"
     else
        if strTipoPedimento  = "2" then
           Response.Addheader "Content-Disposition", "attachment;filename=Reporte_Facturas_Pedimentos_exportacion.xls"
        end if
     end if
  end if
  Response.ContentType = "application/vnd.ms-excel"
  Server.ScriptTimeOut=100000




'Reporte_Facturas_Pedimentos_importacion   Facturas PEDIMENTO DE IMPORTACION
'Reporte_Encabezado_Pedimentos_importacion Encabezado PEDIMENTO DE IMPORTACION
'Reporte_Facturas_Pedimentos_exportacion   Facturas PEDIMENTO DE EXPORTACION
'Reporte_Encabezado_Pedimentos_exportacion Encabezado PEDIMENTO DE EXPORTACION
'Response.Addheader "Content-Disposition", "attachment;filename=hella.xls"




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

	'********************************************************************************************************************************************************
	if rbnTipoReporte  = "1" then 'ENCABEZADO

		if strTipoPedimento  = "1" then
			 tmpTipo = "IMPORTACION"
			'strSQL = "SELECT tipopr01, valmer01,factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecent01 FROM ssdagi01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
			 strSQL = "SELECT concat(concat(concat(concat(concat(concat(ltrim(substring(year(FECPAG01),3,2)),'-'),adusec01 ),'-'),PATENT01),'-'),NUMPED01) as IMPORTA," & _
					  "       adusec01 as Aduana,        " & _
					  "       fecpag01 as pago,          " & _
					  "       TIPCAM01 as TipoCambio,    " & _
					  "       cveped01 as clavePedimento," & _
					  "       FLETES01 as FLETE,         " & _
					  "       SEGROS01 as SEGUROS,       " & _
					  "       Embala01 as embalaje,      " & _
					  "       Incble01 as OtrosIncbles,  " & _
					  "       anexol01 as observa,       " & _
					  "       refcia01 as Referencia,    " & _
					  "       FACTMO01   as FactorMoneda,     " & _
					  "       sum(vaduan02)   as valoraduana, " & _
					  "       sum(vmerme02)   as valorComerExtra,  " & _
					  "       sum(vmerme02*FACTMO01) as valorComerDls,      " & _
					  "       sum(vmerme02*FACTMO01*TIPCAM01) as valorMN " & _
					  "FROM ssdagi01, " & _
					  "     ssfrac02  " & _
					  "WHERE  Refcia01 = refcia02  and " & _
					  "       (firmae01   <> '')   and " & _
					  "       fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
					  "       fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND " & _
					  "       LTRIM(refcia01) <> 'GABBY' " & _
					  Permi & _
					  " GROUP BY REFCIA01"
					  
		
		elseif strTipoPedimento  = "2" then
			tmpTipo = "EXPORTACION"
			'strSQL = "SELECT tipopr01, factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01 FROM ssdage01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
			 strSQL = "SELECT concat(concat(concat(concat(concat(concat(ltrim(substring(year(FECPAG01),3,2)),'-'),adusec01 ),'-'),PATENT01),'-'),NUMPED01) as IMPORTA," & _
					  "       adusec01 as Aduana,        " & _
					  "       fecpag01 as pago,          " & _
					  "       TIPCAM01 as TipoCambio,    " & _
					  "       cveped01 as clavePedimento," & _
					  "       FLETES01 as FLETE,         " & _
					  "       SEGROS01 as SEGUROS,       " & _
					  "       Embala01 as embalaje,      " & _
					  "       Incble01 as OtrosIncbles,  " & _
					  "       anexol01 as observa,       " & _
					  "       refcia01 as Referencia,    " & _
					  "       FACTMO01   as FactorMoneda,     " & _
					  "       sum(vaduan02)   as valoraduana, " & _
					  "       sum(vmerme02)   as valorComerExtra,  " & _
					  "       sum(vmerme02*FACTMO01) as valorComerDls,      " & _
					  "       sum(vmerme02*FACTMO01*TIPCAM01) as valorMN " & _
					  "FROM ssdage01, " & _
					  "     ssfrac02  " & _
					  "WHERE  Refcia01 = refcia02  and " & _
					  "       (firmae01   <> '')   and " & _
					  "       fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
					  "       fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND " & _
					  "       LTRIM(refcia01) <> 'GABBY'  " & _
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


			if not RsRep.eof then

				' Comienza el HTML, se pintan los titulos de las columnas

				strHTML = strHTML & " <p> &nbsp; </p>"
				strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">GRUPO REYES KURI, S.C. </font></strong> <br> "

				if tmpTipo = "EXPORTACION" then
					strHTML = strHTML & "<strong><font color=""#969696"" size=""3"" face=""Arial, Helvetica, sans-serif""> Reporte Encabezado de Operaciones de Exportación del " & strDateIni & " al " & strDateFin & " </font></strong>"
				else
					strHTML = strHTML & "<strong><font color=""#969696"" size=""3"" face=""Arial, Helvetica, sans-serif""> Reporte Encabezado de Operaciones de Importación del " & strDateIni & " al " & strDateFin & " </font></strong>"
				end if

				strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
				strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)

			   'oExcel.Cells(intColumna,18).Value = "PREVALIDACION"

				if tmpTipo = "EXPORTACION" then
					strHTML = strHTML & "<td width=""95""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> EXPORTA         </font></strong></td>" & chr(13) & chr(10)
				else
					strHTML = strHTML & "<td width=""95""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IMPORTA         </font></strong></td>" & chr(13) & chr(10)
				end if

				strHTML = strHTML & "<td width=""70"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ADUANAS         </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""60"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA           </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""80""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TIPO DE CAMBIO  </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""60""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IVA             </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""40"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CLAVE           </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FLETES          </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""70"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SEGUROS         </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> EMBALAJE        </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""60""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> OTROS           </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""40"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DTA             </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""120""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VALOR COMERCIAL </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""120""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VALOR EN ADUANA </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""100"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> OBSERVACIONES   </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""90"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CONSOLIDADO     </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VIRTUAL         </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "<td width=""95""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PREVALIDACION   </font></strong></td>" & chr(13) & chr(10)
				strHTML = strHTML & "</tr>"& chr(13) & chr(10)


				While NOT RsRep.EOF

					strRefer = RsRep.Fields.Item("Referencia").Value

					'traemos los impuestos
					strImpuestos= " SELECT SUM(IF(ltrim(cveimp36)='1' ,import36,0)) as DTA, " & _
								   "        SUM(IF(ltrim(cveimp36)='3' ,import36,0)) as IVA, " & _
								   "        SUM(IF(ltrim(cveimp36)='15',import36,0)) as PRV  " & _
								   " from SSCONT36 " & _
								   " WHERE refcia36='" &  strRefer & "' " & _
								   "      AND FPAGOI36 = 0 " & _
								   " GROUP BY refcia36 "
					'response.write(strHTML)
					'response.write(strImpuestos)
					'response.end

					Set RsImpuestos = Server.CreateObject("ADODB.Recordset")
						RsImpuestos.ActiveConnection = MM_EXTRANET_STRING
						RsImpuestos.Source = strImpuestos
						RsImpuestos.CursorType = 0
						RsImpuestos.CursorLocation = 2
						RsImpuestos.LockType = 1
						RsImpuestos.Open()

					if not RsImpuestos.eof then
						While not RsImpuestos.eof
							dblDTA = RsImpuestos.Fields.Item("DTA").Value
							dblIVA = RsImpuestos.Fields.Item("IVA").Value
							dblPRV = RsImpuestos.Fields.Item("PRV").Value
							RsImpuestos.movenext
						wend
					end if
					RsImpuestos.close
					set RsImpuestos = Nothing


					'vamos a ssfrac02 por valor aduana y valorcomercial
					strVirtual= " SELECT  SUM(if(cveide12='PX' OR cveide12='AE' OR cveide12='MQ' OR cveide12='IM',1,0))  AS VIRTUAL " & _
								  " FROM SSIPAR12 " & _
								  " WHERE refcia12 ='" &  strRefer & "' " & _
								  " GROUP BY refcia12 "
					'response.write(strHTML)
					'response.write(strVirtual)
					'response.end

					Set RsVirtual = Server.CreateObject("ADODB.Recordset")
						RsVirtual.ActiveConnection = MM_EXTRANET_STRING
						RsVirtual.Source = strVirtual
						RsVirtual.CursorType = 0
						RsVirtual.CursorLocation = 2
						RsVirtual.LockType = 1
						RsVirtual.Open()

					dblVirtual = 0
					if not RsVirtual.eof then
					   While not RsVirtual.eof
							dblVirtual = RsVirtual.Fields.Item("VIRTUAL").Value
							RsVirtual.movenext
					   wend
					end if
					RsVirtual.close
					set RsVirtual = Nothing
					
					'Esta validacion no debe de cambiarse o retorna un error, se debe validar el 0 como texto
					if dblVirtual = "0" then
						strvirtual = ""
					else
						strvirtual = "S"
					end if
					
					'*************************************************************************************************
					 'Verificar si es consolidado o no

					 ' strConsolida= " SELECT NUMCON40  AS CONSOLIDA" & _
					 '               " FROM SSCONT40 " & _
					 '               " WHERE REFCIA40  ='" &  strRefer & "' AND" & _
					 '               " NUMCON40 <> '' "
					 ' Set RsConsolida = Server.CreateObject("ADODB.Recordset")
						   ' RsConsolida.ActiveConnection = MM_EXTRANET_STRING
						   ' RsConsolida.Source = strConsolida
					   ' RsConsolida.CursorType = 0
						   ' RsConsolida.CursorLocation = 2
						   ' RsConsolida.LockType = 1
						   ' RsConsolida.Open()
						   ' if not RsConsolida.eof then
					 '   While not RsConsolida.eof
					 '     Strconsolida = "S"
					 '     RsConsolida.movenext
					 '   wend
					 ' else
					 '    Strconsolida = "N"
					 ' end if
					 ' RsConsolida.close
					 ' set RsConsolida = Nothing


					'*************************************************************************************************
					
					strHTML = strHTML&"<tr>" & chr(13) & chr(10)
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("IMPORTA").Value           & "   </font></td>" & chr(13) & chr(10) 'IMPORTA
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("Aduana").Value            & "   </font></td>" & chr(13) & chr(10) 'ADUANAS
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("pago").Value              & "   </font></td>" & chr(13) & chr(10) 'FECHA
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("TipoCambio").Value        & "   </font></td>" & chr(13) & chr(10) 'TIPO DE CAMBIO
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & dblIVA                                       & "   </font></td>" & chr(13) & chr(10) 'IVA
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("clavePedimento").Value    & "   </font></td>" & chr(13) & chr(10) 'CLAVE
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("FLETE").Value             & "   </font></td>" & chr(13) & chr(10) 'FLETES
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("SEGUROS").Value           & "   </font></td>" & chr(13) & chr(10) 'SEGUROS
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("embalaje").Value          & "   </font></td>" & chr(13) & chr(10) 'EMBALAJE
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("OtrosIncbles").Value      & "   </font></td>" & chr(13) & chr(10) 'OTROS
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & dblDTA                                       & "   </font></td>" & chr(13) & chr(10) 'DTA
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & Round(RsRep.Fields.Item("valorMN").Value,0)  & "   </font></td>" & chr(13) & chr(10) 'VALOR COMERCIAL
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("valoraduana").Value       & "   </font></td>" & chr(13) & chr(10) 'VALOR EN ADUANA
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &  "PL4"                                       & "   </font></td>" & chr(13) & chr(10) 'OBSERVACIONES
					 
					'if strvirtual = "S" then
						'Strconsolida = "S"
					'else
						'Strconsolida = "N"
					'end if
					
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">                                                       </font></td>" & chr(13) & chr(10) 'CONSOLIDADO
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &strvirtual                                    & "  </font></td>" & chr(13) & chr(10) 'VIRTUAL
					strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &dblPRV                                        & "  </font></td>" & chr(13) & chr(10) 'PREVALIDACION
					strHTML = strHTML&"</tr>"& chr(13) & chr(10)
					
					RsRep.movenext

				Wend

				strHTML = strHTML & "</table>"& chr(13) & chr(10)
			end if

			'Se cierran las conexiones
			RsRep.close
			Set RsRep = Nothing
			
			'Se valida si existio algun registro
			if strHTML = "" then
			   strHTML = "NO EXISTEN REGISTROS__"
			end if

			'Se pinta todo el HTML formado
			response.Write(strHTML)

		else
			strHTML = "NO EXISTEN REGISTROS"
			response.Write(strHTML)
		end if

	'******************************************************************************************************************************
	else 'DETALLE

		if strTipoPedimento  = "1" then
			tmpTipo = "IMPORTACION"
			strSQL =   " SELECT concat(concat(concat(concat(concat(concat(ltrim(substring(year(FECPAG01),3,2)),'-'),adusec01 ),'-'),PATENT01),'-'),NUMPED01) as IMPORTA, " & _
                   "        desf0101 as facturas,   " & _
                   "        nompro01 as proveedor,  " & _
                   "        ltrim(refcia01) as Referencia, " & _
                   "        TIPCAM01 as TipoCambio, " & _
                   "        FACTMO01 " & _
                   "  from ssdagi01  " & _
                   "  where          " & _
                   "        firmae01 <> ''   and  " & _
                   "       fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                   "       fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND " & _
                   "       LTRIM(refcia01) <> 'GABBY' " & _
                   Permi & _
                 " GROUP BY REFCIA01"
				  

		end if
		if strTipoPedimento  = "2" then
			tmpTipo = "EXPORTACION"
			strSQL =   " SELECT concat(concat(concat(concat(concat(concat(ltrim(substring(year(FECPAG01),3,2)),'-'),adusec01 ),'-'),PATENT01),'-'),NUMPED01) as IMPORTA, " & _
                    "        desf0101 as facturas,   " & _
                    "        nompro01 as proveedor,  " & _
                    "        ltrim(refcia01) as Referencia, " & _
                    "        TIPCAM01 as TipoCambio, " & _
                    "        FACTMO01 " & _
                    "  from ssdage01  " & _
                    "  where          " & _
                    "        firmae01 <> ''   and  " & _
                    "       fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                    "       fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND " & _
                    "       LTRIM(refcia01) <> 'GABBY' " & _
                    Permi & _
                  " GROUP BY REFCIA01"

                  '"  where cveped01 <> 'R1' and  " & _

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

	        if not RsRep.eof then

             ' Comienza el HTML, se pintan los titulos de las columnas
             'strHTML = strHTML & " <p> <img src='../../ext-Images/Gifs/abbot.gif'> </p>"
             'strHTML = strHTML & " <p> <img width='181' eight='38'  src='http://10.66.1.4/PortalMySQL/Extranet/ext-Images/Gifs/abbot.gif'> </p> <P>&nbsp;</P>"

             strHTML = strHTML & " <p> &nbsp; </p>"
             strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">GRUPO REYES KURI, S.C. </font></strong> <br> "

             if tmpTipo = "EXPORTACION" then
	  	         strHTML = strHTML & "<strong><font color=""#969696"" size=""3"" face=""Arial, Helvetica, sans-serif""> Reporte facturas de Operaciones de Exportación del " & strDateIni & " al " & strDateFin & " </font></strong>"
             else
               strHTML = strHTML & "<strong><font color=""#969696"" size=""3"" face=""Arial, Helvetica, sans-serif""> Reporte facturas de Operaciones de Importación del " & strDateIni & " al " & strDateFin & " </font></strong>"
             end if


             strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	           strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)

             if tmpTipo = "EXPORTACION" then
		           strHTML = strHTML & "<td width=""95""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> EXPORTA         </font></strong></td>" & chr(13) & chr(10) 'Folio del pedimento
             else
               strHTML = strHTML & "<td width=""95""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IMPORTA         </font></strong></td>" & chr(13) & chr(10) 'Folio del pedimento
             end if

             strHTML = strHTML & "<td width=""70""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FACTURA       </font></strong></td>" & chr(13) & chr(10) ' Folio de la factura
             strHTML = strHTML & "<td width=""60""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CODIGO        </font></strong></td>" & chr(13) & chr(10) ' Codigo del proveedor
             strHTML = strHTML & "<td width=""80""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA FAC     </font></strong></td>" & chr(13) & chr(10) ' Fecha de la factura
             strHTML = strHTML & "<td width=""40""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> F MONEDA      </font></strong></td>" & chr(13) & chr(10) ' Factor moneda
             strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> NUM PARTE     </font></strong></td>" & chr(13) & chr(10) ' Codigo del Numero de Parte
             strHTML = strHTML & "<td width=""70""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESCRIPCION   </font></strong></td>" & chr(13) & chr(10) ' Descripcion del numero de parte
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TIPO          </font></strong></td>" & chr(13) & chr(10) ' Tipo de Bien
      		 strHTML = strHTML & "<td width=""60""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CLAVE F       </font></strong></td>" & chr(13) & chr(10) ' Fracción Arancelaria
             strHTML = strHTML & "<td width=""40""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TASA          </font></strong></td>" & chr(13) & chr(10) ' Tasa Arancelaria
             strHTML = strHTML & "<td width=""95""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TIPO TASA     </font></strong></td>" & chr(13) & chr(10) ' Tipo de Tasa Arancelaria
             strHTML = strHTML & "<td width=""100"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> UNIDAD        </font></strong></td>" & chr(13) & chr(10) ' Unidad del numero de parte
	  	     strHTML = strHTML & "<td width=""90""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PRECIO UNITARIO     </font></strong></td>" & chr(13) & chr(10) ' Precio Unitario
             strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CANTIDAD      </font></strong></td>" & chr(13) & chr(10) ' Cantidad Facturada
		     strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CONVERSION    </font></strong></td>" & chr(13) & chr(10) ' Factor de Conversión
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ORIGEN        </font></strong></td>" & chr(13) & chr(10) ' Pais de Origen
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VENDEDOR      </font></strong></td>" & chr(13) & chr(10) ' Pais Vendedor
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> F PAGO        </font></strong></td>" & chr(13) & chr(10) ' Forma de Pago
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> INCOTERM      </font></strong></td>" & chr(13) & chr(10) ' Termino Internacional de Comercio

             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> AcuerdoCom    </font></strong></td>" & chr(13) & chr(10) ' Acuerdo Comercial (Agregado el 25/11/09)
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TLCAN         </font></strong></td>" & chr(13) & chr(10) ' Si es TLCAN  (Agregado el 25/11/09)
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TLCUEM        </font></strong></td>" & chr(13) & chr(10) ' Si es TLCUEM (Agregado el 25/11/09)
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TLCAELC       </font></strong></td>" & chr(13) & chr(10) ' Si es TLC (Agregado el 25/11/09)
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> JustTLCAN     </font></strong></td>" & chr(13) & chr(10) ' La justificacion si es TLCAN
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> JustTLCUEM    </font></strong></td>" & chr(13) & chr(10) ' La justificacion si es TLCUEM
             strHTML = strHTML & "<td width=""75""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> JustTLCAELC   </font></strong></td>" & chr(13) & chr(10) ' La justificacion si es TLC



             strHTML = strHTML & "</tr>"& chr(13) & chr(10)


           While NOT RsRep.EOF 'recorremos el cursor de operaciones
                  xRefer     =  RsRep.Fields.Item("Referencia").Value
                  strRefer   =  RsRep.Fields.Item("Referencia").Value
                  xImporta   =  RsRep.Fields.Item("IMPORTA").Value
                  xFactu     =  RsRep.Fields.Item("facturas").Value
                  xProv      =  RsRep.Fields.Item("proveedor").Value
                  xTipCam    =  RsRep.Fields.Item("TipoCambio").Value

                  '*************************************************************************************************
                  'traemos las Fracciones
                  strDetFracciones= " SELECT CONCAT(SUBSTRING(fraarn02,1,4), concat('.',concat(SUBSTRING(fraarn02,5,2),concat('.',SUBSTRING(fraarn02,7,2))) ) ) as fraarn02, " & _
                                    "        cancom02,                      " & _
                                    "        u_medc02,                      " & _
                                    "        paiOri02 as PaisOrigenDestino, " & _
                                    "        paiscv02 as PaisVendedor,      " & _
                                    "        d_mer102   as mercancia,       " & _
                                    "        (vmerme02) as valorcomercial,  " & _
                                    "        PREUNI02,                      " & _
                                    "        ordfra02,                      " & _
                                    "        i_adv102,                       " & _
                                    "        p_adv102,                      " & _
                                    "        tasadv02,                      " & _
                                    "        tt_adv02                       " & _
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
                         xMercan     = RsDetFracciones.Fields.Item("mercancia").Value
                         xOrdfra     = RsDetFracciones.Fields.Item("ordfra02").Value
                         'xtasaIgi    = RsDetFracciones.Fields.Item("tasaIgi").Value
                         'xTipoTasa   = "IGI"
                         ImporteIGI   = RsDetFracciones.Fields.Item("i_adv102").Value ' Importe IGI
                         FormaPagoIGI = RsDetFracciones.Fields.Item("p_adv102").Value ' Forma de Pago IGI
                         TasaIGI      = RsDetFracciones.Fields.Item("tasadv02").Value ' Tasa o Porcentaje del IGI

                     '*************************************************************************************************************************************************
                     ' Buscamos tasa y tipo de Tasa en los permisos
                     '*************************************************************************************************************************************************


                        'Response.Write("Test")
                        'Response.End

                         TipoTasa = ""
                         ComplTipoTasa = ""

                         'sqlpermisos= " select refcia12, cveide12 as TipoTasa "&_
                         sqlpermisos= " SELECT REFCIA12, CVEIDE12 AS TIPOTASA , COMIDE12 "&_
                                      " FROM SSIPAR12 " &_
                                      " WHERE REFCIA12 = '"&strRefer&"' AND ORDFRA12 = "&xOrdfra &" "&_
                                      "   AND CVEIDE12 IN ('PS','TL')  "&_
                                      " GROUP BY refcia12 "

                         'Response.Write(sqlpermisos)
                         'Response.End


                         set RSPermiso = server.CreateObject("ADODB.Recordset")
                         RSPermiso.ActiveConnection = MM_EXTRANET_STRING
                         RSPermiso.Source= sqlpermisos
                         RSPermiso.CursorType = 0
                         RSPermiso.CursorLocation = 2
                         RSPermiso.LockType = 1
                         RSPermiso.Open()
                         pp=0
                         while not RSPermiso.eof
                            if pp=0 then
                                TipoTasa      = RSPermiso.fields.item("TipoTasa").value
                                ComplTipoTasa = RSPermiso.fields.item("COMIDE12").value
                                pp=1
                            else
                                TipoTasa      = TipoTasa      & " , "& RSPermiso.fields.item("TipoTasa").value
                                ComplTipoTasa = ComplTipoTasa & " , "& RSPermiso.fields.item("COMIDE12").value
                            end if
                          RSPermiso.movenext
                         wend
                         RSPermiso.close
                         set RSPermiso= nothing
                     '*********************************************************************************************

                         'AcuerdoCom
                         'TLCAN
                         'TLCUEM
                         'TLCAELC
                         'JustTLCAN
                         'JustTLCUEM
                         'JustTLCAELC

                         StrAcuerdoCom  = ""
                         StrTLCAN       = ""
                         StrTLCUEM      = ""
                         StrTLCAELC     = ""
                         StrJustTLCAN   = ""
                         StrJustTLCUEM  = ""
                         StrJustTLCAELC = ""

                         if TipoTasa = "PS" then ' PS
                             StrAcuerdoCom  = ""
                             StrTLCAN       = "N"
                             StrTLCUEM      = "N"
                             StrTLCAELC     = "N"
                             StrJustTLCAN   = "15"
                             StrJustTLCUEM  = "23"
                             StrJustTLCAELC = "5"
                         end if

                         if TipoTasa = "TL" then ' TL
                            if ComplTipoTasa = "EMU" OR ComplTipoTasa = "ESP" OR ComplTipoTasa = "CZE" THEN 'TLCUEM

                               TipoTasa       = "UE"
                               StrAcuerdoCom  = "UE"
                               StrTLCAN       = "N"
                               StrTLCUEM      = "N"
                               StrTLCAELC     = "N"
                               StrJustTLCAN   = "15"
                               StrJustTLCUEM  = "20"
                               StrJustTLCAELC = "5"

                            else
                               if ComplTipoTasa = "USA" OR ComplTipoTasa = "CAN" THEN'TLCAN

                                   TipoTasa       = "AN"
                                   StrAcuerdoCom  = "AN"
                                   StrTLCAN       = "N"
                                   StrTLCUEM      = "N"
                                   StrTLCAELC     = "N"
                                   StrJustTLCAN   = "10"
                                   StrJustTLCUEM  = "23"
                                   StrJustTLCAELC = "5"

                               else
                                  if ComplTipoTasa = "NOR" OR ComplTipoTasa = "CHE" THEN 'TLCAELC

                                     TipoTasa       = "AE"
                                     StrAcuerdoCom  = "AE"
                                     StrTLCAN       = "N"
                                     StrTLCUEM      = "N"
                                     StrTLCAELC     = "N"
                                     StrJustTLCAN   = "15"
                                     StrJustTLCUEM  = "23"
                                     StrJustTLCAELC = "2"

                                  end if
                               end if
                            end if


                         end if

                         if TipoTasa <> "PS" AND TipoTasa <> "TL" AND TipoTasa <> "UE" AND TipoTasa <> "AN" AND TipoTasa <> "AE" AND ComplTipoTasa <> "ESP" AND ComplTipoTasa <> "CZE" then
                             TipoTasa = "TG"
                             StrAcuerdoCom  = ""
                             StrTLCAN       = "N"
                             StrTLCUEM      = "N"
                             StrTLCAELC     = "N"
                             StrJustTLCAN   = "15"
                             StrJustTLCUEM  = "23"
                             StrJustTLCAELC = "5"
                         end if

                         ImporteIGI   = RsDetFracciones.Fields.Item("i_adv102").Value ' Importe IGI
                         FormaPagoIGI = RsDetFracciones.Fields.Item("p_adv102").Value ' Forma de Pago IGI
                         TasaIGI      = RsDetFracciones.Fields.Item("tasadv02").Value ' Tasa o Porcentaje del IGI
                         'TipoTasa     = "TG" 'Tipo de Tasa


                     if cdbl( TasaIGI ) <= 0 then
                         ImporteIGI   = "0" ' Importe IGI
                         FormaPagoIGI = "0" ' Forma de Pago IGI
                         TasaIGI      = "0" ' Tasa o Porcentaje del IGI
                         'TipoTasa     = "0" ' Tipo de Tasa
                     end if

						
						 ' strPrecUni = " SELECT caco	



                         '*************************************************************************************************
                         'traemos las Mercancias
                         '*************************************************************************************************
							
							if strTipoPedimento = "1" then
							    varP = "i"
								else
								varP = "e"
							end if
                         strDetMercancias= " SELECT caco05,                              " & _
                                           "        umco05,                              " & _
                                           "        vafa05,                              " & _
                                           "        item05,                              " & _
                                           "        fACT05,                              " & _
                                           "        ffactp05,                            " & _
                                           "        year(ffactp05) as aniomerc,          " & _
                                           "        year( ffactp05 ) as aniomerc ,       " & _
                                           "        month( ffactp05 ) as mesmerc ,       " & _
                                           "        DAYOFMONTH( ffactp05 )  as diamerc , " & _
                                           "        PROV05,                              " & _
                                           "        desc05,                              " & _
                                           "        tpmerc05,                            " & _
                                           "        nompro22,                            " & _
                                           "        NPSCLI22,                            " & _
                                           "        ltrim(descri31) as unimed ,          " & _
                                           "        FECFAC39        as FECHAFACTURA,     " & _
                                           "        year( FECFAC39 ) as anio ,           " & _
                                           "        month( FECFAC39 ) as mes ,           " & _
                                           "        DAYOFMONTH(FECFAC39)  as dia ,       " & _
										   "        round((valmex39 * facmon39),2) as 'ValorComercial' ,         " & _
                                           "        TERFAC39        as  INCOTERM,         " & _
										   "        facmon39       as  FACMONEDA,         " & _
										   "        factmo01       as  FACPED,         " & _
										   ' "        if(caco05 is not null and caco05 > 0,round((vafa05/caco05),12),-1) as 'xPreUniok', " & _
										   ' "		if(caco05 is not null and caco05 > 0, round(((vafa05*factmo01)/caco05),12), -1) as xPreUni1, " & _
										   ' "		if(caco05 is not null and caco05 > 0, round(((vafa05/facmon39)/caco05),12), -1) as xPreUni2, " & _
										   "		if(caco05 is not null and caco05 > 0, round(((factmo01/facmon39)*vafa05)/caco05,12), -1) as preuni, " & _
										   "		monfac39	as mExt						  " & _
                                           " FROM D05ARTIC  LEFT JOIN SSPROV22  ON CVEPRO22 = PROV05 " & _
                                           "                LEFT JOIN SSUMED31  ON CLAVEM31 = umco05 " & _
										   "                LEFT JOIN SSFACT39  ON REFCIA39 = REFE05  AND  LTRIM(NUMFAC39) = LTRIM(fACT05) " & _
										   "				LEFT JOIN SSDAG"& varP &"01 ON REFCIA01 = REFCIA39 " & _
                                           " WHERE  refe05='" &  strRefer & "'  AND " & _
                                           " Agru05=" & xOrdfra
										   
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
                                 xCantUniMed = 1
                                 xStrUniMed  = ""

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
                                        CASE 18   'CIENTOS
                                             xStrUniMed  = "CNT"
                                        CASE 20   'CAJA
                                             xStrUniMed  = "CJA"
                                        CASE ELSE
                                             xStrUniMed  = RsDetMercancias.Fields.Item("unimed").Value
                                 END SELECT

                                 xmodelo  = RsDetMercancias.Fields.Item("item05").Value
                                 'xProv    = RsDetMercancias.Fields.Item("nompro22").Value
                                 xProv    = RsDetMercancias.Fields.Item("NPSCLI22").Value

                                 if xProv = "" then
                                    xProv    = RsDetMercancias.Fields.Item("nompro22").Value
                                 end if

								 ' if RsDetMercancias.Fields.Item("FACMONEDA").value <> RsDetMercancias.Fields.Item("FACPED").value then
									' if RsDetMercancias.Fields.Item("FACMONEDA").value = 1.00 then
										' xPreUni = RsDetMercancias.Fields.Item("xPreUni1").value
										' Else
										' xPreUni = RsDetMercancias.Fields.Item("xPreUni2").value
									' End if
								' Else
								    'ok
									xPreUni = RsDetMercancias.Fields.Item("preuni").value
								' End if

								 xCant    = RsDetMercancias.Fields.Item("caco05").Value
                                 xFactmo    =  RsDetMercancias.Fields.Item("FACMONEDA").Value
								 xValor   = RsDetMercancias.Fields.Item("ValorComercial").Value
                                 xFactu   = RsDetMercancias.Fields.Item("fACT05").Value
                                 xUniMed  = RsDetMercancias.Fields.Item("unimed").Value
                                 xMercan  = RsDetMercancias.Fields.Item("desc05").Value
                                 xTipoMer = RsDetMercancias.Fields.Item("tpmerc05").Value
                                 xIncoterm= RsDetMercancias.Fields.Item("INCOTERM").Value
                                 xFecpag = RsDetMercancias.Fields.Item("dia").Value & "/" & RsDetMercancias.Fields.Item("mes").Value & "/" & RsDetMercancias.Fields.Item("anio").Value

                                 'xTipoTasa   = "IGI"

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


                               if xTipoMer = "PM"  then
                                xTipoMer = "MP"
                               ELSE
                                 if xTipoMer = "R"  then
                                  xTipoMer = "RE"
                                 ELSE
                                     if xTipoMer = "PT"  then
                                        xTipoMer = "OT"
                                     ELSE
                                        if xTipoMer = "PA"  then
                                           xTipoMer = "OT"
                                        end if
                                     end if
                                 end if
                               end if

                               strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xImporta         & "    </font></td>" & chr(13) & chr(10) 'Folio del pedimento
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactu           & "    </font></td>" & chr(13) & chr(10) 'FACTURA      Folio de la factura
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xProv            & "    </font></td>" & chr(13) & chr(10) 'CODIGO       Codigo del proveedor
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFecpag          & "    </font></td>" & chr(13) & chr(10) 'FECHA FAC    Fecha de la factura
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactmo          & "    </font></td>" & chr(13) & chr(10) 'F MONEDA     Factor moneda
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xmodelo          & "    </font></td>" & chr(13) & chr(10) 'NUM PARTE    Codigo del Numero de Parte
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xMercan          & "    </font></td>" & chr(13) & chr(10) 'DESCRIPCION  Descripcion del numero de parte
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xTipoMer         & "    </font></td>" & chr(13) & chr(10) 'TIPO         Tipo de Bien
      		                     strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFrac            & "    </font></td>" & chr(13) & chr(10) 'CLAVE F      Fracción Arancelaria
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & TasaIGI          & "    </font></td>" & chr(13) & chr(10) 'TASA         Tasa Arancelaria
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & TipoTasa         & "    </font></td>" & chr(13) & chr(10) 'TIPO TASA   Tipo de Tasa Arancelaria
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xStrUniMed       & "    </font></td>" & chr(13) & chr(10) 'UNIDAD       Unidad del numero de parte
	  	                         strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xPreUni        & "    </font></td>" & chr(13) & chr(10) 'PRECIO       Precio Unitario
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xCant            & "    </font></td>" & chr(13) & chr(10) 'CANTIDAD     Cantidad Facturada
		                           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "          & "    </font></td>" & chr(13) & chr(10) 'CONVERSION   Factor de Conversión
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xpaiOri          & "    </font></td>" & chr(13) & chr(10) 'ORIGEN       Pais de Origen
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xpaiscv          & "    </font></td>" & chr(13) & chr(10) 'VENDEDOR     Pais Vendedor
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & FormaPagoIGI     & "    </font></td>" & chr(13) & chr(10) 'F PAGO       Factor de Pago
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xIncoterm        & "    </font></td>" & chr(13) & chr(10) 'INCOTERM     Termino Internacional de Comercio
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & StrAcuerdoCom    & "    </font></td>" & chr(13) & chr(10) ' Acuerdo Comercial (Agregado el 25/11/09)
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & StrTLCAN         & "    </font></td>" & chr(13) & chr(10) ' Si es TLCAN  (Agregado el 25/11/09)
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & StrTLCUEM        & "    </font></td>" & chr(13) & chr(10) ' Si es TLCUEM (Agregado el 25/11/09)
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & StrTLCAELC       & "    </font></td>" & chr(13) & chr(10) ' Si es TLC (Agregado el 25/11/09)
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & StrJustTLCAN     & "    </font></td>" & chr(13) & chr(10) ' La justificacion si es TLCAN
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & StrJustTLCUEM    & "    </font></td>" & chr(13) & chr(10) ' La justificacion si es TLCUEM
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & StrJustTLCAELC   & "    </font></td>" & chr(13) & chr(10) ' La justificacion si es TLC


                               strHTML = strHTML&"</tr>"& chr(13) & chr(10)
                               ' ImporteIGI   = "" ' Importe IGI


                               'Response.End


                                RsDetMercancias.movenext
                            wend
                         else
                                '
                                'Fecha de la factura
                                'Unidad del numero de parte
                                'INCOTERM

                               ' si no tiene Mercancias
                               strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xImporta    & "    </font></td>" & chr(13) & chr(10) 'Folio del pedimento
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactu      & "    </font></td>" & chr(13) & chr(10) 'FACTURA      Folio de la factura
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xProv       & "    </font></td>" & chr(13) & chr(10) 'CODIGO       Codigo del proveedor
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "    </font></td>" & chr(13) & chr(10) 'FECHA FAC    Fecha de la factura
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactmo     & "    </font></td>" & chr(13) & chr(10) 'F MONEDA     Factor moneda
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "    </font></td>" & chr(13) & chr(10) 'NUM PARTE    Codigo del Numero de Parte
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xMercan     & "    </font></td>" & chr(13) & chr(10) 'DESCRIPCION  Descripcion del numero de parte
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "    </font></td>" & chr(13) & chr(10) 'TIPO         Tipo de Bien
      		                     strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFrac       & "    </font></td>" & chr(13) & chr(10) 'CLAVE F      Fracción Arancelaria
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &               "    </font></td>" & chr(13) & chr(10) 'TASA         Tasa Arancelaria
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &               "    </font></td>" & chr(13) & chr(10)  'TIPO TASA   Tipo de Tasa Arancelaria
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " &nbsp; "  & "    </font></td>" & chr(13) & chr(10) 'UNIDAD       Unidad del numero de parte
	  	                         strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xPreUni     & "    </font></td>" & chr(13) & chr(10) 'PRECIO       Precio Unitario
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xCant       & "    </font></td>" & chr(13) & chr(10) 'CANTIDAD     Cantidad Facturada
		                           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">       1                </font></td>" & chr(13) & chr(10) 'CONVERSION   Factor de Conversión
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xpaiOri     & "    </font></td>" & chr(13) & chr(10) 'ORIGEN       Pais de Origen
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xpaiscv     & "    </font></td>" & chr(13) & chr(10) 'VENDEDOR     Pais Vendedor
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " 0 "  & "         </font></td>" & chr(13) & chr(10) 'F PAGO       Forma de Pago
                               strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "  & "          </font></td>" & chr(13) & chr(10) 'INCOTERM     Termino Internacional de Comercio
                               strHTML = strHTML&"</tr>"& chr(13) & chr(10)

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
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xImporta    & "    </font></td>" & chr(13) & chr(10) 'Folio del pedimento
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactu      & "    </font></td>" & chr(13) & chr(10) 'FACTURA      Folio de la factura
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xProv       & "    </font></td>" & chr(13) & chr(10) 'CODIGO       Codigo del proveedor
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "        & "    </font></td>" & chr(13) & chr(10) 'FECHA FAC    Fecha de la factura
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & xFactmo     & "    </font></td>" & chr(13) & chr(10) 'F MONEDA     Factor moneda
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "        & "    </font></td>" & chr(13) & chr(10) 'NUM PARTE    Codigo del Numero de Parte
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "        & "    </font></td>" & chr(13) & chr(10) 'DESCRIPCION  Descripcion del numero de parte
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "        & "    </font></td>" & chr(13) & chr(10) 'TIPO         Tipo de Bien
      		          strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "        & "    </font></td>" & chr(13) & chr(10) 'CLAVE F      Fracción Arancelaria
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "        & "    </font></td>" & chr(13) & chr(10) 'TASA         Tasa Arancelaria
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "        & "    </font></td>" & chr(13) & chr(10)  'TIPO TASA   Tipo de Tasa Arancelaria
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "        & "    </font></td>" & chr(13) & chr(10) 'UNIDAD       Unidad del numero de parte
	  	              strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "       & "    </font></td>" & chr(13) & chr(10) 'PRECIO       Precio Unitario
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "        & "    </font></td>" & chr(13) & chr(10) 'CANTIDAD     Cantidad Facturada
		              strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">       1                </font></td>" & chr(13) & chr(10) 'CONVERSION   Factor de Conversión
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "  & "          </font></td>" & chr(13) & chr(10) 'ORIGEN       Pais de Origen
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "  & "          </font></td>" & chr(13) & chr(10) 'VENDEDOR     Pais Vendedor
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "  & "          </font></td>" & chr(13) & chr(10) 'F PAGO       Forma de Pago
                      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "  "  & "          </font></td>" & chr(13) & chr(10) 'INCOTERM     Termino Internacional de Comercio
                      strHTML = strHTML&"</tr>"& chr(13) & chr(10)

                  end if
                  RsDetFracciones.close
                  set RsDetFracciones = Nothing
                  '**********************************************************************************
                  '*****   Final de fracciones
                  '**********************************************************************************

             RsRep.movenext


             'response.end


           Wend

             strHTML = strHTML & "</table>"& chr(13) & chr(10)
          end if





          RsRep.close
          Set RsRep = Nothing
          'Se pinta todo el HTML formado
          response.Write(strHTML)
          if strHTML = "" then
             strHTML = "NO EXISTEN REGISTROS"
             response.Write(strHTML)
          end if

		 else
			strHTML = "NO EXISTEN REGISTROS"
			response.Write(strHTML)
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
