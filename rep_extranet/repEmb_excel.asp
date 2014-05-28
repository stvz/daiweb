<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->


<%
' ESTE ASP ES EL SEGUNDO Y ES PARA ADMINISTRADORES
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
Response.Buffer = TRUE
Response.Addheader "Content-Disposition", "attachment;filename=CONTROL_EMBARQUES.xls"
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

'response.Write("Permisos="&permi)
















'strPermisos = Request.Form("Permisos")
'permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

'response.Write("Permisos="&strPermisos)

'if not permi = "" then
'  permi = "  and (" & permi & ") "
'end if
'AplicaFiltro = false
'strFiltroCliente = ""
'strFiltroCliente = request.Form("txtCliente")
'if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
'   blnAplicaFiltro = true
'end if
'if blnAplicaFiltro then
'   permi = " AND cvecli01 =" & strFiltroCliente
'end if
'if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
'   permi = ""
'end if





'response.Write("Permisos="&permi)

'strPermisos = Request.Form("Permisos")
'response.Write("Permisos="&strPermisos)
'if not permi = "" then
'  permi = "  and (" & permi & ") "
'end if
'response.Write("Permisos="&permi)
'AplicaFiltro = false
'strFiltroCliente = ""
'strFiltroCliente = request.Form("txtCliente")
'if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
'   blnAplicaFiltro = true
'end if
'if blnAplicaFiltro then
'   permi = " AND cvecli01 =" & strFiltroCliente
'end if
'if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
'   permi = ""
'end if
'response.Write("Permisos="&permi)












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



if strTipoFiltro  = "Fechas" then    'Filtro por fechas

   if strTipoPedimento  = "1" then
      tmpTipo = "IMPORTACION"
      'strSQL = "SELECT tipopr01, valmer01,factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecent01 FROM ssdagi01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"

      strSQL = " SELECT  refcia01, " & _
               "         patent01, " & _
               "         numped01, " & _
               "         fecpag01, " & _
               "         tipcam01, " & _
               "         factmo01, " & _
               "         fraarn02 as fraccion, " & _
               "         ordfra02 as orden,    " & _
               "         vmerme02*FACTMO01 as valorComercialDolares,     " & _
               "         vmerme02*FACTMO01*TIPCAM01 as valorComercialMN, " & _
               "         tasadv02    as tasaIGI,   " & _
               "         (I_adv102) as IGIMN,      " & _
               "         dtafpp02   as DTAMN,      " & _
               "         (I_iva102) AS IVAMN,      " & _
               "         desf0101 AS FactuaNo,     " & _
               "         d_mer102 AS descripcion,  " & _
               "         cancom02 AS Cantidad,     " & _
               "         nompro01 AS Facturador,   " & _
               "         cveped01 AS CvePed,       " & _
               "         adusec01 AS Aduana,       " & _
               "         FLETES01/TIPCAM01 AS FleteUSD, " & _
               "         SEGROS01 AS Seguros,      " & _
               "         fecent01 AS FechaEntrada, " & _
               "         paiscv02 as paisVendedor, " & _
               "         paiOri02 as PaisOrigen,   " & _
               "         RCLI01,                   " & _
               "         fdsp01,                   " & _
               "        alea01                     " & _
               " from  SSDAGI01,  SSFRAC02, C01REFER " & _
               " WHERE  REFCIA02 = REFCIA01 AND  " & _
               "        REFE01   = REFCIA01 AND  " & _
               "        fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND " & _
               "        fecpag01 <='"&FormatoFechaInv(strDateFin)&"'     " & _
               Permi & " and " & _
               "   firmae01 !=''  and " & _
               "LTRIM(CVEPED01) != 'R1' " & _
               "order by refcia01"
   end if
   if strTipoPedimento  = "2" then
      tmpTipo = "EXPORTACION"
      'strSQL = "SELECT tipopr01, factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01 FROM ssdage01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
      strSQL = " SELECT  refcia01, " & _
               "         patent01, " & _
               "         numped01, " & _
               "         fecpag01, " & _
               "         tipcam01, " & _
               "         factmo01, " & _
               "         fraarn02 as fraccion, " & _
               "         ordfra02 as orden,    " & _
               "         vmerme02*FACTMO01 as valorComercialDolares,     " & _
               "         vmerme02*FACTMO01*TIPCAM01 as valorComercialMN, " & _
               "         tasadv02    as tasaIGI,   " & _
               "         (I_adv102) as IGIMN,      " & _
               "         dtafpp02   as DTAMN,      " & _
               "         (I_iva102) AS IVAMN,      " & _
               "         desf0101 AS FactuaNo,     " & _
               "         d_mer102 AS descripcion,  " & _
               "         cancom02 AS Cantidad,     " & _
               "         nompro01 AS Facturador,   " & _
               "         cveped01 AS CvePed,       " & _
               "         adusec01 AS Aduana,       " & _
               "         FLETES01/TIPCAM01 AS FleteUSD, " & _
               "         SEGROS01 AS Seguros,      " & _
               "         '' AS FechaEntrada, " & _
               "         paiscv02 as paisVendedor, " & _
               "         paiOri02 as PaisOrigen,   " & _
               "         RCLI01,                   " & _
               "         fdsp01,                   " & _
               "         alea01                    " & _
               " from  SSDAGE01,  SSFRAC02, C01REFER " & _
               " WHERE  REFCIA02 = REFCIA01 AND  " & _
               "        REFE01   = REFCIA01 AND  " & _
               "        fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND " & _
               "        fecpag01 <='"&FormatoFechaInv(strDateFin)&"'     " & _
               Permi & " and " & _
               "   firmae01 !=''  and " & _
               "LTRIM(CVEPED01) != 'R1' " & _
               "order by refcia01"

   end if

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
       strHTML = strHTML & " <p> <img width='181' eight='38'  src='http://10.66.1.5/PortalMySQL/Extranet/ext-Images/Gifs/abbot.gif'> </p> <P>&nbsp;</P>"
	     strHTML = strHTML & "<strong><font color=""#969696"" size=""4"" face=""Arial, Helvetica, sans-serif""> CONTROL DE EMBARQUES  DEL " & strDateIni & " al " & strDateFin & " </font></strong>"

       'strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>GRUPO REYES KURI, S.C. </p></font></strong>"

       strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	     strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)

		   Hd1 = "Cliente"
 	     if strTipoUsuario = MM_Cod_Cliente_Division then
	        Hd1 = "Division"
		   end if
		   if strTipoUsuario = MM_Cod_Admon or strTipoUsuario = MM_Cod_Ejecutivo_Grupo then
	        Hd1 = "Cliente"
		   end if


		   strHTML = strHTML & "<td width=""60""   bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> O.C. No                  </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""120""  bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FACTURA Nº               </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""110""  bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SHIPMENT NUMBER          </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""80""   bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CODIGO NAL               </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""80""   bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CODIGO INTL              </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""190""  bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESCRIPCION              </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""60""   bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CANTIDAD                 </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""100""  bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FACTURADOR               </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""60""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PATENTE                  </font></strong></td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">REFERENCIA </td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""90""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PEDIMENTO No             </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""130""  bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA PEDIMENTO          </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""75""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PERMISO No               </font></strong></td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""75""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FRACCION                 </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""140""  bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VALOR COMERCIAL MXP      </font></strong></td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""140""  bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VALOR COMERCIAL USD      </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""70""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> Fletes USD               </font></strong></td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""50""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IGI %                    </font></strong></td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""80""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IGI MXP                  </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""80""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DTA MXP                  </font></strong></td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""125""  bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PREVALIDACION MXP        </font></strong></td>" & chr(13) & chr(10)
		   strHTML = strHTML & "<td width=""80""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IVA MXP                  </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""95""   bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> RECARGOS MXP             </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""130""   bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA SOL. ANTICIPO      </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""170""   bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ANTICIPO PARA IMPUESTOS  </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""170""   bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DIF SOLICITADO Y PAGADO  </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""120""   bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ANTICIPOS FLETES         </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""100""  bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CTA. GASTOS              </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""140""  bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA CUENTA GASTOS      </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""145""  bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> GASTOS ADUANALES         </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""130""  bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> RECEPCION IMPORT         </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""145""  bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> RECEPCION DE FINANZAS    </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""135""  bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA PAGO CUENTA        </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""90""  bgcolor=""#0000FF""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> % PLAN                   </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""155""  bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA LLEGADA ADUANA     </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""160""  bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA ENTRADA ALMACEN    </font></strong></td>" & chr(13) & chr(10) 'Fecha Entrada Abbot
       strHTML = strHTML & "<td width=""115""  bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DIAS EN TRANSITO         </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""105""  bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PAIS VENDEDOR            </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""95""   bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PAIS DE ORIGEN           </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""145""   bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PUERTO DE DESCARGA       </font></strong></td>" & chr(13) & chr(10)
       strHTML = strHTML & "<td width=""100""   bgcolor=""#FF0000""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> COMENTARIOS              </font></strong></td>" & chr(13) & chr(10)



       'strHTML = strHTML & "<td width=""90""  bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CVE. PED.  </td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#339966""  nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ADUANA  </td>" & chr(13) & chr(10)
		   'strHTML = strHTML & "<td width=""100""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Identificadores </font></td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""60""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Partidas </font></td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""105""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Guia BL Master</td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""105""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Guia BL House</td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""120""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Fecha Despacho</td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""70""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Semaforo</td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""70""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Seguros</td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""80""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Anticipos</td>" & chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""175""  nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Saldo de la Cta de Gastos</td>" & chr(13) & chr(10)





       strHTML = strHTML & "</tr>"& chr(13) & chr(10)


	     While NOT RsRep.EOF
       'Se asigna el nombre de la referencia
           strRefer = RsRep.Fields.Item("refcia01").Value
           strOrden = RsRep.Fields.Item("orden").Value

           dblMonto=0
           strIdentificadores= " "
           strpermisos= " "
           dblRsPRV=0
           intDiasTransito=0
           strSemaforo= " "
           strGuiaMaster= " "
           strGuiaHouse= " "
           StrCuentaGastos = " "
           StrFechCG = " "
           dblMontoCG = 0
           dblGastAdu = 0
           dblAnt     = 0

           if RsRep.Fields.Item("alea01").Value = 1 then
              strSemaforo = "ROJO"
           else
              if RsRep.Fields.Item("alea01").Value = 2 then
                 strSemaforo = "VERDE"
              else
                 strSemaforo = " "
              end if

           end if

           intDiasTransito= Int(RsRep.Fields.Item("fdsp01").Value - RsRep.Fields.Item("FechaEntrada").Value )



             ' Calculamos el recargo en MXP
                 strSQL1= "SELECT  IF(ltrim(cveimp36)='7' OR ltrim(cveimp36)='13' ,import36,0) as monto " & _
                          " from  SSCONT36 " & _
                          " WHERE   refcia36='" &  strRefer & "' AND  " & _
                          " (ltrim(cveimp36)='7' OR ltrim(cveimp36)='13') " & _
                          " order by refcia36 "

                 'response.Write("strSQL1="&strSQL1)
                 Set RsRep1 = Server.CreateObject("ADODB.Recordset")
                 RsRep1.ActiveConnection = MM_EXTRANET_STRING
                 RsRep1.Source = strSQL1
                 RsRep1.CursorType = 0
                 RsRep1.CursorLocation = 2
                 RsRep1.LockType = 1
                 RsRep1.Open()

                 if not RsRep1.eof then
                   while not RsRep1.eof
                     dblMonto=dblMonto + RsRep1.Fields.Item("monto").Value
                     RsRep1.movenext
                   wend
                 end if
		             RsRep1.close
                 Set RsRep1=Nothing


                              ' Traemos los identificadores de cada partida
                 strSQL2= "select REFCIA12,ordfra12,cveide12,numper12 from ssipar12 where ordfra12 =" &  strOrden & " and refcia12='" &  strRefer & "'"
                 Set RsIdent = Server.CreateObject("ADODB.Recordset")
			           RsIdent.ActiveConnection = MM_EXTRANET_STRING
			           RsIdent.Source = strSQL2
			           RsIdent.CursorType = 0
			           RsIdent.CursorLocation = 2
			           RsIdent.LockType = 1
			           RsIdent.Open()

			           if not RsIdent.eof then
                    While not RsIdent.eof
                      strIdentificadores = strIdentificadores  & "  " & RsIdent.Fields.Item("cveide12").Value  & " &nbsp; "
                      strpermisos= strpermisos & "  " & RsIdent.Fields.Item("numper12").Value  & " &nbsp; "
                      RsIdent.movenext
                    wend
                 end if
                 RsIdent.close
                 set RsIdent = Nothing

                 if strpermisos="" then
                   strpermisos=" &nbsp;"
                 end if

                 if strIdentificadores="" then
                   strIdentificadores=" &nbsp;"
                 end if


            ' Calculamos el PRV
                 strSQL3= "SELECT  IF(ltrim(cveimp36)='15' ,import36,0) as PRV " & _
                          " from  SSCONT36 " & _
                          " WHERE   refcia36='" &  strRefer & "' AND  " & _
                          " ltrim(cveimp36)='15' " & _
                          " order by refcia36 "
                 Set RsPRV = Server.CreateObject("ADODB.Recordset")
			           RsPRV.ActiveConnection = MM_EXTRANET_STRING
			           RsPRV.Source = strSQL3
     	           RsPRV.CursorType = 0
			           RsPRV.CursorLocation = 2
			           RsPRV.LockType = 1
			           RsPRV.Open()

			           if not RsPRV.eof then
                   While not RsPRV.eof
                     dblRsPRV = dblRsPRV + RsPRV.Fields.Item("PRV").Value
                     RsPRV.movenext
                   wend
                 end if
                 RsPRV.close
                 set RsPRV = Nothing



            ' Traemos las Guias
                 strSQL4= "SELECT  IF( IDNGUI04=1,numgui04,'') AS guiaMaster, " & _
                          "        IF( IDNGUI04=2,numgui04,'') AS guiaHouse " & _
                          " from  ssguia04 " & _
                          " WHERE  refcia04='" &  strRefer & "' "

                 Set RsGuia = Server.CreateObject("ADODB.Recordset")
			           RsGuia.ActiveConnection = MM_EXTRANET_STRING
			           RsGuia.Source = strSQL4
     	           RsGuia.CursorType = 0
			           RsGuia.CursorLocation = 2
			           RsGuia.LockType = 1
			           RsGuia.Open()

			           if not RsGuia.eof then
                   While not RsGuia.eof
                     if not RsGuia.Fields.Item("guiaMaster").Value = "" then
                        strGuiaMaster = strGuiaMaster & "  "& RsGuia.Fields.Item("guiaMaster").Value & "&nbsp "
                     end if
                     if not RsGuia.Fields.Item("guiaHouse").Value ="" then
                        strGuiaHouse  = strGuiaHouse  & "  "& RsGuia.Fields.Item("guiaHouse").Value & " &nbsp"
                     end if
                     RsGuia.movenext
                   wend
                 end if
                 RsGuia.close
                 set RsGuia = Nothing


            ' Traemos la Cuenta de Gastos
                 strSQL5= "select A.cgas31," & _
                          "       B.fech31," & _
                          "       B.sald31 " & _
                          " from d31refer as A LEFT JOIN e31cgast as B ON A.cgas31=B.cgas31 " & _
                          " where B.esta31='I' and " & _
                          "       A.refe31='"&strRefer&"' "

                 Set RsCG = Server.CreateObject("ADODB.Recordset")
			           RsCG.ActiveConnection = MM_EXTRANET_STRING
			           RsCG.Source = strSQL5
     	           RsCG.CursorType = 0
			           RsCG.CursorLocation = 2
			           RsCG.LockType = 1
			           RsCG.Open()

			           if not RsCG.eof then
                   While not RsCG.eof
                        StrCuentaGastos = CStr(RsCG.Fields.Item("cgas31").Value)& "&nbsp;"
                        StrFechCG       = RsCG.Fields.Item("fech31").Value
                        dblMontoCG      = cdbl(RsCG.Fields.Item("sald31").Value)
                     RsCG.movenext
                   wend
                 end if
                 RsCG.close
                 set RsCG = Nothing


            ' Traemos Gastos Aduanales
                 strSQL6= " select  sum( (if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) ) as gastosaduanales " & _
                          " from e21paghe , d21paghe " & _
                          " where  d21paghe.refe21='"&strRefer&"'  and " & _
                          "        YEAR(d21paghe.fech21) = YEAR(e21paghe.fech21) and  " & _
                          "        e21paghe.foli21 = d21paghe.foli21 and " & _
                          "        e21paghe.tmov21 = d21paghe.tmov21  and " & _
                          "        (e21paghe.esta21 <> 'S'  ) and " & _
                          "        e21paghe.tpag21 = 2 " & _
                          " group by refe21"

                 Set RsGastAdu = Server.CreateObject("ADODB.Recordset")
			           RsGastAdu.ActiveConnection = MM_EXTRANET_STRING
			           RsGastAdu.Source = strSQL6
     	           RsGastAdu.CursorType = 0
			           RsGastAdu.CursorLocation = 2
			           RsGastAdu.LockType = 1
			           RsGastAdu.Open()

			           if not RsGastAdu.eof then
                   While not RsGastAdu.eof
                        dblGastAdu  = cdbl(RsGastAdu.Fields.Item("gastosaduanales").Value)
                     RsGastAdu.movenext
                   wend
                 end if
                 RsGastAdu.close
                 set RsGastAdu = Nothing




            ' Traemos los Anticipos
                 strSQL7= " SELECT sum(mont11) AS ANTICIPO " & _
                          " FROM d11movim " & _
                          " where refe11 ='"&strRefer&"'  and " & _
                          "        conc11='ANT'  " & _
                          " group by refe11"

                 Set RsAnti = Server.CreateObject("ADODB.Recordset")
			           RsAnti.ActiveConnection = MM_EXTRANET_STRING
			           RsAnti.Source = strSQL7
     	           RsAnti.CursorType = 0
			           RsAnti.CursorLocation = 2
			           RsAnti.LockType = 1
			           RsAnti.Open()

			           if not RsAnti.eof then
                   While not RsAnti.eof
                        dblAnt  = cdbl(RsAnti.Fields.Item("ANTICIPO").Value)
                     RsAnti.movenext
                   wend
                 end if
                 RsAnti.close
                 set RsAnti = Nothing


           strHTML = strHTML&"<tr>" & chr(13) & chr(10)
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("RCLI01").Value&"&nbsp;         </font></td>" & chr(13) & chr(10) 'O.C. Nº
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("FactuaNo").Value&"&nbsp;        </font></td>" & chr(13) & chr(10) 'Factura Nº
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'SHIPMENT NUMBER
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'CODIGO NAL
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'CODIGO INTL
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("descripcion").Value&"&nbsp;     </font></td>" & chr(13) & chr(10) 'Descripción
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("Cantidad").Value&"&nbsp;        </font></td>" & chr(13) & chr(10) 'Cantidad
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("Facturador").Value&"&nbsp;      </font></td>" & chr(13) & chr(10) 'Facturador
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("patent01").Value&"&nbsp;        </font></td>" & chr(13) & chr(10) 'Patente
           ''strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strRefer&"&nbsp;</font></td>" & chr(13) & chr(10) 'Referencia
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("numped01").Value&"              </font></td>" & chr(13) & chr(10) 'Pedimento Nº
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("fecpag01").Value&"              </font></td>" & chr(13) & chr(10) 'Fecha Pedimento
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &strpermisos&"                                      </font></td>" & chr(13) & chr(10) 'Permiso Nº
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("fraccion").Value&"&nbsp;"&"     </font></td>" & chr(13) & chr(10) 'Fracciones
			     strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("valorComercialMN").Value&"      </font></td>" & chr(13) & chr(10) 'Valor Comercial MXP
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("valorComercialDolares").Value&" </font></td>" & chr(13) & chr(10) 'Valor Comercial USD
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("FleteUSD").Value&"              </font></td>" & chr(13) & chr(10) 'Fletes
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("tasaIGI").Value&"               </font></td>" & chr(13) & chr(10) 'IGI %
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("IGIMN").Value&"                 </font></td>" & chr(13) & chr(10) 'IGI MXP
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("DTAMN").Value&"                 </font></td>" & chr(13) & chr(10) 'DTA MXP
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &dblRsPRV&"                                         </font></td>" & chr(13) & chr(10) 'Prevalidación MXP
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("IVAMN").Value&"                 </font></td>" & chr(13) & chr(10) 'IVA MXP
				   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &dblMonto&"                                         </font></td>" & chr(13) & chr(10) 'Recargos MXP
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'FECHA SOL. ANTICIPO
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'ANTICIPO PARA IMPUESTOS
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'DIF SOLICITADO Y PAGADO
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'ANTICIPO FLETES
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &StrCuentaGastos&"                                  </font></td>" & chr(13) & chr(10) 'Cta de Gastos
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &StrFechCG&"                                        </font></td>" & chr(13) & chr(10) 'Fecha Cta de Gastos
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &dblGastAdu&"                                       </font></td>" & chr(13) & chr(10) 'Gastos Aduanales


           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'RECEPCION IMPORT
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'RECEPCION FINANZAS
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'FECHA PAGO CUENTA
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) '%PLAN


           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("FechaEntrada").Value&"          </font></td>" & chr(13) & chr(10) 'Fecha llegada aduana
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("fdsp01").Value&"                </font></td>" & chr(13) & chr(10) 'Fecha Entrada Abbot
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &intDiasTransito&"                                  </font></td>" & chr(13) & chr(10) 'Dias en transito
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("paisVendedor").Value&"          </font></td>" & chr(13) & chr(10) 'Pais Vendedor
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " &RsRep.Fields.Item("PaisOrigen").Value&"            </font></td>" & chr(13) & chr(10) 'Pais Origen


           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) 'PUERTO DE DESCARGA
           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">   &nbsp;                                              </font></td>" & chr(13) & chr(10) '%COMENTARIOS


           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("CvePed").Value&"</font></td>" & chr(13) & chr(10) 'Cve. Ped
           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("Aduana").Value&"</font></td>" & chr(13) & chr(10) 'Aduana
           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strOrden&"</font></td>" & chr(13) & chr(10) 'Partidas
				   'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strIdentificadores&"</font></td>" & chr(13) & chr(10) 'Identificadores
           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strGuiaMaster&"</font></td>" & chr(13) & chr(10) 'Guia BL Master
           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strGuiaHouse&"</font></td>" & chr(13) & chr(10) 'Guia BL House
           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fdsp01").Value&"</font></td>" & chr(13) & chr(10) 'Fecha Despacho
           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strSemaforo&"</font></td>" & chr(13) & chr(10) 'Semaforo
           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("Seguros").Value&"</font></td>" & chr(13) & chr(10) 'Seguros
           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblAnt&"</font></td>" & chr(13) & chr(10) 'Anticipos
           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblMontoCG&"</font></td>" & chr(13) & chr(10) 'Saldo de la Cta de Gastos

           strHTML = strHTML&"</tr>"& chr(13) & chr(10)
           RsRep.movenext
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