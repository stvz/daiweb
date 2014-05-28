<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<%
   ' ESTE ASP ES EL SEGUNDO Y ES PARA ADMINISTRADORES

     MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
     Response.Buffer = TRUE

     Response.Addheader "Content-Disposition", "attachment;"
     Response.ContentType = "application/vnd.ms-excel"

     Server.ScriptTimeOut=100000

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
     strDateIni=trim(request.Form("txtDateIni"))
     strDateFin=trim(request.Form("txtDateFin"))
     strTipoPedimento=trim(request.Form("rbnTipoDate"))
     strUsuario = request.Form("user")
     strTipoUsuario = request.Form("TipoUser")

     tmpTipo = ""
     strSQL = ""

     if strTipoPedimento  = "1" then
         tmpTipo = "IMPORTACION"
         strSQL = " SELECT REFCIA01, " & _
                  "        NUMPED01, " & _
                  "        FECPAR04, " & _
                  "        HORPAR04, " & _
                  "        REFCIA01, " & _
                  "        IMPPAG04, " & _
                  "        NOMBAN04, " & _
                  "        ITPPEE04, " & _
                  "        ITPPDE04, " & _
                  "        IDDDTA04, " & _
                  "        'IMP' as tipo, " & _
                  "        nomcli01, " & _
                  "        refeco04 " & _
                  " FROM SSDAGI01 , SBPEPA04       " & _
                  " WHERE NUMPED01=NUMDOC04   AND cveaad04=patent01 AND " & _
                  "       FIRMAE01  <> ''     AND  " & _
                  "       CVEPED01  <> 'R1'   AND  " & _
                  "       FECPAR04  <> ''     AND  " & _
                  "       FECPAG01  >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                  "       FECPAG01  <= '"&FormatoFechaInv(strDateFin)&"' " & _
                          Permi
     end if
     if strTipoPedimento  = "2" then
         tmpTipo = "EXPORTACION"
         strSQL = " SELECT REFCIA01, " & _
                  "        NUMPED01, " & _
                  "        FECPAR04, " & _
                  "        HORPAR04, " & _
                  "        REFCIA01, " & _
                  "        IMPPAG04, " & _
                  "        NOMBAN04, " & _
                  "        ITPPEE04, " & _
                  "        ITPPDE04, " & _
                  "        IDDDTA04, " & _
                  "        'EXP' as tipo, " & _
                  "        nomcli01, " & _
                  "        refeco04 " & _
                  "FROM SSDAGE01 , SBPEPA04       " & _
                  " WHERE NUMPED01=NUMDOC04   AND cveaad04=patent01 AND " & _
                  "      FIRMAE01  <> ''     AND  " & _
                  "      CVEPED01  <> 'R1'   AND  " & _
                  "      FECPAR04  <> ''     AND  " & _
                  "      FECPAG01  >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                  "      FECPAG01  <= '"&FormatoFechaInv(strDateFin)&"' " & _
                         Permi
     end if

     'response.Write("Query="&strSQL)

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
	         strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE PAGO ELECTRONICO DE " & tmpTipo & " </p></font></strong>"
	         strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>GRUPO REYES KURI, S.C. </p></font></strong>"
	         strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>" & strDateIni & " al " & strDateFin & "</p></font></strong>"
           strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	         strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
           strHTML = strHTML & "<td width=""40"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Tipo           </td>" &          chr(13) & chr(10)
		       strHTML = strHTML & "<td width=""60"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Fecha          </td>" &          chr(13) & chr(10)
		       strHTML = strHTML & "<td width=""60"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Hora           </font></td>" &   chr(13) & chr(10)
           strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Banco          </font></td>" &   chr(13) & chr(10)
           strHTML = strHTML & "<td width=""75"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Referencia     </font></td>" &   chr(13) & chr(10)
           strHTML = strHTML & "<td width=""75"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Pedimento      </font></td>" &   chr(13) & chr(10)
           strHTML = strHTML & "<td width=""60"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Monto          </font></td>" &   chr(13) & chr(10)
           strHTML = strHTML & "<td width=""120"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> Cliente        </font></td>" &   chr(13) & chr(10)
           strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif""> RFC            </font></td>" &   chr(13) & chr(10)
           strHTML = strHTML & "</tr>"& chr(13) & chr(10)

	         While NOT RsRep.EOF
               'Se asigna el nombre de la referencia
               strRefer = RsRep.Fields.Item("refcia01").Value
                   'REFCIA01
                   'NUMPED01
                   'FECPAR04
                   'HORPAR04
                   'REFCIA01
                   'IMPPAG04
                   'NOMBAN04
                   'ITPPEE04
                   'ITPPDE04
                   'IDDDTA04
                   strHTML = strHTML & "<tr>"& chr(13) & chr(10)
                   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep.Fields.Item("tipo").Value &"</font></td>" & chr(13) & chr(10)  'Tipo
                   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep.Fields.Item("FECPAR04").Value &"</font></td>" & chr(13) & chr(10)  'Fecha
                   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep.Fields.Item("HORPAR04").Value &"</font></td>" & chr(13) & chr(10) 'Hora
                   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep.Fields.Item("NOMBAN04").Value &"</font></td>" & chr(13) & chr(10) 'Banco
                   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep.Fields.Item("REFCIA01").Value &"</font></td>" & chr(13) & chr(10) 'Referencia
				           strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep.Fields.Item("NUMPED01").Value &"</font></td>" & chr(13) & chr(10) 'Pedimento
                   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep.Fields.Item("IMPPAG04").Value &"</font></td>" & chr(13) & chr(10) 'Monto
                   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep.Fields.Item("NOMCLI01").Value &"</font></td>" & chr(13) & chr(10) 'Cliente
                   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep.Fields.Item("refeco04").Value &"</font></td>" & chr(13) & chr(10) 'RFC
                   strHTML = strHTML & "</tr>"& chr(13) & chr(10)
               Set RsRep1=Nothing
               RsRep.movenext
           Wend

'************************

   strHTML = strHTML & "</table>"& chr(13) & chr(10)
   'Se pinta todo el HTML formado
   response.Write(strHTML)
   end if
   RsRep.close
   Set RsRep = Nothing

   if strHTML = "" then
      strHTML = "NO EXISTEN REGISTROS"
      response.Write(strHTML)
   end if
 else
   strHTML = "NO EXISTEN REGISTROS"
   response.Write(strHTML)
end if
%>

<%

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


