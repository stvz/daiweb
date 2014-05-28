<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%


 Response.Buffer = TRUE
 Response.Addheader "Content-Disposition", "attachment;"
 Response.ContentType = "application/vnd.ms-excel"

 strHTML = ""
 strTipoUsuario = Session("GTipoUsuario")
 fechaini = trim(request.Form("txtDateIni"))
 fechafin = trim(request.Form("txtDateFin"))
 strTipoOperaciones = request.Form("rbnTipoDate")

 dim Rsio,Rsio2,Rsio3,Rsio4

strPermisos = Request.Form("Permisos")

strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")


 if not fechaini="" and not fechafin="" then


    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    strDateIni = tmpAnioIni & "/" &tmpMesIni & "/"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    strDateFin = tmpAnioFin & "/" &tmpMesFin & "/"& tmpDiaFin

    if strTipoOperaciones = 1 then
	     strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE OPERACIONES DE IMPORTACION</p></font></strong>"
    else
	     strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE OPERACIONES DE EXPORTACION</p></font></strong>"
    end if

	  strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & fechaini & " Al " & fechafin & "</p></font></strong>"
    strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	  strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
	  strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
	  strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FechaPago</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TipoOper" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CvePed</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Aduana</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cliente</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Mercancia</td>" & chr(13) & chr(10)
    strHTML = strHTML & "</tr>"& chr(13) & chr(10)
  MM_EXTRANET_STRING_TEMP = ""
  for x=1 to 4
    if x=1 then
       aduana="VER"
    end if
    if x=2 then
       aduana="MAN"
    end if
    if x=3 then
       aduana="TAM"
    end if
    if x=4 then
       aduana="MEX"
    end if
    if x=5 then
       aduana="LZR"
    end if
    'if x=6 then
    '   aduana="LAR"
    'end if

    MM_EXTRANET_STRING_TEMP = ODBC_POR_ADUANA(aduana)
    permi = PermisoClientes(aduana,strPermisos,"cvecli01")


    if not permi = "" then
       permi = "  and (" & permi & ") "
    end if

    AplicaFiltro = false
    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
       blnAplicaFiltro = true
    end if
    if blnAplicaFiltro then
       permi = " AND cvecli01 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
       permi = ""
    end if

     if not permi= "" then
      set Rsio = server.CreateObject("ADODB.Recordset")
      Rsio.ActiveConnection =    MM_EXTRANET_STRING_TEMP

      if strTipoOperaciones = 1 then

         strSQL = "select numped01 Pedimento,patent01 as patente, fecpag01 FechaPago,cveped01 as cveped, adusec01 as aduana,vaduan02 as valor,nomcli01 as cliente,d_mer102 as Mercancia from ssdagi01,ssfrac02 WHERE refcia01=refcia02 and adusec01 = adusec02 and firmae01<>'' and  fecpag01>='"&strDateIni&"' and  fecpag01<='"&strDateFin&"' and cveped01<>'R1' " &  permi & " order by fecpag01"
      else
          strSQL = "select numped01 Pedimento,patent01 as patente, fecpag01 FechaPago,cveped01 as cveped, adusec01 as aduana,vaduan02 as valor,nomcli01 as cliente,d_mer102 as Mercancia from ssdage01,ssfrac02 WHERE refcia01=refcia02 and adusec01 = adusec02 and firmae01<>'' and  fecpag01>='"&strDateIni&"' and  fecpag01<='"&strDateFin&"' and cveped01<>'R1' " &  permi & " order by fecpag01"
      end if


      Rsio.Source= strSQL
      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()

      While NOT Rsio.EOF


                 strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Pedimento").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("patente").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fechapago").Value&"</font></td>" & chr(13) & chr(10)
                 if strTipoOperaciones = 1 then
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">IMPO</font></td>" & chr(13) & chr(10)
                 else
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">EXPO</font></td>" & chr(13) & chr(10)
                 end if
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("cveped").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("aduana").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("valor").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("cliente").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Mercancia").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML & "</tr>"& chr(13) & chr(10)
                 Response.Write(strHTML)
                 strHTML = ""
          Rsio.MoveNext()
      Wend

   Rsio.Close()
   Set Rsio = Nothing
   end if
 next


   strHTML = strHTML & "</table>"& chr(13) & chr(10)
   response.Write(strHTML)

 end if


%>