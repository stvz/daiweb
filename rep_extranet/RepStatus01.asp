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
	     strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE ESTATUS</p></font></strong>"
    end if

	  strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & fechaini & " Al " & fechafin & "</p></font></strong>"
    strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	  strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
	  strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
	  strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Hora</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Guia M</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Guia H</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Piezas</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Mercancia</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Impuestos</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Pago</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Firma Electronica</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente</td>" & chr(13) & chr(10)

    strHTML = strHTML & "</tr>"& chr(13) & chr(10)
  MM_EXTRANET_STRING_TEMP = ""
  for x=4 to 4
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
    MM_EXTRANET_STRING_STATUS = ODBC_POR_ADUANA(aduana&"_STATUS")
    permi = PermisoClientes(aduana,strPermisos,"clie01")


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
      Rsio.ActiveConnection =   MM_EXTRANET_STRING_STATUS

      strSQL = "SELECT A.c_referencia as Referencia,A.f_fecha as Fecha,A.m_observ as Observacion,B.tipo01 as Tipo FROM dai_status.ETXPD as A,dai_extranet.c01refer as B where A.c_referencia = B.refe01 and n_etapa = 15  " & permi
      strSQL = "SELECT A.c_referencia as Referencia,A.f_fecha as Fecha,SUBSTRING(A.m_observ,1,2)  as hora ,SUBSTRING(A.m_observ,4,2)  as MINUTOS ,SUBSTRING(A.m_observ,7,2)  as SEGUNDOS,B.tipo01 as Tipo FROM dai_status.ETXPD as A,dai_extranet.c01refer as B where A.c_referencia = B.refe01 and n_etapa = 15  " & permi
      strSQL = "SELECT A.c_referencia as Referencia,max(A.f_fecha) as Fecha,MAX(CAST(SUBSTRING(A.m_observ,1,2)  AS UNSIGNED))  as hora ,MAX(CAST(SUBSTRING(A.m_observ,4,2)  AS UNSIGNED))  as MINUTOS ,MAX(CAST(SUBSTRING(A.m_observ,7,2)  AS UNSIGNED))  as SEGUNDOS,B.tipo01 as Tipo FROM dai_status.ETXPD as A,dai_extranet.c01refer as B where A.c_referencia = B.refe01 and n_etapa = 15  " & permi & " GROUP BY A.c_referencia"
    '  Response.Write(strSQL)
    '  Response.End
      Rsio.Source= strSQL
      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()

      While NOT Rsio.EOF

            strSQL1 = "SELECT patent01,adusec01,numped01,fecpag01,firmae01 from ssdagi01 where refcia01 = '" & Rsio.Fields.Item("Referencia").Value & "' UNION SELECT patent01,adusec01,numped01,fecpag01,firmae01 from SSDAGE01 where refcia01 = '" & Rsio.Fields.Item("Referencia").Value & "'"
            set Rsestutus = server.CreateObject("ADODB.Recordset")
            Rsestutus.ActiveConnection =   MM_EXTRANET_STRING_TEMP
            Rsestutus.Source= strSQL1
            Rsestutus.CursorType = 0
            Rsestutus.CursorLocation = 2
            Rsestutus.LockType = 1
            Rsestutus.Open()
            strPedimento = ""
            strFechaPago =""
            strFirma =""
            strAduanaSec=""
            strPatente=""
            if not Rsestutus.EOF then
               strPedimento =  Rsestutus.Fields.Item("numped01").Value
               strFechaPago = Rsestutus.Fields.Item("fecpag01").Value
               strFirma = Rsestutus.Fields.Item("firmae01").Value
               strAduanaSec= Rsestutus.Fields.Item("adusec01").Value
               strPatente= Rsestutus.Fields.Item("patent01").Value

            end if
            Rsestutus.close
            set Rsestutus = Nothing

            set Rsestutus = server.CreateObject("ADODB.Recordset")
            Rsestutus.ActiveConnection =    MM_EXTRANET_STRING_TEMP
            strSQL = "SELECT * from ssguia04 where refcia04 = '" & Rsio.Fields.Item("Referencia").Value& "'"

            Rsestutus.Source= strSQL
            Rsestutus.CursorType = 0
            Rsestutus.CursorLocation = 2
            Rsestutus.LockType = 1
            Rsestutus.Open()
            strGuiaH = ""
            strGuiaM = ""
            While NOT  Rsestutus.EOF
               if  Rsestutus.Fields.Item("idngui04").Value = "1" then
                  strGuiaM =  Rsestutus.Fields.Item("numgui04").Value
               end if
               if  Rsestutus.Fields.Item("idngui04").Value = "2" then
                  strGuiaH =  Rsestutus.Fields.Item("numgui04").Value
               end if
                Rsestutus.MoveNext()
            Wend
            Rsestutus.close
            set Rsestutus = Nothing

            set Rsestutus = server.CreateObject("ADODB.Recordset")
            Rsestutus.ActiveConnection =   MM_EXTRANET_STRING_TEMP
            strSQL = "SELECT d_mer102 as desc1,sum(cancom02) as piezas from ssfrac02 where refcia02 = '" & Rsio.Fields.Item("Referencia").Value & "' group by refcia02"

            Rsestutus.Source= strSQL
            Rsestutus.CursorType = 0
            Rsestutus.CursorLocation = 2
            Rsestutus.LockType = 1
            Rsestutus.Open()
            intPiezas = 0
            strDesc= ""
            if not Rsestutus.EOF then
                 intPiezas =  Rsestutus.Fields.Item("piezas").Value
                 strDesc= Rsestutus.Fields.Item("desc1").Value
            end if
            Rsestutus.close
            set Rsestutus = Nothing

            Set OBJdbConnection = Server.CreateObject("ADODB.Connection")
            OBJdbConnection.Open MM_EXTRANET_STRING_TEMP

            SQlImpuestos="select sum(ifnull(import36,0)) as Total from sscont36 where refcia36='"&trim(Rsio.Fields.Item("Referencia").Value)&"'"&" and fpagoi36='0'"
            set RsioImpuestos = OBJdbConnection.Execute(SQlImpuestos)
            dblimporte1 = 0
            if not RsioImpuestos.eof then
               dblimporte1 =  RsioImpuestos("Total")
            end if

                 strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Referencia").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Fecha").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("HORA").Value & ":"& Rsio.Fields.Item("MINUTOS").Value & ":"& Rsio.Fields.Item("SEGUNDOS").Value & "</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strGuiaM &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strGuiaH&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& intPiezas&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strDesc&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& dblimporte1 &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&  strPedimento &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strFechaPago &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&  strFirma &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strAduanaSec &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strPatente &"</font></td>" & chr(13) & chr(10)
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