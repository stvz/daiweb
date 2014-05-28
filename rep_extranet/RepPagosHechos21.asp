<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

 Response.Buffer = TRUE
 Response.Addheader "Content-Disposition", "attachment;"
 Response.ContentType = "application/vnd.ms-excel"

 Server.ScriptTimeOut=200000

 strHTML = ""

 strDate=trim(request.Form("txtDateIni"))
 strDate2 = trim(request.Form("txtDateFin"))

 if not strDate="" and not strDate2="" then


   tmpDiaFin = cstr(datepart("d",strDate))
   tmpMesFin = cstr(datepart("m",strDate))
   tmpAnioFin = cstr(datepart("yyyy",strDate))
   strDateFin = tmpAnioFin & "/" &tmpMesFin & "/"& tmpDiaFin

   tmpDiaFin2 = cstr(datepart("d",strDate2))
   tmpMesFin2 = cstr(datepart("m",strDate2))
   tmpAnioFin2 = cstr(datepart("yyyy",strDate2))
   strDateFin2 = tmpAnioFin2 & "/" &tmpMesFin2 & "/"& tmpDiaFin2


   dim con,Rsio,Rsio2


strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"clie01")

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
   permi = " AND clie01 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
end if



   set Rsio = server.CreateObject("ADODB.Recordset")
   Rsio.ActiveConnection = MM_EXTRANET_STRING

   
   ' CON REFACTURACIONES
   'strSQL = "select piva21,facpro21,fefac21,e21paghe.bene21 as Beneficiario,clie01 as Cliente, d21paghe.cgas21 as Cuenta,e21paghe.fech21 as FechaPago,d21paghe.refe21 as referencia,sum(If(e21paghe.deha21='A', d21paghe.mont21,(d21paghe.mont21)*-1)) as Importe,c21paghe.desc21 as Concepto from e21paghe,c01refer inner join d21paghe,c21paghe on (e21paghe.foli21=d21paghe.foli21 and year(e21paghe.fech21)=year(d21paghe.fech21) and e21paghe.tmov21=d21paghe.tmov21 and d21paghe.refe21 = c01refer.refe01 and e21paghe.conc21=c21paghe.clav21) and (e21paghe.fech21>='"&strDateFin&"' and e21paghe.fech21<='"&strDateFin2&"') and e21paghe.esta21 <> 'S' and e21paghe.tpag21 <> 3 and e21paghe.tmov21='P'" & permi & " group by d21paghe.refe21,c21paghe.desc21  having Importe > 0 "
   
   ' SIN REFACTURACIONES
   strSQL = "select piva21,facpro21,fefac21,e21paghe.bene21 as Beneficiario,clie01 as Cliente, d21paghe.cgas21 as Cuenta,e21paghe.fech21 as FechaPago,d21paghe.refe21 as referencia,sum(If(e21paghe.deha21='A', d21paghe.mont21,(d21paghe.mont21)*-1)) as Importe,c21paghe.desc21 as Concepto from e21paghe,c01refer inner join d21paghe,c21paghe on (e21paghe.foli21=d21paghe.foli21 and year(e21paghe.fech21)=year(d21paghe.fech21) and e21paghe.tmov21=d21paghe.tmov21 and d21paghe.refe21 = c01refer.refe01 and e21paghe.conc21=c21paghe.clav21) and (e21paghe.fech21>='"&strDateFin&"' and e21paghe.fech21<='"&strDateFin2&"') and e21paghe.esta21 <> 'S' and e21paghe.tpag21 <> 3 and e21paghe.tmov21='P'" & permi & " and c01refer.clie01 = e31cgast.clie31 LEFT JOIN e31cgast ON d21paghe.cgas21 = e31cgast.cgas31 group by d21paghe.refe21,c21paghe.desc21  having Importe > 0 "
   
   


   'Response.Write(strSQL)
   'Response.End


   Rsio.Source= strSQL
   Rsio.CursorType = 0
   Rsio.CursorLocation = 2
   Rsio.LockType = 1
   Rsio.Open()

	 strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE PAGOS HECHOS</p></font></strong>"
	 strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p> Del " & strDate & " al " & strDate2 & " </p></font></strong>"
   strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	 strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
   ' strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de pago</td>" & chr(13) & chr(10)
   ' strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia" & chr(13) & chr(10)

   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Concepto</td>" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Cuenta" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Beneficiario" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">RFC" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Documento" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Documento" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe</td>" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TASA IVA" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Total" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TipoMercancia" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DescriptionCode" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""140"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripción Fracción" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento" & chr(13) & chr(10)


   strHTML = strHTML & "</tr>"& chr(13) & chr(10)


   While NOT Rsio.EOF


        strHTML = strHTML&"<tr>" & chr(13) & chr(10)
        'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("FechaPago").Value&"</font></td>" & chr(13) & chr(10)
        ' strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Referencia").Value&"</font></td>" & chr(13) & chr(10)

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Concepto").Value&"</font></td>" & chr(13) & chr(10)

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Cuenta").Value&"</font></td>" & chr(13) & chr(10)

        set RsBene2 = server.CreateObject("ADODB.Recordset")
        RsBene2.ActiveConnection = MM_EXTRANET_STRING
        strSQL = "select fech31 from e31cgast where cgas31 ='" & Rsio.Fields.Item("Cuenta").Value &"'"
		
        RsBene2.Source= strSQL
        RsBene2.CursorType = 0
        RsBene2.CursorLocation = 2
        RsBene2.LockType = 1
        RsBene2.Open()
        strFechaCuenta= ""
        if not RsBene2.eof then
            strFechaCuenta = RsBene2.Fields.Item("fech31").Value
        end if
        RsBene2.Close()
        Set RsBene2  = Nothing

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFechaCuenta&"</font></td>" & chr(13) & chr(10)

        set RsBene = server.CreateObject("ADODB.Recordset")
        RsBene.ActiveConnection = MM_EXTRANET_STRING
        strSQL = "select * from c20benef where clav20 = '"&Rsio.Fields.Item("Beneficiario").Value&"'"
        RsBene.Source= strSQL
        RsBene.CursorType = 0
        RsBene.CursorLocation = 2
        RsBene.LockType = 1
        RsBene.Open()
        strBene =""
        strBeneRFC =""
        if not RsBene.eof then
          strBene = RsBene.Fields.Item("nomb20").Value
          strBeneRFC = RsBene.Fields.Item("RFC20").Value
        end if
        'strBene =""
        'strBeneRFC =""
        RsBene.close
        set RsBene = Nothing

        set RsBene2 = server.CreateObject("ADODB.Recordset")
        RsBene2.ActiveConnection = MM_EXTRANET_STRING
        strSQL = "select patent01,numped01 from ssdagi01  where refcia01='" & Rsio.Fields.Item("referencia").Value &"' UNION select patent01,numped01 from ssdage01 where refcia01='" & Rsio.Fields.Item("referencia").Value &"' "

        RsBene2.Source= strSQL
        RsBene2.CursorType = 0
        RsBene2.CursorLocation = 2
        RsBene2.LockType = 1
        RsBene2.Open()
        strNumPed= ""
        if not RsBene2.eof then
          strNumPed= RsBene2.Fields.Item("patent01").Value & "-" & RsBene2.Fields.Item("numped01").Value
        end if
        RsBene2.close
        set RsBene2 = Nothing


        set RsBene2 = server.CreateObject("ADODB.Recordset")
        RsBene2.ActiveConnection = MM_EXTRANET_STRING
        strSQL = "select tpmerc05,descod05 from d05artic where refe05 ='" & Rsio.Fields.Item("referencia").Value &"' "

        RsBene2.Source= strSQL
        RsBene2.CursorType = 0
        RsBene2.CursorLocation = 2
        RsBene2.LockType = 1
        RsBene2.Open()
        strTipoMercancia= ""
        strDescCode =""
        While NOT RsBene2.EOF
          strTipoMercancia= strTipoMercancia & " " & RsBene2.Fields.Item("tpmerc05").Value
          strDescCode =strDescCode  & " " & RsBene2.Fields.Item("descod05").Value
          RsBene2.MoveNext()
        Wend
        RsBene2.Close()
        Set RsBene2  = Nothing


        '*************************************************
        set Rsfracc2 = server.CreateObject("ADODB.Recordset")
        Rsfracc2.ActiveConnection = MM_EXTRANET_STRING
        ' strSQL = " select tpmerc05,descod05 from d05artic where refe05 ='" & Rsio.Fields.Item("referencia").Value &"' "
        strSQLFracc = " SELECT REFCIA02, D_MER102 FROM SSFRAC02 WHERE REFCIA02='"& Rsio.Fields.Item("referencia").Value & "'"
        Rsfracc2.Source= strSQLFracc
        Rsfracc2.CursorType = 0
        Rsfracc2.CursorLocation = 2
        Rsfracc2.LockType = 1
        Rsfracc2.Open()
        strDescFracc =""
        While NOT Rsfracc2.EOF
           if strDescFracc = "" then
              strDescFracc = Rsfracc2.Fields.Item("D_MER102").Value
           else
              strDescFracc = strDescFracc & "; " & Rsfracc2.Fields.Item("D_MER102").Value
           end if
           Rsfracc2.MoveNext()
        Wend
        Rsfracc2.Close()
        Set Rsfracc2  = Nothing
        '*************************************************

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&  strBene &"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strBeneRFC  &"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("facpro21").Value&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fefac21").Value&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Importe").Value/1.15&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("piva21").Value&"</font></td>" & chr(13) & chr(10)
        tempIva = 15
        if Rsio.Fields.Item("piva21").Value = 0 then
           tempIva = 15
        else
           tempIva = Rsio.Fields.Item("piva21").Value
        end if
        intIva =  tempIva / 100
        intIvaEntero = intIva + 1
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&(Rsio.Fields.Item("Importe").Value/ intIvaEntero) * intIva  &"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Importe").Value&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strTipoMercancia &"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strDescCode&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strDescFracc&"</font></td>" & chr(13) & chr(10)
        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strNumPed&"</font></td>" & chr(13) & chr(10)

        Response.Write( strHTML )
         strHTML = ""



  Rsio.MoveNext()
  Wend

Rsio.Close()
Set Rsio = Nothing

end if
strHTML = strHTML & "</tr>"& chr(13) & chr(10)
strHTML = strHTML & "</td>"& chr(13) & chr(10)
strHTML = strHTML & "</table>"& chr(13) & chr(10)
response.Write(strHTML)
%>
