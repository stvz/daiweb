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
permi = PermisoClientes(Session("GAduana"),strPermisos,"RE.clie01")


if not permi = "" then
  permi = "  and (" & permi & ") "
end if
'esponse.write(permi&"000--->>")

blnAplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")
'response.write(strFiltroCliente)

if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if
'response.write(permi&"001--->>")
if blnAplicaFiltro then
   permi = " AND RE.clie01 =" & strFiltroCliente
end if


if Session("GTipoUsuario") = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
end if



   '---------------------------------------------------Verificamos si el cliente es SAMSUNG y agregamos una condicion al where de la consulta pricipal
   bd=""
   bd=adu_ofi(Session("GAduana"))&"_EXTRANET."
  
   if(adu_ofi(Session("GAduana"))="")Then
		bd="DAI_EXTRANET."
   End If
   
   sQueryRFC= FiltroUsuario(Session("GTipoUsuario"))
   strRFCSEM= ""
   strCLient="SEM950215S98"
   sCondic=""
   sData=""
	if (strFiltroCliente <> "Todos")then
			
		set RsBene21 = server.CreateObject("ADODB.Recordset")
			RsBene21.ActiveConnection = MM_EXTRANET_STRING
			RsBene21.Source= sQueryRFC
			RsBene21.CursorType = 0
			RsBene21.CursorLocation = 2
			RsBene21.LockType = 1
			RsBene21.Open()
			
			While NOT RsBene21.EOF	
				'response.write(RsBene21.Fields.Item("CVECLI18").Value)	
				if(strFiltroCliente="")then
					if(strCLient = RsBene21.Fields.Item("RFCCLI18").Value)Then
						strRFCSEM = RsBene21.Fields.Item("RFCCLI18").Value
						
					end if					
					
				Else
					if(CInt(strFiltroCliente) = CInt(RsBene21.Fields.Item("CVECLI18").Value))Then
						strRFCSEM = RsBene21.Fields.Item("RFCCLI18").Value
						
					end if
				End If				
				RsBene21.MoveNext()
			Wend
			RsBene21.Close()
			Set RsBene21 = Nothing
		if(strRFCSEM = "SEM950215S98")	Then				
				sData= " , ( select ifnull(EPI.nEpi,'S/EPI') from tol_status.cuentasepi01 as EPI where EPI.nCuentGast = CG.cgas31 and i.numped01= EPI.nPedimento and EPI.nPatente = i.patent01 limit 1 ) as  EPI"			
		End If
	End If
	'response.write(strRFCSEM&"---"&sCondic)	
	'---------------------------------------------------------------------------------------------
   set Rsio = server.CreateObject("ADODB.Recordset")
   Rsio.ActiveConnection = MM_EXTRANET_STRING
   
   ' CON REFACTURACIONES
   'strSQL = "select piva21,facpro21,fefac21,e21paghe.bene21 as Beneficiario,clie01 as Cliente, d21paghe.cgas21 as Cuenta,e21paghe.fech21 as FechaPago,d21paghe.refe21 as referencia,sum(If(e21paghe.deha21='A', d21paghe.mont21,(d21paghe.mont21)*-1)) as Importe,c21paghe.desc21 as Concepto from e21paghe,c01refer inner join d21paghe,c21paghe on (e21paghe.foli21=d21paghe.foli21 and year(e21paghe.fech21)=year(d21paghe.fech21) and e21paghe.tmov21=d21paghe.tmov21 and d21paghe.refe21 = c01refer.refe01 and e21paghe.conc21=c21paghe.clav21) and (e21paghe.fech21>='"&strDateFin&"' and e21paghe.fech21<='"&strDateFin2&"') and e21paghe.esta21 <> 'S' and e21paghe.tpag21 <> 3 and e21paghe.tmov21='P'" & permi & " group by d21paghe.refe21,c21paghe.desc21  having Importe > 0 "
   
   ' SIN REFACTURACIONES
   ''strSQL = "select piva21,facpro21,fefac21,e21paghe.bene21 as Beneficiario,clie01 as Cliente, d21paghe.cgas21 as Cuenta,e21paghe.fech21 as FechaPago,d21paghe.refe21 as referencia,sum(If(e21paghe.deha21='A', d21paghe.mont21,(d21paghe.mont21)*-1)) as Importe,c21paghe.desc21 as Concepto from e21paghe,c01refer inner join d21paghe,c21paghe on (e21paghe.foli21=d21paghe.foli21 and year(e21paghe.fech21)=year(d21paghe.fech21) and e21paghe.tmov21=d21paghe.tmov21 and d21paghe.refe21 = c01refer.refe01 and e21paghe.conc21=c21paghe.clav21) and (e21paghe.fech21>='"&strDateFin&"' and e21paghe.fech21<='"&strDateFin2&"') and e21paghe.esta21 <> 'S' and e21paghe.tpag21 <> 3 and e21paghe.tmov21='P'" & permi & " and c01refer.clie01 = e31cgast.clie31 LEFT JOIN e31cgast ON d21paghe.cgas21 = e31cgast.cgas31 group by d21paghe.refe21,c21paghe.desc21  having Importe > 0 "
   

   strSQL ="select distinct " &_
	" 	EP.piva21, " &_
	" 	DP.facpro21, " &_
	" 	DP.fefac21, " &_
	"	EP.bene21 AS Beneficiario, " &_
	" 	CBE.nomb20 as Bene, " &_
	" 	CG.cgas31, " &_
 	" 	DP.cgas21 as Cuenta, " &_
	" 	EP.fech21 as FechaPago, " &_
 	" 	DP.refe21 as referencia, " &_
	" 	sum(If(EP.deha21='A', DP.mont21,(DP.mont21)*-1)) as Importe, " &_
 	" 	CP.desc21 as Concepto " &_
		sData &_
	" from  " &_
	" 	"& bd &"ssdagi01 as i  " &_
	"     left join "& bd &"c01refer as RE ON i.refcia01 = RE.refe01 " &_
	" 	LEFT JOIN "& bd &"d31refer AS RF ON RF.refe31 = i.refcia01 " &_
	" 	LEFT JOIN "& bd &"e31cgast AS CG ON CG.cgas31 = RF.cgas31 " &_
	" 	inner JOIN "& bd &"d21paghe AS DP ON DP.refe21 = i.refcia01 and DP.cgas21 = RF.cgas31 " &_
	" 	LEFT JOIN "& bd &"e21paghe AS EP ON EP.foli21=DP.foli21 and year(EP.fech21)=year(DP.fech21) and EP.tmov21=DP.tmov21 and EP.esta21 <> 'S' " &_
	" 	LEFT JOIN "& bd &"c21paghe AS CP ON CP.clav21 = EP.conc21  " &_
	" 	left join "& bd &"c20benef as CBE on CBE.clav20 = EP.bene21  and CBE.aplic20 <> 'T' " &_
	" WHERE " &_
	" 	i.firmae01 <> ''  " &_
	" 	"& permi &"  " &_
	"	AND CG.esta31<> 'C' " &_
	" 	AND CG.fech31>='"&strDateFin&"' and  CG.fech31<='"&strDateFin2&"' " &_
	" group by CP.clav21, CG.cgas31,  RF.refe31 "  &_
	" ORDER BY EP.fech21, DP.refe21 asc "

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
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Pago H.</td>" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Cuenta" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Beneficiario" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">RFC" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Factura Proveedor" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><st rong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Factura" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe</td>" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TASA IVA" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Total" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TipoMercancia" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DescriptionCode" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""140"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripción Fracción" & chr(13) & chr(10)
   strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento" & chr(13) & chr(10)
  
  If (strRFCSEM = "SEM950215S98") Then
	strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta EPI" & chr(13) & chr(10)
  End If
   

   strHTML = strHTML & "</tr>"& chr(13) & chr(10)


   While NOT Rsio.EOF


        strHTML = strHTML&"<tr>" & chr(13) & chr(10)
        'strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("FechaPago").Value&"</font></td>" & chr(13) & chr(10)
        ' strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Referencia").Value&"</font></td>" & chr(13) & chr(10)

        strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Concepto").Value&"</font></td>" & chr(13) & chr(10)
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("FechaPago").Value&"</font></td>" & chr(13) & chr(10)

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
		If (strRFCSEM = "SEM950215S98") Then
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("EPI").Value&"</font></td>" & chr(13) & chr(10)
		End If
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
