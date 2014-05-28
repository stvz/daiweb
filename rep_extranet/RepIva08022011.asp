<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

 Response.Buffer = TRUE
 Response.Addheader "Content-Disposition", "attachment; filename=ReporteIVA.xls" 
 Response.ContentType = "application/vnd.ms-excel"

 Server.ScriptTimeOut=200000

 strHTML = ""
 strHTML2 = ""
 strDate=trim(request.Form("txtDateIni"))
 strDate2 = trim(request.Form("txtDateFin"))
 nCount= 0
 if not strDate="" and not strDate2="" then


   tmpDiaFin = cstr(datepart("d",strDate))
   tmpMesFin = cstr(datepart("m",strDate))
   tmpAnioFin = cstr(datepart("yyyy",strDate))
   strDateFin = tmpAnioFin & "-" &tmpMesFin & "-"& tmpDiaFin

   tmpDiaFin2 = cstr(datepart("d",strDate2))
   tmpMesFin2 = cstr(datepart("m",strDate2))
   tmpAnioFin2 = cstr(datepart("yyyy",strDate2))
   strDateFin2 = tmpAnioFin2 & "-" &tmpMesFin2 & "-"& tmpDiaFin2


   dim con,Rsio,Rsio2


strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"clie01")



if not permi = "" then
  permi = "  and (" & permi & ") "
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("rfcCliente")
if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if
if blnAplicaFiltro then
   permi = strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
end if



   set Rsio = server.CreateObject("ADODB.Recordset")
   Rsio.ActiveConnection ="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=dai_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
   'Rsio.ActiveConnection = MM_EXTRANET_STRING

     
  strSQL = " select distinct " & _
			"	i.refcia01  as REFERENCIA,  " & _
			"	r.cgas31   AS CGASTOS,  " & _
			"	cp.desc21 AS CONCEPTO,  " & _
			"   cp.clav21 as CLAVE,  " & _
			"   date_format(cta.fech31,'%d/%m/%Y')  AS FCGAST,  " & _
			"	date_format(ep.fech21,'%d/%m/%Y')  AS FPAGOH,  " & _
			"	ep.piva21  AS TIVAPH,  " & _
			"	round(( dp.mont21 /  ((ep.piva21 *  .01)+1)),2) as SUBTOTPH,  " & _
			"	round((( dp.mont21 /  ((ep.piva21 *  .01)+1))  * (ep.piva21 *  .01)),2)  as  IVAPH,  " & _
			"	round(dp.mont21,2) AS TOTALPH,  " & _
			"   dp.mfle21 as FLETE,  " &_
			"	( cta.chon31 +  ((cta.piva31*.01) * cta.chon31)) AS HONORARIOS,  " & _
			"	cta.chon31 AS SUBTOTHO,  " & _
			"	 ( cta.chon31   * (cta.piva31 *  .01))  as  IVAHONOR,  " & _
			"	cta.piva31 AS TIVAHON,  " & _
			"	cta.tota31 AS TOTCGASTOS,  " & _
			"	upper(be.nomb20) as BENEFICIARIO,  " & _
			"	upper(be.rfc20) as RFC,  " & _
			"   dp.facpro21 as FACTPROV,  " & _
			"	of.facofna  AS OFICINA  " & _
			" from  " & _
			" 	dai_extranet.ssdagi01 as i " & _
			" 	inner join dai_extranet.d31refer as r on r.refe31 = i.refcia01   " & _
			" 	inner join dai_extranet.ssclie18 as of on of.cvecli18 = i.cvecli01  " & _
			" 	inner join dai_extranet.e31cgast as cta on cta.cgas31 = r.cgas31  " & _
			" 	inner join dai_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31  " & _
			" 	left join  dai_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21  " & _
			" 	left join  dai_extranet.c21paghe as cp on cp.clav21 = ep.conc21  " & _
			" 	left join  dai_extranet.c20benef as be on be.clav20   = ep.bene21  " & _
		    " where  " & _
			" 	i.rfccli01 = '"&permi&"'  " & _
			" 	and  i.firmae01 <> ''  " & _
			" 	and cta.esta31 <> 'C'   " & _
			" 	and ep.deha21 <> 'C'  " & _
			" 	and (cta.fech31>='"&strDateFin&"' and cta.fech31<='"&strDateFin2&"')   " & _
			" union all   " & _
			" select distinct " & _
			"	i.refcia01  as REFERENCIA,  " & _
			"	r.cgas31   AS CGASTOS,  " & _
			"	cp.desc21 AS CONCEPTO,  " & _
			"   cp.clav21 as CLAVE,  " & _
			"   date_format(cta.fech31,'%d/%m/%Y')  AS FCGAST,  " & _
			"	date_format(ep.fech21,'%d/%m/%Y')  AS FPAGOH,  " & _
			"	ep.piva21  AS TIVAPH,  " & _
			"	round(( dp.mont21 /  ((ep.piva21 *  .01)+1)),2) as SUBTOTPH,  " & _
			"	round((( dp.mont21 /  ((ep.piva21 *  .01)+1))  * (ep.piva21 *  .01)),2)  as  IVAPH,  " & _
			"	round(dp.mont21,2) AS TOTALPH,  " & _
			"   dp.mfle21 as FLETE,  " &_
			"	( cta.chon31 +  ((cta.piva31*.01) * cta.chon31)) AS HONORARIOS,  " & _
			"	cta.chon31 AS SUBTOTHO,  " & _
			"	 ( cta.chon31   * (cta.piva31 *  .01))  as  IVAHONOR,  " & _
			"	cta.piva31 AS TIVAHON,  " & _
			"	cta.tota31 AS TOTCGASTOS,  " & _
			"	upper(be.nomb20) as BENEFICIARIO,  " & _
			"	upper(be.rfc20) as RFC,  " & _
			"   dp.facpro21 as FACTPROV,  " & _
			"	of.facofna  AS OFICINA  " & _
			" from  " & _
			" 	lzr_extranet.ssdagi01 as i " & _
			" 	inner join lzr_extranet.d31refer as r on r.refe31 = i.refcia01   " & _
			" 	inner join lzr_extranet.ssclie18 as of on of.cvecli18 = i.cvecli01  " & _
			" 	inner join lzr_extranet.e31cgast as cta on cta.cgas31 = r.cgas31  " & _
			" 	inner join lzr_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31  " & _
			" 	left join  lzr_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21  " & _
			" 	left join  lzr_extranet.c21paghe as cp on cp.clav21 = ep.conc21  " & _
			" 	left join  lzr_extranet.c20benef as be on be.clav20   = ep.bene21  " & _
		    " where  " & _
			" 	i.rfccli01 = '"&permi&"'  " & _
			" 	and  i.firmae01 <> ''  " & _
			" 	and cta.esta31 <> 'C'   " & _
			" 	and ep.deha21 <> 'C'  " & _
			" 	and (cta.fech31>='"&strDateFin&"' and cta.fech31<='"&strDateFin2&"')   " & _
			" union all   " & _
			" select distinct " & _
			"	i.refcia01  as REFERENCIA,  " & _
			"	r.cgas31   AS CGASTOS,  " & _
			"	cp.desc21 AS CONCEPTO,  " & _
			"   cp.clav21 as CLAVE,  " & _
			"   date_format(cta.fech31,'%d/%m/%Y')  AS FCGAST,  " & _
			"	date_format(ep.fech21,'%d/%m/%Y')  AS FPAGOH,  " & _
			"	ep.piva21  AS TIVAPH,  " & _
			"	round(( dp.mont21 /  ((ep.piva21 *  .01)+1)),2) as SUBTOTPH,  " & _
			"	round((( dp.mont21 /  ((ep.piva21 *  .01)+1))  * (ep.piva21 *  .01)),2)  as  IVAPH,  " & _
			"	round(dp.mont21,2) AS TOTALPH,  " & _
			"   dp.mfle21 as FLETE,  " &_
			"	( cta.chon31 +  ((cta.piva31*.01) * cta.chon31)) AS HONORARIOS,  " & _
			"	cta.chon31 AS SUBTOTHO,  " & _
			"	 ( cta.chon31   * (cta.piva31 *  .01))  as  IVAHONOR,  " & _
			"	cta.piva31 AS TIVAHON,  " & _
			"	cta.tota31 AS TOTCGASTOS,  " & _
			"	upper(be.nomb20) as BENEFICIARIO,  " & _
			"	upper(be.rfc20) as RFC,  " & _
			"   dp.facpro21 as FACTPROV,  " & _
			"	of.facofna  AS OFICINA  " & _
			" from  " & _
			" 	rku_extranet.ssdagi01 as i " & _
			" 	inner join rku_extranet.d31refer as r on r.refe31 = i.refcia01   " & _
			" 	inner join rku_extranet.ssclie18 as of on of.cvecli18 = i.cvecli01  " & _
			" 	inner join rku_extranet.e31cgast as cta on cta.cgas31 = r.cgas31  " & _
			" 	inner join rku_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31  " & _
			" 	left join  rku_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21  " & _
			" 	left join  rku_extranet.c21paghe as cp on cp.clav21 = ep.conc21  " & _
			" 	left join  rku_extranet.c20benef as be on be.clav20   = ep.bene21  " & _
		    " where  " & _
			" 	i.rfccli01 = '"&permi&"'  " & _
			" 	and  i.firmae01 <> ''  " & _
			" 	and cta.esta31 <> 'C'   " & _
			" 	and ep.deha21 <> 'C'  " & _
			" 	and (cta.fech31>='"&strDateFin&"' and cta.fech31<='"&strDateFin2&"')   " & _
			" union all   " & _
			" select distinct " & _
			"	i.refcia01  as REFERENCIA,  " & _
			"	r.cgas31   AS CGASTOS,  " & _
			"	cp.desc21 AS CONCEPTO,  " & _
			"   cp.clav21 as CLAVE,  " & _
			"   date_format(cta.fech31,'%d/%m/%Y')  AS FCGAST,  " & _
			"	date_format(ep.fech21,'%d/%m/%Y')  AS FPAGOH,  " & _
			"	ep.piva21  AS TIVAPH,  " & _
			"	round(( dp.mont21 /  ((ep.piva21 *  .01)+1)),2) as SUBTOTPH,  " & _
			"	round((( dp.mont21 /  ((ep.piva21 *  .01)+1))  * (ep.piva21 *  .01)),2)  as  IVAPH,  " & _
			"	round(dp.mont21,2) AS TOTALPH,  " & _
			"   dp.mfle21 as FLETE,  " &_
			"	( cta.chon31 +  ((cta.piva31*.01) * cta.chon31)) AS HONORARIOS,  " & _
			"	cta.chon31 AS SUBTOTHO,  " & _
			"	 ( cta.chon31   * (cta.piva31 *  .01))  as  IVAHONOR,  " & _
			"	cta.piva31 AS TIVAHON,  " & _
			"	cta.tota31 AS TOTCGASTOS,  " & _
			"	upper(be.nomb20) as BENEFICIARIO,  " & _
			"	upper(be.rfc20) as RFC,  " & _
			"   dp.facpro21 as FACTPROV,  " & _
			"	of.facofna  AS OFICINA  " & _
			" from  " & _
			" 	sap_extranet.ssdagi01 as i " & _
			" 	inner join sap_extranet.d31refer as r on r.refe31 = i.refcia01   " & _
			" 	inner join sap_extranet.ssclie18 as of on of.cvecli18 = i.cvecli01  " & _
			" 	inner join sap_extranet.e31cgast as cta on cta.cgas31 = r.cgas31  " & _
			" 	inner join sap_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31  " & _
			" 	left join  sap_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21  " & _
			" 	left join  sap_extranet.c21paghe as cp on cp.clav21 = ep.conc21  " & _
			" 	left join  sap_extranet.c20benef as be on be.clav20   = ep.bene21  " & _
		    " where  " & _
			" 	i.rfccli01 = '"&permi&"'  " & _
			" 	and  i.firmae01 <> ''  " & _
			" 	and cta.esta31 <> 'C'   " & _
			" 	and ep.deha21 <> 'C'  " & _
			" 	and (cta.fech31>='"&strDateFin&"' and cta.fech31<='"&strDateFin2&"')   " & _
			" union all   " & _
			" select distinct " & _
			"	i.refcia01  as REFERENCIA,  " & _
			"	r.cgas31   AS CGASTOS,  " & _
			"	cp.desc21 AS CONCEPTO,  " & _
			"   cp.clav21 as CLAVE,  " & _
			"   date_format(cta.fech31,'%d/%m/%Y')  AS FCGAST,  " & _
			"	date_format(ep.fech21,'%d/%m/%Y')  AS FPAGOH,  " & _
			"	ep.piva21  AS TIVAPH,  " & _
			"	round(( dp.mont21 /  ((ep.piva21 *  .01)+1)),2) as SUBTOTPH,  " & _
			"	round((( dp.mont21 /  ((ep.piva21 *  .01)+1))  * (ep.piva21 *  .01)),2)  as  IVAPH,  " & _
			"	round(dp.mont21,2) AS TOTALPH,  " & _
			"   dp.mfle21 as FLETE,  " &_
			"	( cta.chon31 +  ((cta.piva31*.01) * cta.chon31)) AS HONORARIOS,  " & _
			"	cta.chon31 AS SUBTOTHO,  " & _
			"	 ( cta.chon31   * (cta.piva31 *  .01))  as  IVAHONOR,  " & _
			"	cta.piva31 AS TIVAHON,  " & _
			"	cta.tota31 AS TOTCGASTOS,  " & _
			"	upper(be.nomb20) as BENEFICIARIO,  " & _
			"	upper(be.rfc20) as RFC,  " & _
			"   dp.facpro21 as FACTPROV,  " & _
			"	of.facofna  AS OFICINA  " & _
			" from  " & _
			" 	tol_extranet.ssdagi01 as i " & _
			" 	inner join tol_extranet.d31refer as r on r.refe31 = i.refcia01   " & _
			" 	inner join tol_extranet.ssclie18 as of on of.cvecli18 = i.cvecli01  " & _
			" 	inner join tol_extranet.e31cgast as cta on cta.cgas31 = r.cgas31  " & _
			" 	inner join tol_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31  " & _
			" 	left join  tol_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21  " & _
			" 	left join  tol_extranet.c21paghe as cp on cp.clav21 = ep.conc21  " & _
			" 	left join  tol_extranet.c20benef as be on be.clav20   = ep.bene21  " & _
		    " where  " & _
			" 	i.rfccli01 = '"&permi&"'  " & _
			" 	and  i.firmae01 <> ''  " & _
			" 	and cta.esta31 <> 'C'   " & _
			" 	and ep.deha21 <> 'C'  " & _
			" 	and (cta.fech31>='"&strDateFin&"' and cta.fech31<='"&strDateFin2&"')   "
			
				

   'Response.Write(strSQL)
   'Response.End


   Rsio.Source= strSQL
   Rsio.CursorType = 0
   Rsio.CursorLocation = 2
   Rsio.LockType = 1
   Rsio.Open()

   strHTML2 = strHTML2 & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE IVA</p></font></strong>"
   strHTML2 = strHTML2 & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
   strHTML2 = strHTML2 & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p> Del " & strDate & " al " & strDate2 & " </p></font></strong>"
   strHTML2 = strHTML2 & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
		strHTML2 = strHTML2 & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
   
		   strHTML2 = strHTML2 & "<td width=""300"" align=""left"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PROVEEDOR</td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""left"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">R.F.C." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""120"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FECHA DOCUMENTO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FACTURA PROVEEDOR" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IMPORTE" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TASA IVA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA</td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TOTAL" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TIPO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">OBSERV" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">RETENCIONES" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""120"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Nº DE CUENTA GASTOS" & chr(13) & chr(10)
		strHTML = strHTML2 & "</tr>"& chr(13) & chr(10)

   RefAux=""
      
   While NOT Rsio.EOF
		'Variables
		nSubtot=   Rsio.Fields.Item("SUBTOTPH").Value
		nIvaPH=    Rsio.Fields.Item("IVAPH").Value
		nFactProv= cStr(Rsio.Fields.Item("FACTPROV").Value)
		fPagoH=    cStr(Rsio.Fields.Item("FPAGOH").Value)
		sRfcProv=  Rsio.Fields.Item("RFC").Value
		nReten= 0
		sBenef="--"
		sRfc="--"		
		'----------------------------------------------------------------------------------------------------------------------------------
		set RsBene2 = server.CreateObject("ADODB.Recordset")
        RsBene2.ActiveConnection = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=dai_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		
        'Oficinas
		Select Case Rsio.Fields.Item("OFICINA").Value
			 Case 1
			   sRfc = "GRK-030919-NX4"
			   sBenef = "GRUPO REYES KURI S.C."
			   strSQL = "select ar.tpmerc05, ar.descod05 from rku_extranet.d05artic as ar where ar.refe05 ='" & Rsio.Fields.Item("REFERENCIA").Value &"' "
			 Case 4
			   sRfc = "SAP-960117-TM6"
			   sBenef = "SERVICIOS ADUANALES DEL PACIFICO S.C." 
			   strSQL = "select ar.tpmerc05, ar.descod05 from sap_extranet.d05artic as ar where ar.refe05 ='" & Rsio.Fields.Item("REFERENCIA").Value &"' "
			 Case 5
			   sRfc = "DAI-920805-RH0"
			   sBenef = "DESPACHOS AEREOS INTEGRADOS, S.C."
			   strSQL = "select ar.tpmerc05, ar.descod05 from dai_extranet.d05artic as ar where ar.refe05 ='" & Rsio.Fields.Item("REFERENCIA").Value &"' "
			 Case 9
			   sRfc = "SAP-960117-TM6"
			   sBenef = "SERVICIOS ADUANALES DEL PACIFICO S.C." 
			   strSQL = "select ar.tpmerc05, ar.descod05 from lzr_extranet.d05artic as ar where ar.refe05 ='" & Rsio.Fields.Item("REFERENCIA").Value &"' "
			 Case 12
			   sRfc = "Tol"
			   sBenef = "COMERCIO EXTERIOR DEL GOLFO, S.C."
			   strSQL = "select ar.tpmerc05, ar.descod05 from tol_extranet.d05artic as ar where ar.refe05 ='" & Rsio.Fields.Item("REFERENCIA").Value &"' "
			 Case Else 
			   sRfc = "--"
			   sBenef = "--"
			   strSQL = "select ar.tpmerc05, ar.descod05 from rku_extranet.d05artic as ar where ar.refe05 ='" & Rsio.Fields.Item("REFERENCIA").Value &"' "
			End Select		

        RsBene2.Source= strSQL
        RsBene2.CursorType = 0
        RsBene2.CursorLocation = 2
        RsBene2.LockType = 1
        RsBene2.Open()
        strTipoMercancia= ""
        strDescCode =""
        While NOT RsBene2.EOF
          strTipoMercancia= RsBene2.Fields.Item("tpmerc05").Value
          RsBene2.MoveNext()
        Wend
        RsBene2.Close()
        Set RsBene2  = Nothing 
		'----------------------------------------------------------------------------------------------------------------------------------
						
		'Verificamos si es Flete para poder sacar los montos de Subtotales, IVA y Retenciones
		if(Rsio.Fields.Item("CLAVE").Value= 7 and Rsio.Fields.Item("FLETE").Value <> 0 ) then
			nReten= Round(Rsio.Fields.Item("FLETE").Value * 0.04,2)
			nSubtot= Round((Rsio.Fields.Item("TOTALPH").Value + nReten ) / (((Rsio.Fields.Item("TIVAPH").Value*.01)+1) ),2)
			nIvaPH= Round(nSubtot * (Rsio.Fields.Item("TIVAPH").Value*.01),2)
		end if  
		
		'Campos que pueden veir vacios
		if(fPagoH = "") then
			fPagoH="--"
		end if  		
		if(nFactProv = "") then
			nFactProv="--"
		end if
		if(sRfcProv = "") then 
			sRfcProv="--"
		end if
		'----------------------------------------------------------------------------------------------------------------------------------
		if(RefAux <> cStr(Rsio.Fields.Item("REFERENCIA").Value)) then 
		
			strHTML = strHTML&"<tr>" & chr(13) & chr(10)        
				strHTML = strHTML&"<td width=""300"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sBenef &"</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sRfc&"</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("FCGAST").Value&"</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("CGASTOS").Value&"</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("SUBTOTHO").Value&"</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TIVAHON").Value&"%</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("IVAHONOR").Value&"</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("HONORARIOS").Value&"</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strTipoMercancia &"</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> HONORARIOS </font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)
				strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TOTCGASTOS").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"</tr>" & chr(13) & chr(10)
			
			RefAux= cStr(Rsio.Fields.Item("REFERENCIA").Value)
		end if
		
		strHTML = strHTML&"<tr>" & chr(13) & chr(10)        
			strHTML = strHTML&"<td width=""300"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("BENEFICIARIO").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sRfcProv &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& fPagoH &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& nFactProv &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& nSubtot &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TIVAPH").Value&"%</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& nIvaPH &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TOTALPH").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strTipoMercancia &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("CONCEPTO").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& nReten &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">--</font></td>" & chr(13) & chr(10)
	
			
		strHTML = strHTML&"</tr>" & chr(13) & chr(10)
       '----------------------------------------------------------------------------------------------------------------------------------
	   Response.Write( strHTML )
       strHTML = ""
		
		
        nCount=nCount+1

  Rsio.MoveNext()
  Wend

Rsio.Close()
Set Rsio = Nothing
end if

if(nCount < 1)then
	strHTML= strHTML2 & "</tr>"& chr(13) & chr(10)
	strHTML= strHTML &  "<tr> <td colspan=""12"">"& chr(13) & chr(10)
	strHTML= strHTML &  "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>No se encontr&oacute; ning&uacute;n registro</p></font></strong>"& chr(13) & chr(10)
	strHTML= strHTML & "</tr> </td>"& chr(13) & chr(10)
	strHTML = strHTML & "</table>"& chr(13) & chr(10)
	else
	strHTML = strHTML & "</table>"& chr(13) & chr(10)
end if

response.Write(strHTML)
%>
