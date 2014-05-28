<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%

'MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

			Response.Buffer = TRUE
 strHTML = ""
 strTipoUsuario = Session("GTipoUsuario")
 fechaini = trim(request.Form("txtDateIni"))
 fechafin = trim(request.Form("txtDateFin"))
 strTipoOperaciones = request.Form("rbnTipoDate")

 dim Rsio,Rsio2,Rsio3,Rsio4

strPermisos = Request.Form("Permisos")

 strFiltroCliente = ""
 strFiltroCliente = request.Form("txtCliente")

jnxadu=Session("GAduana")

		select case jnxadu
			case "VER"
				strOficina="rku"
			case "MEX"
				strOficina="dai"
			case "MAN"
				strOficina="sap"
			case "TAM"
				strOficina="ceg"
			case "LZR"
				strOficina="lzr"
			case "TOL"
				strOficina="tol"
		end select
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
		strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PagoElectronico</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia que rectifica</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FechaPago</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FechaDespacho</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Semaforo</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TipoOper" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CvePed</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana Seccion</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cliente</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Mercancia</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">COVE</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">E-document</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Factura</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Factura:Cove</td>" & chr(13) & chr(10)
		
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IGI</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DTA</td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ECI</td>" & chr(13) & chr(10)
		
		strHTML = strHTML & "</tr>"& chr(13) & chr(10)
		MM_EXTRANET_STRING_TEMP = ""
		MM_EXTRANET_STRING_TEMP = ODBC_POR_ADUANA(Session("GAduana"))
		
		permi = PermisoClientes(jnxadu,strPermisos,"i.cvecli01")
	
     if not permi = "" then
        permi = "  and (" & permi & ") "
     end if
	
		
     AplicaFiltro = false
     if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
        blnAplicaFiltro = true
     end if
     if blnAplicaFiltro then
        permi = " AND i.cvecli01 =" & strFiltroCliente
     end if
     if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
        permi = ""
     end if
	  
	 '  if not permi= "" then
    	   query=GeneraSQL
		   'response.write(query)
		Set ConnStr = Server.CreateObject ("ADODB.Connection")
		ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		Set RSops = CreateObject("ADODB.RecordSet")
		Set RSops = ConnStr.Execute(query)
		IF RSops.BOF = True And RSops.EOF = True Then
			Response.Write("No hay datos para esas condiciones")
			strHTML=""
		Else
			
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
	
			 While NOT RSops.EOF
			
                   strHTML = strHTML&"<tr>" & chr(13) & chr(10)
   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("PagoElectronico").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("refcia01").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("Recti").Value&"</font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("Pedimento").Value&"</font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("patente").Value&"</font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("fechapago").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields("FechaDespacho")&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("Semaforo").Value&"</font></td>" & chr(13) & chr(10)
                   if strTipoOperaciones = 1 then
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">Importacion</font></td>" & chr(13) & chr(10)
                   else
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">Exportacion</font></td>" & chr(13) & chr(10)
                   end if
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("cveped").Value&"</font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("aduana").Value&"</font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("cliente").Value&"</font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("Mercancia").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("Cove").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""350"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("E-document").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""350"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("Factura").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""350"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("Factura:COVE").Value&"</font></td>" & chr(13) & chr(10)
				   
				   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("iva").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("igi").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("dta").Value&"</font></td>" & chr(13) & chr(10)
				   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RSops.Fields.Item("eci").Value&"</font></td>" & chr(13) & chr(10)
				   
                   strHTML = strHTML & "</tr>"& chr(13) & chr(10)
                 'Response.Write(strHTML)
                  ' strHTML = ""
				 RSops.MoveNext()
			 Wend
			 RSops.Close()
			 Set RSops = Nothing
	'	end if
	end if

    strHTML = strHTML & "</table>"& chr(13) & chr(10)
   response.Write(strHTML)

 end if

 function GeneraSQL
	dim mov
		mov=""
		strSQL=""
		if strTipoOperaciones = 1 then
			mov="i"
		else
			mov="e"
		end if 
		  strSQL="Select ifnull((select group_concat(distinct r.refcia06) "&_
				" from "&strOficina&"_extranet.ssrecp06 as r "&_
				" where r.reforg06 =i.refcia01 and r.patent06 =i.patent01 and r.adseor06 =i.adusec01),'') as Recti, "&_
				" (select  group_concat(distinct if(bs.Detsit01 in('730','760'),'Cumplido','Pagado')separator '/')   " & _
			"from  trackingbahia.bit_soia  as bs "& _
			" where bs.Numped01 = i.numped01 and bs.Numpat01 = i.patent01 and bs.Tipope01 = i.tipopr01 and bs.frmsaai01 = i.refcia01 and bs.Detsit01 in ('003','005','008','011','730','760'))PagoElectronico , "& _
				"i.refcia01 ,i.numped01 as Pedimento ,i.patent01 as patente,i.fecpag01 as FechaPago, " & _
				"(select cast((max(str_to_date(rob.Fechst01,'%d%m%Y'))) as char)  from trackingbahia.bit_soia as rob where rob.Numped01 = i.numped01  and rob.Numpat01 = i.patent01 and rob.Tipope01 = i.tipopr01 and rob.Detsit01 in('730','760') and rob.frmsaai01 = i.refcia01) as FechaDespacho, " & _
				"i.cveped01 cveped ,i.adusec01 as aduana ,i.nomcli01 cliente," &_
				"(select  IF(SUM(IF(cds.desdsc01 IS NULL, -1, IF(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1, 0))) > 0, 'ROJO', " &_
				"IF(SUM(IF(cds.desdsc01 IS NULL,-1, IF(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1, 0))) = 0 , 'VERDE', 'SIN CAPTURAR')) " &_
				"FROM trackingbahia.bit_soia            AS bs  " &_
				"LEFT JOIN trackingbahia.cat_situaciones     AS cs  ON bs.cvesit01 = cs.cvesit01 " &_
				"LEFT JOIN trackingbahia.cat_det_situaciones AS cds ON cds.detsit01 = bs.detsit01 " &_
				"WHERE bs.frmsaai01 =i.refcia01 ) as Semaforo, " &_
				"(select group_concat(distinct sf2.d_mer102) from "&strOficina&"_extranet.ssfrac02 as sf2 where sf2.refcia02 =i.refcia01 and i.adusec01=sf2.adusec02 ) as Mercancia, "&_
				"(select group_concat(distinct s.edocum39) from "&strOficina&"_extranet.ssfact39 as s	where s.refcia39 =i.refcia01  and s.adusec39 =i.adusec01 and s.patent39 =i.patent01) as Cove, " &_
				"(select group_concat(distinct ip.cveide11 ,'-', ip.comide11 ) from "&strOficina&"_extranet.ssiped11  as ip where ip.refcia11=i.refcia01 and ip.cveide11='ED' and ip.patent11 =i.patent01 and ip.adusec11 =i.adusec01) as 'E-document', " &_
				" (select group_concat(concat_ws(' ',f.numfac39,' ')) from "&strOficina&"_extranet.ssfact39 as f where f.refcia39 =i.refcia01 and f.patent39 =i.patent01 and f.adusec39 =i.adusec01 ) as Factura, " &_
				"(select group_concat(concat(f.numfac39,':',f.edocum39)) from "&strOficina&"_extranet.ssfact39 as f where f.refcia39 =i.refcia01 and f.patent39 =i.patent01 and f.adusec39 =i.adusec01 ) as 'Factura:COVE', "&_
				
				"if(i.cveped01 ='R1', ifnull((select sum(iva.import33) from "&strOficina&"_extranet.sscont33 as iva where iva.refcia33 =i.refcia01 and iva.cveimp33=3),0),ifnull((select sum(iva.import36) from "&strOficina&"_extranet.sscont36 as iva where iva.refcia36 =i.refcia01 and iva.cveimp36=3),0) ) as 'iva',	" & _	
				"if(i.cveped01 ='R1', ifnull((select sum(igi.import33) from "&strOficina&"_extranet.sscont33 as igi where igi.refcia33 =i.refcia01 and igi.cveimp33=6),0),ifnull((select sum(igi.import36) from "&strOficina&"_extranet.sscont36 as igi where igi.refcia36 =i.refcia01 and igi.cveimp36=6),0) ) as 'igi',	" & _
				"if(i.cveped01 ='R1', ifnull((select sum(dta.import33) from "&strOficina&"_extranet.sscont33 as dta where dta.refcia33 =i.refcia01 and dta.cveimp33=1),0),ifnull((select sum(dta.import36) from "&strOficina&"_extranet.sscont36 as dta where dta.refcia36 =i.refcia01 and dta.cveimp36=1),0) ) as 'dta'	" & _
				",  if(i.cveped01 ='R1', ifnull((select sum(eci.import33) from "&strOficina&"_extranet.sscont33 as eci where eci.refcia33 =i.refcia01 and eci.cveimp33=18),0),ifnull((select sum(eci.import36) from "&strOficina&"_extranet.sscont36 as eci where eci.refcia36 =i.refcia01 and eci.cveimp36=18),0) ) as 'eci' "&_
				"from "&strOficina&"_extranet.ssdag"&mov&"01 as i where  i.fecpag01 >='"&strDateIni&"'  and  i.fecpag01<='"&strDateFin&"'  " &  permi & _
				"  and i.firmae01 <>'' and i.firmae01 is not null order by i.fecpag01 "
		' response.write(strSQL)
		' response.end()
		 GeneraSQL=strSQL
 end function

%>