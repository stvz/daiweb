<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
 Response.Buffer = TRUE
 strHTML = ""
 strTipoUsuario = Session("GTipoUsuario")
 fechaini = trim(request.Form("fi"))
 fechafin = trim(request.Form("ff"))
 strTipoOperaciones = request.Form("mov")
 treporte=request.Form("tipRep")
 aduana=request.Form("OficinaG")
 dim Rsio,Rsio2,Rsio3,Rsio4

 strFiltroCliente = ""
 strFiltroCliente = request.Form("rfcCliente")
 
if strFiltroCliente="0" then 
strFiltroCliente=""
end if
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
     elseif strTipoOperaciones = 2 then 
	     strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE OPERACIONES DE EXPORTACION</p></font></strong>"
	
     end if

		strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
		strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & strDateIni & " Al " & tmpDiaFin & "</p></font></strong>"
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
		if treporte = 2 Then
					Response.Addheader "Content-Disposition", "attachment;;filename=Rep_Reporte_.xls"
					Response.ContentType = "application/vnd.ms-excel"
			End If
		if aduana="0" then 
			
			dim conta 
			conta=7
			for i=0 to conta
				Select Case i
					Case 0
						aduana = 43
					Case 1
						aduana = 47
					Case 2
						aduana = 16
					Case 3
						aduana = 65
					Case 4
						aduana=51
					case 5
						aduana=81
					case 6
						aduana=80
					case 7 
						aduana=24
				End Select
				strHTML=strHTML& Rep(aduana)
			next 
		else 
			strHTML=strHTML& Rep(aduana)
		end if 
	
	 strHTML = strHTML & "</table>"& chr(13) & chr(10)
	
  ' response.Write(strHTML)
 end if ' if que inicia en fechas
 
 
 function Rep(aduana2)
 
 set miCon=Server.CreateObject("ADODB.Connection")
		ConnectionString="DRIVER={SQL Server};SERVER=10.66.1.19;UID=sa;PWD=S0l1umF0rW;DATABASE=SIR"
		strSQL = "GSI_REP_OPE_VUCEM_SIR '"&aduana2&"',"&strTipoOperaciones&",'"&strDateIni&"','"&strDateFin&"','"&strFiltroCliente&"','RFC'"

	Set miRS = Server.CreateObject("ADODB.Recordset")
		miRS.Open strSQL, ConnectionString
	
	if err.number =0 then
	strHTML2=""
		IF Not  miRS.eof= false Then
			
			strHTML2=""
		Else
			
		
			While Not  miRS.eof
				strHTML2 = strHTML2&"<tr>" & chr(13) & chr(10)
				strHTML2= strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Robot(miRS("AduSec"),miRS("sPatente"),miRS("sPedimento"),strTipoOperaciones,"PE")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("sReferencia")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("RefRectificacion")&"</font></td>" & chr(13) & chr(10)
                strHTML2 = strHTML2&"<td width=""90"" nowrap align=""center""><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("sPedimento")&"</font></td>" & chr(13) & chr(10)
                strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("sPatente")&"</font></td>" & chr(13) & chr(10)
                strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("dFechaPago")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Robot(miRS("AduSec"),miRS("sPatente"),miRS("sPedimento"),strTipoOperaciones,"FD")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Robot(miRS("AduSec"),miRS("sPatente"),miRS("sPedimento"),strTipoOperaciones,"SMF")&"</font></td>" & chr(13) & chr(10)
                if strTipoOperaciones = 1 then
					strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">Importacion</font></td>" & chr(13) & chr(10)
                else
					strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">Exportacion</font></td>" & chr(13) & chr(10)
                end if
                strHTML2= strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("sCvePed")&"</font></td>" & chr(13) & chr(10)
                strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("AduSec")&"</font></td>" & chr(13) & chr(10)
                strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("sRazonsocialCliente")&"</font></td>" & chr(13) & chr(10)
                strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("sMercanciaDesc")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("COVES")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""350"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("ED")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""350"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("Facturas")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""350"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("FacturasCOVE")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("IVA")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("IGI")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("DTA")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&miRS("ECI")&"</font></td>" & chr(13) & chr(10)
				strHTML2 = strHTML2& "</tr>"& chr(13) & chr(10)
				miRS.movenext
			Wend
			
		 end if
		 miRS.close
    set miRS=nothing
    else 
        response.write err.description
    end if 
	Rep=strHTML2
 end function
 
 
function Robot(adusec,patent,numped,topera,opcion)
 DIM MySQL,valor
 valor=""
 MySQL=""

	if opcion="PE" then 
		MySQL="select ifnull(group_concat(distinct if(bs.Detsit01 in('730','760'),'Cumplido','Pagado')separator '/'),'-') valor "&_ 
				" from  trackingbahia.bit_soia  as bs "&_
				" where bs.Numped01 = '"&numped&"' and bs.Numpat01 = '"&patent&"' and bs.Tipope01 = "&topera&_
				" and bs.Adusec01='"&adusec&"' "&_
				" and bs.Detsit01 in ('003','005','008','011','730','760') "
				 
	elseif opcion="FD"		then 
		MySQL="select ifnull(cast((max(str_to_date(rob.Fechst01,'%d%m%Y'))) as char),'-') as valor "&_
				" from trackingbahia.bit_soia as rob "&_
				" where rob.Numped01 ='"&numped&"'  and rob.Numpat01 = '"&patent&"' "&_
				" and rob.Adusec01='"&adusec&"' "&_
				" and rob.Tipope01 = "&topera&" and rob.Detsit01 in('730','760')"			
	elseif opcion="SMF" then 	  
		MySQL="select  ifnull(IF(SUM(IF(cds.desdsc01 IS NULL, -1, IF(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1, 0))) > 0, 'ROJO',  "&_ 
				"IF(SUM(IF(cds.desdsc01 IS NULL,-1, IF(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1, 0))) = 0 , 'VERDE', 'SIN CAPTURAR')),'-')as valor  "&_ 
				" FROM trackingbahia.bit_soia            AS bs   "&_ 
				" LEFT JOIN trackingbahia.cat_situaciones     AS cs  ON bs.cvesit01 = cs.cvesit01  "&_ 
				" LEFT JOIN trackingbahia.cat_det_situaciones AS cds ON cds.detsit01 = bs.detsit01  "&_ 
				" WHERE bs.Numped01='"&numped&"' and bs.Adusec01='"&adusec&"' and bs.Numpat01='"&patent&"' and bs.Tipope01="&topera
	end if				
	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = MySQL
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	if not(act2.eof) then
		valor=act2.fields("valor").value
	else
		valor="-"
	end if
	act2.close()
	Robot=valor
end function
 
%>
<HTML>
	<HEAD>
		<TITLE>::.... Reporte de operaciones .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=strHTML%>
	</BODY>
</HTML>