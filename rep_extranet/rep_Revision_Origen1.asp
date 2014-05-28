<!-- #include virtual =  "PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp"-->
<%
' 
m_scon = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=marcosro; PWD=grk32455; OPTION=16427"
Response.Buffer = TRUE
Response.Addheader "Content-Disposition", "attachment;"
Response.ContentType = "application/vnd.ms-excel"
Server.ScriptTimeOut = 100000
'
' ------------------------------------------------------- *
' R E P O R T E   -   R E V I S I O N   E N   O R I G E N *
'               M U L T I - A D U A N A                   *
' ------------------------------------------------------- *
' Autor: Marcos Rosete                                    *
' ------------------------------------------------------- *
' 
fechaIni     = FormatoFechaInv(Trim(Request.Form("txtDateIni")))
fechaFin     = FormatoFechaInv(Trim(Request.Form("txtDateFin")))
datosCliente = Trim(Request.Form("rfcCliente"))
strPermisos  = Request.Form("Permisos")
arreglo = Split(datosCliente, "|")

rfc = Trim(arreglo(0))
nombre = Replace(trim(arreglo(1)),"(CANCELADO)","")

'
Dim oficina()
Call initOficinas()
'
quemadas()
HTML = imprimeEncabezados()


Set rsRep = Server.CreateObject("ADODB.Recordset")
rsRep.ActiveConnection = m_scon
strSQL = getSentencia(rfc, fechaIni, fechaFin)
' response.write(strSQL)
' response.end()
rsRep.Open strSQL

While Not rsRep.EOF	
	HTML = HTML & "<tr>" & chr(13) & chr(10)	
	
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("Referencia").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("Nombre").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("RFC").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("AduSec").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("Pedimento").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("Patente").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("CvePedimento").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("PedimentoQueRectifica").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("Rectificado").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("Proveedor").Value)
	HTML = HTML & imprimeDetalle("<center>" & rsRep.Fields.Item("Factura").Value & "</center>")
	HTML = HTML & imprimeDetalle("<center>" & rsRep.Fields.Item("PO").Value & "</center>")
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("FechadePago").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("FechadeDespacho").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("VencParaR1").Value)
	HTML = HTML & imprimeDetalle("<center>" & rsRep.Fields.Item("Status").Value & "</center>")
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("SEMAFORO").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("IGI").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("IVA").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("DTA").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("ECI").Value)
	HTML = HTML & imprimeDetalle(rsRep.Fields.Item("PRV").Value)
	HTML = HTML & imprimeDetalle("<center>" & rsRep.Fields.Item("Comentarios").Value & "</center>")

	HTML = HTML & "</tr>" & chr(13) & chr(10)
	
	rsRep.MoveNext
Wend
'
rsRep.Close
Set rsRep = Nothing
'
'
HTML = HTML & "</table>"& chr(13) & chr(10)
'
Response.Write(HTML)
Response.End()

'
function getSentencia(rfcs, fechaIni, fechaFin)

	SQL = ""
	
	For i = 0 to 5
		Select Case oficina(i)
			Case "rku"
				aduanaTmp = "VER"
			Case "dai"
				aduanaTmp = "MEX"
			Case "sap"
				aduanaTmp = "MAN"
			Case "ceg"
				aduanaTmp = "TAM"
			Case Else
				aduanaTmp = UCase(oficina(i))
		End Select

		permisosUsuario = Trim(Replace(PermisoClientes(aduanaTmp ,strPermisos,"cvecli01"),"cvecli01", "imp.cvecli01"))
		If permisosUsuario <> "" Then
	
			SQL = SQL & "SELECT  " & chr(13) & chr(10)
				SQL = SQL & "imp.refcia01                              AS 'Referencia', " & chr(13) & chr(10)
				SQL = SQL & "imp.nomcli01                              AS 'Nombre', " & chr(13) & chr(10)
				SQL = SQL & "imp.rfccli01                              AS 'RFC', " & chr(13) & chr(10)
				SQL = SQL & "imp.adusec01                              AS 'AduSec', " & chr(13) & chr(10)
				SQL = SQL & "imp.numped01                              AS 'Pedimento', " & chr(13) & chr(10)
				SQL = SQL & "imp.patent01                              AS 'Patente', " & chr(13) & chr(10)
				SQL = SQL & "imp.cveped01                              AS 'CvePedimento', " & chr(13) & chr(10)
				SQL = SQL & "rec.pedorg06                              AS 'PedimentoQueRectifica', " & chr(13) & chr(10)
				SQL = SQL & "if(rep.reforg06 is null,'N','S')          AS 'Rectificado', " & chr(13) & chr(10)
				SQL = SQL & "imp.nompro01                              AS 'Proveedor', " & chr(13) & chr(10)
				SQL = SQL & "GROUP_CONCAT(DISTINCT(fac.numfac39) SEPARATOR ' , ')      AS 'Factura', " & chr(13) & chr(10)
				SQL = SQL & "GROUP_CONCAT(DISTINCT(art.pedi05) SEPARATOR ' , ')        AS 'PO', " & chr(13) & chr(10)
				SQL = SQL & "imp.fecpag01                              AS 'FechadePago', " & chr(13) & chr(10)
				SQL = SQL & "ref.fdsp01                         AS 'FechadeDespacho', " & chr(13) & chr(10)
				SQL = SQL & "ADDDATE(imp.fecpag01,30)                  AS 'VencParaR1', " & chr(13) & chr(10)
				SQL = SQL & "ifnull(Sts.estado,'REV. ORG [EN REVISION]')                                AS 'Status', " & chr(13) & chr(10)
				sql = sql & "if(sum(if(cds.desdsc01 is null, -1, if(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1, 0))) > 0, 'ROJO', if(sum(if(cds.desdsc01 is null,-1, if(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1, 0))) = 0 , 'VERDE', 'SIN CAPTURAR')) AS 'SEMAFORO', " & chr(13) & chr(10)
				'SQL = SQL & "if(sum(if(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1,0))=0,'VERDE','ROJO') AS 'SEMAFORO', " & chr(13) & chr(10)
				SQL = SQL & "if(igi.import36 IS Null,0,igi.import36)   AS 'IGI', " & chr(13) & chr(10)
				SQL = SQL & "if(iva.import36 IS Null,0,iva.import36)   AS 'IVA', " & chr(13) & chr(10)
				SQL = SQL & "if(dta.import36 IS Null,0,dta.import36)   AS 'DTA', " & chr(13) & chr(10)
				SQL = SQL & "if(eci.import36 IS Null,0,eci.import36)   AS 'ECI', " & chr(13) & chr(10)
				SQL = SQL & "if(prv.import36 IS Null,0,prv.import36)   AS 'PRV', " & chr(13) & chr(10)
				SQL = SQL & "Com.comentarios                           AS 'Comentarios' " & chr(13) & chr(10)
			SQL = SQL & "FROM " & oficina(i) & "_extranet.ssdagi01 AS imp " & chr(13) & chr(10)
			    SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.c01refer AS ref  ON ref.refe01 = imp.refcia01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.ssfact39             AS fac ON fac.refcia39 = imp.refcia01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.ssrecp06             AS rec ON rec.refcia06 = imp.refcia01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.ssiped11             AS ide ON ide.refcia11 = imp.refcia01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.d05artic             AS art ON art.refe05 = imp.refcia01  " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.sscont36             AS igi ON igi.refcia36 = imp.refcia01 AND igi.cveimp36 = 6 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.sscont36             AS iva ON iva.refcia36 = imp.refcia01 AND iva.cveimp36 = 3 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.sscont36             AS dta ON dta.refcia36 = imp.refcia01 AND dta.cveimp36 = 1 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.sscont36             AS eci ON eci.refcia36 = imp.refcia01 AND eci.cveimp36 = 18 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.sscont36             AS prv ON prv.refcia36 = imp.refcia01 AND prv.cveimp36 = 15 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN " & oficina(i) & "_extranet.ssrecp06             AS rep ON rep.reforg06 = imp.refcia01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN trackingbahia.bit_soia            AS bs  ON bs.frmsaai01 = imp.refcia01  " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN trackingbahia.cat_situaciones     AS cs  ON bs.cvesit01 = cs.cvesit01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN trackingbahia.cat_det_situaciones AS cds ON cds.detsit01 = bs.detsit01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN ( SELECT c.c_referencia AS referencia, e.d_nombre AS estado FROM usuarios.cat01_issues_subjects AS c " & chr(13) & chr(10)
				SQL = SQL & "   LEFT JOIN " & oficina(i) & "_status.etaps AS e ON e.n_etapa = c.i_etapa WHERE c.i_etapa IN (26,27,28,29) " & chr(13) & chr(10)
				SQL = SQL & "   GROUP BY c.c_referencia) AS Sts ON Sts.referencia = imp.refcia01 " & chr(13) & chr(10)
				SQL = SQL & "LEFT JOIN (SELECT c.c_referencia AS referencia, d.t_comentario AS comentarios FROM usuarios.cat01_issues_subjects AS c " & chr(13) & chr(10)
				SQL = SQL & "   LEFT JOIN usuarios.det01_issues_comments AS d ON d.i_cve_issue = c.i_cve_issue " & chr(13) & chr(10)
				SQL = SQL & "   LEFT JOIN " & oficina(i) & "_status.etaps AS e ON e.n_etapa = c.i_etapa WHERE c.i_etapa IN (26,27,28,29) " & chr(13) & chr(10)
				SQL = SQL & "   GROUP BY c.c_referencia) AS Com ON Com.referencia = imp.refcia01 " & chr(13) & chr(10)
			SQL = SQL & "WHERE imp.rfccli01 IN ('" & rfcs & "') " & chr(13) & chr(10)
			If permisosUsuario <> "imp.cvecli01=0" Then ' Si tiene permisos para todo, no se escribe..
				SQL = SQL & "AND (" &  permisosUsuario & ") " & chr(13) & chr(10)
			End if
				SQL = SQL & "AND ( (ref.fdsp01 BETWEEN '" & fechaIni & "' AND '" & fechaFin & "' )  " & chr(13) & chr(10)
				SQL = SQL & " or (imp.fecpag01 BETWEEN '" & fechaIni & "' AND '" & fechaFin & "'  ) )" & chr(13) & chr(10)
				SQL = SQL & "AND imp.firmae01 <> '' AND imp.firmae01 IS NOT NULL " & chr(13) & chr(10)
				SQL = SQL & "AND ide.cveide11 = 'RO' AND ADDDATE(imp.fecpag01,30) > CURDATE() " & chr(13) & chr(10)
			SQL = SQL & "GROUP BY imp.refcia01 " & chr(13) & chr(10)
			SQL = SQL & "having Status not in ('REV. ORG [NO REQUIERE R1]','REV. ORG [R1 REALIZADA]','FUERA DE TIEMPO, SIN CONFIRMAR')  " & chr(13) & chr(10)
						SQL = SQL & "UNION ALL " & chr(13) & chr(10)

		End if
	Next

	SQL = Mid(SQL,1, Len(SQL) - 12)
	
	getSentencia = SQL
	
end function

function imprimeEncabezados()
	Dim strHTML
		strHTML = strHTML & "<p><strong><font size=""4"" face=""Arial, Helvetica, sans-serif"">Reporte Revisión en Origen, " & nombre & "</font></strong></p>" & chr(13) & chr(10) & chr(13) & chr(10)
		strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	    strHTML = strHTML & "<tr bgcolor=""#08088A"" align=""center"">"& chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""300"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Nombre</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">R.F.C.</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana/Sección</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cve.Pedimento</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento al que rectifica</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Rectificado S/N</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""300"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Proveedor</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""130"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Factura(s)</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""130"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PO</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de pago</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de despacho</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Vencimiento para presentar R1</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Status</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Semáforo</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IGI</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DTA</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ECI</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""100"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PRV</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""300"" height=""60"" ><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Incidencia/Sobrantes/Comentarios</font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "</tr>" & chr(13) & chr(10)
		
	imprimeEncabezados = strHTML
end function

function imprimeDetalle(Valor)
	imprimeDetalle = "<td width=""100"" ><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & Valor & "</font></td>" & chr(13) & chr(10)
end function

function quemadas()
	For i = 0 to 5
		SQL = ""

		Select Case oficina(i)
			Case "rku"
				aduanaTmp = "VER"
			Case "dai"
				aduanaTmp = "MEX"
			Case "sap"
				aduanaTmp = "MAN"
			Case "ceg"
				aduanaTmp = "TAM"
			Case Else
				aduanaTmp = UCase(oficina(i))
		End Select

		SQL = SQL & "update usuarios.cat01_issues_subjects as c  " & chr(13) & chr(10)
			SQL = SQL & "left join " & oficina(i) & "_extranet.ssdagi01 as i on c.c_referencia = i.refcia01 " & chr(13) & chr(10)
		SQL = SQL & "set c.i_etapa = 31  " & chr(13) & chr(10)
		SQL = SQL & "where CURDATE() > ADDDATE(i.fecpag01,30) AND c.i_etapa in(26) AND c.t_asunto = 'RO' AND c.c_aduana='" & aduanaTmp & "' " & chr(13) & chr(10)
		
		'Response.Write(SQL)
		'Response.End()
		
		Set rsRep = Server.CreateObject("ADODB.Connection")
		rsRep.Open m_scon
		rsRep.Execute(SQL)
		rsRep.Close
		Set rsRep = Nothing
		
	Next
end function


sub initOficinas()
	Redim oficina(5)
	
	oficina(0) = "rku"
	oficina(1) = "dai"
	oficina(2) = "sap"
	oficina(3) = "ceg"
	oficina(4) = "tol"
	oficina(5) = "lzr"
end sub

%>