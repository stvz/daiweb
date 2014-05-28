<!-- #include virtual =  "PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp"-->
<%
Response.Buffer = TRUE
response.Charset = "utf-8"
Response.Addheader "Content-Disposition", "attachment; filename=Reporte_Issues_Digitalizacion.xls"
Response.ContentType = "application/vnd.ms-excel"

fechaIni = FormatoFechaInv(Trim(Request.Form("txtDateIni")))
fechaFin = FormatoFechaInv(Trim(Request.Form("txtDateFin")))

sCon = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=usuarios; UID=marcosro; PWD=grk32455; OPTION=3"

On Error Resume Next
	generaReporte()
If err.Number <> 0 Then
	Response.Write("Ocurrio un error en la generación del reporte. Error No." & err.Number)
	Response.End()
End if


function generaReporte()

	strSQL = "SELECT i.id_issue AS Issue, i.c_aduana AS Aduana, i.c_referencia AS Referencia, u.c_rfc_emp AS Cliente, " 
	strSQL = strSQL & "c.c_username AS Usuario, i.d_asunto AS Asunto, c.t_comentario AS Comentario, c.f_emision AS Fecha, " 
	strSQL = strSQL & "DATE_FORMAT(c.h_emision, GET_FORMAT(TIME,'ISO')) AS Hora, i.m_status AS Estado " 
	strSQL = strSQL & "FROM usuarios.isu_issues_blog AS i "
	strSQL = strSQL & "LEFT JOIN usuarios.isc_comentarios_blog AS c ON c.id_issue = i.id_issue AND c.id_comentario = 1 "
	strSQL = strSQL & "LEFT JOIN usuarios.use_users_extra AS u ON u.c_user_nam = c.c_username "
	strSQL = strSQL & "WHERE c.f_emision BETWEEN '" & fechaIni & "' AND '" & fechaFin & "' AND i.m_status IN ('P','S') "
	strSQL = strSQL & "ORDER BY i.m_status DESC"
	
	Set rsIssues = Server.CreateObject("ADODB.Recordset")
	rsIssues.ActiveConnection = sCon
	rsIssues.Open strSQL

	If Not rsIssues.EOF Then
	
		strHtml = imprimeEncabezados()

		While not rsIssues.EOF 

			strHTML = strHTML & "<tr>" & chr(13) & chr(10)	

			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Issue").Value)
			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Aduana").Value)
			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Referencia").Value)
			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Cliente").Value)
			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Usuario").Value)
			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Asunto").Value)
			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Comentario").Value)
			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Fecha").Value)
			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Hora").Value)
			strHtml = strHtml & imprimeDetalle(rsIssues.Fields.Item("Estado").Value)

			strHTML = strHTML & "</tr>"& chr(13) & chr(10)
			
			rsIssues.MoveNext

		Wend

		strHTML = strHTML & "</table>"& chr(13) & chr(10)

		rsIssues.Close
		Set rsIssues = Nothing

		Response.Write(strHtml)
		Response.End()
	Else
		Response.Write("La consulta no arrojó resultados")
		Response.End()
	End if
	
end function


function imprimeDetalle(Valor)
	imprimeDetalle = "<td width=""90"" nowrap><font color=""#000000"" size=""2"" face=""Arial, Helvetica, sans-serif"" align=""center"">" & Valor & "&nbsp;</font></td>" & chr(13) & chr(10)
end function



function imprimeEncabezados()
	Dim strHtml
	strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	strHTML = strHTML & "<tr bgcolor=""#08088A"" align=""center"">"& chr(13) & chr(10)
	strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Id Issue</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""200"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cliente</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Usuario</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Asunto</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""200"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Comentario</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Alta</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Hora Alta</td>" & chr(13) & chr(10)
	strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Estado</td>" & chr(13) & chr(10)
	strHTML = strHTML & "</tr>" & chr(13) & chr(10)
	
	imprimeEncabezados = strHtml
end function

%>