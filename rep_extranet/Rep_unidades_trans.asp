<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<!--#include virtual="PortalMySQL/Extranet/ext-Asp/ext_val_user.asp" -->
<% 
	feci = Request.Form("txtDateIni")
	fecf = Request.Form("txtDateFin")
	placa = Request.Form("Placa")
	cveclie = Request.Form("txtCliente")
	filtro = ""
	if cveclie = "Todos" Then
		filtro = ""
	else
		filtro = "and p.cvecli01 = '" & cveclie & "'"
	End If
	faltaff = False
	faltafi = False
	' Valores = "Fecha Inicio = " & CStr(feci) & "<br> Fecha Final = " & CStr(fecf) & "<br> Placa = " & Iden
	' Response.Write(Valores)
	
	If IsNull(feci) Or feci = "" Then
		fi = "2005-01-01"
		feci = "01-01-2005"
	Else
		If IsDate(feci) = True Then
			diai = DatePart("d",feci)
			mesi = DatePart("m",feci)
			anoi = DatePart("yyyy",feci)
			fi = CStr(anoi) & "-" & CStr(mesi) & "-" & CStr(diai)
		Else
			faltafi = True
		End If		
	End If
	
	If IsNull(fecf) = True Or fecf = "" Then
		fecf = Date()
		diaf = DatePart("d",Date())
		mesf = DatePart("m",Date())
		anof = DatePart("yyyy",Date())
		ff = CStr(anof) & "-" & CStr(mesf) & "-" & CStr(diaf)
	Else
		If IsDate(fecf) = True Then
			diaf = DatePart("d",fecf)
			mesf = DatePart("m",fecf)
			anof = DatePart("yyyy",fecf)
			ff = CStr(anof) & "-" & CStr(mesf) & "-" & CStr(diaf)
		Else
			faltaff = True
		End If
	End If
	
	If faltaff = True Or faltafi = True Then 
		htmlerror = "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""7"" class=""titulosconsultas"">" &_
						"<tr>" &_
							"<td>Una de las fechas en el reporte es incorrecta</td>" &_
						"</tr>" &_
					"</table>"
	Else
		'Response.Write("fecha inicio = " & fi & "<br>" & "fecha final = " & ff)
		Query = "Select t.Refe01, CONCAT_WS('-', t.patent01, t.pedimento) as pedi, t.frac01, t.Consec01, t.Nload, t.Placas, t.PlacaJau, t.Economico, " &_
				"t.PesoBruto, t.PesoTara, t.Pesoaudita, t.mercan01, t.Operador, t.Transpor, t.Ticket, t.fechtick, t.fechdoc, " &_
				"t.destino, t.inicio, t.sellos, t.contrato, t.sellosE " &_
				"FROM rku_cpsimples.tcepartidas as t " &_
				"LEFT JOIN rku_cpsimples.pedimentos as p ON p.refe01 = t.refe01 " &_
				"WHERE fectick >= '" & fi & "' and fectick <= '" & ff & "' and Placas like '%" & placa & "%' " & filtro
		' Response.Write(Query)
		' Response.End()
		Set Conn = Server.CreateObject("ADODB.Connection")
		
		Conn.Open = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		Set RSUni = Server.CreateObject("ADODB.Recordset")
		Set RSUni = Conn.Execute(Query)
		
		If RSUni.BOF = True or RSUni.EOF = True Then
			Response.Write("No hay datos que coincidan con la fecha o numero de placa")
			Response.End()
		End If
	End If
	
Function CeldaCuerpo(txtcelda)
	tags = ""
	tags = "<td align='center'><font size='1' face='Arial'>" & txtcelda & "</font></font></td>"
	celdacuerpo = tags
End Function

Function CeldaHead(txtcelda)
	tags = ""
	tags = "<td width='100' nowrap><strong><font color='#FFFFFF' size='2' face='Arial, Helvetica, sans-serif'>" & txtcelda & "</font></td>"
	CeldaHead = tags
End Function 
%>
<html>
	<head>
		<title>
			BUSCADOR DE UNIDADES
		</title>
	</head>
	<body>
		<% 
			if faltaff = True Or faltafi = True Then
				Response.Write(htmlerror)
			Else
		%>
		<table align="center" Width="1000" bordercolor="#C1C1C1" border="2" align="center" cellpadding="0" cellspacing="0">
			<tr>
				<td colspan=22>
					BUSCADOR DE PLACAS
				</td>
			</tr>
			<tr>
				<td colspan=22>
					Del <%=feci%> al <%=fecf%>
				</td>
			</tr>
			<tr bgcolor='#006699' class='boton'>
				<%=CeldaHead("Referencia")%>
				<%=CeldaHead("Pedimento")%>
				<%=CeldaHead("Fraccion")%>
				<%=CeldaHead("Consecutivo")%>
				<%=CeldaHead("No. Load")%>
				<%=CeldaHead("Placa Transporte")%>
				<%=CeldaHead("Placa Jaula")%>
				<%=CeldaHead("No. Economico")%>
				<%=CeldaHead("Peso Bruto")%>
				<%=CeldaHead("Peso Tara")%>
				<%=CeldaHead("Peso Audita")%>
				<%=CeldaHead("Mercancia")%>
				<%=CeldaHead("Operador")%>
				<%=CeldaHead("Transporte")%>
				<%=CeldaHead("Ticket")%>
				<%=CeldaHead("Fecha Ticket")%>
				<%=CeldaHead("Fecha Documentos")%>
				<%=CeldaHead("Destino")%>
				<%=CeldaHead("Contrato")%>
				<%=CeldaHead("Sellos")%>
			</tr>
			<% 
			RSUni.MoveFirst
			Do Until RSUni.EOF
				Response.Write("<tr>")
				Response.Write(CeldaCuerpo(RSUni("refe01")))
				Response.Write(CeldaCuerpo(RSUni("pedi")))
				Response.Write(CeldaCuerpo(RSUni("frac01")))
				Response.Write(CeldaCuerpo(RSUni("consec01")))
				Response.Write(CeldaCuerpo(RSUni("nload")))
				Response.Write(CeldaCuerpo(RSUni("placas")))
				Response.Write(CeldaCuerpo(RSUni("placajau")))
				Response.Write(CeldaCuerpo(RSUni("economico")))
				Response.Write(CeldaCuerpo(RSUni("pesobruto")))
				Response.Write(CeldaCuerpo(RSUni("pesotara")))
				Response.Write(CeldaCuerpo(RSUni("pesoaudita")))
				Response.Write(CeldaCuerpo(RSUni("mercan01")))
				Response.Write(CeldaCuerpo(RSUni("operador")))
				Response.Write(CeldaCuerpo(RSUni("transpor")))
				Response.Write(CeldaCuerpo(RSUni("ticket")))
				Response.Write(CeldaCuerpo(RSUni("fechtick")))
				Response.Write(CeldaCuerpo(RSUni("fechdoc")))
				Response.Write(CeldaCuerpo(RSUni("destino")))
				Response.Write(CeldaCuerpo(RSUni("contrato")))
				Response.Write(CeldaCuerpo(RSUni("sellos")))
				Response.Write("</tr>")
				RSUni.MoveNext()
			Loop
			%>
		</table>
		<%
			End If
		%>
	</body>
</html>