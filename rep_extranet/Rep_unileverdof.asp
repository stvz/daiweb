<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% 
	Server.ScriptTimeout=150000
	 strTipoUsuario = request.Form("TipoUser")
	 strPermisos = Request.Form("Permisos")
	 permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
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
		 permi = " AND cvecli01 =" & strFiltroCliente
     end if
     if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
		 permi = ""
     end if

%>
<HTML>
	<HEAD>
		<TITLE>:: .... REPORTE DE FRACCIONES EN EL Unilever vs DOF .... ::</TITLE>
	</HEAD>
	<BODY>
<% 
	if  Session("GAduana") = "" then %>
		<table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
			<tr>
				<td><%=strMenjError%></td>
			</tr>
		</table>
<% 
	else 
		jnxadu=Session("GAduana")
		select case jnxadu
			case "VER"
				strOficina="rku"
			case "MEX"
				strOficina="dai"
			case "MAN"
				strOficina="sap"
			case "GUA"
				strOficina="rku"
			case "TAM"
				strOficina="ceg"
			case "LAR"
				strOficina="LAR"
			case "LZR"
				strOficina="lzr"
			case "TOL"
				strOficina="tol"
		end select

        cve=request.form("cve") ' NO TRAJO NADA con RFC e IMPO   "" 
        mov=request.form("mov") ' Tipo de movimiento IMPO o EXPO "i"
        fi=	trim(request.form("fi")) ' Fecha inicio reporte "10/10/2010"
        ff=	trim(request.form("ff")) ' Fecha final reporte "14/10/2010"
		fid= trim(request.form("fid")) ' Fecha inicio del DOF "07/10/2010"	
		ffd= trim(request.form("ffd")) ' Fecha final del DOF "14/10/2010"
		Vrfc= Request.Form("rfcCliente") ' RFC cliente "UME651115N48"
        Vckcve=	Request.Form("ckcve") ' la seleccion si usares RFC o CVECLIE si es 0 es por RFC y si es 1 es por CVE CLI  0
        Vclave=	Request.Form("cveCliente") ' nada usando rfc  ""
		txtcli= Request.Form("txtCliente") ' clave de cliente "Todos"	
		multiofi =	Request.Form("multi") ' Multioficina "t"
		deUsa = Request.Form("pais") 'Filtrar por procedentes de USA  ""
		Filtropais = ""
		'if deUsa = "t" Then
			'Filtropais = "AND fr.paiori02='USA' "
		'End If
		
		'response.write("clave " & cve) 
		'response.write("mov " & mov)
		'response.write("Fecha inicio " & fi)
		'response.write("Fecha Final " & ff)
		'response.write("RFC " & vrfc)
		'response.write("clave cliente " & vclave)
		'response.write("ckcve " & vckcve)
		'Response.Write("CLAVE DE CLIENTE " & txtcli)
		'Response.Write("Multioficina = " & multiofi)
		'Response.End()
		  
        if isdate(fi) and isdate(ff) then
			DiaI = cstr(datepart("d",fi))
            Mesi = cstr(datepart("m",fi))
            AnioI = cstr(datepart("yyyy",fi))
            DateI = Anioi & "/" & Mesi & "/" & Diai            
			DiaF = cstr(datepart("d",ff))
            MesF = cstr(datepart("m",ff))
            AnioF = cstr(datepart("yyyy",ff))
            DateF = AnioF & "/" & MesF & "/" & DiaF
			
			DiaId = cstr(datepart("d",fid))
            MesId = cstr(datepart("m",fid))
            AnioId = cstr(datepart("yyyy",fid))
			DateIDof = AnioId & "/" & MesId & "/" & DiaId
			
			DiaFd = cstr(datepart("d",ffd))
            MesFd = cstr(datepart("m",ffd))
            AnioFd = cstr(datepart("yyyy",ffd))
            DateFDof = AnioFd & "/" & MesFd & "/" & DiaFd
            		
			'Response.Write(DateIDof &"---"&DateFDof)
			 if request.form("tipRep") = "2" then
				 Response.Addheader "Content-Disposition", "attachment; filename=ReporteDOF"
				 Response.ContentType = "application/vnd.ms-excel"
				 
			 end if
			if multiofi = "t" and Vckcve = "1" Then
			Response.Write("<table border='0' align='center' cellpadding='0' cellspacing='7' class='titulosconsultas'>" &_
								"<tr>" &_
									"<td>No es posible elegir por clave de cliente y MultiOficina elijalo por RFC</td>" &_
								"</tr>" &_
							"</table>")
			Else
				log_act = 	"SELECT 'RKU' as Ofi, MAX(d_fechahora_act) as fecha " &_
							"FROM rku_extranet.log_actualiza " &_
							"GROUP BY ofi " &_
							"UNION ALL " &_
							"SELECT 'DAI' as Ofi, MAX(d_fechahora_act) as fecha " &_
							"FROM dai_extranet.log_actualiza " &_
							"GROUP BY ofi " &_
							"UNION ALL " &_
							"SELECT 'SAP' as Ofi, MAX(d_fechahora_act) as fecha " &_
							"FROM sap_extranet.log_actualiza " &_
							"GROUP BY ofi " &_
							"UNION ALL " &_
							"SELECT 'LZR' as Ofi, MAX(d_fechahora_act) as fecha " &_
							"FROM lzr_extranet.log_actualiza " &_
							"GROUP BY ofi " &_
							"UNION ALL " &_
							"SELECT 'CEG' as Ofi, max(d_fechahora_act) as fecha " &_
							"FROM ceg_extranet.log_actualiza " &_
							"group by ofi " &_
							"UNION ALL " &_
							"SELECT 'TOL' as Ofi, max(d_fechahora_act) as fecha " &_
							"FROM tol_extranet.log_actualiza " &_
							"group by ofi "
				
				' response.write("<BR><B>Query issued:</B> " + SQL + "<BR><BR>")
				' response.write("<BR><B>DATA:</B><BR>")
				' rs.MoveFirst
				' while not rs.EOF
				
						' for i = 0 to (rs.Fields.Count - 1)
							' response.write ( rs(i).value & " | " )
						' next 
						' response.write ( "<BR>" )
						' rs.MoveNext
				' wend
				' response.write("<BR><B>Done</B>")
				Set oConn = Server.CreateObject ("ADODB.Connection")
				Set RS = Server.CreateObject ("ADODB.RecordSet")
				oConn.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
				Set RS = oConn.Execute(log_act)
				RS.MoveFirst
				'Response.Write(Actualizaciones(RS))
				set RS = Nothing
				'set RS = Server.CreateObject("ADODB.Recordset")
				query = GeneraSQL
				'response.Write(query)
				'response.End			
				set RS = oConn.Execute(query)
				if RS.BOF = True and RS.EOF = True Then
					Response.Write("<table border='1' align='center' cellpadding='0' cellspacing='7' class='titulosconsultas'>" &_
										"<tr>" &_
											"<td>No existen datos que mostrar</td>" &_
										"</tr>" &_
									"</table>")
				Else
					RS.MoveFirst
					encabezado = ""
					encabezado = "<table align='center' Width='1000' bordercolor='#C1C1C1' border='2' align='center' cellpadding='0' cellspacing='0'>"
					encabezado = encabezado	&		"<tr>" &_
														"<td colspan='7'>Fracciones Usadas Mencionadas en el Diario de la Federacion (DOF)</td>" &_
													"</tr>" &_
													"<tr>" &_
														"<td colspan='7'>Del " & fi & " al " & ff & "</td>" &_
													"</tr>"
					encabezado = encabezado & 		"<tr bgcolor='#006699' class='boton'>" &_
														CeldaHead("Aduana / Sección") &_
														CeldaHead("Pais Origen") &_
														CeldaHead("Codigo Producto") &_
														CeldaHead("Desc. Producto") &_
														CeldaHead("Fracción") &_
														CeldaHead("Fecha Publicacion")&_
														CeldaHead("Fecha Entrada Vigor")&_
														CeldaHead("Acuerdo")&_
													"</tr>"
					cuerpo = ""
					cuerpo = generahtml(RS)
					html = encabezado & cuerpo & "</table>"
					Response.Write(html)
					RS.Close
					oConn.Close
					set oConn = Nothing
					set RS = Nothing
				End If
			End if
		End If
	End If



Function GeneraSQL
	SQL = ""
	movim = mov
	if multiofi <> "t" Then
		if movim = "a" Then
			mov = "i"
			SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
			mov = "e"
			SQL = SQL & OfiSQL(mov, strOficina) & "  "
		Else
			SQL = SQL & OfiSQL(mov, strOficina) & " "
		End If
		'SQL = OfiSQL(mov, strOficina)
	Else
		For indi = 1 To 6
			Select Case indi
				Case 1
					strOficina = "rku"
					if movim = "a" Then
						mov = "i"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
						mov = "e"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					Else
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					End If
				Case 2
					strOficina = "dai"
					if movim = "a" Then
						mov = "i"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
						mov = "e"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					Else
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					End If
				Case 3
					strOficina = "sap"
					if movim = "a" Then
						mov = "i"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
						mov = "e"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					Else
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					End If
				Case 4
					strOficina = "lzr"
					if movim = "a" Then
						mov = "i"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
						mov = "e"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					Else
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					End If
				Case 5
					strOficina = "ceg"
					if movim = "a" Then
						mov = "i"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
						mov = "e"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					Else
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
					End If
				Case 6
					strOficina = "tol"
					if movim = "a" Then
						mov = "i"
						SQL = SQL & OfiSQL(mov, strOficina) & " UNION ALL "
						mov = "e"
						SQL = SQL & OfiSQL(mov, strOficina) & " "
					Else
						SQL = SQL & OfiSQL(mov, strOficina) & " "
					End If
			End Select
		Next
	End If
	'SQL = SQL & "HAVING fracciondof IS NOT NULL "
		'esponse.Write(SQL)
		'Response.End()
	GeneraSQL = SQL
End Function

Function OfiSQL(movi, ofi)
	SQL2 = ""
	if movi = "i" then
		movto = "'IMPO' as Mov, "
		fecentpre = "fecent01"
	Else
		movto = "'EXPO' as Mov, "
		fecentpre = "fecpre01"
	End If
	SQL2 = 	"SELECT " & movto &_
			"i.nomcli01 AS 'nomcliente', " &_
			"i.adusec01 AS 'aduana', " &_
			"i.cvepod01 AS 'PaisO', " &_
			"fr.fraarn02 AS 'fraccion', " &_
			"fr.d_mer102 AS 'descripcion', " &_
			"d.cpro05 AS 'codigoprod', " &_
			"dof.fraccion as 'fracciondof', " &_
			"dof.descripcion as 'descdof', " &_
			"dof.FechPub as 'fechapub', " &_
			"dof.FechVigor as 'fechvigor', " &_
			"dof.Acuerdo as 'acuerdo' " &_
			"FROM " & Ofi & "_extranet.ssdag" & movi & "01 AS i " &_
			"LEFT JOIN " & Ofi & "_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
			"LEFT JOIN " & Ofi & "_extranet.d05artic AS d ON d.refe05 = i.refcia01 AND d.agru05 = fr.ordfra02 AND d.frac05 = fr.fraarn02 " &_
			"LEFT JOIN sistemas.dof_unilever as dof ON fr.fraarn02 = REPLACE(dof.fraccion,'.','') " &_
			"WHERE i.rfccli01 like '" & Vrfc & "' " &_
			"AND i.firmae01 <> '' AND i.firmae01 IS NOT NULL AND i.fecpag01 >= '" & DateI & "' AND i.fecpag01 <='" & Datef & "' "&_
			"AND dof.fechpub >= '" & DateIDof & "' AND dof.fechpub <='" & DatefDof & "' " &_
			
			filtropais &_
			" GROUP BY fracciondof , codigoprod, acuerdo " &_
			"HAVING fracciondof IS NOT NULL "
		
	OfiSQL = SQL2
End Function

Function generahtml(RecSet)
	codigo = ""
	RecSet.MoveFirst
	Do Until RecSet.EOF
		codigo = codigo & 	"<tr>" &_
								CeldaCuerpo(RecSet("aduana")) &_
								CeldaCuerpo(RecSet("PaisO")) &_
								CeldaCuerpo(Recset("codigoprod")) &_
								CeldaCuerpo(Recset("descripcion")) &_
								CeldaCuerpo(Recset("fraccion")) &_
								CeldaCuerpo(Recset("fechapub")) &_
								CeldaCuerpo(Recset("fechvigor")) &_
								CeldaCuerpo(Recset("acuerdo")) &_
							"</tr>"
	RecSet.MoveNext
	Loop
	generahtml = codigo
End Function

Function CeldaCuerpo(txtcelda)
	tags = ""
	tags = "<td align='center'><font size='1' face='Arial'>" & txtcelda & "</font></font></td>"
	celdacuerpo = tags
End Function

Function CeldaHead(txtcelda)
	tags = ""
	tags = "<td align='center' nowrap><strong><font color='#FFFFFF' size='2' face='Arial, Helvetica, sans-serif'>" & txtcelda & "</font></td>"
	CeldaHead = tags
End Function

Function Actualizaciones(RSact)
	html = ""
	cont = 0
	log_act =	"SELECT 'RKU' as Ofi, MAX(d_fechahora_act) as fecha " &_
				"FROM rku_extranet.log_actualiza " &_
				"GROUP BY ofi " &_
				"UNION ALL " &_
				"SELECT 'DAI' as Ofi, MAX(d_fechahora_act) as fecha " &_
				"FROM dai_extranet.log_actualiza " &_
				"GROUP BY ofi " &_
				"UNION ALL " &_
				"SELECT 'SAP' as Ofi, MAX(d_fechahora_act) as fecha " &_
				"FROM sap_extranet.log_actualiza " &_
				"GROUP BY ofi " &_
				"UNION ALL " &_
				"SELECT 'LZR' as Ofi, MAX(d_fechahora_act) as fecha " &_
				"FROM lzr_extranet.log_actualiza " &_
				"GROUP BY ofi " &_
				"UNION ALL " &_
				"SELECT 'CEG' as Ofi, max(d_fechahora_act) as fecha " &_
				"FROM ceg_extranet.log_actualiza " &_
				"group by ofi " &_
				"UNION ALL " &_
				"SELECT 'TOL' as Ofi, max(d_fechahora_act) as fecha " &_
				"FROM tol_extranet.log_actualiza " &_
				"group by ofi " &_
				"order by ofi "
	
	'Set RSact = CreateObject("ADODB.RecordSet")
	'Set RSact = ConnStr.Execute(log_act)
	RsAct.MoveFirst
	
	
	html = html &	"<table border='1' align='center' cellpadding='0' cellspacing='7' class='titulosconsultas'>" &_
						"<tr>" &_
							"<td colspan=4><center>Ultimas Actualizaciones</center></td>" &_
						"</tr>" &_
						"<tr>"
	
	 Do Until RsAct.EOF = true
		html = html & 		"<td>" & RsAct("ofi") & "</td>" &_
							"<td>" & RsAct("fecha") & "</td>"
		cont = cont + 1
		if cont = 2 then
			html = html & "</tr><tr>"
			cont = 0
		End If
		RsAct.MoveNext
	Loop
	
	html = html & 		"</tr>" &_
					"</table><br><br>"
	
	RsAct.Close()
	Set RsAct = Nothing
	
	Actualizaciones = html
End Function
%>
	</BODY>
</HTML>