<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<%Server.ScriptTimeout=15000 %>

<HTML>
	<HEAD>
		<TITLE>
			:: ....REPORTE DE SEGUIMIENTO DE OPERACIONES.... ::
		</TITLE>
	</HEAD>
	<BODY>
		<% 
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

			if  Session("GAduana") <> "" then 
				
				
				oficina_adu=GAduana
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

				cve=request.form("cve")
				mov=request.form("mov")
				fi=trim(request.form("fi"))
				ff=trim(request.form("ff"))
				Vrfc=Request.Form("rfcCliente")
				Vckcve=Request.Form("ckcve")
				
				if isdate(fi) and isdate(ff) then
					DiaI = cstr(datepart("d",fi))
					Mesi = cstr(datepart("m",fi))
					AnioI = cstr(datepart("yyyy",fi))
					DateI = Anioi&"/"&Mesi&"/"&Diai
					DiaF = cstr(datepart("d",ff))
					MesF = cstr(datepart("m",ff))
					AnioF = cstr(datepart("yyyy",ff))
					DateF = AnioF&"/"&MesF&"/"&DiaF
				End If
				'Response.Write("Desde TOEXCEL <br> oficina = " & Stroficina & " <br> movimiento = " & mov & " <br> fecha inicial = " & fi & " <br> fecha final = " & ff & " <br> RFC = " & Vrfc & "<br><br><br>")
				query = 	"SELECT " &_
							"i.refcia01 as 'Referencia', " &_
							"CONCAT_WS('-',i.adusec01,i.patent01,i.numped01) as 'Pedimento', " &_
							"i.cvecli01 as 'cveCliente', " &_
							"i.nomcli01 as 'Cliente', " &_
							"'' as 'Ejecutivo', " &_
							"fr.d_mer102 as 'Mercancia', " &_
							"fr.paiori02 as 'PaisOrigen', " &_
							"i.totbul01 as 'Bultos', " &_
							"IF(r.feta01 = '0000-00-00', '', DATE_FORMAT(r.feta01,'%d/%m/%Y')) as 'ETA', " &_
							"IF(r.fdoc01 = '0000-00-00', '', DATE_FORMAT(r.fdoc01,'%d/%m/%Y')) as 'Documentos', " &_
							"IF(r.frev01 = '0000-00-00', '', DATE_FORMAT(r.frev01,'%d/%m/%Y')) as 'Revalidacion', " &_
							"IF(r.fpre01 = '0000-00-00', '', DATE_FORMAT(r.fpre01,'%d/%m/%Y')) as 'Previo', " &_
							"IF(i.fecpag01 = '0000-00-00', '', DATE_FORMAT(i.fecpag01,'%d/%m/%Y')) as 'Pago', " &_
							"IF(r.fdsp01 = '0000-00-00', '', DATE_FORMAT(r.fdsp01,'%d/%m/%Y')) as 'Despacho', " &_
							"IF(i.fecent01 = '0000-00-00', '', DATE_FORMAT(i.fecent01,'%d/%m/%Y')) as 'Entrada', " &_
							"( TO_DAYS(r.fdsp01) - TO_DAYS(i.fecent01) ) -  " &_
							"if( ((DAYOFWEEK(i.fecent01) -1) = 6 )   ," &_
							"          (  FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(r.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *1.5) - 0.5, " &_
							"          if( (DAYOFWEEK(i.fecent01) -1) = 7  ,  " &_
							"             (  FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(r.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *1.5) - 1, " &_
							"           if(  ( (DAYOFWEEK(r.frev01) -1)+TO_DAYS(r.fdsp01) - TO_DAYS(r.frev01) )  = 6, 0.5," &_
							"             (  FLOOR((( (DAYOFWEEK(i.fecent01) -1) + (TO_DAYS(r.fdsp01) - TO_DAYS(i.fecent01)) )/ 7)) *1.5) )))  as 'KPICTE', " &_
							"( TO_DAYS(r.fdsp01) - TO_DAYS(r.frev01) ) -  " &_
							"if( ((DAYOFWEEK(r.frev01) -1) = 6 )   ," &_
							"          (  FLOOR((( (DAYOFWEEK(r.frev01) -1) + (TO_DAYS(r.fdsp01) - TO_DAYS(r.frev01)) )/ 7)) *1.5) - 0.5, " &_
							"          if( (DAYOFWEEK(r.frev01) -1) = 7  ,  " &_
							"             (  FLOOR((( (DAYOFWEEK(r.frev01) -1) + (TO_DAYS(r.fdsp01) - TO_DAYS(r.frev01)) )/ 7)) *1.5) - 1, " &_
							"           if(  ( (DAYOFWEEK(r.frev01) -1)+TO_DAYS(r.fdsp01) - TO_DAYS(r.frev01) )  = 6, 0.5," &_
							"             (  FLOOR((( (DAYOFWEEK(r.frev01) -1) + (TO_DAYS(r.fdsp01) - TO_DAYS(r.frev01)) )/ 7)) *1.5) )))  as 'KPIGRK', " &_
							"r.obser01 as 'Observaciones Trafico', " &_
							"cta.fech31 as 'FechaCG', " &_
							"cta.cgas31 as 'CG', " &_
							"( TO_DAYS(cta.fech31) - TO_DAYS(r.fdsp01) ) -  " &_
							"if( ((DAYOFWEEK(r.fdsp01) -1) = 6 )   ," &_
							"          (  FLOOR((( (DAYOFWEEK(r.fdsp01) -1) + (TO_DAYS(r.fdsp01) - TO_DAYS(r.fdsp01)) )/ 7)) *1.5) - 0.5, " &_
							"          if( (DAYOFWEEK(r.frev01) -1) = 7  ,  " &_
							"             (  FLOOR((( (DAYOFWEEK(r.fdsp01) -1) + (TO_DAYS(r.fdsp01) - TO_DAYS(r.fdsp01)) )/ 7)) *1.5) - 1, " &_
							"           if(  ( (DAYOFWEEK(r.fdsp01) -1)+TO_DAYS(r.fdsp01) - TO_DAYS(r.fdsp01) )  = 6, 0.5," &_
							"             (  FLOOR((( (DAYOFWEEK(r.fdsp01) -1) + (TO_DAYS(cta.fech31) - TO_DAYS(r.fdsp01)) )/ 7)) *1.5) )))  as 'KPIADMIN', " &_
							"'' as 'Observaciones Administrador' " &_
							"from " & stroficina & "_extranet.ssdagi01 as i " &_
							"left join " & stroficina & "_extranet.c01refer as r on r.refe01 = i.refcia01 " &_
							"left join " & stroficina & "_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " &_
							"left join " & stroficina & "_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " &_
							"left join " & stroficina & "_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01 " &_
							"left join " & stroficina & "_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02 " &_
							"where i.rfccli01 like '" & Vrfc & "' and i.firmae01 <> ''  " &_
							"and  i.fecpag01 >=  '" & DateI &"' and i.fecpag01 <= '" & DateF & "'  and i.cveped01 <> 'R1' " &_
							"group by i.refcia01"
				'Response.Write(query)
				Set ConnStr = Server.CreateObject("ADODB.Connection")
				Set RSkpis = Server.CreateObject("ADODB.Recordset")
				'ConnStr.Open("DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427")
				ConnStr.Open("DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427")
				Set RSkpis = ConnStr.Execute(query)
				IF RSkpis.BOF = True AND RSkpis.EOF = True Then
					response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: NO HAY REGISTROS QUE CUMPLAN CON ESAS CONDICIONES ::</div></p></div>")
				Else
					%>
					<script language=VBScript>
							Set objExcel = CreateObject("Excel.Application")
							objExcel.visible = True
							Set objWorkBook = objExcel.Workbooks.open("http://rkzego.no-ip.org/kpis.xls")
							Set objHojaDatos = objExcel.ActiveWorkBook.Worksheets.item(1)
							set celda = objHojaDatos.cells(3,1)
							<% 
							dim fila
							fila = 7
							dim col
							col = 0
							cont = 0
							if mov = "i" Then
								%>
								celda.value = ":: IMPORTACION ::"
								<% 
							Else
								%>
								celda.value = ":: EXPORTACION ::"
								<% 
							End If
							%>
							set celda = objHojaDatos.cells(5,1)
							celda.value = "Del " & "<%=fi%>" & " Al " & "<%=ff%>"
							<% 
							Set RSconte = CreateObject("ADODB.Recordset")
							
							Do Until RSkpis.EOF
								fila = fila + 1
								cont = cont + 1
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=cont%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=cont%>"
								celda.value = "<%=RSkpis.Fields.Item("referencia").value%>"
								<% 
								refe = RSkpis.Fields.Item("referencia").value
								query = "SELECT numcon40 FROM " & strOficina & "_extranet.sscont40 WHERE refcia40 = '" & refe & "'"%>
								'document.write("<%=query%>" & "<br>")
								<% 
								set RSconte = ConnStr.Execute(query)
								if RSconte.BOF = true and RSconte.EOF then
									contenedores = ""
								else
									RSconte.MoveFirst
									Do Until RSconte.EOF
										contenedores = contenedores & RSconte.Fields.Item("numcon40").value & ", "
										RSconte.MoveNext()
									Loop
									contenedores = MID(contenedores,1,LEN(contenedores)-2)
								end if
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("pedimento").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("cvecliente").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("cliente").value%>"
								<% 
								col = col + 2
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("mercancia").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("PaisOrigen").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=contenedores%>"
								<% 
								contenedores = ""
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("Bultos").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("ETA").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("Documentos").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("revalidacion").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("previo").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("pago").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("despacho").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("entrada").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("KPICTE").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("KPIGRK").value%>"
								<% 
								col = col + 4
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("FechaCG").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("CG").value%>"
								<% 
								col = col + 1
								%>
								set celda = objHojaDatos.cells(<%=fila%>,<%=col%>)
								celda.value = "<%=RSkpis.Fields.Item("KPIADMIN").value%>"
								<% 
								col = 0
								RSkpis.MoveNext()
							Loop
							%>
					</script>
					<% 
				End If
				RSkpis.Close()
				ConnStr.Close()
				Set ConnStr = Nothing
				Set RSkpis = Nothing
			else
			  response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>")
			end if
		%>
	</BODY>
</HTML>