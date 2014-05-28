 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
 
 <style type="text/css">
.style20 {color: #FFFFFF}
 </style>

<%
dim Corte, treporte,titulo

 Corte=request.form("Corte")
 treporte=request.form("treporte")




if  Session("GAduana") = "" then
	html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
else
	jnxadu=Session("GAduana")

	select case jnxadu
			case "VER"
				strOficina="rku"
				nomOficina="Veracruz"
			case "MEX"
				strOficina="dai"
				nomOficina="Mexico"
			case "MAN"
				strOficina="sap"
				nomOficina="Manzanillo"
			case "TAM"
				strOficina="ceg"
				nomOficina="Altamira"
			case "LZR"
				strOficina="lzr"
				nomOficina="Lazaro"
			case "TOL"
				strOficina="tol"
				nomOficina="Toluca"
	end select
		
	if treporte="Buque" then
		titulo=": : REPORTE POR BUQUE : :"
		nocolumns=14
	elseif treporte="Referencia" then
		titulo=": : REPORTE POR REFERENCIA : :"
		nocolumns = 16
	end if
	
 
			query = GeneraSQL(Corte,treporte,strOficina)

			
	
	'Response.Write(query)
	'Response.End
	
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)

	IF RSops.BOF = True And RSops.EOF = True Then
		
		Response.Write("No hay datos para esas condiciones")
	Else
		
	
			Response.Addheader "Content-Disposition", "attachment;filename=Rep_Control_Cierre_Buques.xls"
			Response.ContentType = "application/vnd.ms-excel"
		
		info = 	"<table  width = ""2929""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""4"" face=""Arial"">" &_
											"<b>" &_
												"GRUPO REYES KURI, S.C" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""3"" face=""Arial"">" &_
											"<b>" &titulo&_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
								"</td>" &_
							"</tr>" &_
				"</table>"
				
		header = 			"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr bgcolor = ""#006699"" class = ""boton"">"
							if treporte="Referencia" then
								header = header & celdahead("Registro de la referencia")
								header = header & celdahead("F. Despacho") 
								header = header & celdahead("Referencia") 
								header = header & celdahead("Cliente") 
								header = header & celdahead("Bultos") 
								header = header & celdahead("Barco")
								header = header & celdahead("Registro de barco")
								header = header & celdahead("Fecha de registro del barco") 
								header = header & celdahead("Naviera") 
								header = header & celdahead("Fecha de cierre documental") 
								header = header & celdahead("Hora de cierre documental") 
								header = header & celdahead("Fecha de cierre fisico") 
								header = header & celdahead("Hora de cierre fisico") 
								header = header & celdahead("Ultima modificacion de la fecha de cierre documental") 				 
								header = header & celdahead("Dias para cierre fisico") 	
								header = header & celdahead("Dias para cierre documental") 
							else
								header = header & celdahead("Oficina")
								header = header & celdahead("Registro del barco")
								header = header & celdahead("Barco")
								header = header & celdahead("Naviera")
								header = header & celdahead("Referencias")
								header = header & celdahead("Clientes")
								header = header & celdahead("F. Cierre Doc.")
								header = header & celdahead("Hr. Cierre Doc.")
								header = header & celdahead("F. Cierre Fisico")
								header = header & celdahead("Hr. Cierre Fisico")
								header = header & celdahead("Fecha de registro Barco")
								header = header & celdahead("Ultima modificacion del cierre documental")
								header = header & celdahead("Dias para cierre fisico") 	
								header = header & celdahead("Dias para cierre documental") 
							end if
									
						

		header = header &	"</tr>"
		dim snco
			snco="#FFFFFF"
		Do Until RSops.EOF
						
						
						datos = datos & "<tr> " 
							if treporte="Referencia" then
								datos = datos &	celdadatos(RSops.Fields.Item("frec01").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("fdsp01").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Referencia").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Cliente").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Bultos").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Barco").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("regb06").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Fecha de registroGRK").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Naviera").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("F. CierreD").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Hr. CierreD").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("F. CierreF").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Hr. CierreF").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("FechaHoraRegistroActualizado").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("DiasPCierreFisico").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("DiasPCierrDocumental").Value) 
							else
								datos = datos &	celdadatos(nomOficina) 
								datos = datos &	celdadatos(RSops.Fields.Item("regb06").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Barco").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Naviera").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Referencia2").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Cliente2").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("F. CierreD").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Hr. CierreD").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("F. CierreF").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("Hr. CierreF").Value)
								datos = datos &	celdadatos(RSops.Fields.Item("Fecha de registroGRK").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("FechaHoraRegistroActualizado").Value)
								datos = datos &	celdadatos(RSops.Fields.Item("DiasPCierreFisico").Value) 
								datos = datos &	celdadatos(RSops.Fields.Item("DiasPCierrDocumental").Value) 
							end if
							
								
							
							
						datos = datos &	"</tr>"
			Rsops.MoveNext()
		Loop
	Response.Write(info & header & datos & "</table><br>")
	Response.End()
	ConnStr.Close()
	html = info & header & datos & "</table><br>"
	End If
end if

function celdahead(texto)'Celda de encabezado de la tabla
	cell = "<td bgcolor = ""#006699"" width=""100"" nowrap>" &_
				"<center>" &_
					"<strong>" &_
						"<font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" &_
							texto &_
						"</font>" &_
					"</strong>" &_
				"</center>" &_
			"</td>"
	celdahead = cell
end function

function celdadatos(texto)'Celda de datos de la tabla
'On error resume next
	If IsNull(texto) = True Or texto = "" Then
		texto = " "
	End If
	cell = 	"<td align=""center"">" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdadatos = cell
end function

function GeneraSQL(tipocorte,tiporeporte,oficina)
Ordenar=""
SQL=""
if tipocorte="general" then 
	tipocorte=" "
elseif tipocorte="margendias" then
	tipocorte=" (r.fdsp01 ='0000-00-00' or r.fdsp01 is null or (r.fdsp01 >=date_sub(curdate(), interval 15 day))) and  "
end if

SQL="SELECT r.frec01,reb.regb06 ,r.fdsp01 ,e.refcia01 Referencia, group_concat(e.refcia01) Referencia2 ,group_concat(distinct e.nomcli01) Cliente2, e.nomcli01 Cliente, e.totbul01 Bultos ,bar.nomb06 Barco , "
SQL=SQL & "c55.nom01 Naviera , reb.cierrd06 'F. CierreD' ,reb.hrcid06 'Hr. CierreD' , reb.cierrf06 'F. CierreF' , reb.hrcif06 'Hr. CierreF',reb.falt06 'Fecha de registroGRK', "
SQL=SQL & "(select fec.dati06  from "&oficina&"_Extranet.d06fecci   AS fec where fec.regb06 = r.regb01  order by fec.orden06 desc limit 1) FechaHoraRegistroActualizado, "
SQL=SQL & "(DATEDIFF(reb.cierrf06 , date(sysdate())))DiasPCierreFisico, (DATEDIFF(reb.cierrd06  , date(sysdate())))DiasPCierrDocumental, "
SQL=SQL & "(select group_concat(e31.cgas31) from "&oficina&"_extranet.d31refer as d "
	SQL=SQL & "inner join "&oficina&"_Extranet.e31cgast as e31 on e31.cgas31 =d.cgas31 "
	SQL=SQL & "where d.refe31 =e.refcia01 and e31.esta31 ='I') as cg "
SQL=SQL & "FROM "&oficina&"_Extranet.ssdage01 AS e "
	SQL=SQL & "LEFT JOIN "&oficina&"_extranet.c01refer AS r ON r.refe01 =e.refcia01  "
		SQL=SQL & "LEFT JOIN "&oficina&"_Extranet.c06rebar AS reb ON reb.clav06 =r.cbuq01 AND reb.regb06 =r.regb01  and r.regb01 !='' "
		SQL=SQL & "left join "&oficina&"_Extranet.c06barco as bar on bar.clav06 =r.cbuq01 "
		SQL=SQL & "left join "&oficina&"_Extranet.c55navie as c55 on c55.cve01 =bar.navi06 and c55.Status55 ='T' "
	SQL=SQL &"WHERE  "  &tipocorte
	SQL=SQL & " year(e.fecpag01)>=2012 and r.csit01 <>'FIN' and r.modo01<>'C' "
	SQL=SQL & "AND r.tipo01 =2 and reb.tipo06 =2 "
	
	
	if tiporeporte="Buque" then  
		SQL=SQL& " group by reb.regb06 having cg is null  order by DiasPCierrDocumental asc"
		
	elseif  tiporeporte="Referencia" then 
	     SQL=SQL &" group by e.refcia01 having cg is null order by DiasPCierrDocumental asc"
	end if
	'response.write(SQL)
	'response.end()
	
GeneraSQL=SQL
end function
%>

<HTML>
	<HEAD>
		<TITLE>::.... CONTROL DE CIERRE DE BUQUES .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>
