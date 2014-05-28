<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_libreria.asp" -->
 


 
 <style type="text/css">
.style20 {color: #FFFFFF}

table thead td {
	white-space: nowrap ;
	color: black;
	text-align: center;
	 padding:0 20px;
}
table thead{
	font-family: Arial,Georgia, "Times New Roman",Times, serif;
	font-size: 16;
	font-weight: bold;
	height: 30px;
	
}
table tbody{
	font-family: Arial,Georgia, "Times New Roman",Times, serif;
	font-size: 12;
	text-align: center;
}

 </style>

<%

' Por: Manuel Alejandro Estevez Fernandez
'		13 Noviembre, 2012

' Script que tiene como finalidad realizar el reporte de relacion cuentas contables de groupe seb y los conceptos de los pagos hechos.

'Variables:
'			html_ : Variable que almacena una cadena con la estructura del documento a mostrar como resultado
'			

' Comprobacion de permisos

if  Session("GAduana") = "" then
	html_ = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
else
	' Obteniendo el prefijo de la db y la oficina
	db_ = get_db(Session("GAduana"))
	oficina_ = get_nombre_oficina(Session("GAduana"))
	' Asignacion de fechas  a variables
	fechaini_ = Request.form("txtDateIni")
	fechafin_ = Request.form("txtDateFin")
	
	Response.Addheader "Content-Disposition", "attachment;filename=Rep Conteo de Unidades "& oficina_ &" del "& fechaini_ &" al "& fechafin_ &".xls"
	Response.ContentType = "application/vnd.ms-excel"

	html= ""
	consulta_ = genera_sql()
	'html_ = html_ & consulta_
	html_ = html_ & "<table id=""reporte"">"
	html_ = html_ & genera_encabezado()
	html_ = html_ & genera_reporte(consulta_)
	html_ = html_ & "</table>"
	Response.Write(html_)
	Response.End()
	
end if ' Fin de la verificacion de los permisos

'Response.Addheader "Content-Disposition", "attachment;filename=C:\wwwroot\wstemp\test.xls"
'Response.ContentType = "application/vnd.ms-excel"



'Este metodo  tiene como finalidad la creacion de la cadena de consulta
' tomando como parametro de entrada la oficina sobre la cual se genera 
' la consulta
function  genera_sql()

	sql_ = "" &_
" SELECT p.cvecli01 as 'Clave', p.Nomcli01 AS 'Cliente'," & _  
"         t.Patent01 AS 'Patente', " & _  
"         t.pedimento AS 'Pedimento',  " & _  
"         t.refe01 AS 'Referencia',  " & _  
"         CASE t.tipveh " & _  
"                 WHEN 'H' THEN 'Torthon'  " & _  
"                 WHEN 'F' THEN 'Fugon'  " & _  
"                 WHEN 'G' THEN 'Gondola'" & _  
"                 WHEN 'C' THEN 'Camion'  " & _  
"                 WHEN 'T' THEN 'Tolva'  " & _  
"                 WHEN 'l' THEN 'Full'  " & _  
"                 WHEN 'J' THEN 'Jaula'  " & _  
"                 WHEN 'P' THEN 'Pipa' " & _  
"                 WHEN 'D' THEN 'Doble Pipa' " & _  
"                 WHEN 'R' THEN 'Carro Tanque' " & _  
"                 ELSE 'Invalido'  " & _  
"         END AS 'Tipo Vehiculo', " & _  
"         count(t.tipveh) AS 'No. Unidades' " & _  
"     FROM rku_cpsimples.pedimentos AS p " & _  
"     LEFT JOIN rku_cpsimples.tcepartidas as t ON t.Refe01 = p.Refe01 " & _  
"     WHERE t.fectick >= '"& fechaini_ &"' AND t.fectick <= '"& fechafin_ &"' AND p.Status <> 3 " & _  
"     GROUP BY t.refe01, t.tipveh " & _  
"     ORDER BY referencia, 'Tipo Vehiculo' " 
	genera_sql = sql_
	
end function 

' Este metodo Recorre la consulta y la genera en una tabla HTML
'function genera_xls( _obj_cursor)

'end function

function genera_encabezado()
	cabecera_ = "<thead><tr bgcolor= ""#8B0000"" > " &_
				"<td ><center><b><font color=""white""> Clave </font></center></td>" &_
				"<td ><center><b><font color=""white""> Cliente </font></center></td>" &_
				"<td ><center><b><font color=""white"">Patente</font></center></td>" &_
				"<td ><center><b><font color=""white"">Pedimento</font></center></td>" &_
				"<td ><center><b><font color=""white"">Referencia </font></center></td>" &_
				"<td ><center><b><font color=""white"">Tipo Vehiculo </font></center></td>" &_
				"<td ><center><b><font color=""white"">No. Unidades </font></center></td>" &_
				"</tr></thead>"
	genera_encabezado = cabecera_
	
end function






 %>

 <HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE CONTEO DE UNIDADES DE CARGILL .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html_%>
	</BODY>
</HTML>