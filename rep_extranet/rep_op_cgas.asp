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
tbody th, tbody td {
  text-align: center;
}

 </style>

<%

' Por: Manuel Alejandro Estevez Fernandez
'		29 Noviembre, 2012

' Script que tiene como finalidad realizar el reporte donde se muestren las operaciones por cuenta de gastos de un cliente de pedimento en periodo de facturacion

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
	
	cliente_ = Request.Form("txtCliente")
	
	
	Response.Addheader "Content-Disposition", "attachment;filename=Rep Operaciones por Cuenta de Gastos del "&fechaini_&" al "&fechafin_&".xls"
	Response.ContentType = "application/vnd.ms-excel"

	html= ""
	consulta_ = genera_sql()
	html_ = html_ & "<h3> Operaciones por Cuenta de Gastos</h3>"
	html_ = html_ & "<p> Del " & fechaini_& " al " & fechafin_
	html_ = html_ & "<p>"
	html_ = html_ & "<table id=""reporte"">"
	html_ = html_ & genera_encabezado()
	html_ = html_ & genera_cuerpo(consulta_)
	html_ = html_ & "</table>"
	Response.Write(html_)
	Response.End()
	
end if ' Fin de la verificacion de los permisos


'Este metodo  tiene como finalidad la creacion de la cadena de consulta
' tomando como parametro de entrada la oficina sobre la cual se genera 
' la consulta
function  genera_sql()
	
	if cliente_ = "Todos" Then
		sql_cliente_ = " "
	else
		sql_cliente_ = "cvecli01 = "& cliente_ &" and "
	end if

 ' Consulta 2012 11 29
	sql_ = "   " & _ 
			  "select cli_.nomcli18 as 'Cliente Pedimento',"& _ 
			  "cg_.cgas31 'Cuenta de Gastos', date_format(cg_.fech31,'%d/%m/%Y') 'Fecha CG' ,op_.refcia01 'Referencia', tipo 'Tipo',  "& _ 
			  "op_.patent01 'Patente', op_.adusec01 'Aduana Sección', "& _ 
			  "op_.numped01 'Pedimento',date_format(op_.fecpag01,'%d/%m/%Y') 'Fecha de Pago', "& _ 
			  "op_.cveped01 'Clave Pedimento', ifnull(group_concat(guia_.numgui04),'') Guias "& _ 
			  ",fact_.numfac39 'Factura', date_format(fact_.fecfac39,'%d/%m/%Y') 'Fecha Factura',fact_.edocum39 'Cove', fact_.terfac39 'Incoterm',fact_.valmex39 'Valor Moneda Extranjera', fact_.valdls39 'Valor Dls.' "& _ 
			  "from e31cgast cg_ join d31refer dref_ using(cgas31) "& _ 
			  "left join ( "& _ 
			  "select refcia01,numped01,fecpag01,cveped01,cveadu01,cvesec01, patent01,'impo' as tipo,adusec01,cvecli01 "& _ 
			  "from ssdagi01 in_ where "& sql_cliente_ &" in_.firmae01 != '' and in_.firmae01 is not null "& _ 
			  "union all "& _ 
			  "select refcia01,numped01,fecpag01,cveped01,cveadu01,cvesec01, patent01,'expo' as tipo,adusec01,cvecli01 "& _ 
			  "from ssdage01 exp_ where "& sql_cliente_ &" exp_.firmae01 != '' and exp_.firmae01 is not null "& _ 
			  ") as op_ on op_.refcia01 = dref_.refe31 "& _ 
			  "left join ssguia04 guia_ on op_.refcia01 = guia_.refcia04 "& _ 
			  "left join ssfact39 fact_ on fact_.refcia39 = op_.refcia01 "& _
			  " left join ssclie18 cli_ on op_.cvecli01 = cli_.cvecli18 " & _
			  "where cg_.fech31 between '" & fechaini_ & "' and '" & fechafin_ & "' and fact_.terfac39 != '' and cg_.esta31 ='I' "& _ 
			  "group by fact_.numfac39 "& _ 
			  ""
	genera_sql = sql_
	
end function 

' Este metodo Recorre la consulta y la genera en una tabla HTML
'function genera_xls( _obj_cursor)

'end function

function genera_encabezado()
	cabecera_ = "<thead><tr bgcolor= ""#0066cc"" > " &_
				"<td >Cliente Pedimento</td>" &_
				"<td >Cuenta de Gastos</td>" &_
				"<td >Fecha CG</td>" &_
				"<td >Referencia</td>" &_
				"<td >Tipo </td>" &_
				"<td >Patente </td>" &_
				"<td >Aduana Secci&oacute;n </td>" &_
				"<td >Pedimento </td>" &_
				"<td >Fecha de Pago </td>" &_
				"<td >Clave Pedimentos </td>" &_
				"<td >Gu&iacute;as </td>" &_
				"<td >Factura </td>" &_
				"<td >Fecha Factura </td>" &_
				"<td >Cove </td>" &_
				"<td >Incoterm </td>" &_
				"<td >Valor Moneda Extranjera </td>" &_
				"<td >Valor Dls. </td>" &_
				"</tr></thead>"
	genera_encabezado = cabecera_
	
end function



function genera_cuerpo(sql)
	
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; DATABASE="& db_ &"_extranet; OPTION=16427"
		
	Set RS = CreateObject("ADODB.RecordSet")
	Set RS = ConnStr.Execute(sql)
	tabla_ = ""
    If Not RS.EOF Then
		
		' Variable de bandera para resaltar los renglones pares
		ban_ = 1
        Do Until RS.EOF
			if ban_ = 1 then
				tabla_ = tabla_ & "<tr>"
				ban_ = 0
			else
				tabla_ = tabla_ & "<tr bgcolor= ""#EBECF2"">"
				ban_ = 1
			end if
			For i = 0 To RS.fields.Count-1
				on error resume next
					tabla_ = tabla_ & "<td align=""center"">" & RS.fields(i) & "</td>"
				if Err.Number<> 0 Then
					tabla_ = tabla_ & "<td align=""center"">" & VarType(RS.fields(i)) & "</td>"

				end if
				Err.Clear
            Next
			tabla_ = tabla_ & "</tr>" &VBcrlf
            RS.MoveNext
        Loop
    End If
    RS.Close
    Set RS=Nothing
	genera_cuerpo = tabla_
 End function

 ' Function get_cliente( _cliente )
	' consulta_ = "Select nomcli18 from ssclie18 where cvecli18 = "&_cliente
	
	' Set ConnStr = Server.CreateObject ("ADODB.Connection")
	' ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; DATABASE="& db_ &"_extranet; OPTION=16427"
	' Set RS = CreateObject("ADODB.RecordSet")
	' Set RS = ConnStr.Execute(consulta_)
	' Do Until RS.EOF
		' nombre_cliente_ = RS.fields(0)
	' Loop
    ' RS.Close
    ' Set RS=Nothing
	
	' get_cliente = nombre_cliente_
 
 ' End Function

 %>

 <HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE OPERACIONES POR CUENTA DE GASTOS .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html_%>
	</BODY>
</HTML>