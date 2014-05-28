<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_libreria.asp" -->
 


 
 <style type="text/css">
.style20 {color: #FFFFFF}

table thead td {
	white-space: nowrap ;
	color: white;
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
'		26 Octubre, 2012

' Script que tiene como finalidad realizar el reporte de indicadores de desempeño según Samsung.

'Variables:
'			html_ : Variable que almacena una cadena con la estructura del documento a mostrar como resultado
'			

'inicio_ = Request.QueryString("txtDateIni")
'fin _ = Request.QueryString("txtDateFin")

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
	'Obteniendo valor de días de descanso según aduana
	dias_inhabiles_  = get_dias_descanso(Session("GAduana"))
	html= ""
	' Realizando la verificacion de las oficinas, para el caso de manzanillo no se realiza y el resto no maneja samsung
	
	if db_ = "lrz" Or db_ = "dai" or db_ = "tol" then
		
		if verificacion() = "TRUE" then ' Positivo : Si pasa la verificacion | Falso: Falta por asignar clasificacion a una o mas operaciones
			'consulta_ = genera_sql()
			html_ = html_ &"<p> genera reporte </p>"
			'html_ = genera_reporte(consulta_)
		else ' En caso de no pasar la validacion devolver una lista con las operaciones pendientes por asignar clasificacion para el indicador
			html_ = html_ & lista_pendientes()
		end if 'Fin de la prueba de validacion
	else
		consulta_ = genera_sql()
		html_ = html_ & "<table id=""reporte"">"
		html_ = html_ & genera_encabezado()
		html_ = html_ & genera_reporte(consulta_)
		html_ = html_ & "</table>"
		
	end if ' Fin de la verificacion para las oficinas que se debe validar la clasificacion de samsung
end if ' Fin de la verificacion de los permisos

'Response.Addheader "Content-Disposition", "attachment;filename=C:\wwwroot\wstemp\test.xls"
'Response.ContentType = "application/vnd.ms-excel"


'Este metodo  tiene como finalidad la creacion de la cadena de consulta
' tomando como parametro de entrada la oficina sobre la cual se genera 
' la consulta
function  genera_sql()

	sql_ = ""
	sql_ = sql_ & " SELECT  (@rownum:=@rownum+1) as ' ',"  
	sql_ = sql_ & " 	date_format(i.fecpag01,'%Y-%m') AnoMes, "  
	sql_ = sql_ & " 	i.refcia01 AS referencia,  "  
	sql_ = sql_ & " 	i.numped01 AS pedimento, "  
	sql_ = sql_ & " 	( "  
	sql_ = sql_ & " 		SELECT GROUP_CONCAT(DISTINCT art.desc05) "  
	sql_ = sql_ & " 		FROM " & db_ & "_extranet.d05artic AS art "  
	sql_ = sql_ & " 		WHERE art.refe05 = i.refcia01) AS Descpro,  "  
	sql_ = sql_ & " 	IF(c.fdoc01 IS NULL OR c.fdoc01 = '0000-00-00', 'No Capturada', DATE_FORMAT(c.fdoc01,'%d-%m-%Y')) AS Documentos, "  
	sql_ = sql_ & " 	IF(i.fecent01 IS NULL OR i.fecent01 = '0000-00-00', 'No Capturada', DATE_FORMAT(i.fecent01,'%d-%m-%Y')) AS fentrada,  "  
	sql_ = sql_ & " 	IF(c.frev01 IS NULL OR c.frev01 = '0000-00-00', 'No Capturada', DATE_FORMAT(c.frev01,'%d-%m-%Y')) AS Revalidacioc,  "  
	sql_ = sql_ & " 	IF(i.fecpag01 IS NULL OR i.fecpag01 = '0000-00-00', 'No Capturada', DATE_FORMAT(i.fecpag01,'%d-%m-%Y')) AS Pago,  "  
	sql_ = sql_ & " 	case when DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%m-%Y') is null then date_format(c.fdsp01 ,'%d-%m-%Y') else DATE_FORMAT(MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')),'%d-%m-%Y') end  AS Despacho,  "  
	sql_ = sql_ & " 		trackingbahia.get_dias_habiles( "  
	sql_ = sql_ & " 			case when MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) is null then c.fdsp01 else MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) end  , "  
	sql_ = sql_ & " 			case mid(i.refcia01,1,3) when 'SAP'  "  
	sql_ = sql_ & " 				then  case (select count(*) from " & db_ & "_extranet.sscont40 where refe01=refcia40)  "  
	sql_ = sql_ & " 						when 0 then c.frev01 "  
	sql_ = sql_ & " 						else i.fecent01 "  
	sql_ = sql_ & " 						end "  
	sql_ = sql_ & " 				else i.fecent01  "  
	sql_ = sql_ & " 			end , "  
	sql_ = sql_ & " 			" & dias_inhabiles_ &") as KPI_cfdsp01, "  	
	sql_ = sql_ & " 	( "  
	sql_ = sql_ & " 		SELECT GROUP_CONCAT(DISTINCT x2.cgas31) "  
	sql_ = sql_ & " 		FROM " & db_ & "_extranet.e31cgast AS x2 "  
	sql_ = sql_ & " 		INNER JOIN " & db_ & "_extranet.d31refer AS x3 ON x3.cgas31 = x2.cgas31 "  
	sql_ = sql_ & " 		WHERE x2.esta31 <>'C' AND x3.refe31 = i.refcia01) AS CG, "  
	sql_ = sql_ & " 	( "  
	sql_ = sql_ & " 		SELECT DATE_FORMAT(MAX(x2.fech31), '%d-%m-%Y') "  
	sql_ = sql_ & " 		FROM " & db_ & "_extranet.e31cgast AS x2 "  
	sql_ = sql_ & " 		INNER JOIN " & db_ & "_extranet.d31refer AS x3 ON x3.cgas31 = x2.cgas31 "  
	sql_ = sql_ & " 		WHERE x2.esta31 <>'C' AND x3.refe31 = i.refcia01) AS FCG, "  
	sql_ = sql_ & " 		 "  
	sql_ = sql_ & " 		trackingbahia.get_dias_habiles(  "  
	sql_ = sql_ & " 			( "  
	sql_ = sql_ & " 			SELECT MAX(x2.fech31) "  
	sql_ = sql_ & " 			FROM " & db_ & "_extranet.e31cgast AS x2 "  
	sql_ = sql_ & " 			INNER JOIN " & db_ & "_extranet.d31refer AS x3 ON x3.cgas31 = x2.cgas31 "  
	sql_ = sql_ & " 			WHERE x2.esta31 <>'C' AND x3.refe31 = i.refcia01) "  
	sql_ = sql_ & " 			, 	(case when MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) is null then c.fdsp01 else MAX(STR_TO_DATE(bs.Fechst01,'%d%m%Y')) end ) "  
	sql_ = sql_ & " 			,2) as KPI_ADMIN , "   
	sql_ = sql_ & " 	(  "  
	sql_ = sql_ & " 		SELECT IF(MAX(x2.frec31) IS NOT NULL AND MAX(x2.frec31) <> '0000-00-0', DATE_FORMAT(MAX(x2.frec31), '%d-%m-%Y'),'No capturada') "  
	sql_ = sql_ & " 		FROM " & db_ & "_extranet.e31cgast AS x2 "  
	sql_ = sql_ & " 		INNER JOIN " & db_ & "_extranet.d31refer AS x3 ON x3.cgas31 = x2.cgas31 "  
	sql_ = sql_ & " 		WHERE x2.esta31 <>'C' AND x3.refe31 = i.refcia01) AS 'facuse', "  
	sql_ = sql_ & " 	( "  
	sql_ = sql_ & " 		SELECT GROUP_CONCAT(DISTINCT date_format(x2.frec31,'%d-%m-%Y')) "  
	sql_ = sql_ & " 		FROM " & db_ & "_extranet.e31cgast AS x2 "  
	sql_ = sql_ & " 		INNER JOIN " & db_ & "_extranet.d31refer AS x3 ON x3.cgas31 = x2.cgas31 "  
	sql_ = sql_ & " 		WHERE x2.esta31 <>'C' AND x3.refe31 = i.refcia01) AS facuse33, "  
	sql_ = sql_ & " 	date_format(( "  
	sql_ = sql_ & " 		SELECT MAX(x2.frec31) "  
	sql_ = sql_ & " 		FROM " & db_ & "_extranet.e31cgast AS x2 "  
	sql_ = sql_ & " 		INNER JOIN " & db_ & "_extranet.d31refer AS x3 ON x3.cgas31 = x2.cgas31 "  
	sql_ = sql_ & " 		WHERE x2.esta31 <>'C' AND x3.refe31 = i.refcia01),'%d-%m-%Y' )AS 'Maxfacuse', "  
	sql_ = sql_ & "  	( "  
	sql_ = sql_ & " 		SELECT SUM(cta.chon31) AS campo "  
	sql_ = sql_ & " 		FROM " & db_ & "_extranet.e31cgast AS cta "  
	sql_ = sql_ & " 		INNER JOIN " & db_ & "_extranet.d31refer AS r ON cta.cgas31 = r.cgas31 "  
	sql_ = sql_ & " 		WHERE r.refe31 =i.refcia01 AND cta.esta31 = 'I') AS Honorarios, "  
	sql_ = sql_ & " 	( "  
	sql_ = sql_ & " 		SELECT IF(SUM(IF(cds.desdsc01 IS NULL, -1, IF(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1, 0))) > 0, 'ROJO', IF(SUM(IF(cds.desdsc01 IS NULL,-1, IF(cds.desdsc01 = 'ROJO SS' OR cds.desdsc01 = 'ROJO PS', 1, 0))) = 0, 'VERDE', 'SIN CAPTURAR')) "  
	sql_ = sql_ & " 		FROM trackingbahia.bit_soia AS bs "  
	sql_ = sql_ & " 		LEFT JOIN trackingbahia.cat_situaciones AS cs ON bs.cvesit01 = cs.cvesit01 "  
	sql_ = sql_ & " 		LEFT JOIN trackingbahia.cat_det_situaciones AS cds ON cds.detsit01 = bs.detsit01 "  
	sql_ = sql_ & " 		WHERE bs.frmsaai01 =i.refcia01) AS Semaforo, "  
	sql_ = sql_ & " 	( "  
	sql_ = sql_ & " 		SELECT SUM(vaduan02) AS campo "  
	sql_ = sql_ & " 		FROM " & db_ & "_extranet.ssfrac02 AS sf2 "  
	sql_ = sql_ & " 		WHERE sf2.refcia02 =i.refcia01 AND sf2.adusec02 =i.adusec01 AND sf2.patent02=i.patent01) AS 'Valor aduan'  "  
	sql_ = sql_ & " FROM  (SELECT @rownum:=0) r,"  
	sql_ = sql_ & " 	" & db_ & "_extranet.ssdagi01 AS i "  
	sql_ = sql_ & " 	LEFT JOIN " & db_ & "_extranet.c01refer AS c ON i.refcia01 = c.refe01 "  
	sql_ = sql_ & " 	LEFT JOIN " & db_ & "_extranet.d18mails AS d18 ON d18.cveeje18 = c.ejecli01 "  
	sql_ = sql_ & " 	LEFT JOIN " & db_ & "_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 "  
	sql_ = sql_ & " 	LEFT JOIN " & db_ & "_extranet.d31refer AS ctar ON ctar.refe31 = i.refcia01 "  
	sql_ = sql_ & " 	LEFT JOIN " & db_ & "_extranet.e31cgast AS cta ON cta.cgas31 = ctar.cgas31 AND cta.esta31 <> 'C' AND (cta.esta31= 'I') "  
	sql_ = sql_ & " 	LEFT JOIN " & db_ & "_extranet.d11movim AS d11 ON d11.refe11 = i.refcia01 AND d11.conc11 = 'ANT' "  
	sql_ = sql_ & " 	LEFT JOIN trackingbahia.bit_soia AS bs ON bs.frmsaai01 = i.refcia01 AND bs.Numpat01=i.patent01 AND bs.Detsit01 ='730'  "  
	sql_ = sql_ & " WHERE i.firmae01 IS NOT NULL  "  
	sql_ = sql_ & " 	AND i.firmae01 <> ''  "  
	sql_ = sql_ & " 	AND i.cveped01 <> 'R1'  "  
	sql_ = sql_ & " 	AND i.fecpag01 >= '"& fechaini_ &"' "  
	sql_ = sql_ & " 	AND i.fecpag01<= '" & fechafin_&"' "  
	sql_ = sql_ & " 	AND i.rfccli01 = 'SEM950215S98'  "  
	sql_ = sql_ & " GROUP BY i.refcia01   "  
	sql_ = sql_ & " ORDER BY i.refcia01;  "  


	genera_sql = sql_
	
end function 

' Este metodo Recorre la consulta y la genera en una tabla HTML
'function genera_xls( _obj_cursor)

'end function

function genera_encabezado()
	cabecera_ = "<thead style= ""background:'#006699'"" ><tr> " &_
				"<td > </td>" &_
				"<td >A&ntilde;o-Mes</td>" &_
				"<td >Referencia</td>" &_
				"<td >Pedimento </td>" &_
				"<td >Descripcion Fraccion </td>" &_
				"<td >Entrega de Documentos </td>" &_
				"<td >Fecha de Entrada </td>" &_
				"<td >Fecha de Revalidacion </td>" &_
				"<td >Fecha de Pago </td>" &_
				"<td >Fecha de Despacho </td>" &_
				"<td >KPI Operacion </td>" &_
				"<td >CG </td>" &_
				"<td >Fecha CG </td>" &_
				"<td >KPI Administrativo </td>" &_
				"<td >MM & ISPS </td>" &_
				"<td >Incidencias Causas de Desvio </td>" &_
				"<td >Honorarios </td>" &_
				"<td >Reconocimiento </td>" &_
				"<td >Transporte </td>" &_
				"<td >Valor Aduana </td>" &_
				"<td >Observaciones de Administracion </td>" &_
				"<td >Observaciones de Trafico </td>" &_
				"</tr></thead>"
	genera_encabezado = cabecera_
	
end function

function encabezado_pendientes()
	cabecera_ = "<thead style= ""background:'#006699'"" ><tr> " &_
				"<td > </td>" &_
				"<td >Pedimento</td>" &_
				"<td >Fecha de Pago</td>" &_
				"<td >Descripcion</td>" &_
				"<td >Dias KPI</td>" &_
				"</tr></thead>"
	encabezado_pendientes = cabecera_
end function

function lista_pendientes()
	' Generando consulta
	verificacion_sql_ = ""
	verificacion_sql_ = verificacion_sql_ & " select (@rownum:=@rownum+1) as ' ', ssi.refcia01 AS 'referencia',  "  
	verificacion_sql_ = verificacion_sql_ & " 		ssi.numped01 AS 'pedimento',  "  
	verificacion_sql_ = verificacion_sql_ & " 		ssi.fecpag01 as 'Fecha de Pago', "  
	verificacion_sql_ = verificacion_sql_ & " 		clas.descls64 as 'Descripcion', "  
	verificacion_sql_ = verificacion_sql_ & " 		clas.dias64 as 'Dias KPI' "  
	verificacion_sql_ = verificacion_sql_ & " from  (SELECT @rownum:=0) r,"  
	verificacion_sql_ = verificacion_sql_ & " 	"&db_&"_extranet.ssdagi01 as ssi "  
	verificacion_sql_ = verificacion_sql_ & " 	join "&db_&"_extranet.c01refer ref on ssi.refcia01 = ref.refe01 "  
	verificacion_sql_ = verificacion_sql_ & " 	left join "&db_&"_extranet.c64clsmg clas on ref.clsmg01=clas.clsmg64 "  
	verificacion_sql_ = verificacion_sql_ & " where ssi.firmae01 IS NOT NULL  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND ssi.firmae01 <> ''  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND ssi.cveped01 <> 'R1'  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND ssi.fecpag01 >= '" & fechaini_ & "'  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND ssi.fecpag01<= '" & fechafin_ & "'  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND (ref.clsmg01 is null " 
	verificacion_sql_ = verificacion_sql_ & " 	or ref.clsmg01 = '') "
	' Generando pagina
	loc_html_ = "<h2> Referencias que no tienen clasificacion para reporte de indicadores. </h2>"
	loc_html_ = loc_html_ & "<h3>Favor de capturar las clasificaciones en las siguientes referencias: </h3>"
	loc_html_ = loc_html_ & "<table id=""reporte"">"
	loc_html_ = loc_html_ & encabezado_pendientes
	loc_html_ = loc_html_ & genera_reporte(verificacion_sql_)
	loc_html_ = loc_html_ & "</table>"
	lista_pendientes = loc_html_
	
end function



' Este metodo se encarga de realizar la verificación si todas las operaciones registradas cuentan
' con una clasificicacion para el caso de Manzanillo no cuenta ya que son contenedores y carga suelta.
function verificacion( )
	' Consulta de verificacion, esta cuenta debe dar cero
	verificacion_sql_ = ""
	verificacion_sql_ = verificacion_sql_ & " select count(*) "  
	verificacion_sql_ = verificacion_sql_ & " from  "  
	verificacion_sql_ = verificacion_sql_ & " 	"&db_&"_extranet.ssdagi01 as ssi "  
	verificacion_sql_ = verificacion_sql_ & " 	join "&db_&"_extranet.c01refer ref on ssi.refcia01 = ref.refe01 "  
	verificacion_sql_ = verificacion_sql_ & " 	left join "&db_&"_extranet.c64clsmg clas on ref.clsmg01=clas.clsmg64 "  
	verificacion_sql_ = verificacion_sql_ & " where ssi.firmae01 IS NOT NULL  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND ssi.firmae01 <> ''  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND ssi.cveped01 <> 'R1'  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND ssi.fecpag01>= '"& fechaini_ &"'  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND ssi.fecpag01<= '"& fechafin_ &"'  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND ssi.rfccli01 = 'SEM950215S98'  "  
	verificacion_sql_ = verificacion_sql_ & " 	AND (ref.clsmg01 is null " 
	verificacion_sql_ = verificacion_sql_ & " 	or ref.clsmg01 = ''); "
	
	' Realizando conexion
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	' Abriendo un cursor
	Set RS = CreateObject("ADODB.RecordSet")
	Set RS = ConnStr.Execute(verificacion_sql_)
	
	valor_ = RS.fields(0) 
	
	'Do Until RS.EOF
		' Recorrido de los campos de la fila
	'	For i = 0 To RS.fields.Count-1
	'		valor_ = RS.fields(0) 
	'	Next
	'	RS.MoveNext()
	'Loop
	'Cerrando el cursor
	RS.close()
	'Cerrando conexion
	ConnStr.close()
	
	IF valor_ <> 0 THEN
		verificacion = "FALSE"
	ELSE
		verificacion = "TRUE"
	END IF
	
end function

 %>

 <HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE KPI DE SAMSUNG .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html_%>
	</BODY>
</HTML>