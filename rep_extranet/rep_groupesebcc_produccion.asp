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
	'fechaini_ = Request.form("txtDateIni")
	'fechafin_ = Request.form("txtDateFin")
	cgas_ = Request.form("cgas")
	
	Response.Addheader "Content-Disposition", "attachment;filename=Rep Conceptos vs Cuentas Contables GroupeSeb "& oficina_ &" Cuenta "& cgas_ &".xls"
	Response.ContentType = "application/vnd.ms-excel"

	html= ""
	consulta_ = genera_sql()
	'html_ = html_ & consulta_
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



 ' Nueva consulta 2012 11 26		  
	sql_ = " select proveedor, cuenta_gastos,fecha_cuenta,Pedimento,Cuenta,Subcuenta,Monto,lower(iva), concepto  " & _ 
  "   from (   " & _ 
  "      " & _ 
  "    	select case mid(refe31,1,3)      " & _ 
  "    			when 'ALC' then '02231047'     " & _ 
  "    			when 'RKU' then '02231031'     " & _ 
  "    		end as proveedor,     " & _ 
  "    		cgas31 as cuenta_gastos,     " & _ 
  "    		date_format(fech31,'%m/%d/%Y') as fecha_cuenta,     " & _ 
  "    		concat(mid(refe31,4,2),mid(adusec01,1,2),'-',patent01,'-',numped01 ) as Pedimento,     " & _ 
  "    		clave as Clave,     " & _ 
  "    		case armado_.clave_concepto when 1 or 37      " & _ 
  "    			then (select cat_.cuenta from sistemas.cat_grpseb cat_ where cat_.cvecon = armado_.clave_concepto and cat_.tiptax =impuesto)      " & _ 
  "    			else (select cat_.cuenta from sistemas.cat_grpseb cat_ where cat_.cvecon = armado_.clave_concepto)     " & _ 
  "    		end as ""Cuenta"",     " & _ 
  "    		case armado_.clave_concepto when 1 or 37      " & _ 
  "    			then (select cat_.sbcta from sistemas.cat_grpseb cat_ where cat_.cvecon = armado_.clave_concepto and cat_.tiptax =impuesto)      " & _ 
  "    			else (select cat_.sbcta from sistemas.cat_grpseb cat_ where cat_.cvecon = armado_.clave_concepto)     " & _ 
  "    		end as ""Subcuenta"",     " & _ 
  "    		case armado_.clave_concepto when 1 or 37 then import36 else monto end as ""Monto"",     " & _ 
  "   			(select cat_.tax  from sistemas.cat_grpseb cat_ where cat_.cvecon = armado_.clave_concepto limit 1) as iva,    " & _ 
  "    		armado_.clave_concepto as clave_con,      " & _ 
  "   			(select pa_.desc21 from c21paghe pa_ where pa_.clav21= armado_.clave_concepto) as concepto,   " & _ 
  "   			1 as orden_   " & _ 
  "    from      " & _ 
  "    (     " & _ 
  "        select semi.cgas31,semi.refe31,semi.fech31,semi.clie31, semi.patent01,semi.numped01,semi.clave,clave_concepto,monto , adusec01    " & _ 
  "        from     " & _ 
  "            (     " & _ 
  "                select *     " & _ 
  "                from      " & _ 
  "                (    " & _ 
  "                        select cg_.cgas31,dref_.refe31,cg_.fech31,cg_.clie31     " & _ 
  "                      	from e31cgast cg_      " & _ 
  "                         join d31refer dref_ on cg_.cgas31 = dref_.cgas31     " & _ 
  "                         where dref_.cgas31 = '"& cgas_ &"' " & _ 
  "  " & _ 
  "                ) as crs_     " & _ 
  "                left join     " & _ 
  "                (     " & _ 
  "                        select in_.refcia01,in_.adusec01, in_.patent01,in_.numped01, 'GIT5'  as clave     " & _ 
  "                        from     " & _ 
  "                                (select distinct dref_.refe31     " & _ 
  "                                from d31refer dref_     " & _ 
  "                                where dref_.cgas31 = '"& cgas_ &"' " & _ 
  "                                ) as refes_     " & _ 
  "                                     " & _ 
  "                                join ssdagi01 in_      " & _ 
  "                                on in_.refcia01 = refes_.refe31     " & _ 
  "                        union all     " & _ 
  "                        select in_.refcia01,in_.adusec01,in_.patent01,in_.numped01, 'GIT5'  as clave     " & _ 
  "                        from     " & _ 
  "                                (select distinct dref_.refe31     " & _ 
  "                                from d31refer dref_     " & _ 
  "                                where dref_.cgas31 = '"& cgas_ &"' " & _ 
  "                                ) as refes_     " & _ 
  "                                     " & _ 
  "                                join ssdage01 in_      " & _ 
  "                                on in_.refcia01 = refes_.refe31     " & _ 
  "                ) as ops_     " & _ 
  "                on crs_.refe31 = ops_.refcia01     " & _ 
  "            )     " & _ 
  "            as semi     " & _ 
  "            left join         " & _ 
  "                (     " & _ 
  "                        select dpag_.cgas21, pag_.conc21 as Clave_concepto,     " & _ 
  "                                sum(case pag_.deha21 when 'A' then  dpag_.mont21 else dpag_.mont21*-1 end) as monto     " & _ 
  "                        from d21paghe dpag_  " & _ 
  "                        join  e21paghe pag_ on year(pag_.fech21) = year(dpag_.fech21) and pag_.foli21 = dpag_.foli21 and pag_.tmov21= dpag_.tmov21      " & _ 
  "                        where pag_.tpag21 != 3 and  dpag_.cgas21 = '"& cgas_ &"' " & _ 
  "                        group by dpag_.cgas21,pag_.conc21     " & _ 
  "                ) as montos_cgs on montos_cgs.cgas21 = semi.cgas31     " & _ 
  "    ) as armado_     " & _ 
  "    left join      " & _ 
  "    	(     " & _ 
  "    		select refcia36,1 as Clave_concepto,case cveimp36      " & _ 
  "    								when 1 then 'dta'      " & _ 
  "    								when 3 then 'iva'     " & _ 
  "    								when 15 then 'prv'      " & _ 
  "    								when 50 then 'dfc'      " & _ 
  "    								when 6 then 'igi'      " & _ 
  "    								when 18 then 'eci'      " & _ 
  "    							end as impuesto, import36     " & _ 
  "    		from (select distinct dref_.refe31     " & _ 
  "    				from d31refer dref_         " & _ 
  "    				where dref_.cgas31 = '"& cgas_ &"' " & _ 
  "    				) as refes_     " & _ 
  "    		left join sscont36 cn_ on cn_.refcia36 = refes_.refe31     " & _ 
  "    	)      " & _ 
  "        as impuestos_ on impuestos_.refcia36 = armado_.refe31 and armado_.clave_concepto = impuestos_.clave_concepto     " & _ 
  "      " & _ 
  "   union all   " & _ 
  "      " & _ 
  "   select    " & _ 
  "   		CASE MID(refe31,1,3) WHEN 'ALC' THEN '02231047' WHEN 'RKU' THEN '02231031' END AS ""No de Proveedor"",   " & _ 
  "   		cgas31 ""Cuenta de Gastos"",    " & _ 
  "   		date_format(fech31,'%m/%d/%Y'),   " & _ 
  "   		case when in_.refcia01 is not null    " & _ 
  "   				then  CONCAT(mid(in_.refcia01,4,2),mid(in_.adusec01,1,2),'-',in_.patent01,'-',in_.numped01)   " & _ 
  "   			else CONCAT(mid(exp_.refcia01,4,2),mid(exp_.adusec01,1,2),'-',exp_.patent01,'-',exp_.numped01)   " & _ 
  "   		end 'pedimento',   " & _ 
  "   		'GIT5' as ""CLAVE"",   " & _ 
  "   		Cuenta,   " & _ 
  "   		Subcuenta,   " & _ 
  "   		chon31 as Monto,   " & _ 
  "   		Iva,   " & _ 
  "   		""Clave Concepto"",   " & _ 
  "   		Concepto,   " & _ 
  "   		orden_   " & _ 
  "   from(    " & _ 
  "      " & _ 
  " 	  	select cg_.cgas31,dref_.refe31,cg_.fech31 ,'39000401' Cuenta, '60' Subcuenta , cg_.chon31, 'Si' as ""Iva"",'' as ""Clave Concepto"",   " & _ 
  " 	  	'honorarios' as ""Concepto"", 2 as orden_   " & _ 
  " 	  	       from e31cgast cg_    " & _ 
  " 	          join	d31refer dref_ on cg_.cgas31 = dref_.cgas31   " & _ 
  " 	  	where dref_.cgas31 = '"& cgas_ &"' " & _ 
  " 	  	union all   " & _ 
  " 	  	   " & _ 
  " 	  	select   " & _ 
  " 	  		cgas32 ""Cuenta de Gastos"",    " & _ 
  " 	  		refe31,   " & _ 
  " 	  		fech32,   " & _ 
  " 	  		""Cuenta"",   " & _ 
  " 	  		""Subcuenta"",   " & _ 
  " 	  		mont32,   " & _ 
  " 	  		'Si',   " & _ 
  " 	  		ttar32 ,   " & _ 
  " 	  		dcrp32,   " & _ 
  " 	  		4   " & _ 
  " 	  	from (   " & _ 
  " 		  		SELECT DISTINCT dref_.refe31   " & _ 
  " 		  		FROM d31refer dref_   " & _ 
  " 		  		WHERE dref_.cgas31 = '"& cgas_ &"' " & _ 
  " 	  		) AS refes_    " & _ 
  " 	  		join e32rserv rs_ on rs_.refe32 = refes_.refe31    " & _ 
  " 	  		join d32rserv using(refe32)   " & _ 
  " 	  " & _ 
  " 	  	   " & _ 
  " 	  	UNION ALL   " & _ 
  " 	  	select cg_.cgas31,dref_.refe31,cg_.fech31, '' Cuenta, '' Subcuenta , cg_.caho31, 'si' ,""Clave Concepto"",'adicional a honorarios' as ""Concepto"", 3   " & _ 
  " 	  	       from e31cgast cg_    " & _ 
  " 	          join	d31refer dref_ on cg_.cgas31 = dref_.cgas31   " & _ 
  " 	  	 	where dref_.cgas31 = '"& cgas_ &"' " & _ 
  " 	  	union all   " & _ 
  " 	  	select cg_.cgas31,dref_.refe31,cg_.fech31,'44551101' Cuenta,'70' Subcuenta, (cg_.caho31+ cg_.chon31 + cg_.csce31 ) * (cg_.piva31/100), 'No' as ""Aplica Iva"",""Clave Concepto"",'IVA' as ""Concepto"",5   " & _ 
  " 	  	       from e31cgast cg_    " & _ 
  " 	          join	d31refer dref_ on cg_.cgas31 = dref_.cgas31   " & _ 
  " 	  	where dref_.cgas31 = '"& cgas_ &"' " & _ 
  "   	                  " & _ 
  "   ) as montos_   " & _ 
  "   	left join ssdagi01 in_ on in_.refcia01 = refe31   " & _ 
  "   	left join ssdage01 exp_ on exp_.refcia01 = refe31   " & _ 
  "      " & _ 
  "   where chon31 != 0   " & _ 
  "   ) as Reporte_   " & _ 
  "   order by Pedimento, orden_  "
 


	genera_sql = sql_
	
end function 

' Este metodo Recorre la consulta y la genera en una tabla HTML
'function genera_xls( _obj_cursor)

'end function

function genera_encabezado()
	cabecera_ = "<thead><tr bgcolor= ""yellow""> " &_
				"<td >No Proveedor</td>" &_
				"<td >Cuenta de Gastos</td>" &_
				"<td >Fecha de Cuenta </td>" &_
				"<td >Pedimento </td>" &_
				"<td >Cuenta </td>" &_
				"<td >Sub Cuenta </td>" &_
				"<td >Monto </td>" &_
				"<td >Aplica IVA </td>" &_
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
			'tabla_ = tabla_ & "<td align=""center""> " & RS.fields(0) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(1) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(2) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(3) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(4) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(5) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(6) & "</td>"
			'tabla_ = tabla_ & "<td align=""right"">" & RS.fields(7) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(8) & "</td>"
			'on error resume next
			'	tabla_ = tabla_ & "<td >" & RS.fields(9) & "</td>"
			'	tabla_ = tabla_ & "<td >" & RS.fields(10) & "</td>"
			'Next
			' Recorrido de los campos de la fila
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


 %>

 <HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE CUENTAS CONTABLES GROUPE SEB .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html_%>
	</BODY>
</HTML>