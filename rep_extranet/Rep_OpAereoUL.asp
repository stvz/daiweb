<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000
On Error Resume Next

strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"i.cvecli01")
strOficina=Request.Form("OficinaG")
Corte=Request.Form("opc")
TOperacion=Request.Form("mov")


	fi=trim(request.form("fi"))
	ff=trim(request.form("ff"))
	Vrfc=Request.Form("rfcCliente")
	bclientes=Request.Form("Enviar")


	DiaI = cstr(datepart("d",fi))
	Mesi = cstr(datepart("m",fi))
	AnioI = cstr(datepart("yyyy",fi))
	DateI = Anioi & "/" & Mesi & "/" & DiaI

	DiaF = cstr(datepart("d",ff))
	MesF = cstr(datepart("m",ff))
	AnioF = cstr(datepart("yyyy",ff))
	DateF = AnioF & "/" & MesF & "/" & DiaF
	
if not permi = "" then
	permi = "  and (" & permi & ") "
end if
AplicaFiltro = False
strFiltroCliente = ""
strFiltroCliente = request.Form("rfcCliente")
Tiporepo = Request.Form("TipRep")

if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
	blnAplicaFiltro = true
end if
if blnAplicaFiltro then
	permi = " AND cvecli01 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
	permi = ""
end if
dim car
Car=Cargas()
if  Session("GAduana") = "" or Car=false  then
	if Car=false then 
		html = "<br></br><div align=""center""><p  class=""Titulo1"">:: INFORMACION EN ACTUALIZACION, ESPERE UN MOMENTO E INTENTE DE NUEVO ::</div></p></div>"
	else
		html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"	
	end if
else

	query = GeneraSQL()

	
	'response.write(query)
	'response.end()
	nocolumns = 34
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)
	
	IF RSops.BOF = True And RSops.EOF = True Then
		
		Response.Write("No hay datos para esas condiciones")
	Else
		
		if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;filename=Rep_Reporte_"&DiaI&"-"&Mesi&"_"&DiaF&"-"&MesF&".xls"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		info = 	"<table  width = ""2929""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""4"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td  align=""center"" colspan=""" & nocolumns & """></font></p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p><font color=""#000000"" size=""4"" face=""Arial, Helvetica, sans-serif"">: : R E P O R T E : :</font></p>" &_
											"<p><font color=""#000000"" size=""2"" face=""Arial, Helvetica, sans-serif"">"&movi&" DEL "&DateI&" al "&DateF&"</font>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
		header = 			"<tr class = ""boton"">" &_
								celdahead("COMPAÑIA","#0B6121") &_
								celdahead("REFERENCIA","#0B6121") &_
								celdahead("PLANTA DESTINO","#0B6121") &_
								celdahead("CLASE DE PRODUCTO","#0B6121")&_
								celdahead("PROVEEDOR","#0B6121") &_
								celdahead("PEDIMENTO","#0B6121") &_
								celdahead("CANT DE OPERACIONES","#0B6121") &_
								celdahead("CANTIDAD COMERCIAL","#0B6121") &_
								celdahead("UNIDAD DE MEDIDA (PZA,CAJAS,KILOS,ETC.)","#0B6121") &_
								celdahead("PAIS ORIGEN","#0B6121") &_
								celdahead("CIUDAD ORIGEN","#0B6121") &_
								celdahead("PAIS DESTINO","#0B6121") &_
								celdahead("FECHA DE EMBARQUE","#0B6121") &_
								celdahead("INCOTERM","#0B6121") &_
								celdahead("COSTO DEL FLETE AEREO","#0B6121") &_
								celdahead("TIPO DE CAMBIO","#0B6121") &_
								celdahead("LINEA AEREA","#0B6121")&_
								celdahead("NUMERO DE PALLETS","#0B6121")&_
								celdahead("NUMERO DE BULTOS","#0B6121")&_
								celdahead("PESO","#0B6121")&_
								celdahead("DESCRIPCION DEL MATERIAL","#0B6121")&_
								celdahead("TIPO DE MATERIAL","#0B6121")&_
								celdahead("EXP","#0B6121")&_
								celdahead("FLETE INTERNACIONAL USD","#0B6121")&_
								celdahead("UNIDADES CM","#0B6121")&_
								celdahead("COSTO POR UNIDAD MARITIMO","#0B6121")&_
								celdahead("UNIDADES OC AEREA","#0B6121")&_
								celdahead("COSTO AEREO USD","#0B6121")&_
								celdahead("COSTO POR UNIDAD AÉREO","#0B6121")&_
								celdahead("DIF","#0B6121")&_
								celdahead("EXP REAL (USD)","#0B6121")&_
								celdahead("ODC","#0B6121")&_
								celdahead("FACTURA DEL PROVEEDOR","#0B6121")&_
								celdahead("CLAVE DE PEDIMENTO","#0B6121")
						header = header &	"</tr>"
				dim snco 
				dim aux , cf,tc,bultos,pesobr
				aux=""
				cf=""
				tc="" 
				bultos=""
				pesobr=""
			Do Until RSops.EOF
			'response.write(RSops.Fields.Item("Planta").Value)
							if aux="" or aux<>RSops.Fields.Item("refcia01").Value then
								aux=RSops.Fields.Item("refcia01").Value
								cf=RSops.Fields.Item("COSTO DEL FLETE AEREO USD").Value
								tc=RSops.Fields.Item("Tipo de Cambio").Value
								bultos=RSops.Fields.Item("Numero de bultos").Value
								pesobr=RSops.Fields.Item("Peso Bruto").Value
							else
								cf=""
								tc=""
								bultos=""
								pesobr=""
							end if 
							snco="#FFFFFF"
							datos = datos & "<tr> " &_
							celdadatos(RSops.fields.Item("Compania").Value,snco) &_
							celdadatos(RSops.Fields.Item("refcia01").Value,snco) 
							datos= datos & celdadatos(RSops.Fields.Item("Planta").value,snco) &_
							celdadatos(RSops.Fields.Item("Clase Producto").Value,snco) &_
							celdadatos(RSops.Fields.Item("Proveedor").Value,snco) &_
							celdadatos(RSops.Fields.Item("Pedimento").Value,snco) &_
							celdadatos(RSops.Fields.Item("CANT DE OPERACIONES").Value,snco) &_
							celdadatos(RSops.Fields.Item("Cant Comercial").Value,snco) &_
							celdadatos(RSops.Fields.Item("Unidad de Medida").Value,snco) &_
							celdadatos(RSops.Fields.Item("Pais Origen").Value,snco) &_
							celdadatos(RSops.Fields.Item("Ciudad Origen").Value,snco) &_
							celdadatos(RSops.Fields.Item("Pais Destino").Value,snco) &_
							celdadatos(RSops.Fields.Item("Fecha de embarque").Value,snco) &_
							celdadatos(RSops.Fields.Item("Incoterm").Value,snco) &_
							celdadatos(cf,snco) &_
							celdadatos(tc,snco) &_
							celdadatos(RSops.Fields.Item("Linea Aerea").Value,snco) &_
							celdadatos(RSops.Fields.Item("Numero de Pallets").Value,snco) &_
							celdadatos(bultos,snco) &_
							celdadatos(pesobr,snco) &_
							celdadatos(RSops.Fields.Item("Descripcion del Material").Value,snco) &_														
							celdadatos(RSops.Fields.Item("Tipo Material").Value,snco) &_
							celdadatos(RSops.Fields.Item("EXP").Value,snco) &_
							celdadatos(RSops.Fields.Item("FLETE INTERNACIONAL USD").Value,snco) &_
							celdadatos(RSops.Fields.Item("UNIDADES CM").Value,snco) &_
							celdadatos(RSops.Fields.Item("COSTO POR UNIDAD MARITIMO").Value,snco) &_
							celdadatos(RSops.Fields.Item("UNIDADES OC AEREA").Value,snco) &_
							celdadatos(RSops.Fields.Item("COSTO AÉREO USD").Value,snco) &_
							celdadatos(RSops.Fields.Item("COSTO POR UNIDAD AEREO").Value,snco) &_
							celdadatos(RSops.Fields.Item("DIF").Value,snco) &_
							celdadatos(RSops.Fields.Item("EXP REAL (USD)").Value,snco) &_
							celdadatos(RSops.Fields.Item("ODC").Value,snco) &_
							celdadatos(RSops.Fields.Item("Factura del proveedor").Value,snco) &_
							celdadatos(RSops.Fields.Item("Clave de pedimento").Value,snco) 
							datos = datos &	"</tr>"
				Rsops.MoveNext()
			Loop
	Response.Write(info & header & datos & "</table><br>")
	Response.End()
	ConnStr.Close()
	html = info & header & datos & "</table><br>"
	End If
end if

function celdahead(texto,colorh)'Celda de encabezado de la tabla
	cell = "<td bgcolor = """&colorh&""" width=""200"" nowrap>" &_
				"<center>" &_
					"<strong>" &_
						"<font color=""#FFFFFF"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
							texto &_
						"</font>" &_
					"</strong>" &_
				"</center>" &_
			"</td>"
	celdahead = cell
end function

function celdadatos(texto,pcolor)'Celda de datos de la tabla
On error resume next

	If IsNull(texto) = True Or texto = "" Then
		texto = "&nbsp;"
	End If
	 dim c 
	 c=chr(34)
	cell = 	"<td align=""center""nowrap bgcolor="&c&pcolor&c&" >" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdadatos = cell
end function

function GeneraSQL()
	SQL = ""
	if strOficina <> "Todas" then
		'Se selecciono una oficina en especifico
		
		SQL=subSQL(strOficina)
		
	elseif strOficina="Todas" then 
		dim strOficina2
		for ii=1 to 2
			'Aqui se realiza el llamado de la digitalizacion de todas las oficinas segun el tipo de operacion seleccionado
			select case ii
				case 1
					strOficina2="dai"
				case 2
					strOficina2="tol"
				end select
					
					SQL=SQL & subSQL(strOficina2)
				
				if ii < 2 then 
				 SQL=SQL &" UNION ALL "& chr(13) & chr(10)
				end if
		next
		'response.write(SQL)
		'response.end()
	end if
	GeneraSQL = SQL
	
end function

function subSQL (oficina)'Aqui se construye el query segun el tipo de operacion y la oficina
	SQL=""
		SQL=" select i.nomcli01 'Compania',i.refcia01, "
		SQL=SQL&" cast(if(i.cvecli01 in('11000','15001'),'CDU', "
		SQL=SQL&"if(i.cvecli01 in('11001','15002'),'TULTITLAN', "
		SQL=SQL&"if(i.cvecli01 in('11002','15003'),'LERMA', "
SQL=SQL&"if(i.cvecli01 in('11003','15004'),'CIVAC', "
SQL=SQL&"if(i.cvecli01 in('11004'),'ESPECIALES',if(i.cvecli01='15005','TULTITLAN, LERMA Y CIVAC',i.cvecli01)))))) as char) as Planta, "
SQL=SQL&"'' as 'Clase Producto', "
SQL=SQL&"i.nompro01 Proveedor,  "
SQL=SQL&"concat_ws(' ',DATE_FORMAT(i.fecpag01,'%y'),left(i.adusec01,2), i.patent01, i.numped01) Pedimento, "
SQL=SQL&"'' as 'CANT DE OPERACIONES', "
SQL=SQL&"f.cancom02'Cant Comercial', "
SQL=SQL&"u.descri31 'Unidad de Medida', "
SQL=SQL&"i.cvepod01 'Pais Origen', "
SQL=SQL&"r.ptoemb01 'Ciudad Origen', "
SQL=SQL&"if(i.tipopr01=1,'Mexico',i.cvepod01) as 'Pais Destino', "
SQL=SQL&"i.fecpag01 as'Fecha de embarque', "
'SQL=SQL&" if(i.cveped01='G1',r.forf01,fa.terfac39) 'Incoterm', "
SQL=SQL&" if(i.cveped01='G1',if(i.cvecli01=15004 and (r.forf01='' or r.forf01='FOB') ,'DDU',if(i.cvecli01=15002 and (r.forf01='' or r.forf01='FOB'),'DAP',r.forf01)),fa.terfac39) 'Incoterm', "
SQL=SQL&" r.fleint01 as 'COSTO DEL FLETE AEREO USD', "
'SQL=SQL&"round((i.fletes01/i.tipcam01 ),2) as 'COSTO DEL FLETE AEREO USD', "
SQL=SQL&"i.tipcam01 as 'Tipo de Cambio', "
SQL=SQL&"if (l.desc01='',l.dir01,l.desc01) 'Linea Aerea', "
SQL=SQL&"'' as 'Numero de Pallets', "
SQL=SQL&"i.totbul01 as 'Numero de bultos', "
SQL=SQL&"i.pesobr01 as 'Peso Bruto', "
SQL=SQL&"f.d_mer102 as 'Descripcion del Material', "
SQL=SQL&"group_concat(distinct d.tpmerc05 separator ' ') as 'Tipo Material', "
SQL=SQL&"'' as 'EXP', "
SQL=SQL&"'' as 'FLETE INTERNACIONAL USD', "
SQL=SQL&"'' AS 'UNIDADES CM', "
SQL=SQL&"'' AS 'COSTO POR UNIDAD MARITIMO', "
SQL=SQL&"'' AS 'UNIDADES OC AEREA', "
SQL=SQL&"'' AS 'COSTO AÉREO USD', "
SQL=SQL&"'' AS 'COSTO POR UNIDAD AEREO', "
SQL=SQL&"'' AS 'DIF', "
SQL=SQL&"'' AS 'EXP REAL (USD)', "
SQL=SQL&"group_concat(distinct d.pedi05) AS 'ODC', "
SQL=SQL&" if(i.cveped01='G1',group_concat(distinct d.fact05 separator ' / '),group_concat(distinct fa.numfac39)) 'Factura del proveedor', "
SQL=SQL&"i.cveped01 as 'Clave de pedimento' "
SQL=SQL&"FROM "&oficina&"_extranet.ssdag"&TOperacion&"01 as i "
SQL=SQL&"LEFT JOIN "&oficina&"_extranet.c01refer as r on r.refe01=i.refcia01 "
SQL=SQL&"LEFT JOIN "&oficina&"_extranet.c01airln as l on l.cvela01=r.cvela01  "
SQL=SQL&"LEFT JOIN "&oficina&"_extranet.ssfrac02 as f on f.refcia02=i.refcia01 and f.adusec02=i.adusec01 and f.patent02=i.patent01 "
SQL=SQL&"LEFT JOIN "&oficina&"_extranet.d05artic as d on d.refe05=f.refcia02 and d.agru05=f.ordfra02  "
SQL=SQL&"LEFT JOIN "&oficina&"_extranet.ssumed31 as u on u.clavem31=f.u_medc02 "
SQL=SQL&"LEFT JOIN "&oficina&"_extranet.ssfact39 as fa on fa.refcia39=i.refcia01 and fa.adusec39=i.adusec01 and fa.patent39=i.patent01 and fa.numfac39=d.fact05 and fa.terfac39<>''"
SQL=SQL&"where i.rfccli01 in('"&strFiltroCliente&"') and i.fecpag01>='"&Datei&"' and i.fecpag01<='"&DateF&"' and i.firmae01<>'' and i.firmae01 is not null "
SQL=SQL&"group by i.refcia01, f.ordfra02 "
	'response.write(SQL)
	'response.end()
	subSQL=SQL
end function 
function Cargas()

	sqlAct="select count(*) as conteo from intranet.ban_extranet as b where b.m_bandera <> 'NA'"

	Set act2= Server.CreateObject("ADODB.Recordset")
	
	conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	
	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	if act2.fields("conteo").value>0 then
		'cambiar 
		Cargas=false
	else 
		Cargas=true
	end if
	act2.Close()
end function 
%>
<HTML>
	<HEAD>
		<TITLE>::.... Reporte de Aereos .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>