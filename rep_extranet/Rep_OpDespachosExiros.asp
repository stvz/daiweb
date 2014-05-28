<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000
On Error Resume Next

strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
strOficina=Request.Form("OficinaG")
Corte=Request.Form("opc")


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
mov=request.form("mov")

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

if  Session("GAduana") = "" then
	html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
else

	if mov = "i" then
		movi = "IMPORTACION "
		query = GeneraSQL(mov)
	elseif mov="e" then
		movi="EXPORTACION "
		query=GeneraSQL(mov)
	elseif mov="a" then
		movi = "IMPORTACION / EXPORTACION"
		query = GeneraSQL(mov)

	end if
	'response.write(query)
	'response.end()
	nocolumns = 22
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
								celdahead("Referencia Erycia","#81BEF7") &_
								celdahead("Kilos","#81BEF7") &_
								celdahead("Material","#81BEF7") &_
								celdahead("Contenedores","#81BEF7")&_
								celdahead("Cantidad de Unidades","#81BEF7") &_
								celdahead("Proveedor","#81BEF7") &_
								celdahead("Numero de Legajo","#81BEF7") &_
								celdahead("Facturas","#81BEF7") &_
								celdahead("Numero de Orden de Compra","#81BEF7") &_
								celdahead("Buque","#81BEF7") &_
								celdahead("Eta","#81BEF7") &_
								celdahead("BL","#81BEF7") &_
								celdahead("Numero de Pedimentos","#81BEF7") &_
								celdahead("Terminal","#81BEF7") &_
								celdahead("Fecha Libre de Almacenajes contenedor","#81BEF7") &_
								celdahead("Fecha libre de demoras","#81BEF7")&_
								celdahead("Fecha de Desconsolidacion","#81BEF7")&_
								celdahead("Fecha libre de almacenaje tuberia","#81BEF7")&_
								celdahead("Se tiene prioridades","#81BEF7")&_
								celdahead("Destinos","#81BEF7")&_
								celdahead("Desconsolidado","#81BEF7")&_
								celdahead("Obsercaciones","#81BEF7")
						header = header &	"</tr>"
				dim snco 
				
			Do Until RSops.EOF
							snco="#FFFFFF"
							datos = datos & "<tr> " &_
							celdadatos("",snco) &_
							celdadatos(RSops.Fields.Item("Kilos").Value,snco) &_
							celdadatos(RSops.Fields.Item("Material").Value,snco) &_
							celdadatos(RSops.Fields.Item("Contenedor").Value,snco) &_
							celdadatos(RSops.Fields.Item("Cantidad de Unidades").Value,snco) &_
							celdadatos(RSops.Fields.Item("Proveedor").Value,snco) &_
							celdadatos("",snco) &_
							celdadatos(RSops.Fields.Item("Facturas").Value,snco) &_
							celdadatos(RSops.Fields.Item("Numero de Orden de Compra").Value,snco) &_
							celdadatos(RSops.Fields.Item("Barco").Value,snco) &_
							celdadatos(RSops.Fields.Item("Eta").Value,snco) &_
							celdadatos(RSops.Fields.Item("BL").Value,snco) &_
							celdadatos(RSops.Fields.Item("Numero de Pedimentos").Value,snco) &_
							celdadatos(RSops.Fields.Item("Terminal").Value,snco) &_
							celdadatos("",snco) &_
							celdadatos("",snco) &_
							celdadatos("",snco) &_
							celdadatos("",snco) &_
							celdadatos("",snco) &_
							celdadatos("",snco) &_
							celdadatos("",snco) &_														
							celdadatos("",snco) 
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

function GeneraSQL(op)
	SQL = ""
	if strOficina <> "Todas" then
		'Se selecciono una oficina en especifico
		if op="a" then 
			SQL=subSQL("IMPORTACION","i",strOficina)
			SQL= SQL & " UNION ALL "& subSQL("EXPORTACION","e",strOficina)
		elseif op="i" then 
			SQL=subSQL("IMPORTACION","i",strOficina)
		elseif op="e" then 
			SQL=subSQL("EXPORTACION","e",strOficina)
		end if
	elseif strOficina="Todas" then 
		dim strOficina2
		for ii=1 to 6
			'Aqui se realiza el llamado de la digitalizacion de todas las oficinas segun el tipo de operacion seleccionado
			select case ii
				case 1
					strOficina2="rku"
				case 2
					strOficina2="dai"
				case 3
					strOficina2="sap"
				case 4
					strOficina2="lzr"	
				case 5
					strOficina2="tol"
				case 6
					strOficina2="ceg"
				end select
				if op="a" then 
					SQL=SQL & subSQL("IMPORTACION","i",strOficina2)
					 SQL=SQL &" UNION ALL "& subSQL("EXPORTACION","e",strOficina2)
				elseif op="i" then
					SQL= SQL & subSQL("IMPORTACION",op,strOficina2)
				elseif op="e" then
					SQL=SQL & subSQL("EXPORTACION",op,strOficina2)
				end if 
				if ii < 6 then 
				 SQL=SQL &" UNION ALL "& chr(13) & chr(10)
				end if
		next
		'response.write(SQL)
		'response.end()
	end if
	GeneraSQL = SQL
	
end function

function subSQL (operacion,movimiento,oficina)'Aqui se construye el query segun el tipo de operacion y la oficina
dim Aduana 
if oficina="rku" then
	Aduana="Veracruz"
elseif oficina="dai" then
	Aduana="AICM"
elseif oficina="sap" then 
	Aduana="Manzanillo" 
elseif oficina="tol" then
	Aduana="Toluca"
elseif oficina="lzr" then
	Aduana="Lazaro Cardenas"
else 
	Aduana="Altamira"
end if
	SQL=""
		SQL=" select r.fdsp01 Despachado, i.pesobr01 'Kilos', " 
		SQL=SQL & "group_concat(distinct replace(f.d_mer102,'.','')) 'Material', "
		SQL=SQL & "group_concat(distinct con.numcon40) 'Contenedor', "
		SQL=SQL & "i.totbul01 'Cantidad de Unidades', "
		SQL=SQL & "group_concat(distinct fa.nompro39) 'Proveedor', "
		SQL=SQL & "group_concat(distinct fa.numfac39 separator '| ')'Facturas', "
		SQL=SQL & "group_concat(distinct d.pedi05 separator ' |') 'Numero de Orden de Compra',  "
		SQL=SQL & "b.nomb06 'Barco', "
		SQL=SQL & "r.feta01 'Eta', "
		SQL=SQL & "r.feorig01 'BL', "
		SQL=SQL & "concat_ws(' ',DATE_FORMAT(i.fecpag01,'%y'),left(i.adusec01,2), i.patent01, i.numped01)'Numero de Pedimentos', "
		SQL=SQL & "group_concat(distinct d01.zona01) Terminal "
		SQL=SQL & "from "&oficina&"_extranet.ssdag"&movimiento&"01 as i "
		SQL=SQL & "left join "&oficina&"_extranet.c01refer as r on r.refe01=i.refcia01 "
		SQL=SQL & "left join "&oficina&"_extranet.c06barco as b on b.clav06=r.cbuq01 "
		SQL=SQL & "left join "&oficina&"_extranet.ssfrac02 as f on f.refcia02=i.refcia01 and f.adusec02=i.adusec01 and f.patent02=i.patent01 "
		SQL=SQL & "left join "&oficina&"_extranet.d05artic as d on d.refe05 =f.refcia02 and d.agru05=f.ordfra02  "
		SQL=SQL & "left join "&oficina&"_extranet.sscont40 as con on con.refcia40=i.refcia01 and con.adusec40=i.adusec01 and con.patent40=i.patent01 "
		SQL=SQL & "left join "&oficina&"_extranet.d01conte as d01 on d01.refe01=i.refcia01 "
		SQL=SQL & "left join "&oficina&"_extranet.ssguia04 as g on g.refcia04=i.refcia01 and f.adusec02=i.adusec01 and f.patent02 =i.patent01 "
		SQL=SQL & "left join "&oficina&"_extranet.ssfact39 as fa on fa.refcia39=i.refcia01 and fa.adusec39=i.adusec01 and fa.patent39=i.patent01 and fa.numfac39=d.fact05 "
		SQL=SQL & "left join "&oficina&"_extranet.ssprov22 as p on p.cvepro22=fa.cvepro39 "
		SQL=SQL & "where i.rfccli01 in('"&strFiltroCliente&"' )  "
		if Corte="p" then 
			SQL=SQL & " and ((i.fecpag01 between  '"&DateI&"'  and '"&DateF&"' ) and r.fdsp01='0000-00-00') " 
		else
			SQL=SQL & "and r.fdsp01 between '"&DateI&"'  and '"&DateF&"' and r.fdsp01<>'0000-00-00' " 
		end if
		SQL=SQL & "and i.firmae01 is not null "
		SQL=SQL & "and i.firmae01<>'' "
		SQL=SQL & "group by i.refcia01 "
	'response.write(subSQL)
	'response.end()
	subSQL=SQL
end function 
%>
<HTML>
	<HEAD>
		<TITLE>::.... DESPACHOS .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>