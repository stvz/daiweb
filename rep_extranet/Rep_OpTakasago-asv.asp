<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000
On Error Resume Next

strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
strOficina=Request.Form("OficinaG")


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
dim car
Car=Cargas()
if  Session("GAduana") = "" or Car=false then
	if Car=false then 
		html = "<br></br><div align=""center""><p  class=""Titulo1"">:: INFORMACION EN ACTUALIZACION, ESPERE UN MOMENTO E INTENTE DE NUEVO ::</div></p></div>"
	else
		html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"	
	end if
	
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
	nocolumns = 45
	
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
								celdahead("PO SAP","#81BEF7") &_
								celdahead("No. Pedimento","#81BEF7") &_
								celdahead("Fec. Pago","#81BEF7") &_
								celdahead("Referencia","#81BEF7")&_
								celdahead("No. Guia","#81BEF7") &_
								celdahead("Factura solicitada para pago","#81BEF7") &_
								celdahead("Comentarios","#81BEF7") &_
								celdahead("Flete Pagado por","#81BEF7") &_
								celdahead("Proveedor","#81BEF7") &_
								celdahead("TaxID","#81BEF7") &_
								celdahead("Procedencia","#81BEF7") &_
								celdahead("Origen","#81BEF7") &_
								celdahead("Agente Aduanal","#81BEF7") &_
								celdahead("Aduana","#81BEF7") &_
								celdahead("Codigo","#81BEF7") &_
								celdahead("Producto","#81BEF7") &_
								celdahead("Nombre Quimico","#81BEF7") &_
								celdahead("Fraccion","#81BEF7") &_
								celdahead("Incoterm","#81BEF7") &_
								celdahead("Cantidad","#81BEF7") &_
								celdahead("Precio Proveedor","#81BEF7")&_
								celdahead("Monto Total","#81BEF7")&_
								celdahead("Moneda","#81BEF7")&_
								celdahead("Factura Proveedor","#81BEF7")&_
								celdahead("Fec. Factura Proveedor","#81BEF7")&_
								celdahead("Linea Flete Inter","#81BEF7")&_
								celdahead("Monto Flete Inter","#81BEF7")&_
								celdahead("Moneda","#81BEF7")&_
								celdahead("Quien paga el flete internacional","#81BEF7")&_
								celdahead("Anticipo","#81BEF7")&_
								celdahead("Monto del anticipo","#81BEF7")&_
								celdahead("Fondeo","#81BEF7")&_
								celdahead("Monto Fondeo","#81BEF7")&_
								celdahead("TLC","#81BEF7")&_
								celdahead("IGI","#81BEF7")&_
								celdahead("DTA","#81BEF7")&_
								celdahead("PRV","#81BEF7")&_
								celdahead("Honorarios","#81BEF7")&_
								celdahead("Maniobras","#81BEF7")&_
								celdahead("Desconsolidacion/REVALIDACION","#81BEF7")&_
								celdahead("Serv. Compl.","#81BEF7")&_
								celdahead("IVA PEDIMENTO","#81BEF7")&_
								celdahead("","#81BEF7")&_
								celdahead("Fecha entrega TDM","#81BEF7")&_
								celdahead("Linea Flete Nac.","#81BEF7")&_
								celdahead("Monto Flete Nac.","#81BEF7")&_
								celdahead("Moneda","#81BEF7")&_
								celdahead("Factura Flete Nac.","#81BEF7")&_
								celdahead("No. Cta. de Gastos","#81BEF7")&_
								celdahead("Fecha entrega TDM CG","#81BEF7")&_
								celdahead("Fecha entrega Contabilidad","#81BEF7")&_
								celdahead("Usuario","#81BEF7")&_
								celdahead("Division","#81BEF7")
						header = header &	"</tr>"
				dim snco 
				dim Refrenciaactual, refaux,PRV ,Maniobras,Honorarios,Desconsolidacion,ServCom, Beneficiario,MontofleteSI
				refaux=""
				Beneficiario=""
				MontofleteSI=0
				PRV=0
				Maniobras=0
				Desconsolidacion
				Honorarios=0
				ServCom=0
			Do Until RSops.EOF
							snco="#FFFFFF"
							Referenciaactual=RSops.Fields.Item("Referencia").Value
							if Referenciaactual<>refaux or refaux="" then 
								refaux=Referenciaactual
								PRV=RSops.Fields.Item("PRV").Value
								
								if RSops.Fields.Item("No. Cta. de Gastos").Value<>"" then 
									Honorarios=RetornaHonorarioBit(RSops.Fields.Item("Referencia").Value,RSops.Fields.Item("No. Cta. de Gastos").Value,RSops.Fields.Item("Toperacion").value,"H")
									Honorarios=Honorarios+RetornaHonorarioBit(RSops.Fields.Item("Referencia").Value,RSops.Fields.Item("No. Cta. de Gastos").Value,RSops.Fields.Item("Toperacion").value,"S")
								elseif RSops.Fields.Item("No. Cta. de Gastos").Value="" then
								
									Honorarios=HonorarioCot(RSops.Fields.Item("Referencia").Value)
								end if
								Maniobras=RSops.Fields.Item("Maniobras").Value
								Desconsolidacion=RSops.Fields.Item("Desconsolidacion").Value
								ServCom= RetornaServCom(RSops.Fields.Item("Referencia").Value,RSops.Fields.Item("Toperacion").value)
								if ServCom=0 then 
									ServCom=RSops.Fields.Item("Serv. Compl.").Value
								end if
								Beneficiario=RetornaFleteNacional(RSops.Fields.Item("Referencia").Value,"Beneficiario")
								MontofleteSI=RetornaFleteNacional(RSops.Fields.Item("Referencia").Value,"ImporteSIVA")
							else 
								PRV=""
								Honorarios=""
								Desconsolidacion=""
								Maniobras=""
								ServCom=""
								Beneficiario=""
								MontofleteSI=""
							end if
			
							datos = datos & "<tr> " &_
							celdadatos(RSops.Fields.Item("PO SAP").Value,snco) &_
							celdadatos(RSops.Fields.Item("No. Pedimento").Value,snco) &_
							celdadatos(RSops.Fields.Item("Fec. Pago").Value,snco) &_
							celdadatos(RSops.Fields.Item("Referencia").Value,snco) &_
							celdadatos(RSops.Fields.Item("No. Guia").Value,snco) &_
							celdadatos("",snco) &_
							celdadatos("",snco) &_
							celdadatos(RSops.Fields.Item("Flete Pagado por").Value,snco) &_
							celdadatos(RSops.Fields.Item("Proveedor").Value,snco) &_
							celdadatos(RSops.Fields.Item("TaxID").Value,snco) &_
							celdadatos(RSops.Fields.Item("Procedencia").Value,snco) &_
							celdadatos(RSops.Fields.Item("Origen").Value,snco) &_
							celdadatos(RSops.Fields.Item("Agente").Value,snco) &_
							celdadatos(RSops.Fields.Item("Aduana").Value,snco) &_
							celdadatos("",snco) &_
							celdadatos("",snco) &_
							celdadatos(RSops.Fields.Item("Producto").Value,snco) &_
							celdadatos(RSops.Fields.Item("Fraccion").Value,snco) &_
							celdadatos(RSops.Fields.Item("Incoterm").Value,snco) &_
							celdadatos(RSops.Fields.Item("Cantidad").Value,snco) &_
							celdadatos(RSops.Fields.Item("Precio Proveedor").Value,snco) &_
							celdadatos(RSops.Fields.Item("MontoTotal").Value,snco) &_
							celdadatos(RSops.Fields.Item("Moneda").Value,snco) &_
							celdadatos(RSops.Fields.Item("Factura Proveedor").Value,snco) &_
							celdadatos(RSops.Fields.Item("FechaProveedor").Value,snco) &_
							celdadatos(RSops.Fields.Item("Linea Flete Inter").Value,snco) &_
							celdadatos(RSops.Fields.Item("Monto Flete Inter").Value,snco) &_
							celdadatos(RSops.Fields.Item("Moneda3").Value,snco) &_
							celdadatos("",snco) &_
							celdadatos(retornaMontoAnticipo(RSops.Fields.Item("Referencia").value,"ANT",mid(RSops.Fields.Item("Referencia").value,1,3),"CountAnt"),snco) &_
							celdadatos(retornaMontoAnticipo(RSops.Fields.Item("Referencia").value,"ANT",mid(RSops.Fields.Item("Referencia").value,1,3),"campo"),snco) &_
							celdadatos(RSops.Fields.Item("Fondeo").Value,snco) &_
							celdadatos(RSops.Fields.Item("Monto Fondeo").Value,snco) &_
							celdadatos(RSops.Fields.Item("TLC").Value,snco) &_
							celdadatos(RSops.Fields.Item("IGI").Value,snco) &_
							celdadatos(RSops.Fields.Item("DTA").Value,snco) &_
							celdadatos(PRV,snco) &_
							celdadatos(Honorarios,snco) &_
							celdadatos(Maniobras,snco) &_
							celdadatos(Desconsolidacion,snco) &_
							celdadatos(ServCom,snco) &_
							celdadatos(RSops.Fields.Item("IVA").Value,snco) &_
							celdadatos("",snco) &_
							celdadatos(RSops.Fields.Item("Fecha Entrega TDM").Value,snco) &_
							celdadatos(Beneficiario,snco) &_
							celdadatos(MontofleteSI,snco) &_
							celdadatos(RSops.Fields.Item("Moneda2").Value,snco) &_
							celdadatos(RSops.Fields.Item("Factura Flete Nac.").Value,snco) &_
							celdadatos(RSops.Fields.Item("No. Cta. de Gastos").Value,snco) &_
							celdadatos(RSops.Fields.Item("Fecha entrega TDM CG").Value,snco) &_
							celdadatos(RSops.Fields.Item("Fecha entrega Contabilidad").Value,snco) &_
							celdadatos(RSops.Fields.Item("Usuario").Value,snco) &_
							celdadatos(RSops.Fields.Item("Division").Value,snco) 
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
function HonorarioCot(Referencia)
dim ofi,valor

 
ofi=mid(Referencia,1,3)
if ofi="ALC" then
	ofi="LZR"
elseif ofi="PAN" then 
	ofi="DAI"
end if
	sqlAct="select cast(ifnull(sum(d.mont08),'No se encontro en cotizacion')as char) as Monto from "&ofi&"_extranet.d08cotsv as d where d.refe08 in('"&Referencia&"') and d.clav08 in('HON','VAL','SCO','COM') "

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&ofi&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then

	valor = act2.fields("Monto").value

   HonorarioCot = valor
 
 end if
 HonorarioCot=valor
	
end function 
function subSQL (operacion,movimiento,oficina)'Aqui se construye el query segun el tipo de operacion y la oficina
dim Aduana , Agente
if oficina="rku" then
	Aduana="Veracruz"
	Agente="Grupo Reyes Kuri S.C."
elseif oficina="dai" then
	Aduana="Mexico"
	Agente="Despachos Aéreos Integrados S.C."
elseif oficina="sap" then 
	Aduana="Manzanillo" 
	Agente="Servicios Aduanales del Pacifico S.C."
elseif oficina="tol" then
	Aduana="Toluca"
	Agente="Comercio Exterior del Golfo S.C."
elseif oficina="lzr" then
	Aduana="Lazaro Cardenas"
	Agente="Servicios Aduanales del Pacifico S.C."
else 
	Aduana="Altamira"
	Agente="Comercio Exterior del Golfo S.C."
end if
	SQL=""
			SQL=SQL & "select 		c.rcli01, if(i.tipopr01=1,'i','e') as Toperacion, " 
		SQL=SQL & "group_concat(distinct cast(a.pedi05 as char) separator ' / ')as 'PO SAP', " 
		SQL=SQL & "cast(concat(i.cveadu01,' ', i.patent01 , ' ', i.numped01 )as char) as 'No. Pedimento', " 
		SQL=SQL & "i.fecpag01 as 'Fec. Pago', " 
		SQL=SQL & "i.refcia01 as 'Referencia', " 
		SQL=SQL & "group_concat(distinct g.numgui04)  as 'No. Guia', " 
		SQL=SQL & "replace(c.tartm01,'FLETE PAGADO X ','')as 'Flete Pagado por', " 
		SQL=SQL & "i.nompro01 as 'Proveedor', " 
		SQL=SQL & "i.taxpro01 as 'TaxID', " 
		SQL=SQL & "c.paisem01  as 'Procedencia', " 
		SQL=SQL & "fra.paiori02  as 'Origen', " 
		SQL=SQL & "'"&Agente&"' as 'Agente', " 
		SQL=SQL & "'"&Aduana&"' as 'Aduana', " 
		SQL=SQL & "group_concat( distinct replace(replace(replace(a.desc05,'\n',''),'\r',''),'\a','') )as 'Producto', " 
		SQL=SQL & "fra.fraarn02  as 'Fraccion', " 
		SQL=SQL & "group_concat( distinct f.terfac39)  as 'Incoterm', " 
		SQL=SQL & "fra.cancom02 as 'Cantidad', " 
		SQL=SQL & "vmerme02/fra.cancom02  as 'Precio Proveedor', " 
		SQL=SQL & "vmerme02 MontoTotal, " 
		SQL=SQL & "cast(f.monfac39 as char)  as 'Moneda', " 
		SQL=SQL & "group_concat(distinct cast(f.numfac39 as char), ' ') as 'Factura Proveedor', " 
		SQL=SQL & "cast(group_concat(distinct f.fecfac39) as char) as 'FechaProveedor', " 
		SQL=SQL & "n.nom01  as 'Linea Flete Inter', " 
		SQL=SQL & "i.fletes01 as 'Monto Flete Inter', " 
		SQL=SQL & "''  as 'Moneda3', " 
		SQL=SQL & "'SI' as 'Fondeo', " 
		SQL=SQL & "if(i.cveped01<>'R1',(select ifnull(sum(import36),0) as campo from "& oficina &"_extranet.sscont36 as cf1  where cf1.cveimp36 in ('1', '3','15')   and refcia36 = i.refcia01),(select ifnull(sum(import33),0) as campo from "& oficina &"_extranet.sscont33 as cf1  where cf1.cveimp33 in ('1', '3','15')   and refcia33 = i.refcia01)) as 'Monto Fondeo', " 
		SQL=SQL & "if ((select sum(if(ip.cveide12 in('AL','TL'),1,0)) from "&oficina&"_extranet.ssfrac02 as f3 " 
		SQL=SQL & "left join "&oficina&"_extranet.ssipar12 as ip on ip.refcia12=f3.refcia02 and ip.patent12 =f3.patent02 and ip.adusec12=f3.adusec02 and ip.ordfra12=f3.ordfra02 " 
		SQL=SQL & "where f3.refcia02=i.refcia01 and f3.adusec02 =i.adusec01 and f3.patent02=i.patent01 group by f3.refcia02)>0,'SI','NO')	TLC, " 
		SQL=SQL & "IFNULL( fra.i_adv102,0)  AS 'IGI',  " 
		SQL=SQL & "IFNULL((fra.dtafpp02),0)  AS 'DTA',  " 
		SQL=SQL & "IFNULL((SELECT SUM(prv.import36) FROM "&oficina&"_extranet.sscont36 AS prv WHERE prv.refcia36 = i.refcia01 and prv.patent36 =i.patent01 and prv.adusec36 =i.adusec01 AND prv.cveimp36 = 15 GROUP BY prv.refcia36 ),0)  AS 'PRV', " 
		SQL=SQL & "iFNULL((SELECT SUM(e31.chon31) FROM "&oficina&"_extranet.d31refer AS d31 INNER JOIN "&oficina&"_extranet.e31cgast AS e31 ON e31.cgas31 = d31.cgas31 AND e31.esta31 = 'I' WHERE d31.refe31 = i.refcia01), 0)  as 'Honorarios 2', " 
		SQL=SQL & "IFNULL((SELECT round(SUM(d21.mont21 /(((e21.piva21/100)+1))* IF(e21.deha21 = 'C', -1, 1)),2) FROM "&oficina&"_extranet.d21paghe AS d21 LEFT JOIN "&oficina&"_extranet.e21paghe AS e21 ON e21.foli21 = d21.foli21 AND YEAR(e21.fech21) = YEAR(d21.fech21) AND e21.tmov21 = d21.tmov21 LEFT JOIN "&oficina&"_extranet.c21paghe AS c21 ON c21.clav21 = e21.conc21 WHERE d21.refe21 = i.refcia01 AND c21.desc21 LIKE '%MANIOBR%' GROUP BY d21.refe21),0) AS 'Maniobras',  " 
		SQL=SQL & "IFNULL((SELECT round(SUM(d21.mont21 /(((e21.piva21/100)+1))* IF(e21.deha21 = 'C', -1, 1)),2)  FROM "&oficina&"_extranet.d21paghe AS d21 LEFT JOIN "&oficina&"_extranet.e21paghe AS e21 ON e21.foli21 = d21.foli21 AND YEAR(e21.fech21) = YEAR(d21.fech21) AND e21.tmov21 = d21.tmov21 LEFT JOIN "&oficina&"_extranet.c21paghe AS c21 ON c21.clav21 = e21.conc21 WHERE d21.refe21 = i.refcia01 AND (c21.desc21 LIKE '%DESCON%' or c21.desc21 like'%REVALIDACION%') GROUP BY d21.refe21),0) AS 'Desconsolidacion', " 
		SQL=SQL & "IFNULL((SELECT round(SUM(d21.mont21 /(((e21.piva21/100)+1))* IF(e21.deha21 = 'C', -1, 1)),2)  FROM "&oficina&"_extranet.d21paghe AS d21 LEFT JOIN "&oficina&"_extranet.e21paghe AS e21 ON e21.foli21 = d21.foli21 AND YEAR(e21.fech21) = YEAR(d21.fech21) AND e21.tmov21 = d21.tmov21 LEFT JOIN "&oficina&"_extranet.c21paghe AS c21 ON c21.clav21 = e21.conc21 WHERE d21.refe21 = i.refcia01 AND (c21.desc21 LIKE '%MONTACARGA%' or c21.desc21 like'%PREVIO%') GROUP BY d21.refe21),0) as 'Serv. Compl.', " 
		SQL=SQL & "IFNULL(fra.i_iva102,0)  AS 'IVA',  " 
		SQL=SQL & "c.fdsp01 as  'Fecha entrega TDM', " 
		SQL=SQL & "'' as  'Linea Flete Nac.', " 
		SQL=SQL & "'' as  'Monto Flete Nac.', " 
		SQL=SQL & "'MONEDA NACIONAL' as  'Moneda2', " 
		SQL=SQL & "'' as  'Factura Flete Nac.', " 
		SQL=SQL & "ifnull(group_concat(distinct e31.cgas31),'') as  'No. Cta. de Gastos', " 
		SQL=SQL & "e31.frec31 as  'Fecha entrega TDM CG', " 
		SQL=SQL & "'' as  'Fecha entrega Contabilidad', " 
		SQL=SQL & "'' as  'Usuario', " 
		SQL=SQL & "'' as  'Division' " 
		SQL=SQL & "from 	"&oficina&"_extranet.ssdag"&movimiento&"01 as i " 
		SQL=SQL & "left join "&oficina&"_extranet.c01refer as c on i.refcia01=c.refe01 " 
		SQL=SQL & "left join "&oficina&"_extranet.c06barco as b on c.cbuq01 =b.clav06  " 
		SQL=SQL & "left join "&oficina&"_extranet.c55navie as n on b.navi06 =n.cve01  " 
		SQL=SQL & "left join "&oficina&"_extranet.ssfrac02 as fra on i.refcia01 =fra.refcia02 and i.patent01=fra.patent02 and i.adusec01 =fra.adusec02  " 
		SQL=SQL & "left join "&oficina&"_extranet.d05artic as a on fra.refcia02 = a.refe05 and fra.ordfra02 =a.agru05  " 
		SQL=SQL & "left join "&oficina&"_extranet.ssfact39 as f on i.refcia01 =f.refcia39 and i.patent01 =f.patent39 and i.adusec01 =f.adusec39 and a.fact05=f.numfac39 " 
		SQL=SQL & "left join "&oficina&"_extranet.ssguia04 as g on i.refcia01=g.refcia04 and i.patent01=g.patent04 and i.adusec01=g.adusec04    " 
		SQL=SQL & "left join "&oficina&"_extranet.d31refer as d31 on d31.refe31=i.refcia01 "
		SQL=SQL & "left join "&oficina&"_extranet.e31cgast as e31 on e31.cgas31=d31.cgas31 and e31.esta31='I' "
		SQL=SQL & "where 	i.rfccli01 ='"&strFiltroCliente&"' and i.firmae01 is not null and i.firmae01<>''  and i.fecpag01 between '"&DateI&"' and '"&DateF&"' " 
		SQL=SQL & "group by i.refcia01, fra.ordfra02,fra.fraarn02 "& chr(13) & chr(10)
	subSQL=SQL
	
	'response.write(subSQL)
	'response.end()
end function 

function retornaMontoAnticipo(referencia,campo,oficina,opcion)
dim c,valor
 c=chr(34)
 valor=""
 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	elseif(ucase(oficina)="PAN") then 
		oficina="DAI"
	end if
 
sqlAct= "SELECT refe11, " &_
				"DATE_FORMAT(MAX(fech11), '%d-%m-%Y') AS 'fecha', " &_
				"conc11, " &_
				"SUM(IF(conc11 = 'CAN', mont11*-1, mont11)) AS 'campo', " &_
				"if (SUM(IF(conc11 = 'CAN', mont11*-1, mont11))>0,'SI','NO') CountAnt "&_
				"FROM " & oficina & "_extranet.d11movim " &_
				"WHERE (conc11 = 'ANT' OR conc11 = 'CAN') AND refe11 = '" & referencia & "' " &_
				"GROUP BY refe11 "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields(opcion).value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields(opcion).value
   act2.movenext()
 wend
   retornaMontoAnticipo = valor
 else
 if valor="" and opcion="CountAnt" then
	valor="NO"
 end if
  retornaMontoAnticipo =valor
 end if
end function

function RetornaFleteNacional(Referencia,Opcion)
dim valor,ofi
Valor=0
ofi=mid(Referencia,1,3)
if ofi="ALC" then 
	ofi="LZR"
elseif ofi="PAN" then
	ofi="DAI"
end if

	sqlAct="select r.refe31 as Ref, r.cgas31,ep.conc21,round(sum((dp.mont21*if(ep.deha21 = 'C',-1,1)) )/if(ep.piva21<>0,round(ep.piva21/100,2)+1,1),2) as ImporteSIVA, cp.desc21 ,ep.bene21 , b.nomb20 Beneficiario, dp.facpro21 "
	sqlAct=sqlAct &"from "&ofi&"_extranet.d31refer as r "
	sqlAct=sqlAct &" inner join "&ofi&"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 and cta.esta31<>'C' "
	sqlAct=sqlAct &" inner join "&ofi&"_extranet.d21paghe as dp on dp.refe21 = r.refe31 and dp.cgas21 = r.cgas31 "
	sqlAct=sqlAct &" inner join "&ofi&"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 "
	sqlAct=sqlAct &" inner join  "&ofi&"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 "
	sqlAct=sqlAct &" left join "&ofi&"_extranet.c20benef as b on b.clav20=ep.bene21 "
	sqlAct=sqlAct &" where  cta.esta31 <> 'C'  and ep.conc21 in (if(mid(r.refe31,1,3)in('dai','tol'),7,if(mid(r.refe31,1,3) in('rku','lzr','ceg'),15,5))) "
	sqlAct=sqlAct &" and r.refe31 ='"&Referencia&"' group by Ref "

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&ofi&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
	valor = act2.fields(Opcion).value
   RetornaFleteNacional = valor
 else
	if Opcion="Beneficiario" then 
		RetornaFleteNacional =""
	else 
		RetornaFleteNacional=valor
	end if
 end if
End Function

function RetornaHonorarioBit(Ref,Cta,mov,opc)
	dim ofi, valor, ActSQL,hono,servcom,cveofi
	ofi=mid(Ref,1,3)
	if ofi="ALC" then 
		ofi="LZR"
	elseif ofi="PAN" then 
		ofi="DAI"
	end if
	valor=0
	select case ofi
		case "RKU"
			cveofi="0001"
			hono="4100000100000000"
			servcom="4100000200080000"
		case "DAI"
			cveofi="0005"
			hono="5501000100000000"
			servcom="5501000200000000"
		case "SAP"
			cveofi="0004"
			hono="4100000100000000"
			servcom="4100000200000000"
		case "TOL"
			cveofi="0010"
			hono="4100001000010000"
			servcom="4100001000020000"
		case "CEG"
			cveofi="0003"
			hono="4100000100010000"
			servcom="4100000100020000"
		case "LZR"
			cveofi="0009"
			hono="4100000100000000"
			servcom="4100000200000000"
	end select
	
	if opc="H" then 
		opc=hono
	elseif opc="S" then
		opc=servcom
	end if
	ActSQL=" SELECT  SUM(IF(A.ASIE11 = '"&opc&"', IF(A.CONC11 REGEXP 'FA1|SCA|DEV|CAR|FA2' , A.MONT11, IF(A.CONC11 REGEXP 'LIQ|CF1|SCR|ABO|BOH|CF2' , A.MONT11*-1,0)), 0)) as Monto "&_
		   "	FROM "&ofi&"_extranet.ssdag"&mov&"01 as ix "&_
			"	inner join "&ofi&"_extranet.ssclie18 as cli on cli.cvecli18 = ix.cvecli01 "&_
			"	inner join "&ofi&"_extranet.d31refer as rx on rx.refe31 = ix.refcia01 "&_
			"	inner join "&ofi&"_extranet.e31cgast as cta on cta.cgas31 = rx.cgas31 "&_
			"	inner join  "&ofi&"_extranet.D11MOVIM AS A  on  A.cgas11 = rx.cgas31 "&_
			"	inner join "&ofi&"_extranet.E11Movim AS B ON A.Foli11 = B.Foli11 "&_
			"	WHERE (A.ASIE11 = '"&servcom&"' or A.ASIE11 = '"&hono&"') and "&_
			"	(cli.facofna = '"&cveofi&"'  or cli.facofna = '') 	"&_
			"	AND A.CURE11 <> 'R'  and B.cont11 <> 'C'   and rx.refe31 = '"&Ref&"' "&_
			"	GROUP BY rx.refe31 "
	
	Set act2= Server.CreateObject("ADODB.Recordset")
	
	conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&ofi&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	
	act2.ActiveConnection = conn12
	act2.Source = ActSQL
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	if not(act2.eof) then
		valor = act2.fields("Monto").value
		RetornaHonorarioBit = valor
	else
		RetornaHonorarioBit=valor
	end if

end function

Function RetornaServCom(Ref,mov)
	dim ofi ,mon,prev,valor
	valor=0
	ofi=mid(Ref,1,3)
	if ofi="ALC" then
		ofi="LZR"
	elseif ofi="PAN" then
		ofi="DAI"
	end if
	select case ofi
		case "RKU"
			mon="166"
			prev="111"
		case "SAP"
			mon="195"
			prev="175"
		case "LZR"
			mon="166,171"
			prev="111,401,325,183,390"
		case "TOL"
			mon="11"
			prev="102,12,179"
		case "DAI"
			mon="11"
			prev="12,102"
		case "CEG"
			mon=""
			prev="305,306"
	end select
	
	sqlAct="select i.refcia01 as Ref,cta.fech31, r.cgas31,ep.conc21,ep.piva21,ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as Importe, cp.desc21,cta.csce31 as TFlat , ep.tpag21,ep.deha21 "&_
			"	from "&ofi&"_extranet.ssdag"&mov&"01 as i  "&_
			"	inner join "&ofi&"_extranet.d31refer as r on r.refe31 = i.refcia01  "&_
			"	inner join "&ofi&"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 "&_
			"	inner join "&ofi&"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 "&_
			"	inner join "&ofi&"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  and ep.tmov21 =dp.tmov21  "&_
			"	inner join  "&ofi&"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 "&_
			"	where  i.firmae01 <> ''  and cta.esta31 <> 'C' "&_
			"	and i.refcia01='"&Ref&"' and ep.conc21 in("&mon&","&prev&") "&_
			"	group by i.refcia01"

	Set act2= Server.CreateObject("ADODB.Recordset")
	
	conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&ofi&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	
	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	if not(act2.eof) then
		valor = act2.fields("Importe").value
		RetornaServCom = valor
	else
		RetornaServCom=valor
	end if

			
end Function

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
		Cargas=false
	else 
		Cargas=true
	end if
	act2.Close()
end function 
%>
<HTML>
	<HEAD>
		<TITLE>::.... TAKASAGO.... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>