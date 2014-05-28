<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000


strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
strOficina=Request.Form("OficinaG")
strCortaR=Request.Form("cortaR")

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
strFiltroCliente = request.Form("txtCliente")
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
	'response.write(query&strOficina)
	'response.end()
	nocolumns = 16
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)
	
	IF RSops.BOF = True And RSops.EOF = True Then
		
		Response.Write("No hay datos para esas condiciones")
	Else
		
		if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;filename=Rep_Digitalizacion_"&DiaI&"-"&Mesi&"_"&DiaF&"-"&MesF&".xls"
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
											"<p><font color=""#000000"" size=""4"" face=""Arial, Helvetica, sans-serif"">: : D I G I T A L I Z A C I O N : :</font></p>" &_
											"<p><font color=""#000000"" size=""2"" face=""Arial, Helvetica, sans-serif"">"&movi&" DEL "&DateI&" al "&DateF&"</font>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
		header = 			"<tr class = ""boton"">" &_
								celdahead("Operacion","#81BEF7") &_
								celdahead("Referencia","#81BEF7") &_
								celdahead("Cuentas de Gastos","#00FF00") &_
								celdahead("Clave Pedimento","#81BEF7")&_
								celdahead("Cliente","#81BEF7") &_
								celdahead("Fecha de Pago Pedimento","#81BEF7") &_
								celdahead("Fecha CG Reciente","#00FF00") &_
								celdahead("Fecha Recepcion CG Reciente","#00FF00") &_
								celdahead("Registro digitalizacion Reciente","#81BEF7") &_
								celdahead("Doc. Digitalizados en General","#81BEF7") &_
								celdahead("CG Digitalizada","#FF4000") &_
								celdahead("Pedimento digitalizado","#FF4000") &_
								celdahead("Cant. Comprobantes PH Digitalizados","#FF4000") &_
								celdahead("Cant. Facturas Comerciales Digitalizadas","#FF4000") &_
								celdahead("Cant. PH de la Referencia","#FACC2E") &_
								celdahead("Auditoria digitalizacion","#FACC2E")
						
				header = header &	"</tr>"
				dim repCG, tieneCGD, snco,snco2, catidadph
				
			Do Until RSops.EOF
				repCG=RSops.Fields.Item("Fecha_RecepCG_reciente").Value 
				tieneCGD=RSops.Fields.Item("DocumentosCG").Value
				cantidadph=Cint(RSops.Fields.Item("Cant_ComprobantesPHDigitalizados").Value)
				
					if repCG<>"" and tieneCGD="" then
						'Filtro: Si la fecha de recepcion de cuenta de gastos se encuentra capturada y no se encuentra digitalizada la cuenta de gastos entonces se marcara la fila 
						snco="#F2F5A9"
						if cantidadph<Cint(CantPH(RSops.Fields.Item("refcia01").Value,RSops.Fields.Item("CuentasGastos").Value,mid(RSops.Fields.Item("refcia01").Value,1,3)))  then 
							snco2="#FFFF00"
						else
						snco2=snco
						end if
					else
						snco="#FFFFFF"
						snco2=snco
					end if
							datos = datos & "<tr> " &_
							celdadatos(RSops.Fields.Item("tipo").Value,snco) &_
							celdadatos(RSops.Fields.Item("refcia01").Value,snco) &_
							celdadatos(RSops.Fields.Item("CuentasGastos").Value,snco) &_
							celdadatos(RSops.Fields.Item("cveped01").Value,snco) &_ 
							celdadatos(RSops.Fields.Item("nomcli01").Value,snco) &_
							celdadatos(RSops.Fields.Item("fecpag01").Value,snco) &_
							celdadatos(RSops.Fields.Item("FechaCG_reciente").Value,snco) &_
							celdadatos(RSops.Fields.Item("Fecha_RecepCG_reciente").Value,snco) &_
							celdadatos(RSops.Fields.Item("f_alta").Value,snco) &_
							celdadatos(RSops.Fields.Item("DocumentosenGeneral").Value,snco) &_
							celdadatos(RSops.Fields.Item("DocumentosCG").Value,snco) &_
							celdadatos(RSops.Fields.Item("DocumentosCG2").Value,snco) &_
							celdadatos(RSops.Fields.Item("Cant_ComprobantesPHDigitalizados").Value,snco) &_
							celdadatos(RSops.Fields.Item("Cant_FacturasComercialesDigitalizadas").Value,snco)&_
							celdadatos(CantPH(RSops.Fields.Item("refcia01").Value,RSops.Fields.Item("CuentasGastos").Value,mid(RSops.Fields.Item("refcia01").Value,1,3)),snco2) &_
							celdadatos(RSops.Fields.Item("Auditoria_digitalizacion").Value,snco)
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
						"<font color=""#000000"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
							texto &_
						"</font>" &_
					"</strong>" &_
				"</center>" &_
			"</td>"
	celdahead = cell
end function

function celdadatos(texto,pcolor)'Celda de datos de la tabla
'On error resume next
'response.write(texto & "   ..")
texto = texto & ""
'response.end()
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
	SQL=""
		if strCortaR="Todo" then
			strCortaR="left"
		else
			strCortaR="inner"
		end if
	SQL=SQL & "select '"&operacion&"' as tipo,ax.refcia01,ax.cveped01,ax.nomcli01,ax.fecpag01,max(ax.fech31) as FechaCG_reciente, max(ax.frec31) as Fecha_RecepCG_reciente, group_concat(distinct ax.cgas31) as CuentasGastos, doc.f_alta, "& chr(13) & chr(10)
			SQL=SQL & "count(distinct doc.d_image) as DocumentosenGeneral, "& chr(13) & chr(10)
			SQL=SQL & "if(sum(if(doc.d_documen ='CGT',1,0)) >0,'Si tiene CG Digitalizada','') as DocumentosCG, "& chr(13) & chr(10)
			SQL=SQL & "if(sum(if(doc.d_documen ='PDI',1,0)) >0,'Si tiene Pedimento Digitalizado','') as DocumentosCG2, "& chr(13) & chr(10)
			SQL=SQL & "ifnull(round( (sum(if(doc.d_documen ='CPG',1,0))/count(distinct ax.cgas31)),0),0) as Cant_ComprobantesPHDigitalizados, "& chr(13) & chr(10)
			SQL=SQL & "ifnull(round( (sum(if(doc.d_documen ='FAC',1,0))/count(distinct ax.cgas31)),0),0) as Cant_FacturasComercialesDigitalizadas, "& chr(13) & chr(10)
			SQL=SQL & "round( (sum(if(doc.m_pdf ='E',1,0))/count(distinct ax.cgas31)),0) as Cant_JPG_con_error, "& chr(13) & chr(10)
			SQL = SQL & " (SELECT cast( (concat('[',ds.f_emision,' ',ds.h_emision,'] ',et.d_nombre,'. ',ds.t_comentario )) as char) as Auditoria_digitalizacion "& chr(13) & chr(10)
			SQL = SQL & " FROM usuarios.cat01_issues_subjects as cs "& chr(13) & chr(10)
			SQL = SQL & " inner join usuarios.det01_issues_comments as ds on ds.i_cve_issue = cs.i_cve_issue "& chr(13) & chr(10)
			SQL = SQL & " left join rku_status.etaps as et on et.n_etapa = cs.i_etapa and et.l_auto = 'D' "& chr(13) & chr(10)
			SQL = SQL & " where cs.c_referencia =doc.c_referencia "& chr(13) & chr(10)
			SQL = SQL & " limit 1) as Auditoria_digitalizacion " & chr(13) & chr(10)
			SQL=SQL & " from ( (SELECT i.nomcli01,i.refcia01,i.cveped01,i.fecpag01,e31.fech31,e31.frec31,e31.cgas31"& chr(13) & chr(10)
			SQL=SQL & "		FROM "&oficina&"_extranet.ssdag"&movimiento&"01 as i "& chr(13) & chr(10)
			SQL=SQL & "			"&strCortaR&" join "&oficina&"_extranet.d31refer as d31 on d31.refe31 = i.refcia01 "& chr(13) & chr(10)
			SQL=SQL & "			"&strCortaR&" join "&oficina&"_extranet.e31cgast as e31 on e31.cgas31 = d31.cgas31 "& chr(13) & chr(10)
			SQL=SQL & "			where "& chr(13) & chr(10)
			if bclientes="T" then 
					SQL=SQL &" i.rfccli01  in('"&Vrfc&"') and  "& chr(13) & chr(10)
			end if
			
			SQL=SQL & "					i.firmae01 <> '' "& chr(13) & chr(10)
			if strCortaR="inner" then
				SQL=SQL & "			and e31.esta31 = 'I' "& chr(13) & chr(10)
			end if
			SQL=SQL & "			and i.fecpag01>='"&DateI&"' and i.fecpag01<='"&DateF&"')"& chr(13) & chr(10)
			SQL=SQL & "		) as ax "& chr(13) & chr(10)
			SQL=SQL & "		left join digitalizacion.doc_doctos_digitales as doc on doc.c_referencia = ax.refcia01 and ucase(doc.m_delete) <> 'X'"& chr(13) & chr(10)
			SQL=SQL & "		group by ax.refcia01 "& chr(13) & chr(10)
	subSQL=SQL
	
	'response.write(subSQL)
	'response.end()
end function 

function CantPH(referencia,CGT,offi) 'Funcion para calcular la cantidad de pagos hechos de la referencia y cuenta de gastos especificada
	if offi="ALC" then 
		offi="lzr"
	elseif offi="PAN" then
		offi="dai"
	end if
	
	dim Valor
		Valor=0
	SQL=""
		SQL="select count(x.Cuenta) Cantidad from (select  ifnull(sum(if(ep.piva21>=15,(dp.mont21*if(ep.deha21= 'C',-1,1))/((ep.piva21/100)+1),(dp.mont21*if(ep.deha21 = 'C',-1,1)))),0) as  cuenta "
		SQL=SQL &"  from  "&offi&"_extranet.d31refer as r "
		SQL=SQL &"      inner join "&offi&"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 and cta.esta31 <> 'C' "
		SQL=SQL &"      inner join "&offi&"_extranet.d21paghe as dp on dp.refe21 = '"&referencia&"' and dp.cgas21 = r.cgas31 "
		SQL=SQL &"      inner join "&offi&"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  and ep.tmov21 = dp.tmov21 "
		SQL=SQL &"      inner join  "&offi&"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 "
		SQL=SQL &"  where   ep.conc21 <> 1 and cta.cgas31 = '"&CGT&"' and  r.refe31 ='"&referencia&"' group by cp.clav21 "
		SQL=SQL &" 			) as x "
		
		
		Set act2= Server.CreateObject("ADODB.Recordset")
		conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&offi&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		act2.ActiveConnection = conn12
		act2.Source = SQL
		act2.cursortype=0
		act2.cursorlocation=2
		act2.locktype=1
		act2.open()
		if not(act2.eof) then
			Valor =act2.fields("Cantidad").value
		end if
		act2.Close()
	CantPH=Valor
end function
%>
<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE DIGITALIZACION .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>