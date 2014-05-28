<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%Server.ScriptTimeout=15000


strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

	fi=trim(request.form("fi"))
	ff=trim(request.form("ff"))
	Vrfc=Request.Form("rfcCliente")
	Vckcve=Request.Form("ckcve")
	Vclave=Request.Form("txtCliente")

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

	jnxadu=Session("GAduana")

	select case jnxadu
		case "VER"
			strOficina="rku"
		case "MEX"
			strOficina="dai"
		case "MAN"
			strOficina="sap"
		case "LZR"
			strOficina="lzr"
	end select
	if mov = "i" then
		movi = "IMPORTACION ::"
		query = GeneraSQL(mov)
	else
		movi = "EXPORTACION ::"
		query = GeneraSQL(mov)
	end if
	
	nocolumns = 25
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)
	IF RSops.BOF = True And RSops.EOF = True Then
		
		Response.Write("No hay datos para esas condiciones")
	Else
	
		if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;"
			Response.ContentType = "application/vnd.ms-excel"
		End If
		info = 	"<table  width = ""2929""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<strong>" &_
									"<font color=""#000066"" size=""4"" face=""Arial, Helvetica, sans-serif"">" &_
										"<td colspan=""" & nocolumns & """>" &_
											"<p align=""left""><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">" &_
												":: OPERACIONES DE " & movi &_
											"</font></p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p><font color=""#000066"" size=""2"" face=""Arial, Helvetica, sans-serif"">" &_
												"Del " & fi & " Al " & ff &_
											"</font></p>" &_
											"<p>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
		header = 			"<tr class = ""boton"">" &_
								celdahead("Compañia") &_
								celdahead("Referencia") &_
								celdahead("Contenedores") &_
								celdahead("Cliete") &_
								celdahead("Trasporte/Buque") &_
								celdahead("Pedimento") &_
								celdahead("Facturas") &_
								celdahead("Fecha Fact") &_
								celdahead("Proveedor") &_
								celdahead("Valor Factura") &_
								celdahead("Moneda Factura") &_
								celdahead("Ref. Cliente") &_
								celdahead("Mercancia") &_
								celdahead("Recepcion de documentos") &_
								celdahead("Estimated time arrival") &_
								celdahead("Revalidacion") &_
								celdahead("Solicitud de anticipo") &_
								celdahead("Deposito de anticipo") &_
								celdahead("Previo") &_
								celdahead("Pago de pedimento") &_
								celdahead("Programado para embarque") &_
								celdahead("Despacho") &_
								celdahead("Vacio de contenedores") &_
								celdahead("Emision de cuenta de gastos") &_
								celdahead("Comentarios ¿Cual es el pendiente por resolver para realizar el despacho?")								
				header = header &	"</tr>"
		
		Do Until RSops.EOF
 				datos = datos&"<tr> "&_
								celdadatos(RSops.Fields.Item("Company").Value)&_
								celdadatos(RSops.Fields.Item("Referencia").Value)&_
								celdadatos(RSops.Fields.Item("Contenedores").Value)&_ 
								celdadatos(RSops.Fields.Item("Cliente").Value) &_
								celdadatos(RSops.Fields.Item("Transporte/Buque").Value) &_
								celdadatos(RSops.Fields.Item("No.Pedimento").Value) &_
								celdadatos(RSops.Fields.Item("Factura").Value) &_
								celdadatos(RSops.Fields.Item("Fec.Fact").Value) &_
								celdadatos(RSops.Fields.Item("Provedor").Value) &_
								celdadatos(RSops.Fields.Item("VALOR_FACTURA").Value) &_
								celdadatos(RSops.Fields.Item("MonedaFac").Value) &_
								celdadatos(RSops.Fields.Item("Ref.Cliente").Value) &_
								celdadatos(RSops.Fields.Item("Desc.Prod.").Value) &_
								celdadatos(RSops.Fields.Item("Rec.Documentos").Value) &_
								celdadatos(RSops.Fields.Item("F.Est.Arribo").Value) &_
								celdadatos(RSops.Fields.Item("Revalidacion").Value) &_
								celdadatos(RSops.Fields.Item("F.Solicitud.Ant").Value) &_
								celdadatos(RSops.Fields.Item("Deposito.Ant").Value) &_
								celdadatos(RSops.Fields.Item("Previo").Value) &_
								celdadatos(RSops.Fields.Item("Pago.Ped.").Value) &_
								celdadatos(RSops.Fields.Item("Prog.embarque").Value) &_
								celdadatos(RSops.Fields.Item("F.Desp").Value) &_
								celdadatos(RSops.Fields.Item("Vacio.cont").Value) &_
								celdadatos(RSops.Fields.Item("Emi.CG").Value) &_
								celdadatos(GeneraObser(RSops.Fields.Item("Referencia").Value,strOficina,mov)) 
				datos = datos &	"</tr>"
							
			Rsops.MoveNext()
		Loop
	
	Response.Write(info & header & datos & "</table><br>")
	Response.End()
	
'html=info&header& "</table><br>"
	html = info & header & datos & "</table><br>"
	End If
end if

function celdahead(texto)
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

function celdadatos(texto)
'On error resume next
	If IsNull(texto) = True Or texto = "" Then
		texto = "&nbsp;"
	End If
	cell = 	"<td align=""center"">" &_
				"<font size=""1"" face=""Arial"">" &_
					texto &_
				"</font>" &_
			"</td>"
	celdadatos = cell
end function

function GeneraObser(ref,ofi,topera)
sql=""
	sql="SELECT group_concat( concat_ws(' ', etx.f_fecha, etx.m_observ)) as Observaciones " 
	sql=sql&"FROM "&ofi&"_extranet.ssdag"&topera&"01 AS i " 
	sql=sql&	"LEFT JOIN "&ofi&"_extranet.c01refer AS c ON i.refcia01 = c.refe01 " 
		sql=sql&"LEFT JOIN "&ofi&"_status.etxpd as etx on etx.c_referencia = i.refcia01 and etx.clavec <> 0 "
		sql=sql&"LEFT JOIN "&ofi&"_status.c01caus as cau on cau.c01clavec = etx.clavec  " 
		IF Vckcve=0 then
			sql=sql&" WHERE i.rfccli01 in('"&Vcrfc&"') and etx.n_etapa=9 and i.refcia01 ='"&ref&"'"
		ELSEIF Vckcve=1 and Vclave="Todos" THEN
			sql=sql&"WHERE etx.n_etapa=9 and i.refcia01 ='"&ref&"'"
		ELSE
			sql=sql&"WHERE i.cvecli01 in("&Vclave&") and etx.n_etapa=9 and i.refcia01 ='"&ref&"'"
		END IF
		
Set act2= Server.CreateObject("ADODB.Recordset")
conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&ofi&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sql
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
  if not(act2.eof) then
 valor = act2.fields("Observaciones").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("Observaciones").value
   act2.movenext()
 wend
  GeneraObser = valor
 else
  GeneraObser =valor
 end if

	
end function

function GeneraSQL(op)
	SQL = ""
	jnxadu=Session("GAduana")

	select case jnxadu
		case "VER"
			strOficina="rku"
		case "MEX"
			strOficina="dai"
		case "MAN"
			strOficina="sap"
		case "LZR"
			strOficina="lzr"
	end select
		
			SQL = SQL &	"select cc.nomcli18 as 'Company'," & chr(13) & chr(10)
			SQL = SQL & "i.refcia01 as 'Referencia', " & chr(13) & chr(10)
			SQL = SQL & "ifnull((if(dco.clas01 <> 'CON',CONCAT_ws(',',left(cc.nomcli18,14),(select  concat_WS(' ',sum(sd.cant01 ),sd.clas01)  " & chr(13) & chr(10)
			SQL = SQL & "	from "&strOficina&"_extranet.ssdag"&op&"01 as si " & chr(13) & chr(10)
			SQL = SQL & "		left join "&strOficina&"_extranet.d01conte as sd on sd.refe01 =si.refcia01 " & chr(13) & chr(10)
			SQL = SQL & "			where si.refcia01  =i.refcia01)),group_concat(distinct cont.numcon40))),i.totbul01) as Contenedores, " & chr(13) & chr(10)
			SQL = SQL & "cc.nomcli18 as 'Cliente', " & chr(13) & chr(10)
			SQL = SQL & "i.nombar01 AS 'Transporte/Buque', " & chr(13) & chr(10)
			SQL = SQL & "i.numped01 as 'No.Pedimento', " & chr(13) & chr(10)
			SQL = SQL & "f.numfac39 as 'Factura', " & chr(13) & chr(10)
			SQL = SQL & "f.fecfac39 as 'Fec.Fact', " & chr(13) & chr(10)
			SQL = SQL & "p22.nompro22   as 'Provedor', " & chr(13) & chr(10)
			SQL = SQL & "round((f.valmex39 * f.facmon39 * i.tipcam01),0) AS 'VALOR_FACTURA', " & chr(13) & chr(10)
			SQL = SQL & "f.monfac39 as 'MonedaFac', " & chr(13) & chr(10)
			SQL = SQL & "r.rcli01 as 'Ref.Cliente', " & chr(13) & chr(10)
			SQL = SQL & "ar.desc05 as 'Desc.Prod.'," & chr(13) & chr(10)
			SQL = SQL & "r.fdoc01 as 'Rec.Documentos', " & chr(13) & chr(10)
			SQL = SQL & "r.feta01 as 'F.Est.Arribo', " & chr(13) & chr(10)
			SQL = SQL & "r.frev01 as 'Revalidacion', " & chr(13) & chr(10)
			SQL = SQL & "r.fcot01 as 'F.Solicitud.Ant', " & chr(13) & chr(10)
			SQL = SQL & "d11.fech11 as 'Deposito.Ant', " & chr(13) & chr(10)
			SQL = SQL & "r.fpre01 as 'Previo', " & chr(13) & chr(10)
			SQL = SQL & "i.fecpag01 as 'Pago.Ped.', " & chr(13) & chr(10)
			SQL = SQL & "r.fpro01  as 'Prog.embarque', " & chr(13) & chr(10)
			SQL = SQL & "r.fdsp01 as 'F.Desp', " & chr(13) & chr(10)
			SQL = SQL & "dco.fcarta01 as 'Vacio.cont'," & chr(13) & chr(10)
			SQL = SQL & "cta.fech31  as 'Emi.CG' " & chr(13) & chr(10)
			SQL = SQL & "  from "&strOficina&"_extranet.ssdag"&op&"01 as i  " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.ssclie18 as cc on cc.rfccli18 = i.rfccli01   " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.ssprov22 as p22 on p22.cvepro22 =i.cvepro01 " & chr(13) & chr(10)
			SQL = SQL & "left join "&strOficina&"_extranet.c01refer as r on r.refe01 = i.refcia01  " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & chr(13) & chr(10)
			SQL = SQL & " LEFT JOIN "&strOficina&"_extranet.d11movim AS d11 ON d11.refe11 = i.refcia01 AND d11.conc11 = 'ANT'" & chr(13) & chr(10)
			SQL = SQL & "left join "&strOficina&"_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01  " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.sscont40 as cont on cont.refcia40=i.refcia01 " & chr(13) & chr(10)
			SQL = SQL & "left join "&strOficina&"_extranet.d01conte as dco on dco.refe01 =i.refcia01  " & chr(13) & chr(10)
			IF Vckcve=0 then
				SQL = SQL & " where cc.rfccli18 in('"&Vrfc&"') and i.cvecli01 =cc.cvecli18 and i.fecpag01>='"&DateI&"' and i.fecpag01 <='"&DateF&"' and i.firmae01 <>''" & chr(13) & chr(10)
			ELSEIF Vckcve=1 and Vclave="Todos" THEN
				SQL = SQL & "where i.fecpag01>='"&DateI&"' and i.fecpag01 <='"&DateF&"' and i.firmae01 <>''" &chr(13) & chr(10)
			ELSE
				SQL = SQL & "where cc.cvecli18  in("&Vclave&") and i.cvecli01 =cc.cvecli18 and i.fecpag01>='"&DateI&"' and i.fecpag01 <='"&DateF&"' and i.firmae01 <>''" &chr(13) & chr(10)
			END IF
			SQL = SQL & "group by i.refcia01,i.fecpag01" & chr(13) & chr(10)
	GeneraSQL = SQL
end function


%>
<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE OPERACIONES COCA-COLA Y JDV .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>