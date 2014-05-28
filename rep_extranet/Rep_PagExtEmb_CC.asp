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
		case "TAM"
			strOficina="ceg"
	end select
	if mov = "i" then
		movi = "IMPORTACION ::"
		query = GeneraSQL(mov)
	else
		movi = "EXPORTACION ::"
		query = GeneraSQL(mov)
	end if
	
	nocolumns = 13
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
										"<td colspan=""" & nocolumns & """></font></p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p><font color=""#000066"" size=""2"" face=""Arial, Helvetica, sans-serif""></font></p>" &_
											"<p>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
		header = 			"<tr class = ""boton"">" &_
								celdahead("Plant Name") &_
								celdahead("Period") &_
								celdahead("Referencia Grupo Zego") &_
								celdahead("Purchase Orden") &_
								celdahead("ITEM") &_
								celdahead("Material description") &_
								celdahead("Concept") &_
								celdahead("Amount Paid (FC)(SIN IVA)") &_
								celdahead("Amount Paid $USD") &_
								celdahead("Vendor") &_
								celdahead("Cause Explanations") &_
								celdahead("Actions Plans") &_
								celdahead("Owner /Action Plan")
				header = header &	"</tr>"
		Aux=""
		i=0
		nuevo=true
		item1=""
	
		
			Do Until RSops.EOF
							datos = datos & "<tr> " &_
							celdadatos(RSops.Fields.Item("Cliente").Value) &_
							celdadatos(RSops.Fields.Item("Periodo").Value) &_
							celdadatos(RSops.Fields.Item("Referencia").Value) &_ 
							celdadatos(RSops.Fields.Item("Orden de compra").Value) &_
							celdadatos(RSops.Fields.Item("Item").Value) &_
							celdadatos(RSops.Fields.Item("DescProd").Value) &_
								celdadatos("Demurrages") &_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,11,mov,strOficina,"SinIVA")) &_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,11,mov,strOficina,"Conversion")) &_
								celdadatos(RSops.Fields.Item("Vendor").Value) &_
								celdadatos(GeneraObser(RSops.Fields.Item("Referencia").Value,strOficina,mov)) &_
								celdadatos("") &_
								celdadatos("") 
							datos = datos &	"</tr>"
							
					for ii=1 to 4
								datos = datos&"<tr> " &_
								celdadatos("") &_
								celdadatos("") &_
								celdadatos("") &_ 
								celdadatos("") &_
								celdadatos("") &_
								celdadatos("") 
							select case ii
							case 1
								datos=datos & celdadatos("Storages") &_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,4,mov,strOficina,"SinIVA")) &_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,4,mov,strOficina,"Conversion"))
							case 2
								datos=datos & celdadatos("Additional Duties") &_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,180,mov,strOficina,"SinIVA")) &_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,180,mov,strOficina,"Conversion"))
							case 3
								datos=datos & celdadatos("Additional Freight Charge") &_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,"2,166,177",mov,strOficina,"SinIVA"))&_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,"2,166,177",mov,strOficina,"Conversion"))
			
							case 4
								datos=datos & celdadatos("Additional Broker Charge") &_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,45,mov,strOficina,"SinIVA")) &_
								celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").Value,45,mov,strOficina,"Conversion"))
								
							end select
							datos=datos & celdadatos("") &_
									celdadatos("") &_
									celdadatos("") &_
									celdadatos("") 
								datos = datos &	"</tr>"
								
						next 
			Rsops.MoveNext()
		Loop
		datos=datos&"</table><br><br><br>"
		datos=datos&"<table  width = ""2929""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">"
	datos=datos&"<tr>"&celdahead("")&celdahead("")&celdahead("")&celdahead("")&celdahead("")&celdahead("")&celdahead("TOTALES") &celdahead(retornaTotal("SinIva",strOficina,mov,"4,11,2,166,177,45,180")) &celdahead(retornaTotal("Conversion",strOficina,mov,"4,11,2,166,177,45,180")) &celdahead("") &celdahead("") &celdahead("") &celdahead("") &"</tr>"
	Response.Write(info & header & datos & "</table><br>")
	Response.End()
	html = info & header & datos & "</table><br>"
	End If
end if

function celdahead(texto)
	cell = "<td bgcolor = ""#FFFFFF"" width=""135"" nowrap>" &_
				"<center>" &_
					"<strong>" &_
						"<font color=""#000000"" size=""2"" face=""Arial, Helvetica, sans-serif"">" &_
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
	sql="SELECT group_concat( if(etx.m_observ<>'',concat_ws(' ', etx.f_fecha, etx.m_observ),'')) as Observaciones   " 
	sql=sql&"FROM "&ofi&"_extranet.ssdag"&topera&"01 AS i " 
	sql=sql&	"LEFT JOIN "&ofi&"_extranet.c01refer AS c ON i.refcia01 = c.refe01 " 
		sql=sql&"LEFT JOIN "&ofi&"_status.etxpd as etx on etx.c_referencia = i.refcia01 "
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
		case "TAM"
			strOficina="ceg"
	end select
		
			SQL = SQL &	"select cc.nomcli18 as 'Cliente', " & chr(13) & chr(10)
			SQL = SQL & "DATE_FORMAT(cta.fech31,'%M')  as Periodo, " & chr(13) & chr(10)
			SQL = SQL & "i.refcia01 as 'Referencia',  " & chr(13) & chr(10)
			SQL = SQL & "group_concat(distinct ar.pedi05) as 'Orden de compra'," & chr(13) & chr(10)
			SQL = SQL & "group_concat(distinct ar.item05) as 'Item', " & chr(13) & chr(10)
			SQL = SQL & "group_concat(distinct ar.desc05) as 'DescProd', " & chr(13) & chr(10)
			SQL = SQL & "f2.paiscv02 as 'Vendor' " & chr(13) & chr(10)
			SQL = SQL & "from "&strOficina&"_extranet.ssdag"&op&"01 as i " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.ssclie18 as cc on cc.rfccli18 =i.rfccli01  " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.c01refer as r on r.refe01 = i.refcia01  " & chr(13) & chr(10)
			SQL = SQL & "INNER join "&strOficina&"_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01   " & chr(13) & chr(10)
			SQL = SQL & " INNER join "&strOficina&"_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C'  " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01   " & chr(13) & chr(10)
			SQL = SQL & " left join "&strOficina&"_extranet.ssfrac02 as f2 on f2.refcia02 =i.refcia01 " & chr(13) & chr(10)
			SQL = SQL & "left join "&strOficina&"_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01      " & chr(13) & chr(10)
			IF Vckcve=0 then
				SQL = SQL & " where i.rfccli01 in('"&Vrfc&"') and i.cvecli01 =cc.cvecli18  and cta.fech31  >='"&DateI&"' and cta.fech31 <='"&DateF&"'  " & chr(13) & chr(10)
			ELSEIF Vckcve=1 and Vclave="Todos" THEN
				SQL = SQL & " where i.cvecli01 =cc.cvecli18  and cta.fech31  >='"&DateI&"' and cta.fech31 <='"&DateF&"'  " & chr(13) & chr(10)			
			ELSE
				SQL = SQL & " where i.cvecli01="&Vclave&" and i.cvecli01 =cc.cvecli18  and cta.fech31  >='"&DateI&"' and cta.fech31 <='"&DateF&"'  " & chr(13) & chr(10)
			END IF
			SQL = SQL & " group by i.refcia01 " ',ar.pedi05 ,ar.item05,f2.paiscv02 " & chr(13) & chr(10)
			
			'response.write(SQL)
			'response.end()
	GeneraSQL = SQL
	
end function

function retornaPagosHechos(referencia,conceptos,tipope,oficina,campo)
dim c,valor
 valor=0
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then
	sqlAct="select   i.refcia01 as Ref, r.cgas31,ep.conc21, ep.piva21, if (ep.piva21 =0 ,truncate(ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0),1),truncate((ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0))*100/(100+ep.piva21),2)) as SinIVA, "
		sqlAct=sqlAct & "truncate(if (ep.piva21 =0 ,truncate(ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0),1),truncate((ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0))*100/(100+ep.piva21),2))/i.tipcam01,2) as Conversion,    i.tipcam01,  cp.desc21 "
		sqlAct=sqlAct & "from "& oficina &"_extranet.ssdag"&tipope&"01 as i "
		sqlAct=sqlAct & "inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 "
        sqlAct=sqlAct & "  inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 "
        sqlAct=sqlAct & "inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  "
		sqlAct=sqlAct & "and ep.tmov21 =dp.tmov21 inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 "
		IF Vckcve=0 then
			sqlAct=sqlAct & " where i.rfccli01 in('"&Vrfc&"') and firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in("& conceptos &")  and i.refcia01 in('"& referencia &"') and cta.fech31 >='"&DateI&"'	 and cta.fech31 <='"&DateF&"' "
		ELSEIF Vckcve=1 and Vclave="Todos" THEN
			sqlAct=sqlAct & " where i.cvecli01 =cc.cvecli18  and firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in("& conceptos&")  and i.refcia01 in('"& referencia &"') and cta.fech31 >='"&DateI&"'	 and cta.fech31 <='"&DateF&"' "		
		ELSE
			sqlAct=sqlAct & " where i.cvecli01="&Vclave&" and  firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in("& conceptos &")  and i.refcia01 in('"& referencia &"') and cta.fech31 >='"&DateI&"'	 and cta.fech31 <='"&DateF&"' "
		END IF

		Set act3= Server.CreateObject("ADODB.Recordset")
		conn123 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act3.ActiveConnection = conn123
	act3.Source = sqlAct
	act3.cursortype=0
	act3.cursorlocation=2
	act3.locktype=1
	act3.open()
	valor=0
	if not(act3.eof) then
		valor = act3.fields(campo).value
		retornaPagosHechos = valor
	else
		retornaPagosHechos = valor
	end if
 
else 
 retornaPagosHechos = valor
end if

End function

function retornaTotal(campo,oficina,tipope,conceptos)
dim c,valor
 valor=0
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then
	sqlAct="select   sum(if (ep.piva21 =0 ,truncate(ifnull((dp.mont21*if(ep.deha21 = 'C',-1,1)),0),1),truncate((ifnull((dp.mont21*if(ep.deha21 = 'C',-1,1)),0))*100/(100+ep.piva21),2))) as SinIVA, "
		sqlAct=sqlAct & "sum(truncate(if (ep.piva21 =0 ,truncate(ifnull((dp.mont21*if(ep.deha21 = 'C',-1,1)),0),1),truncate((ifnull((dp.mont21*if(ep.deha21 = 'C',-1,1)),0))*100/(100+ep.piva21),2))/i.tipcam01,2)) as Conversion"
		sqlAct=sqlAct & " from "& oficina &"_extranet.ssdag"&tipope&"01 as i "
		sqlAct=sqlAct & "inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 "
        sqlAct=sqlAct & "  inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 "
        sqlAct=sqlAct & "inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  "
		sqlAct=sqlAct & "and ep.tmov21 =dp.tmov21 inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 "
		IF Vckcve=0 then
			sqlAct=sqlAct & " where i.rfccli01 in('"&Vrfc&"') and firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in("& conceptos &")   and cta.fech31 >='"&DateI&"'	 and cta.fech31 <='"&DateF&"' "
		ELSEIF Vckcve=1 and Vclave="Todos" THEN
			sqlAct=sqlAct & " where i.cvecli01 =cc.cvecli18  and firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in("& conceptos&")  and cta.fech31 >='"&DateI&"'	 and cta.fech31 <='"&DateF&"' "		
		ELSE
			sqlAct=sqlAct & " where i.cvecli01="&Vclave&" and  firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in("& conceptos &")   and cta.fech31 >='"&DateI&"'	 and cta.fech31 <='"&DateF&"' "
		END IF

		Set act3= Server.CreateObject("ADODB.Recordset")
		conn123 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act3.ActiveConnection = conn123
	act3.Source = sqlAct
	act3.cursortype=0
	act3.cursorlocation=2
	act3.locktype=1
	act3.open()
	valor=0
	if not(act3.eof) then
		valor = act3.fields(campo).value
		retornaTotal = valor
	else
		retornaTotal = valor
	end if
 
else 
 retornaTotal = valor
end if

End function
%>
<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE OPERACIONES COCA-COLA Y JDV .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html%>
	</BODY>
</HTML>