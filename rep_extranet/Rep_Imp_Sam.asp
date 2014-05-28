<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<%Server.ScriptTimeout=15000
Dim sRuta, codigo, archivocomparativo, archivoreporte, XML
Dim strFInicial,strFFinal
Dim reporterevisar, reporteenviar, reportecomparar,reportenuevo
nocolumns = 0
tablamov = ""
reporterevisar = 1
reporteenviar = 2
reportecomparar = 3
reportenuevo = 4
reporteSIR = 5
abreformatofuente = "<strong><br><font color=""#006699"" size=""4"" face=""Arial, Helvetica, sans-serif"">"
cierraformatofuente = "</font></strong>"
archivocomparativo = "C:\Reporte_Comparativo_Impuestos.xls"
archivoreporte = "C:\Reporte_Impuestos.xls"
sRuta = ""
guarda = ""



if  Session("GAduana") = "" then
	html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
else
	jnxadu=Session("GAduana")
	
	if Request.Form("Enviar") = "t" then
		if Request.Form("txtCorreo") = "" then
			Response.Write(abreformatofuente & "Debe de escribir por lo menos un correo.<br> Gracias." & cierraformatofuente)
			Response.End()	
		end if
	end if

	select case jnxadu
		case "VER"
			strOficina="rku"
		case "MEX"
			strOficina="dai"
		case "MAN"
			strOficina="sap"
		case "GUA"
			strOficina="rku"
		case "TAM"
			strOficina="ceg"
		case "LAR"
			strOficina="LAR"
		case "LZR"
			strOficina="lzr"
		case "TOL"
			strOficina="tol"
	end select
	
	if checaCargas then
		Response.Write(abreformatofuente & "Las Bases de Datos se estan actualizando y no es posible llevar a cabo su solicitud.<br>  Por Favor intente de nuevo en unos momentos.<br>  Gracias." & cierraformatofuente)
		Response.End()
	end if

	if Request.Form("Enviar") = "t" then
		'Regresa true si todas las validaciones estan bien, si no es true, se regresa el mensaje de error
		mensaje = checaValidacion("tienefechasoia")
		if not mensaje = "true" then
			Response.Write(abreformatofuente & mensaje & cierraformatofuente)
			Response.End()
		end if
		set file_FSO = createObject("scripting.filesystemobject")
		if (file_FSO.FileExists(archivoreporte)) then
			response.write(abreformatofuente & "Error 0001 " & archivoreporte & " ya Existe.<br> Hubo un error al generar el archivo, comuniquese con el área de sitemas.<br> Gracias.." & cierraformatofuente)
			response.end()
		end if
		
		html = GeneraArchivo
		
		set Stream = file_FSO.CreateTextFile(archivoreporte,true)
		stream.write(html)
		stream.close
		if (file_FSO.FileExists(archivoreporte)) then
			guarda = mid(guarda,1,len(guarda)-1)
			
			Set adodb = Server.CreateObject ("ADODB.Connection")
			adodb.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
			adodb.Execute(guarda)
			adodb.Close
			a = EnviarEmail("contacto@rkzego.com","Estimado",Request.Form("txtCorreo"),"Reporte de Impuestos Grupo Zego","","",1,archivoreporte,"Reporte de Impuestos")
		else
			Response.Write(abreformatofuente & "El archivo no pudo se creado.<br> Pongase en contacto con el area de Informàtica." & cierraformatofuente)
		end if
		
		if (file_FSO.FileExists(archivoreporte)) then
			file_FSO.DeleteFile(archivoreporte) 
		end if

	end if
	if Request.Form("Comparativo") = "s" then
		'html = GeneraComparativo
		html = GeneraLazaro
	end if
	if Request.Form("Nuevo") = "n" then
		html = GeneraArchivoNuevo

	end if
	if Request.Form("Enviar") <> "t" and Request.Form("Comparativo") <> "s" then
		html = GeneraArchivoNuevo'GeneraArchivo
	end if
	
end if

Function EnviarEmail(pFrom, pFromName, pTo, pSubject, pCC, pBCC, pPriority,Attm,nombrearchivo)
	strError = ""
	cuerpo = "Estimado cliente se adjunta " & nombrearchivo & " generado el día " & cstr(date()) & " a las " & cstr(TIME())
		if not pFrom="" and not pTo="" then
				sch = "http://schemas.microsoft.com/cdo/configuration/"
				Set cdoConfig = CreateObject("CDO.Configuration")
				With cdoConfig.Fields
				.Item(sch & "sendusing") = 2
				.Item(sch & "smtpserver") = "smtp.gmail.com"
				.Item(sch & "smtpserverport") = 465
				.Item(sch & "smtpconnectiontimeout") = 30
				.Item(sch & "smtpusessl") = true
				.Item(sch & "smtpauthenticate") = 1
				.Item(sch & "sendusername") = pFrom
				.Item(sch & "sendpassword") = "grk3yzqp2"
				.update
				End With

				Set MailObject = Server.CreateObject("CDO.Message")
				Set MailObject.Configuration = cdoConfig
				MailObject.From = pFrom
				MailObject.To = pTo
				MailObject.Subject = pSubject
				MailObject.HTMLBody = "<HTML><p><h4 align=""left""><font face=""Arial, Helvetica, Calibri, sans-serif"">" & cuerpo & "</font></h4></p>" &_
										"<br><h3 align=""left""><font face=""Helvetica, Calibri"">Reciba un Cordial Saludo.</font></h3><br><br>" &_
										"Si requiere mayor información sobre el contenido de este mensaje, por favor pongase en contacto con nuestro ejecutivo de cuenta.<br>" &_
										"<br>http://www.grupozego.com<BODY><CENTER>" & pHTML & "</CENTER></BODY></HTML>"
				MailObject.AddAttachment Attm
			
				On Error Resume Next
				MailObject.Send
				If Err <> 0 Then
					strError = Err.Description
				End If
				
				Set MailObject = Nothing
				Set cdoConfig = Nothing
		end if
	EnviarEmail = strError
End Function

function GeneraSQL(opcion,condicionopcional)'opcion 2
	SQL = ""
	if condicionopcional <> "" then
		condicion = condicionopcional
	else
		condicion = filtro(opcion)
	end if
	
	tablamov = "ssdagi01"
	select case opcion
		case reporterevisar
			SQL = 	"SELECT cast(CONCAT_WS('', i.patent01, '-', i.numped01) as char) Pedimento," &_
			" ifnull(i.refcia01,'-') Referencia, " &_
			" ifnull(fr.fraarn02,'0') FraccionAranc, " &_
			" ifnull(group_concat(fr.d_mer102),'-') Descripcion, " &_
			" i.fecent01 as 'Fecha de Entrada', " &_
			" i.tipcam01 as 'Tipo de Cambio', " &_
			" i.factmo01 as 'Factor Moneda', " &_
			" sum(ifnull(fr.prepag02,0)) 'Valor Merc Mon Nac', " &_
			" format(i.valdol01,2) as 'Valor Dolares', " &_
			" sum(ifnull(fr.cancom02,0)) as 'Total Quantity', " &_
			" sum(ifnull(fr.vmerme02,0)) as 'Invoice Amount', " &_
			" (select ifnull(sum(frac.prepag02),0) from " & strOficina & "_extranet.ssfrac02 as frac where frac.refcia02 = i.refcia01  and frac.patent02 = i.patent01 and frac.adusec02 = i.adusec01 ) as 'Tot Fac Mon Nac', " &_
			" i.valseg01 as 'Valor Seguros', " &_
			" i.segros01 as 'Seguros', " &_
			" i.fletes01 as 'Fletes', " &_
			" i.embala01 as 'Embalajes', " &_
			" i.incble01 as 'OtrosInc', " &_
			" format(sum(ifnull(fr.vaduan02,0)),0) 'Valor Aduana', " &_
			" format(if(i.cveped01 = 'R1',ifnull(cf13.import33,0),ifnull(ifnull(cf1.import36,0),0)),0) DTA, " &_
			" format(sum(ifnull(fr.i_adv102,0) + ifnull(fr.i_adv202,0) + ifnull(fr.i_adv302,0)),0) IGI, " &_
			" format(if(i.cveped01 = 'R1',ifnull(ifnull(cf153.import33,0),0),ifnull(ifnull(cf15.import36,0),0)),0) PRV, " &_
			" format(sum(ifnull(fr.i_iva102,0) + ifnull(fr.i_iva202,0) + ifnull(fr.i_iva302,0)),0) 'IVA', " &_
			" format(sum(ifnull(fr.i_cc0102,0) + ifnull(fr.i_cc0202,0) + ifnull(fr.i_cc0302,0)),0) 'CC', " &_
			" format(ifnull(cf7.import36,0) + ifnull(cf11.import36,0),0) 'MultasyRecargos', " &_
			" '' as 'OtrosPar', " &_
			" '' as 'OtrosPed', " &_
			" format((if(i.cveped01 = 'R1',ifnull(cf13.import33,0),ifnull(cf1.import36,0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf33.import33,0),ifnull(cf3.import36,0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf63.import33,0),ifnull(cf6.import36,0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf153.import33,0),ifnull(cf15.import36,0)) " &_
			
			" + if(i.cveped01 = 'R1',ifnull(cf23.import33,0),ifnull(ifnull(cf2.import36,0),0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf73.import33,0),ifnull(ifnull(cf7.import36,0),0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf113.import33,0),ifnull(ifnull(cf11.import36,0),0)) " &_
			
			
			" ),0) TotalImpuestos, " &_
			" i.tt_dta01, " &_
			" fr.tt_cc_02, " &_
			" fr.tasacc02, " &_
			" format(sum(ifnull(fr.cantar02,0)),3) as cantar02, " &_
			" ifnull(fr.ordfra02,0) as ordfra02," &_
			" ifnull(fr.tasadv02,0) as tasadv02, " &_
			" ifnull(if(i.facact01=0,1,i.facact01),0) as facact01, " &_
			" i.adusec01 as adu, " &_
			" i.patent01 " &_
			
			"from " & strOficina & "_extranet." & tablamov & " as i " &_
			" left join " & strOficina & "_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " &_
			"      left join " & strOficina & "_extranet.ssfrac02 as fr on fr.refcia02 = i.refcia01  and fr.patent02 = i.patent01 and fr.adusec02 = i.adusec01  " &_
			"           left join " & strOficina & "_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1' and cf1.adusec36 = i.adusec01 and cf1.patent36 =i.patent01  " &_
			"            left join " & strOficina & "_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3' and cf3.adusec36 = i.adusec01 and cf3.patent36 =i.patent01" &_
			"             left join " & strOficina & "_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6' and cf6.adusec36 = i.adusec01 and cf6.patent36 =i.patent01 " &_
			"              left join " & strOficina & "_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15' and cf15.adusec36 = i.adusec01 and cf15.patent36 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf13 on cf13.refcia33 = i.refcia01 and cf13.cveimp33 = '1' and cf13.adusec33 = i.adusec01 and cf13.patent33 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf33 on cf33.refcia33 = i.refcia01 and cf33.cveimp33 = '3' and cf33.adusec33 = i.adusec01 and cf33.patent33 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf63 on cf63.refcia33 = i.refcia01 and cf63.cveimp33 = '6' and cf63.adusec33 = i.adusec01 and cf63.patent33 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf153 on cf153.refcia33 = i.refcia01 and cf153.cveimp33 = '15' and cf153.adusec33 = i.adusec01 and cf153.patent33 =i.patent01 " &_
			"                 left join " & strOficina & "_extranet.sscont36 as cf11 on cf11.refcia36 = i.refcia01 and cf11.cveimp36 = '11' and cf11.adusec36 = i.adusec01 and cf11.patent36 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont33 as cf113 on cf113.refcia33 = i.refcia01 and cf113.cveimp33 = '11' and cf113.adusec33 = i.adusec01 and cf113.patent33 =i.patent01 " &_
			"                 left join " & strOficina & "_extranet.sscont36 as cf7 on cf7.refcia36 = i.refcia01 and cf7.cveimp36 = '7' and cf7.adusec36 = i.adusec01 and cf7.patent36 =i.patent01 " &_
			"                 left join " & strOficina & "_extranet.sscont33 as cf73 on cf73.refcia33 = i.refcia01 and cf73.cveimp33 = '7' and cf73.adusec33 = i.adusec01 and cf73.patent33 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '2' and cf2.adusec36 = i.adusec01 and cf2.patent36 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont33 as cf23 on cf23.refcia33 = i.refcia01 and cf23.cveimp33 = '2' and cf23.adusec33 = i.adusec01 and cf23.patent33 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont36 as cf20 on cf20.refcia36 = i.refcia01 and cf20.cveimp36 = '20' and cf20.adusec36 = i.adusec01 and cf20.patent36 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont33 as cf203 on cf203.refcia33 = i.refcia01 and cf203.cveimp33 = '20' and cf203.adusec33 = i.adusec01 and cf203.patent33 =i.patent01  " &_
			" where 1 " & condicion &_
			" group by i.refcia01, fr.fraarn02,fr.tasadv02  " & " order by 2,ordfra02 "
		case reporteenviar
			SQL = 	"SELECT i.patent01,i.numped01," &_
					" i.adusec01,i.refcia01 " &_
					" from " & strOficina & "_extranet." & tablamov & " as i " &_
					" left join " & strOficina & "_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " &_
					" where cc.rfccli18 = 'SEM950215S98' " & condicion
	
		case reportecomparar
			SQL = 	"SELECT i.refcia01 as refere " &_
					"from " & strOficina & "_extranet." & tablamov & " as i " &_
					" left join " & strOficina & "_extranet.ssclie18 as cc on i.rfccli01 = cc.rfccli18 " &_
					" right join " & strOficina & "_extranet.c01refer refe on i.refcia01 = refe.refe01 " &_
					" where cc.rfccli18 = 'SEM950215S98' " & condicion & " group by i.refcia01"
		case reportenuevo
			SQL = 	"SELECT cast(CONCAT_WS('', i.patent01, '-', i.numped01) as char) Pedimento," &_
			" ifnull(i.refcia01,'-') Referencia, " &_
			" ifnull(group_concat(distinct fr.fraarn02  separator '|'),'0') FraccionAranc, " &_
			" ifnull(group_concat(distinct fr.d_mer102),'-') Descripcion, " &_
			" i.fecent01 as 'Fecha de Entrada', " &_
			" i.tipcam01 as 'Tipo de Cambio', " &_
			" i.factmo01 as 'Factor Moneda', " &_
			" sum(ifnull(fr.prepag02,0)) 'Valor Merc Mon Nac', " &_
			" format(i.valdol01,2) as 'Valor Dolares', " &_
			" sum(ifnull(fr.cancom02,0)) as 'Total Quantity', " &_
			" sum(ifnull(fr.vmerme02,0)) as 'Invoice Amount', " &_
			" i.valseg01 as 'Valor Seguros', " &_
			" i.segros01 as 'Seguros', " &_
			" i.fletes01 as 'Fletes', " &_
			" i.embala01 as 'Embalajes', " &_
			" i.incble01 as 'OtrosInc', " &_
			" format(sum(ifnull(fr.vaduan02,0)),0) 'Valor Aduana', " &_
			" format(if(i.cveped01 = 'R1',ifnull(cf13.import33,0),ifnull(ifnull(cf1.import36,0),0)),0) DTA, " &_
			" format(sum(ifnull(fr.i_adv102,0) + ifnull(fr.i_adv202,0) + ifnull(fr.i_adv302,0)),0) IGI, " &_
			" format(if(i.cveped01 = 'R1',ifnull(ifnull(cf153.import33,0),0),ifnull(ifnull(cf15.import36,0),0)),0) PRV, " &_
			" format(sum(ifnull(fr.i_iva102,0) + ifnull(fr.i_iva202,0) + ifnull(fr.i_iva302,0)),0) 'IVA', " &_
			" format(sum(ifnull(fr.i_cc0102,0) + ifnull(fr.i_cc0202,0) + ifnull(fr.i_cc0302,0)),0) 'CC', " &_
			" format(ifnull(cf7.import36,0) + ifnull(cf11.import36,0),0) 'MultasyRecargos', " &_
			" '' as 'Otros', " &_
			" format((if(i.cveped01 = 'R1',ifnull(cf13.import33,0),ifnull(cf1.import36,0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf33.import33,0),ifnull(cf3.import36,0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf63.import33,0),ifnull(cf6.import36,0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf153.import33,0),ifnull(cf15.import36,0)) " &_
			
			" + if(i.cveped01 = 'R1',ifnull(cf23.import33,0),ifnull(ifnull(cf2.import36,0),0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf73.import33,0),ifnull(ifnull(cf7.import36,0),0)) " &_
			" + if(i.cveped01 = 'R1',ifnull(cf113.import33,0),ifnull(ifnull(cf11.import36,0),0)) " &_
			
			
			" ),0) TotalImpuestos, " &_
			" i.tt_dta01, " &_
			" ifnull(fr.ordfra02,0) as ordfra02," &_
			" ifnull(if(i.facact01=0,1,i.facact01),0) as facact01, " &_
			" i.adusec01 as adu, " &_
			" i.patent01 " &_
			
			"from " & strOficina & "_extranet." & tablamov & " as i " &_
			" left join " & strOficina & "_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " &_
			"      left join " & strOficina & "_extranet.ssfrac02 as fr on fr.refcia02 = i.refcia01  and fr.patent02 = i.patent01 and fr.adusec02 = i.adusec01  " &_
			"           left join " & strOficina & "_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1' and cf1.adusec36 = i.adusec01 and cf1.patent36 =i.patent01  " &_
			"            left join " & strOficina & "_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3' and cf3.adusec36 = i.adusec01 and cf3.patent36 =i.patent01" &_
			"             left join " & strOficina & "_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6' and cf6.adusec36 = i.adusec01 and cf6.patent36 =i.patent01 " &_
			"              left join " & strOficina & "_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15' and cf15.adusec36 = i.adusec01 and cf15.patent36 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf13 on cf13.refcia33 = i.refcia01 and cf13.cveimp33 = '1' and cf13.adusec33 = i.adusec01 and cf13.patent33 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf33 on cf33.refcia33 = i.refcia01 and cf33.cveimp33 = '3' and cf33.adusec33 = i.adusec01 and cf33.patent33 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf63 on cf63.refcia33 = i.refcia01 and cf63.cveimp33 = '6' and cf63.adusec33 = i.adusec01 and cf63.patent33 =i.patent01 " &_
			"               left join " & strOficina & "_extranet.sscont33 as cf153 on cf153.refcia33 = i.refcia01 and cf153.cveimp33 = '15' and cf153.adusec33 = i.adusec01 and cf153.patent33 =i.patent01 " &_
			"                 left join " & strOficina & "_extranet.sscont36 as cf11 on cf11.refcia36 = i.refcia01 and cf11.cveimp36 = '11' and cf11.adusec36 = i.adusec01 and cf11.patent36 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont33 as cf113 on cf113.refcia33 = i.refcia01 and cf113.cveimp33 = '11' and cf113.adusec33 = i.adusec01 and cf113.patent33 =i.patent01 " &_
			"                 left join " & strOficina & "_extranet.sscont36 as cf7 on cf7.refcia36 = i.refcia01 and cf7.cveimp36 = '7' and cf7.adusec36 = i.adusec01 and cf7.patent36 =i.patent01 " &_
			"                 left join " & strOficina & "_extranet.sscont33 as cf73 on cf73.refcia33 = i.refcia01 and cf73.cveimp33 = '7' and cf73.adusec33 = i.adusec01 and cf73.patent33 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '2' and cf2.adusec36 = i.adusec01 and cf2.patent36 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont33 as cf23 on cf23.refcia33 = i.refcia01 and cf23.cveimp33 = '2' and cf23.adusec33 = i.adusec01 and cf23.patent33 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont36 as cf20 on cf20.refcia36 = i.refcia01 and cf20.cveimp36 = '20' and cf20.adusec36 = i.adusec01 and cf20.patent36 =i.patent01  " &_
			"                 left join " & strOficina & "_extranet.sscont33 as cf203 on cf203.refcia33 = i.refcia01 and cf203.cveimp33 = '20' and cf203.adusec33 = i.adusec01 and cf203.patent33 =i.patent01  " &_
			" where cc.rfccli18 = 'SEM950215S98' " & condicion &_
			" group by i.refcia01  " & " order by 2 "
	end select 

	GeneraSQL = SQL
end function


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
	If IsNull(texto) = True Or texto = "" Then
		texto = "-"
	End If
	cell = 	"<td align=""center"">" &_
				
					texto &_
			
			"</td>"
	celdadatos = cell
end function

function celdanumero(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "0.00"
	End If
	cell = 	"<td align=""center"" style=""mso-number-format:'#,##0.00';"" >" &_
				
					texto &_
			
			"</td>"
	celdanumero = cell
end function

function celdanumeroentero(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "0"
	End If
	cell = 	"<td align=""center"" style=""mso-number-format:'#,##0';"" >" &_
				
					texto &_
			
			"</td>"
	celdanumeroentero = cell
end function


function celdasumas(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "-"
	End If
	cell = 	"<td align=""center"" style=""font-weight: bold"" >" &_
				
					texto &_
	
			"</td>"
	celdasumas = cell
end function

function celdasumasnumero(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "0.00"
	End If
	cell = 	"<td align=""center"" style=""font-weight: bold"" style=""mso-number-format:'#,##0.00';"" >" &_
				
					texto &_
	
			"</td>"
	celdasumasnumero = cell
end function

function celdasumasnumeroentero(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "0"
	End If
	cell = 	"<td align=""center"" style=""font-weight: bold"" style=""mso-number-format:'#,##0';"">" &_
				
					texto &_
	
			"</td>"
	celdasumasnumeroentero = cell
end function

function filtro(opcion)
	select case opcion
		case reporterevisar
			cadena = "'" & replace ((replace(replace(Request.Form("CServisWeb"),vbCrLf,"")," ","")),",","','") & "'"
			condicion = " and i.refcia01 in (" & cadena & ")"
		case reporteenviar
			cadena = "'" & replace ((replace(replace(Request.Form("CServisWeb"),vbCrLf,"")," ","")),",","','") & "'"
			condicion = " and i.refcia01 in (" & cadena & ")"
		case reportenuevo
			cadena = "'" & replace ((replace(replace(Request.Form("CServisWeb"),vbCrLf,"")," ","")),",","','") & "'"
			condicion = " and i.refcia01 in (" & cadena & ")"
		case reporteSIR
			cadena = "'" & replace ((replace(replace(Request.Form("CServisWeb"),vbCrLf,"")," ","")),",","','") & "'"
			condicion = " and REFERENCIA in (" & cadena & ")"
		case reportecomparar
			if not request.Form("txtdateIni") = "" then
			   strFInicial =request.Form("txtdateIni")
			end if
			if not request.Form("txtdateFin")  = "" then
			   strFFinal = request.Form("txtdateFin")
			end if
			if not isdate(strFInicial) then
			  Response.End()
			end if
			if not isdate(strFFinal) then
			  Response.End()
			end if

			condicion = " and refe.fdsp01 >= '" & FormatoFechaInv(strFInicial) & "' and refe.fdsp01 <= '" & FormatoFechaInv(strFFinal) & "'"
	end select
	
	filtro = condicion
end function

function VFME(referencia,oficina,fraccion,aduana,patente)
	dim valor
	 valor ="0"
	 
	 if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	 end if
	  if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	 end if

	sqlAct=" select format(sum(fact.valmex39),2) as val from " & oficina & "_extranet.ssfact39 as fact  " &_
			" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'"

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
		if not(act2.eof) then
			VFME =act2.fields("val").value
		else
			VFME =valor
		end if
end function

function PO(referencia,oficina,fraccion,aduana,patente)
	dim valor
	 valor ="-"
	 
	 if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	 end if
	  if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	 end if
	 
	sqlAct=" select group_concat(distinct ar.pedi05,char(05) separator '') as val from " & oficina & "_extranet.ssfact39 as f " &_
		" left join " & oficina & "_extranet.d05artic as ar on ar.refe05 = f.refcia39 and  ar.fact05 =f.numfac39 " &_
		" where f.refcia39 = '" & referencia & "' and  f.adusec39 = '" & aduana & "' and f.patent39 = '" & patente & "' "

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
		if not(act2.eof) then
			PO =act2.fields("val").value
		else
			PO =valor
		end if
end function

'Material number
function MAN(referencia,oficina,fraccion,aduana,patente)
	dim valor
	valor ="-"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	 
	sqlAct=" select group_concat(distinct ar.item05) as val from " & oficina & "_extranet.ssfact39 as f " &_
		" left join " & oficina & "_extranet.d05artic as ar on ar.refe05 = f.refcia39 and  ar.fact05 =f.numfac39 " &_
		" where f.refcia39 = '" & referencia & "' and  f.adusec39 = '" & aduana & "' and f.patent39 = '" & patente & "' " 

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()

	if not(act2.eof) then
		MAN =act2.fields("val").value
	else
		MAN =valor
	end if
end function

function FACTS(referencia,oficina,fraccion,aduana,patente)
	dim valor
	 valor ="-"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	sqlAct=" select group_concat(distinct fact.numfac39,char(05) separator '') as val from " & oficina & "_extranet.ssfact39 as fact  " &_
			" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'"
	 
	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	
	if not(act2.eof) then
		FACTS =act2.fields("val").value
	else
		FACTS =valor
	end if
end function

function INCO(referencia,oficina,fraccion,aduana,patente)
	dim valor
	valor ="-"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	 
	sqlAct=" select group_concat(distinct fact.terfac39,char(05) separator '') as val from " & oficina & "_extranet.ssfact39 as fact  " &_
			" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'"

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	
	if not(act2.eof) then
		INCO =act2.fields("val").value
	else
		INCO =valor
	end if
end function

function FACMON(referencia,oficina,fraccion,aduana,patente)
	dim valor
	valor ="0"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if

	sqlAct=" select ifnull(cast(group_concat(distinct fact.facmon39) as char),0) as val from " & oficina & "_extranet.ssfact39 as fact  " &_
			" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'" &_
			" and fact.numfac39 in (" & " select f.numfac39 from " & oficina & "_extranet.ssfact39 as f " &_
		" left join " & oficina & "_extranet.d05artic as ar on ar.refe05 = f.refcia39 and  ar.fact05 =f.numfac39 " &_
		" where f.refcia39 = '" & referencia & "' and  f.adusec39 = '" & aduana & "' and f.patent39 = '" & patente & "'" &")"

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	
	if not(act2.eof) then
		FACMON =act2.fields("val").value
	else
		FACMON =valor
	end if
end function

function MONFAC(referencia,oficina,fraccion,aduana,patente)
	dim valor
	 valor ="-"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if

	sqlAct=" select group_concat(distinct fact.monfac39) as val from " & oficina & "_extranet.ssfact39 as fact  " &_
			" where fact.refcia39 = '" & referencia & "' and fact.adusec39 = '" & aduana & "' and  fact.patent39 = '" & patente & "'"

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	
	if not(act2.eof) then
		MONFAC =act2.fields("val").value
	else
		MONFAC =valor
	end if
end function

function FechasAct()

	sqlAct=" SELECT * FROM registro_monitor WHERE ofic00 in ('RKU','DAI','SAP','CEG','LZR','TOL')"

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=intranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()

	Do Until act2.eof
		dato = dato & act2.Fields.Item("ofic00").Value & ":" & act2.Fields.Item("fecha_hora_act").Value & "  || "
	act2.MoveNext()
	Loop
	FechasAct = dato
end function

function conte(referencia,oficina,fraccion,aduana,patente)
	dim valor
	 valor ="-"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if

	if (ucase(oficina) = "LZR") then
		sqlAct=" select ifnull(group_concat(cont.numcon40),'-') as val from " & oficina & "_extranet.sscont40 as cont  " &_
			" where cont.refcia40 = '" & referencia & "' and cont.patent40 = '" & patente & "' and cont.adusec40 = '" & aduana & "' "
	else
		sqlAct=" select ifnull(group_concat(gui1.numgui04),'-') as val from " & oficina & "_extranet.ssguia04 as gui1  " &_
			" where gui1.refcia04 = '" & referencia & "' and gui1.patent04 = '" & patente & "' and gui1.adusec04 = '" & aduana & "' and gui1.idngui04 = 2 "
	end if

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	
	if not(act2.eof) then
		conte =act2.fields("val").value
	else
		conte =valor
	end if
end function

function TRATADO(referencia,oficina,fraccion,aduana,patente)
	dim valor
	 valor ="0"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if

	sqlAct=" select count(ipar.cveide12) as val from " & oficina & "_extranet.ssipar12 ipar " &_
			" where ipar.refcia12 = '" & referencia & "' and ipar.patent12 = '" & patente & "' and ipar.adusec12 = '" & aduana & "' and ipar.cveide12 = 'TL'"

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.open()
	
	if (act2.Fields.Item("val").Value > 0) then
		TRATADO = 1
	else
		TRATADO =valor
	end if
end function

Function checaCargas

	strSQL = "select count(*) as conteo from intranet.ban_extranet as b where b.m_bandera <> 'NA'"
	
	Set conn = Server.CreateObject ("ADODB.Connection")
	conn.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set recset = CreateObject("ADODB.RecordSet")
	Set recset = conn.Execute(strSQL)
	recset.MoveFirst()
	if recset.Fields.Item("conteo").Value = 0 then
		checaCargas = false
	else
		checaCargas = true
	end if
End Function

Function checaValidacion(valida)
	
	cumple = "true"
	select case valida 'Realiza validaciones por cada referencia
			case "tienefechasoia" 'Validaciones si esta seleccionado Enviar por correo
				strSQL = GeneraSQL(reporteenviar,"")
	end select

	Set conn1 = Server.CreateObject ("ADODB.Connection")
	conn1.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = conn1.Execute(strSQL)

	Do Until RSops.EOF
		select case valida 'Realiza validaciones por cada referencia
			case "tienefechasoia" 'Validaciones si esta seleccionado Enviar por correo
					'Revisa si tiene fecha de despacho en tabla soia
				if tieneFechasoia(RSops.Fields.Item("numped01").Value,RSops.Fields.Item("patent01").Value,RSops.Fields.Item("adusec01").Value,"730",RSops.Fields.Item("refcia01").Value,"fecdsp") <> "0" then
					cumple = "No puede enviar una solicitud de impuestos de una referencia despachada.<br> El archivo no se ha enviado.<br> Gracias."
					exit do
				end if
			case "comparable" 'Validaciones si esta seleccionado Enviar comparativo
					'Revisa si tiene fecha de despacho en tabla soia
				if tieneFechasoia(RSops.Fields.Item("numped01").Value,RSops.Fields.Item("patent01").Value,RSops.Fields.Item("adusec01").Value,"730",RSops.Fields.Item("refcia01").Value,"fecdsp") = "0" then
					cumple = "No puede enviar un comparativo de una operación que aún no tiene fecha de despacho.<br> Gracias"
					exit do
				end if
					'Revisa si se envió una solicitud de impuestos anteriomente para poder comparar
				if tieneFechasoia(RSops.Fields.Item("numped01").Value,RSops.Fields.Item("patent01").Value,RSops.Fields.Item("adusec01").Value,"730",RSops.Fields.Item("refcia01").Value,"fecenvio") = "0" then
					cumple = "No puede enviar un comparativo de una operación de la que no envió un reporte previo de solicitud de impuestos.<br> Gracias"
					exit do
				end if
					'Revisa si la fecha de despacho es mayor a la úlima vez en que se generó el repote de solicitud de impuestos
				if tieneFechasoia(RSops.Fields.Item("numped01").Value,RSops.Fields.Item("patent01").Value,RSops.Fields.Item("adusec01").Value,"730",RSops.Fields.Item("refcia01").Value,"fecdspmenor") = "registrodespues" then
					cumple = "Envió una solicitud de impuestos después haber despachado la operación, no se debe generar el comparativo con la información acual.<br> Favor " &_
							"de verificar con el área de sistemas.<br> Gracias"
					exit do
				end if
		end select
	Rsops.MoveNext()
	Loop

		checaValidacion = cumple
	
End Function
	'Realiza validacines deacuerdo a la opcion recibida regresa el valor devuelto por la consulta
Function tieneFechasoia(pedimento,patente,seccionaduana,situacion,referencia,opcion)
	select case opcion
			case "fecdsp" 'Revisa si tiene fecha de despacho en tabla soia
				strSQL = "select count(*) as valor from trackingbahia.bit_soia as b where b.Numped01 = '" & pedimento & "'" &_
						" and Numpat01 = " & patente & " and Adusec01 = '" & seccionaduana & "' and Detsit01 ='" & situacion & "'"
			case "fecenvio" 'Revisa si se envió una solicitud de impuestos anteriomente para poder comparar
				strSQL = "select count(fe.f_FechaRegistro) as valor from sistemas.enc004repimpsam as fe where fe.t_Referencia = '" & referencia & "'"
			case "fecdspmenor" 'Revisa si la fecha de despacho es mayor a la úlima vez en que se generó el repote de solicitud de impuestos
				strSQL = " select if((timestampdiff(second," &_
				" (select max(fe.f_FechaRegistro) as fe from sistemas.enc004repimpsam as fe where fe.t_Referencia = '" & referencia & "'),b.fecreg01) " &_
				" )>0,'depachoprimero','registrodespues') as valor " &_
				" from trackingbahia.bit_soia as b where b.Numped01 = '" & pedimento & "'" &_
				" and Numpat01 = " & patente & " and Adusec01 = '" & seccionaduana & "' and Detsit01 ='" & situacion & "'"
		end select

	Set conn = Server.CreateObject ("ADODB.Connection")
	conn.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set recset = CreateObject("ADODB.RecordSet")
	Set recset = conn.Execute(strSQL)
	recset.MoveFirst()

	tieneFechasoia = recset.Fields.Item("valor").Value
End Function

Function tieneFechaenviosolicitud(referencia)
	cumple = false
	
	strSQL = "select count(fe.f_FechaRegistro) as conteo from sistemas.enc004repimpsam as fe where fe.t_Referencia = '" & referencia & "'"

	Set conn = Server.CreateObject ("ADODB.Connection")
	conn.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set recset = CreateObject("ADODB.RecordSet")
	Set recset = conn.Execute(strSQL)
	recset.MoveFirst()

	if recset.Fields.Item("conteo").Value <> 0 then
		tieneFechaenviosolicitud = true
	end if
End Function

Function GeneraArchivo()
	nocolumns = 52
	
	query = GeneraSQL(reporterevisar,"")

	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)
	
	IF  RSops.BOF = True And RSops.EOF = True Then
		Response.Write( abreformatofuente & "No hay datos para esas condiciones.." & cierraformatofuente)
	Else
		
		info = 	"<table  width = ""2929""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""4"" face=""Arial"">" &_
											"<b>" &_
												"GRUPO ZEGO" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""3"" face=""Arial"">" &_
											"<b>" &_
												" SOLICITUD DE IMPUESTOS" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td>" &_
								"</td>" &_
							"</tr>" &_
				"</table>"
		
		header = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr bgcolor = ""#006699"" class = ""boton"">" &_
							    celdahead("PEDIMENTO") &_
								celdahead("REFERENCIA") &_
								celdahead("FRACCION") &_
								celdahead("DESCRIPCION") &_
								celdahead("MATERIAL NUMBER") &_
								celdahead("FACTURAS") &_
								celdahead("HOUSE BL o CONTENEDOR(LAZ)") &_
								celdahead("P/O No") &_
								celdahead("INCOTERMS") &_
								celdahead("FECHA DE ENTRADA") &_
								celdahead("TIPO DE CAMBIO") &_
								celdahead("FACTOR MONEDA") &_
								celdahead("VALOR FACTURA ME") &_
								celdahead("VALOR MERCANCIA MON NAC") &_
								celdahead("VALOR DOLARES") &_
								celdahead("TOTAL QUANTITY") &_
								celdahead("INVOICE AMOUNT") &_
								celdahead("VALOR TOTAL FACT MON NAC") &_
								celdahead("INVOICE CURRENCY") &_
								celdahead("VALOR SEGUROS PED") &_
								celdahead("VALOR SEGUROS FRACCION") &_
								celdahead("SEGUROS PED") &_
								celdahead("SEGUROS FRACCION") &_
								celdahead("FLETES PED") &_
								celdahead("FLETES FRACCION") &_
								celdahead("EMBALAJES PED") &_
								celdahead("EMBALAJES FRACCION") &_
								celdahead("OTROS INCREMENTABLES PED") &_
								celdahead("OTROS INC FRACCION") &_
								celdahead("VALOR ADUANA") &_
								celdahead("VALOR ADUANA CALC") &_
								celdahead("DTA") &_
								celdahead("DTA (CALC)") &_
								celdahead("IGI") &_
								celdahead("IGI (CALC)") &_
								celdahead("PRV") &_
								celdahead("CC") &_
								celdahead("CC (CALC)") &_
								celdahead("MULTAS Y RECARGOS") &_
								celdahead("OTROS PARTIDA") &_
								celdahead("OTROS PEDIMENTO") &_
								celdahead("IVA") &_
								celdahead("IVA (CALC)") &_
								celdahead("TOTAL IMPUESTOS") &_
								celdahead("TOTAL IMPUESTOS (CALC)") &_
								celdahead("CVE TIPO TASA DTA") &_
								celdahead("TASA ADV") &_
								celdahead("CVE TIPO TASA CC") &_
								celdahead("TASA CC") &_
								celdahead("CANTIDAD TARIFA") &_
								celdahead("FACTOR DE ACTUALIZACION") &_
								celdahead("TL")
		header = header &	"</tr>"
		datos = ""
		
		contador=4
		nueva=true
		refcia = ""
		
		guarda = guarda & "insert into sistemas.enc004repimpsam  (t_Pedimento,t_Referencia,t_Fraccion,t_Descripcion," &_
						" t_NumeroMaterial,t_Facturas,t_BL,t_OrdenDeCompra,t_Incoterms,t_FechaEntrada,t_TipoCambio,t_FactorMoneda," &_
						" t_ValFactMonExt,t_ValMercMonNac,t_ValorDolares,t_CantidadTotal,t_ValorFactura,t_MontoFactura,t_MonedaFactura,t_ValorSeguro," &_
						" t_Seguros,t_Fletes,t_Embalajes,t_OtrosInc,t_ValorAduana,t_DTA,t_IGI,t_PRV,t_CC,t_MultasyRecargos,t_OtrosPar,t_OtrosPed,t_IVA," &_
						" t_TotalImpuestos,t_CveTipoTasaDTA,t_TasaADV,t_CveTipoTasaCC,t_TasaCC,t_Cantidadtarifa,t_FactorActualizacion,t_AplicaTL) values "
			
		Do Until RSops.EOF
			contador = contador + 1
			
			if refcia <> "" then
				if RSops.Fields.Item("Referencia").Value = refcia  then
					nueva = false
				else
					nueva = true
				end if
			end if
			
			if nueva then
				prv = RSops.Fields.Item("PRV").Value
				dta = RSops.Fields.Item("DTA").Value
				otrosped = RSops.Fields.Item("OtrosPed").Value
				totimp = RSops.Fields.Item("TotalImpuestos").Value
			else
				prv = 0
				dta = 0
				otrosped = 0
				totimp = 0
			end if
			
			matnum = MAN(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			facturas = FACTS(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			contene = conte(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			puror = PO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			incoterms = INCO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			vafame = VFME(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			monfact = MONFAC(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			tienetl = TRATADO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
		
			guarda = 	guarda & " ('" & RSops.Fields.Item("Pedimento").Value & "' " &_
						", '" & RSops.Fields.Item("Referencia").Value & "' " &_
						", '" & RSops.Fields.Item("FraccionAranc").Value & "' " &_
						", '" & RSops.Fields.Item("Descripcion").Value & "' " &_
						", '" & matnum & "' " &_
						", '" & facturas & "' " &_
						", '" & contene & "' " &_
						", '" & puror & "' " &_
						", '" & incoterms & "' " &_
						", '" & RSops.Fields.Item("Fecha de Entrada").Value & "' " &_
						", '" & RSops.Fields.Item("Tipo de Cambio").Value & "' " &_
						", '" & RSops.Fields.Item("Factor Moneda").Value & "' " &_
						", '" & vafame & "' " &_
						", '" & "0" & "' " &_
						", '" & "0" & "' " &_
						", '" & RSops.Fields.Item("Total Quantity").Value & "' " &_
						", '" & RSops.Fields.Item("Invoice Amount").Value & "' " &_
						", '" & RSops.Fields.Item("Tot Fac Mon Nac").Value & "' " &_
						", '" & monfact & "' " &_
						", '" & RSops.Fields.Item("Valor Seguros").Value & "' " &_
						", '" & RSops.Fields.Item("Seguros").Value & "' " &_
						", '" & RSops.Fields.Item("Fletes").Value & "' " &_
						", '" & RSops.Fields.Item("Embalajes").Value & "' " &_
						", '" & RSops.Fields.Item("OtrosInc").Value & "' " &_
						", '" & RSops.Fields.Item("Valor Aduana").Value & "' " &_
						", '" & dta & "' " &_
						", '" & RSops.Fields.Item("IGI").Value & "' " &_
						", '" & prv & "' " &_
						", '" & RSops.Fields.Item("CC").Value & "' " &_
						", '" & RSops.Fields.Item("MultasyRecargos").Value & "' " &_
						", '" & RSops.Fields.Item("OtrosPar").Value & "' " &_
						", '" & otrosped & "' " &_
						", '" & RSops.Fields.Item("IVA").Value & "' " &_
						", '" & totimp & "' " &_
						", '" & RSops.Fields.Item("tt_dta01").Value & "' " &_
						", '" & RSops.Fields.Item("tasadv02").Value & "' " &_
						", '" & RSops.Fields.Item("tt_cc_02").Value & "' " &_
						", '" & RSops.Fields.Item("tasacc02").Value & "' " &_
						", '" & RSops.Fields.Item("cantar02").Value & "' " &_
						", '" & RSops.Fields.Item("facact01").Value & "' " &_
						", '" & tienetl & "' " &_
						"),"			
			
			datos = datos &	"<tr>" &_
			celdadatos(RSops.Fields.Item("Pedimento").Value) &_
			celdadatos(RSops.Fields.Item("Referencia").Value) &_
			celdadatos(RSops.Fields.Item("FraccionAranc").Value) &_
			celdadatos(RSops.Fields.Item("Descripcion").Value) &_
			celdadatos(matnum) &_
			celdadatos(facturas) &_
			celdadatos(contene) &_
			celdadatos(puror) &_
			celdadatos(incoterms) &_
			celdadatos(RSops.Fields.Item("Fecha de Entrada").Value) &_
			celdadatos(RSops.Fields.Item("Tipo de Cambio").Value) &_
			celdadatos(RSops.Fields.Item("Factor Moneda").Value) &_
			celdadatos(vafame) &_
			celdanumero("=REDONDEAR(O" & cstr(contador) & "*K" & cstr(contador) & ",0)") &_
			celdanumero("=Q" & cstr(contador) & "*L" & cstr(contador)) &_
			celdanumeroentero(RSops.Fields.Item("Total Quantity").Value) &_
			celdanumero(RSops.Fields.Item("Invoice Amount").Value) &_
			celdanumero(RSops.Fields.Item("Tot Fac Mon Nac").Value) &_
			celdadatos(monfact) &_
			celdanumero(RSops.Fields.Item("Valor Seguros").Value) &_
			celdanumero("=SI(R" & cstr(contador) & "=0,0,(N" & cstr(contador) & "/R" & cstr(contador) & ")*T" & cstr(contador) & ")") &_
			celdanumero(RSops.Fields.Item("Seguros").Value) &_
			celdanumero("=SI(R" & cstr(contador) & "=0,0,(N" & cstr(contador) & "/R" & cstr(contador) & ")*V" & cstr(contador) & ")") &_
			celdanumero( RSops.Fields.Item("Fletes").Value) &_
			celdanumero("=SI(R" & cstr(contador) & "=0,0,(N" & cstr(contador) & "/R" & cstr(contador) & ")*X" & cstr(contador) & ")") &_
			celdanumero(RSops.Fields.Item("Embalajes").Value) &_
			celdanumero("=SI(R" & cstr(contador) & "=0,0,(N" & cstr(contador) & "/R" & cstr(contador) & ")*Z" & cstr(contador) & ")") &_
			celdanumero(RSops.Fields.Item("OtrosInc").Value) &_
			celdanumero("=SI(R" & cstr(contador) & "=0,0,(N" & cstr(contador) & "/R" & cstr(contador) & ")*AB" & cstr(contador) & ")") &_
			celdanumeroentero(RSops.Fields.Item("Valor Aduana").Value) &_
			celdanumero("=REDONDEAR(N"&cstr(contador) & "+W"&cstr(contador) & "+Y"&cstr(contador)& "+AA"&cstr(contador) & "+AC"&cstr(contador)& ",0)") &_
			celdanumeroentero(dta) &_
			celdanumero("=SI(AZ" & cstr(contador) & "=0,SI(AT" & cstr(contador) & "=7,0.008*AD" & cstr(contador) & "*AY"&cstr(contador) & ",SI(AT" & cstr(contador) & "=4,AF" & cstr(contador) & ",0)),AF" & cstr(contador) & ")") &_
			celdanumeroentero(RSops.Fields.Item("IGI").Value) &_
			celdanumero("=REDONDEAR((AD"&cstr(contador) & "* (AU"&cstr(contador) & ")/100)*AY"&cstr(contador) & ",0)") &_
			celdanumeroentero(prv) &_
			
			
			celdanumeroentero(RSops.Fields.Item("CC").Value) &_
			celdanumeroentero("=REDONDEAR(SI(AV"&cstr(contador) & "=1,AD"&cstr(contador) & "*(AW"&cstr(contador) & "/100),SI(AV"&cstr(contador) & "=2,AX"&cstr(contador) & "*AW"&cstr(contador) & "*K"&cstr(contador) & ",0)),0)") &_
			celdanumeroentero(RSops.Fields.Item("MultasyRecargos").Value) &_
			celdanumeroentero(RSops.Fields.Item("OtrosPar").Value) &_
			celdanumeroentero(otrosped) &_
			
			celdanumeroentero(RSops.Fields.Item("IVA").Value) &_
			celdanumero("=REDONDEAR(((AD"&cstr(contador) & "+AG"&cstr(contador) & "+AI"&cstr(contador) & "+AL"&cstr(contador) & ")*0.16),0)") &_
			celdanumeroentero(totimp) &_
			celdanumero("=REDONDEAR(AG"&cstr(contador) & "+AI"&cstr(contador) & "+AJ"&cstr(contador)& "+AL"&cstr(contador) & "+AM"&cstr(contador) & "+AN"&cstr(contador)  &_
						"+AO"&cstr(contador) & "+AQ"&cstr(contador)  & ",0)") &_
			celdadatos(RSops.Fields.Item("tt_dta01").Value) &_
			celdadatos(RSops.Fields.Item("tasadv02").Value) &_
			celdadatos(RSops.Fields.Item("tt_cc_02").Value) &_
			celdadatos(RSops.Fields.Item("tasacc02").Value) &_
			celdadatos(RSops.Fields.Item("cantar02").Value) &_
			celdadatos(RSops.Fields.Item("facact01").Value) &_
			celdadatos(tienetl)
			datos = datos &	"</tr>"
		
			refcia =  RSops.Fields.Item("Referencia").Value
			Rsops.MoveNext()
		Loop
		'response.write(query)
		'response.end()
		sumas = ""
		sumas = "<tr>" &_
					"<td colspan=""" & 12 & """>" &_
						"<center>" &_
									"" &_
						"</center>" &_
					"</td>" &_
								
		celdasumas("SUMAS") &_
		celdasumasnumero("=SUMA(N5:N"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(O5:O"&cstr(contador)&")") &_
		celdasumasnumeroentero("=SUMA(P5:P"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(Q5:Q"&cstr(contador)&")") &_
		celdadatos("") &_
		celdadatos("") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(U5:U"&cstr(contador)&")") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(W5:W"&cstr(contador)&")") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(Y5:Y"&cstr(contador)&")") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(AA5:AA"&cstr(contador)&")") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(AC5:AC"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AD5:AD"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AE5:AE"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AF5:AF"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AG5:AG"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AH5:AH"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AI5:AI"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AJ5:AJ"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AK5:AK"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AL5:AL"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AM5:AM"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AN5:AN"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AO5:AO"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AP5:AP"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AQ5:AQ"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AR5:AR"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AS5:AS"&cstr(contador)&")") &_
		celdadatos("") &_
		celdadatos("") &_
		celdadatos("") &_
		celdadatos("") &_
		celdadatos("") &_
		celdadatos("") &_
		celdadatos("")
		sumas =  sumas & "</tr>"
	 	
		Response.Addheader "Content-Disposition", "attachment; filename=Reporte_Impuestos.xls"
		Response.ContentType = "application/vnd.ms-excel"
		GeneraArchivo = info & header & datos & sumas & "</table><br>" 
	end if
End Function

Function GeneraComparativo()
Dim referencias, info, header, datos, sumas, Info2, datos2, sumas2
info = ""
header = ""
datos = ""
sumas = ""
Info2 = ""
datos2 = ""
sumas2 = ""
referencias = ""

	nocolumns = 36
	contador=4
	info = 	"<table  width = ""2929""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">" &_
				"<tr>" &_
					"<td colspan=""" & nocolumns & """>" &_
						"<center>" &_
							"<font color=""#000000"" size=""4"" face=""Arial"">" &_
								"<b>" &_
									"GRUPO ZEGO" &_
								"</b>" &_
							"</font>" &_
						"</center>" &_
					"</td>" &_
				"</tr>" &_
				"<tr>" &_
					"<td colspan=""" & nocolumns & """>" &_
						"<center>" &_
							"<font color=""#000000"" size=""3"" face=""Arial"">" &_
								"<b>" &_
									" REPORTE COMPARATIVO DE IMPUESTOS" &_
								"</b>" &_
							"</font>" &_
						"</center>" &_
					"</td>" &_
				"</tr>" &_
				"<tr>" &_
					"<td bgcolor = ""#632523"" colspan=""" & 3 & """>" &_
						"<center>" &_
							"<font color=""#FFFFFF"" size=""3"" face=""Arial"">" &_
								"<b>" &_
									" ENVIADO ANTERIORMENTE" &_
								"</b>" &_
							"</font>" &_
						"</center>" &_
					"</td>" &_
				"</tr>" &_
			"</table>"
		
	header = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
					"<tr bgcolor = ""#006699"" class = ""boton"">" &_
						celdahead("PEDIMENTO") &_
						celdahead("REFERENCIA") &_
						celdahead("FRACCION") &_
						celdahead("DESCRIPCION") &_
						celdahead("MATERIAL NUMBER") &_
						celdahead("FACTURAS") &_
						celdahead("HOUSE BL o CONTENEDOR(LAZ)") &_
						celdahead("P/O No") &_
						celdahead("INCOTERMS") &_
						celdahead("FECHA DE ENTRADA") &_
						celdahead("TIPO DE CAMBIO") &_
						celdahead("FACTOR MONEDA") &_
						celdahead("VALOR FACTURA ME") &_

						celdahead("TOTAL QUANTITY") &_
						celdahead("INVOICE AMOUNT") &_
						celdahead("VALOR TOTAL FACT MON NAC") &_
						celdahead("INVOICE CURRENCY") &_
						celdahead("VALOR SEGUROS") &_
						celdahead("SEGUROS") &_
						celdahead("FLETES") &_
						celdahead("EMBALAJES") &_
						celdahead("OTROS INCREMENTABLES") &_
						celdahead("VALOR ADUANA") &_
						celdahead("DTA") &_
						celdahead("IGI") &_
						celdahead("PRV") &_
						celdahead("CC") &_
						celdahead("MULTAS Y RECARGOS") &_
						celdahead("OTROS PARTIDA") &_
						celdahead("OTROS PEDIMENTO") &_
						celdahead("IVA") &_
						celdahead("TOTAL IMPUESTOS") &_
						celdahead("CVE TIPO TASA DTA") &_
						celdahead("TASA ADV") &_
						celdahead("CVE TIPO TASA CC") &_
						celdahead("TASA CC") &_
						celdahead("CANTIDAD TARIFA") &_
						celdahead("FACTOR DE ACTUALIZACION") &_
						celdahead("TL")
						
	header = header &	"</tr>"
	info2 = "<table  width = ""2929""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">" &_
				"<tr>" &_
				"</tr>" &_
				"<tr>" &_
					"<td bgcolor = ""#632523"" colspan=""" & 3 & """>" &_
						"<center>" &_
							"<font color=""#FFFFFF"" size=""3"" face=""Arial"">" &_
								"<b>" &_
									" ACTUALIZADO " &_
								"</b>" &_
							"</font>" &_
						"</center>" &_
					"</td>" &_
				"</tr>" &_
			"</table>"
			
	query = GeneraSQL(reportecomparar,"")

	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)

	IF  not (RSops.BOF = True And RSops.EOF = True) Then
	
			Do Until RSops.EOF
				referencias = referencias & "'" & RSops.Fields.Item("refere").Value & "'"
				Rsops.MoveNext()
				if not RSops.EOF then
					referencias = referencias & ","
				end if
			Loop
		
		SQLReferencias = "select a.t_Referencia ,max(a.f_FechaRegistro) as fechas " &_ 
							" from sistemas.enc004repimpsam as a  where a.t_Referencia in " &_
							" (" & referencias & ")" &_
							" group by  a.t_Referencia "
		
		Set conn1 = Server.CreateObject ("ADODB.Connection")
		conn1.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		Set RSops = CreateObject("ADODB.RecordSet")
		Set RSops = conn1.Execute(SQLReferencias)
		
		IF  not (RSops.BOF = True And RSops.EOF = True) Then
		
			Do Until RSops.EOF
				fecha = datepart("yyyy",RSops.Fields.Item("fechas").Value) & "-" & datepart("m",RSops.Fields.Item("fechas").Value) & "-" & datepart("d",RSops.Fields.Item("fechas").Value) &_
							" " & datepart("h",RSops.Fields.Item("fechas").Value) & ":" & datepart("n",RSops.Fields.Item("fechas").Value) & ":" & datepart("s",RSops.Fields.Item("fechas").Value)

				SQLComparativo = SQLComparativo & "select * " &_ 
								" from sistemas.enc004repimpsam as a " &_
								" where a.t_Referencia = '" & RSops.Fields.Item("t_Referencia").Value & "'" &_
								" and f_FechaRegistro = '" & fecha & "'"
								
			Rsops.MoveNext()
				if not RSops.EOF then
					SQLComparativo = SQLComparativo & " union all "
				end if
			Loop
			
			
			Set conn1 = Server.CreateObject ("ADODB.Connection")
			conn1.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
			Set RSops = CreateObject("ADODB.RecordSet")
			Set RSops = conn1.Execute(SQLComparativo)
			
			datos = ""
			nueva=true
			refcia = ""

			Do Until RSops.EOF
				contador = contador + 1
				
				datos = datos &	"<tr>" &_
				celdadatos(RSops.Fields.Item("t_Pedimento").Value) &_
				celdadatos(RSops.Fields.Item("t_Referencia").Value) &_
				celdadatos(RSops.Fields.Item("t_Fraccion").Value) &_
				celdadatos(RSops.Fields.Item("t_Descripcion").Value) &_
				celdadatos(RSops.Fields.Item("t_NumeroMaterial").Value) &_
				celdadatos(RSops.Fields.Item("t_Facturas").Value) &_
				celdanumero(RSops.Fields.Item("t_BL").Value) &_
				celdanumero(RSops.Fields.Item("t_OrdenDeCompra").Value) &_
				celdadatos(RSops.Fields.Item("t_Incoterms").Value) &_
				celdadatos(RSops.Fields.Item("t_FechaEntrada").Value) &_
				celdadatos(RSops.Fields.Item("t_TipoCambio").Value) &_
				celdadatos(RSops.Fields.Item("t_FactorMoneda").Value) &_
				celdanumero(RSops.Fields.Item("t_ValFactMonExt").Value) &_
				
				celdanumeroentero(RSops.Fields.Item("t_CantidadTotal").Value) &_
				celdanumero(RSops.Fields.Item("t_ValorFactura").Value) &_
				celdanumero(RSops.Fields.Item("t_MontoFactura").Value) &_
				celdadatos(RSops.Fields.Item("t_MonedaFactura").Value) &_
				celdanumero(RSops.Fields.Item("t_ValorSeguro").Value) &_
				celdanumero(RSops.Fields.Item("t_Seguros").Value) &_
				celdanumero(RSops.Fields.Item("t_Fletes").Value) &_
				celdanumero(RSops.Fields.Item("t_Embalajes").Value) &_
				celdanumero(RSops.Fields.Item("t_OtrosInc").Value) &_
				celdanumeroentero(RSops.Fields.Item("t_ValorAduana").Value) &_
				celdanumeroentero(RSops.Fields.Item("t_DTA").Value) &_
				celdanumeroentero(RSops.Fields.Item("t_IGI").Value) &_
				celdanumeroentero(RSops.Fields.Item("t_PRV").Value) &_
				
				celdanumeroentero(RSops.Fields.Item("t_CC").Value) &_
				celdanumeroentero(RSops.Fields.Item("t_MultasyRecargos").Value) &_
				celdanumeroentero(RSops.Fields.Item("t_OtrosPar").Value) &_
				celdanumeroentero(RSops.Fields.Item("t_OtrosPed").Value) &_

				celdanumeroentero(RSops.Fields.Item("t_IVA").Value) &_				
				celdanumeroentero(RSops.Fields.Item("t_TotalImpuestos").Value) &_
				celdadatos(RSops.Fields.Item("t_CveTipoTasaDTA").Value) &_
				celdadatos(RSops.Fields.Item("t_TasaADV").Value) &_
				celdadatos(RSops.Fields.Item("t_CveTipoTasaCC").Value) &_
				celdadatos(RSops.Fields.Item("t_TasaCC").Value) &_
				celdadatos(RSops.Fields.Item("t_Cantidadtarifa").Value) &_
				celdadatos(RSops.Fields.Item("t_FactorActualizacion").Value) &_
				celdadatos(RSops.Fields.Item("t_AplicaTL").Value) &_
				datos = datos &	"</tr>"
			
				refcia =  RSops.Fields.Item("t_Referencia").Value
				Rsops.MoveNext()
			Loop

			sumas = ""
			sumas = "<tr>" &_
						"<td colspan=""" & 12 & """>" &_
							"<center>" &_
										"" &_
							"</center>" &_
						"</td>" &_
									
			celdasumas("SUMAS") &_
			celdasumasnumero("=SUMA(N5:N"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(O5:O"&cstr(contador)&")") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdasumasnumero("=SUMA(W5:W"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(X5:X"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(Y5:Y"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(Z5:Z"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(AA5:AA"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(AB5:AB"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(AB5:AC"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(AB5:AD"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(AB5:AE"&cstr(contador)&")") &_
			celdasumasnumero("=SUMA(AB5:AF"&cstr(contador)&")") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("")
			sumas =  sumas & "</tr>" & "</table>"
		else
			sumas =  "<tr>" & "</tr>" & "</table>"
		end if
			
		datos2 = ""
		
		contador2 = contador + 4
		nueva2 = true
		refcia2 = ""
		

		query = GeneraSQL(reporterevisar," and i.refcia01 in (" & referencias & ")")

		Set ConnStr = Server.CreateObject ("ADODB.Connection")
		ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		Set RSops = CreateObject("ADODB.RecordSet")
		Set RSops = ConnStr.Execute(query)
		
		IF  not (RSops.BOF = True And RSops.EOF = True) Then
		
			Do Until RSops.EOF
				contador2 = contador2 + 1
				
				if refcia2 <> "" then
					if RSops.Fields.Item("Referencia").Value = refcia2  then
						nueva2 = false
					else
						nueva2 = true
					end if
				end if
				
				if nueva2 then
					prv = RSops.Fields.Item("PRV").Value
					dta = RSops.Fields.Item("DTA").Value
					otrosped = RSops.Fields.Item("OtrosPed").Value
					totimp = RSops.Fields.Item("TotalImpuestos").Value
				else
					prv = 0
					dta = 0
					otrosped = 0
					totimp = 0
				end if
					
				matnum = MAN(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
				facturas = FACTS(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
				contene = conte(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
				puror = PO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
				incoterms = INCO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
				vafame = VFME(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
				monfact = MONFAC(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
				tienetl = TL(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
				
				datos2 = datos2 &	"<tr>" &_
				celdadatos(RSops.Fields.Item("Pedimento").Value) &_
				celdadatos(RSops.Fields.Item("Referencia").Value) &_
				celdadatos(RSops.Fields.Item("FraccionAranc").Value) &_
				celdadatos(RSops.Fields.Item("Descripcion").Value) &_
				celdadatos(matnum) &_
				celdadatos(facturas) &_
				celdadatos(contene) &_
				celdadatos(puror) &_
				celdadatos(incoterms) &_
				celdadatos(RSops.Fields.Item("Fecha de Entrada").Value) &_
				celdadatos(RSops.Fields.Item("Tipo de Cambio").Value) &_
				celdadatos(RSops.Fields.Item("Factor Moneda").Value) &_
				celdadatos(vafame) &_

				celdanumeroentero(RSops.Fields.Item("Total Quantity").Value) &_
				celdanumero(RSops.Fields.Item("Invoice Amount").Value) &_
				celdanumero(RSops.Fields.Item("Tot Fac Mon Nac").Value) &_
				celdadatos(monfact) &_
				celdanumero(RSops.Fields.Item("Valor Seguros").Value) &_
				celdanumero(RSops.Fields.Item("Seguros").Value) &_
				celdanumero( RSops.Fields.Item("Fletes").Value) &_
				celdanumero(RSops.Fields.Item("Embalajes").Value) &_
				celdanumero(RSops.Fields.Item("OtrosInc").Value) &_
				celdanumeroentero(RSops.Fields.Item("Valor Aduana").Value) &_
				celdanumeroentero(dta) &_
				celdanumeroentero(RSops.Fields.Item("IGI").Value) &_
				celdanumeroentero(prv) &_
				
				celdanumeroentero(RSops.Fields.Item("CC").Value) &_
				celdanumeroentero(RSops.Fields.Item("MultasyRecargos").Value) &_
				celdanumeroentero(RSops.Fields.Item("OtrosPar").Value) &_
				celdanumeroentero(otrosped) &_

				celdanumeroentero(RSops.Fields.Item("IVA").Value) &_				
				celdanumeroentero(totimp) &_
				celdadatos(RSops.Fields.Item("tt_dta01").Value) &_
				celdadatos(RSops.Fields.Item("tasadv02").Value) &_
				celdadatos(RSops.Fields.Item("tt_cc_02").Value) &_
				celdadatos(RSops.Fields.Item("tasacc02").Value) &_
				celdadatos(RSops.Fields.Item("cantar02").Value) &_
				celdadatos(RSops.Fields.Item("facact01").Value) &_
				celdadatos(tienetl)
				datos2 = datos2 &	"</tr>"
			
				refcia2 =  RSops.Fields.Item("Referencia").Value
				Rsops.MoveNext()
			Loop
				
			sumas2 = ""
			sumas2 ="<tr>" &_
						"<td colspan=""" & 12 & """>" &_
							"<center>" &_
										"" &_
							"</center>" &_
						"</td>" &_
									
			celdasumas("SUMAS") &_
			celdasumasnumero("=SUMA(N" & contador + 5 & ":N"&cstr(contador2)&")") &_
			celdasumasnumero("=SUMA(O" & contador + 5 & ":O"&cstr(contador2)&")") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdasumasnumero("=SUMA(W" & contador + 5 & ":W"&cstr(contador2)&")") &_
			celdasumasnumero("=SUMA(X" & contador + 5 & ":X"&cstr(contador2)&")") &_
			celdasumasnumero("=SUMA(Y" & contador + 5 & ":Y"&cstr(contador2)&")") &_
			celdasumasnumero("=SUMA(Z" & contador + 5 & ":Z"&cstr(contador2)&")") &_
			celdasumasnumero("=SUMA(AA" & contador + 5 & ":AA"&cstr(contador2)&")") &_
			celdasumasnumero("=SUMA(AB" & contador + 5 & ":AB"&cstr(contador2)&")") &_
			
			celdasumasnumero("=SUMA(AC" & contador + 5 & ":AC"&cstr(contador2)&")") &_
			celdasumasnumero("=SUMA(AD" & contador + 5 & ":AD"&cstr(contador2)&")") &_
			celdasumasnumero("=SUMA(AE" & contador + 5 & ":AE"&cstr(contador2)&")") &_
			celdasumasnumero("=SUMA(AF" & contador + 5 & ":AF"&cstr(contador2)&")") &_
			
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("") &_
			celdadatos("")
			sumas2 =  sumas2 & "</tr>"
		else
			sumas2 =  "<tr>" & "</tr>"
		end if
	end if
	Response.Addheader "Content-Disposition", "attachment; filename=Reporte_Comparativo_Impuestos.xls"
	Response.ContentType = "application/vnd.ms-excel"
	GeneraComparativo = info & header & datos & sumas & Info2 & header & datos2 & sumas2 & "</table><br>"

End Function


Function GeneraArchivoNuevo()
	nocolumns = 41

	query = GeneraSQL(reportenuevo,"")
	
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)

	IF  RSops.BOF = True And RSops.EOF = True Then
		Response.Write( abreformatofuente & "No hay datos para esas condiciones.." & cierraformatofuente)
	Else
		
		info = 	"<table  width = ""2929""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""4"" face=""Arial"">" &_
											"<b>" &_
												"GRUPO ZEGO" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""3"" face=""Arial"">" &_
											"<b>" &_
												" SOLICITUD DE IMPUESTOS" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td>" &_
								"</td>" &_
							"</tr>" &_
				"</table>"

		header = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr bgcolor = ""#006699"" class = ""boton"">" &_
							    celdahead("PEDIMENTO") &_
								celdahead("REFERENCIA") &_
								celdahead("FRACCION") &_
								celdahead("DESCRIPCION") &_
								celdahead("MATERIAL NUMBER") &_
								celdahead("FACTURAS")
								if strOficina = "LZR" then
									header = header & celdahead("CONTENEDOR")
								else
									header = header & celdahead("HOUSE BL")
								end if				
								header = header & celdahead("P/O No") &_
								celdahead("INCOTERMS") &_
								celdahead("FECHA DE ENTRADA") &_
								celdahead("TIPO DE CAMBIO") &_
								celdahead("FACTOR MONEDA") &_
								celdahead("VALOR FACTURA ME") &_
								celdahead("VALOR MERCANCIA MON NAC") &_
								celdahead("VALOR DOLARES") &_
								celdahead("TOTAL QUANTITY") &_
								celdahead("MONTO EN PESOS") &_
								celdahead("MONTO EN DOLARES") &_
								celdahead("INVOICE AMOUNT") &_
								celdahead("INVOICE CURRENCY") &_
								celdahead("VALOR SEGUROS") &_
								celdahead("SEGUROS") &_
								celdahead("FLETES") &_
								celdahead("EMBALAJES") &_
								celdahead("OTROS INCREMENTABLES") &_
								celdahead("VALOR ADUANA") &_
								celdahead("VALOR ADUANA CALC") &_
								celdahead("DTA") &_
								celdahead("DTA (CALC)") &_
								celdahead("IGI") &_
								celdahead("PRV") &_
								celdahead("CC") &_
								celdahead("MULTAS Y RECARGOS") &_
								celdahead("OTROS") &_
								celdahead("IVA") &_
								celdahead("IVA (CALC)") &_
								celdahead("TOTAL IMPUESTOS") &_
								celdahead("TOTAL IMPUESTOS (CALC)") &_
								celdahead("CVE TIPO TASA DTA") &_
								celdahead("FACTOR DE ACTUALIZACION") &_
								celdahead("TL")
		header = header &	"</tr>"
		datos = ""
		
		contador=4

		
		guarda = guarda & "insert into sistemas.enc004repimpsam  (t_Pedimento,t_Referencia,t_Fraccion,t_Descripcion," &_
						" t_NumeroMaterial,t_Facturas,t_BL,t_OrdenDeCompra,t_Incoterms,t_FechaEntrada,t_TipoCambio,t_FactorMoneda," &_
						" t_ValFactMonExt,t_ValMercMonNac,t_ValorDolares,t_CantidadTotal,t_ValorFactura,t_MonedaFactura,t_ValorSeguro," &_
						" t_Seguros,t_Fletes,t_Embalajes,t_OtrosInc,t_ValorAduana,t_DTA,t_IGI,t_PRV,t_CC,t_MultasyRecargos,t_OtrosPar,t_OtrosPed,t_IVA," &_
						" t_TotalImpuestos,t_CveTipoTasaDTA,t_TasaADV,t_CveTipoTasaCC,t_TasaCC,t_Cantidadtarifa,t_FactorActualizacion,t_AplicaTL) values "
			
		Do Until RSops.EOF
			contador = contador + 1
			
			matnum = MAN(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			facturas = FACTS(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			contene = conte(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			puror = PO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			incoterms = INCO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			vafame = VFME(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			monfact = MONFAC(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)
			tienetl = TRATADO(RSops.Fields.Item("Referencia").Value,mid(RSops.Fields.Item("Referencia").Value,1,3),RSops.Fields.Item("FraccionAranc").Value,RSops.Fields.Item("adu").Value,RSops.Fields.Item("patent01").Value)

			guarda = 	guarda & " ('" & RSops.Fields.Item("Pedimento").Value & "' " &_
						", '" & RSops.Fields.Item("Referencia").Value & "' " &_
						", '" & RSops.Fields.Item("FraccionAranc").Value & "' " &_
						", '" & RSops.Fields.Item("Descripcion").Value & "' " &_
						", '" & matnum & "' " &_
						", '" & facturas & "' " &_
						", '" & contene & "' " &_
						", '" & puror & "' " &_
						", '" & incoterms & "' " &_
						", '" & RSops.Fields.Item("Fecha de Entrada").Value & "' " &_
						", '" & RSops.Fields.Item("Tipo de Cambio").Value & "' " &_
						", '" & RSops.Fields.Item("Factor Moneda").Value & "' " &_
						", '" & vafame & "' " &_
						", '" & "0" & "' " &_
						", '" & "0" & "' " &_
						", '" & RSops.Fields.Item("Total Quantity").Value & "' " &_
						", '" & RSops.Fields.Item("Invoice Amount").Value & "' " &_
						", '" & monfact & "' " &_
						", '" & RSops.Fields.Item("Valor Seguros").Value & "' " &_
						", '" & RSops.Fields.Item("Seguros").Value & "' " &_
						", '" & RSops.Fields.Item("Fletes").Value & "' " &_
						", '" & RSops.Fields.Item("Embalajes").Value & "' " &_
						", '" & RSops.Fields.Item("OtrosInc").Value & "' " &_
						", '" & RSops.Fields.Item("Valor Aduana").Value & "' " &_
						", '" & RSops.Fields.Item("DTA").Value & "' " &_
						", '" & RSops.Fields.Item("IGI").Value & "' " &_
						", '" & RSops.Fields.Item("PRV").Value & "' " &_
						", '" & RSops.Fields.Item("CC").Value & "' " &_
						", '" & RSops.Fields.Item("MultasyRecargos").Value & "' " &_
						", '" & RSops.Fields.Item("Otros").Value & "' " &_
						", '" & RSops.Fields.Item("IVA").Value & "' " &_
						", '" & RSops.Fields.Item("TotalImpuestos").Value & "' " &_
						", '" & RSops.Fields.Item("facact01").Value & "' " &_
						", '" & tienetl & "' " &_
						"),"			
			
			
			datos = datos &	"<tr>" &_
			celdadatos(RSops.Fields.Item("Pedimento").Value) &_
			celdadatos(RSops.Fields.Item("Referencia").Value) &_
			celdadatos(RSops.Fields.Item("FraccionAranc").Value) &_
			celdadatos(RSops.Fields.Item("Descripcion").Value) &_
			celdadatos(matnum) &_
			celdadatos(facturas) &_
			celdadatos(contene) &_
			celdadatos(puror) &_
			celdadatos(incoterms) &_
			celdadatos(RSops.Fields.Item("Fecha de Entrada").Value) &_
			celdadatos(RSops.Fields.Item("Tipo de Cambio").Value) &_
			celdadatos(RSops.Fields.Item("Factor Moneda").Value) &_
			celdadatos(vafame) &_
			celdanumero("=REDONDEAR(O" & cstr(contador) & "*K" & cstr(contador) & ",0)") &_
			celdanumero("=S" & cstr(contador) & "*L" & cstr(contador)) &_
			celdanumeroentero(RSops.Fields.Item("Total Quantity").Value)
			if RSops.Fields.Item("Factor Moneda").Value <1 then
				datos = datos & celdanumero(RSops.Fields.Item("Invoice Amount").Value) &_
				celdanumeroentero("")
			else
				datos = datos & celdanumeroentero("") &_
				celdanumero(RSops.Fields.Item("Invoice Amount").Value)
			end if
			
			datos = datos & celdanumero(RSops.Fields.Item("Invoice Amount").Value) &_
			celdadatos(monfact) &_
			celdanumero(RSops.Fields.Item("Valor Seguros").Value) &_
			celdanumero(RSops.Fields.Item("Seguros").Value) &_
			celdanumero( RSops.Fields.Item("Fletes").Value) &_
			celdanumero(RSops.Fields.Item("Embalajes").Value) &_
			celdanumero(RSops.Fields.Item("OtrosInc").Value) &_
			celdanumeroentero(RSops.Fields.Item("Valor Aduana").Value) &_
			celdanumero("=REDONDEAR(N" & cstr(contador) & "+U"&cstr(contador) & "+V"&cstr(contador)& "+W"&cstr(contador) & "+X"&cstr(contador) & "+Y"&cstr(contador)& ",0)") &_
			celdanumeroentero(RSops.Fields.Item("DTA").Value) &_
			celdanumero("=SI(AO" & cstr(contador) & "=0,SI(AM" & cstr(contador) & "=7,0.008*AA" & cstr(contador) & "*AN"&cstr(contador) & ",SI(AM" & cstr(contador) & "=4,AB" & cstr(contador) & ",0)),AB" & cstr(contador) & ")") &_
			celdanumeroentero(RSops.Fields.Item("IGI").Value) &_
			celdanumeroentero(RSops.Fields.Item("PRV").Value) &_
			celdanumeroentero(RSops.Fields.Item("CC").Value) &_
			celdanumeroentero(RSops.Fields.Item("MultasyRecargos").Value) &_
			celdanumeroentero(RSops.Fields.Item("Otros").Value) &_
			
			celdanumeroentero(RSops.Fields.Item("IVA").Value) &_
			celdanumero("=REDONDEAR(((AA"&cstr(contador) & "+AC"&cstr(contador) & "+AD"&cstr(contador) & "+AF"&cstr(contador) & ")*0.16),0)") &_
			celdanumeroentero(RSops.Fields.Item("TotalImpuestos").Value) &_
			celdanumero("=REDONDEAR(AC"&cstr(contador) & "+AD"&cstr(contador) & "+AE"&cstr(contador)& "+AF"&cstr(contador) & "+AG"&cstr(contador) & "+AH"&cstr(contador)  &_
						"+AJ"&cstr(contador)  & ",0)") &_
			celdadatos(RSops.Fields.Item("tt_dta01").Value) &_
			celdadatos(RSops.Fields.Item("facact01").Value) &_
			celdadatos(tienetl)
			datos = datos &	"</tr>"

			Rsops.MoveNext()
		Loop
		'response.write(query)
		'response.end()

		sumas = ""
		sumas = "<tr>" &_
					"<td colspan=""" & 11 & """>" &_
						"<center>" &_
									"" &_
						"</center>" &_
					"</td>" &_
								
		celdasumas("SUMAS") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(N5:N"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(O5:O"&cstr(contador)&")") &_
		celdasumasnumeroentero("=SUMA(P5:P"&cstr(contador)&")") &_
		celdasumasnumeroentero("=SUMA(Q5:Q"&cstr(contador)&")") &_
		celdasumasnumeroentero("=SUMA(R5:R"&cstr(contador)&")") &_
		celdadatos("") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(U5:U"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(V5:V"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(W5:W"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(X5:X"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(Y5:Y"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(Z5:Z"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AA5:AA"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AB5:AB"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AC5:AC"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AD5:AD"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AE5:AE"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AF5:AF"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AG5:AG"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AH5:AH"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AI5:AI"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AJ5:AJ"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AK5:AK"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AL5:AL"&cstr(contador)&")") &_
		celdadatos("") &_
		celdadatos("") &_
		celdadatos("")	
		sumas =  sumas & "</tr>"
	 	
		Response.Addheader "Content-Disposition", "attachment; filename=Reporte_Impuestos.xls"
		Response.ContentType = "application/vnd.ms-excel"
response.write(info & header & datos & sumas & "</table><br>")
response.end()
		GeneraArchivoNuevo = info & header & datos & sumas & "</table><br>" 
		
	end if
End Function




Function GeneraLazaro()
	nocolumns = 41

	condicion = filtro(reporteSIR)
	query = "SELECT * FROM sir.dbo.GSI_VT_ImpuestosSamsung_I WHERE [RFC Imp/Exp] = 'SEM950215S98' " & condicion
			
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "PROVIDER=SQLOLEDB;DATA SOURCE=10.66.1.19;UID=sa;PWD=S0l1umF0rW;DATABASE=sir"
	Set RSops = CreateObject("ADODB.RecordSet")
	Set RSops = ConnStr.Execute(query)

	IF  RSops.BOF = True And RSops.EOF = True Then
		Response.Write( abreformatofuente & "No hay datos para esas condiciones.." & cierraformatofuente)
	Else
		
		info = 	"<table  width = ""2929""  border = ""0"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""4"" face=""Arial"">" &_
											"<b>" &_
												"GRUPO ZEGO" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td colspan=""" & nocolumns & """>" &_
									"<center>" &_
										"<font color=""#000000"" size=""3"" face=""Arial"">" &_
											"<b>" &_
												" SOLICITUD DE IMPUESTOS" &_
											"</b>" &_
										"</font>" &_
									"</center>" &_
								"</td>" &_
							"</tr>" &_
							"<tr>" &_
								"<td>" &_
								"</td>" &_
							"</tr>" &_
				"</table>"

		header = 	"<table  width = ""778""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" &_
							"<tr bgcolor = ""#006699"" class = ""boton"">" &_
							    celdahead("PEDIMENTO") &_
								celdahead("REFERENCIA") &_
								celdahead("FRACCION") &_
								celdahead("DESCRIPCION") &_
								celdahead("MATERIAL NUMBER") &_
								celdahead("FACTURAS")
								if strOficina = "LZR" then
									header = header & celdahead("CONTENEDOR")
								else
									'header = header & celdahead("HOUSE BL")
									header = header & celdahead("CONTENEDOR")
								end if				
								header = header & celdahead("P/O No") &_
								celdahead("INCOTERMS") &_
								celdahead("FECHA DE ENTRADA") &_
								celdahead("TIPO DE CAMBIO") &_
								celdahead("FACTOR MONEDA") &_
								celdahead("VALOR FACTURA ME") &_
								celdahead("VALOR MERCANCIA MON NAC") &_
								celdahead("VALOR DOLARES") &_
								celdahead("TOTAL QUANTITY") &_
								celdahead("MONTO EN PESOS") &_
								celdahead("MONTO EN DOLARES") &_
								celdahead("INVOICE AMOUNT") &_
								celdahead("INVOICE CURRENCY") &_
								celdahead("VALOR SEGUROS") &_
								celdahead("SEGUROS") &_
								celdahead("FLETES") &_
								celdahead("EMBALAJES") &_
								celdahead("OTROS INCREMENTABLES") &_
								celdahead("VALOR ADUANA") &_
								celdahead("VALOR ADUANA CALC") &_
								celdahead("DTA") &_
								celdahead("DTA (CALC)") &_
								celdahead("IGI") &_
								celdahead("PRV") &_
								celdahead("CC") &_
								celdahead("MULTAS Y RECARGOS") &_
								celdahead("OTROS") &_
								celdahead("IVA") &_
								celdahead("IVA (CALC)") &_
								celdahead("TOTAL IMPUESTOS") &_
								celdahead("TOTAL IMPUESTOS (CALC)") &_
								celdahead("CVE TIPO TASA DTA") &_
								celdahead("FACTOR DE ACTUALIZACION") &_
								celdahead("TL")
		header = header &	"</tr>"
		datos = ""
		
		contador=4

		
		'guarda = guarda & "insert into sistemas.enc004repimpsam  (t_Pedimento,t_Referencia,t_Fraccion,t_Descripcion," &_
		'				" t_NumeroMaterial,t_Facturas,t_BL,t_OrdenDeCompra,t_Incoterms,t_FechaEntrada,t_TipoCambio,t_FactorMoneda," &_
		'				" t_ValFactMonExt,t_ValMercMonNac,t_ValorDolares,t_CantidadTotal,t_ValorFactura,t_MonedaFactura,t_ValorSeguro," &_
		'				" t_Seguros,t_Fletes,t_Embalajes,t_OtrosInc,t_ValorAduana,t_DTA,t_IGI,t_PRV,t_CC,t_MultasyRecargos,t_OtrosPar,t_OtrosPed,t_IVA," &_
		'				" t_TotalImpuestos,t_CveTipoTasaDTA,t_TasaADV,t_CveTipoTasaCC,t_TasaCC,t_Cantidadtarifa,t_FactorActualizacion,t_AplicaTL) values "
			
		Do Until RSops.EOF
			contador = contador + 1
			
		'guarda = 	guarda & " ('" & RSops.Fields.Item("Pedimento").Value & "' " &_
		'				", '" & RSops.Fields.Item("Referencia").Value & "' " &_
		'				", '" & RSops.Fields.Item("FraccionAranc").Value & "' " &_
		'				", '" & RSops.Fields.Item("Descripcion").Value & "' " &_
		'				", '" & matnum & "' " &_
		'				", '" & facturas & "' " &_
		'				", '" & contene & "' " &_
		'				", '" & puror & "' " &_
		'				", '" & incoterms & "' " &_
		'				", '" & RSops.Fields.Item("Fecha de Entrada").Value & "' " &_
		'				", '" & RSops.Fields.Item("Tipo de Cambio").Value & "' " &_
		'				", '" & RSops.Fields.Item("Factor Moneda").Value & "' " &_
		'				", '" & vafame & "' " &_
		'				", '" & "0" & "' " &_
		'				", '" & "0" & "' " &_
		'				", '" & RSops.Fields.Item("Total Quantity").Value & "' " &_
		'				", '" & RSops.Fields.Item("Invoice Amount").Value & "' " &_
		'				", '" & monfact & "' " &_
		'				", '" & RSops.Fields.Item("Valor Seguros").Value & "' " &_
		'				", '" & RSops.Fields.Item("Seguros").Value & "' " &_
		'				", '" & RSops.Fields.Item("Fletes").Value & "' " &_
		'				", '" & RSops.Fields.Item("Embalajes").Value & "' " &_
		'				", '" & RSops.Fields.Item("OtrosInc").Value & "' " &_
		'				", '" & RSops.Fields.Item("Valor Aduana").Value & "' " &_
		'				", '" & RSops.Fields.Item("DTA").Value & "' " &_
		'				", '" & RSops.Fields.Item("IGI").Value & "' " &_
		'				", '" & RSops.Fields.Item("PRV").Value & "' " &_
		'				", '" & RSops.Fields.Item("CC").Value & "' " &_
		'				", '" & RSops.Fields.Item("MultasyRecargos").Value & "' " &_
		'				", '" & RSops.Fields.Item("Otros").Value & "' " &_
		'				", '" & RSops.Fields.Item("IVA").Value & "' " &_
		'				", '" & RSops.Fields.Item("TotalImpuestos").Value & "' " &_
		'				", '" & RSops.Fields.Item("facact01").Value & "' " &_
		'				", '" & tienetl & "' " &_
		'				"),"			
			
			
			datos = datos &	"<tr>" &_
			celdadatos(RSops.Fields.Item("PEDIMENTO").Value) &_
			celdadatos(RSops.Fields.Item("REFERENCIA").Value) &_
			celdadatos(RSops.Fields.Item("FRACCION").Value) &_
			celdadatos(RSops.Fields.Item("DESCRIPCION").Value) &_
			celdadatos(RSops.Fields.Item("MATERIAL NUMBER").Value) &_
			celdadatos(RSops.Fields.Item("FACTURAS").Value) &_
			celdadatos(RSops.Fields.Item("CONTENEDOR").Value) &_
			celdadatos(RSops.Fields.Item("P/O No").Value) &_
			celdadatos(RSops.Fields.Item("INCOTERMS").Value) &_
			celdadatos(RSops.Fields.Item("FECHA DE ENTRADA").Value) &_
			celdadatos(RSops.Fields.Item("TIPO DE CAMBIO").Value) &_
			celdadatos(RSops.Fields.Item("FACTOR MONEDA").Value) &_
			celdadatos(RSops.Fields.Item("VALOR FACTURA ME").Value) &_
			celdanumero("=REDONDEAR(O" & cstr(contador) & "*K" & cstr(contador) & ",0)") &_
			celdanumero("=S" & cstr(contador) & "*L" & cstr(contador)) &_
			
			
			celdanumeroentero(RSops.Fields.Item("TOTAL QUANTITY").Value)
			
			
				datos = datos & celdanumero(RSops.Fields.Item("MONTO EN PESOS").Value) &_
							celdanumero(RSops.Fields.Item("MONTO EN DOLARES").Value)
			
			
			datos = datos & celdanumero(RSops.Fields.Item("INVOICE AMOUNT").Value) &_
			celdadatos(RSops.Fields.Item("INVOICE CURRENCY").Value) &_
			celdanumero(RSops.Fields.Item("VALOR SEGUROS").Value) &_
			celdanumero(RSops.Fields.Item("SEGUROS").Value) &_
			celdanumero( RSops.Fields.Item("FLETES").Value) &_
			celdanumero(RSops.Fields.Item("EMBALAJES").Value) &_
			celdanumero(RSops.Fields.Item("OTROS INCREMENTABLES").Value) &_
			celdanumeroentero(RSops.Fields.Item("VALOR ADUANA").Value) &_
			celdanumero("=REDONDEAR(N" & cstr(contador) & "+U"&cstr(contador) & "+V"&cstr(contador)& "+W"&cstr(contador) & "+X"&cstr(contador) & "+Y"&cstr(contador)& ",0)") &_
			
			celdanumeroentero(RSops.Fields.Item("DTA").Value) &_
			celdanumero("=SI(AO" & cstr(contador) & "=0,SI(AM" & cstr(contador) & "=7,0.008*AA" & cstr(contador) & "*AN"&cstr(contador) & ",SI(AM" & cstr(contador) & "=4,AB" & cstr(contador) & ",0)),AB" & cstr(contador) & ")") &_
			
			celdanumeroentero(RSops.Fields.Item("IGI").Value) &_
			celdanumeroentero(RSops.Fields.Item("PRV").Value) &_
			celdanumeroentero(RSops.Fields.Item("CC").Value) &_
			celdanumeroentero(RSops.Fields.Item("MULTAS Y RECARGOS").Value) &_
			celdanumeroentero(RSops.Fields.Item("OTROS").Value) &_
			celdanumeroentero(RSops.Fields.Item("IVA").Value) &_
			celdanumero("=REDONDEAR(((AA"&cstr(contador) & "+AC"&cstr(contador) & "+AD"&cstr(contador) & "+AF"&cstr(contador) & ")*0.16),0)") &_
			
			celdanumeroentero(RSops.Fields.Item("TOTAL IMPUESTOS").Value) &_
			celdanumero("=REDONDEAR(AC"&cstr(contador) & "+AD"&cstr(contador) & "+AE"&cstr(contador)& "+AF"&cstr(contador) & "+AG"&cstr(contador) & "+AH"&cstr(contador)  &_
						"+AJ"&cstr(contador)  & ",0)") &_
			
			celdadatos(RSops.Fields.Item("CVE TIPO TASA DTA").Value) &_
			celdadatos(RSops.Fields.Item("FACTOR DE ACTUALIZACION").Value) &_
			celdadatos(RSops.Fields.Item("TL").Value)
			datos = datos &	"</tr>"

			Rsops.MoveNext()
			'celdanumero(RSops.Fields.Item("VALOR MERCANCIA MON NAC").Value) &_
			'celdanumero(RSops.Fields.Item("VALOR DOLARES").Value) &_
			'celdanumero(RSops.Fields.Item("VALOR ADUANA CALC").Value) &_
			'celdanumero(RSops.Fields.Item("DTA (CALC)").Value) &_
			'celdanumero(RSops.Fields.Item("IVA (CALC)").Value) &_
			'celdanumero(RSops.Fields.Item("TOTAL IMPUESTOS (CALC)").Value) &_
			'if RSops.Fields.Item("FACTOR MONEDA").Value <1 then
			'	datos = datos & celdanumero(RSops.Fields.Item("MONTO EN PESOS").Value) &_
			'	celdanumeroentero("")
			'else
			'	datos = datos & celdanumeroentero("") &_
			'	celdanumero(RSops.Fields.Item("MONTO EN DOLARES").Value)
			'end if
		Loop
		'response.write(query)
		'response.end()
			
		sumas = ""
		sumas = "<tr>" &_
					"<td colspan=""" & 11 & """>" &_
						"<center>" &_
									"" &_
						"</center>" &_
					"</td>" &_
								
		celdasumas("SUMAS") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(N5:N"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(O5:O"&cstr(contador)&")") &_
		celdasumasnumeroentero("=SUMA(P5:P"&cstr(contador)&")") &_
		celdasumasnumeroentero("=SUMA(Q5:Q"&cstr(contador)&")") &_
		celdasumasnumeroentero("=SUMA(R5:R"&cstr(contador)&")") &_
		celdadatos("") &_
		celdadatos("") &_
		celdasumasnumero("=SUMA(U5:U"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(V5:V"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(W5:W"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(X5:X"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(Y5:Y"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(Z5:Z"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AA5:AA"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AB5:AB"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AC5:AC"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AD5:AD"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AE5:AE"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AF5:AF"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AG5:AG"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AH5:AH"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AI5:AI"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AJ5:AJ"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AK5:AK"&cstr(contador)&")") &_
		celdasumasnumero("=SUMA(AL5:AL"&cstr(contador)&")") &_
		celdadatos("") &_
		celdadatos("") &_
		celdadatos("")	
		sumas =  sumas & "</tr>"
	 	
		Response.Addheader "Content-Disposition", "attachment; filename=Reporte_Impuestos.xls"
		Response.ContentType = "application/vnd.ms-excel"
response.write(info & header & datos & sumas & "</table><br>")
response.end()
		GeneraLazaro = info & header & datos & sumas & "</table><br>" 
		
	end if
End Function
%>

<HTML>
	<HEAD>
	<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
	<meta name=ProgId content=Excel.Sheet>
	<meta name=Generator content="Microsoft Excel 11">
		<TITLE>::.... REPORTE DE IMPUESTOS SAMSUNG.... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html
		%>
	</BODY>
</HTML>

