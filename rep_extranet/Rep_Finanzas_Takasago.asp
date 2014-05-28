<!-- #include virtual =  "PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp"-->
<%
On Error Resume Next
	Dim FechaI, FechaF, Corte,Permiso
	
	FechaI=FormatoFechaInv(Trim(Request.Form("fi")))
	FechaF=FormatoFechaInv(Trim(Request.Form("ff")))
	Corte=Request.Form("opc")
	strTipoUsuario = request.Form("TipoUser")
	strPermisos = Request.Form("Permisos")
	permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
	strOficina=Request.Form("OficinaG")
	Vrfc=Request.Form("rfcCliente")
	Vckcve=Request.Form("ckcve")
	Vclave=Request.Form("txtCliente")
	Tiporepo = Request.Form("TipRep")
	nocolumns=34
	if not permi = "" then
		permi = "  and (" & permi & ") "
	end if
	AplicaFiltro = False
	strFiltroCliente = ""
	strFiltroCliente = request.Form("txtCliente")
	if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
		blnAplicaFiltro = true
	end if
	if blnAplicaFiltro then
		permi = " AND i.cvecli01 =" & strFiltroCliente
	end if
	if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
		permi = ""
	end if
	
	if  Session("GAduana") = "" then
	html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
	else	
		dim Consulta
		if strOficina<>"Todas" then
				Select Case strOficina
					Case  "rku"
						condicion = filtro("VER")
					case  "dai"
						condicion = filtro("MEX")
					Case  "tol"
						condicion = filtro("TOL")
					Case  "sap"
						condicion = filtro("MAN")
					Case "lzr"
						condicion = filtro("LZR")
					Case "ceg"
							condicion = filtro("TAM")
				End Select
			consulta= QuerySQL(strOficina,condicion)
		else
			dim i 
			for i=0 to 5
				Select Case i
					Case 0
						aduanaTmp = "rku"
						condicion = filtro("VER")
					case 1
						aduanaTmp = "dai"
						condicion = filtro("MEX")
					Case 2
						aduanaTmp = "tol"
						condicion = filtro("TOL")
					Case 3
						aduanaTmp = "sap"
						condicion = filtro("MAN")
					Case 4
						aduanaTmp = "lzr"
						condicion = filtro("LZR")
					Case 5
						aduanaTmp = "ceg"
							condicion = filtro("TAM")
				End Select
				
				consulta=consulta& QuerySQL(aduanaTmp,condicion)
				
				if i<5 then
					consulta=consulta &" UNION ALL "
				end if
			next 
		end if
			
		Set ConnStr = Server.CreateObject ("ADODB.Connection")
		ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		Set RSops = CreateObject("ADODB.RecordSet")
		Set RSops = ConnStr.Execute(consulta)
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
								"<font color=""#000066"" size=""8"" face=""Arial, Helvetica, sans-serif"">" &_
									"<td colspan=""" & nocolumns & """>" &_
										"<p align=""left"">" &_
											"REPORTE FINANZAS" &_
										"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
											"<p>" &_
												"Del " & Request.Form("fi") & " Al " & Request.Form("ff") &_
											"</p>" &_
											"<p>" &_
											"</p>" &_
										"</td>" &_
									"</font>" &_
								"</strong>" &_
							"</tr>"
		
			header = 			"<tr class = ""boton"">" 
								header =header&	celdahead("Numero de pedimento") 
								header =header&	celdahead("Patente Aduanal")
								header =header&	celdahead("CTA GASTOS AGENTE ADUANAL") 
								header =header&	celdahead("Agente aduanal") 
								header =header&	celdahead("TIPO PEDIMENTO") 
								header =header&	celdahead("Clave de la Aduana") 
								header =header&	celdahead("Fecha de Pago") 
								header =header&	celdahead("ID") 
								header =header&	celdahead("PROVEEDOR EXTRANJERO")
								header =header&	celdahead("IMPORTE EN DLLS") 
								header =header&	celdahead("FACTURAS") 
								header =header&	celdahead("INCREMENTABLE FLETE INTERNACIONAL")
								header =header&	celdahead("OTROS INCREMENTABLE")
								header =header&	celdahead("T.C. PEDIMENTO")
								header =header&	celdahead("Valor de la Mercancia en Aduana")
								header =header&	celdahead("IGI")
								header =header&	celdahead("DTA")
								header =header&	celdahead("REC /IEPS U OTROS CONCEPTOS")
								header =header&	celdahead("BASE PARA IVA")
								header =header&	celdahead("Importe del IVA de la Mercancia en Aduana")
								header =header&	celdahead("PREV")
								header =header&	celdahead("FLETE MARITIMO/AEREO INTERNACIONAL")
								header =header&	celdahead("PROVEEDOR FLETE NACIONAL")
								header =header&	celdahead("RFC PROVEEDOR FLETE NACIONAL")
								header =header&	celdahead("FLETE 16% NACIONAL")
								header =header&	celdahead("FLETE 11% NACIONAL")
								header =header&	celdahead("AUTOPISTAS")
								header =header&	celdahead("IVA 16%")
								header =header&	celdahead("IVA 11%")
								header =header&	celdahead("IVA RETENIDO")
								header =header&	celdahead("MANIOBRAS")
								header =header&	celdahead("DESCONSOLIDACION")
								header =header&	celdahead("OTROS TASA 16%")
								header =header&	celdahead("OTROS TASA 11%")
								header =header&	celdahead("OTROS TASA 0%")
								header =header&	celdahead("IVA 16%")
								header =header&	celdahead("IVA 11%")
								header =header&	celdahead("HONORARIOS")
								header =header&	celdahead("SERVICIOS COMPLEMENTARIOS")
								header =header&	celdahead("IVA 16%")
								header =header&	celdahead("IVA 11%")
								header =header&	celdahead("TOTAL")
								header =header&	celdahead("TOTAL ANTICIPO A AGENTE/FONDEO CTA BANAMEX")
								header =header&	celdahead("DIFERENCIA A CARGO(FAVOR)")
				header = header &	"</tr>"
		dim Maniobra ,Desconsolidacion, Otros ,igi,dta,impma,prv,fleten,ivaflete,ivaretenido,ivaconceptos,honorarios,servcomp,ivaCG,contador
		Maniobra=0
		Desconsolidacion=0 
		Otros=0
		igi=0
		dta=0
		impma=0
		prv=0
		fleten=0
		ivaretenido=0
		ivaflete=0
		ivaconceptos=0
		honorarios=0
		servcomp=0
		ivaCG
		contador=5
			Do Until RSops.EOF
 
				datos = datos &	"<tr>" 
							datos = datos &celdadatos(RSops.Fields.Item("Pedimento").Value)
							datos = datos &celdadatos(RSops.Fields.Item("Patente").Value)
							datos = datos &celdadatos(RSops.Fields.Item("CG").Value)
							datos = datos &celdadatos(RSops.Fields.Item("Agente").Value)
							datos = datos &celdadatos(RSops.Fields.Item("cvePed").Value)
							datos = datos &celdadatos(RSops.Fields.Item("cveAdu").Value)
							datos = datos  &celdadatos(RSops.Fields.Item("Fecha Pago").value)
							datos = datos &celdadatos(RSops.Fields.Item("ID").Value)
							datos = datos &celdadatos(RSops.Fields.Item("Proveedor Extranjero").Value) 
							datos = datos &celdadatos(RSops.Fields.Item("Imp DLLS Facturas").Value) 
							datos = datos &celdadatos(RSops.Fields.Item("Facturas").Value) 
							datos = datos &celdadatos(RSops.Fields.Item("FleteInt").Value)
							datos = datos &celdadatos(RSops.Fields.Item("Otros Inc").Value)
							datos = datos &celdadatos(RSops.Fields.Item("TC Ped").Value) 
							datos = datos &celdadatos(RSops.Fields.Item("Valor Aduana").Value) 
							igi=RSops.Fields.Item("IGI").Value
							datos = datos &celdanumero(igi)
							dta=RSops.Fields.Item("DTA").Value
							datos = datos &celdanumero(dta)
							datos = datos &celdadatos(RSops.Fields.Item("REC/IEPS u OTROS").Value)
							datos = datos &celdadatos(RSops.Fields.Item("BASE PARA IVA").Value)
							impma=RSops.Fields.Item("IVA MercAduana").Value
							datos = datos &celdanumero(impma)
							prv=RSops.Fields.Item("PREV").Value
							datos = datos &celdanumero(prv)
							datos = datos &celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"FLETEAM"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"","Importe")) 'FLETE MARITIMO /AEREO INTERNACIONAL
							datos = datos &celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"FLETE"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"","Beneficiario")) 'PROVEEDOR FLETE NACIONAL
							datos = datos &celdadatos(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"FLETE"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"","RFC")) 'RFC PROVEEDOR FLETE NACIONAL
							datos = datos &celdanumero(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"FLETE"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"16.00","ImporteF")) 'FLETE 16% NACIONAL
							datos = datos &celdanumero(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"FLETE"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"11.00","ImporteF")) 'FLETE 11%
							datos = datos &celdadatos("") 'AUTOPISTAS
							datos = datos &celdanumero(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"FLETE"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"16.00","ImpIVA")) 'IVA 16%
							datos = datos &celdanumero(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"FLETE"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"11.00","ImpIVA")) 'IVA 11%
							datos = datos &celdanumero(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(Rsops.Fields.Item("Referencia").Value,1,3),"FLETE"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"","Retencion")) 'IVA RETENIDO
								Maniobra=retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"MANIOBRAS"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"","Importe")
							datos = datos &celdanumero(Maniobra) 'MANIOBRAS
								Desconsolidacion=retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"DESCONSOLIDACION"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"","Importe")
							datos = datos &celdanumero(Desconsolidacion) 'DESCONSOLIDACION
								Otros=retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"OTROS"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"16.00","Importe")
							datos = datos &celdanumero(Otros) 'OTROS TASA 16%
							datos = datos &celdanumero(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"OTROS"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"11.00","Importe")) 'OTROS TASA 11 %
							datos = datos &celdanumero(retornaPagosHechos(RSops.Fields.Item("Referencia").value,retornaConceptosPH(mid(RSops.Fields.Item("Referencia").value,1,3),"OTROS"),mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"0.00","Importe")) 'OTROS TASA 0%
							datos = datos &celdanumero(retornaPagosHechos(RSops.Fields.Item("Referencia").value,"",mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"16.00","ImpIVAMDO")) 'IVA 16%
							datos = datos &celdanumero(retornaPagosHechos(RSops.Fields.Item("Referencia").value,"",mid(RSops.Fields.Item("Referencia").value,1,3),RSops.Fields.Item("CG").Value,"11.00","ImpIVAMDO")) 'IVA 11%
								Honorarios=RetornaHonorarioBit(RSops.Fields.Item("Referencia").Value,RSops.Fields.Item("CG").Value,"i","H")
									Servcomp=RetornaHonorarioBit(RSops.Fields.Item("Referencia").Value,RSops.Fields.Item("CG").Value,"i","S")			
			datos = datos &celdanumero(Honorarios) 'HONORARIOS
							datos = datos &celdanumero(Servcomp) 'SERVICIOS COMPLEMENTARIOS
							if RSops.Fields.Item("piva31")="16" then 
								datos = datos &celdanumero((Honorarios+Servcomp)*.16) 'IVA 16%
								else
								datos = datos &celdanumero("") 'IVA 16%
							end if
							if RSops.Fields.Item("piva31")="11" then 
								datos = datos &celdanumero((Honorarios+Servcomp)*.11) 'IVA 11%
								else
								datos = datos &celdanumero("") 'IVA 11%
							end if
							'=SUMA(P5,Q5,T5,U5,Y5,Z5,AA5,AB5,AC5,AD5,AE5,AF5,AG5,AH5,AI5,AJ5,AK5,AL5,AM5,AN5,AO5)
							datos = datos &celdanumero("=SUMA(P"&contador&",Q"&contador&",T"&contador&",U"&contador&",V"&contador&",Y"&contador&",Z"&contador&",AA"&contador&",AB"&contador&",AC"&contador&",AD"&contador&",AE"&contador&",AF"&contador&",AG"&contador&",AH"&contador&",AI"&contador&",AJ"&contador&",AK"&contador&",AL"&contador&",AM"&contador&",AN"&contador&",AO"&contador&")") 'TOTAL
							datos = datos &celdadatos(RSops.Fields.Item("CuadroLiquidacion").value) 'TOTAL ANTICIPO A AGENTE /FONDEO CTA BANAMEX
							datos = datos &celdadatos("=AP"&contador&"-AQ"&contador) 'DEFERENCIA A CARGO(FAVOR)
				datos = datos &	"</tr>"
				contador=contador+1
				Rsops.MoveNext()
			Loop
	
			html = info & header & datos & "</table><br>"
		end if
	end if
	function celdanumero(texto)
	If IsNull(texto) = True Or texto = "" Then
		texto = "0.00"
	End If
	cell = 	"<td align=""center"" style=""mso-number-format:'#,##0.00';"" >" &_
				
					texto &_
			
			"</td>"
	celdanumero = cell
end function


	function retornaConceptosPH(oficina,topico)
		dim cad
		cad = "NA"
	
		if (ucase(oficina) = "ALC")then
			oficina = "LZR"
		elseif (ucase(oficina)="PAN")then
			oficina="DAI"
		end if
 
		if oficina = "SAP" then
			if topico = "FLETE" then
				cad= "322,5"
			end if
			if topico="FLETEAM" then 
				cad="32,4"
			end if
			if topico = "MANIOBRAS" then
				cad= "111,181,258,289,2,86,183,196,209,287,290,389,390"
			end if
			if topico = "DESCONSOLIDACION" then
				cad="17,18,140,168"
			end if
			if topico="OTROS" THEN
				cad="NE"
			end if
		else 
			if oficina = "CEG" then
  
				if topico = "FLETE" then
					cad= "15"
				end if
				if topico="FLETEAM" then 
				cad="5"
			end if
				if topico = "MANIOBRAS" then
					cad="2,77,248,249,250,252,287,288,289,290,294,295,304,336,342" 
				end if
				if topico = "DESCONSOLIDACION" then
					cad= "85,254,306,331,373"
				end if
				IF topico="OTROS" THEN
					cad="NE"
				END IF
		else 
			if oficina = "TOL" then
		
				if topico = "FLETE" then
					cad= "7"
				end if
				if topico="FLETEAM" then 
				cad="3,19,74"
			end if
				if topico = "MANIOBRAS" then
					cad= "2,127"
				end if
				if topico = "DESCONSOLIDACION" then
					cad="6,94"
				end if
				if topico = "OTROS" then
					cad="NE"
				end if
		else 
			if oficina = "LZR" then
				if topico = "FLETE" then
					cad= "15,96"
				end if
				if topico="FLETEAM" then 
				cad="35,5"
			end if
				if topico = "MANIOBRAS" then
					cad="2,115,116,125,160,167,230,297"
				end if
				if topico="DESCONSOLIDACION" then
					cad="85,324"
				end if
				if topico="OTROS" then
					cad="NE"
				end if
		else 
	       if oficina = "RKU" then
		        if topico = "FLETE" then
					cad="15"
				end if
				if topico="FLETEAM" then 
				cad="5,35"
			end if
				if topico = "MANIOBRAS" then
					cad="2"
				end if
				if topico="DESCONSOLIDACION" then 	
					cad="85,368"
				end if
				if topico="OTROS" then 	
					cad="NE"
				end if
        else 
		    if oficina = "DAI" then
		        if topico = "FLETE" then
					cad="7"
				end if
				if topico="FLETEAM" then 
				cad="3,19"
			end if
				if topico = "MANIOBRAS" then
					cad="2,127"
				end if
				if topico="DESCONSOLIDACION" then 	
					cad="6"
				end if
				if topico="OTROS" then 	
					cad="NE"
				end if
			else 
   		    cad = "NA"
        end if
        end if
		end if
		end if
		end if
		end if
		retornaConceptosPH = cad
	end function
	
	Function HonorariosAtlas(Ref,op)
if op ="H" then 
	op="in('00428','00027')"
elseif op="S" then 
	op="in('00012','00316')"
end if
	sqlAct="SELECT sum(d32.mont32) Importe"&_
			"	FROM atv_extranet.d31refer AS d "&_
			"	LEFT JOIN atv_extranet.e31cgast AS e31 ON e31.cgas31=d.cgas31 AND e31.esta31<>'C' "&_
			"	left join  atv_extranet.e32rserv e on e.refe32=d.refe31 and e.cgas32=e31.cgas31 "&_
			"	left join atv_extranet.d32rserv as d32 on d32.refe32=e.refe32 "&_
			"	where d.refe31='"&Ref&"' and d32.ttar32  "&op
			
	Set act2= Server.CreateObject("ADODB.Recordset")
	
	conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=atv_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	
	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	if not(act2.eof) then
		HonorariosAtlas=act2.fields("Importe").value
	else 
		HonorariosAtlas=""
	end if
	act2.Close()
	
end Function
	function retornaPagosHechos(referencia,conceptos,oficina,cta,iva,Opc)
		dim c,valor
		c=chr(34)
		valor=0
		if iva<>"" then 
			iva=" and ep.piva21='"&iva&"' "
		end if
		if (ucase(oficina) = "ALC")then	
			oficina = "LZR"
		elseif (ucase(oficina="PAN")) then 
			oficina="DAI" 
		end if
		if conceptos="" then 
			conceptos=retornaConceptosPH(oficina,"FLETE")
			conceptos=" and ep.conc21 not in("&conceptos&") "
		elseif conceptos <> "NE" then 
			conceptos=" and ep.conc21 in ("&conceptos&")"
		elseif conceptos="NE" then
			conceptos=retornaConceptosPH(oficina,"FLETE")&","&retornaConceptosPH(oficina,"MANIOBRAS")&","&retornaConceptosPH(oficina,"DESCONSOLIDACION")&","&retornaconceptosPH(oficina,"FLETEAM")
			conceptos=" and ep.conc21 not in ("&conceptos&")"
		
		end if
		
		if(conceptos <> "NA" )then

			sqlAct="select r.refe31 as Ref,bn2.nomb20 Beneficiario,bn2.rfc20 RFC, r.cgas31,ep.conc21,ep.piva21,round((sum(dp.mont21*if(ep.deha21 = 'C',-1,1))/((ep.piva21/100)+1)),2) as Importe, round((ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0)/((ep.piva21/100)+1-.04)),2) ImporteF,cp.desc21 ,(round((ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0)/((ep.piva21/100)+1-.04)),2)*(if(ep.piva21<>0,ep.piva21/100,1))) as ImpIVA ,round((ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0)/((ep.piva21/100)+1-.04)),2)*.04 Retencion, " & _
			" (round((ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0)),2)*(if(ep.piva21<>0,ep.piva21/100,1))) as ImpIVAMDO "&_
			" from  "& oficina &"_extranet.d31refer as r  " & _
			"     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 " & _
			"          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = r.refe31 and dp.cgas21 = r.cgas31 " & _
			"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
			"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
			" 				   left join (select distinct bn.clav20, bn.nomb20, bn.rfc20 from  "& oficina &"_extranet.c20benef as bn  ) as bn2 on bn2.clav20=ep.bene21 "&_
			"    where  cta.esta31 <> 'C' "&conceptos&" and r.refe31 = '"&referencia&"' and cta.cgas31='"&cta&"' "&iva

			Set act2= Server.CreateObject("ADODB.Recordset")
			conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
			'if Opc="ImpIVAMDO" then 
			'	response.write(sqlAct)
			'	response.end()
			'end if
			act2.ActiveConnection = conn12
			act2.Source = sqlAct
			act2.cursortype=0
			act2.cursorlocation=2
			act2.locktype=1
			act2.open()
			if not(act2.eof) then
				valor = act2.fields(Opc).value
				retornaPagosHechos = valor
			else
				retornaPagosHechos = valor
			end if
		else
			retornaPagosHechos =valor
		end if

	end function
	
	function retornaTransportista(oficina,referencia,opcion)
	Trans=""
	
		if oficina="ALC" then 
			oficina="LZR" 
		elseif oficina="PAN" THEN
			oficina="DAI" 
		end if
			sqlAct="select group_concat(distinct c.nom02) as Transportista ,'' as RFC from "&oficina&"_extranet.d01conte d" & _
			" left join "&oficina&"_extranet.e01oemb e on d.peri01 = e.peri01 and d.nemb01 = e.nemb01" & _
			" left join "&oficina&"_extranet.c56trans c on c.cve02 = e.ctra01 " & _
			"where d.refe01 ='"&referencia&"' and d.nemb01 != 0 group by d.refe01"
		
		Set act2= Server.CreateObject("ADODB.Recordset")
		conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427" 
		act2.ActiveConnection = conn12
		act2.Source = sqlAct
		act2.cursortype=0
		act2.cursorlocation=2
		act2.locktype=1
		act2.open()
		if not(act2.eof) then
			while not act2.eof
				Trans =act2.fields(opcion).value
				act2.movenext()
			wend
			retornaTransportista = Trans
		else
			retornaTransportista =Trans
		end if 
		
	end function 
	
	function filtro(strGaduanaG)
	if Vckcve = 0  and Vckcve <>"" then
	'ckcve
		if Vrfc <> "0" then
			condicion = "AND i.rfccli01 = '" & Vrfc & "' "
		else
			condicion = " "
		end if
	else
		if Vclave <> "Todos" and Vclave<>"" Then
			condicion = "AND i.cvecli01 = " & Vclave & " "
		Else
			permi = PermisoClientes(strGaduanaG,strPermisos,"i.cvecli01")
			condicion = permi
			condicion = "AND (" & condicion &")"
			if condicion = "AND (i.cvecli01=0 )" Then
				condicion = ""
			end if
		End If
	end if
	filtro = condicion
	end function
	
	
	function QuerySQL(oficina,condi)
		if oficina="ALC" then 
			oficina="LZR"
		elseif oficina="PAN" then
			oficina="DAI"
		end if
		
		dim sql
		sqlAct=""
		sqlAct=" select i.numped01 'Pedimento', i.refcia01 Referencia, "
		sqlAct=sqlAct&" i.patent01 'Patente', "
		sqlAct=sqlAct&" e.cgas31 'CG', "
		sqlAct=sqlAct&" if( "
		sqlAct=sqlAct&" mid(i.refcia01,1,3)='RKU','GRUPO REYES KURI S.C.', "
		sqlAct=sqlAct&" if((mid(i.refcia01,1,3)='SAP' OR mid(i.refcia01,1,3)='ALC'),'SERVICIOS ADUANALES DEL PACIFICO S.C.', "
		sqlAct=sqlAct&" if((mid(i.refcia01,1,3)='TOL' or mid(i.refcia01,1,3)='CEG'),'COMERCIO EXTERIOR DEL GOLFO S.C.', "
		sqlAct=sqlAct&" if(mid(i.refcia01,1,3)='DAI','DESPACHOS AEREOS INTEGRADOS, S.C.','')))) as'Agente', "
		sqlAct=sqlAct&" i.cveped01 as 'cvePed', "
		sqlAct=sqlAct&" i.cveadu01 as 'cveAdu', "
		sqlAct=sqlAct&" i.fecpag01 as 'Fecha Pago', "
		sqlAct=sqlAct&" p.irspro22 as 'ID', "
		sqlAct=sqlAct&" p.nompro22 as 'Proveedor Extranjero', "
		sqlAct=sqlAct&" sum(f.valdls39)  as 'Imp DLLS Facturas', "
		sqlAct=sqlAct&" 	group_concat(distinct f.numfac39) as 'Facturas', "
		sqlAct=sqlAct&" 	i.fletes01 as 'FleteInt', "
		sqlAct=sqlAct&" 	i.incble01 as 'Otros Inc', "
		sqlAct=sqlAct&" 	i.tipcam01 as 'TC Ped', "
		sqlAct=sqlAct&" 	(select sum(ifnull(fr.vaduan02,0))  from "&oficina&"_extranet.ssfrac02 as fr where fr.refcia02 =i.refcia01 and fr.adusec02=i.adusec01 and fr.patent02=i.patent01 ) as 'Valor Aduana', "
		sqlAct=sqlAct&" ((select sum(fr.vaduan02) from "&oficina&"_extranet.ssfrac02 as fr where fr.refcia02 =i.refcia01 and fr.adusec02=i.adusec01 and fr.patent02=i.patent01 )+i_dta101+ifnull(if(i.cveped01<>'R1' , igi.import36, igi2.import33),0) ) 'BASE PARA IVA' ,  "
		sqlAct=sqlAct&" 	ifnull(if(i.cveped01<>'R1' , igi.import36, igi2.import33),0) IGI, "
		sqlAct=sqlAct&" 	ifnull(if(i.cveped01<>'R1',dta.import36, dta2.import33),0) DTA, "
		sqlAct=sqlAct&" 	ifnull(if(i.cveped01<>'R1', sum(otros.import36),sum(otros2.import33)),0) 'REC/IEPS u OTROS', "
		sqlAct=sqlAct&" 	ifnull(if(i.cveped01 <>'R1',iva.import36 , iva2.import33),0) 'IVA MercAduana', "
		sqlAct=sqlAct&" 	ifnull(if(i.cveped01<>'R1',prv.import36, prv2.import33 ),0) PREV, "
		sqlAct=sqlAct&" 	e.chon31 'Honorarios', e.piva31 , "
		sqlAct=sqlAct&"if(i.cveped01<>'R1',(select ifnull(sum(import36),0) as campo from "& oficina &"_extranet.sscont36 as cf1  where refcia36 = i.refcia01),(select ifnull(sum(import33),0) as campo from "& oficina &"_extranet.sscont33 as cf1  where refcia33 = i.refcia01)) as 'CuadroLiquidacion' "
		sqlAct=sqlAct&" From "&oficina&"_extranet.ssdagi01 as i  "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.d31refer as d on d.refe31=i.refcia01 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.e31cgast as e on e.cgas31=d.cgas31 and e.esta31<>'C' "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.ssfact39 as f on f.refcia39=i.refcia01 and f.adusec39=i.adusec01 and f.patent39=i.patent01 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.ssprov22 as p on p.cvepro22=f.cvepro39 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont36 as igi on igi.refcia36=i.refcia01 and igi.patent36=i.patent01 and igi.adusec36=i.adusec01 and igi.cveimp36=6 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont33 as igi2 on igi2.refcia33=i.refcia01 and igi2.patent33=i.patent01 and igi2.adusec33=i.adusec01 and igi2.cveimp33=6 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont36 as dta on dta.refcia36=i.refcia01 and dta.adusec36=i.adusec01 and dta.patent36=i.patent01 and dta.cveimp36=1 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont33 as dta2 on dta2.refcia33=i.refcia01 and dta2.adusec33=i.adusec01 and dta2.patent33=i.patent01 and dta2.cveimp33=1 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont36 as otros on otros.refcia36=i.refcia01 and otros.adusec36=i.adusec01 and otros.patent36=i.patent01 and otros.cveimp36 not in (1,3,15,6) "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont33 as otros2 on otros2.refcia33=i.refcia01 and otros2.adusec33=i.adusec01 and otros2.patent33=i.patent01 and otros2.cveimp33  not in (1,3,15,6) "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont36 as iva on iva.refcia36=i.refcia01 and iva.adusec36=i.adusec01 and iva.patent36=i.patent01 and iva.cveimp36 =3 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont33 as iva2 on iva2.refcia33=i.refcia01 and iva2.adusec33=i.adusec01 and iva2.patent33=i.patent01 and iva2.cveimp33 =3 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont36 as prv on prv.refcia36=i.refcia01 and prv.adusec36=i.adusec01 and prv.patent36=i.patent01 and prv.cveimp36 =15 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.sscont33 as prv2 on prv2.refcia33=i.refcia01 and prv2.adusec33=i.adusec01 and prv2.patent33=i.patent01 and prv2.cveimp33 =15 "
		sqlAct=sqlAct&" left join "&oficina&"_extranet.ssfrac02 as fra on fra.refcia02=i.refcia01 and fra.adusec02=i.adusec01 and fra.patent02=i.patent01 "
		sqlAct=sqlAct&" where  i.firmae01<>'' and i.firmae01 is not null " &condi &"  and i.fecpag01>='"&FechaI&"' and i.fecpag01<='"&FechaF&"'"
		sqlAct=sqlAct&" group by i.refcia01, e.cgas31 "
		'response.write(sqlAct)
		'response.end()
		QuerySQL=sqlAct
		
	end function
	
	function celdahead(texto)'Celda de encabezado de la tabla
		cell = "<td bgcolor = ""#000000"" width=""200"" nowrap>" &_
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

	function celdadatos(texto)'Celda de datos de la tabla
		On error resume next
			If IsNull(texto) = True Or texto = "" Then
				texto = "&nbsp;"
			End If
			dim c 
			c=chr(34)
			cell = 	"<td align=""center""nowrap bgcolor=#FFFFFF >" &_
						"<font size=""1"" face=""Arial"">" &texto &"</font>" &_
					"</td>"
			celdadatos = cell
	end function
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
%>

<HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE FINANZAS .... ::</TITLE>
	</HEAD>
	<BODY>
		<%=html%>
	</BODY>
</HTML>