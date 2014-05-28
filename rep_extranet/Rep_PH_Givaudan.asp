<!-- #include virtual =  "PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp"-->
<% 'On Error Resume Next

'http://10.66.1.9/portalmysql/extranet/ext-asp/reportes/Rep_PH_Givaudan.asp

	Dim FechaI, FechaF, Permiso,ColPH
	FechaI=FormatoFechaInv(Trim(Request.Form("fi"))) 'Fecha inicio
	FechaF=FormatoFechaInv(Trim(Request.Form("ff"))) 'Fecha fin
	strOficina= retornaOficina(Session("GAduana")) 'Oficina a generar	
	Tiporepo = Request.Form("TipRep") ' HTML y Excel
	mov=Request.Form("mov") 'Importacion / Exportacion / Ambos
	strFiltroCliente = request.Form("rfcCliente")
	ColPH=ColumnasPH(mov,strOficina,FechaI,FechaF,"NC",0)
	nocolumns=21+ColPH 'Obtener numero de columnas a generar 


	if  Session("GAduana") = "" then
		html = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
	else	
	
		 if Tiporepo = 2 Then
			Response.Addheader "Content-Disposition", "attachment;; filename=Indicadores_PH.xls"
			 Response.ContentType = "application/vnd.ms-excel"
		 End If
				
		datos=retornaMYSQL(mov,strOficina,FechaI,FechaF) & "</table>"
		html=datos
	End if
		
	function celdanumero(texto)
		If IsNull(texto) = True Or texto = "" Then
			texto = "0"
		End If
		cell = 	"<td align=""center"" style=""mso-number-format:'#,##0.00';"" ><font size=""1"" face=""Arial"">" &_
					texto &_
			"</font></td>"
		celdanumero = cell
	end function	
	
	function celdadatos(texto)'Celda de datos de la tabla
		On error resume next
			If IsNull(texto) = True Or texto = "" Then
				texto = "&nbsp;"
			End If
			dim c 
			c=chr(34)
			cell = 	"<td align=""center""nowrap bgcolor=#FFFFFF ><font size=""1"" face=""Arial"">" &texto &"</font></td>"
			celdadatos = cell
	end function
		
	function celdahead(texto)'Celda de encabezado de la tabla
		cell = "<td bgcolor = ""#FF0000"" width=""200"" nowrap>" &_
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
	
	function ColumnasPH(ie,ofi,fi,ff,campo,R) 'Esta funcion puede obtener la sub consulta para pagos Hechos, Numero de columnas de PH a utilizar y los encabezados de los mismos
		consulta=""
		dim valor 
		if campo="NC" then 
			valor =0
		else 
			valor=""
		end if 
		consulta="select cast(group_concat(Tab.Sentencia)as char) as CPH , count(Tab.Sentencia) NC ,Tab.NConcepto as Nconceptos "&_
				 " from(	select "&_
				"	concat('sum(if(T.conc21 =',ep.conc21,',T.Importe,0)) as ""',replace(cp.desc21,'.',''),'""' ) as Sentencia ,replace(cp.desc21,'.','')NConcepto "&_
				"	from "&ofi&"_extranet.ssdag"&ie&"01 as i "&_
				"	inner join "&ofi&"_extranet.d31refer as r on r.refe31 = i.refcia01  "&_
				"	inner join "&ofi&"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 "&_
				"	inner join "&ofi&"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 "&_
				"	inner join "&ofi&"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  and ep.tmov21 =dp.tmov21  "&_
				"	inner join  "&ofi&"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 "&_
				"	where  i.rfccli01 in('"&strFiltroCliente&"') and i.fecpag01>='"&fi&"' and i.fecpag01<='"&ff&"' "&_
				"	and cta.esta31 <> 'C' and ep.tpag21<>3 "&_
				"	group by cp.clav21 "&_
				"	) as Tab  "
		if R>0 then 
			consulta=consulta & "group by Tab.Sentencia"
		end if
		consulta=consulta &" order by Tab.Sentencia desc"
		
		Set ConnStr = Server.CreateObject ("ADODB.Connection")
		ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		Set RSops = CreateObject("ADODB.RecordSet")
		Set RSops = ConnStr.Execute(consulta)
	
		IF RSops.BOF = True And RSops.EOF = True Then
			
		else
			if R=0 then
				valor=RSops.Fields.Item(campo).Value
			elseif R>0 and campo="Nconceptos" then 
					Do Until RSops.EOF
						valor=valor & celdahead(RSops.Fields.Item(campo).value)
					Rsops.MoveNext()
					Loop
				
			end if
		end if
		ConnStr.close()
		
		ColumnasPH=valor
	end function 
	
	function retornaMYSQL(ie,ofi,fi,ff)
	consulta=""
		consulta="Select  "&_
					"T.REFERENCIA, T.patent01 PATENTE, T.numped01 'NUMERO DE PEDIMENTO', T.cveped01 'CLAVE DE PEDIMENTO',  "&_
					"T.regime01 REGIMEN, T.ValorAduana 'VALOR ADUANA',if(T.Ncontenedores>0,'Contenedor','Carga Suelta')'TIPO DE CARGA',  "&_
					"T.DTA, T.PRV,T.ECI, T.IVA, T.IGI,rr.fdoc01 'RECEPCIÓN DE DOCUMENTOS',T.Entrada 'FECHA DE ENTRADA',rr.frev01 'FECHA DE REVALIDACIÓN',rr.fpre01 'FECHA DE PREVIO',T.fecpag01 'FECHA DE PAGO',  "&_
					"rr.fdsp01 'FECHA DE DESPACHO',T.fech31 'FACTURACIÓN','' as 'INDICADOR DE DESPACHO'  "
					if ColPH<>0 then 
						consulta=consulta & " , " & ColumnasPH(ie,ofi,fi,ff,"CPH",0)
					end if
					
		consulta=consulta&", T.cgas31 CG,''AS 'INDICADOR DE FACTURACIÓN', "&_
						KPI("", "rr.fdsp01", "T.fech31") & "AS 'KPIACUSE' " &_
						"from (  "&_
						"select i.refcia01 as REFERENCIA  "&_
						", i.patent01  "&_
						", i.numped01  "&_
						", i.cveped01  "&_
						", i.regime01  "&_
						", i.fecpag01  "
						if ie="i" then 
							consulta=consulta &", i.fecent01 as Entrada"
						else 
							consulta=consulta &", i.fecpre01 as Entrada "
						end if 
		consulta=consulta&", cta.fech31  "&_
						", r.cgas31  "&_
						", ep.conc21  "&_
						", ep.piva21  "&_
						", i.pesobr01  "&_
						", ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as Importe  "&_
						", cp.desc21  "&_
						", (select sum(f.vaduan02 )from "&ofi&"_extranet.ssfrac02 as f where f.refcia02=i.refcia01 and f.patent02 =i.patent01 )ValorAduana  "&_
						",(select count(g.numcon40)  "&_
						"from "&ofi&"_extranet.sscont40 as g where g.refcia40=i.refcia01 and g.adusec40=i.adusec01 and g.patent40=i.patent01 ) as NContenedores  "&_
						", Contri.*  "&_
						"from "&ofi&"_extranet.ssdag"&ie&"01 as i  "&_
							"left join (  "&_
										"select ii.refcia01  "&_
										",if(ii.cveped01<>'R1', sum(if(s.cveimp36 =1, ifnull(s.import36,0),0)),sum(if(s3.cveimp33 =1, ifnull(s3.import33,0),0))) DTA  "&_
										",if(ii.cveped01<>'R1', sum(if(s.cveimp36 =15, ifnull(s.import36,0),0)),sum(if(s3.cveimp33 =15, ifnull(s3.import33,0),0))) PRV  "&_
										",if(ii.cveped01<>'R1', sum(if(s.cveimp36 =18, ifnull(s.import36,0),0)),sum(if(s3.cveimp33 =18, ifnull(s3.import33,0),0))) ECI  "&_
										",if(ii.cveped01<>'R1', sum(if(s.cveimp36 =3, ifnull(s.import36,0),0)),sum(if(s3.cveimp33 =3, ifnull(s3.import33,0),0))) IVA  "&_
										",if(ii.cveped01<>'R1', sum(if(s.cveimp36 =6, ifnull(s.import36,0),0)),sum(if(s3.cveimp33 =6, ifnull(s3.import33,0),0))) IGI  "&_
										"from "&ofi&"_extranet.ssdag"&ie&"01 as ii  "&_
										"left join "&ofi&"_extranet.sscont36 as s on s.refcia36=ii.refcia01 and s.patent36=ii.patent01  "&_
										"left join "&ofi&"_extranet.sscont33 as s3 on s3.refcia33=ii.refcia01 and s3.patent33=ii.patent01  "&_
										"where ii.rfccli01 ='"&strFiltroCliente&"' and ii.firmae01 is not null and ii.firmae01<>'' and ii.fecpag01>='"&fi&"' and ii.fecpag01 <= '"&ff&"'  "&_
										"group by ii.refcia01  "&_
									")as Contri on Contri.refcia01=i.refcia01  "&_
									"left join "&ofi&"_extranet.d31refer as r on r.refe31 = i.refcia01  "&_
									"left join "&ofi&"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31  and cta.esta31 <> 'C' "&_
									"left join "&ofi&"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31   "&_
									"left join "&ofi&"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  and ep.tmov21 =dp.tmov21 and ep.tpag21<>3  "&_
									"left join  "&ofi&"_extranet.c21paghe as cp on cp.clav21 = ep.conc21  "&_
									"where  i.rfccli01 in('"&strFiltroCliente&"') and i.fecpag01>='"&fi&"' and i.fecpag01<='"&ff&"' and i.firmae01 is not null and i.firmae01 <>''"&_
									"group by i.refcia01, cta.cgas31,cp.clav21  "&_
							") as T  "&_
							"left join "&ofi&"_extranet.c01refer as rr on rr.refe01=T.REFERENCIA  "&_
							" group by T.REFERENCIA, T.cgas31  "
		'response.write(consulta)
		'response.end()
		Set ConnStr = Server.CreateObject ("ADODB.Connection")
		ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		Set RSops = CreateObject("ADODB.RecordSet")
		Set RSops = ConnStr.Execute(consulta)
		
		IF RSops.BOF = True And RSops.EOF = True Then
			datos = "<br></br><div align=""center""><p  class=""Titulo1"">:: NO SE ENCONTRO INFORMACIÓN PARA LOS PARAMETROS SOLICITADOS ::</div></p></div>"
		else
			info = 	"<table  width = ""2929""  border = ""2"" cellspacing = ""0"" cellpadding = ""0"">" 
				header ="<tr class = ""boton"">" 
			dim ii ,registro
			registro=2
			ii=0
				Do Until RSops.EOF 'Aqui se imprime el reporte
							datos=datos & "<tr>"
							for i=0 to nocolumns
								if ii=0 then 
									header =header&	celdahead(RSops.Fields(i).name) 'imprime el encabezado del reporte
								end if 
								IF RSops.Fields(i).name="INDICADOR DE DESPACHO" then 
								 
									 datos=datos &celdanumero("=SI(O(IZQUIERDA(A"&registro&",3)<>""TOL"", IZQUIERDA(A"&registro&",3)<>""DAI"",IZQUIERDA(A"&registro&",3)<>""PAN""),DIAS.LAB(O"&registro&",R"&registro&")+ENTERO((R"&registro&"-DIASEM(R"&registro&"-6)-O"&registro&"+8)/7)*0.5-1,DIAS.LAB(O"&registro&",R"&registro&")-SI(O(DIASEM(O"&registro&")=1,DIASEM(O"&registro&")=7),0,1))")
								 ELSEIF RSops.Fields(i).name="INDICADOR DE FACTURACIÓN" THEN
									 datos=datos &celdanumero(RSops.Fields.Item(i+1))
								 ELSE
									datos=datos& celdadatos(RSops.Fields.Item(i)) 'imprime los registros del reporte
								 end if 
								
								
							next
							registro=registro+1
							datos=datos & "</tr>"
					Rsops.MoveNext()
					ii=+1
					
				Loop
				header=header &"</tr>"
		end if 
			datos=info & header &datos
		
		retornaMYSQL=datos
	end function 
	

	Function KPI(opera, finicio, ffinal)
	SQL = 	""
	SQL = 	opera & "(IF(MID(T.REFERENCIA,1,3) = 'RKU' OR MID(T.REFERENCIA,1,3) = 'CEG' OR MID(T.REFERENCIA,1,3) = 'SAP' OR MID(T.REFERENCIA,1,3) = 'ALC', " &_
			"(( TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " ) ) -   " &_
			"if( ((DAYOFWEEK( " & finicio & " ) -1) = 6 )   , " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *1.5) - 0.5,  " &_
			"if( (DAYOFWEEK( " & finicio & " ) -1) = 7  ,   " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *1.5) - 1,  " &_
			"if(  ( (DAYOFWEEK( " & finicio & " ) -1)+TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " ) )  = 6, 0.5, " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *1.5) ))) " &_
			" - if( ((DAYOFWEEK( " & finicio & " ) -1) = 5 ), 0.5, 0)), " &_
			"(( TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " ) ) -   " &_
			"if( ((DAYOFWEEK( " & finicio & " ) -1) = 6 )   , " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *2) - 1,  " &_
			"if( (DAYOFWEEK( " & finicio & " ) -1) = 7  ,   " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) *2) - 1,  " &_
			"if(  ( (DAYOFWEEK( " & finicio & " ) -1)+TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " ) )  = 6, 1, " &_
			"(  FLOOR((( (DAYOFWEEK( " & finicio & " ) -1) + (TO_DAYS( " & ffinal & " ) - TO_DAYS( " & finicio & " )) )/ 7)) * 2) ))) " &_
			" - if( ((DAYOFWEEK( " & finicio & " ) -1) = 5 ),1, 0) " &_
			" - if( ((DAYOFWEEK(" & ffinal & ") ) = 7 ),1, 0)))) "
			' Response.Write(SQL)
			' Response.End
	KPI = SQL
End Function

Function retornaOficina(Aduana)
	select case Aduana
				case "VER"
					retornaOficina="rku"
				case "TAM"
					retornaOficina="ceg"
				case "MEX"
					retornaOficina="dai"
				case "MAN"
					retornaOficina="sap"
				case "LZR" 
					retornaOficina="lzr"
				case "TOL"
					retornaOficina="tol"
				end select
end function
%>
<HTML>
	<HEAD>		<TITLE>::.... REPORTE DE INDICADORES .... ::</TITLE>	</HEAD>
	<BODY>		<%=html%>	</BODY>
</HTML>