<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 11">
<%
'Response.Buffer = TRUE
'Response.Charset = "iso-8859-1"
'Response.Addheader "Content-Disposition", "attachment; filename=reporte1.xls"
'Response.ContentType = "application/vnd.ms-excel"

dim oficina,cvesoficina,validacion,codigo

codigo = ""
oficina="RKU"
cvesoficina=""
validacion=""

'oficina=Request.QueryString("ofi")
'tipope=Request.QueryString("tipope")
'det=Request.QueryString("det")
'mes=Request.QueryString("mes")

 strTipoUsuario = Session("GTipoUsuario")
 fechaini = trim(request.Form("txtDateIni"))
 fechafin = trim(request.Form("txtDateFin"))
 strTipoOperaciones = request.Form("rbnTipoDate")
 
 
if not fechaini="" and not fechafin="" then

dim finicio,ffinal

'finicio="2009-07-28"
'ffinal="2009-07-30"

    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    finicio = tmpAnioIni & "-" &tmpMesIni & "-"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    ffinal = tmpAnioFin & "-" &tmpMesFin & "-"& tmpDiaFin




dim strHTML 
strHTML = ""

Server.ScriptTimeOut=100000
%>
<title> Reporte1 Becton Dickinson</title>
</head>
<body>

<%



Dim XML : XML = "<?xml version='1.0'?>" & _
              "<?mso-application progid='Excel.Sheet'?>" & _
               "<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet' " & _
               " xmlns:o='urn:schemas-microsoft-com:office:office' "& _
               " xmlns:x='urn:schemas-microsoft-com:office:excel' " & _
               " xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet' " & _
               " xmlns:html='http://www.w3.org/TR/REC-html40'>" & _
               " <DocumentProperties xmlns='urn:schemas-microsoft-com:office:office'>" & _
               "    <Author>Pedro Bautista</Author>" & _
               "    <LastAuthor>Pedro Bautista</LastAuthor>" & _
               "    <Created>" & ISODate & "</Created>" & _
               "    <Company>Grupo Zego</Company>" & _
               "    <Version>12.00</Version>" & _
               "</DocumentProperties>" & _
               "<OfficeDocumentSettings xmlns='urn:schemas-microsoft-com:office:office'>" & _
               "    <DownloadComponents/>" & _
               "</OfficeDocumentSettings>" & _
               "<ExcelWorkbook xmlns='urn:schemas-microsoft-com:office:excel'>" & _
               "    <WindowHeight>10425</WindowHeight>" & _
               "    <WindowWidth>18015</WindowWidth>" & _
               "    <WindowTopX>240</WindowTopX>" & _
               "    <WindowTopY>60</WindowTopY>" & _
               "    <ProtectStructure>False</ProtectStructure>" & _
               "    <ProtectWindows>False</ProtectWindows>" & _
               "</ExcelWorkbook>" & _
               "<Styles>" & _
               "    <Style ss:ID='Default' ss:Name='Normal'>" & _
               "        <Alignment ss:Vertical='Bottom'/>" & _
               "        <Borders/>" & _
               "        <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='11' ss:Color='#000000'/>" & _
               "        <Interior/>" & _
               "        <NumberFormat/>" & _
               "        <Protection/>" & _			   
               "    </Style>" & _
			   " <Style ss:ID='s21'> " & _
			   " <Interior ss:Color='#330099' ss:Pattern='Solid'></Interior> " & _
			   "  <Font ss:FontName='Calibri' x:Family='Swiss' ss:Size='12' ss:Color='#FFFFFF'/>" & _
			   "  </Style> " & _
               "</Styles>" & _
               "<Worksheet ss:Name='ENC IMPO'>" & _
                        " <Table x:FullColumns='1' " & _
						" x:FullRows='1' ss:DefaultRowHeight='15'>" & _
						 genera_registros("ENC","IMPO",finicio,ffinal) & _
						" </Table>" & _
           "<WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>" & _
           "    <PageSetup>" & _
           "        <Header x:Margin='0.3'/>" & _
           "        <Footer x:Margin='0.3'/>" & _
           "        <PageMargins x:Bottom='0.75' x:Left='0.7' x:Right='0.7' x:Top='0.75'/>" & _
           "    </PageSetup>" & _
           "    <Selected/>" & _
           "    <Panes>" & _
           "        <Pane>" & _
           "           <Number>3</Number>" & _
           "           <ActiveRow>16</ActiveRow>" & _
           "           <ActiveCol>8</ActiveCol>" & _
           "       </Pane>" & _
           "    </Panes>" & _
           "    <ProtectObjects>False</ProtectObjects>" & _
           "    <ProtectScenarios>False</ProtectScenarios>" & _
           "</WorksheetOptions>" & _
           "</Worksheet>" & _
           "<Worksheet ss:Name='FACTURAS IMPO'>" & _
						"    <Table x:FullColumns='1' " & _
						"      x:FullRows='1' ss:DefaultRowHeight='15'>" & _
					   genera_registros("DET","IMPO",finicio,ffinal) & _
						 "</Table>" & _
           " <WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>" & _
           "   <PageSetup>" & _
           "       <Header x:Margin='0.3'/>" & _
           "       <Footer x:Margin='0.3'/>" & _
           "       <PageMargins x:Bottom='0.75' x:Left='0.7' x:Right='0.7' x:Top='0.75'/>" & _
           "    </PageSetup>" & _
           "   <ProtectObjects>False</ProtectObjects>" & _
           "   <ProtectScenarios>False</ProtectScenarios>" & _
           "  </WorksheetOptions>" & _
           " </Worksheet>" & _
           " <Worksheet ss:Name='ENC EXPO'>" & _
						"    <Table x:FullColumns='1' " & _
						"      x:FullRows='1' ss:DefaultRowHeight='15'>" & _
						 genera_registros("ENC","EXPO",finicio,ffinal) & _
						"</Table>" & _
           "  <WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>" & _
           "   <PageSetup>" & _
           "    <Header x:Margin='0.3'/>" & _
           "    <Footer x:Margin='0.3'/>" & _
           "    <PageMargins x:Bottom='0.75' x:Left='0.7' x:Right='0.7' x:Top='0.75'/>" & _
           "   </PageSetup>" & _
           "   <ProtectObjects>False</ProtectObjects>" & _
           "   <ProtectScenarios>False</ProtectScenarios>" & _
           "  </WorksheetOptions>" & _
           " </Worksheet>" & _
		   " <Worksheet ss:Name='FACTURAS EXPO'>" & _
						"    <Table x:FullColumns='1' " & _
						"      x:FullRows='1' ss:DefaultRowHeight='15'>" & _
						 genera_registros("DET","EXPO",finicio,ffinal) & _
						"</Table>" & _
           "  <WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>" & _
           "   <PageSetup>" & _
           "    <Header x:Margin='0.3'/>" & _
           "    <Footer x:Margin='0.3'/>" & _
           "    <PageMargins x:Bottom='0.75' x:Left='0.7' x:Right='0.7' x:Top='0.75'/>" & _
           "   </PageSetup>" & _
           "   <ProtectObjects>False</ProtectObjects>" & _
           "   <ProtectScenarios>False</ProtectScenarios>" & _
           "  </WorksheetOptions>" & _
           " </Worksheet>" & _
           "</Workbook>"
'genera_registros("ENC","EXPO","2009-07-01","2009-07-30") & _		   
' genera_registros("DET","IMPO","2009-07-01","2009-07-30") & _
'genera_registros("ENC","EXPO","2009-07-01","2009-07-30") & _
'genera_registros("DET","EXPO","2009-07-01","2009-07-30") & _


'XML = Replace(XML, "<", "(")
'XML = Replace(XML, ">", ")")
'response.Write(XML)
'response.End()

 With Response
   .Charset = "UTF-8"
   .Clear
   .ContentType = "excel/ms-excel"
   .AddHeader "Content-Disposition","attachment; filename=ReporteBD.xls"
   .AddHeader "Content-Length", Len(XML)
   .Write XML
   .Flush
   .End
 End With
%>

</body>
</html>
<%

end if

'BDM571004IZ6
' 'IFF610526PQ6','IF&610526C95'

'cveped01 <> 'R1'




function genera_registros(det,tipope,finicio,ffinal)
dim c,nparte
nparte=""
codigo = ""
 c=chr(34)
 
 if (det = "ENC" and tipope = "IMPO") then

 codigo=codigo &"<Row>"
 
 genera_html "e","Importa","center"
 genera_html "e","Aduanas","center"
 genera_html "e","Fecha","center"
 genera_html "e","TipoCambio","center"
 genera_html "e","IVA","center"
 genera_html "e","Clave","center"
 genera_html "e","Fletes","center"
 genera_html "e","Seguros","center"
 genera_html "e","Embalaje","center"
 genera_html "e","Otros","center"
 genera_html "e","DTA","center"
 genera_html "e","ValorCom","center"
 genera_html "e","ValorAd","center"
 genera_html "e","Observaciones","center"
 genera_html "e","Consolidado","center"
 genera_html "e","Virtual","center"
 genera_html "e","Prev","center"
 codigo=codigo &"</Row>"

sqlAct= "  select trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Importa, " & _
" i.adusec01 as Aduanas, date_format(i.fecpag01,'%d/%m/%Y') as Fecha,i.tipcam01 as TipoCambio,ifnull(sum(fr.I_iva102 + fr.I_iva202),0) as IVA,i.cveped01 as Clave,  " & _ 
" ifnull(i.fletes01,0) as Fletes,  ifnull(i.segros01,0) as Seguros, ifnull(i.embala01,0) as Embalaje,ifnull(i.incble01,0) as Otros,ifnull(i.i_dta101,0) as DTA,cf3.import36 as DTA2,sum(fr.prepag02) as ValorCom,sum(fr.vaduan02) as ValorAd,  " & _
" if(i.tipped01=1,'S','N') as Consolidado,'' as Virtual,cf2.import36 as Prev,i.anexol01 as Observaciones " & _
"   from rku_extranet.ssdagi01 as i " & _
"      inner join rku_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"        inner join rku_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  " & _
"              left join rku_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '15'  " & _
"                 left join rku_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '1'  " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null  and i.firmae01 <> '' and cveped01 <>'R1'  and cveped01 <> 'R1' and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' " & _
" group by importa " & _
" union all " & _
" select trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Importa, " & _
" i.adusec01 as Aduanas, date_format(i.fecpag01,'%d/%m/%Y') as Fecha,i.tipcam01 as TipoCambio,ifnull(sum(fr.I_iva102 + fr.I_iva202),0) as IVA,i.cveped01 as Clave,  " & _ 
" ifnull(i.fletes01,0) as Fletes,  ifnull(i.segros01,0) as Seguros, ifnull(i.embala01,0) as Embalaje,ifnull(i.incble01,0) as Otros,ifnull(i.i_dta101,0) as DTA,cf3.import36 as DTA2,sum(fr.prepag02) as ValorCom,sum(fr.vaduan02) as ValorAd,  " & _
" if(i.tipped01=1,'S','N') as Consolidado,'' as Virtual,cf2.import36 as Prev,i.anexol01 as Observaciones " & _
"   from dai_extranet.ssdagi01 as i " & _
"      inner join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"        inner join dai_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  " & _
"              left join dai_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '15'  " & _
"                 left join dai_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '1'  " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null  and i.firmae01 <> '' and cveped01 <>'R1'  and cveped01 <> 'R1' and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' " & _
" group by importa"  & _
" union all " & _
" select trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Importa, " & _
" i.adusec01 as Aduanas, date_format(i.fecpag01,'%d/%m/%Y') as Fecha,i.tipcam01 as TipoCambio,ifnull(sum(fr.I_iva102 + fr.I_iva202),0) as IVA,i.cveped01 as Clave,  " & _ 
" ifnull(i.fletes01,0) as Fletes,  ifnull(i.segros01,0) as Seguros, ifnull(i.embala01,0) as Embalaje,ifnull(i.incble01,0) as Otros,ifnull(i.i_dta101,0) as DTA,cf3.import36 as DTA2,sum(fr.prepag02) as ValorCom,sum(fr.vaduan02) as ValorAd,  " & _
" if(i.tipped01=1,'S','N') as Consolidado,'' as Virtual,cf2.import36 as Prev,i.anexol01 as Observaciones " & _
"   from sap_extranet.ssdagi01 as i " & _
"      inner join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"        inner join sap_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  " & _
"              left join sap_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '15'  " & _
"                 left join sap_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '1'  " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null  and i.firmae01 <> '' and cveped01 <>'R1'  and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"'" & _
" group by importa" & _
" union all " & _
" select trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Importa, " & _
" i.adusec01 as Aduanas, date_format(i.fecpag01,'%d/%m/%Y') as Fecha,i.tipcam01 as TipoCambio,ifnull(sum(fr.I_iva102 + fr.I_iva202),0) as IVA,i.cveped01 as Clave,  " & _ 
" ifnull(i.fletes01,0) as Fletes,  ifnull(i.segros01,0) as Seguros, ifnull(i.embala01,0) as Embalaje,ifnull(i.incble01,0) as Otros,ifnull(i.i_dta101,0) as DTA,cf3.import36 as DTA2,sum(fr.prepag02) as ValorCom,sum(fr.vaduan02) as ValorAd,  " & _
" if(i.tipped01=1,'S','N') as Consolidado,'' as Virtual,cf2.import36 as Prev,i.anexol01 as Observaciones " & _
"   from lzr_extranet.ssdagi01 as i " & _
"      inner join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"        inner join lzr_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  " & _
"              left join lzr_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '15'  " & _
"                 left join lzr_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '1'  " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null  and i.firmae01 <> '' and cveped01 <>'R1'  and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"'" & _
" group by importa" 

else
 if (det = "DET" and tipope = "IMPO") then
 codigo=codigo &"<Row>"
 genera_html "e","Importa","center"
 genera_html "e","Factura","center"
 genera_html "e","CodigoP","center"
 genera_html "e","FechaFac","center"
 genera_html "e","FMoneda","center"
 genera_html "e","NumParte","center"
 genera_html "e","DscNP","center"
 genera_html "e","TipoBien","center"
 genera_html "e","FraccionImpo","center"
 genera_html "e","Tasa","center"
 genera_html "e","TipoTasa","center"
 genera_html "e","Unidad","center"
 genera_html "e","Precio","right"
 genera_html "e","Cantidad","center"
 genera_html "e","Conversion","center"
 genera_html "e","Origen","center"
 genera_html "e","Vendedor","center"
 genera_html "e","Fpago" ,"center"
 genera_html "e","Incoterm" ,"center"
 genera_html "e","AcuedoCom" ,"center"
 genera_html "e","TLCAN" ,"center"
 genera_html "e","TLCUEM","center"
 genera_html "e","TLCAELC","center"

 codigo=codigo &"</Row>"
 
 sqlAct= "select  i.refcia01 as referencia,trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Importa,	 " & _
" f.numfac39 as Factura, prv.nompro22 as CodigoP, date_format(f.fecfac39,'%d/%m/%Y') as FechaFac,i.factmo01 as FMoneda,ar.item05 as Numparte, ar.desc05 as DscNP, " & _
"'MP' as TipoBien,concat(concat(concat(concat(substring(fr.fraarn02,1,4),'.'),substring(fr.fraarn02,5,2)),'.'),substring(fr.fraarn02,7,2)) as FraccionImpo ,fr.tasadv02 as Tasa, " & _
" if(ipar2.cveide12 ='TL',if(ipar.comide12 = 'EMU','UE','??'),ifnull(ipar2.cveide12,'TG')) as TipoTasa , " & _
" if(um.descri31 = 'CIENTOS','CNT',if(um.descri31 = 'PIEZA.','PZA',if(um.descri31 = 'MILLAR.','MIL',if(um.descri31 = 'KILOS.','KGS',if(um.descri31 = 'TONELADA.','TON',if(um.descri31 = 'LITRO.','LTS',if(um.descri31 = 'DOCENAS','DOC',um.descri31))))))) as Unidad , " & _
" fr.preuni02  as PrecioMN,fr.cancom02 as CantidadFRACCION,(i.factmo01*i.tipcam01*(ar.vafa05/ar.caco05)) as PrecioMN_cal, round(ar.vafa05/ar.caco05,10) as Precio, ar.caco05 as Cantidad,ar.vafa05,ar.caco05, " & _
" 1 as Conversion,i.cvepod01 as Origen,i.cvepvc01 as Vendedor,cf4.fpagoi36 as Fpago, f.terfac39 as Incoterm , if(concat(ipar2.cveide12,ipar.comide12)='TLEMU','UE',if(concat(ipar2.cveide12,ipar.comide12)='TLCAN','AN',if(concat(ipar2.cveide12,ipar.comide12)='TLUSA','US',concat(ipar2.cveide12,ipar.comide12)))) as AcuerdoCom, " & _
" if(ipar.comide12 in ('CAN'),'S','N') as TLCAN,if(ipar.comide12 in ('EMU'),'S','N') as TLCUEM,if(ipar.comide12 in ('AEL'),'S','N') as TLCAELC  " & _
"   from rku_extranet.ssdagi01 as i " & _
"      inner join rku_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"       left join rku_extranet.ssfact39 as f on f.refcia39 = i.refcia01   " & _
"         left join rku_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  " & _
"           left join rku_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05  " & _
"             left join rku_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01  " & _
"               left join rku_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02  " & _
"                left join rku_extranet.sscont36 as cf4 on cf4.refcia36 = i.refcia01 and cf4.cveimp36 = '6'  " & _
"                  left join rku_extranet.ssipar12 as ipar on ipar.refcia12 = i.refcia01 and ipar.ordfra12 = fr.ordfra02 and ipar.cveide12 = 'TL' and ipar.comide12 in ('CAN','USA','EMU','NOR','CHE','ISL','LIE') " & _
"                  left join rku_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','PS','TL','OC') " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95')  and i.firmae01 is not null and i.firmae01 <> ''  and cveped01 <> 'R1' " & _
" and  i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' " & _
" union all " & _
"select i.refcia01 as referencia,trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Importa,	 " & _
" f.numfac39 as Factura, prv.nompro22 as CodigoP, date_format(f.fecfac39,'%d/%m/%Y') as FechaFac,i.factmo01 as FMoneda,ar.item05 as Numparte, ar.desc05 as DscNP, " & _
"'MP' as TipoBien,concat(concat(concat(concat(substring(fr.fraarn02,1,4),'.'),substring(fr.fraarn02,5,2)),'.'),substring(fr.fraarn02,7,2)) as FraccionImpo ,fr.tasadv02 as Tasa, " & _
" if(ipar2.cveide12 ='TL',if(ipar.comide12 = 'EMU','UE','??'),ifnull(ipar2.cveide12,'TG')) as TipoTasa , " & _
" if(um.descri31 = 'CIENTOS','CNT',if(um.descri31 = 'PIEZA.','PZA',if(um.descri31 = 'MILLAR.','MIL',if(um.descri31 = 'KILOS.','KGS',if(um.descri31 = 'TONELADA.','TON',if(um.descri31 = 'LITRO.','LTS',if(um.descri31 = 'DOCENAS','DOC',um.descri31))))))) as Unidad , " & _
" fr.preuni02  as PrecioMN,fr.cancom02 as CantidadFRACCION,(i.factmo01*i.tipcam01*(ar.vafa05/ar.caco05)) as PrecioMN_cal, round(ar.vafa05/ar.caco05,10) as Precio, ar.caco05 as Cantidad,ar.vafa05,ar.caco05, " & _
" 1 as Conversion,i.cvepod01 as Origen,i.cvepvc01 as Vendedor,cf4.fpagoi36 as Fpago, f.terfac39 as Incoterm , if(concat(ipar2.cveide12,ipar.comide12)='TLEMU','UE',if(concat(ipar2.cveide12,ipar.comide12)='TLCAN','AN',if(concat(ipar2.cveide12,ipar.comide12)='TLUSA','US',concat(ipar2.cveide12,ipar.comide12)))) as AcuerdoCom, " & _
" if(ipar.comide12 in ('CAN'),'S','N') as TLCAN,if(ipar.comide12 in ('EMU'),'S','N') as TLCUEM,if(ipar.comide12 in ('AEL'),'S','N') as TLCAELC  " & _
"   from dai_extranet.ssdagi01 as i " & _
"      inner join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"       left join dai_extranet.ssfact39 as f on f.refcia39 = i.refcia01   " & _
"         left join dai_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  " & _
"           left join dai_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05  " & _
"             left join dai_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01  " & _
"               left join dai_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02  " & _
"                left join dai_extranet.sscont36 as cf4 on cf4.refcia36 = i.refcia01 and cf4.cveimp36 = '6'  " & _
"                  left join dai_extranet.ssipar12 as ipar on ipar.refcia12 = i.refcia01 and ipar.ordfra12 = fr.ordfra02 and ipar.cveide12 = 'TL' and ipar.comide12 in ('CAN','USA','EMU','NOR','CHE','ISL','LIE') " & _
"                  left join dai_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','PS','TL','OC') " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95')  and i.firmae01 is not null and i.firmae01 <> ''  and cveped01 <> 'R1' " & _
"  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' " & _
" union all " & _
"select i.refcia01 as referencia,trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Importa,	 " & _
" f.numfac39 as Factura, prv.nompro22 as CodigoP, date_format(f.fecfac39,'%d/%m/%Y') as FechaFac,i.factmo01 as FMoneda,ar.item05 as Numparte, ar.desc05 as DscNP, " & _
"'MP' as TipoBien,concat(concat(concat(concat(substring(fr.fraarn02,1,4),'.'),substring(fr.fraarn02,5,2)),'.'),substring(fr.fraarn02,7,2)) as FraccionImpo ,fr.tasadv02 as Tasa, " & _
" if(ipar2.cveide12 ='TL',if(ipar.comide12 = 'EMU','UE','??'),ifnull(ipar2.cveide12,'TG')) as TipoTasa , " & _
" if(um.descri31 = 'CIENTOS','CNT',if(um.descri31 = 'PIEZA.','PZA',if(um.descri31 = 'MILLAR.','MIL',if(um.descri31 = 'KILOS.','KGS',if(um.descri31 = 'TONELADA.','TON',if(um.descri31 = 'LITRO.','LTS',if(um.descri31 = 'DOCENAS','DOC',um.descri31))))))) as Unidad , " & _
" fr.preuni02  as PrecioMN,fr.cancom02 as CantidadFRACCION,(i.factmo01*i.tipcam01*(ar.vafa05/ar.caco05)) as PrecioMN_cal, round(ar.vafa05/ar.caco05,10) as Precio, ar.caco05 as Cantidad,ar.vafa05,ar.caco05, " & _
" 1 as Conversion,i.cvepod01 as Origen,i.cvepvc01 as Vendedor,cf4.fpagoi36 as Fpago, f.terfac39 as Incoterm , if(concat(ipar2.cveide12,ipar.comide12)='TLEMU','UE',if(concat(ipar2.cveide12,ipar.comide12)='TLCAN','AN',if(concat(ipar2.cveide12,ipar.comide12)='TLUSA','US',concat(ipar2.cveide12,ipar.comide12)))) as AcuerdoCom, " & _
" if(ipar.comide12 in ('CAN'),'S','N') as TLCAN,if(ipar.comide12 in ('EMU'),'S','N') as TLCUEM,if(ipar.comide12 in ('AEL'),'S','N') as TLCAELC  " & _
"   from sap_extranet.ssdagi01 as i " & _
"      inner join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"       left join sap_extranet.ssfact39 as f on f.refcia39 = i.refcia01   " & _
"         left join sap_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  " & _
"           left join sap_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05  " & _
"             left join sap_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01  " & _
"               left join sap_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02  " & _
"                left join sap_extranet.sscont36 as cf4 on cf4.refcia36 = i.refcia01 and cf4.cveimp36 = '6'  " & _
"                  left join sap_extranet.ssipar12 as ipar on ipar.refcia12 = i.refcia01 and ipar.ordfra12 = fr.ordfra02 and ipar.cveide12 = 'TL' and ipar.comide12 in ('CAN','USA','EMU','NOR','CHE','ISL','LIE') " & _
"                  left join sap_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','PS','TL','OC') " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95')  and i.firmae01 is not null and i.firmae01 <> ''  and cveped01 <> 'R1' " & _
"  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' " & _
" union all " & _
"select i.refcia01 as referencia,trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Importa,	 " & _
" f.numfac39 as Factura, prv.nompro22 as CodigoP, date_format(f.fecfac39,'%d/%m/%Y') as FechaFac,i.factmo01 as FMoneda,ar.item05 as Numparte, ar.desc05 as DscNP, " & _
"'MP' as TipoBien,concat(concat(concat(concat(substring(fr.fraarn02,1,4),'.'),substring(fr.fraarn02,5,2)),'.'),substring(fr.fraarn02,7,2)) as FraccionImpo ,fr.tasadv02 as Tasa, " & _
" if(ipar2.cveide12 ='TL',if(ipar.comide12 = 'EMU','UE','??'),ifnull(ipar2.cveide12,'TG')) as TipoTasa , " & _
" if(um.descri31 = 'CIENTOS','CNT',if(um.descri31 = 'PIEZA.','PZA',if(um.descri31 = 'MILLAR.','MIL',if(um.descri31 = 'KILOS.','KGS',if(um.descri31 = 'TONELADA.','TON',if(um.descri31 = 'LITRO.','LTS',if(um.descri31 = 'DOCENAS','DOC',um.descri31))))))) as Unidad , " & _
" fr.preuni02  as PrecioMN,fr.cancom02 as CantidadFRACCION,(i.factmo01*i.tipcam01*(ar.vafa05/ar.caco05)) as PrecioMN_cal, round(ar.vafa05/ar.caco05,10) as Precio, ar.caco05 as Cantidad,ar.vafa05,ar.caco05, " & _
" 1 as Conversion,i.cvepod01 as Origen,i.cvepvc01 as Vendedor,cf4.fpagoi36 as Fpago, f.terfac39 as Incoterm , if(concat(ipar2.cveide12,ipar.comide12)='TLEMU','UE',if(concat(ipar2.cveide12,ipar.comide12)='TLCAN','AN',if(concat(ipar2.cveide12,ipar.comide12)='TLUSA','US',concat(ipar2.cveide12,ipar.comide12)))) as AcuerdoCom, " & _
" if(ipar.comide12 in ('CAN'),'S','N') as TLCAN,if(ipar.comide12 in ('EMU'),'S','N') as TLCUEM,if(ipar.comide12 in ('AEL'),'S','N') as TLCAELC  " & _
"   from lzr_extranet.ssdagi01 as i " & _
"      inner join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"       left join lzr_extranet.ssfact39 as f on f.refcia39 = i.refcia01   " & _
"         left join lzr_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  " & _
"           left join lzr_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05  " & _
"             left join lzr_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01  " & _
"               left join lzr_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02  " & _
"                left join lzr_extranet.sscont36 as cf4 on cf4.refcia36 = i.refcia01 and cf4.cveimp36 = '6'  " & _
"                  left join lzr_extranet.ssipar12 as ipar on ipar.refcia12 = i.refcia01 and ipar.ordfra12 = fr.ordfra02 and ipar.cveide12 = 'TL' and ipar.comide12 in ('CAN','USA','EMU','NOR','CHE','ISL','LIE') " & _
"                  left join lzr_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','PS','TL','OC') " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95')  and i.firmae01 is not null and i.firmae01 <> ''  and cveped01 <> 'R1' " & _
"  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' "





 else
    if (det = "ENC" and tipope = "EXPO") then

	 codigo=codigo &"<Row>"
	 
 	 genera_html "e","Exporta","center"
	 genera_html "e","Aduanas","center"
	 genera_html "e","Fecha","center"
	 genera_html "e","TipoCambio","center"
	 genera_html "e","IVA","center"
	 genera_html "e","Clave","center"
	 genera_html "e","Fletes","center"
	 genera_html "e","Seguros","center"
	 genera_html "e","Embalaje","center"
	 genera_html "e","Otros","center"
	 genera_html "e","DTA","center"
	 genera_html "e","ValorCom","center"
	 genera_html "e","ValorAd","center"
	 genera_html "e","Observaciones","center"
	 genera_html "e","Consolidado","center"
	 genera_html "e","Virtual","center"
	 genera_html "e","Prev","center"
	 
	 codigo=codigo &"</Row>"

	sqlAct= "  select trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Exporta,  " & _
" i.adusec01 as Aduanas, date_format(i.fecpag01,'%d/%m/%Y') as Fecha,i.tipcam01 as TipoCambio,ifnull(sum(fr.I_iva102 + fr.I_iva202),0) as IVA,i.cveped01 as Clave,  " & _
" ifnull(i.fletes01,0) as Fletes,  ifnull(i.segros01,0) as Seguros, ifnull(i.embala01,0) as Embalaje,ifnull(i.incble01,0) as Otros,ifnull(i.i_dta101,0) as DTA,cf3.import36 as DTA2,sum(fr.prepag02) as ValorCom,(i.tipcam01*i.valfac01) as ValorComBAD,sum(fr.vaduan02) as ValorAd,i.anexol01 as Observaciones, " & _
" if(i.tipped01=1,'S','N') as Consolidado,'' as Virtual,cf2.import36 as Prev " & _
"   from rku_extranet.ssdage01 as i " & _
"      inner join rku_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"        inner join rku_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  " & _
"              left join rku_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '15'  " & _
"                 left join rku_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '1'  " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null and i.firmae01 <> '' and cveped01 <>'R1'  and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' " & _
" group by Exporta " & _
" union all " & _ 
"  select trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Exporta,  " & _
" i.adusec01 as Aduanas, date_format(i.fecpag01,'%d/%m/%Y') as Fecha,i.tipcam01 as TipoCambio,ifnull(sum(fr.I_iva102 + fr.I_iva202),0) as IVA,i.cveped01 as Clave,  " & _
" ifnull(i.fletes01,0) as Fletes,  ifnull(i.segros01,0) as Seguros, ifnull(i.embala01,0) as Embalaje,ifnull(i.incble01,0) as Otros,ifnull(i.i_dta101,0) as DTA,cf3.import36 as DTA2,sum(fr.prepag02) as ValorCom,(i.tipcam01*i.valfac01) as ValorComBAD,sum(fr.vaduan02) as ValorAd,i.anexol01 as Observaciones, " & _
" if(i.tipped01=1,'S','N') as Consolidado,'' as Virtual,cf2.import36 as Prev " & _
"   from dai_extranet.ssdage01 as i " & _
"      inner join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"        inner join dai_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  " & _
"              left join dai_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '15'  " & _
"                 left join dai_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '1'  " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null and i.firmae01 <> '' and cveped01 <>'R1'  and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' " & _
" group by Exporta " & _
" union all " & _ 
"  select trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Exporta,  " & _
" i.adusec01 as Aduanas, date_format(i.fecpag01,'%d/%m/%Y') as Fecha,i.tipcam01 as TipoCambio,ifnull(sum(fr.I_iva102 + fr.I_iva202),0) as IVA,i.cveped01 as Clave,  " & _
" ifnull(i.fletes01,0) as Fletes,  ifnull(i.segros01,0) as Seguros, ifnull(i.embala01,0) as Embalaje,ifnull(i.incble01,0) as Otros,ifnull(i.i_dta101,0) as DTA,cf3.import36 as DTA2,sum(fr.prepag02) as ValorCom,(i.tipcam01*i.valfac01) as ValorComBAD,sum(fr.vaduan02) as ValorAd,i.anexol01 as Observaciones, " & _
" if(i.tipped01=1,'S','N') as Consolidado,'' as Virtual,cf2.import36 as Prev " & _
"   from sap_extranet.ssdage01 as i " & _
"      inner join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"        inner join sap_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  " & _
"              left join sap_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '15'  " & _
"                 left join sap_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '1'  " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null and i.firmae01 <> '' and cveped01 <>'R1'  and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' " & _
" group by Exporta " & _
" union all " & _ 
"  select trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Exporta,  " & _
" i.adusec01 as Aduanas, date_format(i.fecpag01,'%d/%m/%Y') as Fecha,i.tipcam01 as TipoCambio,ifnull(sum(fr.I_iva102 + fr.I_iva202),0) as IVA,i.cveped01 as Clave,  " & _
" ifnull(i.fletes01,0) as Fletes,  ifnull(i.segros01,0) as Seguros, ifnull(i.embala01,0) as Embalaje,ifnull(i.incble01,0) as Otros,ifnull(i.i_dta101,0) as DTA,cf3.import36 as DTA2,sum(fr.prepag02) as ValorCom,(i.tipcam01*i.valfac01) as ValorComBAD,sum(fr.vaduan02) as ValorAd,i.anexol01 as Observaciones, " & _
" if(i.tipped01=1,'S','N') as Consolidado,'' as Virtual,cf2.import36 as Prev " & _
"   from lzr_extranet.ssdage01 as i " & _
"      inner join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"        inner join lzr_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  " & _
"              left join lzr_extranet.sscont36 as cf2 on cf2.refcia36 = i.refcia01 and cf2.cveimp36 = '15'  " & _
"                 left join lzr_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '1'  " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null and i.firmae01 <> '' and cveped01 <>'R1'  and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' " & _
" group by Exporta "



    else
       if (det = "DET" and tipope = "EXPO") then
		 codigo=codigo &"<Row>"

 		 genera_html "e","Exporta","center"
		 genera_html "e","Factura","center"
		 genera_html "e","CodigoC","center"
		 genera_html "e","FechaFac","center"
		 genera_html "e","FMoneda","right"
		 genera_html "e","codigo de producto","right"
		 genera_html "e","DscNP","center"
		 genera_html "e","TipoBien","center"
		 genera_html "e","FraccionExpo","right"
		 genera_html "e","Tasa","right"
		 genera_html "e","TipoTasa","center"
		 genera_html "e","Unidad","right"
		 genera_html "e","Precio","right"
		 genera_html "e","Cantidad","right"
		 genera_html "e","Conversion","right"
		 genera_html "e","Destino","center"
		 genera_html "e","Comprador","center"
		 genera_html "e","Fpago" ,"right"
		 genera_html "e","Incoterm" ,"center"
		 
 		 codigo=codigo &"</Row>"
		 
 'TipoBien ar.tpmerc05
		 sqlAct= "select i.refcia01 as referencia,trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Exporta,	 " & _
" f.numfac39 as Factura, prv.nompro22 as CodigoC, date_format(f.fecfac39,'%d/%m/%Y') as FechaFac,i.factmo01 as FMoneda,ar.item05 as Numparte, ar.desc05 as DscNP, " & _
"'MP' as TipoBien,concat(concat(concat(concat(substring(fr.fraarn02,1,4),'.'),substring(fr.fraarn02,5,2)),'.'),substring(fr.fraarn02,7,2)) as FraccionExpo ,fr.tasadv02 as Tasa, " & _
" if(ipar2.cveide12 ='TL',if(ipar.comide12 = 'EMU','UE','??'),ifnull(ipar2.cveide12,'TG')) as TipoTasa , " & _
" if(um.descri31 = 'CIENTOS','CNT',if(um.descri31 = 'PIEZA.','PZA',if(um.descri31 = 'MILLAR.','MIL',if(um.descri31 = 'KILOS.','KGS',if(um.descri31 = 'TONELADA.','TON',if(um.descri31 = 'LITRO.','LTS',if(um.descri31 = 'DOCENAS','DOC',um.descri31))))))) as Unidad , " & _
" fr.preuni02  as PrecioMN,fr.cancom02 as CantidadFRACCION, round(ar.vafa05/ar.caco05,10) as Precio, ar.caco05 as Cantidad,ar.vafa05,ar.caco05, " & _
" 1 as Conversion,i.cvepod01 as Destino,i.cvepvc01 as Comprador,cf4.fpagoi36 as Fpago, f.terfac39 as Incoterm,fr.ordfra02,ar.agru05  " & _
"   from rku_extranet.ssdage01 as i " & _
"    inner join rku_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"      left join rku_extranet.ssfact39 as f on f.refcia39 = i.refcia01   " & _
"       left join rku_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  " & _
"         left join rku_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"          left join rku_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01  " & _
"            left join rku_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02  " & _
"              left join rku_extranet.sscont36 as cf4 on cf4.refcia36 = i.refcia01 and cf4.cveimp36 = '6'  " & _
"                left join rku_extranet.ssipar12 as ipar on ipar.refcia12 = i.refcia01 and ipar.ordfra12 = fr.ordfra02 and ipar.cveide12 = 'TL'  " & _
"                 left join rku_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','PS','TL','OC') " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null  and i.firmae01 <> ''  and cveped01 <>'R1' and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"'" & _
" union all " & _
" select i.refcia01 as referencia,trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Exporta,	 " & _
" f.numfac39 as Factura, prv.nompro22 as CodigoC, date_format(f.fecfac39,'%d/%m/%Y') as FechaFac,i.factmo01 as FMoneda,ar.item05 as Numparte, ar.desc05 as DscNP, " & _
"'MP' as TipoBien,concat(concat(concat(concat(substring(fr.fraarn02,1,4),'.'),substring(fr.fraarn02,5,2)),'.'),substring(fr.fraarn02,7,2)) as FraccionExpo ,fr.tasadv02 as Tasa, " & _
" if(ipar2.cveide12 ='TL',if(ipar.comide12 = 'EMU','UE','??'),ifnull(ipar2.cveide12,'TG')) as TipoTasa , " & _
" if(um.descri31 = 'CIENTOS','CNT',if(um.descri31 = 'PIEZA.','PZA',if(um.descri31 = 'MILLAR.','MIL',if(um.descri31 = 'KILOS.','KGS',if(um.descri31 = 'TONELADA.','TON',if(um.descri31 = 'LITRO.','LTS',if(um.descri31 = 'DOCENAS','DOC',um.descri31))))))) as Unidad , " & _
" fr.preuni02  as PrecioMN,fr.cancom02 as CantidadFRACCION, round(ar.vafa05/ar.caco05,10) as Precio, ar.caco05 as Cantidad,ar.vafa05,ar.caco05, " & _
" 1 as Conversion,i.cvepod01 as Destino,i.cvepvc01 as Comprador,cf4.fpagoi36 as Fpago, f.terfac39 as Incoterm,fr.ordfra02,ar.agru05  " & _
"   from dai_extranet.ssdage01 as i " & _
"    inner join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"      left join dai_extranet.ssfact39 as f on f.refcia39 = i.refcia01   " & _
"       left join dai_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  " & _
"         left join dai_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"          left join dai_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01  " & _
"            left join dai_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02  " & _
"              left join dai_extranet.sscont36 as cf4 on cf4.refcia36 = i.refcia01 and cf4.cveimp36 = '6'  " & _
"                left join dai_extranet.ssipar12 as ipar on ipar.refcia12 = i.refcia01 and ipar.ordfra12 = fr.ordfra02 and ipar.cveide12 = 'TL'  " & _
"                 left join dai_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','PS','TL','OC') " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null  and i.firmae01 <> ''  and cveped01 <>'R1' and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"'" & _
" union all " & _
" select i.refcia01 as referencia,trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Exporta,	 " & _
" f.numfac39 as Factura, prv.nompro22 as CodigoC, date_format(f.fecfac39,'%d/%m/%Y') as FechaFac,i.factmo01 as FMoneda,ar.item05 as Numparte, ar.desc05 as DscNP, " & _
"'MP' as TipoBien,concat(concat(concat(concat(substring(fr.fraarn02,1,4),'.'),substring(fr.fraarn02,5,2)),'.'),substring(fr.fraarn02,7,2)) as FraccionExpo ,fr.tasadv02 as Tasa, " & _
" if(ipar2.cveide12 ='TL',if(ipar.comide12 = 'EMU','UE','??'),ifnull(ipar2.cveide12,'TG')) as TipoTasa , " & _
" if(um.descri31 = 'CIENTOS','CNT',if(um.descri31 = 'PIEZA.','PZA',if(um.descri31 = 'MILLAR.','MIL',if(um.descri31 = 'KILOS.','KGS',if(um.descri31 = 'TONELADA.','TON',if(um.descri31 = 'LITRO.','LTS',if(um.descri31 = 'DOCENAS','DOC',um.descri31))))))) as Unidad , " & _
" fr.preuni02  as PrecioMN,fr.cancom02 as CantidadFRACCION, round(ar.vafa05/ar.caco05,10) as Precio, ar.caco05 as Cantidad,ar.vafa05,ar.caco05, " & _
" 1 as Conversion,i.cvepod01 as Destino,i.cvepvc01 as Comprador,cf4.fpagoi36 as Fpago, f.terfac39 as Incoterm,fr.ordfra02,ar.agru05  " & _
"   from sap_extranet.ssdage01 as i " & _
"    inner join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"      left join sap_extranet.ssfact39 as f on f.refcia39 = i.refcia01   " & _
"       left join sap_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  " & _
"         left join sap_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"          left join sap_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01  " & _
"            left join sap_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02  " & _
"              left join sap_extranet.sscont36 as cf4 on cf4.refcia36 = i.refcia01 and cf4.cveimp36 = '6'  " & _
"                left join sap_extranet.ssipar12 as ipar on ipar.refcia12 = i.refcia01 and ipar.ordfra12 = fr.ordfra02 and ipar.cveide12 = 'TL'  " & _
"                 left join sap_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','PS','TL','OC') " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null  and i.firmae01 <> ''  and cveped01 <>'R1' and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"'" & _
" union all " & _
" select i.refcia01 as referencia,trim(concat(concat(concat(concat(concat(concat(date_format(i.fecpag01,'%y'),'-'),i.adusec01),'-'),i.patent01),'-'),i.numped01)) as Exporta,	 " & _
" f.numfac39 as Factura, prv.nompro22 as CodigoC, date_format(f.fecfac39,'%d/%m/%Y') as FechaFac,i.factmo01 as FMoneda,ar.item05 as Numparte, ar.desc05 as DscNP, " & _
"'MP' as TipoBien,concat(concat(concat(concat(substring(fr.fraarn02,1,4),'.'),substring(fr.fraarn02,5,2)),'.'),substring(fr.fraarn02,7,2)) as FraccionExpo ,fr.tasadv02 as Tasa, " & _
" if(ipar2.cveide12 ='TL',if(ipar.comide12 = 'EMU','UE','??'),ifnull(ipar2.cveide12,'TG')) as TipoTasa , " & _
" if(um.descri31 = 'CIENTOS','CNT',if(um.descri31 = 'PIEZA.','PZA',if(um.descri31 = 'MILLAR.','MIL',if(um.descri31 = 'KILOS.','KGS',if(um.descri31 = 'TONELADA.','TON',if(um.descri31 = 'LITRO.','LTS',if(um.descri31 = 'DOCENAS','DOC',um.descri31))))))) as Unidad , " & _
" fr.preuni02  as PrecioMN,fr.cancom02 as CantidadFRACCION, round(ar.vafa05/ar.caco05,10) as Precio, ar.caco05 as Cantidad,ar.vafa05,ar.caco05, " & _
" 1 as Conversion,i.cvepod01 as Destino,i.cvepvc01 as Comprador,cf4.fpagoi36 as Fpago, f.terfac39 as Incoterm,fr.ordfra02,ar.agru05  " & _
"   from lzr_extranet.ssdage01 as i " & _
"    inner join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
"      left join lzr_extranet.ssfact39 as f on f.refcia39 = i.refcia01   " & _
"       left join lzr_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  " & _
"         left join lzr_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"          left join lzr_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01  " & _
"            left join lzr_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02  " & _
"              left join lzr_extranet.sscont36 as cf4 on cf4.refcia36 = i.refcia01 and cf4.cveimp36 = '6'  " & _
"                left join lzr_extranet.ssipar12 as ipar on ipar.refcia12 = i.refcia01 and ipar.ordfra12 = fr.ordfra02 and ipar.cveide12 = 'TL'  " & _
"                 left join lzr_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','PS','TL','OC') " & _
"   where cc.rfccli18 in ('IFF610526PQ6','IF&610526C95') and i.firmae01 is not null  and i.firmae01 <> ''  and cveped01 <>'R1' and cveped01 <> 'R1'  and i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"'"

	   end if
    end if
 end if 
end if



Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE="&oficina&"_extranet; UID=pedrobm; PWD=123; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()


if (det = "ENC" and tipope = "IMPO") then
while not act2.eof
'response.Write("<tr align="&c&"center"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")
codigo=codigo &"<Row>"
 genera_html "d",act2.fields("Importa").value,"center"
 genera_html "d",act2.fields("Aduanas").value,"center"
 genera_html "d",act2.fields("Fecha").value,"center"
 genera_html "d",act2.fields("tipoCambio").value,"right"
 genera_html "d",act2.fields("IVA").value,"right"

 genera_html "d",act2.fields("Clave").value,"right"
 genera_html "d",act2.fields("Fletes").value,"right"
 genera_html "d",act2.fields("Seguros").value,"right"
 genera_html "d",act2.fields("Embalaje").value,"right"
 genera_html "d",act2.fields("Otros").value,"right"
 genera_html "d",act2.fields("DTA").value,"right"
 genera_html "d",act2.fields("ValorCom").value,"right"
 genera_html "d",act2.fields("ValorAd").value,"right"
 genera_html "d",act2.fields("Observaciones").value,"center"
 genera_html "d",act2.fields("Consolidado").value,"center"
 genera_html "d",act2.fields("Virtual").value,"center"
  genera_html "d",act2.fields("Prev").value,"right"


' response.Write("</tr>")
codigo=codigo &"</Row>"
 act2.movenext()
wend
else
 if (det = "DET" and tipope = "IMPO") then
 while not act2.eof
'response.Write("<tr align="&c&"right"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")
		 codigo=codigo &"<Row>"

			if(existeNumeroParte(mid(act2.fields("referencia").value,1,3),act2.fields("referencia").value,act2.fields("Numparte").value)=true)then
  			 nparte= act2.fields("Numparte").value
			else
   			 nparte= "(false) "& act2.fields("Numparte").value
			end if

   genera_html "d",act2.fields("Importa").value,"right"
   genera_html "d",act2.fields("Factura").value,"right"
   genera_html "d",codigoProveedor(act2.fields("CodigoP").value),"right"
   genera_html "d",act2.fields("FechaFac").value,"center"
   genera_html "d",act2.fields("FMoneda").value,"right"
   'genera_html "d",act2.fields("Numparte").value,"right"
   genera_html "d",nparte,"right"
   genera_html "d",act2.fields("DscNP").value,"center"
   genera_html "d",act2.fields("TipoBien").value,"center"
   genera_html "d",act2.fields("FraccionImpo").value,"right"
   genera_html "d",act2.fields("Tasa").value,"right"
   genera_html "d",act2.fields("TipoTasa").value,"center"
   genera_html "d",act2.fields("Unidad").value,"center"
   genera_html "d",act2.fields("Precio").value,"right"
   genera_html "d",act2.fields("Cantidad").value,"right"
   genera_html "d",act2.fields("Conversion").value,"center"
   genera_html "d",act2.fields("Origen").value,"center"
   genera_html "d",act2.fields("Vendedor").value,"center"
   genera_html "d",act2.fields("FPago").value,"right"
   genera_html "d",act2.fields("Incoterm").value,"center"
   genera_html "d",act2.fields("AcuerdoCom").value,"center"
   genera_html "d",act2.fields("TLCAN").value,"center"
   genera_html "d",act2.fields("TLCUEM").value,"center"
   genera_html "d",act2.fields("TLCAELC").value,"center"
 'response.Write("</tr>")
		 codigo=codigo &"</Row>"
 act2.movenext()
 wend
 
 else
    if (det = "ENC" and tipope = "EXPO") then
		while not act2.eof
		'response.Write("<tr align="&c&"center"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")
				 codigo=codigo &"<Row>"
		 genera_html "d",act2.fields("Exporta").value,"center"
		 genera_html "d",act2.fields("Aduanas").value,"center"
		 genera_html "d",act2.fields("Fecha").value,"center"
		 genera_html "d",act2.fields("tipoCambio").value,"right"
		 genera_html "d",act2.fields("IVA").value,"right"

		 genera_html "d",act2.fields("Clave").value,"center"
		 genera_html "d",act2.fields("Fletes").value,"right"
		 genera_html "d",act2.fields("Seguros").value,"right"
		 genera_html "d",act2.fields("Embalaje").value,"right"
		 genera_html "d",act2.fields("Otros").value,"right"
		 genera_html "d",act2.fields("DTA").value,"right"
		 genera_html "d",act2.fields("ValorCom").value,"right"
		 genera_html "d",act2.fields("ValorAd").value,"right"
		 genera_html "d",act2.fields("Observaciones").value,"center"
		 genera_html "d",act2.fields("Consolidado").value,"center"
		 genera_html "d",act2.fields("Virtual").value,"center"
		 genera_html "d",act2.fields("Prev").value,"right"

			 codigo=codigo &"</Row>"
	    'response.Write("</tr>")
	
	    act2.movenext()
	   wend
    else
       if (det = "DET" and tipope = "EXPO") then
	      while not act2.eof
			'response.Write("<tr align="&c&"center"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")
		 codigo=codigo &"<Row>"
			if(existeNumeroParte(mid(act2.fields("referencia").value,1,3),act2.fields("referencia").value,act2.fields("Numparte").value)=true)then
  			 nparte= act2.fields("Numparte").value
			else
   			 nparte= "(false) "& act2.fields("Numparte").value
			end if
			
			   genera_html "d",act2.fields("Exporta").value,"center"
			   genera_html "d",act2.fields("Factura").value,"right"
'			   genera_html "d",act2.fields("CodigoC").value,"center"
			   genera_html "d",codigoCliente(act2.fields("CodigoC").value),"right"  '& "-" & act2.fields("CodigoC").value,"right"
			   genera_html "d",act2.fields("FechaFac").value,"center"
			   genera_html "d",act2.fields("FMoneda").value,"center"
			   'genera_html "d",act2.fields("Numparte").value,"right"
			   genera_html "d",nparte,"right"
			   genera_html "d",act2.fields("DscNP").value,"center"
			   genera_html "d",act2.fields("TipoBien").value,"center"
			   genera_html "d",act2.fields("FraccionExpo").value,"center"
			   genera_html "d",act2.fields("Tasa").value,"right"
			   genera_html "d",act2.fields("TipoTasa").value,"center"
			   genera_html "d",act2.fields("Unidad").value,"center"
			   genera_html "d",act2.fields("Precio").value,"right"
			   genera_html "d",act2.fields("Cantidad").value,"right"
			   genera_html "d",act2.fields("Conversion").value,"center"
			   genera_html "d",act2.fields("Destino").value,"center"
			   genera_html "d",act2.fields("Comprador").value,"center"
			   genera_html "d",act2.fields("FPago").value,"right"
			   genera_html "d",act2.fields("Incoterm").value,"center"
			 'response.Write("</tr>")
				 codigo=codigo &"</Row>"
			 act2.movenext()
			 wend
			 
	   end if
    end if
 end if
end if

codigo = Replace((codigo), "á", "a")
codigo = Replace((codigo), "é", "e")
codigo = Replace((codigo), "í", "i")
codigo = Replace((codigo), "ó", "u")
codigo = Replace((codigo), "ú", "u")
'codigo = Replace((codigo), "Á", "A")
'codigo = Replace((codigo), "É", "E")
'codigo = Replace((codigo), "Í", "I")
'codigo = Replace((codigo), "Ó", "O")
'codigo = Replace((codigo), "Ú", "U")
'codigo = Replace((codigo), "´", "")


'codigo = Replace(codigo, ">", ")")
'codigo = Replace(codigo, "<", "(")
'response.Write(codigo)
'response.End()
genera_registros = codigo
end function

sub genera_html(tipo,valor,alineacion)

 if(tipo = "e")then
  'response.Write("<td width="&c&"100"&c&" align="&c&alineacion&c&" nowrap bgcolor="&c&"#CCFF99"&c&"><div align="&c&alineacion&c&"><strong><em><font size="&c&"2"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></em></strong></div></td>")
  codigo = codigo&"<Cell ss:StyleID='s21'><Data ss:Type='String'>"&valor&"</Data></Cell>" 
 else 
 '  response.Write("<td align="&c&alineacion&c&" nowrap><div align="&c&alineacion&c&"><font color="&c&"#000000"&c&" size="&c&"1"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></div></td>")
  codigo = codigo&"<Cell><Data ss:Type='String'>"&valor&"</Data></Cell>"
 end if


end sub

function regresa_fecha_cuenta_gastos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
sqlAct="select r.refe31,min(cta.fech31) as fech31 from e31cgast as cta, d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
" r.refe31 = '"&referencia&"'  "&_
" group by r.refe31"

Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = cadena_de_conexion()
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE="&oficina&"_extranet; UID=pedrobm; PWD=123; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
if not(act2.eof) then
 regresa_fecha_cuenta_gastos =act2.fields("fech31").value
else
  regresa_fecha_cuenta_gastos =valor
   end if
end function

function codigoProveedor(desc)
dim res,desc2
res = "no"
Path=Server.MapPath("catprobd.xls")
desc2=replace(desc," ","%")
Set ConexionBD = Server.CreateObject("ADODB.Connection") 
ConexionBD.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & Path
Set rsVac = Server.CreateObject("ADODB.Recordset") 
rsVac.Open "Select * From A1:B50 where descpro like '" & desc2 & "'", ConexionBD,3,3 

if not(rsVac.eof)then
 res =rsVac.fields("cvepro") '&","&rsVac.fields("descpro")
else
 res = desc
end if 

codigoProveedor = res
end function

function codigoCliente(desc)
dim res,desc2
res = "no"
Path=Server.MapPath("catclibd.xls")
desc2=replace(desc," ","%")
Set ConexionBD = Server.CreateObject("ADODB.Connection") 
ConexionBD.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & Path
Set rsVac = Server.CreateObject("ADODB.Recordset") 
rsVac.Open "Select * From A1:B35 where desccli like '" & desc2 & "'", ConexionBD,3,3 
 'response.write("Select * From A1:B15 where desccli like '%" & desc2 & "%'")
 'response.End()
if not(rsVac.eof)then
 res =rsVac.fields("cvecli")  '&","&rsVac.fields("desccli")

else
 res = desc
end if 

codigoCliente = res
end function


function existeNumeroParte(oficina,referencia,item05)
dim val
val = false

if oficina = "ALC" then
oficina = "LZR"
end if


 c=chr(34)

	sqlAct="select item05 as campo from d05artic as ar "&_
	" where ar.refe05 = '"&referencia&"' and item05='"& item05 &"' and item05 in ('13006'," & _
"'30791'," & _
"'31391'," & _
"'31933'," & _
"'470821030R'," & _
"'471137050RPDC'," & _
"'471180030RPDC'," & _
"'471262030RPDC'," & _
"'47314002'," & _
"'8002795'," & _
"'8003955'," & _
"'8011660'," & _
"'8011762'," & _
"'8011974'," & _
"'8012196'," & _
"'8012414'," & _
"'8080410'," & _
"'8015930'," & _
"'8016350'," & _
"'8016905'," & _
"'8019237'," & _
"'8020619'," & _
"'8034744'," & _
"'8034758'," & _
"'8034759'," & _
"'8080598'," & _
"'8082530'," & _
"'8083378'," & _
"'8083485'," & _
"'8084438'," & _
"'8085259'," & _
"'8304539'," & _
"'POLIPROPILENO'," & _
"'R00920010PDC'," & _
"'R00944080'," & _
"'R00945080PDC'," & _
"'R01003081PDC'," & _
"'R01008050'," & _
"'R01013010PDC'," & _
"'R01047080PDC'," & _
"'R01048080PDC'," & _
"'R01051080PDC'," & _
"'R01052080PDC'," & _
"'R01681080PDC'," & _
"'U03428080PDC'," & _
"'8304669'," & _
"'8034770'," & _
"'10254') "
	
	
	'Set act2= Server.CreateObject("ADODB.Recordset")

	'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE="&oficina&"_extranet; UID=pedrobm; PWD=123; OPTION=16427"
	
	'act2.ActiveConnection = conn12
	'act2.Source = sqlAct
	'act2.cursortype=0
	'act2.cursorlocation=2
	'act2.locktype=1
	'act2.open()
	'if not(act2.eof) then
	'  existeNumeroParte = true 
	'else
	'  existeNumeroParte = false 
	'end if
	existeNumeroParte = true 

end function

%>

