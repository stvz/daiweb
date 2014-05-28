<!-- #include virtual="/PortalMySQL/Extranet/ext-Asp/Clases/cConexion.asp" -->

<META HTTP-EQUIV="Content-Type" CONTENT="text/html"; charset="utf-8">
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<%
'																										 		'
'																												'
' ---------------------------------------        EXTRACOSTOS       --------------------------------------------	'
'																												'
' 
'se manda a llamar de la siguiente forma:
'http://10.66.1.9/portalmysql/extranet/ext-asp/reportes/reporte_erp_unilever_99_XC.asp?finicio=13/09/2010&ffinal=28/09/2010&tipope=i&det=


Response.Buffer = TRUE
response.Charset = "utf-8"
'Response.Addheader "Content-Disposition", "attachment; filename=BookletUNILEVER_EXTRACOSTOS_.xls"'
'Response.ContentType = "application/vnd.ms-excel"


dim oficina,cvesoficina,validacion

cvesoficina = ""
validacion = ""
num=0

oficina 	= "RKU"
tipope		= Request.QueryString("tipope")
det			= Request.QueryString("det")
fechaini	= Request.QueryString("finicio")
fechafin	= Request.QueryString("ffinal")


if not fechaini="" and not fechafin="" then

    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    finicio = tmpAnioIni & "-" &tmpMesIni & "-"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    ffinal = tmpAnioFin & "-" &tmpMesFin & "-"& tmpDiaFin

	' Response.Write("fecha inicio = " & finicio & "<br>fecha final = " & ffinal & "<br>")
	' Response.End

	dim orden(50)
	dim subrefaux,subref,bgcolor,strHTML
	subrefaux=""
	subref=""
	bgcolor="#FFFFFF"
	strHTML = ""

	Server.ScriptTimeOut=10000000
%>
<!-- head -->
<title> Reporte1.. </title>
<link href="reporte_erp_unilever_99_XC.css" type="text/css" rel="stylesheet">
</head>
<body>
<table x:str border=0 cellpadding=0 cellspacing=0 width=12637 style='border-collapse:
 collapse;table-layout:fixed;width:9479pt'>
 <col width=125 style='mso-width-source:userset;mso-width-alt:4571;width:94pt'>
 <col width=100 span=2 style='mso-width-source:userset;mso-width-alt:3657;
 width:75pt'>
 <col class=xl26 width=226 style='mso-width-source:userset;mso-width-alt:8265;
 width:170pt'>
 <col width=100 span=4 style='mso-width-source:userset;mso-width-alt:3657;
 width:75pt'>
 <col width=212 style='mso-width-source:userset;mso-width-alt:7753;width:159pt'>
 <col width=100 span=11 style='mso-width-source:userset;mso-width-alt:3657;
 width:75pt'>
 <col class=xl26 width=214 style='mso-width-source:userset;mso-width-alt:7826;
 width:161pt'>
 <col width=100 span=8 style='mso-width-source:userset;mso-width-alt:3657;
 width:75pt'>
 <col class=xl26 width=100 style='mso-width-source:userset;mso-width-alt:3657;
 width:75pt'>
 <col width=100 span=67 style='mso-width-source:userset;mso-width-alt:3657;
 width:75pt'>
 <col width=80 span=32 style='width:60pt'>
 <tr class=xl27 height=52 style='height:39.0pt'>
<% genera_registros det,tipope %>
</table>
</body>
</html>
<%
	'c12 Este end if viene de la cabezera
end if

sub genera_registros(det,tipope)
	dim c
	c=chr(34)
%>
</tr>
<%
sqlAct= "select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'MARITIMO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01)  as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" IFNULL(cf6.import36,0) as '75', " & _
" IFNULL(cf1.import36,0) as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" IFNULL(cf3.import36,0)  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from rku_extranet.ssdagi01 as i  " & _
"   left join rku_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join rku_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" INNER join rku_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" INNER join rku_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join rku_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join rku_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join rku_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join rku_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join rku_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join rku_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join rku_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join rku_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join rku_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join rku_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join rku_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and  i.fecpag01 >=  '"& finicio &"' and i.fecpag01 <= '"& ffinal &"' " & _
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
" select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'AEREO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01)  as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" IFNULL(cf6.import36,0) as '75', " & _
" IFNULL(cf1.import36,0) as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" IFNULL(cf3.import36,0)  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from dai_extranet.ssdagi01 as i  " & _
"   left join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join dai_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" INNER join dai_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" INNER join dai_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join dai_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join dai_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join dai_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join dai_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join dai_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join dai_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join dai_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join dai_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join dai_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join dai_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join dai_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and   cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
" select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'MARITIMO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01)  as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _   
" 'unilever' as '74', " & _
" IFNULL(cf6.import36,0) as '75', " & _
" IFNULL(cf1.import36,0) as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" IFNULL(cf3.import36,0)  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from sap_extranet.ssdagi01 as i  " & _
"   left join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join sap_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" INNER join sap_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" INNER join sap_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join sap_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join sap_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join sap_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join sap_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join sap_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join sap_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join sap_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join sap_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join sap_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join sap_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join sap_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and  cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
" select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'AEREO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01) as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" IFNULL(cf6.import36,0) as '75', " & _
" IFNULL(cf1.import36,0) as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" IFNULL(cf3.import36,0)  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from tol_extranet.ssdagi01 as i  " & _
"   left join tol_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join tol_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" INNER join tol_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" INNER join tol_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join tol_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join tol_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join tol_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join tol_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join tol_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join tol_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join tol_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join tol_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join tol_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join tol_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join tol_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> '' and  cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
" select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'MARITIMO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01) as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" IFNULL(cf6.import36,0) as '75', " & _
" IFNULL(cf1.import36,0) as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" IFNULL(cf3.import36,0)  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from ceg_extranet.ssdagi01 as i  " & _
"   left join ceg_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join ceg_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" INNER join  ceg_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" INNER join ceg_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join ceg_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join ceg_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join ceg_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join ceg_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join ceg_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join ceg_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join ceg_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join ceg_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join ceg_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join ceg_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join ceg_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and  cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
" select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'MARITIMO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01) as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" IFNULL(cf6.import36,0) as '75', " & _
" IFNULL(cf1.import36,0) as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" IFNULL(cf3.import36,0)  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82' , " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from lzr_extranet.ssdagi01 as i  " & _
"   left join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join lzr_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" INNER join lzr_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" INNER join lzr_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join lzr_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join lzr_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join lzr_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join lzr_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join lzr_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join lzr_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join lzr_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join lzr_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join lzr_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join lzr_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join lzr_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and  cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 "

'Response.write(sqlAct)
'response.end()

'c15 Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

dim ref,refAux
dim cambio,rcli 
cambio = 1
rcli=""
refAux = ""

while not act2.eof
response.Write("<tr align="&c&"center"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")
ref = act2.fields("29").value
if (ref <> refAux)then
 if cambio = 1 then 
	 cambio = 2
	 else 
	 if cambio = 2 then
	    cambio = 1
	 end if
 end if
end if
 
if cambio = 1 then
	bgcolor="#D7ECF4"
else
	bgcolor="#FFFFFF"
end if 
  
genera_html "d",retornaDivision(act2.fields("1").value,act2.fields("71").value),"center"  'DIVISION 
genera_html "d","","center"  'IMPORTANCIA
genera_html "d","","center"  'CATEGORIA
genera_html "d",act2.fields("4").value,"center"  'Nombre del Material
genera_html "d","","center"  'Codigo SAP
genera_html "d","","center"  'Clase de Producto
genera_html "d","","center"  'Contrato Marco
genera_html "d","","center"  'Código SAP Proveedor
genera_html "d",act2.fields("9").value,"center"  'Proveedor

rcli =act2.fields("11").value
genera_html "d",retornaCuenta(rcli),"center"  'Cuenta 
genera_html "d",retornaCECO(rcli),"center"  'CECO

if(act2.fields("177").value<>"")then
	PuertEmb= retornaCampoPuertoEmb(act2.fields("177").value,act2.fields("17").value,"cvepai01",mid(act2.fields("29").value,1,3))	
else
	PuertEmb= ""	
end if 
 
genera_html "d",act2.fields("12").value,"center"  'ODC
genera_html "d","","center"  'No IE
genera_html "d",act2.fields("14").value,"center"  'País de Origen 
genera_html "d",PuertEmb,"center"  'País de Procedencia
genera_html "d",retornaRegion(act2.fields("14").value),"center"  'Region
genera_html "d",act2.fields("177").value,"center" 'PTO./CD DE ORIGEN
genera_html "d",act2.fields("18").value,"center"  'TAX ID/ RFC
genera_html "d",act2.fields("19").value,"center"  'Factura
genera_html "d",act2.fields("20").value,"center"  'Fecha de Factura

genera_html "d",retornaIMPORTADOR(act2.fields("21").value,mid(act2.fields("29").value,1,3)),"center"  'IMPORTADOR
genera_html "d",act2.fields("22").value,"center"  'Cantidad 
genera_html "d",act2.fields("23").value,"center"  'Unidad de Medida
genera_html "d",act2.fields("24").value,"center"  'Incoterms
genera_html "d",act2.fields("25").value,"center"  'Tipo de Transporte
genera_html "d",retornaAduana(act2.fields("26").value),"center"  'Aduana
genera_html "d",retornaAgenteAduanal(act2.fields("27").value),"center"  'Agente Aduanal
genera_html "d",act2.fields("28").value,"center"  'Patente Agente Aduanal
genera_html "d",act2.fields("29").value,"center"  'No. De Trafico
 
ref = act2.fields("29").value
if (ref <> refAux)then
  refAux=ref

  
'Lote 2
genera_html "d",retornaCampoContenedores(act2.fields("29").value,"marc01",mid(act2.fields("29").value,1,3)),"center"  'No de Contenedor 
genera_html "d",act2.fields("31").value,"center"  'No. Pedimento
genera_html "d",act2.fields("32").value,"center"  'Fecha Pedimento
genera_html "d",act2.fields("33").value,"center"  'Mes
genera_html "d",act2.fields("34").value,"center"  'No.Semana
genera_html "d",act2.fields("35").value,"center"  'Cantidad de Operaciones 
genera_html "d",retornaCantContenedores(act2.fields("29").value,"'ISO','CON'",mid(act2.fields("29").value,1,3)),"center"  'Cantidad de Contenedores
genera_html "d",retornaCantContenedores(act2.fields("29").value,"'BUL','CAJ','BID','PAL'",mid(act2.fields("29").value,1,3)),"center"  'PALLETS/BULTOS 
genera_html "d",retornaTipoContenedores(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  'TIPO DE CONTENEDOR/ CAJA
genera_html "d",act2.fields("39").value,"center"  'Fecha Factura
genera_html "d",act2.fields("40").value,"center"  'Fecha BL
genera_html "d",act2.fields("41").value,"center"  'Fecha de arribo a la aduana
genera_html "d",act2.fields("42").value,"center"  'Fecha Desaduanamiento
genera_html "d","","center"  'KPI Desaduanamiento
genera_html "d","","center"  'KPI lead TIME
genera_html "d","","center"  'TARGET TIME
genera_html "d","","center"  'Fecha de arribo a la planta
genera_html "d","","center"  'BW ? 
genera_html "d","","center"  ' BW $  
genera_html "d","","center"  'PL
genera_html "d","","center"  'TT
genera_html "d","","center"  'PR
genera_html "d","","center"  'AA
genera_html "d","","center"  'CO
genera_html "d","","center"  'AL
genera_html "d","","center"  'NUMERO DE EMBARQUE
genera_html "d","","center"  'IDOT 
genera_html "d",retornaCampoCtaGastos(act2.fields("29").value,"cgas31",mid(act2.fields("29").value,1,3)),"center"  'No. CTA DE GASTOS
genera_html "d",regresa_fecha_cuenta_gastos(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  'Fecha C.Gastos
genera_html "d",regresa_tipo_Cgastos(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"    'Tipo C.Gastos
genera_html "d",retornaMontoAnticipo(act2.fields("29").value,"ANT",mid(act2.fields("29").value,1,3)),"center"  ' Monto de Anticipo 

Subref = act2.fields("81").value
if (ordenOcupado(Subref,act2.fields("29").value) = False)then
	ocuparOrd(Subref)
	
	'-----
	genera_html "d",act2.fields("59").value,"center"  ' Precio Pagado / valor comercial 
	genera_html "d",act2.fields("60").value,"center"  '  Valor comercial USD 
	genera_html "d",act2.fields("61").value,"center"  ' VALOR FLETES INTERNACIONAL M.N. 
	genera_html "d",act2.fields("62").value,"center"  ' SEGUROS 
	genera_html "d",act2.fields("63").value,"center"  ' OTROS INCREMENTABLES 
	genera_html "d",act2.fields("64").value,"center"  ' VALOR ADUANA M.N. 
	genera_html "d",act2.fields("65").value,"center"  ' T.C. 
	genera_html "d",act2.fields("66").value,"center"  ' VALOR ADUANA  DLLS 

	if( act2.fields("25").value = "MARITIMO") then
		genera_html "d",act2.fields("67").value,"center"  ' VALOR FLETES AEREO DLLS 
		genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
		genera_html "d",act2.fields("69").value,"center"  ' VALOR FLETES MARITIMO DLLS. 
	else
		genera_html "d",act2.fields("69").value,"center"  ' VALOR FLETES AEREO DLLS 
		genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
		genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
	end if
	
	genera_html "d","","center"  ' SAVING 
	genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
	genera_html "d",act2.fields("72").value,"center"  'ARANCEL %
	genera_html "d",act2.fields("73").value,"center"  'ARANCEL PREFERENCIAL 
	genera_html "d",retornaECI(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3)),"center"  ' MONTO DE RECUPERACION $  
	'-----
      
	genera_html "d",act2.fields("761").value,"center"  ' ADV FRACC. $ 

	genera_html "d",act2.fields("76").value,"center"  ' DTA $ 
	genera_html "d",act2.fields("77").value,"center"  'IVA %
	genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
else  
    '-----
	genera_html "d","","center"  ' Precio Pagado / valor comercial 
	genera_html "d","","center"  '  Valor comercial USD 
	genera_html "d",act2.fields("61").value,"center"  ' VALOR FLETES INTERNACIONAL M.N. 
	genera_html "d",act2.fields("62").value,"center"  ' SEGUROS 
	genera_html "d",act2.fields("63").value,"center"  ' OTROS INCREMENTABLES 
	genera_html "d","","center"  ' VALOR ADUANA M.N. 
	genera_html "d",act2.fields("65").value,"center"  ' T.C. 
	genera_html "d","","center"  ' VALOR ADUANA  DLLS 

	if( act2.fields("25").value = "MARITIMO") then
		genera_html "d",act2.fields("67").value,"center"  ' VALOR FLETES AEREO DLLS 
		genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
		genera_html "d",act2.fields("69").value,"center"  ' VALOR FLETES MARITIMO DLLS. 
	else
		genera_html "d",act2.fields("69").value,"center"  ' VALOR FLETES AEREO DLLS 
		genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
		genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
	end if

	genera_html "d","","center"  ' SAVING 
	genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
	genera_html "d",act2.fields("72").value,"center"  'ARANCEL %
	genera_html "d",act2.fields("73").value,"center"  'ARANCEL PREFERENCIAL 
	genera_html "d",retornaECI(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3)),"center"  ' MONTO DE RECUPERACION $  
   '-----
  
	genera_html "d","","center"  ' ADV FRACC. $ 
	genera_html "d",act2.fields("76").value,"center"  ' DTA $ 
	genera_html "d","","center"  'IVA %
	genera_html "d","","center"  ' IVA FRACC. $ 
end if

'/Lote2 c21 aqui termina el lote2
    
'Lote 1 c22
	
	genera_html "d",act2.fields("79").value,"center"  ' PREVAL. 
	genera_html "d",sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  ' TOTAL IMPUESTOS 
	'Response.Write( cdbl( sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3)))/act2.fields("65").value)
	genera_html "d",(CDbl((sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))))/act2.fields("65").value),"center"  ' Total Impuestos USD  
	'Response.Write(sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))&"aki" & act2.fields("65").value)
	'response.end()
	genera_html "d","N/A","center"  ' GTOS. ADUANA USD(SOLO FRONTERA)
	genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"DEMORAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' DEMORAS 
	genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' ESTADIAS 
	 
	if ucase(mid(act2.fields("29").value,1,3)) ="RKU" then
		genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' MANIOBRAS  
		genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' ALMACENAJES 
	else
		genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' MANIOBRAS
		genera_html "d","?","center"  ' ALMACENAJES 
	end if
	
	dim TPH,DEM,EST,ALMMAN,MAN
	TPH=0.0
	DEM=0.0
	EST=0.0
	ALMAN=0.0
	MAN=0.0
	TPH	  = retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3))
	DEM	  = retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"DEMORAS"),"I",mid(act2.fields("29").value,1,3))
	EST	  = retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"I",mid(act2.fields("29").value,1,3))
	ALMAN = retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"I",mid(act2.fields("29").value,1,3))
	
	if ucase(mid(act2.fields("29").value,1,3)) ="RKU" then
		MAN=retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3))
	end if
	
	if(revisaImpuestosFacturados( act2.fields("29").value,"I",mid(act2.fields("29").value,1,3)) <> 0 )then
		STIMP=sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))
		STIVA=sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3))
end if
	
	

	genera_html "d",TPH-DEM-EST-ALMAN-MAN-STIMP-STIVA,"center"  ' OTROS 


	'genera_html "d",retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3))-sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))-sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  ' TOTAL GASTOS DIVERSOS 
	genera_html "d",retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3))-STIMP-STIVA,"center"  ' TOTAL GASTOS DIVERSOS 
	
	'genera_html "d",((retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3))-sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))-sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3))) /act2.fields("65").value),"center"  ' TOTAL GASTOS DIVERSOS USD 
	genera_html "d",(CDbl(retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3))-STIMP-STIVA)/act2.fields("65").value),"center"  ' TOTAL GASTOS DIVERSOS USD 

	genera_html "d",retornaHonorarios(act2.fields("29").value,"chon31",mid(act2.fields("29").value,1,3)),"center"  ' HONORARIOS AG AD. $ 
	
'/Lote 1 termina el lote1
 else
'bgcolor="#D7ECF4"


 
 'Lote 2
genera_html "d","","center"  'No de Contenedor 
genera_html "d",act2.fields("31").value,"center"  'No. Pedimento
genera_html "d",act2.fields("32").value,"center"  'Fecha Pedimento
genera_html "d",act2.fields("33").value,"center"  'Mes
genera_html "d",act2.fields("34").value,"center"  'No.Semana
genera_html "d",act2.fields("35").value,"center"  'Cantidad de Operaciones 
genera_html "d","","center"  'Cantidad de Contenedores
genera_html "d","","center"  'PALLETS/BULTOS 
genera_html "d","","center"  'TIPO DE CONTENEDOR/ CAJA
genera_html "d",act2.fields("39").value,"center"  'Fecha Factura
genera_html "d",act2.fields("40").value,"center"  'Fecha BL
genera_html "d",act2.fields("41").value,"center"  'Fecha de arribo a la aduana
genera_html "d",act2.fields("42").value,"center"  'Fecha Desaduanamiento
 genera_html "d","","center"  'KPI Desaduanamiento
 genera_html "d","","center"  'KPI lead TIME
 genera_html "d","","center"  'TARGET TIME
 genera_html "d","","center"  'Fecha de arribo a la planta
 genera_html "d","","center"  'BW ? 
 genera_html "d","","center"  ' BW $  
 genera_html "d","","center"  'PL
 genera_html "d","","center"  'TT
 genera_html "d","","center"  'PR
 genera_html "d","","center"  'AA
 genera_html "d","","center"  'CO
 genera_html "d","","center"  'AL
 genera_html "d","","center"  'NUMERO DE EMBARQUE
 genera_html "d","","center"  'IDOT 
genera_html "d",retornaCampoCtaGastos(act2.fields("29").value,"cgas31",mid(act2.fields("29").value,1,3)),"center"  'No. CTA DE GASTOS
genera_html "d",regresa_fecha_cuenta_gastos(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"      'Fecha Cta de Gastos
genera_html "d","","center"    'Tipo C.Gastos
genera_html "d","","center"  ' Monto de Anticipo 
'genera_html "d",act2.fields("59").value,"center"  ' Precio Pagado / valor comercial 
'genera_html "d",act2.fields("60").value,"center"  '  Valor comercial USD 
'genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL M.N. 
'genera_html "d","","center"  ' SEGUROS 
'genera_html "d","","center"  ' OTROS INCREMENTABLES 
'genera_html "d",act2.fields("64").value,"center"  ' VALOR ADUANA M.N. 
'genera_html "d",act2.fields("65").value,"center"  ' T.C. 
'genera_html "d",act2.fields("66").value,"center"  ' VALOR ADUANA  DLLS 
'genera_html "d","","center"  ' VALOR FLETES AEREO DLLS 
'genera_html "d","","center"  ' VALOR FLETES TERRESTRE DLLS. 
'genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
'genera_html "d",act2.fields("70").value,"center"  ' SAVING 
'genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
'genera_html "d",act2.fields("72").value,"center"  'ARANCEL %
'genera_html "d",act2.fields("73").value,"center"  'ARANCEL PREFERENCIAL 
'genera_html "d",act2.fields("74").value,"center"  ' MONTO DE RECUPERACION $  



Subref = act2.fields("81").value
 if (ordenOcupado(Subref,act2.fields("29").value) = False)then
   ocuparOrd(Subref)
   
   '---------
    genera_html "d",act2.fields("59").value,"center"  ' Precio Pagado / valor comercial 
	genera_html "d",act2.fields("60").value,"center"  '  Valor comercial USD 
	genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL M.N. 
	genera_html "d","","center"  ' SEGUROS 
	genera_html "d","","center"  ' OTROS INCREMENTABLES 
	genera_html "d",act2.fields("64").value,"center"  ' VALOR ADUANA M.N. 
	genera_html "d",act2.fields("65").value,"center"  ' T.C. 
	genera_html "d",act2.fields("66").value,"center"  ' VALOR ADUANA  DLLS 
	genera_html "d","","center"  ' VALOR FLETES AEREO DLLS 
	genera_html "d","","center"  ' VALOR FLETES TERRESTRE DLLS. 
	genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
	genera_html "d","","center"  ' SAVING 
	genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
	genera_html "d",act2.fields("72").value,"center"  'ARANCEL %
	genera_html "d",act2.fields("73").value,"center"  'ARANCEL PREFERENCIAL 
	genera_html "d","","center"  ' MONTO DE RECUPERACION $  
   '---------
   
   
   genera_html "d",act2.fields("761").value,"center"  ' ADV FRACC. $ 
  
   genera_html "d","","center"  ' DTA $ 
   genera_html "d",act2.fields("77").value,"center"  'IVA %
   genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
 else
 
   '---------
    genera_html "d","","center"  ' Precio Pagado / valor comercial 
	genera_html "d","","center"  '  Valor comercial USD 
	genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL M.N. 
	genera_html "d","","center"  ' SEGUROS 
	genera_html "d","","center"  ' OTROS INCREMENTABLES 
	genera_html "d","","center"  ' VALOR ADUANA M.N. 
	genera_html "d",act2.fields("65").value,"center"  ' T.C. 
	genera_html "d","","center"  ' VALOR ADUANA  DLLS 
	genera_html "d","","center"  ' VALOR FLETES AEREO DLLS 
	genera_html "d","","center"  ' VALOR FLETES TERRESTRE DLLS. 
	genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
	genera_html "d",act2.fields("70").value,"center"  ' SAVING 
	genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
	genera_html "d",act2.fields("72").value,"center"  'ARANCEL %
	genera_html "d",act2.fields("73").value,"center"  'ARANCEL PREFERENCIAL 
	genera_html "d","","center"  ' MONTO DE RECUPERACION $  
   '---------
 
   genera_html "d","","center"  ' ADV FRACC. $ 
   
   genera_html "d","","center"  ' DTA $ 
   genera_html "d","","center"  'IVA %
   genera_html "d","","center"  ' IVA FRACC. $ 
 end if

'/Lote2 termina lote2
 

'Lote 1
   'genera_html "d","","center"  ' DTA $ 
   'genera_html "d",act2.fields("77").value,"center"  'IVA %
   'genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
   genera_html "d","","center"  ' PREVAL. 
   genera_html "d","","center"  ' TOTAL IMPUESTOS 
   genera_html "d","","center"  ' Total Impuestos USD  
   
    genera_html "d","N/A","center"  ' GTOS. ADUANA USD(SOLO FRONTERA) 
	genera_html "d","","center"  ' DEMORAS 
	genera_html "d","","center"  ' ESTADIAS 
	genera_html "d","","center"  ' MANIOBRAS  
	genera_html "d","","center"  ' ALMACENAJES 
	genera_html "d","","center"  ' OTROS 
	genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS 
	
	genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS USD 
	genera_html "d","","center"  ' HONORARIOS AG AD. $ 
'/Lote 1
 end if
 
'genera_html "d","","center"  ' FLETE TARIFA NORMAL  
'genera_html "d","?","center"  'TRANSPORTISTA
'genera_html "d","?","center"  ' COSTO EXTRA EN FLETE 
'genera_html "d","?","center"  ' TOTAL FLETE NAL 
genera_html "d","","center"  ' Total Gastos Indirectos  USD 
genera_html "d","","center"  'INLAND
genera_html "d","","center"  'Impacto Valor Factura. 
genera_html "d","","center"  ' VAL FACT USD 
genera_html "d",retornaPlantaEntrega(act2.fields("1").value),"center"  'PLANTA DE ENTREGA

genera_html "d",act2.fields("81").value,"center"  'ORDEN FRACCION
genera_html "d",act2.fields("82").value,"center"  'CUENTA ORDEN FRACCION

genera_html "d",act2.fields("Item05").value,"center"  'item
genera_html "d",act2.fields("firmita").value,"center"  'firmita
genera_html "d",act2.fields("sDescTransp").value,"center"  'descripcion del transporte

 response.Write("</tr>")
num=num+1
 act2.movenext()
 
wend

end sub

sub genera_html(tipo,valor,alineacion)
 if(tipo = "e")then
  'response.Write("<td width="&c&"100"&c&" align="&c&alineacion&c&" nowrap bgcolor="&c&"#CCFF99"&c&"><div align="&c&alineacion&c&"><strong><em><font size="&c&"2"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></em></strong></div></td>")
   response.Write("<td width="&c&"100"&c&" align="&c&alineacion&c&" nowrap bgcolor="&c&"#CCFF99"&c&"><div align="&c&alineacion&c&"><strong><em><font size="&c&"2"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></em></strong></div></td>")
 else 
  'response.Write("<td align="&c&alineacion&c&" nowrap background="&c&bgcolor&c&"><div align="&c&alineacion&c&"><font color="&c&"#000000"&c&" size="&c&"1"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></div></td>")
  'response.Write("<td align="&c&alineacion&c&" nowrap background="&c&bgcolor&c&"><div align="&c&alineacion&c&"><font color="&c&"#000000"&c&" size="&c&"1"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></div></td>")
   if bgcolor ="#D7ECF4" then
     response.Write("<td align="&c&alineacion&c&" class=xl73>"&valor&"</td>")
   else
     response.Write("<td align="&c&alineacion&c&" class=xl78>"&valor&"</td>")
   end if
 '"#D7ECF4"
 end if

end sub

function revisaImpuestosFacturados(referencia,tipoop,oficina)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select count(i.refcia01) as Ref " & _
" from "& oficina &"_extranet.ssdag" & tipoop &"01 as i  " & _
"  inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  " & _
"     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 " & _
"          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 " & _
"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
"    where  i.firmae01 <> ''  and cta.esta31 <> 'C'  and i.refcia01 = '"& referencia &"' and ep.conc21 = 1"

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = cadena_de_conexion()
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="& oficina &"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

if not(act2.eof) then
 revisaImpuestosFacturados =act2.fields("Ref").value
else
  revisaImpuestosFacturados = nothing
end if


end function


function regresa_fecha_cuenta_gastos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select max(date_format(cta.fech31,'%d/%m/%Y')) as fech31 from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
" r.refe31 = '"&referencia&"' and cta.esta31 <> 'C' "

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

if not(act2.eof) then
 regresa_fecha_cuenta_gastos =act2.fields("fech31").value
else
  regresa_fecha_cuenta_gastos =valor
   end if
end function


function regresa_tipo_Cgastos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select if(COUNT(cta.cgas31) > 1, 'COMPLEMENTARIA','NORMAL')  as tipo from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
" r.refe31 = '"&referencia&"'  "&_
"  and cta.esta31 <> 'C' "

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

if not(act2.eof) then
 regresa_tipo_Cgastos =act2.fields("tipo").value
else
  regresa_tipo_Cgastos =valor
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

function retornaCampoCtaGastos(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
 sqlAct = "select r."& campo &" as campo from "&oficina&"_extranet.e31cgast as cta " &_
 " inner join  "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
 " where  r.refe31 = '"& referencia &"' and cta.esta31 <> 'C' "

 'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCampoCtaGastos = valor
 else
  retornaCampoCtaGastos =valor
 end if
end function


function retornaMontoAnticipo(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
 sqlAct = " select sum(dm.mont11) as campo " & _
			" from "&oficina&"_extranet.d11movim as dm " & _
			" where dm.refe11='"& referencia&"' and dm.conc11 = '"&campo&"' "
	'and dm.cgas11 <> ''"

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaMontoAnticipo = valor
 else
  retornaMontoAnticipo =valor
 end if
end function


function retornaIMPORTADOR(clave,oficina)
ON ERROR RESUME NEXT
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

 sqlAct2 = "select c.nomcli18 as campo from "&oficina&"_extranet.ssclie18 as c where c.cvecli18 = "&clave

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct2,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct2
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

if err.number <> 0 then
	retornaIMPORTADOR = err.description
else

 if not(act2.eof) then
 
 
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaIMPORTADOR = valor
 else
  retornaIMPORTADOR =valor
 end if
end if   
 
end function




function retornaCampoPuertoEmb(pto,val,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
'sqlAct="select "& campo &" as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"'  "
sqlAct="SELECT "& campo &" as campo FROM "&oficina&"_extranet.c01ptoemb where cvepto01 ="& val &" and nompto01 like '"& pto &"%'"

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCampoPuertoEmb = valor
 else
  retornaCampoPuertoEmb =valor
 end if
end function

function retornaAgenteAduanal(valor)
dim val
val = ""
if(valor = "3921")then
 val = "Luis E. de la Cruz Reyes"
else
if(valor = "3210")then
 val = "Rolando Reyes Kuri"
else
if(valor = "3945")then
 val = "Jesús Gómez Reyes"
else
if(valor = "3931")then
 val = "Sergio Alvarez Ramírez"
else
if(valor = "3044")then
 val = "Carlos Humberto Zesati Andrade"
else
val =""
end if
end if
end if
end if
end if

retornaAgenteAduanal = val
end function


function retornaAduana(valor)
dim val
val = ""
if(valor = "470")then
 val = "México"
else
if(valor = "430")then
 val = "Veracruz"
else
if(valor = "810")then
 val = "Altamira"
else
if(valor = "160")then
 val = "Manzanillo"
else
if(valor = "510")then
 val = "Lázaro Cardenas"
else
if(valor = "650")then
 val = "Toluca"
else
val =""
end if
end if
end if
end if
end if
end if

retornaAduana = val
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


function retornaHonorarios(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

sqlAct=" select cta."&campo&" as campo from "&oficina&"_extranet.e31cgast as cta  " & _
       " inner join "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
       " where  r.refe31 = '"& referencia &"' and cta.esta31 = 'I' "

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaHonorarios = valor
 else
  retornaHonorarios =valor
 end if
end function

function retornaConceptosPH(oficina,topico)
dim cad
cad = "NA"

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if oficina = "SAP" then

if topico = "ALMACENAJES-MANIOBRAS" then
  cad= "2,3,63,65,86,111,111,112,142,174,181,183,183,186,188,189,190,196,196,208,209,210,211,212,214,216,218,234,251,256,258,265,269,284,286,287,288,289,290,291,292,293,294,295,296,297,298,299,300,301,303,304,305,307,309,323,331,336,351"
end if
if topico = "DEMORAS" then
  cad= "6,14,46,63,129,156,352"
end if
if topico = "ESTADIAS" then
  cad="144"
end if
if topico = "OTROS" then
  cad="306,313,350,351,352"
end if


else 
  if oficina = "CEG" then
  
    if topico = "ALMACENAJES-MANIOBRAS" then
     cad= "2,4,59,77,100,223,223,235,235,241"
    end if
	if topico = "DEMORAS" then
	 cad= "11,48,99,150"
	end if
	if topico = "ESTADIAS" then
	 cad="NA"
	end if
	if topico = "OTROS" then
	 cad="239"
    end if
  
  else 
     if oficina = "TOL" then
	 
	    if topico = "ALMACENAJES-MANIOBRAS" then
		 cad= "2,2,10,127,128"
		end if
		if topico = "DEMORAS" then
		 cad= "79"
		end if
		if topico = "ESTADIAS" then
		 cad="123"
		end if
		if topico = "OTROS" then
		 cad="NA"
		end if
		
     else 
	   if oficina = "LZR" then
	         if topico = "ALMACENAJES-MANIOBRAS" then
			 cad= "4,78,115,116,119,125,160,167,167,203,230,230,244,297,312"
			end if
			if topico = "DEMORAS" then
			 cad= "11"
			end if
			if topico = "ESTADIAS" then
			 cad="77"
			end if
			if topico = "OTROS" then
			 cad="NA"
			end if
       else 
	       if oficina = "RKU" then
		          if topico = "ALMACENAJES-MANIOBRAS" then
					 'cad= "2,4,78,115,116,119,125,160,167,167,203,230,230,244,297,304,311,312,313,359,359"
					 cad="4"
					end if
					if topico = "DEMORAS" then
					 cad= "11,310,376"
					end if
					if topico = "ESTADIAS" then
					 cad="77"
					end if
					if topico = "MANIOBRAS" then
					 cad="2"
					end if
           else 
		       if oficina = "DAI" then
			           if topico = "ALMACENAJES-MANIOBRAS" then
						 cad= "2,2,10,93,127,128,155,163,163,166,170"
						end if
						if topico = "DEMORAS" then
						 cad= "79,171"
						end if
						if topico = "ESTADIAS" then
						 cad="123"
						end if
						if topico = "OTROS" then
						 cad="NA"
						end if
               else 
   			          cad = "NE"
               end if
           end if
       end if
     end if
  end if
end if
retornaConceptosPH = cad
end function


function retornaECI(referencia,tipope,oficina)
dim c,valor
 c=chr(34)
 valor=0

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 


sqlAct =" select c.import36 as Campo ,c.cveimp36,c.refcia36  from "& oficina &"_extranet.ssdag"& tipope &"01 as i " & _
		"  inner  join  "& oficina &"_extranet.sscont36 as c on i.refcia01 = c.refcia36 " & _
		"    where c.refcia36 = '"& referencia &"' and c.cveimp36 = '18' and i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3')"

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

	 if not(act2.eof) then
	 valor = act2.fields("Campo").value
	
	  retornaECI = valor
	 else
	  retornaECI = valor
	 end if

end function


function retornaPagosHechos(referencia,conceptos,tipope,oficina)
dim c,valor
 c=chr(34)
 valor=0

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then


'sqlAct =" select i.refcia01 as Ref,sum(dp.mont21*if(ep.deha21 = 'C',-1,1)) as Importe " & _
'		" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
'		" inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 " & _
'		" inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and ep.fech21 = dp.fech21 and ep.conc21 in ("&conceptos&") and ep.esta21 <> 'S' and ep.esta21 <> 'C' " & _
'		" where i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.refcia01 = '"&referencia&"'  and i.firmae01 <> ''   group by i.refcia01 "

sqlAct="select i.refcia01 as Ref, r.cgas31,ep.conc21,ep.piva21,CAST(ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as Decimal(20,4)) as Importe, cp.desc21 " & _
" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
"  inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  " & _
"     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 " & _
"          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 " & _
"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.tmov21 =dp.tmov21 " & _
"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
"    where  i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3')  and i.firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in ("&conceptos&") and i.refcia01 = '"&referencia&"'  group by Ref,cgas31,conc21"

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("Importe").value
 'act2.movenext()
 'while not act2.eof
 '  valor = valor&", "&act2.fields("Importe").value
 '  act2.movenext()
 'wend
  retornaPagosHechos = valor
 else
  retornaPagosHechos = valor
 end if
 else
   retornaPagosHechos =valor
 end if

end function


function retornaTOTALPagosHechos(referencia,tipope,oficina)
dim c,valor
 c=chr(34)
 valor=0

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then

sqlAct =" select i.refcia01 as Ref,CAST(ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0) AS decimal(20,4)) as Importe " & _
		" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
		" inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 " & _
		" inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  and ep.tmov21 =dp.tmov21 " & _
		" where i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3')  and i.refcia01 = '"&referencia&"'  and i.firmae01 <> ''  group by i.refcia01 "

'response.Write(sqlAct)
'response.End()

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("Importe").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("Importe").value
   act2.movenext()
 wend
  retornaTOTALPagosHechos = valor
 else
  retornaTOTALPagosHechos = valor
 end if
 else
   retornaTOTALPagosHechos =0
 end if

end function

function retornaCampoContenedores(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select "& campo &" as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"'  "

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCampoContenedores = valor
 else
  retornaCampoContenedores =valor
 end if
end function


function retornaCantContenedores(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
sqlAct="select count(*) as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"' and clas01 in ("&campo&") "

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCantContenedores = valor
 else
  retornaCantContenedores =valor
 end if
end function


function DatosContenedor(Tipo)
dim val
val = Tipo
if(Tipo = "1")then
  'val="Contenedor Estandar 40 pulg (Standard Container 40 pulg)"
   val="20"
end if

if(Tipo = "2")then
'  val="Contenedor Estandar 40 pulg (Standard Container 40 pulg)"
    val="40"
end if

if(Tipo = "3")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="40 HighCube"
end if

if(Tipo = "4")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="20 Hardtop"
end if

if(Tipo = "5")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="40 Hardtop"
end if


if(Tipo = "6")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="20 OpenTop"
end if

if(Tipo = "7")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="40 OpenTop"
end if


if(Tipo = "17")then
'val="Contenedor Refrigerante Cubo Alto 17 pulg (High Cube Refrigerated Container 40 pulg"
val="17 HighCube"
end if
DatosContenedor = val
end function

function retornaTipoContenedores(referencia,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select distinct cn4.tipcon40 as campo from "& oficina &"_extranet.sscont40 as cn4 where cn4.refcia40 = '"& referencia &"' "

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = DatosContenedor(act2.fields("campo").value)
 act2.movenext()
 while not act2.eof
   valor = DatosContenedor(valor) &", "& DatosContenedor(act2.fields("campo").value)
   act2.movenext()
 wend
  retornaTipoContenedores = valor
 else
  retornaTipoContenedores =valor
 end if
end function

function revisaFraccion(fraccion)
dim val
val ="(pendiente)"
fraccion = trim(fraccion)
if (fraccion = "90230001" or _
 fraccion = "49111099" or _
 fraccion = "48219099" or _
 fraccion = "39235001" or _
 fraccion = "34011101" or _
 fraccion = "33072001" or _
 fraccion = "33059099" or _
 fraccion = "33051001" or _
 fraccion = "33049999" or _
  fraccion = "76129099" or _
  fraccion = "33079099" or _
      fraccion = "39202099" or _
	  fraccion = "85234099" or _
	   fraccion = "49119999" or _
	   	 fraccion = "09081001" or _
	 fraccion = "13023902" or _
	 fraccion = "21069003" or _
	 fraccion = "39209999" or _
	 fraccion = "39233099" or _
	 fraccion = "84212199" or _
	 fraccion = "33071001" or _
	 fraccion = "39231001" or _
	 fraccion = "49019906" or _	 
	     fraccion = "33029099") then
 val ="HPC"
 else
	 if (fraccion = "21039099" or _
	 fraccion = "12119001" or _
 	 fraccion = "9023001" or _
	 fraccion = "07103001" or _
	 fraccion = "07108099" or _
	 fraccion = "18069099" or _
 	 fraccion = "21069099" or _
	 fraccion = "21041001" or _
 	 fraccion = "07129099" or _
 	 fraccion = "11061001" or _
	 fraccion = "15119099" or _
	 fraccion = "17029099" or _
	 fraccion = "39219099" or _
	 fraccion = "09023001") then
	 val ="FOODS"
	 else
	 val = "ERROR (CD) "&fraccion
	 end if
 end if

revisaFraccion = val

end function

function retornaDivision(clave,fraccion)
dim val,res
val= ""
res= "(pendiente)"

if clave = "11000" then
  res = revisaFraccion(fraccion)
  val = res '"Centro de Distribución"
end if

if clave = "11001" then
val = "ICE CREAM" '"Planta Helados"
end if
if clave = "11002" then
val = "FOODS"
end if
if clave = "11003" then
val = "HPC" '"Planta HPC"
end if
if clave = "11004" then
val = "Todas"
end if

'OTRAS
if clave = "13000" then
val = "Todas"
end if

if clave = "14000" then
val = "Todas"
end if



retornaDivision = val
end function


function retornaPlantaEntrega(clave)
dim val
val= ""
if clave = "11000" then
val = "CDU"
end if
if clave = "11001" then
val = "TULTITLAN"
end if
if clave = "11002" then
val = "LERMA"
end if
if clave = "11003" then
val = "CIVAC"
end if
if clave = "11004" then
val = "ESPECIALES"
end if

'OTRAS
if clave = "13000" then
val = "Todas"
end if
if clave = "14000" then
val = "Todas"
end if


retornaPlantaEntrega = val
end function


function sumaTotalImpuestos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="0"
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct=" select ifnull(sum(import36),0) as campo from "& oficina &"_extranet.sscont36 as cf1 " & _
       " where cf1.cveimp36 in ('1', '6','15')   and refcia36 = '"&referencia&"' "
      ' " where cf1.cveimp36 in ('1','3','6','15')   and refcia36 = '"&referencia&"' "

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor &", "& act2.fields("campo").value
   act2.movenext()
 wend
  sumaTotalImpuestos = valor
 else
  sumaTotalImpuestos =valor
 end if

'ADV/IGI+DTA+IVA+PREVAL

end function
function sumaTotalIVA(referencia,oficina)
dim c,valor
 c=chr(34)
 valor=0
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct=" select ifnull(sum(import36),0) as campo from "& oficina &"_extranet.sscont36 as cf1 " & _
       " where cf1.cveimp36 in ('3')   and refcia36 = '"&referencia&"' "
      ' " where cf1.cveimp36 in ('1','3','6','15')   and refcia36 = '"&referencia&"' "

'Llamada a la conexion de MySQL mediante la clase cConexion en el archivo cConexion.asp
Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2

'Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

'act2.ActiveConnection = conn12
'act2.Source = sqlAct
'act2.cursortype=0
'act2.cursorlocation=2
'act2.locktype=1
'act2.open()

 if not(act2.eof) then
 valor = act2.fields("campo").value
 'act2.movenext()
 'while not act2.eof
   'valor = valor &", "& act2.fields("campo").value
 '  act2.movenext()
 'wend
  sumaTotalIVA = valor
 else
  sumaTotalIVA =valor
 end if

'ADV/IGI+DTA+IVA+PREVAL

end function


function retornaRegion(clave)
dim val
val= "("& clave &")"

if clave = "MEX" then
val = "LATINOAMERICA"
end if

if clave = "TUR" then
val = "ASIA"
end if

if clave = "ARG" then
val = "SUDAMERICA"
end if
if clave = "BRA" then
val = "SUDAMERICA"
end if
if clave = "CHL" then
val = "SUDAMERICA"
end if
if clave = "COL" then
val = "SUDAMERICA"
end if
if clave = "DEU" then
val = "EUROPA"
end if


if clave = "DOM" then
val = "ANTILLAS"
end if
if clave = "ESP" then
val = "EUROPA"
end if
if clave = "FRA" then
val = "EUROPA"
end if
if clave = "GBR" then
val = "EUROPA"
end if
if clave = "SLV" then
val = "CENTROAMERICA"
end if

if clave = "THA" then
val = "ASIA"
end if
if clave = "USA" then
val = "NORTE AMERICA"
end if
if clave = "ZYA" then
val = "EUROPA"
end if

if clave = "CAN" then
val = "NORTE AMERICA"
end if

if clave = "CHN" then
val = "ASIA"
end if

if clave = "ITA" then
val = "EUROPA"
end if

if clave = "DNK" then
val = "EUROPA"
end if

if clave = "PHL" then
val = "ASIA"
end if

if clave = "BEL" then
val = "EUROPA"
end if

if clave = "BOL" then
val = "SUDAMERICA"
end if

if clave = "IND" then
val = "ASIA"
end if

if clave = "JPN" then
val = "ASIA"
end if



if clave = "PER" then
val = "SUDAMERICA"
end if

if clave = "VEN" then
val = "SUDAMERICA"
end if

if clave = "CRI" then
val = "LATINOAMERICA"
end if

if clave = "AUT" then
val = "ASIA"
end if


if clave = "SGP" then
val = "ASIA"
end if

if clave = "NZL" then
val = "EUROPA"
end if

if clave = "CHE" then
val = "EUROPA"
end if

if clave = "MYS" then
val = "ASIA"
end if



retornaRegion = val
end function


function ordenOcupado(Subref,referencia)
dim res
res = False


if(referencia <> subrefAux)then
subrefaux=referencia
  orden(1)=""
  orden(2)=""
  orden(3)=""
   orden(4)=""
	orden(5)=""
	 orden(6)=""
	  orden(7)=""
	   orden(8)=""
		orden(9)=""
		orden(10)=""
		orden(11)=""
	 orden(12)=""
	  orden(13)=""
	   orden(14)=""
	    orden(15)=""
		 orden(16)=""
		  orden(17)=""
end if

if(subref <> "" ) then

'for i=0 to 50 
if orden(Subref) = "1" then
 res = True
else
 res = False
end if
'next
end if


ordenOcupado = res
end function

function ocuparOrd(Subref)
dim res
res = False

if(subref <> "" ) then

orden(Subref) ="1"
'redim orden
end if

ocuparOrd = res
end function

function retornaCECO(rcli)
dim val,aux
val =""



If InStr(rcli,"CECO")>0 then
'If inStr(ucase(rcli),"CUENTA")>0 then

 aux=split(rcli," ")
' Response.write(rcli & ", "& Ubound(aux) & ":" & aux(0) &"," & aux(1))
' response.End()
 if Ubound(aux) = 0  then
'  aux=ucase(mid(aux,"CECO:")
	  if InStr(aux(0),"CECO")>0then
	   val = aux(0)
	  else
	  
	  
	  
	    if InStr(aux(1),"CECO")>0then
		 val = aux(1)
	    else
	     val = "N/E"
	    end if
		
		
		
		
	  end if
 else
  if(Ubound(aux) = 1) then
  
      if InStr(aux(0),"CECO")>0then
	    val = aux(0)
	  else
	    if InStr(aux(1),"CECO")>0then
		 val = aux(1)
	    else
	     val = "N/E"
	    end if
	  end if
	  
   else
    val="ERROR:"&rcli& "," &Ubound(aux)
   end if
 end if
Else
val = "N/E"
End if

retornaCECO = val
end function






function retornaCuenta(rcli)
dim val,aux
val =""



If InStr(rcli,"CUENTA")>0 then
 aux=split(rcli," ")
 if Ubound(aux) = 0  then
'  aux=ucase(mid(aux,"CECO:")
	  if InStr(aux(0),"CUENTA")>0then
	   val = aux(0)
	  else
	    'if InStr(aux(1),"CUENTA")>0then
		' val = aux(1)
	    'else
	     val = "N/E"
	    'end if
	  end if
 else
  if(Ubound(aux) = 1) then
  
      if InStr(aux(0),"CUENTA")>0then
	    val = aux(0)
	  else
	    if InStr(aux(1),"CUENTA")>0then
		 val = aux(1)
	    else
	     val = "N/E"
	    end if
	  end if
	  
   else
    val="ERROR:"&rcli
   end if
 end if
Else
val = "N/E"
End if


retornaCuenta = val
end function
%>

