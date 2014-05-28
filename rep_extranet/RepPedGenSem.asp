<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
 <%
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

Response.Buffer = TRUE
Response.Addheader "Content-Disposition", "attachment; filename=ReportePed.xls" 
Response.ContentType = "application/vnd.ms-excel"

 Server.ScriptTimeOut=200000
 strHTML = ""
 strHTML2 = ""
 sQueryComp= " inner  " 
 strDate=trim(request.Form("txtDateIni"))
 strDate2 = trim(request.Form("txtDateFin"))
 bConNoFact= request.Form("nValor")
 nCount= 0
'Response.Write("entroo")
'Response.End

 'strDate="01-06-2010"
 'strDate2="10-06-2010"  
 
 if not strDate="" and not strDate2="" then
   
   tmpDiaFin = cstr(datepart("d",strDate))
   tmpMesFin = cstr(datepart("m",strDate))
   tmpAnioFin = cstr(datepart("yyyy",strDate))
   strDateFin = tmpAnioFin & "-" &tmpMesFin & "-"& tmpDiaFin

   tmpDiaFin2 = cstr(datepart("d",strDate2))
   tmpMesFin2 = cstr(datepart("m",strDate2))
   tmpAnioFin2 = cstr(datepart("yyyy",strDate2))
   strDateFin2 = tmpAnioFin2 & "-" &tmpMesFin2 & "-"& tmpDiaFin2
   

    'strDateFin="2010-06-01"
	'strDateFin2="2010-06-05"
   dim con,Rsio,Rsio2


strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"clie01")



if not permi = "" then
  permi = "  and (" & permi & ") "
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("rfcCliente")



if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if
if blnAplicaFiltro then
   permi = strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
end if


   set Rsio = server.CreateObject("ADODB.Recordset")
   Rsio.ActiveConnection ="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=dai_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
   'Rsio.ActiveConnection = MM_EXTRANET_STRING
     
  strSQL = "  select " & _
			" i.refcia01 as 'REFERENCIA', " & _
			" sg.numgui04 as 'BL', " & _
			" ifnull(gh.numgui04,'--') as 'House', " & _
			" pto.nompto01 as 'PORT', " & _
			" ifnull(air.desc01,'--') as 'LINEA AEREA', " & _
			" ifnull(nav.nom01,'--') as 'SHIPPING', " & _
			" ifnull(rex.nomb01,'--') as 'FORWARDER', " & _ 
			" CONCAT(i.PATENT01, CONCAT( '-',i.NUMPED01 ) )  as 'IMPORT DOCUMENT', " & _
			" prv.nompro22 as 'SHIPPER', " & _
			" ar.cpro05 as 'DESCRIPTION CODE', " & _
			" ar.obse05 as 'ObservacionesMerc', " & _
			" replace(replace(ar.desc05,'\n',''),'\r','') as 'DESCRIPTION', " & _
			" ifnull(ar.caco05,0) as 'CANTMERC', " & _
			" mo.descri30 as 'MODALIDAD', " & _
			" i.patent01 as 'Patente', " & _
			" i.numped01 as 'Pedimento', " & _
			" i.adusec01 as 'Aduana', " & _
			" fr.fraarn02 as 'Fraccion', " & _
			" fr.d_mer102 as 'Descripcion', " & _
			" ifnull(cta.cgas31,'S/CG') as 'Cuenta de Gastos', " & _
			" cta.fech31 as 'Fecha de la C.G.', " & _
			" ifnull(cta.csce31,0) as 'ServiciosComp', " & _
			" ifnull(cta.chon31,0) as 'Honorarios', " & _
			" ifnull(igi.import36,0) as 'IGI', " & _
			" ifnull(dta.import36,0) as 'DTA', " & _
			" (fr.i_adv102+fr.i_adv202) as ADVFrac,   " & _
			" ifnull(cp.desc21,'--') as DescBene, " & _
			" ep.conc21 as 'Concepto', " & _
			" ifnull((dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as 'ImportePH', " & _
			" ifnull(cta.tota31,0) as 'CuentaGastos', " & _ 
			" ifnull(ep.fech21,'--') as 'FechaCuenta',  " & _
			" ifnull(cbe.nomb20,'--') as 'Beneficiario', " & _
			"(SELECT con.agente32 FROM dai_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AgenteA', " &_
			" ar.tpmerc05  as 'TipoMercancia' " & _
			" from dai_extranet.ssdagi01 as i " & _
			" inner join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
			" left join dai_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
			" left join dai_extranet.c01airln as air on air.cvela01 = r.cvela01 " & _
			" left join dai_extranet.c01ptoemb as pto on pto.Cvepto01= r.cveptoemb " & _
			" left join dai_extranet.c01reexp  as rex on rex.cvrexp01 = r.cvrexp01 " & _
			" left join dai_extranet.ssmtra30   as mo on mo.clavet30 = i.cvemts01 " & _
			" left join dai_extranet.ssguia04 as sg on sg.refcia04  = i.refcia01 and sg.idngui04 = 1 " & _
			" left join dai_extranet.ssguia04 as gh on gh.refcia04  = i.refcia01 and gh.idngui04 = 2 " & _
			" left join dai_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
			" left join dai_extranet.c01navie as nav on nav.cve01 = r.naim01   " & _
         "    left join dai_extranet.sscont36 as igi on igi.refcia36 = i.refcia01 and igi.cveimp36 = '6' " & _
         "    left join dai_extranet.sscont36 as dta on dta.refcia36 = i.refcia01 and dta.cveimp36 = '1' " & _
			" left join dai_extranet.d05artic as ar on ar.refe05 = i.refcia01 and ar.refe05 = r.refe01 " & _
                  " left join dai_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
                  " left join dai_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _
			" left join dai_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _
			" left join dai_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = ctar.cgas31 " & _
			" left join dai_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
			" left join dai_extranet.c20benef as cbe on ep.bene21 = cbe.clav20 and cbe.aplic20 <> 'T' " & _
		"        left join  dai_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
       " where cc.rfccli18 = '"&permi&"' and i.firmae01 is not null and i.firmae01 <> ''  and  i.fecpag01 >='"&strDateFin&"' and i.fecpag01 <='"&strDateFin2&"' " & _
    " group by i.refcia01, ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05,ctar.cgas31,ep.conc21 " & _
	" union all " & _
		"  select " & _
			" i.refcia01 as 'REFERENCIA', " & _
			" sg.numgui04 as 'BL', " & _
			" ifnull(gh.numgui04,'--') as 'House', " & _
			" pto.nompto01 as 'PORT', " & _
			" ifnull(air.desc01,'--') as 'LINEA AEREA', " & _
			" ifnull(nav.nom01,'--') as 'SHIPPING', " & _
			" ifnull(rex.nomb01,'--') as 'FORWARDER', " & _ 
			" CONCAT(i.PATENT01, CONCAT( '-',i.NUMPED01 ) )  as 'IMPORT DOCUMENT', " & _
			" prv.nompro22 as 'SHIPPER', " & _
			" ar.cpro05 as 'DESCRIPTION CODE', " & _
			" ar.obse05 as 'ObservacionesMerc', " & _
			" replace(replace(ar.desc05,'\n',''),'\r','') as 'DESCRIPTION', " & _
			" ifnull(ar.caco05,0) as 'CANTMERC', " & _
			" mo.descri30 as 'MODALIDAD', " & _
			" i.patent01 as 'Patente', " & _
			" i.numped01 as 'Pedimento', " & _
			" i.adusec01 as 'Aduana', " & _
			" fr.fraarn02 as 'Fraccion', " & _
			" fr.d_mer102 as 'Descripcion', " & _
			" ifnull(cta.cgas31,'S/CG') as 'Cuenta de Gastos', " & _
			" cta.fech31 as 'Fecha de la C.G.', " & _
			" ifnull(cta.csce31,0) as 'ServiciosComp', " & _
			" ifnull(cta.chon31,0) as 'Honorarios', " & _
			" ifnull(igi.import36,0) as 'IGI', " & _
			" ifnull(dta.import36,0) as 'DTA', " & _
			" (fr.i_adv102+fr.i_adv202) as ADVFrac,   " & _
			" ifnull(cp.desc21,'--') as DescBene , " & _
			" ep.conc21 as 'Concepto', " & _
			" ifnull((dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as 'ImportePH', " & _
			" ifnull(cta.tota31,0) as 'CuentaGastos', " & _ 
			" ifnull(ep.fech21,'--') as 'FechaCuenta',  " & _
			" ifnull(cbe.nomb20,'--') as 'Beneficiario', " & _
			"(SELECT con.agente32 FROM tol_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AgenteA', " &_
			" ar.tpmerc05  as 'TipoMercancia' " & _
			" from tol_extranet.ssdagi01 as i " & _
			" inner join tol_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
			" left join tol_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
			" left join tol_extranet.c01airln as air on air.cvela01 = r.cvela01 " & _
			" left join tol_extranet.c01ptoemb as pto on pto.Cvepto01= r.cveptoemb " & _
			" left join tol_extranet.c01reexp  as rex on rex.cvrexp01 = r.cvrexp01 " & _
			" left join tol_extranet.ssmtra30   as mo on mo.clavet30 = i.cvemts01 " & _
			" left join tol_extranet.ssguia04 as sg on sg.refcia04  = i.refcia01 and sg.idngui04 = 1 " & _
			" left join tol_extranet.ssguia04 as gh on gh.refcia04  = i.refcia01 and gh.idngui04 = 2 " & _
			" left join tol_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
			" left join tol_extranet.c01navie as nav on nav.cve01 = r.naim01   " & _
         "    left join tol_extranet.sscont36 as igi on igi.refcia36 = i.refcia01 and igi.cveimp36 = '6' " & _
         "    left join tol_extranet.sscont36 as dta on dta.refcia36 = i.refcia01 and dta.cveimp36 = '1' " & _
			" left join tol_extranet.d05artic as ar on ar.refe05 = i.refcia01 and ar.refe05 = r.refe01 " & _
                  " left join tol_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
                  " left join tol_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _
			" left join tol_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _
			" left join tol_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = ctar.cgas31 " & _
			" left join tol_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
			" left join tol_extranet.c20benef as cbe on ep.bene21 = cbe.clav20 and cbe.aplic20 <> 'T' " & _
		"        left join  tol_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
       " where cc.rfccli18 = '"&permi&"' and i.firmae01 is not null and i.firmae01 <> ''  and  i.fecpag01 >='"&strDateFin&"' and i.fecpag01 <='"&strDateFin2&"' " & _
    " group by i.refcia01, ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05,ctar.cgas31,ep.conc21 " & _
	" union all " & _
		"  select " & _
			" i.refcia01 as 'REFERENCIA', " & _
			" sg.numgui04 as 'BL', " & _
			" ifnull(gh.numgui04,'--') as 'House', " & _
			" pto.nompto01 as 'PORT', " & _
			" ifnull(air.desc01,'--') as 'LINEA AEREA', " & _
			" ifnull(nav.nom01,'--') as 'SHIPPING', " & _
			" ifnull(rex.nomb01,'--') as 'FORWARDER', " & _ 
			" CONCAT(i.PATENT01, CONCAT( '-',i.NUMPED01 ) )  as 'IMPORT DOCUMENT', " & _
			" prv.nompro22 as 'SHIPPER', " & _
			" ar.cpro05 as 'DESCRIPTION CODE', " & _
			" ar.obse05 as 'ObservacionesMerc', " & _
			" replace(replace(ar.desc05,'\n',''),'\r','') as 'DESCRIPTION', " & _
			" ifnull(ar.caco05,0) as 'CANTMERC', " & _
			" mo.descri30 as 'MODALIDAD', " & _
			" i.patent01 as 'Patente', " & _
			" i.numped01 as 'Pedimento', " & _
			" i.adusec01 as 'Aduana', " & _
			" fr.fraarn02 as 'Fraccion', " & _
			" fr.d_mer102 as 'Descripcion', " & _
			" ifnull(cta.cgas31,'S/CG') as 'Cuenta de Gastos', " & _
			" cta.fech31 as 'Fecha de la C.G.', " & _
			" ifnull(cta.csce31,0) as 'ServiciosComp', " & _
			" ifnull(cta.chon31,0) as 'Honorarios', " & _
			" ifnull(igi.import36,0) as 'IGI', " & _
			" ifnull(dta.import36,0) as 'DTA', " & _
			" (fr.i_adv102+fr.i_adv202) as ADVFrac,   " & _
			" ifnull(cp.desc21,'--') as DescBene , " & _
			" ep.conc21 as 'Concepto', " & _
			" ifnull((dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as 'ImportePH', " & _
			" ifnull(cta.tota31,0) as 'CuentaGastos', " & _ 
			" ifnull(ep.fech21,'--') as 'FechaCuenta',  " & _
			" ifnull(cbe.nomb20,'--') as 'Beneficiario', " & _
			"(SELECT con.agente32 FROM sap_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AgenteA', " &_
			" ar.tpmerc05  as 'TipoMercancia' " & _
			" from sap_extranet.ssdagi01 as i " & _
			" inner join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
			" left join sap_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
			" left join sap_extranet.c01airln as air on air.cvela01 = r.cvela01 " & _
			" left join sap_extranet.c01ptoemb as pto on pto.Cvepto01= r.cveptoemb " & _
			" left join sap_extranet.c01reexp  as rex on rex.cvrexp01 = r.cvrexp01 " & _
			" left join sap_extranet.ssmtra30   as mo on mo.clavet30 = i.cvemts01 " & _
			" left join sap_extranet.ssguia04 as sg on sg.refcia04  = i.refcia01 and sg.idngui04 = 1 " & _
			" left join sap_extranet.ssguia04 as gh on gh.refcia04  = i.refcia01 and gh.idngui04 = 2 " & _
			" left join sap_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
			" left join sap_extranet.c01navie as nav on nav.cve01 = r.naim01   " & _
         "    left join sap_extranet.sscont36 as igi on igi.refcia36 = i.refcia01 and igi.cveimp36 = '6' " & _
         "    left join sap_extranet.sscont36 as dta on dta.refcia36 = i.refcia01 and dta.cveimp36 = '1' " & _
			" left join sap_extranet.d05artic as ar on ar.refe05 = i.refcia01 and ar.refe05 = r.refe01 " & _
                  " left join sap_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
                  " left join sap_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _
			" left join sap_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _
			" left join sap_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = ctar.cgas31 " & _
			" left join sap_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
			" left join sap_extranet.c20benef as cbe on ep.bene21 = cbe.clav20 and cbe.aplic20 <> 'T' " & _
		"        left join  sap_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
       " where cc.rfccli18 = '"&permi&"' and i.firmae01 is not null and i.firmae01 <> ''  and  i.fecpag01 >='"&strDateFin&"' and i.fecpag01 <='"&strDateFin2&"' " & _
    " group by i.refcia01, ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05,ctar.cgas31,ep.conc21 " & _
	" union all " & _
		"  select " & _
			" i.refcia01 as 'REFERENCIA', " & _
			" sg.numgui04 as 'BL', " & _
			" ifnull(gh.numgui04,'--') as 'House', " & _
			" pto.nompto01 as 'PORT', " & _
			" ifnull(air.desc01,'--') as 'LINEA AEREA', " & _
			" ifnull(nav.nom01,'--') as 'SHIPPING', " & _
			" ifnull(rex.nomb01,'--') as 'FORWARDER', " & _ 
			" CONCAT(i.PATENT01, CONCAT( '-',i.NUMPED01 ) )  as 'IMPORT DOCUMENT', " & _
			" prv.nompro22 as 'SHIPPER', " & _
			" ar.cpro05 as 'DESCRIPTION CODE', " & _
			" ar.obse05 as 'ObservacionesMerc', " & _
			" replace(replace(ar.desc05,'\n',''),'\r','') as 'DESCRIPTION', " & _
			" ifnull(ar.caco05,0) as 'CANTMERC', " & _
			" mo.descri30 as 'MODALIDAD', " & _
			" i.patent01 as 'Patente', " & _
			" i.numped01 as 'Pedimento', " & _
			" i.adusec01 as 'Aduana', " & _
			" fr.fraarn02 as 'Fraccion', " & _
			" fr.d_mer102 as 'Descripcion', " & _
			" ifnull(cta.cgas31,'S/CG') as 'Cuenta de Gastos', " & _
			" cta.fech31 as 'Fecha de la C.G.', " & _
			" ifnull(cta.csce31,0) as 'ServiciosComp', " & _
			" ifnull(cta.chon31,0) as 'Honorarios', " & _
			" ifnull(igi.import36,0) as 'IGI', " & _
			" ifnull(dta.import36,0) as 'DTA', " & _
			" (fr.i_adv102+fr.i_adv202) as ADVFrac,   " & _
			" ifnull(cp.desc21,'--') as DescBene, " & _
			" ep.conc21 as 'Concepto', " & _
			" ifnull((dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as 'ImportePH', " & _
			" ifnull(cta.tota31,0) as 'CuentaGastos', " & _ 
			" ifnull(ep.fech21,'--') as 'FechaCuenta',  " & _
			" ifnull(cbe.nomb20,'--') as 'Beneficiario', " & _
			"(SELECT con.agente32 FROM lzr_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AgenteA', " &_
			" ar.tpmerc05  as 'TipoMercancia' " & _
			" from lzr_extranet.ssdagi01 as i " & _
			" inner join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01 " & _
			" left join lzr_extranet.c01refer as r on r.refe01 = i.refcia01 " & _    
			" left join lzr_extranet.c01airln as air on air.cvela01 = r.cvela01 " & _
			" left join lzr_extranet.c01ptoemb as pto on pto.Cvepto01= r.cveptoemb " & _
			" left join lzr_extranet.c01reexp  as rex on rex.cvrexp01 = r.cvrexp01 " & _
			" left join lzr_extranet.ssmtra30   as mo on mo.clavet30 = i.cvemts01 " & _
			" left join lzr_extranet.ssguia04 as sg on sg.refcia04  = i.refcia01 and sg.idngui04 = 1 " & _
			" left join lzr_extranet.ssguia04 as gh on gh.refcia04  = i.refcia01 and gh.idngui04 = 2 " & _
			" left join lzr_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
			" left join lzr_extranet.c01navie as nav on nav.cve01 = r.naim01   " & _
         "    left join lzr_extranet.sscont36 as igi on igi.refcia36 = i.refcia01 and igi.cveimp36 = '6' " & _
         "    left join lzr_extranet.sscont36 as dta on dta.refcia36 = i.refcia01 and dta.cveimp36 = '1' " & _
			" left join lzr_extranet.d05artic as ar on ar.refe05 = i.refcia01 and ar.refe05 = r.refe01 " & _
                  " left join lzr_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
                  " left join lzr_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _
			" left join lzr_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _
			" left join lzr_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = ctar.cgas31 " & _
			" left join lzr_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
			" left join lzr_extranet.c20benef as cbe on ep.bene21 = cbe.clav20 and cbe.aplic20 <> 'T' " & _
		"        left join  lzr_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
       " where cc.rfccli18 = '"&permi&"' and i.firmae01 is not null and i.firmae01 <> ''  and  i.fecpag01 >='"&strDateFin&"' and i.fecpag01 <='"&strDateFin2&"' " & _
    " group by i.refcia01, ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05,ctar.cgas31,ep.conc21 "

	'Response.Write(strSQL)
	'Response.End
	
   Rsio.Source= strSQL
   Rsio.CursorType = 0
   Rsio.CursorLocation = 2
   Rsio.LockType = 1
   Rsio.Open()
 	
   
  strHTML2 = strHTML2 & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE GENERAL DE PEDIMENTOS</p></font></strong>"
   strHTML2 = strHTML2 & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
   strHTML2 = strHTML2 & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p> Del " & strDate & " al " & strDate2 & " </p></font></strong>"
   strHTML2 = strHTML2 & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
		strHTML2 = strHTML2 & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">B. OF L. / <br> AW. B. M. </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CONTAINER/ <br> AW. B. H. </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""150"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PORT/AIRPORT <br> OF DEPARTURE </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""200"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">SHIPPING LINE / <br> FORWARDER </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IMPORT <br> DOCUMENT </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""200"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">SHIPPER </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""150"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DESCRIPTION CODE </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""350"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DESCRIPTION </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ATA W-H  </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">MODALIDAD </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CUENTA DE <br> GASTOS </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FECHA DE <br> CUENTA G. </td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""200"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CONCEPTO </td> " & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TOTAL </td>  " & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""300"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">BENEFICIARIO </td> " & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TIPO MERCANCIA </td> " & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PATENTE </td> " & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PEDIMENTO </td> " & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ADUANA </td> " & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FRACCION </td> " & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""300"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DESCRIPCION </td> " & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""100"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">SERVICIOS <br> COMPLEMENTARIOS" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">HONORARIOS" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DTA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IGI" & chr(13) & chr(10)
			      
		strHTML = strHTML2 & "</tr>"& chr(13) & chr(10) 

   RefAux=""
      
   While NOT Rsio.EOF
		'Variables
		sBenef = "--"
		sCGastos = "--"
		sContenedor = ""
		sLineaMar = ""
		strFechaATAWH      = ""
		strComentarioATAWH = ""
		strHoraATAWH       = "--"
		sAdu="dai"
		sHonor=0
		sDTA=0
		sIgi=0
		select case cInt(Rsio.Fields.Item("Aduana").Value)
			case 470
				sAdu="dai"
			case 510
				sAdu="lzr"
			case 160
				sAdu="sap"
			case 650
				sAdu="tol"
		end select 
		
		
		
		if(RefAux <> cStr(Rsio.Fields.Item("REFERENCIA").Value)) then
			
			sCGastos= cStr(Rsio.Fields.Item("Cuenta de Gastos").Value)
			dCGastos= Rsio.Fields.Item("Fecha de la C.G.").Value
			sSerComp= Rsio.Fields.Item("ServiciosComp").Value
			sHonor = Rsio.Fields.Item("Honorarios").Value
			sDTA = Rsio.Fields.Item("DTA").Value
			sIgi = Rsio.Fields.Item("IGI").Value			
			RefAux= cStr(Rsio.Fields.Item("REFERENCIA").Value)
			
			'OBTENGO EL TOTAL DE LA MERCANCIA DE ESTA REFERENCIA
			nTotMerc=0
			Set RTotMer = Server.CreateObject("ADODB.Recordset")
			RTotMer.ActiveConnection = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=dai_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
			strSqlMer = "select  ifnull(sum(ar.caco05),0) as TotMerc from "&sAdu&"_extranet.d05artic as ar where ar.refe05='"&RefAux&"'   " 
						 
			RTotMer.Source = strSqlMer
			RTotMer.CursorType = 0
			RTotMer.CursorLocation = 2
			RTotMer.LockType = 1
			RTotMer.Open()
			if not RTotMer.eof then
				nTotMerc  = RTotMer.Fields.Item("TotMerc").Value
			end if
			RTotMer.close
			set RTotMer = Nothing
			 	
		end if
		
		'**************************************************************************************************************
		'PRORRATEO DE PAGOS HECHOS
		nTotPago= Rsio.Fields.Item("ImportePH").Value
		nCantMerc= Rsio.Fields.Item("CANTMERC").Value
		if(nTotPago <> 0 and nCantMerc <> 0) then 	
			nProrPago= (nTotPago * nCantMerc)/ nTotMerc		
		else
			nProrPago= 0
		end if 
		'Response.Write(nProrPago)
		'Response.End
		'**************************************************************************************************************
		' Contenedores
		strNumConte = ""
		strATDRAIL  = ""
		strETA_CP   = ""
		strATAC_P   = ""
		strETAW_H   = ""
		'----------------
		
			
		if RefAux <> "" then
			 Set RContenedores = Server.CreateObject("ADODB.Recordset")
			 RContenedores.ActiveConnection ="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=dai_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
			 'strSqlSel =  "select marc01 from d01conte where refe01 = '" & ltrim(RefAux) & "' "
			 strSqlSel =  " select marc01, " & _
						  "       fcTren01 as ATDRAIL, " & _
						  "       feCont01 as ETA_CP,  " & _
						  "       frCont01 as ATAC_P,  " & _
						  "       feAlma01 as ETAW_H   " & _
						  " from "&sAdu&"_extranet.d01conte where refe01 = '" & ltrim(RefAux) & "' "

			 'ATD RAIL (Fecha de Carga en Tren) d01Conte.fcTren01
			 'ETA C./P. (Estimada de Arribo Contrimodal)  d01Conte.feCont01
			 'ATA C./P. (Real de Arribo Contrimodal) d01Conte.frCont01
			 'ETA W/H (Fecha de llegada a Almacen de SEM) d01Conte.feAlma01

			 'Response.Write(strSqlSel)
			 'Response.End
			 RContenedores.Source = strSqlSel
			 RContenedores.CursorType = 0
			 RContenedores.CursorLocation = 2
			 RContenedores.LockType = 1
			 RContenedores.Open()
			 if not RContenedores.eof then
			   While NOT RContenedores.EOF
					   strNumConte = RContenedores.Fields.Item("marc01").Value
					   strATDRAIL  = RContenedores.Fields.Item("ATDRAIL").Value
					   strETA_CP   = RContenedores.Fields.Item("ETA_CP").Value
					   strATAC_P   = RContenedores.Fields.Item("ATAC_P").Value
					   strETAW_H   = RContenedores.Fields.Item("ETAW_H").Value
					   '*********************************************
						 strFechaATAWH      = ""
						 strComentarioATAWH = ""
						 strHoraATAWH       = ""
						 Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
						 RConteDetalle.ActiveConnection = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=dai_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
						 strSqlSel = " SELECT ifnull(e.f_fecha,'--'),   " & _
									 "        e.t_hora,   " & _
									 "        e.m_observ  " & _
									 " FROM "&sAdu&"_status.etxcoi as e, "&sAdu&"_status.etaps as a " & _
									 " where e.n_etapa = a.n_etapa and " & _
									 "       e.c_referencia = '" & ltrim(RefAux)    & "' and    " & _
									 "       e.c_conte      = '" & ltrim(strNumConte) & "' and " & _
									 "       a.d_abrev      = 'LLP'             " & _
									 " order by n_secuenc desc                  "
						 'Response.Write(strSqlSel)
						 'Response.End
						 RConteDetalle.Source = strSqlSel
						 RConteDetalle.CursorType = 0
						 RConteDetalle.CursorLocation = 2
						 RConteDetalle.LockType = 1
						 RConteDetalle.Open()
						 if not RConteDetalle.eof then
							 strFechaATAWH       = RConteDetalle.Fields.Item("f_fecha").Value
							 strHoraATAWH        = RConteDetalle.Fields.Item("t_hora").Value
							 strComentarioATAWH  = RConteDetalle.Fields.Item("m_observ").Value
						 end if
						 RConteDetalle.close
						 set RConteDetalle = Nothing
					   '*********************************************
				 RContenedores.movenext
			   Wend
			 else
			 'Response.Write("No hay ningun contenedor")
			 'Response.End

			 end if
			 RContenedores.close
			 set RContenedores = Nothing
		end if

		'*************************************************************
		'Tracking Aereo o Maritimo		
		if (cInt(Rsio.Fields.Item("Aduana").Value) = 470 or cInt(Rsio.Fields.Item("Aduana").Value) = 650) then
			sContenedor = Rsio.Fields.Item("House").Value			
			sLineaMar = Rsio.Fields.Item("FORWARDER").Value
		else
			sContenedor = strNumConte
			sLineaMar = Rsio.Fields.Item("SHIPPING").Value
		end if
		
		
		
		'----------------------------------------------------------------------------------------------------------------------------------
		'CUERPO DEL REPORTE
		strHTML = strHTML&"<tr>" & chr(13) & chr(10)        
			strHTML = strHTML&"<td width=""90""align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("REFERENCIA").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90""align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("BL").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sContenedor& "</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""150"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("PORT").Value &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""200"" align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sLineaMar &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("IMPORT DOCUMENT").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""300"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("SHIPPER").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""150"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("DESCRIPTION CODE").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""350"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("DESCRIPTION").Value& " - " & Rsio.Fields.Item("ObservacionesMerc").Value &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strFechaATAWH &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("MODALIDAD").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sCGastos &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& dCGastos &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""300"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("DescBene").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& nProrPago &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""300"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Beneficiario").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TipoMercancia").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Patente").Value &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Pedimento").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Aduana").Value &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Fraccion").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""300"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Descripcion").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""100"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sSerComp &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sHonor &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sDTA &" </font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sIgi &"</font></td>" & chr(13) & chr(10)
			
		strHTML = strHTML&"</tr>" & chr(13) & chr(10)
       '----------------------------------------------------------------------------------------------------------------------------------
	   Response.Write( strHTML )
       strHTML = ""
		
		
        nCount=nCount+1

  Rsio.MoveNext()
  Wend

Rsio.Close()
Set Rsio = Nothing
end if

if(nCount < 1)then
	strHTML= strHTML2 & "</tr>"& chr(13) & chr(10)
	strHTML= strHTML &  "<tr> <td colspan=""12"">"& chr(13) & chr(10)
	strHTML= strHTML &  "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>No se encontr&oacute; ning&uacute;n registro</p></font></strong>"& chr(13) & chr(10)
	strHTML= strHTML & "</tr> </td>"& chr(13) & chr(10)
	strHTML = strHTML & "</table>"& chr(13) & chr(10)
	else
	strHTML = strHTML & "</table>"& chr(13) & chr(10)
end if

response.Write(strHTML)
%>
