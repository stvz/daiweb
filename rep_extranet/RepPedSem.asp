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
 
 if not strDate="" and not strDate2="" then
 


   tmpDiaFin = cstr(datepart("d",strDate))
   tmpMesFin = cstr(datepart("m",strDate))
   tmpAnioFin = cstr(datepart("yyyy",strDate))
   strDateFin = tmpAnioFin & "-" &tmpMesFin & "-"& tmpDiaFin

   tmpDiaFin2 = cstr(datepart("d",strDate2))
   tmpMesFin2 = cstr(datepart("m",strDate2))
   tmpAnioFin2 = cstr(datepart("yyyy",strDate2))
   strDateFin2 = tmpAnioFin2 & "-" &tmpMesFin2 & "-"& tmpDiaFin2


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

'--------------------------------------------------------------
'Recibe el check si quiere incluir las no facturadas
'--------------------------------------------------------------


if CInt(bConNoFact)=1  then
	sQueryComp= " left  " 	
else
	sQueryComp= " inner  " 
end if

   set Rsio = server.CreateObject("ADODB.Recordset")
   Rsio.ActiveConnection ="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=dai_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
   'Rsio.ActiveConnection = MM_EXTRANET_STRING
     
  strSQL = " select   " & _
		   "   i.refcia01 as Referencia, " & _
		   "  ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1) as DtaPro, " & _
		   "  ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1) as PrePro, " & _
		   "   ar.tpmerc05 as Tipo,  " & _
			"   prv.nompro22 as Proveedor, " & _  
			"   prv.dompro22   as Domicilio, " & _
			"   f.vincul39  as Vincul, " & _
			"   date_format(f.fecfac39,'%d/%m/%Y') as FechFact,   " & _ 
			"  f.numfac39 as Factura, " & _
			"   (fr.prepag02/i.tipcam01) as ValComDolares,     " & _
			"   (fr.prepag02/i.tipcam01)  as VComerBaseP, " & _
			"   (i.fletes01/i.tipcam01) as FletesDlls,   " & _
			"   (i.segros01/i.tipcam01)as SegurosDlls,   " & _
			"   (i.incble01/i.tipcam01) as OtrsIncrementDlls,  " & _  
			"   ((i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) ) * ( i.tipcam01) as  IncrBasePed, " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as ValorFact, " & _
			"(SELECT con.agente32 FROM lzr_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AA', " &_
			"   i.cveped01 as TipoPed, " & _
			"   i.numped01 as Pedimento,   " & _
			"   i.patent01 as Patente,  " & _
			"   date_format(i.fecpag01,'%d/%m/%Y') as fPagoPedim,    " & _
			"   sgui.numgui04  as Guia, " & _
			"   ar.desc05 as Producto,   " & _
			"   ar.caco05 as Cantidad,   " & _
			"   ar.item05 as CodProduct,   " & _ 
			"   ar.pedi05 as OrdenCompra,    " & _
			"   0 as DocPosteo, " & _
			"   i.cvepvc01 as PaisVendedor,      " & _
			"   i.cvepod01 as PaisOrigen,   " & _
			"   i.adusec01 as AduSec,       " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as BaseAduDlls, " & _
			"   i.tipcam01 as TipoCamb,    " & _
			"   f.monfac39 as FactorMoneda, " & _
			"   fr.prepag02   as CompraMN, " & _
			"   fr.prepag02   as BasePed, " & _
			"   fr.tasadv02 as TasaAdv, " & _
			"   i.tsadta01 as TasaDta, " & _
			"   (fr.i_adv102+fr.i_adv202) as ADVFrac,   " & _
			"   cf1.import36 as DTA,   " & _
			"   cf15.import36  as Preval,   " & _
			"   0 as CuotComp, " & _
			"   0 as Recargos, " & _
			"   (fr.prepag02 +  (fr.i_adv102+fr.i_adv202) +  ((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) ) as BaseIva,   " & _
			"   fr.tasiva02 as TasaIva,   " & _
			"   (fr.i_iva102+fr.i_iva202) as IVAFrac, " & _
			"   ((fr.i_adv102+fr.i_adv202) + (ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1)) + (ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1)) + (fr.i_iva102+fr.i_iva202) ) as TotImpuesImpo, " & _
			"	cta.cgas31 as CGastos, " & _
			"	ifnull(cta.tota31,0) as TotCGastos, " & _
			"	f.terfac39 as Incoterms,     " & _    
			"   fr.fraarn02 as FracArancel, " & _
			"   fr.p_adv102 as FormPagAdv, " & _
			"   i.p_dta101 as FormPagDta, " & _
			"   fr.p_iva102 as FormPagIva, " & _
			"   (i.fletes01)  as ValFleteMx,   " & _
			"   (i.fletes01/i.tipcam01)  as ValFleteDlls " & _
		   " from lzr_extranet.ssdagi01 as i    " & _
			"		inner join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01    " & _
			"		left join lzr_extranet.c01refer as r on r.refe01 = i.refcia01   " & _
			"		"&sQueryComp&" join lzr_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _  
			"		"&sQueryComp&" join lzr_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C'   " & _
			"		left join  lzr_extranet.ssguia04 as sgui on  sgui.refcia04 = i.refcia01 " & _
			"		left join lzr_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01    " & _
			"		left join lzr_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01   " & _
			"		left join lzr_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05  " & _   
			"		left join lzr_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01     " & _
			"		left join lzr_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'     " & _
			"		left join lzr_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'     " & _
			"		left join lzr_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'     " & _
			"		left join lzr_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'     " & _
			"		left join lzr_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL')   " & _
		   " where " & _
			"		cc.rfccli18 in ('"&permi&"') " & _
			"		and i.firmae01 is not null  " & _
			"		and i.firmae01 <> '' " & _
			"		and   i.fecpag01 >=  '"&strDateFin&"' and i.fecpag01 <= '"&strDateFin2&"'   " & _
		   " group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05   " & _
		  "	union all  " & _
		  "	select    " & _
			"   i.refcia01 as Referencia, " & _
			"  ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1) as DtaPro, " & _
		    "  ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1) as PrePro, " & _
			"   ar.tpmerc05 as Tipo,  " & _
			"   prv.nompro22 as Proveedor,  " & _ 
			"   prv.dompro22   as Domicilio, " & _
			"   f.vincul39 as Vincul, " & _
			"   date_format(f.fecfac39,'%d/%m/%Y') as FechFact,    " & _
			"   f.numfac39 as Factura, " & _
			"   (fr.prepag02/i.tipcam01) as ValComDolares,     " & _
			"   (fr.prepag02/i.tipcam01)  as VComerBaseP, " & _
			"   (i.fletes01/i.tipcam01) as FletesDlls,   " & _ 
			"   (i.segros01/i.tipcam01)as SegurosDlls,   " & _
			"   (i.incble01/i.tipcam01) as OtrsIncrementDlls,  " & _  
			"   ((i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) ) * ( i.tipcam01) as  IncrBasePed, " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as ValorFact, " & _
			"(SELECT con.agente32 FROM dai_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AA', " &_
			"   i.cveped01 as TipoPed, " & _
			"   i.numped01 as Pedimento,   " & _
			"   i.patent01 as Patente,  " & _
			"   date_format(i.fecpag01,'%d/%m/%Y') as fPagoPedim,    " & _
			"   sgui.numgui04  as Guia, " & _
			"   ar.desc05 as Producto,   " & _
			"   ar.caco05 as Cantidad,   " & _
			"   ar.item05 as CodProduct,   " & _
			"   ar.pedi05 as OrdenCompra,    " & _
			"   0 as DocPosteo, " & _
			"   i.cvepvc01 as PaisVendedor,      " & _
			"   i.cvepod01 as PaisOrigen,   " & _
			"   i.adusec01 as AduSec,       " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as BaseAduDlls, " & _
			"   i.tipcam01 as TipoCamb,    " & _
			"   f.monfac39 as FactorMoneda, " & _
			"   fr.prepag02   as CompraMN, " & _
			"   fr.prepag02   as BasePed, " & _
			"   fr.tasadv02 as TasaAdv, " & _
			"   i.tsadta01 as TasaDta, " & _
			"   (fr.i_adv102+fr.i_adv202) as ADVFrac,   " & _
			"   cf1.import36 as DTA,   " & _
			"   cf15.import36  as Preval,   " & _
			"   0 as CuotComp, " & _
			"   0 as Recargos, " & _
			"   (fr.prepag02 +  (fr.i_adv102+fr.i_adv202) +  ((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) ) as BaseIva,   " & _
			"   fr.tasiva02 as TasaIva,   " & _
			"   (fr.i_iva102+fr.i_iva202) as IVAFrac, " & _
			"   ((fr.i_adv102+fr.i_adv202) + (ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1)) + (ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1)) + (fr.i_iva102+fr.i_iva202) ) as TotImpuesImpo, " & _
		    "	cta.cgas31 as CGastos, " & _
			"	ifnull(cta.tota31,0) as TotCGastos, " & _
			"	f.terfac39 as Incoterms,     " & _   
			"   fr.fraarn02 as FracArancel, " & _
			"   fr.p_adv102 as FormPagAdv, " & _
			"   i.p_dta101 as FormPagDta, " & _
			"   fr.p_iva102 as FormPagIva, " & _
			"   (i.fletes01)  as ValFleteMx,    " & _
			"   (i.fletes01/i.tipcam01)  as ValFleteDlls " & _
		  "	from dai_extranet.ssdagi01 as i    " & _
			"		inner join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01    " & _
			"		left join dai_extranet.c01refer as r on r.refe01 = i.refcia01   " & _
			"		"&sQueryComp&" join dai_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _  
			"		"&sQueryComp&" join dai_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C'   " & _
			"		left join  dai_extranet.ssguia04 as sgui on  sgui.refcia04 = i.refcia01 " & _
			"		left join dai_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01    " & _
			"		left join dai_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01   " & _
			"		left join dai_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05  " & _   
			"		left join dai_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01     " & _
			"		left join dai_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'     " & _
			"		left join dai_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'     " & _
			"		left join dai_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'     " & _ 
			"		left join dai_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'     " & _
			"		left join dai_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL')   " & _
		  "	where  " & _
			"		cc.rfccli18 in ('"&permi&"') " & _
			"		and i.firmae01 is not null  " & _
			"		and i.firmae01 <> '' " & _
			"		and   i.fecpag01 >=  '"&strDateFin&"' and i.fecpag01 <= '"&strDateFin2&"'   " & _
		  "	 group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05   " & _
		  "	union all " & _
		  "	select    " & _
			"   i.refcia01 as Referencia, " & _
			"  ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1) as DtaPro, " & _
		    "  ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1) as PrePro, " & _
			"   ar.tpmerc05 as Tipo,  " & _
			"   prv.nompro22 as Proveedor,  " & _ 
			"   prv.dompro22   as Domicilio, " & _
			"   f.vincul39 as Vincul, " & _
			"   date_format(f.fecfac39,'%d/%m/%Y') as FechFact,    " & _
			"   f.numfac39 as Factura, " & _
			"   (fr.prepag02/i.tipcam01) as ValComDolares,     " & _
			"   (fr.prepag02/i.tipcam01)  as VComerBaseP, " & _
			"   (i.fletes01/i.tipcam01) as FletesDlls,   " & _
			"   (i.segros01/i.tipcam01)as SegurosDlls,   " & _
			"   (i.incble01/i.tipcam01) as OtrsIncrementDlls,  " & _  
			"   ((i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) ) * ( i.tipcam01) as  IncrBasePed, " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as ValorFact, " & _
			"(SELECT con.agente32 FROM sap_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AA', " &_
			"   i.cveped01 as TipoPed, " & _
			"   i.numped01 as Pedimento,   " & _
			"   i.patent01 as Patente,  " & _
			"   date_format(i.fecpag01,'%d/%m/%Y') as fPagoPedim,    " & _
			"   sgui.numgui04  as Guia, " & _
			"   ar.desc05 as Producto,   " & _
			"   ar.caco05 as Cantidad,   " & _
			"   ar.item05 as CodProduct,   " & _
			"   ar.pedi05 as OrdenCompra,    " & _
			"   0 as DocPosteo, " & _
			"   i.cvepvc01 as PaisVendedor,      " & _
			"   i.cvepod01 as PaisOrigen,   " & _
			"   i.adusec01 as AduSec,       " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as BaseAduDlls, " & _
			"   i.tipcam01 as TipoCamb,    " & _
			"   f.monfac39 as FactorMoneda, " & _
			"   fr.prepag02   as CompraMN, " & _
			"   fr.prepag02   as BasePed, " & _
			"   fr.tasadv02 as TasaAdv, " & _
			"   i.tsadta01 as TasaDta, " & _
			"   (fr.i_adv102+fr.i_adv202) as ADVFrac,   " & _
			"   cf1.import36 as DTA,   " & _
			"   cf15.import36  as Preval,   " & _
			"   0 as CuotComp, " & _
			"   0 as Recargos, " & _
			"   (fr.prepag02 +  (fr.i_adv102+fr.i_adv202) +  ((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) ) as BaseIva,   " & _
			"   fr.tasiva02 as TasaIva,   " & _
			"   (fr.i_iva102+fr.i_iva202) as IVAFrac, " & _
			"   ((fr.i_adv102+fr.i_adv202) + (ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1)) + (ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1)) + (fr.i_iva102+fr.i_iva202) ) as TotImpuesImpo, " & _
		    "	cta.cgas31 as CGastos, " & _
			"	ifnull(cta.tota31,0) as TotCGastos, " & _
			"	f.terfac39 as Incoterms,     " & _   
			"   fr.fraarn02 as FracArancel, " & _
			"   fr.p_adv102 as FormPagAdv, " & _
			"   i.p_dta101 as FormPagDta, " & _
			"   fr.p_iva102 as FormPagIva, " & _
			"   (i.fletes01)  as ValFleteMx,    " & _
			"   (i.fletes01/i.tipcam01)  as ValFleteDlls " & _
			" from sap_extranet.ssdagi01 as i    " & _
			"		inner join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01    " & _
			"		left join sap_extranet.c01refer as r on r.refe01 = i.refcia01   " & _
			"		"&sQueryComp&" join sap_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01   " & _
			"		"&sQueryComp&" join sap_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C'   " & _
			"		left join  sap_extranet.ssguia04 as sgui on  sgui.refcia04 = i.refcia01 " & _
			"		left join sap_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01    " & _
			"		left join sap_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01   " & _
			"		left join sap_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05     " & _
			"		left join sap_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01     " & _
			"		left join sap_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'     " & _
			"		left join sap_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'     " & _
			"		left join sap_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'     " & _
			"		left join sap_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'     " & _
			"		left join sap_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL')   " & _
		  "	where " & _
			"		cc.rfccli18 in ('"&permi&"')  " & _
			"		and i.firmae01 is not null  " & _
			"		and i.firmae01 <> '' " & _
			"		and   i.fecpag01 >=  '"&strDateFin&"' and i.fecpag01 <= '"&strDateFin2&"'   " & _
			" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05   " & _
		  "	union all " & _
		  "	select    " & _
			"   i.refcia01 as Referencia, " & _
			"  ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1) as DtaPro, " & _
		    "  ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1) as PrePro, " & _
			"   ar.tpmerc05 as Tipo,  " & _
			"   prv.nompro22 as Proveedor,  " & _ 
			"   prv.dompro22   as Domicilio, " & _
			"   f.vincul39 as Vincul, " & _
			"   date_format(f.fecfac39,'%d/%m/%Y') as FechFact,    " & _
			"   f.numfac39 as Factura, " & _
			"   (fr.prepag02/i.tipcam01) as ValComDolares,     " & _
			"   (fr.prepag02/i.tipcam01)  as VComerBaseP, " & _
			"   (i.fletes01/i.tipcam01) as FletesDlls,   " & _
			"   (i.segros01/i.tipcam01)as SegurosDlls,   " & _
			"   (i.incble01/i.tipcam01) as OtrsIncrementDlls,  " & _  
			"   ((i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) ) * ( i.tipcam01) as  IncrBasePed, " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as ValorFact, " & _
			"(SELECT con.agente32 FROM tol_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AA', " &_
			"   i.cveped01 as TipoPed, " & _
			"   i.numped01 as Pedimento,   " & _
			"   i.patent01 as Patente,  " & _
			"   date_format(i.fecpag01,'%d/%m/%Y') as fPagoPedim,    " & _
			"   sgui.numgui04  as Guia, " & _
			"   ar.desc05 as Producto,   " & _
			"   ar.caco05 as Cantidad,   " & _
			"   ar.item05 as CodProduct,   " & _
			"   ar.pedi05 as OrdenCompra,    " & _
			"   0 as DocPosteo, " & _
			"   i.cvepvc01 as PaisVendedor,      " & _
			"   i.cvepod01 as PaisOrigen,   " & _
			"   i.adusec01 as AduSec,       " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as BaseAduDlls, " & _
			"   i.tipcam01 as TipoCamb,    " & _
			"   f.monfac39 as FactorMoneda, " & _
			"   fr.prepag02   as CompraMN, " & _
			"   fr.prepag02   as BasePed, " & _
			"   fr.tasadv02 as TasaAdv, " & _
			"   i.tsadta01 as TasaDta, " & _
			"   (fr.i_adv102+fr.i_adv202) as ADVFrac,   " & _
			"   cf1.import36 as DTA,   " & _
			"   cf15.import36  as Preval,   " & _
			"   0 as CuotComp, " & _
			"   0 as Recargos, " & _
			"   (fr.prepag02 +  (fr.i_adv102+fr.i_adv202) +  ((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) ) as BaseIva,   " & _
			"   fr.tasiva02 as TasaIva,   " & _
			"   (fr.i_iva102+fr.i_iva202) as IVAFrac, " & _
			"   ((fr.i_adv102+fr.i_adv202) + (ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1)) + (ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1)) + (fr.i_iva102+fr.i_iva202) ) as TotImpuesImpo, " & _
			"	cta.cgas31 as CGastos, " & _
			"	ifnull(cta.tota31,0)  as TotCGastos, " & _
			"	f.terfac39 as Incoterms,     " & _   
			"   fr.fraarn02 as FracArancel, " & _
			"   fr.p_adv102 as FormPagAdv, " & _
			"   i.p_dta101 as FormPagDta, " & _
			"   fr.p_iva102 as FormPagIva, " & _
			"   (i.fletes01)  as ValFleteMx,    " & _
			"   (i.fletes01/i.tipcam01)  as ValFleteDlls " & _
		  " from tol_extranet.ssdagi01 as i    " & _
			"		inner join tol_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01    " & _
			"		left join tol_extranet.c01refer as r on r.refe01 = i.refcia01   " & _
			"		"&sQueryComp&" join tol_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01   " & _
			"		"&sQueryComp&" join tol_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C'   " & _
			"		left join  tol_extranet.ssguia04 as sgui on  sgui.refcia04 = i.refcia01 " & _
			"		left join tol_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01    " & _
			"		left join tol_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01   " & _
			"		left join tol_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05  " & _   
			"		left join tol_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01     " & _
			"		left join tol_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'     " & _
			"		left join tol_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'     " & _
			"		left join tol_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'     " & _
			"		left join tol_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'     " & _
			"		left join tol_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL')   " & _
		  "	where  " & _
			"		cc.rfccli18 in ('"&permi&"') " & _
			"		and i.firmae01 is not null  " & _
			"		and i.firmae01 <> '' " & _
			"		and   i.fecpag01 >=  '"&strDateFin&"' and i.fecpag01 <= '"&strDateFin2&"'   " & _
		  "	 group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05   " & _
		  "	union all  " & _
		  "	select    " & _
			"   i.refcia01 as Referencia, " & _
			"  ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1) as DtaPro, " & _
		    "  ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1) as PrePro, " & _
			"   ar.tpmerc05 as Tipo,  " & _
			"   prv.nompro22 as Proveedor,   " & _
			"   prv.dompro22   as Domicilio, " & _
			"   f.vincul39 as Vincul, " & _
			"   date_format(f.fecfac39,'%d/%m/%Y') as FechFact,    " & _
			"   f.numfac39 as Factura, " & _
			"   (fr.prepag02/i.tipcam01) as ValComDolares,     " & _
			"   (fr.prepag02/i.tipcam01)  as VComerBaseP,  " & _
			"   (i.fletes01/i.tipcam01) as FletesDlls,   " & _
			"   (i.segros01/i.tipcam01)as SegurosDlls,   " & _
			"   (i.incble01/i.tipcam01) as OtrsIncrementDlls,  " & _  
			"   ((i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) ) * ( i.tipcam01) as  IncrBasePed, " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as ValorFact, " & _
			"(SELECT con.agente32 FROM rku_extranet.ssconf32 AS con " &_
						"WHERE con.cveadu32 = i.cveadu01 AND con.cvesec32 = i.cvesec01 AND con.patent32 = i.patent01) AS 'AA', " &_
			"   i.cveped01 as TipoPed, " & _
			"   i.numped01 as Pedimento,   " & _
			"   i.patent01 as Patente,  " & _
			"   date_format(i.fecpag01,'%d/%m/%Y') as fPagoPedim,    " & _
			"   sgui.numgui04  as Guia, " & _
			"   ar.desc05 as Producto,   " & _
			"   ar.caco05 as Cantidad,   " & _
			"   ar.item05 as CodProduct,   " & _
			"   ar.pedi05 as OrdenCompra,    " & _
			"   0 as DocPosteo, " & _
			"   i.cvepvc01 as PaisVendedor,      " & _
			"   i.cvepod01 as PaisOrigen,   " & _
			"   i.adusec01 as AduSec,       " & _
			"   (i.fletes01/i.tipcam01)+(i.segros01/i.tipcam01)+(i.incble01/i.tipcam01) +(fr.prepag02/i.tipcam01)   as BaseAduDlls, " & _
			"   i.tipcam01 as TipoCamb,    " & _
			"   f.monfac39 as FactorMoneda, " & _
			"   fr.prepag02   as CompraMN, " & _
			"   fr.prepag02   as BasePed, " & _
			"   fr.tasadv02 as TasaAdv, " & _
			"   i.tsadta01 as TasaDta, " & _
			"   (fr.i_adv102+fr.i_adv202) as ADVFrac,   " & _
			"   cf1.import36 as DTA,   " & _
			"   cf15.import36  as Preval,   " & _
			"   0 as CuotComp, " & _
			"   0 as Recargos, " & _
			"   (fr.prepag02 +  (fr.i_adv102+fr.i_adv202) +  ((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) ) as BaseIva,   " & _
			"   fr.tasiva02 as TasaIva,   " & _
			"   (fr.i_iva102+fr.i_iva202) as IVAFrac, " & _
			"   ((fr.i_adv102+fr.i_adv202) + (ROUND(((cf1.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)) , 1)) + (ROUND(((cf15.import36 * fr.prepag02) / (i.valdol01 * i.tipcam01)), 1)) + (fr.i_iva102+fr.i_iva202) ) as TotImpuesImpo, " & _
			"	cta.cgas31 as CGastos, " & _
			"	ifnull(cta.tota31,0) as TotCGastos, " & _
			"	f.terfac39 as Incoterms,     " & _   
			"   fr.fraarn02 as FracArancel, " & _
			"   fr.p_adv102 as FormPagAdv, " & _
			"   i.p_dta101 as FormPagDta, " & _
			"   fr.p_iva102 as FormPagIva, " & _
			"   (i.fletes01)  as ValFleteMx,    " & _
			"   (i.fletes01/i.tipcam01)  as ValFleteDlls " & _
		   " from rku_extranet.ssdagi01 as i    " & _
			"		inner join rku_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01    " & _
			"		left join rku_extranet.c01refer as r on r.refe01 = i.refcia01   " & _
			"		"&sQueryComp&" join rku_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _  
			"		"&sQueryComp&" join rku_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C'   " & _
			"		left join  rku_extranet.ssguia04 as sgui on  sgui.refcia04 = i.refcia01 " & _
			"		left join rku_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01    " & _
			"		left join rku_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01   " & _
			"		left join rku_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05  " & _   
			"		left join rku_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01     " & _
			"		left join rku_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'     " & _
			"		left join rku_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'     " & _
			"		left join rku_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'     " & _
			"		left join rku_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'     " & _
			"		left join rku_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL')   " & _
			" where " & _
			"		cc.rfccli18 in ('"&permi&"') " & _
			"		and i.firmae01 is not null  " & _
			"		and i.firmae01 <> '' " & _
			"		and   i.fecpag01 >=  '"&strDateFin&"' and i.fecpag01 <= '"&strDateFin2&"'   " & _
			" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05   "
						
				
	'Response.Write(strSQL)
	'Response.End()
	
	Rsio.Source= strSQL
	Rsio.CursorType = 0
	Rsio.CursorLocation = 2
	Rsio.LockType = 1
	Rsio.Open()

   
	strHTML2 = strHTML2 & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>LAYOUT PEDIMENTOS</p></font></strong>"
	strHTML2 = strHTML2 & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
	strHTML2 = strHTML2 & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p> Del " & strDate & " al " & strDate2 & " </p></font></strong>"
	strHTML2 = strHTML2 & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
		strHTML2 = strHTML2 & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""left"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""left"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TIPO</td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""300"" align=""left"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PROVEEDOR" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""150"" align=""center"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DOMICILIO FISCAL DEL PEDIMENTO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">VINCUL" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FECHA FACTURA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FACTURA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">V. COMERCIAL DE LA MERCANCIA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">V. COMERC. BASE PED.S / incr</td>" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FLETE USD" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">SEGURO USD" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">INCREMENT. DLLS" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">INCREMENT. BASE PED." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">VALOR FACTURA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">AGENTE ADUANAL" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TIPO PED." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PEDIMENTO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TIPO PED. CORRELAC." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PATENTE AGENTE" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FECHA PEDIMENTO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">GUIA DEL PEDIMENTO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""200"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PRODUCTO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CANTIDAD" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CODIGO DE PRODUCTO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ORDEN DE COMPRA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DOCUMENTO DE POSTEO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PAIS VENDEDOR" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PAIS ORIGEN" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ADUANA SECCION" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">BASE ADUANA DLLS." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TIPO CAMBIO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.A." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">COMPRA M.N." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">BASE PEDIMENTO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TASA ADV." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TASA D.T.A." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ADVALOREM" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">D.T.A." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PREVALIDACION" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CUOTA COMP." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">RECARGOS" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">BASE IVA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TASA IVA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TOTAL IMPUESTOS IMPORTACION" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CUENTA DE GASTOS" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TIPO DE VALUACION" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FRACCION ARANCELARIA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FORMA DE PAGO ADV." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FORMA DE PAGO DTA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FORMA DE PAGO IVA" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FORMA DE PAGO COM" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DOCUMENTO" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FLETE TERR. DLLS" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FLETE TERR. MXP" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">BUFFER DLLS." & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">BUFFER MXP" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">COMP IGI" & chr(13) & chr(10)
		   strHTML2 = strHTML2 & "<td width=""90"" align=""right"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">COMP DTA" & chr(13) & chr(10)
		      
		strHTML = strHTML2 & "</tr>"& chr(13) & chr(10) 

   RefAux=""
      
   While NOT Rsio.EOF
		'Variables
		nReten= 0
		sBenef="--"
		sCGastos= "--"
		sVincul=""
		nDtaPro= 0
		nPrevPro= 0	
				
		'Campos para definir la vinculacion
		if(Rsio.Fields.Item("Vincul").Value=1) then
			sVincul="SI"
		end if
		if(Rsio.Fields.Item("Vincul").Value=2) then
		    sVincul="NO"
		end if
				
		nDtaPro= Rsio.Fields.Item("DtaPro").Value		
		nPrevPro= Rsio.Fields.Item("PrePro").Value
		
		if(RefAux <> cStr(Rsio.Fields.Item("Referencia").Value)) then
			sCGastos= cStr(Rsio.Fields.Item("TotCGastos").Value)
			RefAux= cStr(Rsio.Fields.Item("Referencia").Value)
		end if
		
		
		'----------------------------------------------------------------------------------------------------------------------------------
		'CUERPO DEL REPORTE
		strHTML = strHTML&"<tr>" & chr(13) & chr(10)        
			strHTML = strHTML&"<td width=""90""align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Referencia").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90""align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Tipo").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""300"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Proveedor").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""150"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Domicilio").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sVincul &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("FechFact").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Factura").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("ValComDolares").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("VComerBaseP").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("FletesDlls").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("SegurosDlls").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("OtrsIncrementDlls").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("IncrBasePed").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("ValorFact").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("AA").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TipoPed").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Pedimento").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> N/A </font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Patente").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("fPagoPedim").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Guia").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""200"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Producto").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Cantidad").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("CodProduct").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("OrdenCompra").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("DocPosteo").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("PaisVendedor").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""center"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("PaisOrigen").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("AduSec").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("BaseAduDlls").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TipoCamb").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("FactorMoneda").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("CompraMN").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("BasePed").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TasaAdv").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TasaDta").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("ADVFrac").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& nDtaPro &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& nPrevPro &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("CuotComp").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Recargos").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("BaseIva").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TasaIva").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("IVAFrac").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("TotImpuesImpo").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& sCGastos &"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""left"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("Incoterms").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("FracArancel").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("FormPagAdv").Value&" </font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("FormPagDta").Value&" </font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("FormPagIva").Value&" </font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> N/A</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> N/A </font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("ValFleteMx").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& Rsio.Fields.Item("ValFleteDlls").Value&"</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> N/A</font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> N/A </font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> 0 </font></td>" & chr(13) & chr(10)
			strHTML = strHTML&"<td width=""90"" align=""right"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> 0 </font></td>" & chr(13) & chr(10)
			
			
			
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
