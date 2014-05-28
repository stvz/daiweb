<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_libreria.asp" -->
 
 
 <style type="text/css">
.style20 {color: #FFFFFF}

table thead td {
	white-space: nowrap ;
	color: black;
	text-align: center;
	 padding:0 20px;
}
table thead{
	font-family: Arial,Georgia, "Times New Roman",Times, serif;
	font-size: 16;
	font-weight: bold;
	height: 30px;
	
}
table tbody{
	font-family: Arial,Georgia, "Times New Roman",Times, serif;
	font-size: 12;
	text-align: center;
}
tbody th, tbody td {
  text-align: center;
}

 </style>

<%

' Por: Manuel Alejandro Estevez Fernandez
'		Abril , 2013


if  Session("GAduana") = "" then
	html_ = "<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>"
else
	' Obteniendo el prefijo de la db y la oficina
	db_ = get_db(Session("GAduana"))
	oficina_ = get_nombre_oficina(Session("GAduana"))
	' Asignacion de fechas  a variables
	'fechaini_ = Request.form("txtDateIni")
	'fechafin_ = Request.form("txtDateFin")
	cgas_ = Request.form("cgas")
	
	Response.Addheader "Content-Disposition", "attachment;filename=Rep Pendientes de Facturacion "& oficina_ &" ABB.xls"
	Response.ContentType = "application/vnd.ms-excel"

	html= ""
	consulta_ = genera_sql()
  'Response.Write(consulta_)
  'Response.End()
	'html_ = html_ & consulta_
	html_ = html_ & "<table id=""reporte"">"
	html_ = html_ & genera_encabezado()
	html_ = html_ & genera_cuerpo(consulta_)
	html_ = html_ & "</table>"
	'Response.Write(html_)
	'Response.End()
	
end if ' Fin de la verificacion de los permisos


'Este metodo  tiene como finalidad la creacion de la cadena de consulta
' tomando como parametro de entrada la oficina sobre la cual se genera 
' la consulta
function  genera_sql()



 ' Nueva consulta 2012 11 26		  
      ''	sql_ = " select '' as usuario " & _
      ''              " , '' as estatus " & _
      ''              " , cast(concat('\'',adusec01,patent01,numped01) as char) as Pedimento " & _
      ''              " , refcia01 as Referencia " & _
      ''              " , operacion as Observacion " & _
      ''              " , adusec01 as Aduana " & _
      ''              " , patent01 as Patente " & _
      ''              " , numped01 as 'No Pedimento' " & _
      ''              " ,'' as 'Pedimento Reportado por ABB' " & _
      ''              " , cveped01 as 'Clave Pedimento' " & _
      ''              " , valor_aduana as 'Valor Aduana' " & _
      ''              " , valor_comercial as 'Valor Comercial' " & _
      ''              " , '' as PG " & _
      ''              " , '' as CC " & _
      ''              " , '' as Proyecto " & _
      ''              " , '' as 'Cuenta Estandar' " & _
      ''              " , '' as 'Compania' " & _
      ''              " , prv " & _
      ''              " , dta " & _
      ''              " , igi " & _
      ''              " , iva " & _
      ''              " , '' as 'Subtotal Gastos' " & _
      ''              " , total_impuestos as 'Total Impuestos Pedimento' " & _
      ''              " , '' as Proporcion " & _
      ''              " , pesobr01 as 'Kgs' " & _
      ''              " , fecent01 as 'Fecha Entrada' " & _
      ''              " , fecpag01 as 'Fecha de Pago' " & _
      ''              " , datediff(fecpag01, fecent01) as 'Dias para calculo' " & _
      ''              " , guia as 'Guia  House' " & _
      ''              " , custodia " & _
      ''              " , cus_prov " & _
      ''              " , maniobras as 'Maniobra/Manejo' " & _
      ''              " , man_prov as manejo_proveedor " & _
      ''              " , almacenaje " & _
      ''              " , alm_prov " & _
      ''              " , montacargas " & _
      ''              " , mon_prov"  & _ 
      ''              " , desconsolidacion " & _
      ''              " , des_prov"  & _ 
      ''              " , tipcam01 as 'Tipo de Cambio' " & _
      ''              " , reconocimiento_previo " & _
      ''              " , rec_prov " & _
      ''              " , extraordinarios " & _
      ''              " , ifnull(ext_prov,'') as ext_prov " & _
      ''              " , fumigacion " & _
      ''              " , fum_prov " & _
      ''              " , total_impuestos as 'Impuestos segun Pedimento' " & _
      ''              " , total_paghe " & _
      ''              " , '' as 'Base para el Pago de Honorarios' " & _
      ''              " , '' as 'Honorarios' "  & _
      ''              " , '' as 'Complementarios' " & _
      ''              " , '' as 'Embalaje' " & _
      ''              " , '' as 'Validacion' " & _
      ''              " , '' as 'Proveedor Transporte' " & _
      ''              " , '' as 'Servicios por Operacion' " & _
      ''              " , '' as 'Proveedor Servicios Operacion' " & _
      ''              " , extraordinarios " & _
      ''              " , '' as 'Sub Total Honorarios '" & _
      ''              " , '' as 'IVA Honorarios' " & _
      ''              " , '' as 'Total Honorarios' " & _
      ''              " , '' as '' " & _
      ''              " from( " & _
      ''              " select " & _
      ''              "   refcia01 " & _
      ''              "   , operacion " & _
      ''              "   , pesobr01 " & _
      ''              "   , fecpag01 " & _
      ''              "   , fecent01 " & _
      ''              "   , cveped01 " & _
      ''              "   , numped01 " & _
      ''              "   , patent01 " & _
      ''              "   , adusec01 " & _
      ''              "   , group_concat(distinct numgui04) as guia " & _
      ''              "   , tipcam01 " & _
      ''              "   , (select sum(fr_.vaduan02) from ssfrac02 fr_ where fr_.refcia02 = ops_.refcia01) valor_aduana " & _ 
      ''              "   , (select sum(fr_.prepag02 ) from ssfrac02 fr_ where fr_.refcia02 = ops_.refcia01) valor_comercial " & _
      ''              "   , dta " & _
      ''              "   , iva " & _
      ''              "   , prv " & _
      ''              "   , igi " & _
      ''              "   , total_impuestos " & _
      ''              "   , sum(if(clav21=6,dpg_.mont21,0)) as desconsolidacion " & _
      ''              "   , group_concat(distinct if(clav21=6,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1 ),null)) as des_prov " & _
      ''              "   , sum(if(clav21=10,dpg_.mont21,0)) as almacenaje " & _
      ''              "   , group_concat(distinct if(clav21=10,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1 ),null) ) as alm_prov " & _
      ''              "   , sum(if(clav21=82,dpg_.mont21,0)) as custodia " & _
      ''              "   , group_concat(distinct if(clav21=82,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as cus_prov " & _
      ''              "   , sum(if(clav21=219,dpg_.mont21,0)) as reconocimiento_previo " & _
      ''              "   , group_concat( distinct if(clav21=219,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as rec_prov " & _
      ''              "   , sum(if(clav21=63,dpg_.mont21,0)) as extraordinarios " & _
      ''              "   , group_concat(distinct if(clav21=63,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as ext_prov " & _
      ''              "   , sum(if(clav21=111,dpg_.mont21,0)) as fumigacion " & _
      ''              "   , group_concat(distinct if(clav21=111,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as fum_prov " & _
      ''              "   , sum(dpg_.mont21) as total_paghe " & _
      ''              "    , sum(if(clav21=127,epg_.mont21,0)) as maniobras " & _
      ''              "    , group_concat(distinct if(clav21=127,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as man_prov " & _ 
      ''              "   , sum(if(clav21=11,dpg_.mont21,0)) as montacargas " & _
      ''              "   , group_concat(distinct if(clav21=11,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as mon_prov " &_
      ''              " from " & _
      ''              "   ( " & _
      ''              "   select refcia01, pesobr01, fecpag01, fecent01, cveped01, numped01, patent01, adusec01, 'importacion' as operacion, tipcam01 " & _
      ''              "   , sum( if( c36.cveimp36 = 1, c36.import36, 0)) 'dta' " & _
      ''              "    , sum(if( c36.cveimp36 = 3, c36.import36,0))  'iva' " & _
      ''              "    , sum(if( c36.cveimp36 = 15 , c36.import36,0)) 'prv' " & _ 
      ''              "    , sum(if( c36.cveimp36 = 6 , c36.import36,0)) 'igi' " & _
      ''              "    , sum( c36.import36) as total_impuestos  " & _
      ''              "   from ssdagi01 op_ left join sscont36 c36 on op_.refcia01 = c36.refcia36" & _
      ''              "   where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' " & _
      ''              "   group by refcia01 " & _
      ''              "   union all " & _
      ''              "   select refcia01, pesobr01, fecpag01, fecpre01, cveped01, numped01, patent01, adusec01, 'exportacion', tipcam01 " & _
      ''              "   , sum( if( c36.cveimp36 = 1, c36.import36, 0)) 'dta' " & _
      ''              "    , sum(if( c36.cveimp36 = 3, c36.import36,0))  'iva' " & _
      ''              "    , sum(if( c36.cveimp36 = 15 , c36.import36,0)) 'prv' " & _ 
      ''              "    , sum(if( c36.cveimp36 = 6 , c36.import36,0)) 'igi' " & _
      ''              "    , sum( c36.import36) as total_impuestos  " & _
      ''              "   from ssdage01 op_ left join sscont36 c36 on op_.refcia01 = c36.refcia36 " & _
      ''              "   where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' " & _
      ''              "   group by refcia01 " & _
      ''              "   ) as ops_ " & _
      ''              "   left join ssguia04 guia_ on guia_.refcia04 = ops_.refcia01 and idngui04 = 2 " & _
      ''              "   left join d21paghe dpg_ on dpg_.refe21 = ops_.refcia01 " & _
      ''              "   left join e21paghe epg_ on epg_.foli21 = dpg_.foli21 and year(epg_.fech21) = year(dpg_.fech21) " & _
      ''              "   left join  c21paghe cpg_ on cpg_.clav21 = epg_.conc21 " & _
      ''              " where refcia01 not in (select refe31 from e31cgast join d31refer using(cgas31) where esta31='I' ) " & _
      ''              " group by refcia01 " & _
      ''              " ) as consulta_completa_ "
  sql_ = ""
        ''  sql_ =  " select '' as usuario " & _
        ''              " , '' as estatus " & _
        ''              " , cast(concat('\'',adusec01,patent01,numped01) as char) as Pedimento " & _
        ''              " , refcia01 as Referencia " & _
        ''              " , operacion as Observacion " & _
        ''              " , adusec01 as Aduana " & _
        ''              " , patent01 as Patente " & _
        ''              " , numped01 as 'No Pedimento' " & _
        ''              " ,'' as 'Pedimento Reportado por ABB' " & _
        ''              " , cveped01 as 'Clave Pedimento' " & _
        ''              " , valor_aduana as 'Valor Aduana' " & _
        ''              " , valor_comercial as 'Valor Comercial' " & _
        ''              " , '' as PG " & _
        ''              " , '' as CC " & _
        ''              " , '' as Proyecto " & _
        ''              " , '' as 'Cuenta Estandar' " & _
        ''              " , '' as 'Compania' " & _
        ''              " , prv " & _
        ''              " , dta " & _
        ''              " , igi " & _
        ''              " , iva " & _
        ''              " , '' as 'Subtotal Gastos' " & _
        ''              " , total_impuestos as 'Total Impuestos Pedimento' " & _
        ''              " , '' as Proporcion " & _
        ''              " , pesobr01 as 'Kgs' " & _
        ''              " , fecent01 as 'Fecha Entrada' " & _
        ''              " , fecpag01 as 'Fecha de Pago' " & _
        ''              " , datediff(fecpag01, fecent01) as 'Dias para calculo' " & _
        ''              " , guia as 'Guia  House' " & _
        ''              " , custodia " & _
        ''              " , cus_prov " & _
        ''              " , maniobras as 'Maniobra/Manejo' " & _
        ''              " , man_prov as manejo_proveedor " & _
        ''              " , almacenaje " & _
        ''              " , alm_prov " & _
        ''              " , montacargas " & _
        ''              " , mon_prov"  & _ 
        ''              " , desconsolidacion " & _
        ''              " , des_prov"  & _ 
        ''              " , cast(tipcam01 as char) as 'Tipo de Cambio' " & _
        ''              " , reconocimiento_previo " & _
        ''              " , rec_prov " & _
        ''              " , extraordinarios " & _
        ''              " , ifnull(ext_prov,'') as ext_prov " & _
        ''              " , fumigacion " & _
        ''              " , fum_prov " & _
        ''              " , total_impuestos as 'Impuestos segun Pedimento' " & _
        ''              " , total_paghe " & _
        ''              " , '' as 'Base para el Pago de Honorarios' " & _
        ''              " , '' as 'Honorarios' "  & _
        ''              " , '' as 'Complementarios' " & _
        ''              " , '' as 'Embalaje' " & _
        ''              " , '' as 'Validacion' " & _
        ''              " , '' as 'Proveedor Transporte' " & _
        ''              " , '' as 'Servicios por Operacion' " & _
        ''              " , '' as 'Proveedor Servicios Operacion' " & _
        ''              " , extraordinarios " & _
        ''              " , '' as 'Sub Total Honorarios '" & _
        ''              " , '' as 'IVA Honorarios' " & _
        ''              " , '' as 'Total Honorarios' " & _
        ''              " , '' as '' " & _
        ''               "    " & _
        ''               " from( " & _
        ''               "  select " & _
        ''               "    refcia01 " & _
        ''               "    , operacion " & _
        ''               "    , pesobr01 " & _
        ''               "    , fecpag01 " & _
        ''               "    , fecent01 " & _
        ''               "    , cveped01 " & _
        ''               "    , numped01 " & _
        ''               "    , patent01 " & _
        ''               "    , adusec01 " & _
        ''               "    , tipcam01 " & _
        ''               "    , group_concat(distinct numgui04) as guia " & _
        ''               "    , (select sum(fr_.vaduan02) from ssfrac02 fr_ where fr_.refcia02 = ops_.refcia01) valor_aduana " & _
        ''               "    , (select sum(fr_.prepag02 ) from ssfrac02 fr_ where fr_.refcia02 = ops_.refcia01) valor_comercial " & _
        ''               "    , prv " & _
        ''               "    , dta " & _
        ''               "    , igi " & _
        ''               "    , iva  " & _
        ''               "    , custodia " & _
        ''               "    , cus_prov " & _
        ''               "    , maniobras " & _
        ''               "    , manejo  " & _
        ''               "    , man_prov  " & _
        ''               "    , almacenaje " & _
        ''               "    , alm_prov " & _
        ''               "    , montacargas " & _
        ''               "    , mon_prov " & _
        ''               "    , extraordinarios " & _
        ''               "    , ext_prov " & _
        ''               "    , fumigacion " & _
        ''               "    , fum_prov " & _
        ''               "    , total_paghe " & _
        ''               "    , desconsolidacion   " & _
        ''               "    , des_prov " & _
        ''               "    , total_impuestos " & _
        ''               "    , reconocimiento_previo " & _
        ''               "    , reconocimiento_previo_prov as  rec_prov " & _
        ''               "  from " & _
        ''               "    ( " & _
        ''               "    select refcia01, pesobr01, fecpag01, fecent01, cveped01, numped01, patent01, adusec01, 'importacion' as operacion, tipcam01 " & _
        ''               "    , sum( if( c36.cveimp36 = 1, c36.import36, 0)) 'dta' " & _
        ''               "    , sum(if( c36.cveimp36 = 3, c36.import36,0))  'iva' " & _
        ''               "    , sum(if( c36.cveimp36 = 15 , c36.import36,0)) 'prv' " & _
        ''               "    , sum(if( c36.cveimp36 = 6 , c36.import36,0)) 'igi' " & _
        ''               "    , sum( c36.import36) as total_impuestos  " & _
        ''               "    from ssdagi01 op_ left join sscont36 c36 on op_.refcia01 = c36.refcia36 " & _
        ''               "    where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' and cveped01 != 'R1' " & _
        ''               "    group by refcia01 " & _
        ''               "    union all " & _
        ''               "    select refcia01, pesobr01, fecpag01, fecpre01, cveped01, numped01, patent01, adusec01, 'exportacion', tipcam01, sum( if( c36.cveimp36 = 1, c36.import36, 0)) 'dta' " & _
        ''               "    , sum(if( c36.cveimp36 = 3, c36.import36,0))  'iva' " & _
        ''               "    , sum(if( c36.cveimp36 = 15 , c36.import36,0)) 'prv' " & _
        ''               "    , sum(if( c36.cveimp36 = 6 , c36.import36,0)) 'igi' " & _
        ''               "    , sum( c36.import36) as total_impuestos  " & _
        ''               "    from ssdage01 op_ left join sscont36 c36 on op_.refcia01 = c36.refcia36 " & _
        ''               "    where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' and cveped01 != 'R1' " & _
        ''               "    group by refcia01 " & _
        ''               "    ) as ops_ " & _
        ''               "    left join ssguia04 guia_ on guia_.refcia04 = ops_.refcia01 and idngui04 = 2 " & _
        ''               "    left join  " & _
        ''               "    ( " & _
        ''               "      select refcia01 " & _
        ''               "        , sum(if(clav21=6,dpg_.mont21,0)) as desconsolidacion " & _
        ''               "            , group_concat(distinct if(clav21=6,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1 ),null)) as des_prov " & _
        ''               "            , sum(if(clav21=10,dpg_.mont21,0)) as almacenaje " & _
        ''               "            , group_concat(distinct if(clav21=10,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1 ),null) ) as alm_prov " & _
        ''               "            , sum(if(clav21=82,dpg_.mont21,0)) as custodia " & _
        ''               "            , group_concat(distinct if(clav21=82,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as cus_prov " & _
        ''               "            , sum(if(clav21=219 or clav21 =102 ,dpg_.mont21,0)) as reconocimiento_previo " & _
        ''               "            , group_concat( distinct if(clav21=219,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),if(clav21=102,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll))) as reconocimiento_previo_prov " & _
        ''               "            , sum(if(clav21=63,dpg_.mont21,0)) as extraordinarios " & _
        ''               "            , group_concat(distinct if(clav21=63,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as ext_prov " & _
        ''               "            , sum(if(clav21=111,dpg_.mont21,0)) as fumigacion " & _
        ''               "            , group_concat(distinct if(clav21=111,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as fum_prov " & _
        ''               "            , sum(dpg_.mont21) as total_paghe " & _
        ''               "            , sum(if(clav21=11,dpg_.mont21,0)) as montacargas " & _
        ''               "            , group_concat(distinct if(clav21=11,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as mon_prov " & _
        ''               "            , sum(if(clav21=127 ,dpg_.mont21,0)) as maniobras " & _
        ''               "            , sum(if(clav21=141,dpg_.mont21 ,0)) as manejo " & _
        ''               "            , group_concat(distinct if(clav21=127,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1), if(clav21=141,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1) ,Null) )) as man_prov " & _
        ''               "        from  (select refcia01 from ssdagi01 where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' union all  " & _
        ''               "            select refcia01 from ssdage01 where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' ) as ops_ " & _
        ''               "        left join d21paghe dpg_ on dpg_.refe21 = ops_.refcia01 " & _
        ''               "            left join e21paghe epg_ on dpg_.foli21 = epg_.foli21 and year(epg_.fech21) = year(dpg_.fech21) " & _
        ''               "            left join  c21paghe cpg_ on cpg_.clav21 = epg_.conc21 " & _
        ''               "        where refcia01 not in (select refe31 from e31cgast join d31refer using(cgas31) where esta31='I' )  " & _
        ''               "        group by refcia01 " & _
        ''               "    ) as paghes_ using (refcia01) " & _
        ''               "  where refcia01 not in (select refe31 from e31cgast join d31refer using(cgas31) where esta31='I' ) " & _
        ''               "  group by refcia01 " & _
        ''               " ) as consulta_completa_ "
' Consulta modificada Junio 2013

  sql_ = " select '' as usuario  " & _ 
        "  , '' as estatus  " & _ 
        "  , cast(concat('\'',adusec01,patent01,numped01) as char) as Pedimento  " & _ 
        "  , refcia01 as Referencia  " & _ 
        "  , operacion as Observacion  " & _ 
        "  , adusec01 as Aduana  " & _ 
        "  , patent01 as Patente  " & _ 
        "  , numped01 as 'No Pedimento'  " & _ 
        "  ,'' as 'Pedimento Reportado por ABB'  " & _ 
        "  , cveped01 as 'Clave Pedimento'  " & _ 
        "  , valor_aduana as 'Valor Aduana'  " & _ 
        "  , valor_comercial as 'Valor Comercial'  " & _ 
        "  , '' as PG  " & _ 
        "  , '' as CC  " & _ 
        "  , '' as Proyecto  " & _ 
        "  , '' as 'Cuenta Estandar'  " & _ 
        "  , '' as 'Compania'  " & _ 
        "  , prv  " & _ 
        "  , dta  " & _ 
        "  , igi  " & _ 
        "  , iva  " & _ 
        "  , '' as 'Subtotal Gastos'  " & _ 
        "  , total_impuestos as 'Total Impuestos Pedimento'  " & _ 
        "  , '' as Proporcion  " & _ 
        "  , pesobr01 as 'Kgs'  " & _ 
        "  , fecent01 as 'Fecha Entrada'  " & _ 
        "  , fecpag01 as 'Fecha de Pago'  " & _ 
        "  , datediff(fecpag01, fecent01) as 'Dias para calculo'  " & _ 
        "  , guia as 'Guia  House' " & _ 
        "  , cus_rfc as 'Custoria RFC' " & _ 
        "  , cus_fac as  'Custodia Factura' " & _ 
        "  , custodia  " & _ 
        "  , cus_prov as 'Custodia Proveedor' " & _ 
        "  , man_rfc as 'Manejo RFC' " & _ 
        "  , man_fac as 'Manejo Factura' " & _ 
        "  , maniobras as 'Maniobra/Manejo'  " & _ 
        "  , man_prov as manejo_proveedor  " & _ 
        "  , alm_rfc as 'Almacenaje RFC' " & _ 
        "  , alm_fac as 'Almacenaje Factura' " & _ 
        "  , almacenaje  " & _ 
        "  , alm_prov as 'Almacenaje Proveedor' " & _ 
        "  , mon_rfc as 'Montacargas RFC' " & _ 
        "  , mon_fac as 'Montacargas Factura' " & _ 
        "  , montacargas  " & _ 
        "  , mon_prov as 'Montacargas Proveedor' " & _ 
        "  , des_rfc as 'Desconsolidación RFC' " & _ 
        "  , des_fac as 'Desconsolidación Factura' " & _ 
        "  , desconsolidacion  " & _ 
        "  , des_prov as 'Desconsolidación Proveedor' " & _ 
        "  , cast(tipcam01 as char) as 'Tipo de Cambio'  " & _ 
        "  , rec_rfc as 'Previo RFC' " & _ 
        "  , rec_fac as 'Previo Factura' " & _ 
        "  , reconocimiento_previo  as 'Previo' " & _ 
        "  , rec_prov  as 'Previo Proveedor' " & _ 
        "  , ext_rfc as 'Extraordinarios RFC' " & _ 
        "  , ext_fac as 'Extraordinarios Factura' " & _ 
        "  , extraordinarios  " & _ 
        "  , ifnull(ext_prov,'') as 'Extraordinarios Proveedor' " & _ 
        "  , fum_rfc as 'Fumigacion RFC' " & _ 
        "  , fum_fac as 'Fumigacion Factura' " & _ 
        "  , fumigacion  " & _ 
        "  , fum_prov as 'Fumigacion Proveedor' " & _ 
        "  , total_impuestos as 'Impuestos segun Pedimento'  " & _ 
        "  , total_paghe  as 'Total Pagos Hechos' " & _ 
        "  , '' as 'Base para el Pago de Honorarios'  " & _ 
        "  , '' as 'Honorarios'  " & _ 
        "  , '' as 'Complementarios'  " & _ 
        "  , '' as 'Embalaje'  " & _ 
        "  , '' as 'Validacion'  " & _ 
        "  , '' as 'Proveedor Transporte'  " & _ 
        "  , '' as 'Servicios por Operacion'  " & _ 
        "  , '' as 'Proveedor Servicios Operacion'  " & _ 
        "  , extraordinarios as 'Extraordinarios Impte' " & _ 
        "  , '' as 'Sub Total Honorarios ' " & _ 
        "  , '' as 'IVA Honorarios'  " & _ 
        "  , '' as 'Total Honorarios'  " & _ 
        "  , '' as ''  " & _ 
        "      " & _ 
        "  from(  " & _ 
        "   select  " & _ 
        "     refcia01  " & _ 
        "     , operacion  " & _ 
        "     , pesobr01  " & _ 
        "     , fecpag01  " & _ 
        "     , fecent01  " & _ 
        "     , cveped01  " & _ 
        "     , numped01  " & _ 
        "     , patent01  " & _ 
        "     , adusec01  " & _ 
        "     , tipcam01  " & _ 
        "     , group_concat(distinct numgui04) as guia  " & _ 
        "     , (select sum(fr_.vaduan02) from ssfrac02 fr_ where fr_.refcia02 = ops_.refcia01) valor_aduana  " & _ 
        "     , (select sum(fr_.prepag02 ) from ssfrac02 fr_ where fr_.refcia02 = ops_.refcia01) valor_comercial  " & _ 
        "     , prv  " & _ 
        "     , dta  " & _ 
        "     , igi  " & _ 
        "     , iva " & _ 
        "    , cus_rfc " & _ 
        "    , cus_fac   " & _ 
        "     , custodia  " & _ 
        "     , cus_prov  " & _ 
        "     , man_rfc " & _ 
        "    , man_fac " & _ 
        "     , maniobras  " & _ 
        "     , manejo   " & _ 
        "     , man_prov   " & _ 
        "     , alm_rfc " & _ 
        "    , alm_fac " & _ 
        "     , almacenaje  " & _ 
        "     , alm_prov  " & _ 
        "     , mon_rfc " & _ 
        "    , mon_fac " & _ 
        "     , montacargas  " & _ 
        "     , mon_prov  " & _ 
        "     , ext_rfc " & _ 
        "    , ext_fac " & _ 
        "     , extraordinarios  " & _ 
        "     , ext_prov  " & _ 
        "     , fum_rfc " & _ 
        "    , fum_fac " & _ 
        "     , fumigacion  " & _ 
        "     , fum_prov  " & _ 
        "     , total_paghe  " & _ 
        "     , des_rfc " & _ 
        "    , des_fac " & _ 
        "     , desconsolidacion    " & _ 
        "     , des_prov  " & _ 
        "     , total_impuestos " & _ 
        "    , rec_rfc " & _ 
        "    , rec_fac  " & _ 
        "     , reconocimiento_previo  " & _ 
        "     , reconocimiento_previo_prov as  rec_prov  " & _ 
        "   from  " & _ 
        "     (  " & _ 
        "     select refcia01, pesobr01, fecpag01, fecent01, cveped01, numped01, patent01, adusec01, 'importacion' as operacion, tipcam01  " & _ 
        "     , sum( if( c36.cveimp36 = 1, c36.import36, 0)) 'dta'  " & _ 
        "     , sum(if( c36.cveimp36 = 3, c36.import36,0))  'iva'  " & _ 
        "     , sum(if( c36.cveimp36 = 15 , c36.import36,0)) 'prv'  " & _ 
        "     , sum(if( c36.cveimp36 = 6 , c36.import36,0)) 'igi'  " & _ 
        "     , sum( c36.import36) as total_impuestos   " & _ 
        "     from ssdagi01 op_ left join sscont36 c36 on op_.refcia01 = c36.refcia36  " & _ 
        "     where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' and cveped01 != 'R1'  " & _ 
        "     group by refcia01  " & _ 
        "     union all  " & _ 
        "     select refcia01, pesobr01, fecpag01, fecpre01, cveped01, numped01, patent01, adusec01, 'exportacion', tipcam01, sum( if( c36.cveimp36 = 1, c36.import36, 0)) 'dta'  " & _ 
        "     , sum(if( c36.cveimp36 = 3, c36.import36,0))  'iva'  " & _ 
        "     , sum(if( c36.cveimp36 = 15 , c36.import36,0)) 'prv'  " & _ 
        "     , sum(if( c36.cveimp36 = 6 , c36.import36,0)) 'igi'  " & _ 
        "     , sum( c36.import36) as total_impuestos   " & _ 
        "     from ssdage01 op_ left join sscont36 c36 on op_.refcia01 = c36.refcia36  " & _ 
        "     where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' and cveped01 != 'R1'  " & _ 
        "     group by refcia01  " & _ 
        "     ) as ops_  " & _ 
        "     left join ssguia04 guia_ on guia_.refcia04 = ops_.refcia01 and idngui04 = 2  " & _ 
        "     left join   " & _ 
        "     (  " & _ 
        "       select refcia01  " & _ 
        "             , sum(if(clav21=6,dpg_.mont21,0)) as desconsolidacion  " & _ 
        "             , group_concat(distinct if(clav21=6,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1 ),null)) as des_prov  " & _ 
        "             , group_concat(distinct if(clav21=6,(select distinct a.rfc20 from c20benef a where bene21 = a.clav20 limit 1) , null)) as des_rfc " & _ 
        "         , group_concat(distinct if(clav21=6, dpg_.facpro21, null)) as des_fac " & _ 
        "          " & _ 
        "             , sum(if(clav21=10,dpg_.mont21,0)) as almacenaje  " & _ 
        "             , group_concat(distinct if(clav21=10,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1 ),null) ) as alm_prov  " & _ 
        "             , group_concat(distinct if(clav21=10,(select distinct a.rfc20 from c20benef a where bene21 = a.clav20 limit 1) , null)) as alm_rfc " & _ 
        "         , group_concat(distinct if(clav21=10, dpg_.facpro21, null)) as alm_fac " & _ 
        "              " & _ 
        "             , sum(if(clav21=82,dpg_.mont21,0)) as custodia  " & _ 
        "             , group_concat(distinct if(clav21=82,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as cus_prov  " & _ 
        "             , group_concat(distinct if(clav21=82,(select distinct a.rfc20 from c20benef a where bene21 = a.clav20 limit 1) , null)) as cus_rfc " & _ 
        "         , group_concat(distinct if(clav21=82, dpg_.facpro21, null)) as cus_fac " & _ 
        "              " & _ 
        "             , sum(if(clav21=219 or clav21 =102 ,dpg_.mont21,0)) as reconocimiento_previo  " & _ 
        "             , group_concat( distinct if(clav21=219,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),if(clav21=102,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll))) as reconocimiento_previo_prov  " & _ 
        "             , group_concat(distinct if(clav21=219,(select distinct a.rfc20 from c20benef a where bene21 = a.clav20 limit 1),if(clav21=102,(select distinct a.rfc20 from c20benef a where bene21 = a.clav20 limit 1), null))) as rec_rfc " & _ 
        "         , group_concat(distinct if(clav21=219 or clav21=102, dpg_.facpro21, null)) as rec_fac " & _ 
        "              " & _ 
        "             , sum(if(clav21=63,dpg_.mont21,0)) as extraordinarios  " & _ 
        "             , group_concat(distinct if(clav21=63,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as ext_prov  " & _ 
        "             , group_concat(distinct if(clav21=63,(select distinct a.rfc20 from c20benef a where bene21 = a.clav20 limit 1) , null)) as ext_rfc " & _ 
        "         , group_concat(distinct if(clav21=63, dpg_.facpro21, null)) as ext_fac " & _ 
        "              " & _ 
        "             , sum(if(clav21=111,dpg_.mont21,0)) as fumigacion  " & _ 
        "             , group_concat(distinct if(clav21=111,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as fum_prov  " & _ 
        "             , group_concat(distinct if(clav21=111,(select distinct a.rfc20 from c20benef a where bene21 = a.clav20 limit 1) , null)) as fum_rfc " & _ 
        "         , group_concat(distinct if(clav21=111, dpg_.facpro21, null)) as fum_fac " & _ 
        "              " & _ 
        "             , sum(dpg_.mont21) as total_paghe  " & _ 
        "              " & _ 
        "             , sum(if(clav21=11,dpg_.mont21,0)) as montacargas  " & _ 
        "             , group_concat(distinct if(clav21=11,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1),NUll)) as mon_prov  " & _ 
        "             , group_concat(distinct if(clav21=11,(select distinct a.rfc20 from c20benef a where bene21 = a.clav20 limit 1) , null)) as mon_rfc " & _ 
        "         , group_concat(distinct if(clav21=11, dpg_.facpro21, null)) as mon_fac " & _ 
        "              " & _ 
        "             , sum(if(clav21=127 ,dpg_.mont21,0)) as maniobras  " & _ 
        "             , sum(if(clav21=141,dpg_.mont21 ,0)) as manejo  " & _ 
        "             , group_concat(distinct if(clav21=127,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1), if(clav21=141,(select replace(nomb20,'.','') from c20benef a where bene21 = a.clav20 limit 1) ,Null) )) as man_prov  " & _ 
        "             , group_concat(distinct if(clav21=127,(select distinct a.rfc20 from c20benef a where bene21 = a.clav20 limit 1) , null)) as man_rfc " & _ 
        "         , group_concat(distinct if(clav21=127, dpg_.facpro21, Null)) as man_fac " & _ 
        "              " & _ 
        "         from  (select refcia01 from ssdagi01 where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' union all   " & _ 
        "             select refcia01 from ssdage01 where fecpag01 > '20130101' and firmae01 != '' and rfccli01 = 'AME920102SS4' ) as ops_  " & _ 
        "         left join d21paghe dpg_ on dpg_.refe21 = ops_.refcia01  " & _ 
        "             left join e21paghe epg_ on dpg_.foli21 = epg_.foli21 and year(epg_.fech21) = year(dpg_.fech21)  " & _ 
        "             left join  c21paghe cpg_ on cpg_.clav21 = epg_.conc21  " & _ 
        "         where refcia01 not in (select refe31 from e31cgast join d31refer using(cgas31) where esta31='I' )   " & _ 
        "         group by refcia01  " & _ 
        "     ) as paghes_ using (refcia01)  " & _ 
        "   where refcia01 not in (select refe31 from e31cgast join d31refer using(cgas31) where esta31='I' )  " & _ 
        "   group by refcia01  " & _ 
        "  ) as consulta_completa_  "
	genera_sql = sql_
	
end function 

' Este metodo Recorre la consulta y la genera en una tabla HTML
'function genera_xls( _obj_cursor)

'end function

function genera_encabezado()
	'cabecera_ = "<thead><tr bgcolor= ""yellow""> " &_
	''			"<td >Usuario</td>" &_
	''			"<td >Estatus</td>" &_
	''			"<td >Pedimento </td>" &_
	''			"<td >Referencia </td>" &_
	''			"<td >Observacion </td>" &_
	''			"<td >Aduana </td>" &_
	''			"<td >Patente </td>" &_
	''			"<td >No Pedimento </td>" &_
  ''     "<td >Pedimento Reportado por ABB </td>" &_
  ''    "<td >Clave Pedimento </td>" &_
  ''      "<td >Valor Aduana </td>" &_
  ''      "<td >Valor Comercial </td>" &_
  ''      "<td >PG </td>" &_
  ''      "<td >CC </td>" &_
  ''      "<td >Proyecto </td>" &_
  ''      "<td >Cuenta Estandar </td>" &_
  ''      "<td >Compania </td>" &_
  ''      "<td >PRV </td>" &_
  ''      "<td >DTA </td>" &_
  ''      "<td >IGI </td>" &_
  ''      "<td >IVA </td>" &_
  ''      "<td >Sub Total Gastos </td>" &_
  ''      "<td >Total Impuestos Pedimento </td>" &_
  ''      "<td >Proporcion </td>" &_
  ''      "<td >Kgs </td>" &_
  ''      "<td >Fecha Entrada </td>" &_
  ''      "<td >Fecha de Pago </td>" &_
  ''      "<td >Dias para el Calculo </td>" &_
  ''      "<td >Guia House </td>" &_
  ''      "<td >Custodia </td>" &_
  ''      "<td >Proveedor </td>" &_
  ''      "<td >Manejo RFC </td>" &_
  ''      "<td >Custodia Factura </td>" &_
  ''      "<td >Maniobra/Manejo </td>" &_
  ''      "<td >Proveedor </td>" &_
  ''      "<td >Almacenaje </td>" &_
  ''      "<td >Proveedor </td>" &_
  ''      "<td >Montacargas </td>" &_
  ''      "<td >Proveedor </td>" &_
  ''      "<td >Desconsolidacion </td>" &_
  ''      "<td >Proveedor </td>" &_
  ''      "<td >Tipo de Cambio </td>" &_
  ''      "<td >Reconocimiento Previo </td>" &_
  ''      "<td >Proveedor </td>" &_
  ''      "<td >Servicio Extraordinario </td>" &_
  ''      "<td >Proveedor </td>" &_
  ''      "<td >Fumigacion </td>" &_
  ''      "<td >Proveedor </td>" &_
  ''      "<td >Impuestos segun Pedimentos </td>" &_
  ''      "<td >Total de Pagos Hechos </td>" &_
  ''      "<td >Base para el cobro Honorarios </td>" &_
  ''      "<td >Honorarios </td>" &_
  ''      "<td >Complementarios </td>" &_
  ''      "<td >Embalaje </td>" &_
  ''      "<td >Validacion </td>" &_
  ''      "<td >Proveedor Transporte </td>" &_
  ''      "<td >Servicios por Operacion </td>" &_
  ''      "<td >Proveedor </td>" &_
  ''      "<td >Servicio Extraordinario </td>" &_
  ''      "<td >Sub Total Honorarios </td>" &_
  ''      "<td >Iva Honorarios </td>" &_
  ''      "<td >Total Honorarios </td>" &_
	''			"</tr></thead>"

  cabecera_ = "<thead><tr bgcolor= ""yellow""> " &_
        "<td >Usuario</td>" &_
        "<td >Estatus</td>" &_
        "<td >Pedimento </td>" &_
        "<td >Referencia </td>" &_
        "<td >Observacion </td>" &_
        "<td >Aduana </td>" &_
        "<td >Patente </td>" &_
        "<td >No Pedimento </td>" &_
        "<td >Pedimento Reportado por ABB </td>" &_
        "<td >Clave Pedimento </td>" &_
        "<td >Valor Aduana </td>" &_
        "<td >Valor Comercial </td>" &_
        "<td >PG </td>" &_
        "<td >CC </td>" &_
        "<td >Proyecto </td>" &_
        "<td >Cuenta Estandar </td>" &_
        "<td >Compania </td>" &_
        "<td >PRV </td>" &_
        "<td >DTA </td>" &_
        "<td >IGI </td>" &_
        "<td >IVA </td>" &_
        "<td >Sub Total Gastos </td>" &_
        "<td >Total Impuestos Pedimento </td>" &_
        "<td >Proporcion </td>" &_
        "<td >Kgs </td>" &_
        "<td >Fecha Entrada </td>" &_
        "<td >Fecha de Pago </td>" &_
        "<td >Dias para el Calculo </td>" &_
        "<td >Guia House </td>" &_
        "<td >Custodia RFC </td>" &_
        "<td >Custodia Factura </td>" &_
        "<td >Custodia </td>" &_
        "<td >Proveedor </td>" &_
        "<td >Manejo RFC </td>" &_
        "<td >Manejo Factura </td>" &_
        "<td >Maniobra/Manejo </td>" &_
        "<td >Proveedor </td>" &_
        "<td >Almacenaje RFC </td>" &_
        "<td >Almacenaje Factura </td>" &_
        "<td >Almacenaje </td>" &_
        "<td >Proveedor </td>" &_
        "<td >Montacargas RFC </td>" &_
        "<td >Montacargas Factura </td>" &_
        "<td >Montacargas </td>" &_
        "<td >Proveedor </td>" &_
        "<td >Desconsolidacion RFC </td>" &_
        "<td >Desconsolidacion Factura </td>" &_
        "<td >Desconsolidacion </td>" &_
        "<td >Proveedor </td>" &_
        "<td >Tipo de Cambio </td>" &_
        "<td >Reconocimiento Previo </td>" &_
        "<td >Reconocimiento Previo RFC </td>" &_
        "<td >Reconocimiento Previo Factura </td>" &_
        "<td >Proveedor </td>" &_
        "<td >Servicio Extraordinario RFC </td>" &_
        "<td >Servicio Extraordinario Factura </td>" &_
        "<td >Servicio Extraordinario </td>" &_
        "<td >Proveedor </td>" &_
        "<td >Fumigacion RFC </td>" &_
        "<td >Fumigacion Factura </td>" &_
        "<td >Fumigacion </td>" &_
        "<td >Proveedor </td>" &_
        "<td >Impuestos segun Pedimentos </td>" &_
        "<td >Total de Pagos Hechos </td>" &_
        "<td >Base para el cobro Honorarios </td>" &_
        "<td >Honorarios </td>" &_
        "<td >Complementarios </td>" &_
        "<td >Embalaje </td>" &_
        "<td >Validacion </td>" &_
        "<td >Proveedor Transporte </td>" &_
        "<td >Servicios por Operacion </td>" &_
        "<td >Proveedor </td>" &_
        "<td >Servicio Extraordinario </td>" &_
        "<td >Sub Total Honorarios </td>" &_
        "<td >Iva Honorarios </td>" &_
        "<td >Total Honorarios </td>" &_
        "</tr></thead>"

	genera_encabezado = cabecera_
	
end function



function genera_cuerpo(sql)
	
	Set ConnStr = Server.CreateObject ("ADODB.Connection")
	ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; DATABASE="& db_ &"_extranet; OPTION=16427"
		
	Set RS = CreateObject("ADODB.RecordSet")
	Set RS = ConnStr.Execute(sql)
	tabla_ = ""
    If Not RS.EOF Then
		
		' Variable de bandera para resaltar los renglones pares
		ban_ = 1
        Do Until RS.EOF
			if ban_ = 1 then
				tabla_ = tabla_ & "<tr>"
				ban_ = 0
			else
				tabla_ = tabla_ & "<tr bgcolor= ""#EBECF2"">"
				ban_ = 1
			end if
			'tabla_ = tabla_ & "<td align=""center""> " & RS.fields(0) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(1) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(2) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(3) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(4) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(5) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(6) & "</td>"
			'tabla_ = tabla_ & "<td align=""right"">" & RS.fields(7) & "</td>"
			'tabla_ = tabla_ & "<td align=""center"">" & RS.fields(8) & "</td>"
			'on error resume next
			'	tabla_ = tabla_ & "<td >" & RS.fields(9) & "</td>"
			'	tabla_ = tabla_ & "<td >" & RS.fields(10) & "</td>"
			'Next
			' Recorrido de los campos de la fila
			For i = 0 To RS.fields.Count-1
				on error resume next
					tabla_ = tabla_ & "<td align=""center"">" & RS.fields(i) & "</td>"
				if Err.Number<> 0 Then
					tabla_ = tabla_ & "<td align=""center"">" & VarType(RS.fields(i)) & "</td>"

				end if
				Err.Clear
            Next
			tabla_ = tabla_ & "</tr>" &VBcrlf
           RS.MoveNext
        Loop
    End If
    RS.Close
    Set RS=Nothing
	genera_cuerpo = tabla_
 End function


 %>

 <HTML>
	<HEAD>
		<TITLE>::.... REPORTE DE CUENTAS CONTABLES GROUPE SEB .... ::</TITLE>
	</HEAD>
	<BODY>
	<%=html_%>
	</BODY>
</HTML>