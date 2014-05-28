
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<% 
    'serv="10.66.1.5"
    'base_datos="dai_extranet"
    'usu="carlosmg"
    'pass="123456"
    'MM_STRING = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER="&serv&"; DATABASE="&base_datos&"; UID="&usu&"; PWD="&pass&"; OPTION=16427"



    Response.Buffer = TRUE
    Response.Addheader "Content-Disposition", "attachment;filename=HoneyWell.xls"
    Response.ContentType = "application/vnd.ms-excel"
    Server.ScriptTimeOut=100000

    strUsuario     = request.Form("user")
    strTipoUsuario = request.Form("TipoUser")
    strPermisos    = Request.Form("Permisos")
    permi          = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
    permi2         = PermisoClientesTabla("B",Session("GAduana") ,strPermisos,"clie31")

    if not permi2 = "" then
      permi2 = "  and (" & permi2 & ") "
    end if

    AplicaFiltro = false
    strFiltroCliente = ""
    strFiltroCliente = request.Form("txtCliente")
    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
       blnAplicaFiltro = true
    end if
    if blnAplicaFiltro then
       permi2 = " AND B.clie31 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
       permi2 = ""
    end if

    if not permi = "" then
      permi = "  and (" & permi & ") "
    end if

    AplicaFiltro = false
    'strFiltroCliente = ""
    'strFiltroCliente = request.Form("txtCliente")
    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
       blnAplicaFiltro = true
    end if
    if blnAplicaFiltro then
       permi = " AND cvecli01 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
       permi = ""
    end if

    '*****************************************************************************************************************************
    strDateIni = ""
    strDateFin = ""
    strTipoPedimento = ""
    strCodError      = "0"

    '*******************************************************
    strTipoFiltro         = trim(request.Form("TipoFiltro"))
    strDateIni            = trim(request.Form("txtDateIni"))
    strDateFin            = trim(request.Form("txtDateFin"))
    strTipoOperacion      = trim(request.Form("rbnTipoDate"))


    'strLinNav             = trim(request.Form("txtLinNav"))
    'strModalidad          = trim(request.Form("txtMod"))
    'strProv               = trim(request.Form("txtProv"))
    'strfiltrosrestantes   = trim(request.Form("txtfiltrosrestantes"))
    'strTipoOtrosFiltros   = trim(request.Form("txttipoOtrosFiltros"))

    'strTipoConte = trim(request.Form("txttipoConte"))
    '*******************************************************

    if not isdate(strDateIni) then
      strCodError = "5"
    end if
    if not isdate(strDateFin) then
      strCodError = "6"
    end if
    if strDateIni="" or strDateFin="" then
      strCodError = "1"
    end if

    'Response.Write(strTipoFiltro)


    'strHTML = ""


    if strCodError = "0" then

    tempstrOficina = adu_ofi( Session("GAduana") )
    IF NOT enproceso(tempstrOficina) THEN

    '******************************************************************************************************

'+----------+-----------------------------------------------+--------------+
'| cvecli18 | nomcli18       DAI                            | rfccli18     |
'+----------+-----------------------------------------------+--------------+
'|     3297 | HONEYWELL OPERATIVA MEXICO S. DE R.L. DE C.V. | HOM050106JIA |
'|     3309 | HONEYWELL S.A. DE C.V.                        | HON641119JI7 |
'+----------+-----------------------------------------------+--------------+

'ISTRFINI="2008-01-1"
'FSTRFFIN="2008-07-30"


'response.Write(ISTRFINI)
'response.Write(FSTRFFIN)
'Response.End

' CCLIENTE="3297 or cvecli01=3309 "
' CCLIENTE2="3297 or cvecli18=3309 "

'ISTRFINI="2008-01-09"
'FSTRFFIN="2008-01-09"
'CCLIENTE=1188

ISTRFINI = strDateIni
FSTRFFIN = strDateFin

  'base="MEX"
  base = Session("GAduana")
  Select Case (base)
   Case "VER":
       c_fleteTerr="15"
       c_man="2"
       c_fleteAereo="35"
       c_desconsolidado="85"
       c_maniobrasyalmacenajes="230"
   Case "MEX":
       c_fleteTerr="7"
       c_man="127"
       c_fleteAereo="3"
       c_desconsolidado="6"
       c_maniobrasyalmacenajes="2"
   Case "MAN":
       c_fle="5"
       c_man="2"
   Case "TAM":
      c_fle="15"
       c_man="2"
   Case "LZR":
       c_fle="15"
       c_man="2"
   Case "GUA":
       c_fle="15"
       c_man="2"
   Case "LAR":
       c_fle="15"
       c_man="2"
       c_fleteTerr="142"
       c_fleteAereo="129"
       c_desconsolidado="109"
       c_maniobrasyalmacenajes="130"
   End Select


  top="i"
  if strTipoOperacion = 1 then
    tabla       = "ssdagi01"
    num_con_f_m = c_fleteTerr
    fecEntAdu   = "fecent01"
    pTipOpe     = "Importación"
  else
    tabla       = "ssdage01"
    num_con_f_m = c_fleteTerr
    fecEntAdu   = "fecpre01"
    pTipOpe     = "Exportación"
  end if


'Response.Write(tabla)
'Response.End

'sqlReferencias=" Select refcia01,numped01,(i_dta101+i_dta201) as dta, cvepro01, cvepod01,cvepvc01, "&_
'                "tipcam01,fecpag01,"&fecEntAdu&" as fechaEA, firmae01,desf0101,pesobr01,valmer01,segros01,fletes01, sum(fletes01+segros01) as fleyseg,cveped01,"&_
'                "desdoc01,cveadu01,cvepfm01,cvecli01,"&_
'                "fech31,d.cgas31 as cg "&_
'                " FROM  (e31cgast as e inner join d31refer as d on e.cgas31=d.cgas31) "&_
'                " inner join "&tabla&" on d.refe31=refcia01 "&_
'                " where firmae01<>'' and  cveped01<>'R1' and fech31>='"&ISTRFINI&"' and fech31<='"&FSTRFFIN&"'  "&_
'                " and esta31 = 'I' and cvecli01="&CCLIENTE&" "&_
'                "group by refcia01 "

sqlReferencias= " Select refcia01, " &_
                "        numped01," &_
                "        (i_dta101+i_dta201) as dta, " &_
                "        cvepro01, " &_
                "        cvepod01, " &_
                "        cvepvc01, " &_
                "        tipcam01, " &_
                "        DATE_FORMAT(fecpag01, '%d/%m/%Y') AS 'fecpag01', " & _
                "        DATE_FORMAT(" & fecEntAdu & ", '%d/%m/%Y') as fechaEA, " &_
                "        firmae01, " &_
                "        desf0101, " &_
                "        pesobr01, " &_
                "        segros01, " &_
                "        fletes01, " &_
                "        (fletes01+segros01) as fleyseg," &_
                "        cveped01, " &_
                "        desdoc01, " &_
                "        cveadu01, " &_
                "        cvepfm01, " &_
                "        cvecli01, " &_
                "        DATE_FORMAT(fech31, '%d-%m-%Y') AS 'fech31',   " &_
                "        d.cgas31 as cg, "&_
                "        incble01,       "&_
                "        otros01,        "&_
                "        adusec01,       "&_
                "        patent01,       "&_
                "        FACTMO01        "&_
                " FROM "&tabla&" LEFT JOIN d31refer AS d on d.refe31=refcia01 left join  e31cgast as e  on e.cgas31=d.cgas31  " & _
                " where firmae01<>''  and fecpag01 >= '"&FormatoFechaInv(ISTRFINI)&"' and fecpag01 <= '"&FormatoFechaInv(FSTRFFIN)&"'  "& permi & _
                " and refcia01<>'DAI09-0947' " & _
                " group by refcia01 "
				
                'and  cveped01<>'R1'
                '" and ( cvecli01="&CCLIENTE&" )"&_

' Response.Write(sqlReferencias)
' Response.End

     MM_STRING = ODBC_POR_ADUANA(Session("GAduana"))
     set RSReferencias = server.CreateObject("ADODB.Recordset")
     RSReferencias.ActiveConnection = MM_STRING
     RSReferencias.Source= sqlReferencias
     RSReferencias.CursorType = 0
     RSReferencias.CursorLocation = 2
     RSReferencias.LockType = 1
     RSReferencias.Open()




%>

 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Layout Honeywell Operaciones de <%=pTipOpe%> </p></font></strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p></p></font></strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Del <%=ISTRFINI%> al <%=FSTRFFIN%> </p></font></strong>

<table border="1">
  <!--
  <tr bgcolor="#CCFFCC">
      <th>Control</th>
      <th>Empresa</th>
      <th>Cancelacion</th>
      <th>No.CGast</th>
      <th>FechCGtos</th>
      <th>TipOpe</th>
      <th>Nombre Agente Aduanal</th>
      <th>Patente</th>
      <th>Pedimento</th>
      <th>Guia Master</th>
      <th>Guia House</th>
      <th>OrdComp</th>
      <th>SBU</th>
      <th>Proveedor</th>
      <th>PaisOrigMat</th>
      <th>PuertoOrigMat</th>
      <th>FechEmbProv</th>
      <th>Fecha Pago de Impuestos </th>
      <th>NoFactProv</th>
      <th>FechFactProv</th>
      <th>ValDLLS</th>
      <th>TipCamb</th>
      <th>ValComer</th>
      <th>FletesYSegs</th>
      <th bgcolor="#FFCC33">GtosOrig(USD)</th>
      <th>Otros Incrementables</th>
      <th>ValAduana</th>
      <th>IGI</th>
      <th>DTA</th>
      <th bgcolor="#FFCC33">IGIyDTA</th>
      <th>Prev</th>
      <th>Cuota Compensatoria</th>
      <th>Semaforo</th>
      <th bgcolor="#FFFF99">TotDerechos</th>
      <th>IVAPagADuana</th>
      <th>TotPedimento</th>
      <th>Desconsolidado</th>
      <th>Cruces</th>
      <th>Alma's y Mani's</th>
      <th>FleteInter</th>
      <th>FleteLocal</th>
      <th>IVAFleteLocal</th>
      <th>Retención Flete</th>
      <th>Comprobados</th>
      <th>Complementarios</th>
      <th>Honorarios</th>
      <th bgcolor="#FFFF99">Gtos.Mexico</th>
      <th bgcolor="#FFCC33">GtosMexico(USD)</th>
      <th>IVACtaGtos</th>
      <th>Anticipo</th>
      <th>Saldo</th>
      <th>NoCliente</th>
      <th>Peso</th>
      <th>CuentaContableProyectoTarea</th>
      <th>#PartidasFacturas</th>
      <th>FletesXNtra.Cuenta</th>
      <th>Moneda</th>
      <th>FechEntrAduana</th>
      <th>FirmaElect</th>
      <th>CveDoc</th>
      <th>Regimen</th>
      <th>AduDEntrada</th>
      <th>AduDespacho</th>
      <th>RFC.Cliente</th>
  </tr>




      <th>TipOpe</th>
      <th>PaisOrigMat</th>
      <th>FechFactProv</th>
      <th>NoCliente</th>
      <th>FletesXNtra.Cuenta</th>

  -->

  <tr bgcolor="#CCFFCC">
      <th>CONTROL</th>
      <th>EMPRESA</th>
      <th>RFC.Cliente</th>
      <th>Cancelacion</th>
      <th>No.CGast</th>
      <th>FechCGtos</th>
      <th>Patente</th>
      <th>Pedimento</th>
      <th>Nombre Agente Aduanal</th>
      <th>ORDEN DE COMPRA</th>
      <th>SBU</th>
      <th>Guia Master</th>
      <th>Guia House</th>
      <th>PESO</th>
      <th>PuertoOrigMat</th>
      <th>FechEmbProv</th>
      <th>PROVEEDOR</th>
      <th>NO. FACTURA PROVEEDOR</th>
      <th>VALOR DLLS</th>
      <th>TIPO DE CAMBIO</th>
      <th>Fecha Pago</th>
      <th>VALOR COMERCIAL</th>
      <th> FLETES Y SEGUROS </th>
      <th bgcolor="#FFCC33">GASTOS EN ORIGEN (USD)</th>
      <th>OTROS INCREMENTABLES</th>
      <th>VALOR ADUANA</th>
      <th>IGI</th>
      <th>DTA</th>
      <th>Cuota Compensatoria</th>
      <th bgcolor="#FFCC33">IGI Y DTA (USD)</th>
      <th>Prev</th>
      <th bgcolor="#FFFF99">TotDerechos</th>
      <th>IVAPagADuana</th>
      <th>TotPedimento</th>

      <th>Cruces </th>
      <th>Iva Cruce	</th>
      <th>Retención Cruce </th>

      <th>Cta americana</th>
      <th>DESCONSOLIDACION</th>
      <th>ALMACENAJE Y MANIOBRAS</th>
      <th>FLETE INTERNACIONAL</th>
      <th>Moneda</th>
      <th>FLETE LOCAL</th>
      <th>IVA FLETE LOCAL</th>
      <th>RET. 4% FLETE LOCAL</th>
      <th>COMPROBADOS</th>
      <th>COMPLEMENTARIOS</th>
      <th>HONORARIOS</th>
      <th bgcolor="#FFFF99">GASTOS EN MÉXICO</th>
      <th bgcolor="#FFCC33"> GASTOS EN MÉXICO (USD) </th>
      <th>IVA CTA. GASTOS</th>
      <th>ANTICIPO</th>
      <th>SALDO</th>
      <th>CuentaContableProyectoTarea</th>
      <th>#PartidasFacturas</th>
      <th>FechEntrAduana</th>
      <th>FirmaElect</th>
      <th>Clave documento</th>
      <th>Regimen</th>
      <th>Aduana de entrada</th>
      <th>Aduana despacho</th>
      <th>Semaforo</th>
  </tr>


<%


if not RSReferencias.eof then

while not RSReferencias.eof

     claveclientePed = RSReferencias.fields.item("cvecli01").value

     '/////////////////////datos del cliente rfccli18,cvecli18,
     sqlDatosCliente=" SELECT cvecli18,nomcli18,rfccli18,domcli18,ciucli18 "&_
                     " FROM ssclie18 "&_
                     " where cvecli18="& claveclientePed & " "

    'Response.Write(sqlDatosCliente)
    'Response.End

     set RSDatosCliente = server.CreateObject("ADODB.Recordset")
     RSDatosCliente.ActiveConnection = MM_STRING
     RSDatosCliente.Source= sqlDatosCliente
     RSDatosCliente.CursorType = 0
     RSDatosCliente.CursorLocation = 2
     RSDatosCliente.LockType = 1
     RSDatosCliente.Open()
          if not RSDatosCliente.eof then
                'pTotalFleteLocal=RSFleLocalIVARet4.fields.item("TOTALFLETELOCAL").value
                pnombre=RSDatosCliente.fields.item("nomcli18").value
                prfc=RSDatosCliente.fields.item("rfccli18").value
                pcd =RSDatosCliente.fields.item("ciucli18").value
          else
                'pTotalFleteLocal=0
                pnombre=" - - - "
                prfc=" - - - "
                pcd=" - - - "
          end if
     RSDatosCliente.close
    set RSDatosCliente = nothing




'Response.End
' inicializo variables
ptipcam=0
pdta=0
pprev=0
padv_igi=0
sumaT=0

idx=idx +1
'/'/'/'/'////////// cuenta de gastos
         sqlctagtos=" Select d.refe31, DATE_FORMAT(fech31, '%d/%m/%Y') AS fech31, d.cgas31 as ctagts "&_
                    " FROM (e31cgast as e inner join d31refer as d on e.cgas31=d.cgas31) "&_
                    " where  d.refe31='"&RSReferencias.fields.item("refcia01").value&"' "&_
                    "       and esta31 = 'I' "

    'Response.Write(sqlctagtos)
    'Response.End
         set RSctagtos = server.CreateObject("ADODB.Recordset")
         RSctagtos.ActiveConnection = MM_STRING
         RSctagtos.Source= sqlctagtos
         RSctagtos.CursorType = 0
         RSctagtos.CursorLocation = 2
         RSctagtos.LockType = 1
         RSctagtos.Open()
          pcg      = 0
          pctagtos = ""
          fechCG   = ""

         while not RSctagtos.eof

           if pcg=0 and RSctagtos.fields.item("ctagts").value<>"" then
              'pctagtos = RSctagtos.fields.item("ctagts").value&" de "&RSctagtos.fields.item("fech31").value
              pctagtos = RSctagtos.fields.item("ctagts").value
              fechCG   = RSctagtos.fields.item("fech31").value
              pcg=321
           else
              pctagtos = pctagtos & " , "&RSctagtos.fields.item("ctagts").value
              fechCG   = fechCG & " , " & RSctagtos.fields.item("fech31").value
              'pctagtos=pctagtos&" , "&RSctagtos.fields.item("ctagts").value&" de "&RSctagtos.fields.item("fech31").value
           end if

           RSctagtos.movenext()
         wend
          RSctagtos.close
          set RSctagtos = nothing




'/'/'/'/'////////// Ver si hay alguna cuenta de gastos cancelada
         sqlctagtosCanc=" Select d.refe31, DATE_FORMAT(fech31, '%d/%m/%Y') AS fech31, d.cgas31 as ctagts "&_
                    " FROM (e31cgast as e inner join d31refer as d on e.cgas31=d.cgas31) "&_
                    " where  d.refe31='"&RSReferencias.fields.item("refcia01").value&"' "&_
                    "       and esta31 = 'C' "

    'Response.Write(sqlctagtos)
    'Response.End
         set RSctagtosCanc = server.CreateObject("ADODB.Recordset")
         RSctagtosCanc.ActiveConnection = MM_STRING
         RSctagtosCanc.Source= sqlctagtosCanc
         RSctagtosCanc.CursorType = 0
         RSctagtosCanc.CursorLocation = 2
         RSctagtosCanc.LockType = 1
         RSctagtosCanc.Open()
          pcgCanc      = 0
          pctagtosCanc = ""
          fechCGCanc   = ""

         while not RSctagtosCanc.eof

           if pcgCanc=0 and RSctagtosCanc.fields.item("ctagts").value<>"" then
              'pctagtos = RSctagtos.fields.item("ctagts").value&" de "&RSctagtos.fields.item("fech31").value
              pctagtosCanc = RSctagtosCanc.fields.item("ctagts").value
              fechCGCanc   = RSctagtosCanc.fields.item("fech31").value
              pcgCanc=321
           else
              pctagtosCanc = pctagtosCanc & " ; "&RSctagtosCanc.fields.item("ctagts").value
              fechCGCanc   = fechCGCanc & " ; " & RSctagtosCanc.fields.item("fech31").value
              'pctagtos=pctagtos&" , "&RSctagtos.fields.item("ctagts").value&" de "&RSctagtos.fields.item("fech31").value
           end if

           RSctagtosCanc.movenext()
         wend
          RSctagtosCanc.close
          set RSctagtosCanc = nothing


                  '*********************************************************

                         strGuiaMaster      = ""
                         strGuiaMasterHouse = ""
                         if RSReferencias.fields.item("refcia01").value <> "" then
                             Set Recguia = Server.CreateObject("ADODB.Recordset")
                             Recguia.ActiveConnection = MM_STRING
                             'strSqlSel =  "select numgui04 from ssguia04 where refcia04='" & ltrim(StrRefer)&"'"
                             strSqlSel =  " SELECT  IF( IDNGUI04=1,numgui04,'') AS guiaMaster,  " & _
                                          "         IF( IDNGUI04=2,numgui04,'') AS guiaHouse    " & _
                                          " from ssguia04  " & _
                                          " where refcia04='" & ltrim(RSReferencias.fields.item("refcia01").value)&"'"
                             'Response.Write(strSqlSel)
                             'Response.End

                             Recguia.Source = strSqlSel
                             Recguia.CursorType = 0
                             Recguia.CursorLocation = 2
                             Recguia.LockType = 1
                             Recguia.Open()
                             if not Recguia.eof then
                                 strGuiaMaster      = Recguia.Fields.Item("guiaMaster").Value
                                 strGuiaMasterHouse = Recguia.Fields.Item("guiaHouse").Value
                                 intcountguia1=1
                                 intcountguia2=1
                                 While NOT Recguia.EOF
                                    if Recguia.Fields.Item("guiaMaster").Value <> "" then
                                       if intcountguia1 = 1 then
                                           strGuiaMaster      = Recguia.Fields.Item("guiaMaster").Value
                                       else
                                           strGuiaMaster      = strGuiaMaster & "; "& Recguia.Fields.Item("guiaMaster").Value
                                       end if
                                       intcountguia1= intcountguia1 + 1
                                    end if

                                    if Recguia.Fields.Item("guiaHouse").Value <> "" then
                                       if intcountguia2 = 1 then
                                           strGuiaMasterHouse = Recguia.Fields.Item("guiaHouse").Value
                                       else
                                           strGuiaMasterHouse = strGuiaMasterHouse & "; "& Recguia.Fields.Item("guiaHouse").Value
                                       end if
                                       intcountguia2= intcountguia2 + 1
                                    end if

                                 Recguia.movenext

                                 Wend
                             end if
                             Recguia.close
                             set Recguia = Nothing
                         end if
                     '*********************************************************



'////////// Proveedor en ssprov22
         sqlprov="select cvepro22,nompro22 from ssprov22 where cvepro22='"&RSReferencias.fields.item("cvepro01").value&"' "

    'Response.Write(sqlordencomp)
    'Response.End
         set RSProv = server.CreateObject("ADODB.Recordset")
         RSProv.ActiveConnection = MM_STRING
         RSProv.Source= sqlprov
         RSProv.CursorType = 0
         RSProv.CursorLocation = 2
         RSProv.LockType = 1
         RSProv.Open()

          if not RSProv.eof then
            if RSProv.fields.item("nompro22").value<>"" then
                pnomprov=RSProv.fields.item("nompro22").value
            else
                pnomprov=" - - - "
            end if
          else
            pnomprov=" - - - "
          end if

        RSProv.close
        set RSProv = nothing


  'Response.End

'////////// orden de compra en d05artic

         sqlordencomp="select refe05,pedi05 from d05artic where refe05='"&RSReferencias.fields.item("refcia01").value&"' group by pedi05"

    'Response.Write(sqlordencomp)
    'Response.End
         set RSordencomp = server.CreateObject("ADODB.Recordset")
         RSordencomp.ActiveConnection = MM_STRING
         RSordencomp.Source= sqlordencomp
         RSordencomp.CursorType = 0
         RSordencomp.CursorLocation = 2
         RSordencomp.LockType = 1
         RSordencomp.Open()
          poc=0
          pordenc=""
         while not RSordencomp.eof

          if RSordencomp.fields.item("pedi05").value<>"" then
               if poc=0 then
                  pordenc=RSordencomp.fields.item("pedi05").value
                  poc=321
               else
                  pordenc=pordenc&" , "&RSordencomp.fields.item("pedi05").value
               end if
          else
              pordenc=" - - - "
          end if
           RSordencomp.movenext()
         wend
          RSordencomp.close
          set RSordencomp = nothing


         'Response.End

         sqlNumPartFact="select refe05,fact05,max(pfac05) as NumPartFact from d05artic where refe05='"&RSReferencias.fields.item("refcia01").value&"' group by fact05"
         'Response.Write(sqlNumPartFact)
         'Response.End

         set RSNumPartFact = server.CreateObject("ADODB.Recordset")
         RSNumPartFact.ActiveConnection = MM_STRING
         RSNumPartFact.Source= sqlNumPartFact
         RSNumPartFact.CursorType = 0
         RSNumPartFact.CursorLocation = 2
         RSNumPartFact.LockType = 1
         RSNumPartFact.Open()
          pnpf=0
          pNumPartFact=""

         if not RSNumPartFact.eof then

          if RSNumPartFact.fields.item("NumPartFact").value<>"" and RSNumPartFact.fields.item("fact05").value<>"" then
            while not RSNumPartFact.eof
               if pnpf=0 then
                  pNumPartFact=RSNumPartFact.fields.item("NumPartFact").value&" de "&RSNumPartFact.fields.item("fact05").value
                  pnpf=321
               else
                  pNumPartFact=pNumPartFact&" , "&RSNumPartFact.fields.item("NumPartFact").value&" de "&RSNumPartFact.fields.item("fact05").value
               end if
             RSNumPartFact.movenext()
            wend

          else
              pNumPartFact=" - - - "
          end if
         else
           pNumPartFact=" - - - "
         end if

          RSNumPartFact.close
          set RSNumPartFact = nothing

         'Response.End

'/////////honrarios y servicios

    sqlHonoYserv="select refe31, DATE_FORMAT(fech31, '%d/%m/%Y') as fcg,sum(sald31) as saldo, sum(anti31) as anticipo,"&_
                  "(sum(if (esta31='I',(chon31+caho31+csce31), 0 )) * (piva31/100)) as iva_cg,"&_
                  " (e31cgast.cgas31) as cg, DATE_FORMAT(fech31, '%d/%m/%Y') AS 'fech31', sum(if (esta31='I',(chon31+caho31), 0 ) ) as honorarios, "&_
                  " sum( if (esta31='I', (csce31), 0 )) as servicios, "&_
                  " sum(if(esta31='I', (chon31+caho31+csce31), 0)) as ingresos,"&_
                  " sum(if (esta31='I',(suph31), 0)) as o_cargos "&_
                  "from d31refer join e31cgast on d31refer.cgas31=e31cgast.cgas31 "&_
                  "where d31refer.refe31='"&RSReferencias.fields.item("refcia01").value&"' "&_
                  "group by refe31 "&_
                  "order by refe31"
'Response.Write(sqlHonoYserv)
'Response.End

     set RSHonoYserv = server.CreateObject("ADODB.Recordset")
     RSHonoYserv.ActiveConnection = MM_STRING
     RSHonoYserv.Source= sqlHonoYserv
     RSHonoYserv.CursorType = 0
     RSHonoYserv.CursorLocation = 2
     RSHonoYserv.LockType = 1
     RSHonoYserv.Open()
       if not RSHonoYserv.eof then
          if RSHonoYserv.fields.item("honorarios").value <>"" then
                phonorarios=RSHonoYserv.fields.item("honorarios").value
          else
                phonorarios=0
          end if
          if RSHonoYserv.fields.item("servicios").value <>"" then
                pservicios=RSHonoYserv.fields.item("servicios").value
          else
                pservicios=0
          end if
          if RSHonoYserv.fields.item("o_cargos").value <>"" then
                po_cargos=RSHonoYserv.fields.item("o_cargos").value
          else
                po_cargos=0
          end if
          if RSHonoYserv.fields.item("saldo").value <>"" then
                psaldo=RSHonoYserv.fields.item("saldo").value
          else
                psaldo=0
          end if
          if RSHonoYserv.fields.item("anticipo").value <>"" then
                panticipo=RSHonoYserv.fields.item("anticipo").value
          else
                panticipo=0
          end if
          if RSHonoYserv.fields.item("iva_cg").value <>"" then
                piva_cg=RSHonoYserv.fields.item("iva_cg").value
          else
                piva_cg=0
          end if
        else
            phonorarios=0
            po_cargos=0
            pservicios=0
            panticipo=0
            psaldo=0
            piva_cg=0
        end if
     RSHonoYserv.close
    set RSHonoYserv = nothing

    'Response.End


'////////////////////maniobras_flete
    sqlManiobras=" select  sum(if(d.esta21=0,d.mont21,0)) as maniobras_flete "&_
                  " from d21paghe as d join e21paghe as e on  (e.foli21=d.foli21 and e.fech21=d.fech21 and  e.conc21="&num_con_f_m&" )  "&_
                  " where d.refe21= '"&RSReferencias.fields.item("refcia01").value&"' "&_
                  " group by refe21 "
 'Response.Write(sqlManiobras)
 'Response.End

     set RSManiobras = server.CreateObject("ADODB.Recordset")
     RSManiobras.ActiveConnection = MM_STRING
     RSManiobras.Source= sqlManiobras
     RSManiobras.CursorType = 0
     RSManiobras.CursorLocation = 2
     RSManiobras.LockType = 1
     RSManiobras.Open()

          if not RSManiobras.eof then
                pmaniobras_flete=RSManiobras.fields.item("maniobras_flete").value
          else
                pmaniobras_flete=0
          end if
     RSManiobras.close
    set RSManiobras = nothing

    'Response.End

'////////////////////desconsolidado
    sqlDesconsolidado=" select  sum(if(d.esta21=0,d.mont21,0)) as desconsolidado "&_
                  " from d21paghe as d join e21paghe as e on  (e.foli21=d.foli21 and e.fech21=d.fech21 and  e.conc21="&c_desconsolidado&" )  "&_
                  " where d.refe21= '"&RSReferencias.fields.item("refcia01").value&"' "&_
                  " group by refe21 "
'Response.Write(sqlManiobras)
'Response.End
     set RSDesconsolidado = server.CreateObject("ADODB.Recordset")
     RSDesconsolidado.ActiveConnection = MM_STRING
     RSDesconsolidado.Source= sqlDesconsolidado
     RSDesconsolidado.CursorType = 0
     RSDesconsolidado.CursorLocation = 2
     RSDesconsolidado.LockType = 1
     RSDesconsolidado.Open()

          if not RSDesconsolidado.eof then
                pdesconsolidado=RSDesconsolidado.fields.item("desconsolidado").value
          else
                pdesconsolidado=0
          end if
     RSDesconsolidado.close
    set RSDesconsolidado = nothing

    'Response.End

    '////////////////////maniobras y almacenajes
    sqlManiYAlma=" select  sum(if(d.esta21=0,d.mont21,0)) as maniyalma "&_
                  " from d21paghe as d join e21paghe as e on  (e.foli21=d.foli21 and e.fech21=d.fech21 and  e.conc21="&c_maniobrasyalmacenajes&" )  "&_
                  " where d.refe21= '"&RSReferencias.fields.item("refcia01").value&"' "&_
                  " group by refe21 "
'Response.Write(sqlManiobras)
'Response.End
     set RSManiYAlma = server.CreateObject("ADODB.Recordset")
     RSManiYAlma.ActiveConnection = MM_STRING
     RSManiYAlma.Source= sqlManiYAlma
     RSManiYAlma.CursorType = 0
     RSManiYAlma.CursorLocation = 2
     RSManiYAlma.LockType = 1
     RSManiYAlma.Open()

          if not RSManiYAlma.eof then
                pmaniyalma=RSManiYAlma.fields.item("maniyalma").value
          else
                pmaniyalma=0
          end if
     RSManiYAlma.close
    set RSManiYAlma = nothing

    'Response.End

    '************************************************************************************************************************************************************************************************************************************************************************************

	'///// FLETE LOCAL, IVA DEL FLETE LOCAL Y RETENCION DEL 4% DEL FLETE LOCAL
			'SE ELIMINO ESTE QUERY POR QUE NO ESTABA RETORNANDO NADA, SE CAMBIO POR EL DE ALGUNOS RENGLONES ABAJO
			'sqlfleteslocalivaret4="select D.refe21,D.foli21,D.fech21,(D.mfle21) AS FLETELOCAL, ((D.mfle21)*.15) AS IVADFLETEL,((D.mfle21)*.04) AS RET4DFLETEL, E.foli21,E.fech21,E.conc21, E.esta21,E.tpag21,E.piva21, C.clav21,C.desc21 "&_
					'"from d21paghe as D,e21paghe as E, c21paghe as C "&_
					'"where C.clav21='"&c_fleteTerr&"' and C.clav21=E.conc21 and D.foli21=E.foli21 and D.fech21=E.fech21 "&_
					'"and D.refe21='"&RSReferencias.fields.item("refcia01").value&"' "		
					
			sqlfleteslocalivaret4="SELECT D.refe21,D.foli21,D.fech21,(D.mfle21) AS FLETELOCAL, ((D.mfle21)*.15) AS IVADFLETEL,((D.mfle21)*.04) AS RET4DFLETEL, E.foli21,E.fech21,E.conc21, E.esta21,E.tpag21,E.piva21, C.clav21,C.desc21 "&_
					"FROM d21paghe AS D "&_
					"LEFT JOIN e21paghe AS E ON E.foli21 = D.foli21 AND YEAR(E.fech21) = YEAR(D.fech21) AND E.esta21 <> 'S' AND E.esta21 <> 'C' AND E.tmov21 =D.tmov21 "&_
					"LEFT JOIN c21paghe AS C ON C.clav21 = E.conc21 "&_
					"WHERE C.clav21 = '7' AND refe21='"&RSReferencias.fields.item("refcia01").value&"' "
							
			'Response.Write(sqlfleteslocalivaret4)
			'Response.End
			
			set RSFleLocalIVARet4 = server.CreateObject("ADODB.Recordset")
			RSFleLocalIVARet4.ActiveConnection = MM_STRING
			RSFleLocalIVARet4.Source= sqlfleteslocalivaret4
			RSFleLocalIVARet4.CursorType = 0
			RSFleLocalIVARet4.CursorLocation = 2
			RSFleLocalIVARet4.LockType = 1
			RSFleLocalIVARet4.Open()

			if not RSFleLocalIVARet4.eof then
				'pTotalFleteLocal=RSFleLocalIVARet4.fields.item("TOTALFLETELOCAL").value
				pFleteLocal=RSFleLocalIVARet4.fields.item("FLETELOCAL").value
				pIVAFleteL=RSFleLocalIVARet4.fields.item("IVADFLETEL").value
				pRet4FleteL =RSFleLocalIVARet4.fields.item("RET4DFLETEL").value
			else
				'pTotalFleteLocal=0
				pFleteLocal=0
				pIVAFleteL=0
				pRet4FleteL=0
			end if

		RSFleLocalIVARet4.close
		set RSFleLocalIVARet4 = nothing
		'Response.End
		'******************************************************************
		'El siguiente codigo se comento por que arrojaba valores incorrectos, 0 en la mayoria de los casos
				'******************************************************************
		'	   pFleteLocal    = 0
		'	   pIVAFleteL     = 0
		'	   pRet4FleteL    = 0
		'	   montofleteAux2 = 0
		'	   montoSvrComp   = 0
		'
		'	   Set RsValor2 = Server.CreateObject("ADODB.Recordset")
		'	   RsValor2.ActiveConnection = MM_STRING
		'	   strSQLsvr = " SELECT refe32,ttar32, dcrp32,mont32" &_
		'				   " FROM d32rserv  " &_
		'				   " WHERE REFE32 = '" &RSReferencias.fields.item("refcia01").value& "'"
		'
		'	   Response.Write(strSQLsvr)
		'	   Response.End
		'
		'	   RsValor2.Source = strSQLsvr
		'	   RsValor2.CursorType = 0
		'	   RsValor2.CursorLocation = 2
		'	   RsValor2.LockType = 1
		'	   RsValor2.Open()
		'	   if not RsValor2.EOF then
		'		   While NOT RsValor2.EOF
		'			 if (RsValor2.Fields.Item("ttar32").Value = "00033") then
		'				montofleteAux2  = montofleteAux2 + RsValor2.Fields.Item("mont32").Value
		'			 else
		'				montoSvrComp  = montoSvrComp + RsValor2.Fields.Item("mont32").Value
		'			 end if
		'			 RsValor2.movenext
		'		   Wend
		'	   end if
		'	   RsValor2.close
		'	   set  RsValor2 = nothing
		'
		'	   pFleteLocal = montofleteAux2
		'	   pIVAFleteL  = montofleteAux2*0.15
		'	   pRet4FleteL = 0
		'	   pservicios  = montoSvrComp

		'******************************************************************

		'************************************************************************************************************************************************************************************************************************************************************************************

'///////////// igi(advaloren)

  if base= "LAR" then
      slqimpuestos= " select '' as feorig01 ," &_
                    "        '' as cveptoemb," &_
                    "        '' as cvepai01," &_
                    "        '' as nompto01," &_
                    "        '' as adudes01, " &_
                    "              import36, " &_
                    "       (i_adv102 + i_adv202) as adv_igi,  " &_
                    "       SUM(vaduan02) AS vadu," &_
                    "       '' AS vprepag, "&_
                    "       sum(i_iva102+i_iva202) as iva, " &_
                    "       paiori02, " &_
                    "       sum(vmerme02) as ValComer , " &_
                    "       rcli01, " &_
                    "       alea01, " &_
                    "       PTOEMB01, " &_
                    "       FLEINT01 " &_
                " from c01refer,sscont36,ssfrac02 "&_
                " where refe01='"&RSReferencias.fields.item("refcia01").value&"' and "&_
                "       refe01=refcia36 and "&_
                "       refe01=refcia02 and "&_
                "       cveimp36=15 "&_
                "group by refcia02 "

                '"        '' as cveimp36, " &_
  else
      slqimpuestos=" select DATE_FORMAT(feorig01, '%d/%m/%Y') AS 'feorig01', " & _
                   "        cveptoemb," & _
                   "        '' as cvepai01," & _
                   "        '' as nompto01," & _
                   "        adudes01," & _
                   "        import36," & _
                   "        (i_adv102 + i_adv202) as adv_igi," & _
                   "        SUM(vaduan02) AS vadu, " & _
                   "        SUM(prepag02) AS vprepag, "&_
                   "        sum(i_iva102+i_iva202) as iva, paiori02, sum(vmerme02) as ValComer ,"&_
                   "        rcli01, " &_
                   "        alea01, " &_
                   "        PTOEMB01, " &_
                   "        FLEINT01 " &_
                   " from c01refer,sscont36,ssfrac02 "&_
                   " where refe01='"&RSReferencias.fields.item("refcia01").value&"' and "&_
                   "    refe01=refcia36 and "&_
                   "    refe01=refcia02 and "&_
                   "    cveimp36=15 "&_
                   " group by refcia02 "

                   '"        cveimp36," & _
                   '"    cveptoemb=cvepto01 and "&_

  end if

'Response.Write(slqimpuestos)
'Response.end


     usbcliente = ""
     semaforo   = 0
     strptoemb = ""
     strfleteint01 = 0
     set RSimpuestos = server.CreateObject("ADODB.Recordset")
     RSimpuestos.ActiveConnection = MM_STRING
     RSimpuestos.Source= slqimpuestos
     RSimpuestos.CursorType = 0
     RSimpuestos.CursorLocation = 2
     RSimpuestos.LockType = 1
     RSimpuestos.Open()

          if not RSimpuestos.eof then
                pfechaemb      = RSimpuestos.fields.item("feorig01").value
                padudes        = RSimpuestos.fields.item("adudes01").value
                pPyCOrig       = RSimpuestos.fields.item("cvepai01").value&", "&RSimpuestos.fields.item("PTOEMB01").value
                padv_igi       = RSimpuestos.fields.item("adv_igi").value
                pprev          = RSimpuestos.fields.item("import36").value
                padu           = RSimpuestos.fields.item("vadu").value
                piva           = RSimpuestos.fields.item("iva").value
                pporigen       = RSimpuestos.fields.item("paiori02").value
                pvprepag       = RSimpuestos.fields.item("vprepag").value
                pvalcomer      = RSimpuestos.fields.item("ValComer").value
                usbcliente     = RSimpuestos.fields.item("rcli01").value
                semaforo       = RSimpuestos.fields.item("alea01").value
                strptoemb      = RSimpuestos.fields.item("PTOEMB01").value
                strfleteint01  = RSimpuestos.fields.item("FLEINT01").value
                strcveptoEmb01 = RSimpuestos.fields.item("cveptoemb").value

          else
                pfechaemb      = "- - -"
                pPyCOrig       = "- - -"
                padv_igi       = 0
                pprev          = 0
                padu           = 0
                piva           = 0
                pporigen       = " - - -"
                pvprepag       = 0
                pvalcomer      = 0
                usbcliente     = ""
                semaforo       = 0
                strfleteint01  = 0
                strcveptoEmb01 = 0

          end if
            'pfdsp=RSimpuestos.fields.item("fdsp01").value

     RSimpuestos.close
    set RSimpuestos = nothing
' igi(advaloren)
'Response.End



            '**************************************************************************
            ' Puerto de embarque y Pais de Embarque
            '**************************************************************************
                           strNomcveptoEmb01  = ""
                           strNomPaisptoEmb01 = ""
                           sqlNomPaisEmb = " SELECT CVEPTO01,CVEPAI01,NOMPTO01 " & _
                                          " FROM C01PTOEMB                    " & _
                                          " WHERE  CVEPTO01 = "& strcveptoEmb01

                           set RSNomPaisEmb = server.CreateObject("ADODB.Recordset")
                           RSNomPaisEmb.ActiveConnection = MM_STRING
                           RSNomPaisEmb.Source = sqlNomPaisEmb
                           RSNomPaisEmb.CursorType = 0
                           RSNomPaisEmb.CursorLocation = 2
                           RSNomPaisEmb.LockType = 1
                           RSNomPaisEmb.Open()

                           'Response.Write(sqlNomPaisEmb)
                           'Response.End

                           if not RSNomPaisEmb.eof then
                               'strNomcveptoEmb01  = RSNomPaisEmb.fields.item("NOMPTO01").value &","&RSNomPaisEmb.fields.item("CVEPAI01").value
                               strNomcveptoEmb01  = RSNomPaisEmb.fields.item("NOMPTO01").value
                               strNomPaisptoEmb01 = RSNomPaisEmb.fields.item("CVEPAI01").value
                            else
                                strNomcveptoEmb01  = ""
                                strNomPaisptoEmb01 = ""
                            end if
                           RSNomPaisEmb.close
                           set RSNomPaisEmb = nothing
            '**************************************************************************

            '**************************************************************************
            ' Nombre del pais del catalogo de paises
            '**************************************************************************
                           'strCatNomPaisptoEmb01 = ""
                           sqlCatPaisEmb = " SELECT NOMPAI19 " & _
                                           " FROM sspais19 " & _
                                           "  where cvepai19 = '"&strNomPaisptoEmb01&"'"

                           set RSCatPaisEmb = server.CreateObject("ADODB.Recordset")
                           RSCatPaisEmb.ActiveConnection = MM_STRING
                           RSCatPaisEmb.Source = sqlCatPaisEmb
                           RSCatPaisEmb.CursorType = 0
                           RSCatPaisEmb.CursorLocation = 2
                           RSCatPaisEmb.LockType = 1
                           RSCatPaisEmb.Open()

                           'Response.Write(sqlNomPaisEmb)
                           'Response.End

                           if not RSCatPaisEmb.eof then
                               strNomcveptoEmb01 = strNomcveptoEmb01&","&RSCatPaisEmb.fields.item("NOMPAI19").value
                            end if
                           RSCatPaisEmb.close
                           set RSCatPaisEmb = nothing
            '**************************************************************************


            '**************************************************************************
            ' traer impuesto - Cuota compensatoria
            '**************************************************************************
                           StrImpContrCC  = ""
                           IntCountImpContr   = 0
                           ' facturas comerciales
                           sqlImpContr = " SELECT SUM(import36) as Impuestos,               " & _
                                         "        SUM(IF(cveimp36=1,import36,0) )  AS DTA,  " & _
                                         "        SUM(IF(cveimp36=2,import36,0) )  AS CC,   " & _
                                         "        SUM(IF(cveimp36=3,import36,0) )  AS IVA,  " & _
                                         "        SUM(IF(cveimp36=4,import36,0) )  AS ISAN, " & _
                                         "        SUM(IF(cveimp36=5,import36,0) )  AS IEPS, " & _
                                         "        SUM(IF(cveimp36=6,import36,0) )  AS ADV,  " & _
                                         "        SUM(IF(cveimp36=15,import36,0) ) AS PRV,  " & _
                                         "        SUM(import36) AS TOTAL                    " & _
                                         " FROM sscont36                                    " & _
                                         " WHERE  FPAGOI36 = 0                              " & _
                                         "        AND  REFCIA36 = '"&ltrim(RSReferencias.fields.item("refcia01").value)&"' " & _
                                         " GROUP BY refcia36                               "

                           'strSqlSel =  " SELECT SUM(import36) as Impuestos " & _
                           '               " FROM sscont36         " & _
                           '               " WHERE  REFCIA36 = '"&ltrim( RSReferencias.fields.item("refcia01").value )&"'  AND " & _
                           '               "        FPAGOI36 = 0 " & _
                           '               " GROUP BY refcia36 "

                           'Response.Write(sqlImpContr)
                           'Response.End

                           set RSImpContr = server.CreateObject("ADODB.Recordset")
                           RSImpContr.ActiveConnection = MM_STRING
                           RSImpContr.Source= sqlImpContr
                           RSImpContr.CursorType = 0
                           RSImpContr.CursorLocation = 2
                           RSImpContr.LockType = 1
                           RSImpContr.Open()
                           'pnpf=0
                           'pNumPartFact=""

                           if not RSImpContr.eof then
                             'if RSNumPartFact.fields.item("NumPartFact").value<>"" and RSNumPartFact.fields.item("fact05").value<>"" then
                              'while not RSImpContr.eof
                              '   if IntCountImpContr=0 then
                                    'pNumPartFact   = RSNumPartFact.fields.item("MONFAC39").value
                                    StrImpContrCC    = RSImpContr.fields.item("CC").value
                                    StrImpContrTotal = RSImpContr.fields.item("TOTAL").value
                                    padv_igi         = RSImpContr.fields.item("ADV").value '''****************
                                    strdta           = RSImpContr.fields.item("DTA").value
                                    pdta             = RSImpContr.fields.item("DTA").value
                                    piva             = RSImpContr.fields.item("IVA").value

                                    'FecNumPartFact = RSImpContr.fields.item("FECFAC39").value
                                    'MonNumPartFact = RSImpContr.fields.item("MONFAC39").value

                                    '&" de "&RSNumPartFact.fields.item("fact05").value
                                    'IntCountImpContr=321
                              '   else
                              '      'pNumPartFact=pNumPartFact&" , "&RSNumPartFact.fields.item("NumPartFact").value&" de "&RSNumPartFact.fields.item("fact05").value
                              '      NumPartFact    = NumPartFact & ";"& RSImpContr.fields.item("NUMFAC39").value
                              '      FecNumPartFact = FecNumPartFact & ";"& RSImpContr.fields.item("FECFAC39").value
                              '      MonNumPartFact = MonNumPartFact & ";"& RSImpContr.fields.item("MONFAC39").value
                              '      IntCountImpContr= IntCountImpContr + 1
                              '   end if
                              ' RSImpContr.movenext()
                              'wend
                            else
                                pNumPartFact=""
                            end if
                             'else
                             '  pNumPartFact=" - - - "
                             'end if
                           RSImpContr.close
                           set RSImpContr = nothing
            '**************************************************************************



               NumPartFact    = ""
               FecNumPartFact = ""
               MonNumPartFact = ""
               Valfact39      = 0
               pnpf           = 0

               ' facturas comerciales
               sqlNumPartFact = " SELECT NUMFAC39,VALDLS39 , FECFAC39, MONFAC39 " & _
                                " FROM SSFACT39   " & _
                                " WHERE REFCIA39 = '"& RSReferencias.fields.item("refcia01").value &"' "
               'Response.Write(sqlNumPartFact)
               'Response.End

               set RSNumPartFact = server.CreateObject("ADODB.Recordset")
               RSNumPartFact.ActiveConnection = MM_STRING
               RSNumPartFact.Source= sqlNumPartFact
               RSNumPartFact.CursorType = 0
               RSNumPartFact.CursorLocation = 2
               RSNumPartFact.LockType = 1
               RSNumPartFact.Open()
               'pnpf=0
               'pNumPartFact=""

               if not RSNumPartFact.eof then
                 'if RSNumPartFact.fields.item("NumPartFact").value<>"" and RSNumPartFact.fields.item("fact05").value<>"" then
                  while not RSNumPartFact.eof
                     if pnpf=0 then
                        'pNumPartFact   = RSNumPartFact.fields.item("MONFAC39").value
                        NumPartFact     = RSNumPartFact.fields.item("NUMFAC39").value
                        'FecNumPartFact = RSNumPartFact.fields.item("FECFAC39").value
                        MonNumPartFact  = RSNumPartFact.fields.item("MONFAC39").value
                        Valfact39       = RSNumPartFact.fields.item("VALDLS39").value

                        '&" de "&RSNumPartFact.fields.item("fact05").value
                        'pnpf=321
                     else
                        'pNumPartFact=pNumPartFact&" , "&RSNumPartFact.fields.item("NumPartFact").value&" de "&RSNumPartFact.fields.item("fact05").value
                        NumPartFact    = NumPartFact & ";"& RSNumPartFact.fields.item("NUMFAC39").value
                        FecNumPartFact = FecNumPartFact & ";"& RSNumPartFact.fields.item("FECFAC39").value
                        MonNumPartFact = MonNumPartFact & ";"& RSNumPartFact.fields.item("MONFAC39").value
                        Valfact39      = Valfact39 + RSNumPartFact.fields.item("VALDLS39").value
                        pnpf= pnpf + 1
                     end if
                   RSNumPartFact.movenext()
                  wend
                else
                    pNumPartFact=""
                end if
                 'else
                 '  pNumPartFact=" - - - "
                 'end if
               RSNumPartFact.close
               set RSNumPartFact = nothing




ptipcam=RSReferencias.fields.item("tipcam01").value
'pdta=RSReferencias.fields.item("dta").value
''pprev=RSReferencias.fields.item("prev").value
'padv_igi=RSReferencias.fields.item("adv_igi").value

'sumaT=pdta+pprev+padv_igi+phonorarios+pservicios+po_cargos+pmaniobras_flete

''pdta=RSReferencias.fields.item("dta").value

'Response.End


%>







<!--
  <tr bgcolor="#CCFFCC">
      <th>Control</th>
      <th>Empresa</th>
      <th>RFC.Cliente</th>
      <th>Cancelacion</th>
      <th>No.CGast</th>
      <th>FechCGtos</th>
      <th>Patente</th>
      <th>Pedimento</th>
      <th>Nombre Agente Aduanal</th>
      <th>OrdComp</th>
      <th>SBU</th>
      <th>Guia Master</th>
      <th>Guia House</th>
      <th>Peso</th>
      <th>PuertoOrigMat</th>
      <th>FechEmbProv</th>
      <th>Proveedor</th>
      <th>NoFactProv</th>
      <th>ValDLLS</th>
      <th>TipCamb</th>
      <th>ValComer</th>
      <th>FletesYSegs</th>
      <th bgcolor="#FFCC33">GtosOrig(USD)</th>
      <th>Otros Incrementables</th>
      <th>ValAduana</th>
      <th>IGI</th>
      <th>DTA</th>
      <th>Cuota Compensatoria</th>
      <th bgcolor="#FFCC33">IGIyDTA</th>
      <th>Prev</th>
      <th bgcolor="#FFFF99">TotDerechos</th>
      <th>IVAPagADuana</th>
      <th>TotPedimento</th>
      <th>Cruces</th>
      <th>Cta americana</th>
      <th>Desconsolidado</th>
      <th>Alma's y Mani's</th>
      <th>FleteInter</th>
      <th>Moneda</th>
      <th>FleteLocal</th>
      <th>IVAFleteLocal</th>
      <th>Retención Flete</th>
      <th>Comprobados</th>
      <th>Complementarios</th>
      <th>Honorarios</th>
      <th bgcolor="#FFFF99">Gtos.Mexico</th>
      <th bgcolor="#FFCC33">GtosMexico(USD)</th>
      <th>IVACtaGtos</th>
      <th>Anticipo</th>
      <th>Saldo</th>
      <th>CuentaContableProyectoTarea</th>
      <th>#PartidasFacturas</th>
      <th>FechEntrAduana</th>
      <th>FirmaElect</th>
      <th>CveDoc</th>
      <th>Regimen</th>
      <th>AduDEntrada</th>
      <th>AduDespacho</th>
      <th>Semaforo</th>
  </tr>


  <th><font size="1" face="Arial"><%=pTipOpe%></th> <!-- Tipo de operacion>

  <th><font size="1" face="Arial"><%=FecNumPartFact%></th> <!-- Fecha de Factura >
  <th><font size="1" face="Arial"><%=RSReferencias.fields.item("cvecli01").value%></th>
  <th><font size="1" face="Arial"> </th> <!-- FleXNtraCuent .- Flete por nuestra cuenta >


-->

  <tr>
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("refcia01").value%></th>
      <th><font size="1" face="Arial"><%=pnombre%></th>
      <th><font size="1" face="Arial"><%=prfc%></th>
      <th><font size="1" face="Arial"><%=pctagtosCanc%></th>
      <th><font size="1" face="Arial"><%=pctagtos%></th>
      <th><font size="1" face="Arial"><%=fechCG%> </th>         <!-- Fecha de Cuenta de gastos, ya se toma en cuenta en la cuenta de gastos-->


<%
strpatent01 = RSReferencias.fields.item("patent01").value
strNamepatent01 = ""
if  strpatent01 = "3210"  then
   strNamepatent01 = "LIC. ROLANDO REYES KURI"
else
  if  strpatent01 = "3921"  then
   strNamepatent01 = "LUIS ENRIQUE DE LA CRUZ REYES"
  else
     if  strpatent01 = "3931"  then
       strNamepatent01 = "SERGIO ALVAREZ RAMIREZ"
     else
        if  strpatent01 = "3857"  then
           strNamepatent01 = "RAFAEL MENDOZA DIAZ BARRIGA"
        else
          if  strpatent01 = "3407"  then
            strNamepatent01 = "YOLANDA LEYVA SALAZAR"
          else
            if  strpatent01 = "3933"  then
              strNamepatent01 = "MAURICIO MENDOZA SANTA ANA"
            end if
          end if
        end if
     end if

  end if

end if
'3210  LIC. ROLANDO REYES KURI
'3921  LUIS ENRIQUE DE LA CRUZ REYES
'3931  SERGIO ALVAREZ RAMIREZ
'3857  RAFAEL MENDOZA DIAZ BARRIGA
'3407  YOLANDA LEYVA SALAZAR
'3933  MAURICIO MENDOZA SANTA ANA
%>
      <th><font size="1" face="Arial"><%=strpatent01%></th> <!--Patente-->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("numped01").value%></th> <!-- Pedimento -->
      <th><font size="1" face="Arial"><%=strNamepatent01%></th><!--Nombre del Agente-->
      <th><font size="1" face="Arial"><%=pordenc%></th>
      <%
        'if len(usbcliente) = 3 then
      %>
      <th><font size="1" face="Arial"><%=usbcliente%> </th>
      <%
        'else
        'end if
      %>
      <th><font size="1" face="Arial"><%=strGuiaMaster%> </th> <!-- Guia Master -->
      <th><font size="1" face="Arial"><%=strGuiaMasterHouse%> </th> <!-- Guia House -->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("pesobr01").value%></th>
      <!-- <th><font size="1" face="Arial"><%=strptoemb%> , <%=pporigen%>  </th>  -->  <!-- Puerto de Origen --> <!--Pais Origen -->
      <th><font size="1" face="Arial"><%=strNomcveptoEmb01%></th>  <!-- Puerto de Origen --> <!--Pais Origen -->

      <th><font size="1" face="Arial"><%=pfechaemb%></th>
      <th><font size="1" face="Arial"><%=pnomprov%> </th> <!--Nombre del proveedor -->
      <th><font size="1" face="Arial"><%=NumPartFact%></th>

      <th><font size="1" face="Arial"><%=Replace(( Valfact39 ),",",".")%></th> <!-- Valor dolares-->  <!-- pvalcomer*RSReferencias.fields.item("FACTMO01").value -->



      <th><font size="1" face="Arial"><%=Replace(RSReferencias.fields.item("tipcam01").value,",",".")%></th> <!-- -Tipo de cambio-->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("fecpag01").value%></th> <!--Fecha de pago -->
      <th><font size="1" face="Arial"><%=Replace(pvalcomer*RSReferencias.fields.item("tipcam01").value*RSReferencias.fields.item("FACTMO01").value,",",".")%></th> <!-- VALOR COMERCIAL -->
      <th><font size="1" face="Arial"><%=Replace(RSReferencias.fields.item("fleyseg").value,",",".")%></th>
      <th><font size="1" face="Arial"></th> <!-- Gastos de origen USD -->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("incble01").value %></th><!--Otros Incrementables-->
      <th><font size="1" face="Arial"><%=Replace(padu,",",".")%></th> <!-- Valor aduana -->
      <th><font size="1" face="Arial"><%=Replace(padv_igi,",",".")%></th>
      <%
         strdta = ""
         'Response.Write(strdta)
         'Response.Write("<br>")
         'Response.Write(RSReferencias.fields.item("dta").value)
         'Response.End
         'if not isnull(RSReferencias.fields.item("dta").value) then
         if not RSReferencias.fields.item("dta").value = "" then
           strdta = Replace( CStr(RSReferencias.fields.item("dta").value),",",".")
         end if
         'Replace(isnull(RSReferencias.fields.item("dta").value,0),",",".")
      %>
      <th><font size="1" face="Arial"><%=strdta%></th>
      <th><font size="1" face="Arial"><%=StrImpContrCC%></th> <!-- Cuota Compensatoria -->

      
      <th><font size="1" face="Arial"><%=0%> </th> <!-- IGIyDTA  en dolares 02/03/09 solicitaron que quedara en ceros-->

      <th><font size="1" face="Arial"><%=pprev%></th>
      <th><font size="1" face="Arial"><%=cdbl(padv_igi)+cdbl(pdta)+cdbl(pprev)%></th> <!-- Total derechos  -->
      <th><font size="1" face="Arial"><%=piva%></th> <!-- IVAPagADuana -->

      <!-- <th><font size="1" face="Arial"><%=pvprepag %>  --></th> <!-- TotPedimento -->
      <th><font size="1" face="Arial"><%=StrImpContrTotal %></th> <!-- TotPedimento -->



      <th><font size="1" face="Arial">  </th>  <!-- Cruces -->
      <th><font size="1" face="Arial">  </th>  <!-- Iva Cruce -->
      <th><font size="1" face="Arial">  </th>  <!-- Retención Cruce -->

      <th><font size="1" face="Arial">  </th>  <!-- Cuenta americana -->
      <th><font size="1" face="Arial"><%=pdesconsolidado %></th>
      <th><font size="1" face="Arial"><%=pmaniyalma%></th>

      <!-- <th><font size="1" face="Arial">strfleteint01 </th>  <!-- Flete internacional el que se genera en aduana americana y gastos de cuenta americana-->
      <th><font size="1" face="Arial"> </th>  <!-- Flete internacional el que se genera en aduana americana y gastos de cuenta americana-->

      <th><font size="1" face="Arial"></th> <!-- MonNumPartFact-->
      <th><font size="1" face="Arial"><%=pFleteLocal%></th>

      <!-- <th><font size="1" face="Arial"><%=pIVAFleteL%></th> -->
      <th><font size="1" face="Arial"><%=0%></th> <!--02/03/09 solicitaron se quedara en ceros -->

      <th><font size="1" face="Arial"><%=pRet4FleteL%></th>
      <th><font size="1" face="Arial"> </th> <!-- Gastos comprobados -->
      <th><font size="1" face="Arial"><%=pservicios%></th>
      <th><font size="1" face="Arial"><%=Replace(phonorarios,",",".")%></th>
      <th><font size="1" face="Arial">  </th> <!-- Gastos en mexico -->
      <th><font size="1" face="Arial">  </th> <!-- Gastos USD-->
      <th><font size="1" face="Arial"><%=piva_cg%></th>
      <th><font size="1" face="Arial"><%=panticipo%></th>
      <th><font size="1" face="Arial"><%=psaldo%></th>
      <%
        'if len(usbcliente) = 4 then
      %>
      <th><font size="1" face="Arial">  </th> <!-- Cuenta proyecto contable -->
      <% 
	  'usbcliente
	  %>
      <%
        'else
        'end if
      %>
      <th><font size="1" face="Arial"><%=pNumPartFact%></th>
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("fechaEA").value%></th>
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("firmae01").value%></th>
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("cveped01").value%></th>

      <!-- <th><font size="1" face="Arial"><%=RSReferencias.fields.item("desdoc01").value%></th> --> <!-- Regimen -->
      <th><font size="1" face="Arial"><%=pTipOpe%></th> <!-- Regimen -->


      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("adusec01").value%></th>
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("adusec01").value%></th>
      <% 
        tmpsemaforo = ""
        if semaforo = 1 then
           tmpsemaforo = "ROJO"
        else
           if semaforo = 2 then
              tmpsemaforo = "VERDE"
           else
              tmpsemaforo = ""
           end if
        end if
      %>
      <th><font size="1" face="Arial"> <%=tmpsemaforo%> </th> <!-- Semaforo -->





      <!--  <th><font size="1" face="Arial">pprev</th><!--Prevalidación> -->
      <!-- Replace(((padv_igi+pdta)/ptipcam),",",".") -->
      <!-- Replace(pprev,",",".") -->
      <!-- Replace((padv_igi+pdta+pprev),",",".") -->
      <!-- Replace(piva,",",".") -->
      <!-- Replace(pvprepag,",",".") -->
      <!-- Replace(pdesconsolidado,",",".") -->
      <!-- Replace(pmaniyalma,",",".") -->
      <!-- Replace(pFleteLocal,",",".") -->
      <!-- Replace(pIVAFleteL,",",".") -->
      <!-- Replace(pRet4FleteL,",",".") -->
      <!-- Replace(pservicios,",",".") -->
      <!-- Replace(phonorarios,",",".") -->
      <!-- Replace(piva_cg,",",".") -->
      <!-- %=Replace(panticipo,",",".") -->
      <!-- Replace(psaldo,",",".") -->
      <!-- <th><font size="1" face="Arial"> <pcd> </th> -->
      <!-- Replace(pcd,",",".") -->
      <!-- Replace(RSReferencias.fields.item("pesobr01").value,",",".") -->
      <!-- <th><font size="1" face="Arial"> RSReferencias.fields.item("cvepfm01").value </th> -->
      <!-- <th><font size="1" face="Arial">padudes </th> -->
      <!--th> - </th-->
  </tr>





<% 

'Response.Write("aqui")
'Response.End

  RSReferencias.movenext()

wend

RSReferencias.close
set RSReferencias = nothing


else
%>
<tr>
  <th colspan=12>
    <font size="1" face="Arial">No se Encontro ningun registro con esos parametros
  </th>
</tr>
<% 
end if

%>
</table>

<% 



             'Carga en Proceso
             else

                strHTML = "<table>"
                strHTML = strHTML &  "<tr bgcolor='#1B5296'>"
                strHTML = strHTML &  "      <td colspan='4' class='textForm2'><div align='right'></div></td> "
                strHTML = strHTML &  " </tr> "
                strHTML = strHTML &  " <tr>  "
                strHTML = strHTML &  "    <td colspan='4'><div align='center'></div></td> "
                strHTML = strHTML &  " </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td width='250' rowspan='4' align='center'><img src='http://rkzego.no-ip.org/PortalMySQL/Extranet/ext-Images/computadora_animo.jpg' width='150' height='157'></td> "
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=4 COLOR=red>Espere un momento...</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=5 COLOR=red>La Base de Datos se esta Actualizando</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=5 COLOR=red>Genere este Reporte unos minutos mas tarde</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=3 COLOR=red>estamos trabajando para brindarle un mejor servicio</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='4'><div align='center'></div></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr bgcolor='#1B5296'><td colspan='4'></td></tr>"
                strHTML = strHTML &  "  </table>"
                response.Write(strHTML)
            end if



 else
      select case strCodError
        case "1"
         strMenjError = "Campo en Blanco Especifique!.."
      case "5","6"
         strMenjError = "Fechas Erroneas, Verifique!"
      case "7"
         strMenjError = "Registros No Encontrados!"
      end select
%>
    <table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
      <tr>
      <td><%=strMenjError%></td>
      </tr>
    </table>
    <br>
<% 
end if
%>

