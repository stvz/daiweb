
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<%
    ' ESTE ASP ES EL SEGUNDO Y ES PARA ADMINISTRADORES
    MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
    MM_EXTRANET_STRING_STATUS = ODBC_POR_ADUANA(Session("GAduana")&"_STATUS")

    Dim arrRefEtapas()


    Response.Buffer = TRUE
    Response.Addheader "Content-Disposition", "attachment;filename=Status_Contenedores.xls"
    Response.ContentType = "application/vnd.ms-excel"
    Server.ScriptTimeOut=100000

    strUsuario     = request.Form("user")
    strTipoUsuario = request.Form("TipoUser")
    strPermisos    = Request.Form("Permisos")

    permi          = PermisoClientes(Session("GAduana"),strPermisos,"cliE01")
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
    strFiltroCliente = ""
    strFiltroCliente = request.Form("txtCliente")
    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
       blnAplicaFiltro = true
    end if
    if blnAplicaFiltro then
       permi = " AND cvecli01 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
       permi = ""
    end if

    strDateIni = ""
    strDateFin = ""
    strTipoPedimento = ""
    strCodError      = "0"

    '*******************************************************
    strDateIni   = trim(request.Form("txtDateIni"))
    strDateFin   = trim(request.Form("txtDateFin"))
    strTipoConte = trim(request.Form("txttipoConte"))
    '*******************************************************

    '***************************************************************************************************************
    'strDescripcion    = trim(request.Form("txtDescripcion"))
    'strDateIni2       = trim(request.Form("txtDateIni2"))
    'strDateFin2       = trim(request.Form("txtDateFin2"))
    'strTipoPedimento2 = trim(request.Form("rbnTipoDate2"))
    'strTipoFiltro     = trim(request.Form("TipoFiltro"))
    'rbnTipoReporte    = trim(request.Form("rbnTipoReporte"))
    'strDateIni=trim(request.Form("txtDateIni"))
    'strDateFin=trim(request.Form("txtDateFin"))
    'strTipoPedimento=trim(request.Form("rbnTipoDate"))

    if not isdate(strDateIni) then
      strCodError = "5"
    end if
    if not isdate(strDateFin) then
      strCodError = "6"
    end if
    if strDateIni="" or strDateFin="" then
      strCodError = "1"
    end if

    if strCodError = "0" then

    strHTML = ""
    tmpTipo = ""
    strSQL  = ""

    'Todos           - Sin filtro
    'ContePuerto     - Sin fecha de despacho
    'ConteTranPlant  - Con fecha de despacho y sin fecha de llegada  planta
    'ContePlant      - Con fecha de llegada  planta y sin fecha de salida a planta
    'ConteTranPuerto - Con fecha de salida a planta y sin fecha de vacio de contenedor
    'ConteFVacio     - Con fecha de Vacio de contenedor

         if strTipoConte  = "Todos" then 'Sin filtro
          '  'strSQL = "SELECT tipopr01, valmer01,factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecent01 FROM ssdagi01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
            strSQL = " SELECT C01REFER.REFE01 as referencia,     " & _
                     "        MARC01          as contenedor,     " & _
                     "        desf0101        as facturas,       " & _
                     "        fecpag01        as fechaPago,      " & _
                     "        fdsp01          as fechaDespacho,  " & _
                     "        fcarta01        as fechaCartaVacio " & _
                     " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01         " & _
                     "                                     INNER JOIN D01CONTE                " & _
                     "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                     " WHERE FREC01 > '2005-01-01'  AND                                       " & _
                     "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                     "        firmae01   <> ''  and                                           " & _
                     "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                     "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                     "        C01REFER.REFE01 <> ''  " & Permi & _
                     " UNION ALL                                                              " & _
                     " SELECT C01REFER.REFE01 as referencia,     " & _
                     "        MARC01          as contenedor,     " & _
                     "        desf0101        as facturas,       " & _
                     "        fecpag01        as fechaPago,      " & _
                     "        fdsp01          as fechaDespacho,  " & _
                     "        fcarta01        as fechaCartaVacio " & _
                     " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01         " & _
                     "                                     INNER JOIN D01CONTE                " & _
                     "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                     " WHERE FREC01 > '2005-01-01'  AND                                       " & _
                     "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                     "        firmae01   <> ''  and                                           " & _
                     "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                     "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                     "                     C01REFER.REFE01 <> ''      " & Permi & _
                     " order by referencia, contenedor "
             'Response.Write(strSQL)
             'response.end
         else
            if strTipoConte  = "ContePuerto" then  'Sin fecha de despacho
                  strSQL = " SELECT C01REFER.REFE01 as referencia,     " & _
                           "        MARC01          as contenedor,     " & _
                           "        desf0101        as facturas,       " & _
                           "        fecpag01        as fechaPago,      " & _
                           "        fdsp01          as fechaDespacho,  " & _
                           "        fcarta01        as fechaCartaVacio " & _
                           " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01         " & _
                           "                                     INNER JOIN D01CONTE                " & _
                           "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                           " WHERE ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                           "        firmae01   <> ''  and                                           " & _
                           "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                           "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                           "        fdsp01 = ''                                   AND               " & _
                           "        C01REFER.REFE01 <> ''  " & Permi & _
                           " UNION ALL                                                              " & _
                           " SELECT C01REFER.REFE01 as referencia,     " & _
                           "        MARC01          as contenedor,     " & _
                           "        desf0101        as facturas,       " & _
                           "        fecpag01        as fechaPago,      " & _
                           "        fdsp01          as fechaDespacho,  " & _
                           "        fcarta01        as fechaCartaVacio " & _
                           " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01         " & _
                           "                                     INNER JOIN D01CONTE                " & _
                           "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                           " WHERE  ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                           "        firmae01   <> ''  and                                           " & _
                           "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                           "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                           "        fdsp01 = ''                                   AND               " & _
                           "                     C01REFER.REFE01 <> ''      " & Permi & _
                           " order by referencia, contenedor "
                   'Response.Write(strSQL)
                   'response.end
            else
                  if strTipoConte  = "ConteTranPlant" then 'Con fecha de despacho y sin fecha de llegada  planta
                        'strSQL =  " SELECT DISTINCT  C01REFER.REFE01 as referencia,                                      " & _
                        '          "        MARC01          as contenedor,                                                " & _
                        '          "        desf0101        as facturas,                                                  " & _
                        '          "        fecpag01        as fechaPago,                                                 " & _
                        '          "        fdsp01          as fechaDespacho,                                             " & _
                        '          "        fcarta01        as fechaCartaVacio                                            " & _
                        '          " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01                       " & _
                        '          "               INNER JOIN D01CONTE  ON C01REFER.REFE01=D01CONTE.REFE01                " & _
                        '          "               INNER JOIN RKU_STATUS.ETXCOI ON                                        " & _
                        '          "                     (RKU_STATUS.ETXCOI.C_REFERENCIA <> C01REFER.REFE01  AND          " & _
                        '          "                     RKU_STATUS.ETXCOI.C_CONTE<> MARC01)                              " & _
                        '          "               INNER JOIN RKU_STATUS.ETAPS  ON                                        " & _
                        '          "                     RKU_STATUS.ETXCOI.N_ETAPA = RKU_STATUS.ETAPS.N_ETAPA AND         " & _
                        '          "                     D_ABREV = 'LLP'                                                  " & _
                        '          " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') AND                                      " & _
                        '          "       FREC01 > '2005-01-01'                 AND                                      " & _
                        '          "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                        '          "       firmae01 <> ''                        AND                                      " & _
                        '          "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"'               AND               " & _
                        '          "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"'               AND               " & _
                        '          "       fdsp01   <> ''                        AND                                      " & _
                        '          "       fcarta01 = ''                         AND                                      " & _
                        '          "       C01REFER.REFE01 <> ''                 AND                                      " & _
                        '          "       RKU_STATUS.ETXCOI.f_fecha <> ''                                                " & _
                        '                  Permi                                                                            & _
                        '          " UNION ALL                                                                            " & _
                        '          " SELECT DISTINCT  C01REFER.REFE01 as referencia,                                      " & _
                        '          "        MARC01          as contenedor,                                                " & _
                        '          "        desf0101        as facturas,                                                  " & _
                        '          "        fecpag01        as fechaPago,                                                 " & _
                        '          "        fdsp01          as fechaDespacho,                                             " & _
                        '          "        fcarta01        as fechaCartaVacio                                            " & _
                        '          " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01                       " & _
                        '          "               INNER JOIN D01CONTE ON C01REFER.REFE01=D01CONTE.REFE01                 " & _
                        '          "               INNER JOIN RKU_STATUS.ETXCOI ON                                        " & _
                        '          "                     (RKU_STATUS.ETXCOI.C_REFERENCIA <> C01REFER.REFE01  AND          " & _
                        '          "                     RKU_STATUS.ETXCOI.C_CONTE <> MARC01 )                            " & _
                        '          "               INNER JOIN RKU_STATUS.ETAPS  ON                                        " & _
                        '          "                     RKU_STATUS.ETXCOI.N_ETAPA = RKU_STATUS.ETAPS.N_ETAPA AND         " & _
                        '          "                     D_ABREV = 'LLP'                                                  " & _
                        '          " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') AND                                      " & _
                        '          "       FREC01 > '2005-01-01'                 AND                                      " & _
                        '          "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                        '          "       firmae01   <> ''                      AND                                      " & _
                        '          "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND                             " & _
                        '          "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND                             " & _
                        '          "       fdsp01 <> ''                          AND                                      " & _
                        '          "       fcarta01 = ''                         AND                                      " & _
                        '          "       C01REFER.REFE01 <> ''                 AND                                      " & _
                        '          "       RKU_STATUS.ETXCOI.f_fecha <> ''                                                " & _
                        '                  Permi                                                                            & _
                        '          " order by referencia, contenedor                                                      "

                       strSQL =  " SELECT C01REFER.REFE01 as referencia,     " & _
                                 "        MARC01          as contenedor,     " & _
                                 "        desf0101        as facturas,       " & _
                                 "        fecpag01        as fechaPago,      " & _
                                 "        fdsp01          as fechaDespacho,  " & _
                                 "        fcarta01        as fechaCartaVacio " & _
                                 " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01         " & _
                                 "                                     INNER JOIN D01CONTE                " & _
                                 "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                 " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and                        " & _
                                 "       FREC01 > '2005-01-01'  AND                                       " & _
                                 "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                 "        firmae01   <> ''  and                                           " & _
                                 "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                 "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                 "        fdsp01 <> ''                                  AND               " & _
                                 "        fcarta01 = ''                                 AND               " & _
                                 "        C01REFER.REFE01 <> ''  " & Permi & _
                                 " UNION ALL                                                              " & _
                                 " SELECT C01REFER.REFE01 as referencia,     " & _
                                 "        MARC01          as contenedor,     " & _
                                 "        desf0101        as facturas,       " & _
                                 "        fecpag01        as fechaPago,      " & _
                                 "        fdsp01          as fechaDespacho,  " & _
                                 "        fcarta01        as fechaCartaVacio " & _
                                 " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01         " & _
                                 "                                     INNER JOIN D01CONTE                " & _
                                 "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                 " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and " & _
                                 "       FREC01 > '2005-01-01'  AND " & _
                                 "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                 "        firmae01   <> ''  and                                           " & _
                                 "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                 "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                 "        fdsp01 <> ''                                  AND               " & _
                                 "        fcarta01 = ''                                 AND               " & _
                                 "                     C01REFER.REFE01 <> ''      " & Permi & _
                                 " order by referencia, contenedor "

                         'Response.Write(strSQL)
                         'response.end
                  else
                        if strTipoConte  = "ContePlant" then 'Con fecha de llegada  planta y sin fecha de salida a planta
                         strSQL =  " SELECT C01REFER.REFE01 as referencia,     " & _
                                   "        MARC01          as contenedor,     " & _
                                   "        desf0101        as facturas,       " & _
                                   "        fecpag01        as fechaPago,      " & _
                                   "        fdsp01          as fechaDespacho,  " & _
                                   "        fcarta01        as fechaCartaVacio " & _
                                   " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01         " & _
                                   "                                     INNER JOIN D01CONTE                " & _
                                   "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                   " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and                        " & _
                                   "       FREC01 > '2005-01-01'  AND                                       " & _
                                   "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                   "        firmae01   <> ''  and                                           " & _
                                   "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                   "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                   "        fdsp01 <> ''                                  AND               " & _
                                   "        fcarta01 = ''                                 AND               " & _
                                   "        C01REFER.REFE01 <> ''  " & Permi & _
                                   " UNION ALL                                                              " & _
                                   " SELECT C01REFER.REFE01 as referencia,     " & _
                                   "        MARC01          as contenedor,     " & _
                                   "        desf0101        as facturas,       " & _
                                   "        fecpag01        as fechaPago,      " & _
                                   "        fdsp01          as fechaDespacho,  " & _
                                   "        fcarta01        as fechaCartaVacio " & _
                                   " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01         " & _
                                   "                                     INNER JOIN D01CONTE                " & _
                                   "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                   " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and " & _
                                   "       FREC01 > '2005-01-01'  AND " & _
                                   "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                   "        firmae01   <> ''  and                                           " & _
                                   "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                   "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                   "        fdsp01 <> ''                                  AND               " & _
                                   "        fcarta01 = ''                                 AND               " & _
                                   "                     C01REFER.REFE01 <> ''      " & Permi & _
                                   " order by referencia, contenedor "

                             'strSQL = " SELECT DISTINCT  C01REFER.REFE01 as referencia,                              " & _
                             '         "                  MARC01          as contenedor,                              " & _
                             '         "                  desf0101        as facturas,                                " & _
                             '         "                  fecpag01        as fechaPago,                               " & _
                             '         "                  fdsp01          as fechaDespacho,                           " & _
                             '         "                  fcarta01        as fechaCartaVacio                          " & _
                             '         " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01               " & _
                             '         "               INNER JOIN D01CONTE  ON C01REFER.REFE01=D01CONTE.REFE01        " & _
                             '         "               INNER JOIN RKU_STATUS.ETXCOI ON                                " & _
                             '         "                     RKU_STATUS.ETXCOI.C_REFERENCIA = C01REFER.REFE01  AND    " & _
                             '         "                     RKU_STATUS.ETXCOI.C_CONTE= MARC01                        " & _
                             '         "               INNER JOIN RKU_STATUS.ETAPS  ON                                " & _
                             '         "                     RKU_STATUS.ETXCOI.N_ETAPA = RKU_STATUS.ETAPS.N_ETAPA AND " & _
                             '         "                     RKU_STATUS.ETAPS.D_ABREV = 'LLP'                         " & _
                             '         " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') AND                        " & _
                             '         "       FREC01 > '2005-01-01'                 AND                        " & _
                             '         "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and  " & _
                             '         "       firmae01   <> ''        AND                                      " & _
                             '         "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                             '         "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                             '         "       fdsp01 <> ''            AND                                      " & _
                             '         "       C01REFER.REFE01 <> ''   AND                                      " & _
                             '         "       fcarta01 = ''                                 AND                " & _
                             '         "       RKU_STATUS.ETXCOI.f_fecha <> ''                                  " & _
                             '                 Permi & _
                             '         " UNION ALL                                                                    " & _
                             '         " SELECT  DISTINCT  C01REFER.REFE01 as referencia,                             " & _
                             '         "         MARC01          as contenedor,                                       " & _
                             '         "         desf0101        as facturas,                                         " & _
                             '         "         fecpag01        as fechaPago,                                        " & _
                             '         "         fdsp01          as fechaDespacho,                                    " & _
                             '         "         fcarta01        as fechaCartaVacio                                   " & _
                             '         " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01               " & _
                             '         "               INNER JOIN D01CONTE ON C01REFER.REFE01=D01CONTE.REFE01         " & _
                             '         "               INNER JOIN RKU_STATUS.ETXCOI ON                                " & _
                             '         "                     RKU_STATUS.ETXCOI.C_REFERENCIA = C01REFER.REFE01  AND    " & _
                             '         "                     RKU_STATUS.ETXCOI.C_CONTE = MARC01                       " & _
                             '         "               INNER JOIN RKU_STATUS.ETAPS  ON                                " & _
                             '         "                     RKU_STATUS.ETXCOI.N_ETAPA = RKU_STATUS.ETAPS.N_ETAPA AND " & _
                             '         "                     RKU_STATUS.ETAPS.D_ABREV = 'LLP'                         " & _
                             '         " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and                        " & _
                             '         "       FREC01 > '2005-01-01'  AND                                       " & _
                             '         "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and  " & _
                             '         "       firmae01   <> ''  and                                            " & _
                             '         "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                             '         "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                             '         "       fdsp01 <> ''  AND                                                " & _
                             '         "       C01REFER.REFE01 <> ''   AND                                      " & _
                             '         "        fcarta01 = ''                                 AND               " & _
                             '         "       RKU_STATUS.ETXCOI.f_fecha <> ''                                  " & _
                             '                 Permi & _
                             '         " order by referencia, contenedor                                        "


                                       ' " SELECT C01REFER.REFE01 as referencia,     " & _
                                       ' "        MARC01          as contenedor,     " & _
                                       ' "        desf0101        as facturas,       " & _
                                       ' "        fecpag01        as fechaPago,      " & _
                                       ' "        fdsp01          as fechaDespacho,  " & _
                                       ' "        fcarta01        as fechaCartaVacio " & _
                                       ' " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01         " & _
                                       ' "                                     INNER JOIN D01CONTE                " & _
                                       ' "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                       ' " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and                        " & _
                                       ' "       FREC01 > '2005-01-01'  AND                                       " & _
                                       ' "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                       ' "        firmae01   <> ''  and                                           " & _
                                       ' "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                       ' "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                       ' "        fdsp01 <> ''                                  AND               " & _
                                       ' "        C01REFER.REFE01 <> ''  " & Permi & _
                                       ' " UNION ALL                                                              " & _
                                       ' " SELECT C01REFER.REFE01 as referencia,     " & _
                                       ' "        MARC01          as contenedor,     " & _
                                       ' "        desf0101        as facturas,       " & _
                                       ' "        fecpag01        as fechaPago,      " & _
                                       ' "        fdsp01          as fechaDespacho,  " & _
                                       ' "        fcarta01        as fechaCartaVacio " & _
                                       ' " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01         " & _
                                       ' "                                     INNER JOIN D01CONTE                " & _
                                       ' "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                       ' " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and " & _
                                       ' "       FREC01 > '2005-01-01'  AND " & _
                                       ' "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                       ' "        firmae01   <> ''  and                                           " & _
                                       ' "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                       ' "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                       ' "        fdsp01 <> ''                                  AND               " & _
                                       ' "                     C01REFER.REFE01 <> ''      " & Permi & _
                                       ' " order by referencia, contenedor "
                               'Response.Write(strSQL)
                               'response.end
                        else
                            if strTipoConte  = "ConteTranPuerto" then 'ConteTranPuerto - Con fecha de salida a planta y sin fecha de vacio de contenedor
                                 strSQL =  " SELECT C01REFER.REFE01 as referencia,     " & _
                                           "        MARC01          as contenedor,     " & _
                                           "        desf0101        as facturas,       " & _
                                           "        fecpag01        as fechaPago,      " & _
                                           "        fdsp01          as fechaDespacho,  " & _
                                           "        fcarta01        as fechaCartaVacio " & _
                                           " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01         " & _
                                           "                                     INNER JOIN D01CONTE                " & _
                                           "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                           " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and                        " & _
                                           "       FREC01 > '2005-01-01'  AND                                       " & _
                                           "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                           "        firmae01   <> ''  and                                           " & _
                                           "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                           "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                           "        fdsp01 <> ''                                  AND               " & _
                                           "        fcarta01 = ''                                 AND               " & _
                                           "        C01REFER.REFE01 <> ''  " & Permi & _
                                           " UNION ALL                                                              " & _
                                           " SELECT C01REFER.REFE01 as referencia,     " & _
                                           "        MARC01          as contenedor,     " & _
                                           "        desf0101        as facturas,       " & _
                                           "        fecpag01        as fechaPago,      " & _
                                           "        fdsp01          as fechaDespacho,  " & _
                                           "        fcarta01        as fechaCartaVacio " & _
                                           " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01         " & _
                                           "                                     INNER JOIN D01CONTE                " & _
                                           "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                           " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and " & _
                                           "       FREC01 > '2005-01-01'  AND " & _
                                           "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                           "        firmae01   <> ''  and                                           " & _
                                           "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                           "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                           "        fdsp01 <> ''                                  AND               " & _
                                           "        fcarta01 = ''                                 AND               " & _
                                           "                     C01REFER.REFE01 <> ''      " & Permi & _
                                           " order by referencia, contenedor "

                                  'strSQL = " SELECT C01REFER.REFE01 as referencia,     " & _
                                  '         "        MARC01          as contenedor,     " & _
                                  '         "        desf0101        as facturas,       " & _
                                  '         "        fecpag01        as fechaPago,      " & _
                                  '         "        fdsp01          as fechaDespacho,  " & _
                                  '         "        fcarta01        as fechaCartaVacio " & _
                                  '         " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01         " & _
                                  '         "                                     INNER JOIN D01CONTE                " & _
                                  '         "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                  '         " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and                        " & _
                                  '         "       FREC01 > '2005-01-01'  AND                                       " & _
                                  '         "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                  '         "        firmae01   <> ''  and                                           " & _
                                  '         "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                  '         "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                  '         "        fcarta01 = ''                                 AND               " & _
                                  '         "        C01REFER.REFE01 <> ''  " & Permi & _
                                  '         " UNION ALL                                                              " & _
                                  '         " SELECT C01REFER.REFE01 as referencia,     " & _
                                  '         "        MARC01          as contenedor,     " & _
                                  '         "        desf0101        as facturas,       " & _
                                  '         "        fecpag01        as fechaPago,      " & _
                                  '         "        fdsp01          as fechaDespacho,  " & _
                                  '         "        fcarta01        as fechaCartaVacio " & _
                                  '         " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01         " & _
                                  '         "                                     INNER JOIN D01CONTE                " & _
                                  '         "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                  '         " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and " & _
                                  '         "       FREC01 > '2005-01-01'  AND " & _
                                  '         "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                  '         "        firmae01   <> ''  and                                           " & _
                                  '         "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                  '         "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                  '         "        fcarta01 = ''                                 AND               " & _
                                  '         "        C01REFER.REFE01 <> ''      " & Permi & _
                                  '         " order by referencia, contenedor "

                                   'Response.Write(strSQL)
                                   'response.end
                            else
                                  if strTipoConte  = "ConteFVacio" then    'Con fecha de Vacio de contenedor
                                        strSQL = " SELECT C01REFER.REFE01 as referencia,     " & _
                                                 "        MARC01          as contenedor,     " & _
                                                 "        desf0101        as facturas,       " & _
                                                 "        fecpag01        as fechaPago,      " & _
                                                 "        fdsp01          as fechaDespacho,  " & _
                                                 "        fcarta01        as fechaCartaVacio " & _
                                                 " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01         " & _
                                                 "                                     INNER JOIN D01CONTE                " & _
                                                 "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                                 " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and                        " & _
                                                 "       FREC01 > '2005-01-01'  AND                                       " & _
                                                 "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                                 "        firmae01   <> ''  and                                           " & _
                                                 "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                                 "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                                 "        fcarta01 <> ''                                AND               " & _
                                                 "        C01REFER.REFE01 <> ''  " & Permi & _
                                                 " UNION ALL                                                              " & _
                                                 " SELECT C01REFER.REFE01 as referencia,     " & _
                                                 "        MARC01          as contenedor,     " & _
                                                 "        desf0101        as facturas,       " & _
                                                 "        fecpag01        as fechaPago,      " & _
                                                 "        fdsp01          as fechaDespacho,  " & _
                                                 "        fcarta01        as fechaCartaVacio " & _
                                                 " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01         " & _
                                                 "                                     INNER JOIN D01CONTE                " & _
                                                 "                                     ON C01REFER.REFE01=D01CONTE.REFE01 " & _
                                                 " WHERE ( CSIT01 <> 'FIN' and  cgas01 <> 'F') and " & _
                                                 "       FREC01 > '2005-01-01'  AND " & _
                                                 "       ( cvep01 <> 'R1' AND CVEP01 <> 'A3' AND CVEP01 <> 'F4' AND CVEP01 <> 'BB') and " & _
                                                 "        firmae01   <> ''  and                                           " & _
                                                 "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND               " & _
                                                 "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND               " & _
                                                 "        fcarta01 <> ''                                AND               " & _
                                                 "                     C01REFER.REFE01 <> ''      " & Permi & _
                                                 " order by referencia, contenedor "
                                         'Response.Write(strSQL)
                                         'response.end
                                  end if

                            end if

                        end if

                  end if

            end if

         end if


                              'strSQL = "SELECT concat(concat(concat(concat(concat(concat(ltrim(substring(year(FECPAG01),3,2)),'-'),CVEADU01),'-'),PATENT01),'-'),NUMPED01) as IMPORTA," & _
                     '         "       adusec01 as Aduana,        " & _
                     '         "       fecpag01 as pago,          " & _
                     '         "       TIPCAM01 as TipoCambio,    " & _
                     '         "       cveped01 as clavePedimento," & _
                     '         "       FLETES01 as FLETE,         " & _
                     '         "       SEGROS01 as SEGUROS,       " & _
                     '         "       Embala01 as embalaje,      " & _
                     '         "       Incble01 as OtrosIncbles,  " & _
                     '         "       anexol01 as observa,       " & _
                     '         "       refcia01 as Referencia,    " & _
                     '         "       FACTMO01   as FactorMoneda,     " & _
                     '         "       sum(vaduan02)   as valoraduana, " & _
                     '         "       sum(vmerme02)   as valorComer,  " & _
                     '         "       sum(vmerme02*FACTMO01) as valorExtra,      " & _
                     '         "       sum(vmerme02*FACTMO01*TIPCAM01) as valorMN " & _
                     '         "FROM ssdagi01, " & _
                     '         "     ssfrac02  " & _
                     '         "WHERE  Refcia01 = refcia02  and " & _
                     '         "       (cveped01 <> 'R1')   and " & _
                     '         "       (firmae01   <> '')   and " & _
                     '         "       fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                     '         "       fecpag01 <= '"&FormatoFechaInv(strDateFin)&"'  " & _
                     '         Permi & _
                     '         " GROUP BY REFCIA01"
                 'else
                 '    if strTipoConte  = "Todos" then
                 '       tmpTipo = "EXPORTACION"
                 '       'strSQL = "SELECT tipopr01, factmo01, p_dta101, t_reca01, i_dta101, cvecli01, refcia01, fecpag01, valfac01, fletes01, segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01 FROM ssdage01 WHERE fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
                 '        strSQL = "SELECT concat(concat(concat(concat(concat(concat(ltrim(substring(year(FECPAG01),3,2)),'-'),CVEADU01),'-'),PATENT01),'-'),NUMPED01) as IMPORTA," & _
                 '                 "       adusec01 as Aduana,        " & _
                 '                 "       fecpag01 as pago,          " & _
                 '                 "       TIPCAM01 as TipoCambio,    " & _
                 '                 "       cveped01 as clavePedimento," & _
                 '                 "       FLETES01 as FLETE,         " & _
                 '                 "       SEGROS01 as SEGUROS,       " & _
                 '                 "       Embala01 as embalaje,      " & _
                 '                 "       Incble01 as OtrosIncbles,  " & _
                 '                 "       anexol01 as observa,       " & _
                 '                 "       refcia01 as Referencia,    " & _
                 '                 "       FACTMO01   as FactorMoneda,     " & _
                 '                 "       sum(vaduan02)   as valoraduana, " & _
                 '                 "       sum(vmerme02)   as valorComer,  " & _
                 '                 "       sum(vmerme02*FACTMO01) as valorExtra,      " & _
                 '                 "       sum(vmerme02*FACTMO01*TIPCAM01) as valorMN " & _
                 '                 "FROM ssdage01, " & _
                 '                 "     ssfrac02  " & _
                 '                 "WHERE  Refcia01 = refcia02  and " & _
                 '                 "       (cveped01 <> 'R1')   and " & _
                 '                 "       (firmae01   <> '')   and " & _
                 '                 "       fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                 '                 "       fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND " & _
                 '                 "       LTRIM(refcia01) <> 'GABBY'  " & _
                 '                 Permi & _
                 '                 " GROUP BY REFCIA01"
                 '    end if

         'response.write(strSQL)
         'response.end


         if not trim(strSQL)="" then
            Set RsRep = Server.CreateObject("ADODB.Recordset")
            RsRep.ActiveConnection = MM_EXTRANET_STRING
            RsRep.Source = strSQL
            'RsRep.Source = " SELECT * FROM SSDAGI01 "
            RsRep.CursorType = 0
            RsRep.CursorLocation = 2
            RsRep.LockType = 1

            RsRep.Open()

            'response.write( RsRep.recordCount )
            'response.write(strSQL)
            'response.end


            if not RsRep.eof then
               'Comienza el HTML, se pintan los titulos de las columnas
               'strHTML = strHTML & " <p> <img src='../../ext-Images/Gifs/abbot.gif'> </p>"
               'strHTML = strHTML & " <p> <img width='181' eight='38'  src='http://10.66.1.4/PortalMySQL/Extranet/ext-Images/Gifs/abbot.gif'> </p> <P>&nbsp;</P>"

               'referencia, contenedor, factura(s), dias (hoy-despacho), pago, despacho, llp, spl, vcont.

               'strHTML = strHTML & " <p> &nbsp; </p>"
               strHTML = strHTML & " <br> "
               strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">GRUPO REYES KURI, S.C. </font></strong> <br> "

               StrtituloRep = ""
               if strTipoConte  = "Todos" then
                  StrtituloRep = "Contenedores"
               else
                  if strTipoConte  = "ContePuerto" then
                     StrtituloRep = "Contenedores en Puerto"
                  else
                     if strTipoConte  = "ConteTranPlant" then
                        StrtituloRep = "Contenedores en Transito a Planta"
                     else
                        if strTipoConte  = "ContePlant" then
                           StrtituloRep = "Contenedores en Planta"
                        else
                           if strTipoConte  = "ConteTranPuerto" then
                              StrtituloRep = "Contenedores en Transito a Puerto"
                           else
                             if strTipoConte  = "ConteFVacio" then
                                StrtituloRep = "Contenedores ya Retornados"
                             end if
                           end if
                        end if
                     end if
                  end if
               end if

               'Todos           - Sin filtro
               'ContePuerto     - Sin fecha de despacho
               'ConteTranPlant  - Con fecha de despacho y sin fecha de llegada  planta
               'ContePlant      - Con fecha de llegada  planta y sin fecha de salida a planta
               'ConteTranPuerto - Con fecha de salida a planta y sin fecha de vacio de contenedor
               'ConteFVacio     - Con fecha de Vacio de contenedor


               strHTML = strHTML & "<strong><font color=""#969696"" size=""3"" face=""Arial, Helvetica, sans-serif""> Reporte de Status de " & StrtituloRep & " del " & strDateIni & " al " & strDateFin & " </font></strong>"
               strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
               strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)

               strHTML = strHTML & "<td width=""90""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REFERENCIA     </font></strong></td>" & chr(13) & chr(10) 'REFERENCIA
               strHTML = strHTML & "<td width=""90""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CONTENEDOR     </font></strong></td>" & chr(13) & chr(10) 'CONTENEDOR
               strHTML = strHTML & "<td width=""100"" bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FACTURAS       </font></strong></td>" & chr(13) & chr(10) 'FACTURAS
               strHTML = strHTML & "<td width=""30""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DIAS           </font></strong></td>" & chr(13) & chr(10) 'DIAS
               strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PAGO           </font></strong></td>" & chr(13) & chr(10) 'PAGO
               strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESPACHO       </font></strong></td>" & chr(13) & chr(10) 'DESPACHO
               strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> LLP            </font></strong></td>" & chr(13) & chr(10) 'LLP
               strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SPL            </font></strong></td>" & chr(13) & chr(10) 'SPL
               strHTML = strHTML & "<td width=""55""  bgcolor=""#339966"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VCONT          </font></strong></td>" & chr(13) & chr(10) 'VCONT
               strHTML = strHTML & "</tr>"& chr(13) & chr(10)

               'response.end

               intColumna = 1
               While NOT RsRep.EOF
                 'Vamos a filtrar los registros

                  Dim strFechaContellP
                  DIm strFechaConteSPL
                  banFilReg = True
                  strRefer      = RsRep.Fields.Item("Referencia").Value
                  strContenedor = RsRep.Fields.Item("contenedor").Value

                  'banFilReg = 0
                  'if strTipoConte  = "ConteTranPlant" then 'Con fecha de despacho y sin fecha de llegada  planta
                  '       'Response.End
                  'else
                  '  if strTipoConte  = "ContePlant" then 'Con fecha de llegada  planta y sin fecha de salida a planta
                  '
                  '  else
                  '    if strTipoConte  = "ConteTranPuerto" then 'ConteTranPuerto - Con fecha de salida a planta y sin fecha de vacio de contenedor
                  '
                  '    end if
                  '  end if
                  'end if
                         '**************************************************************************************************************
                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                         strSqlSel =  " SELECT D.n_secuenc, " & _
                                      "        D.n_etapa,   " & _
                                      "        D.f_fecha,   " & _
                                      "        D.m_observ , " & _
                                      "        F.d_nombre,  " & _
                                      "        F.d_abrev    " & _
                                      " FROM etxcoi D,      " & _
                                      "      etaps F        " & _
                                      " where D.n_etapa = F.n_etapa and            " & _
                                      "       D.c_referencia = '"& strRefer & "' and " & _
                                      "       D.c_conte      = '"& RsRep.Fields.Item("contenedor").Value & "' and " & _
                                      "       F.d_abrev      = 'LLP' " & _
                                      " order by D.n_secuenc desc "

                         'Response.Write(strSqlSel)
                         'Response.End
                         RConteDetalle.Source = strSqlSel
                         RConteDetalle.CursorType = 0
                         RConteDetalle.CursorLocation = 2
                         RConteDetalle.LockType = 1
                         RConteDetalle.Open()
                         if not RConteDetalle.eof then
                             strFechaContellP = RConteDetalle.Fields.Item("f_fecha").Value
                         else
                             strFechaContellP = ""
                         end if
                         RConteDetalle.close
                         set RConteDetalle = Nothing
                         'Response.End

                         '**************************************************************************************************************
                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                         strSqlSel =  " SELECT D.n_secuenc, " & _
                                      "        D.n_etapa,   " & _
                                      "        D.f_fecha,   " & _
                                      "        D.m_observ , " & _
                                      "        F.d_nombre,  " & _
                                      "        F.d_abrev    " & _
                                      " FROM etxcoi D,      " & _
                                      "      etaps F        " & _
                                      " where D.n_etapa = F.n_etapa and            " & _
                                      "       D.c_referencia = '"& strRefer & "' and " & _
                                      "       D.c_conte      = '"& RsRep.Fields.Item("contenedor").Value & "' and " & _
                                      "       F.d_abrev      = 'SPL' " & _
                                      " order by D.n_secuenc desc "

                         'Response.Write(strSqlSel)
                         'Response.End
                         RConteDetalle.Source = strSqlSel
                         RConteDetalle.CursorType = 0
                         RConteDetalle.CursorLocation = 2
                         RConteDetalle.LockType = 1
                         RConteDetalle.Open()
                         if not RConteDetalle.eof then
                             strFechaConteSPL = RConteDetalle.Fields.Item("f_fecha").Value
                         else
                             strFechaConteSPL = ""
                         end if
                         RConteDetalle.close
                         set RConteDetalle = Nothing
                         'Response.End

                         'if isnull(strFechaContellP) then
                         '   strFechaContellP = ""
                         'end if
                         'if isnull(strFechaConteSPL) then
                         '   strFechaConteSPL = ""
                         'end if

                         '*************************************************************************************************
                         ' referencia
                         ' contenedor
                         ' facturas
                         ' fechaPago
                         ' fechaDespacho
                         ' fechaCartaVacio
                         '*************************************************************************************************

                          banFilReg = True
                          'if strTipoConte  = "ConteTranPlant" then 'Con fecha de despacho y sin fecha de llegada  planta
                          '   if ( isnull(strFechaContellP) ) then
                          '      banFilReg = 0
                          '   else
                          '      banFilReg = 1
                          '   end if
                          'else
                          '  if strTipoConte  = "ContePlant" then 'Con fecha de llegada  planta y sin fecha de salida a planta
                          '     if not isnull(strFechaContellP) and isnull(strFechaConteSPL) then
                          '         banFilReg = 0
                          '     else
                          '         banFilReg = 1
                          '     end if
                          '  else
                          '    if strTipoConte  = "ConteTranPuerto" then 'ConteTranPuerto - Con fecha de salida a planta y sin fecha de vacio de contenedor
                          '       if not isnull(strFechaConteSPL) then
                          '           banFilReg = 0
                          '       else
                          '           banFilReg = 1
                          '       end if
                          '    end if
                          '  end if
                          'end if

                          'Response.Write( "registro=" & CStr(intColumna) )
                          'Response.Write( isempty(strFechaContellP) )
                          'Response.Write( "<br>" )
                          'Response.Write( isempty(strFechaConteSPL) )
                          'Response.Write( "<br>" )

                          if strTipoConte  = "ConteTranPlant" then 'Con fecha de despacho y sin fecha de llegada  planta
                             if strFechaContellP = "" and strFechaConteSPL = "" then
                                banFilReg = True
                             else
                                banFilReg = False
                             end if
                          else
                            if strTipoConte  = "ContePlant" then 'Con fecha de llegada  planta y sin fecha de salida a planta

                                 'Response.Write( intColumna)
                                 'Response.Write( strFechaContellP)
                                 'Response.Write( isempty(strFechaContellP) )
                                 'Response.Write( isnull(strFechaContellP) )
                                 'Response.Write( isDate(strFechaContellP) )

                                 if not isdate(strFechaContellP) then
                                    banFilReg = False
                                 else
                                      if not isdate(strFechaConteSPL)  then
                                           banFilReg = True
                                      else
                                           banFilReg = False
                                      end if
                                 end if
                            else
                              if strTipoConte  = "ConteTranPuerto" then 'ConteTranPuerto - Con fecha de salida a planta y sin fecha de vacio de contenedor
                                 if strFechaConteSPL = "" then
                                     banFilReg = False
                                 else
                                     banFilReg = True
                                 end if
                                 'if isempty(strFechaConteSPL) then
                                 '    banFilReg = True
                                 'else
                                 '    banFilReg = False
                                 'end if
                              end if
                            end if
                          end if


                          if banFilReg = True then
                                 '*************************************************************************************************

                                 redim preserve arrRefEtapas(2,intColumna)
                                 'strRefer      = RsRep.Fields.Item("Referencia").Value
                                 'strContenedor = RsRep.Fields.Item("contenedor").Value
                                 arrRefEtapas(0,intColumna-1) = trim(strRefer)
                                 arrRefEtapas(1,intColumna-1) = trim(strContenedor)
                                 '*************************************************************************************************

                                 intDifDias = 0
                                 Dhoy = Date()


                                 if not isnull(RsRep.Fields.Item("fechaCartaVacio").Value) and isdate(RsRep.Fields.Item("fechaCartaVacio").Value)  then
                                    intDifDias =  DateDiff ("d", RsRep.Fields.Item("fechaDespacho").Value, RsRep.Fields.Item("fechaCartaVacio").Value)
                                 else
                                    intDifDias =  DateDiff ("d", RsRep.Fields.Item("fechaDespacho").Value, Dhoy)
                                 end if


                                 strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("referencia").Value      &"&nbsp;  </font></td>" & chr(13) & chr(10) 'REFERENCIA
                                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("contenedor").Value      &"&nbsp;  </font></td>" & chr(13) & chr(10) 'CONTENEDOR
                                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("facturas").Value        &"&nbsp;  </font></td>" & chr(13) & chr(10) 'FACTURAS
                                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & intDifDias                                 &"&nbsp;  </font></td>" & chr(13) & chr(10) 'DIAS
                                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("fechaPago").Value       &"&nbsp;  </font></td>" & chr(13) & chr(10) 'PAGO
                                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("fechaDespacho").Value   &"&nbsp;  </font></td>" & chr(13) & chr(10) 'DESPACHO
                                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & strFechaContellP                           &"&nbsp;  </font></td>" & chr(13) & chr(10) 'LLP
                                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & strFechaConteSPL                           &"&nbsp;  </font></td>" & chr(13) & chr(10) 'SPL
                                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RsRep.Fields.Item("fechaCartaVacio").Value &"&nbsp;  </font></td>" & chr(13) & chr(10) 'VCONT
                                 strHTML = strHTML&"</tr>"& chr(13) & chr(10)
                                 intColumna = intColumna + 1
                          end if


                     RsRep.movenext


              'Response.Write( intColumna )
              'Response.Write(strHTML)
              'Response.End


               Wend

            strHTML = strHTML & "</table>"& chr(13) & chr(10)


            'response.write( UBound(arrRefEtapas) )

            strHTML = strHTML & "<br> <br> <br> "
            'strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">"& chr(13) & chr(10)
            strHTML = strHTML & "<table >"& chr(13) & chr(10)

            '******************************************************************************************
            '***                                                                                    ***

            if intColumna > 1 then
                for y=0 to (UBound(arrRefEtapas,2) - 1)
                  if not isnull( arrRefEtapas(0,y) ) and  not isnull( arrRefEtapas(1,y) ) then
                '    if ltrim(arrRefEtapas(y) ) = ltrim(CStr(RsRepPermi.Fields.Item("CANTIDADFAC").Value)) then

                       strCadenaRenglon= ""
                       strCadenaColumna= ""

                       strCadenaRenglon = strCadenaRenglon & "<tr>"
                       strCadenaRenglon = strCadenaRenglon & "<td colspan='2'><strong><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> REFERENCIA = " & arrRefEtapas(0,y) & " </font> </strong></td>"
                       strCadenaRenglon = strCadenaRenglon & "<td colspan='6'><strong><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> CONTENEDOR = " & arrRefEtapas(1,y) & " </font> </strong></td>"
                       strCadenaRenglon = strCadenaRenglon & "</tr>"
                       '***********************************************************************

                       strCadenaRenglon = strCadenaRenglon & "<tr>"
                       strCadenaRenglon = strCadenaRenglon & "<td></td>"
                       strCadenaRenglon = strCadenaRenglon & "<td colspan='7'> <font color=""#006600"" size=""1"" face=""Arial, Helvetica, sans-serif""> "

                             Set RConteEtapas = Server.CreateObject("ADODB.Recordset")
                             RConteEtapas.ActiveConnection = MM_EXTRANET_STRING_STATUS
                             strSqlConteEtapas =  " SELECT n_etapa,  " & _
                                                  "        d_abrev   " & _
                                                  " FROM etaps       " & _
                                                  " order by n_etapa "
                             'Response.Write(strSqlSel)
                             'Response.End
                             RConteEtapas.Source = strSqlConteEtapas
                             RConteEtapas.CursorType = 0
                             RConteEtapas.CursorLocation = 2
                             RConteEtapas.LockType = 1
                             RConteEtapas.Open()

                             if not RConteEtapas.eof then
                                while NOT RConteEtapas.EOF
                                  'RConteEtapas.Fields.Item("n_etapa").Value
                                  '****************************************


                                      Set RsDetalle = Server.CreateObject("ADODB.Recordset")
                                        RsDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                        strSQL = " SELECT D.n_etapa,   " & _
                                                 "        D.f_fecha,   " & _
                                                 "        D.m_observ,  " & _
                                                 "        F.d_nombre,  " & _
                                                 "        F.d_abrev,   " & _
                                                 "        D.n_secuenc  " & _
                                                 " FROM ETXPD as D,    " & _
                                                 "      etaps F        " & _
                                                 " WHERE D.c_referencia = '" & arrRefEtapas(0,y) & "' and " & _
                                                 "       D.n_etapa =  F.n_etapa   and  " & _
                                                 "       D.n_etapa ='"& RConteEtapas.Fields.Item("n_etapa").Value & "' AND " & _
                                                 "       trim(D.L_VISIBLE) = 'T'  and " & _
                                                 "       m_observ <> '' " & _
                                                 " ORDER BY D.n_secuenc    desc "

                                        'response.write(strSQL)
                                        'Response.End
                                        RsDetalle.Source = strSQL
                                        RsDetalle.CursorType = 0
                                        RsDetalle.CursorLocation = 2
                                        RsDetalle.LockType = 1
                                        RsDetalle.Open()

                                         if not RsDetalle.eof then
                                             'while NOT RsDetalle.EOF
                                               if RsDetalle.Fields.Item("m_observ").Value <> "" then
                                                   strCadenaColumna = strCadenaColumna &  RsDetalle.Fields.Item("d_nombre").Value&" ("& RsDetalle.Fields.Item("d_abrev").Value&") "&RsDetalle.Fields.Item("f_fecha").Value&".-"&mid(RsDetalle.Fields.Item("m_observ").Value,1,135) & " <br>"
                                               end if
                                        '     '  RsDetalle.MoveNext
                                        '     'wend
                                         end if
                                        RsDetalle.close
                                        set RsDetalle = nothing
                                  '****************************************

                                  RConteEtapas.MoveNext
                                wend
                             end if
                             RConteEtapas.close
                             set RConteEtapas = Nothing
                             'Response.End


                       '***********************************************************************
                             Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                             RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                             strSqlSel =  " SELECT D.n_secuenc, " & _
                                          "        D.n_etapa,   " & _
                                          "        D.f_fecha,   " & _
                                          "        D.m_observ , " & _
                                          "        F.d_nombre,  " & _
                                          "        F.d_abrev    " & _
                                          " FROM etxcoi D,      " & _
                                          "      etaps F        " & _
                                          " where D.n_etapa = F.n_etapa and            " & _
                                          "       D.c_referencia = '"& arrRefEtapas(0,y) & "' and " & _
                                          "       D.c_conte      = '"& arrRefEtapas(1,y) & "' and " & _
                                          "       F.d_abrev      = 'LLP' " & _
                                          " order by D.n_secuenc desc "

                             'Response.Write(strSqlSel)
                             'Response.End
                             RConteDetalle.Source = strSqlSel
                             RConteDetalle.CursorType = 0
                             RConteDetalle.CursorLocation = 2
                             RConteDetalle.LockType = 1
                             RConteDetalle.Open()
                             if not RConteDetalle.eof then
                                 strCadenaColumna = strCadenaColumna & RConteDetalle.Fields.Item("d_nombre").Value&" ("&RConteDetalle.Fields.Item("d_abrev").Value&") "&RConteDetalle.Fields.Item("f_fecha").Value&".-"&mid(RConteDetalle.Fields.Item("m_observ").Value,1,135) & " <br>"
                             end if
                             RConteDetalle.close
                             set RConteDetalle = Nothing
                             'Response.End

                             '**************************************************************************************************************
                             Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                             RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                             strSqlSel =  " SELECT D.n_secuenc, " & _
                                          "        D.n_etapa,   " & _
                                          "        D.f_fecha,   " & _
                                          "        D.m_observ , " & _
                                          "        F.d_nombre,  " & _
                                          "        F.d_abrev    " & _
                                          " FROM etxcoi D,      " & _
                                          "      etaps F        " & _
                                          " where D.n_etapa = F.n_etapa and            " & _
                                          "       D.c_referencia = '"& arrRefEtapas(0,y) & "' and " & _
                                          "       D.c_conte      = '"& arrRefEtapas(1,y) & "' and " & _
                                          "       F.d_abrev      = 'SPL' " & _
                                          " order by D.n_secuenc desc "

                             'Response.Write(strSqlSel)
                             'Response.End
                             RConteDetalle.Source = strSqlSel
                             RConteDetalle.CursorType = 0
                             RConteDetalle.CursorLocation = 2
                             RConteDetalle.LockType = 1
                             RConteDetalle.Open()
                             if not RConteDetalle.eof then
                                 strCadenaColumna = strCadenaColumna & RConteDetalle.Fields.Item("d_nombre").Value&" ("&RConteDetalle.Fields.Item("d_abrev").Value&") "&RConteDetalle.Fields.Item("f_fecha").Value&".-"&mid(RConteDetalle.Fields.Item("m_observ").Value,1,135) & " <br>"
                             end if
                             RConteDetalle.close
                             set RConteDetalle = Nothing

                       '***********************************************************************

                      if strCadenaColumna <> "" then
                       strCadenaRenglon = strCadenaRenglon & strCadenaColumna
                      end if
                      strCadenaRenglon = strCadenaRenglon & " </font> </td>"
                      strCadenaRenglon = strCadenaRenglon & "</tr>"


                      if strCadenaColumna <> "" then
                        strHTML = strHTML &  strCadenaRenglon
                      end if

                '    end if
                  end if

                next
            end if

                'For i = 0 To Ubound(arrRefEtapas)
	              '   response.write("índice(" & i & "):" & arrRefEtapas(i) & "<br>")
                'Next


            '***                                                                                    ***
            '******************************************************************************************
            strHTML = strHTML & "</table>"& chr(13) & chr(10)

            end if

            RsRep.close
            Set RsRep = Nothing
            'Se pinta todo el HTML formado
            response.Write(strHTML)
            if strHTML = "" then
               strHTML = "NO EXISTEN REGISTROS"
               response.Write(strHTML)
            end if

          else
             strHTML = "NO EXISTEN REGISTROS"
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
<%end if%>
