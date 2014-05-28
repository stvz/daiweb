
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% Server.ScriptTimeout=1500 %>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<HTML>
<HEAD>
<TITLE>:: REPORTE DE PESOS .... ::</TITLE>
</HEAD>
<BODY>
<%
	'Response.Write("HOLA MUNDO")
	'response.end()
	IPHost = "localhost"
    strTipoUsuario  = request.Form("TipoUser")
    strPermisos     = Request.Form("Permisos")
    strTipoTransp   = Request.Form("txttipoTransp")

    strDateIni      = Request.Form("txtDateIni")
    strHoraIni      = Request.Form("txtHoraIni")
    strDateFin      = Request.Form("txtDateFin")
    strHoraFin      = Request.Form("txtHoraFin")
    strFilTipoFecha = Request.Form("FilTipoFecha")




    'pedimento=request.form("txtPed")
    strReferencia=request.form("txtRef") ' se cambio por numero de pedimento

    'if strReferencia<>"" then
    '   strReferencia = " and a.refe01 ='"&strReferencia&"' "
    'else
    '   strReferencia = " "
    'end if


    if strReferencia<>"" then
       strReferencia = " and a.pedi01 ='"&strReferencia&"' "
    else
       strReferencia = " "
    end if


    if strFilTipoFecha = "1" then    'Fecha del Ticket

       strFilTipoFecha2 = " and CAST(FECHTICK AS DATETIME)                                                          " &_
                          " BETWEEN CAST( '"&FormatoFechaInvGuion(strDateIni)&" "& strHoraIni & ":00' AS DATETIME)  " &_
                          " AND CAST( '"&FormatoFechaInvGuion(strDateFin)&" "& strHoraFin & ":00' AS DATETIME)      "

       strFilTipoFecha3 = " and CAST(c.FECHTICK AS DATETIME)                                                        " &_
                          " BETWEEN CAST( '"&FormatoFechaInvGuion(strDateIni)&" "& strHoraIni & ":00' AS DATETIME)  " &_
                          " AND CAST( '"&FormatoFechaInvGuion(strDateFin)&" "& strHoraFin & ":00' AS DATETIME)      "

       strFilTipoFecha4 = " and CAST(FECHTICK AS DATETIME ) <= CAST( '"&FormatoFechaInvGuion(strDateFin)&" "& strHoraFin & ":00' AS DATETIME) "

    else ' Fecha de Liberación
       strFilTipoFecha2 = " and CAST(FECHDOC AS DATETIME)                                                           " &_
                          " BETWEEN CAST( '"&FormatoFechaInvGuion(strDateIni)&" "& strHoraIni & ":00' AS DATETIME)  " &_
                          " AND CAST( '"&FormatoFechaInvGuion(strDateFin)&" "& strHoraFin & ":00' AS DATETIME)      "

       strFilTipoFecha3 = " and CAST(c.FECHDOC AS DATETIME)                                                        " &_
                          " BETWEEN CAST( '"&FormatoFechaInvGuion(strDateIni)&" "& strHoraIni & ":00' AS DATETIME)  " &_
                          " AND CAST( '"&FormatoFechaInvGuion(strDateFin)&" "& strHoraFin & ":00' AS DATETIME)      "

       strFilTipoFecha4 = " and CAST(FECHDOC AS DATETIME ) <= CAST( '"&FormatoFechaInvGuion(strDateFin)&" "& strHoraFin & ":00' AS DATETIME) "

    end if






    if request.form("rbnTipoDate") = "2" then
      Response.Addheader "Content-Disposition", "attachment;filename=Rep_pesos.xls"
      Response.ContentType = "application/vnd.ms-excel"
    end if

    permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

    'permi          = PermisoClientes(Session("GAduana"),strPermisos,"cliE01")
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
       permi = " AND a.cvecli01 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
       permi = ""
    end if
    %>

   <%''''''''''''''''''''''''if  Session("GUsuario") <> "" then

   if  Session("GAduana") <> "" then
	'Response.Write(strTipoTransp & "trans")
	'response.end()
      if  strTipoTransp = "1" then

      StrfilTipVeh = " AND TIPVEH IN ('H','K','C','T','l','J','F','G','P','D','R') "

      ' " AND TIPVEH IN ('F','G') " FERROCARRILES
		'Response.write(strfiltipveh)
		'response.end()

              %>


               <table width="778"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td align="center">
                        <%
                        dim cve,mov,fi,ff,tabla,sql1,sql2,sql3,refe

                    oficina_adu=GAduana

                     jnxadu=Session("GAduana")

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
                         end select


                MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
                MM_EXTRANET_STRING2 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="& IPHost &"; DATABASE=rku_cpsimples; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
                Set Conn = Server.CreateObject ("ADODB.Connection")
                Set REFE = Server.CreateObject ("ADODB.RecordSet")
                Conn.Open MM_EXTRANET_STRING

                Set Conn = Server.CreateObject ("ADODB.Connection")
                Set REFE = Server.CreateObject ("ADODB.RecordSet")
                Conn.Open MM_EXTRANET_STRING2

                'STRSQL= " SELECT distinct a.refe01 as ref,     " &_
                '        "        a.pedi01 as ped,              " &_
                '        "        a.nomcli01 as clie,           " &_
                '        "        a.patent01 as pat,            " &_
                '        "        a.fecpag01,                   " &_
                '        "        b.frac01 as frac,             " &_
                '        "        b.desc01 as merc,             " &_
                '        "        a.fecpag01 as fecpag,         " &_
                '        "        a.pesobr01 as pesob,          " &_
                '        "        b.partida01 as partida,       " &_
                '        "        a.cvecli01,                   " &_
                '        "        a.nombar01                    " &_
                '        " FROM pedimentos a,                   " &_
                '        "      fracciones b,                   " &_
                '        "      tcepartidas c                   " &_
                '        " where a.refe01=b.refe01 "&permi&"    " &_
                '        "   and b.refe01=c.refe01 " & strReferencia &_
                '        "   AND ( c.FECTICK >= '"&FormatoFechaInv(strDateIni)&"' AND      " & _
                '        "         c.HORAPES >= '"&strHoraIni&"'                     ) AND " & _
                '        "       ( c.FECTICK <= '"&FormatoFechaInv(strDateFin)&"' AND      " & _
                '        "         c.HORAPES <= '"&strHoraFin&"'                     )     " & _
                '        "   and consec01 > 0            " &_
                '        "   and partComp = 0            " & StrfilTipVeh & _
                '        "   and pesoaudita > 0  "
                '
                STRSQL= " SELECT distinct a.refe01 as ref,     " &_
                        "        a.pedi01 as ped,              " &_
                        "        a.nomcli01 as clie,           " &_
                        "        a.patent01 as pat,            " &_
                        "        a.fecpag01,                   " &_
                        "        b.frac01 as frac,             " &_
                        "        b.desc01 as merc,             " &_
                        "        a.fecpag01 as fecpag,         " &_
						"        a.fecent01 as fecent,         " &_
						"        a.pesobr01 as pesob,          " &_
                        "        b.partida01 as partida,       " &_
                        "        a.cvecli01,                   " &_
                        "        a.nombar01,                   " &_
                        "        a.rmercan,                    " &_
                        "        a.viajeBar,                   " &_
						"		 if(a.status=2,'Descargando','Cerrado') as statusped,	   " &_
                        "        a.Rnombar                     " &_
                        " FROM pedimentos a,                   " &_
                        "      fracciones b,                   " &_
                        "      tcepartidas c                   " &_
                        " where a.refe01=b.refe01 "&permi&"    " &_
                        "   and b.refe01=c.refe01 " & strReferencia  & strFilTipoFecha3 &_
                        "   and consec01 > 0            " &_
                        "   and partComp = 0            " & StrfilTipVeh & _
                        "   and pesoaudita > 0  "
						
						' response.write(STRSQL) 
						' response.end




                        strTipoTransp =  Request.Form("txttipoTransp")
                        strDateIni    =  Request.Form("txtDateIni")
                        strHoraIni    =  Request.Form("txtHoraIni")
                        strDateFin    =  Request.Form("txtDateFin")
                        strHoraFin    =  Request.Form("txtHoraFin")

                'response.Write(strsql)
                'response.end()

                Set REFE= Conn.Execute(strSQL)
				
				QRYCIERRE = "SELECT DATE_ADD(fectick, INTERVAL 60 DAY) as feccierre, consec01 from tcepartidas where consec01 = 1 and pedimento = '" & Trim(request.form("txtRef")) &"'"
				'response.Write(qrycierre)
                'response.end()
				Set CIERRA = Server.CreateObject ("ADODB.RecordSet")
				SET CIERRA= Conn.Execute(QRYCIERRE)

                %>


                <table width="854px" height="50px" bgcolor = "#003366" align= "Center" valign="midle" >
                <tr>
                <td width="98%" height="50px" align= "Center" valign="midle" colspan=9>
                   <strong>
                   <font color="#C0C0C0" size="4" face="Arial, Helvetica, sans-serif">
                   <p >Reporte de Pesos de Camiones </p>
                   </font>
                   </strong>
                </td>

                <tr>
                </table>



               <table align="left" >

                 <%

                     if not REFE.eof then
                     While (NOT  REFE.EOF)

                      referencia = REFE("ref")
					  QRYCIERRE = "SELECT DATE_ADD(fectick, INTERVAL 59 DAY) as feccierre, consec01 from tcepartidas where consec01 = 1 and refe01 = '" & referencia & "'"
					  Set CIERRA = Server.CreateObject ("ADODB.RecordSet")
					  SET CIERRA= Conn.Execute(QRYCIERRE)
                      pedimento  = REFE("ped")
                      cliente    = REFE("clie")
                      patente    = REFE("pat")
                      mercancia  = REFE("merc")
                      fecpago    = REFE("fecpag")
					  fecent	 = REFE("fecent")
					  feccierre	 = CIERRA("feccierre")
					  statusped  = REFE("statusped")
                      peso       = REFE("pesob")
                      vence      = fecpago+60
                      frac       = REFE("frac")
                      saldolib   = 0
                      tmpeso     = trim(cstr(formatnumber(peso,0)))
                      buque      = REFE("nombar01")
					  

                      Strrmercan   = REFE("rmercan")
                      StrviajeBar  = REFE("viajeBar")
                      StrRnombar   = REFE("Rnombar")






                        '************************************partidas tce**********************
                        Set Connx = Server.CreateObject ("ADODB.Connection")
                        Set RSPart = Server.CreateObject ("ADODB.RecordSet")
                        intcontadorunidades = 0
                        Connx.Open MM_EXTRANET_STRING2
                        'strSQL = " select *                            " & _
                        '         " from  tcepartidas                   " & _
                        '         " where frac01 = '"&frac&"'       AND " & _
                        '         "       refe01 = '"&referencia&"' AND " & _
                        '         "       ( FECTICK >= '"&FormatoFechaInv(strDateIni)&"' AND            " & _
                        '         "         HORAPES >= '"&strHoraIni&"'                     ) AND       " & _
                        '         "       ( FECTICK <= '"&FormatoFechaInv(strDateFin)&"' AND            " & _
                        '         "         HORAPES <= '"&strHoraFin&"'                     )      " & _
                        '         "   and consec01 > 0            " &_
                        '         "   and partComp = 0            "

                        strSQL = " select *                            " & _
                                 " from  tcepartidas                   " & _
                                 " where frac01 = '"&frac&"'       AND " & _
                                 "       refe01 = '"&referencia&"'     " & _
                                 "   and consec01 > 0            " & _
                                 "   and partComp = 0            " & _
                                 "   and pesoaudita > 0  "
                        'Response.Write(strSQL)
                        Set RSPart= Connx.Execute(strSQL)
                        if not RSPart.eof then
                            Do while not RSPart.Eof
                                pesoneto = RSPart("pesoneto")
                                tipoveh  = RSPart("tipveh")
                                saldolib    = saldolib + round(pesoneto)
                                saldo       = peso-saldolib
                                intcontadorunidades = intcontadorunidades + 1
                                RSPart.MoveNext  ' de las partidas de tce
                            Loop ' de las partidas de tce
                         else
                           saldo               = peso - saldolib
                           intcontadorunidades = 0
                         end if
                         '------------------------------------para la segunda tablas



                         '************************************partidas tce**********************

                        Set Connx2  = Server.CreateObject ("ADODB.Connection")
                        Set RSPart2 = Server.CreateObject ("ADODB.RecordSet")
                        intcontadorunidades = 0
                        Connx2.Open MM_EXTRANET_STRING2
                        'strSQL = " select COUNT(REFE01) AS PARTIDAS                            " & _
                        '         " from  tcepartidas                   " & _
                        '         " where frac01 = '"&frac&"'       AND " & _
                        '         "       refe01 = '"&referencia&"' AND " & _
                        '         "       ( FECTICK >= '"&FormatoFechaInv(strDateIni)&"' AND            " & _
                        '         "         HORAPES >= '"&strHoraIni&"'                     ) AND       " & _
                        '         "       ( FECTICK <= '"&FormatoFechaInv(strDateFin)&"' AND            " & _
                        '         "         HORAPES <= '"&strHoraFin&"'                     )      " & _
                        '         "   and consec01 > 0            " &_
                        '         "   and partComp = 0            " & StrfilTipVeh &_
                        '         "   and pesoaudita > 0  " &_
                        '         "   GROUP BY REFE01     "



                        strSQL = " select COUNT(REFE01) AS PARTIDAS , SUM(pesoneto) as pesoalcorte                       " & _
                                 " from  tcepartidas                                                                     " & _
                                 " where frac01 = '"&frac&"'       AND                                                   " & _
                                 "       refe01 = '"&referencia&"'                                                       " & strFilTipoFecha2 &_
                                 "   and consec01 > 0                                                                    " &_
                                 "   and partComp = 0            " & StrfilTipVeh &_
                                 "   and pesoaudita > 0  " &_
                                 "   GROUP BY REFE01     "

                        'Response.Write(strSQL)
                        'Response.End
                        dblPesoAlCorte = 0
                        dblSaldoAlcorte = 0
                        set RSPart2 = Connx2.Execute(strSQL)
                        if not RSPart2.eof then
                            Do while not RSPart2.Eof
                                intcontadorunidades = RSPart2("PARTIDAS")
                                dblPesoAlCorte      = RSPart2("pesoalcorte")
                                dblSaldoAlcorte     = peso - dblPesoAlCorte
                                RSPart2.MoveNext
                            Loop ' de las partidas de tce
                         else
                           intcontadorunidades = 0
                         end if
                         '------------------------------------para la segunda tablas



                         '--------------------------------------------------------------------
                            Set Connx3  = Server.CreateObject ("ADODB.Connection")
                            Set RSPart3 = Server.CreateObject ("ADODB.RecordSet")
                            intcontadorunidades3 = 0
                            Connx3.Open MM_EXTRANET_STRING2
                            strSQL3 = " select COUNT(REFE01) AS PARTIDAS , SUM(pesoneto) as pesoalcorte                      " & _
                                     " from  tcepartidas                                                                     " & _
                                     " where frac01 = '"&frac&"'       AND                                                   " & _
                                     "       refe01 = '"&referencia&"'                                                       " & strFilTipoFecha4 &_
                                     "   and consec01 > 0                                                                    " &_
                                     "   and partComp = 0            " & StrfilTipVeh &_
                                     "   and pesoaudita > 0  " &_
                                     "   GROUP BY REFE01     "
                            'Response.Write(strSQL3)
                            'Response.End
                            dblPesoAlCorte3 = 0
                            dblSaldoAlcorte3 = 0
                            set RSPart3 = Connx3.Execute(strSQL3)
                            if not RSPart3.eof then
                                Do while not RSPart3.Eof
                                    intcontadorunidades3 = RSPart3("PARTIDAS")
                                    dblPesoAlCorte3      = RSPart3("pesoalcorte")
                                    dblSaldoAlcorte3     = peso - dblPesoAlCorte3
                                    RSPart3.MoveNext
                                Loop ' de las partidas de tce
                             else
                               intcontadorunidades3 = 0
                             end if
                             '------------------------------------para la segunda tablas

                         '--------------------------------------------------------------------





                         if(peso > 0) then
                            peso = peso/1000
                         end if

                         if(saldolib > 0) then
                            saldolib = saldolib/1000
                         end if

                         if(saldo > 0) then
                            saldo = saldo/1000
                         end if

                         if(dblPesoAlCorte > 0) then
                            dblPesoAlCorte = dblPesoAlCorte/1000
                         end if

                         if(dblSaldoAlcorte > 0) then
                            dblSaldoAlcorte = dblSaldoAlcorte/1000
                         end if

                         if(dblPesoAlCorte3 > 0) then
                            dblPesoAlCorte3 = dblPesoAlCorte3/1000
                         end if

                         if(dblSaldoAlcorte3 > 0) then
                            dblSaldoAlcorte3 = dblSaldoAlcorte3/1000
                         end if

                         'dblPesoAlCorte3
                         'dblSaldoAlcorte3


                    %>

                   <BR>
                     <table width="854"  border="1" cellspacing="3" cellpadding="3">
                       <tr bgcolor="#C0C0C0">
                           <td bgcolor="#C0C0C0" >
                             <font size="1" color="#993300" >
                                <b>Buque:</b>
                             </FONT>
                           </td>
                           <td  align="left" colspan="2" >
                             <font size="1" color="#993300">
                                <!-- <b> RESPONSE.Write(buque) </b> -->
                                <b> <%RESPONSE.Write(StrRnombar)%> </b>

                             </font>
                           </td>
                           <td bgcolor="#C0C0C0" >
                            <font size="1" color="#993300" >
                              <b>Producto</b>
                            </FONT>
                           </td>

                           <td  align="left" colspan="3">
                            <font size="1" color="#993300" >
                                <!-- <b> RESPONSE.Write( mercancia ) </b> -->
                                <b> <%RESPONSE.Write( Strrmercan )%> </b>

                            </font>
                           </td>

                           <td bgcolor="#C0C0C0" colspan="2">
                            <font size="1" color="#993300" >
                                <b>TM PEDIMENTO</b>
                            </FONT>
                           </td>

                           <td  align="left">
                            <font size="1" color="#993300">
                              <b>

                                  <%RESPONSE.Write( trim(cstr(formatnumber(peso,3))) ) %>
                              </b>
                            </font>
                           </td>

                           <td bgcolor="#C0C0C0" colspan="2" >
                            <font size="1" color="#993300" >
                                <b>TM TOTAL LIBERADO </b>
                            </FONT>
                           </td>

                           <td  align="left">
                            <font size="1" color="#993300">
                              <b>
                                  <%RESPONSE.Write( trim(cstr(formatnumber(saldolib,3))) ) %>
                              </b>
                            </font>
                           </td>

                           <td bgcolor="#C0C0C0" colspan="2" >
                            <font size="1" color="#993300" >
                                <b>TM SALDO TOTAL POR LIBERAR </b>
                            </FONT>
                           </td>

                           <td  align="left">
                            <font size="1" color="#993300">
                              <b>
                                  <%RESPONSE.Write( trim(cstr(formatnumber(saldo,3))) ) %>
                              </b>
                            </font>
                           </td>

                       </tr>
                       <tr bgcolor="#C0C0C0" >
                           <td  >
                             <font size="1" color="#993300" >
                               <b>Pedimento:</b>
                             </FONT>
                           </td>
                           <td  align="left" colspan="2">
                             <font size="1" color="#993300">
                               <b> <% RESPONSE.Write( trim(patente)&"-"&trim(pedimento) )%> </b>
                             </font>
                           </td>
                           <td >
                             <font size="1" color="#993300" >
                                <b>Cliente</b>
                             </FONT>
                           </td>
                           <td  align="left" colspan="3">
                             <font size="1" color="#993300" >
                                <b><%RESPONSE.Write(cliente)%> </b>
                             </font>
                           </td>
                           <td  colspan="2">
                             <font size="1" color="#993300" >
                                <b>UNIDADES ACOMULADAS AL CORTE:</b>
                             </FONT>
                           </td>
                           <td  align="left">
                             <font size="1" color="#993300" >
                                <b><%RESPONSE.Write(intcontadorunidades3)%> </b>
                             </font>
                           </td>

                           <td bgcolor="#C0C0C0" colspan="2">
                            <font size="1" color="#993300" >
                                <b> TM LIBERADO ACOMULADO AL CORTE</b>
                            </FONT>
                           </td>

                           <td  align="left">
                            <font size="1" color="#993300">
                              <b>
                                  <%RESPONSE.Write( trim(cstr(formatnumber(dblPesoAlCorte3,3))) ) %>
                              </b>
                            </font>
                           </td>

                           <td bgcolor="#C0C0C0" colspan="2">
                            <font size="1" color="#993300" >
                                <b> TM POR LIBERAR AL CORTE</b>
                            </FONT>
                           </td>

                           <td  align="left">
                            <font size="1" color="#993300">
                              <b>
                                  <%RESPONSE.Write( trim(cstr(formatnumber(dblSaldoAlcorte3,3))) ) %>
                              </b>
                            </font>
                           </td>

                       </tr>
                       <tr >

                           <td colspan="3">
							
                           </td>
						   <td bgcolor="#C0C0C0" bgcolor="#C0C0C0" colspan="1">
                            <font size="1" color="#993300" >
                                <b>STATUS PEDIMENTO</b>
                            </FONT>
                           </td>
						   <td  align="left" colspan="3">
                            <font size="1" color="#993300">
                              <b>
                                  <%RESPONSE.Write(statusped) %>
                              </b>
                            </font>
                           </td>
                           <td bgcolor="#C0C0C0" bgcolor="#C0C0C0" colspan="2">
                            <font size="1" color="#993300" >
                                <b>UNIDADES AL CORTE</b>
                            </FONT>
                           </td>

                           <td  align="left">
                            <font size="1" color="#993300">
                              <b>

                                  <%RESPONSE.Write( trim(cstr(formatnumber(intcontadorunidades,0))) ) %>
                              </b>
                            </font>
                           </td>

                           <td bgcolor="#C0C0C0" colspan="2" >
                            <font size="1" color="#993300" >
                                <b>TM LIBERADO AL CORTE</b>
                            </FONT>
                           </td>

                           <td  align="left">
                            <font size="1" color="#993300">
                              <b>
                                  <%RESPONSE.Write( trim(cstr(formatnumber(dblPesoAlCorte,3))) ) %>
                              </b>
                            </font>
                           </td>

                           <td bgcolor="#C0C0C0" colspan="2" >
                            <font size="1" color="#993300" >
                                <b>VENCIMIENTO REGLA</b>
                            </FONT>
                           </td>

                           <td bgcolor="#C0C0C0" align="left">
                            <font size="1" color="#993300">
                              <b>
								<% Response.Write(cstr(feccierre)) %>
                              </b>
                            </font>
                           </td>

                       </tr>


                     </table>
                   <BR>





                      <%	   '*******************************************************

                      '--------------------------------------------------------------------------------- 
					  %>
                        <table width="854" border="1" cellspacing="3" cellpadding="3">
                              <!--
                              <tr>
                                 <th bgcolor="#009999">&nbsp;</th>
                                 <th bgcolor="#009999">&nbsp;</th>
                                 <th bgcolor="#009999">&nbsp;</th>
                                 <th bgcolor="#009999" colspan="2"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Transporte</FONT></th>
                                 <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Inicios</FONT></th>
                                 <th bgcolor="#009999">&nbsp;</th>
                              </tr>
                              -->
                              <tr bgcolor="#C0C0C0" >
                                 <!--
                                 <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Fecha</FONT></th>
                                 <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Partida</FONT></th>
                                 <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">No.ticket</FONT></th>
                                 <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Tipo</FONT></th>
                                 <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Placas</FONT></th>
                                 <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">FL</FONT></th>
                                 <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Peso Neto</FONT></th>
                                 -->

                                 <td width="50px" ><font size="1" color="#000000" > <b> CONSECUTIVO         </b> </FONT></th>
                                 <td width="50px" ><font size="1" color="#000000" > <b> No. LOAD            </b> </FONT></th>
								 <td width="50px" ><font size="1" color="#000000" > <b> RECINTO         	</b> </FONT></th>
                                 <td width="50px" ><font size="1" color="#000000" > <b> No. FICHA           </b> </FONT></th> <!-- Ticket -->
                                 <td width="50px" ><font size="1" color="#000000" > <b> FECHA PESAJE        </b> </FONT></th> <!-- Fecha Ticket -->
                                 <td width="50px" ><font size="1" color="#000000" > <b> HORA PESAJE         </b> </FONT></th> <!-- Hora -->

                                 <td width="50px" ><font size="1" color="#000000" > <b> FECHA LIBERACIÓN    </b> </FONT></th> <!-- Fecha Ticket -->
                                 <td width="50px" ><font size="1" color="#000000" > <b> HORA LIBERACIÓN     </b> </FONT></th> <!-- Hora -->

                                 <td width="50px" ><font size="1" color="#000000" > <b> TIPO DE VEHICULO    </b> </FONT></th> <!-- Tipo de Vehiculo -->
								 <td width="50px" ><font size="1" color="#000000" > <b> TRANSPORTISTA    </b> </FONT></th> <!-- Tipo de Vehiculo -->
								 <td width="50px" ><font size="1" color="#000000" > <b> OPERADOR             </b> </FONT></th>
                                 <td width="50px" ><font size="1" color="#000000" > <b> PLACA TRACTOR       </b> </FONT></th>
								 <td width="50px" ><font size="1" color="#000000" > <b> SELLOS       </b> </FONT></th>
								 <td width="50px" ><font size="1" color="#000000" > <b> SELLOS ESLINGA       </b> </FONT></th>
                                 <td width="50px" ><font size="1" color="#000000" > <b> PLACA JAULA         </b> </FONT></th>
                                 <td width="50px" ><font size="1" color="#000000" > <b> P.TARA(KG)          </b> </FONT></th>
                                 <td width="50px" ><font size="1" color="#000000" > <b> P.BRUTO(KG)         </b> </FONT></th>
                                 <td width="50px" ><font size="1" color="#000000" > <b> P.NETO(KG)          </b> </FONT></th>
                                 <td width="50px" ><font size="1" color="#000000" > <b> DESTINO             </b> </FONT></th>
								 

                              </tr>

                      <% '************************************partidas tce**********************

                      saldolib = 0
                      Set Connx = Server.CreateObject ("ADODB.Connection")
                      Set RSPart = Server.CreateObject ("ADODB.RecordSet")
                      Connx.Open MM_EXTRANET_STRING2
                      'strSQL= " SELECT a.refe01 as ref,        " &_
                      '        "        a.pedi01 as ped,        " &_
                      '        "        a.nomcli01 as clie,     " &_
                      '        "        a.patent01 as pat,      " &_
                      '        "        a.fecpag01,             " &_
                      '        "        b.frac01 as frac,       " &_
                      '        "        b.desc01 as merc,       " &_
                      '        "        a.fecpag01 as fecpag,   " &_
                      '        "        a.pesobr01 as pesob,    " &_
                      '        "        c.consec01 as consec,   " &_
                      '        "        c.refe01,               " &_
                      '        "        c.tipveh as tipveh,     " &_
                      '        "        c.nload,                " &_
                      '        "        c.ticket ,              " &_
                      '        "        c.fectick,              " &_
                      '        "        c.placas ,              " &_
                      '        "        c.placaJau,             " &_
                      '        "        c.pesoBruto,            " &_
                      '        "        c.pesoTara,             " &_
                      '        "        c.pesoNeto,             " &_
                      '        "        TIME_FORMAT(c.horaPes,'%H:%i') as horaPes, " &_
                      '        "        c.destino               " &_
                      '        " FROM pedimentos a,fracciones b,tcepartidas c " &_
                      '        " where a.refe01='"&referencia&"'  " &_
                      '        "   and a.refe01=b.refe01         " &_
                      '        "   and a.refe01=c.refe01         " &_
                      '        "   AND ( c.FECTICK >= '"&FormatoFechaInv(strDateIni)&"' AND      " & _
                      '        "         c.HORAPES >= '"&strHoraIni&"'                     ) AND " & _
                      '        "       ( c.FECTICK <= '"&FormatoFechaInv(strDateFin)&"' AND      " & _
                      '        "         c.HORAPES <= '"&strHoraFin&"'                     )      " &permi &_
                      '        "   and c.consec01 > 0            " &_
                      '        "   and c.partComp = 0            " & StrfilTipVeh & _
                      '        "   and pesoaudita > 0  " & _
                      '        "   order by consec01  "
                      '
                      '
                      '        "    CAST(FECHTICK AS DATETIME)                                                         " &_
                      '        "    BETWEEN CAST( '"&FormatoFechaInv(strDateIni)&" "& strHoraIni & ":00' AS DATETIME)  " &_
                      '        "        AND CAST( '"&FormatoFechaInv(strDateFin)&" "& strHoraFin & ":00' AS DATETIME)  " &_

                      strSQL= " SELECT a.refe01 as ref,        " &_
                              "        a.pedi01 as ped,        " &_
                              "        a.nomcli01 as clie,     " &_
                              "        a.patent01 as pat,      " &_
                              "        a.fecpag01,             " &_
                              "        b.frac01 as frac,       " &_
                              "        b.desc01 as merc,       " &_
                              "        a.fecpag01 as fecpag,   " &_
                              "        a.pesobr01 as pesob,    " &_
                              "        c.consec01 as consec,   " &_
                              "        c.refe01,               " &_
                              "        c.tipveh as tipveh,     " &_
                              "        c.nload,                " &_
                              "        c.ticket ,              " &_
                              "        c.fectick,              " &_
                              "        c.placas ,              " &_
							  "		   c.transpor, 			   " &_
                              "        c.placaJau,             " &_
                              "        c.pesoBruto,            " &_
                              "        c.pesoTara,             " &_
                              "        c.pesoNeto,             " &_
							  "        c.operador,             " &_
                              "        TIME_FORMAT(c.horaPes,'%H:%i') as horaPes, " &_
                              "        c.destino,              " &_
                              "        c.FecDoc,               " &_
							  "        c.Recinto,               " &_
							  "        c.sellos as se,         " &_
							  "        c.sellosE as see,       " &_
                              "        TIME_FORMAT(c.horaDoc,'%H:%i') as horaDoc  " &_
                              " FROM pedimentos a,fracciones b,tcepartidas c " &_
                              " where a.refe01='"&referencia&"'  " &_
                              "   and a.refe01=b.refe01         " &_
                              "   and a.refe01=c.refe01         " & strFilTipoFecha3 &permi &_
                              "   and c.consec01 > 0            " &_
                              "   and c.partComp = 0            " & StrfilTipVeh & _
                              "   and pesoaudita > 0  " & _
                              "   order by consec01  "


                      ' Response.Write(strSQL)
                      ' Response.End
                      Set RSPart= Connx.Execute(strSQL)

                      Do while not RSPart.Eof

                          'fechapartida=RSPart("fectick")
                          'partida=RSPart("consec")
                          'ticket=RSPart("ticket")

                          'placas=RSPart("placas")
                          'inicios=RSPart("inicio")

                          pesoneto = RSPart("pesoneto")


                          strnload     = RSPart("nload")
                          strticket    = RSPart("ticket")
                          strfectick   = RSPart("fectick")
                          strplacas    = RSPart("placas")
						  strtranspor  = RSPart("transpor")
                          strplacaJau  = RSPart("placaJau")
                          strpesoBruto = RSPart("pesoBruto")
                          strpesoTara  = RSPart("pesoTara")
                          strpesoNeto  = RSPart("pesoNeto")
                          strdestino   = RSPart("destino")
                          strhoraPes   = RSPart("horaPes")
                          tipoveh      = RSPart("tipveh")
                          strConsec    = RSPart("consec")
						  stroperador  = RSPart("operador")
						  strrecinto   = RSPart("recinto")
                          strhoraDoc   = RSPart("horaDoc")
                          strfecDOc    = RSPart("FecDoc")
						  strsellos    = RSPart("se")
						  strsellosE   = RSPart("see")


                          saldolib=saldolib+round(pesoneto)
                          select case tipoveh
								case "C"
									Tipotrans="Camión"
								case "T"
                                    Tipotrans="Tolva"
								case "F"
                                    Tipotrans="Furgon"
								case "G"
                                    Tipotrans="Gondola"
								case "l"
									Tipotrans="Full"
								case "J"
                                    Tipotrans="Jaula"
								case "K"
									Tipotrans="Camioneta"
								case "H"
									Tipotrans="Torthon"
								case "P"
									Tipotrans="Pipa"
								case "D"
									Tipotrans="Doble Pipa"
								case "R"
									Tipotrans="Carro Tanque"
                          end select
                          '---------------------------TERCER TABLA PARTIDAS------------------
						  %>
                           <tr>

                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strConsec)  %>  </font></td><!-- CONSECUTIVO   -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strnload)   %>  </font></td><!-- No. LOAD      -->
							 <td><font size="1" color="#000000" > <%RESPONSE.Write(strrecinto)   %>  </font></td><!-- RECINTO      -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strticket)  %>  </font></td><!-- No. FICHA     -->

                             <%
                                  'if strFilTipoFecha = "1" then    'Fecha del Ticket
                                  '<td><font size="1" color="#000000" > <RESPONSE.Write(strfectick) >  </font></td><!-- FECHA         -->
                                  '<td><font size="1" color="#000000" > <RESPONSE.Write(strhoraPes) >  </font></td><!-- HoraPes       -->
                                  'else ' Fecha de Liberación
                                  'end if
                             %>


                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strfectick) %>  </font></td><!-- FECHA    -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strhoraPes) %>  </font></td><!-- HoraPes  -->

                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strfecDOc) %>   </font></td><!-- FECHA LIBERACIÓN -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strhoraDoc) %>  </font></td><!-- HoraPes LIBERACIÓN -->

                             <td align="Center"><font size="1" color="#000000" > <%RESPONSE.Write(Tipotrans)  %>  </FONT></td><!-- Tipo de Vehiculo -->
							 <td ><font size="1" color="#000000" > <%RESPONSE.Write(strtranspor)%>   </font></td><!-- CONDUCTOR TRANSPORTE          -->
							 <td ><font size="1" color="#000000" > <%RESPONSE.Write(stroperador)%>   </font></td><!-- CONDUCTOR TRANSPORTE          -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strplacas)  %>  </font></td><!-- PLACA TRACTOR -->
							 <td><font size="1" color="#000000" > <%RESPONSE.Write(strsellos)  %>  </font></td><!-- SELLO EMBARQUE -->
							 <td><font size="1" color="#000000" > <%RESPONSE.Write(strsellosE)  %>  </font></td><!-- SELLO ESLINGA -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strplacaJau)%>  </font></td><!-- PLACA JAULA   -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write( trim(cstr(formatnumber(strpesoTara,0)))  )%> </font></td><!-- P. TARA       -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write( trim(cstr(formatnumber(strpesoBruto,0))) )%> </font></td><!-- P. BRUTO      -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write( trim(cstr(formatnumber(strpesoNeto,0)))  )%> </font></td><!-- P. NETO       -->
                             <td><font size="1" color="#000000" > <%RESPONSE.Write(strdestino) %>  </font></td><!-- DESTINO       -->
							 
                           </tr>

                          <%

                          RSPart.MoveNext  ' de las partidas de tce
                      Loop ' de las partidas de tce



                          %>
                   <tr>
                        <td colspan="12" align="right">
                          <font size="1" color="#000000" >
                            <b>Saldo Liberado ( KG ) </b>
                          </font>
                        </td>
                        <td>
                          <font size="1" color="#000000" >
                            <% RESPONSE.Write( trim(cstr(formatnumber(saldolib,0))) ) %>
                          </font>
                        </td>
                   </tr>
                   <% 
                      saldolib = saldolib/1000
                   %>
                   <tr>
                        <td colspan="12" align="right">
                          <font size="1" color="#000000" >
                            <b>Saldo Liberado ( TM ) </b>
                          </font>
                        </td>
                        <td>
                          <font size="1" color="#000000" >
                            <% RESPONSE.Write( trim(cstr(formatnumber(saldolib,3))) ) %>
                          </font>
                        </td>
                   </tr>
                </table>
              <%'**********************************************************************

                    ' RESPONSE.Write("</tr>")
                      ' Refe.MoveNext 'avanza referencia  ---->
                  Refe.MoveNext
                       wend 'REFErencia
               else
			  Response.Write(strSQL & "3")
			  response.end()
              %>
			 
              <tr>
                <th colspan=12>
                  <font size="2" face="Arial">No se Encontro ningun registro con esos parametros</font>
                </th>
              </tr>
              <table>

                <%
                  'end if
                  end if
                %>
              </form>





<%
else
    if  strTipoTransp = "2" then  ' Ferrocarriles

                StrfilTipVeh = " AND TIPVEH IN ('F','G','T','R') "
%>


                           <table width="778"  border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td align="center">
                                    <%
                                    'dim cve,mov,fi,ff,tabla,sql1,sql2,sql3,refe

                                oficina_adu=GAduana

                                 jnxadu=Session("GAduana")

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
                                     end select


                            MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
                            MM_EXTRANET_STRING2 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="& IPHost &"; DATABASE=rku_cpsimples; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
                            Set Conn = Server.CreateObject ("ADODB.Connection")
                            Set REFE = Server.CreateObject ("ADODB.RecordSet")
                            Conn.Open MM_EXTRANET_STRING

                            Set Conn = Server.CreateObject ("ADODB.Connection")
                            Set REFE = Server.CreateObject ("ADODB.RecordSet")
                            Conn.Open MM_EXTRANET_STRING2

                            'STRSQL= " SELECT distinct a.refe01 as ref,              " &_
                            '        "        a.pedi01 as ped,              " &_
                            '        "        a.nomcli01 as clie,           " &_
                            '        "        a.patent01 as pat,            " &_
                            '        "        a.fecpag01,                   " &_
                            '        "        b.frac01 as frac,             " &_
                            '        "        b.desc01 as merc,             " &_
                            '        "        a.fecpag01 as fecpag,         " &_
                            '        "        a.pesobr01 as pesob,          " &_
                            '        "        b.partida01 as partida,       " &_
                            '        "        a.cvecli01,                   " &_
                            '        "        a.nombar01                    " &_
                            '        " FROM pedimentos a,                   " &_
                            '        "      fracciones b,                   " &_
                            '        "      tcepartidas c                   " &_
                            '        " where a.refe01=b.refe01 "&permi&"    " &_
                            '        "   and b.refe01=c.refe01 " & strReferencia &_
                            '        "   AND ( c.FECTICK >= '"&FormatoFechaInv(strDateIni)&"' AND      " & _
                            '        "         c.HORAPES >= '"&strHoraIni&"'                     ) AND " & _
                            '        "       ( c.FECTICK <= '"&FormatoFechaInv(strDateFin)&"' AND      " & _
                            '        "         c.HORAPES <= '"&strHoraFin&"'                     )     " & _
                            '        "   and consec01 > 0            " &_
                            '        "   and partComp = 0            " & StrfilTipVeh & _
                            '        "   and pesoaudita > 0  "




                                    STRSQL= " SELECT distinct a.refe01 as ref,     " &_
                                            "        a.pedi01 as ped,              " &_
                                            "        a.nomcli01 as clie,           " &_
                                            "        a.patent01 as pat,            " &_
                                            "        a.fecpag01,                   " &_
                                            "        b.frac01 as frac,             " &_
                                            "        b.desc01 as merc,             " &_
                                            "        a.fecpag01 as fecpag,         " &_
                                            "        a.pesobr01 as pesob,          " &_
                                            "        b.partida01 as partida,       " &_
                                            "        a.cvecli01,                   " &_
                                            "        a.nombar01,                   " &_
                                            "        a.rmercan,                    " &_
                                            "        a.viajeBar,                   " &_
                                            "        a.Rnombar                     " &_
                                            " FROM pedimentos a,                   " &_
                                            "      fracciones b,                   " &_
                                            "      tcepartidas c                   " &_
                                            " where a.refe01=b.refe01 "&permi&"    " &_
                                            "   and b.refe01=c.refe01 " & strReferencia & strFilTipoFecha3 &_
                                            "   and consec01 > 0            " &_
                                            "   and partComp = 0            " & StrfilTipVeh & _
                                            "   and pesoaudita > 0  "


                                    strTipoTransp =  Request.Form("txttipoTransp")
                                    strDateIni    =  Request.Form("txtDateIni")
                                    strHoraIni    =  Request.Form("txtHoraIni")
                                    strDateFin    =  Request.Form("txtDateFin")
                                    strHoraFin    =  Request.Form("txtHoraFin")

                            'response.Write(strsql)
                            'response.end()
                            Set REFE= Conn.Execute(strSQL)


                            %>


                            <table width="854px" height="50px" bgcolor = "#000000" align= "Center" valign="midle" >
                            <tr>
                            <td width="98%" height="50px" align= "Center" valign="midle" colspan=9>
                               <strong>
                               <font color="#FFFFFF" size="4" face="Arial, Helvetica, sans-serif">
                               <p >Reporte de Pesos de FERROCARRILES </p>
                               </font>
                               </strong>
                            </td>

                            <tr>
                            </table>



                           <table align="left" >

                             <% 

                                 if not REFE.eof then
                                 While (NOT  REFE.EOF)

                                  referencia = REFE("ref")
                                  pedimento  = REFE("ped")
                                  cliente    = REFE("clie")
                                  patente    = REFE("pat")
                                  mercancia  = REFE("merc")
                                  fecpago    = REFE("fecpag")
                                  peso       = REFE("pesob")
                                  vence      = fecpago+60
                                  frac       = REFE("frac")
                                  saldolib   = 0
                                  tmpeso     = trim(cstr(formatnumber(peso,0)))
                                  buque      = REFE("nombar01")

                                  Strrmercan      = REFE("rmercan")
                                  StrviajeBar     = REFE("viajeBar")
                                  StrRnombar      = REFE("Rnombar")


                                    '************************************partidas tce**********************
                                    Set Connx = Server.CreateObject ("ADODB.Connection")
                                    Set RSPart = Server.CreateObject ("ADODB.RecordSet")
                                    intcontadorunidades = 0
                                    Connx.Open MM_EXTRANET_STRING2
                                    'strSQL = " select *                            " & _
                                    '         " from  tcepartidas                   " & _
                                    '         " where frac01 = '"&frac&"'       AND " & _
                                    '         "       refe01 = '"&referencia&"' AND " & _
                                    '         "       ( FECTICK >= '"&FormatoFechaInv(strDateIni)&"' AND            " & _
                                    '         "         HORAPES >= '"&strHoraIni&"'                     ) AND       " & _
                                    '         "       ( FECTICK <= '"&FormatoFechaInv(strDateFin)&"' AND            " & _
                                    '         "         HORAPES <= '"&strHoraFin&"'                     )      " & _
                                    '         "   and consec01 > 0            " &_
                                    '         "   and partComp = 0            "

                                    strSQL = " select *                            " & _
                                             " from  tcepartidas                   " & _
                                             " where frac01 = '"&frac&"'       AND " & _
                                             "       refe01 = '"&referencia&"'     " & _
                                             "   and consec01 > 0            " &_
                                             "   and partComp = 0            " & _
                                             "   and pesoaudita > 0  "
                                    'Response.Write(strSQL)
                                    Set RSPart= Connx.Execute(strSQL)
                                    if not RSPart.eof then
                                        Do while not RSPart.Eof
                                            pesoneto = RSPart("pesoneto")
                                            tipoveh  = RSPart("tipveh")
                                            saldolib    = saldolib + round(pesoneto)
                                            saldo       = peso-saldolib
                                            intcontadorunidades = intcontadorunidades + 1
                                            RSPart.MoveNext  ' de las partidas de tce
                                        Loop ' de las partidas de tce
                                     else
                                       saldo               = peso - saldolib
                                       intcontadorunidades = 0
                                     end if
                                     '------------------------------------para la segunda tablas



                                     '************************************partidas tce**********************

                                    Set Connx2  = Server.CreateObject ("ADODB.Connection")
                                    Set RSPart2 = Server.CreateObject ("ADODB.RecordSet")
                                    intcontadorunidades = 0
                                    Connx2.Open MM_EXTRANET_STRING2
                                    'strSQL = " select COUNT(REFE01) AS PARTIDAS                            " & _
                                    '         " from  tcepartidas                   " & _
                                    '         " where frac01 = '"&frac&"'       AND " & _
                                    '         "       refe01 = '"&referencia&"' AND " & _
                                    '         "       ( FECTICK >= '"&FormatoFechaInv(strDateIni)&"' AND            " & _
                                    '         "         HORAPES >= '"&strHoraIni&"'                     ) AND       " & _
                                    '         "       ( FECTICK <= '"&FormatoFechaInv(strDateFin)&"' AND            " & _
                                    '         "         HORAPES <= '"&strHoraFin&"'                     )      " & _
                                    '         "   and consec01 > 0            " &_
                                    '         "   and partComp = 0            " & StrfilTipVeh &_
                                    '         "   and pesoaudita > 0  " & _
                                    '         "   GROUP BY REFE01             "




                                    strSQL = " select COUNT(REFE01) AS PARTIDAS, SUM(pesoneto) as pesoalcorte " & _
                                             " from  tcepartidas                   " & _
                                             " where frac01 = '"&frac&"'       AND " & _
                                             "       refe01 = '"&referencia&"'     " & strFilTipoFecha2 &_
                                             "   and consec01 > 0            " &_
                                             "   and partComp = 0            " & StrfilTipVeh &_
                                             "   and pesoaudita > 0  " & _
                                             "   GROUP BY REFE01             "


                                    'Response.Write(strSQL)
                                    'Response.End
                                    dblPesoAlCorte      = 0
                                    dblSaldoAlcorte     = 0
                                    set RSPart2 = Connx2.Execute(strSQL)
                                    if not RSPart2.eof then
                                        Do while not RSPart2.Eof
                                            intcontadorunidades = RSPart2("PARTIDAS")
                                            dblPesoAlCorte      = RSPart2("pesoalcorte")
                                            dblSaldoAlcorte     = peso - dblPesoAlCorte
                                            RSPart2.MoveNext
                                        Loop ' de las partidas de tce
                                     else
                                       intcontadorunidades = 0
                                     end if
                                     '------------------------------------para la segunda tablas


                                     '--------------------------------------------------------------------
                                      Set Connx3  = Server.CreateObject ("ADODB.Connection")
                                      Set RSPart3 = Server.CreateObject ("ADODB.RecordSet")
                                      intcontadorunidades3 = 0
                                      Connx3.Open MM_EXTRANET_STRING2
                                      strSQL3 = " select COUNT(REFE01) AS PARTIDAS , SUM(pesoneto) as pesoalcorte                      " & _
                                               " from  tcepartidas                                                                     " & _
                                               " where frac01 = '"&frac&"'       AND                                                   " & _
                                               "       refe01 = '"&referencia&"'                                                       " & strFilTipoFecha4 &_
                                               "   and consec01 > 0                                                                    " &_
                                               "   and partComp = 0            " & StrfilTipVeh &_
                                               "   and pesoaudita > 0  " &_
                                               "   GROUP BY REFE01     "
                                      'Response.Write(strSQL3)
                                      'Response.End
                                      dblPesoAlCorte3 = 0
                                      dblSaldoAlcorte3 = 0
                                      set RSPart3 = Connx3.Execute(strSQL3)
                                      if not RSPart3.eof then
                                          Do while not RSPart3.Eof
                                              intcontadorunidades3 = RSPart3("PARTIDAS")
                                              dblPesoAlCorte3      = RSPart3("pesoalcorte")
                                              dblSaldoAlcorte3     = peso - dblPesoAlCorte3
                                              RSPart3.MoveNext
                                          Loop ' de las partidas de tce
                                       else
                                         intcontadorunidades3 = 0
                                       end if
                                       '------------------------------------para la segunda tablas

                                   '--------------------------------------------------------------------


                                     if(peso > 0) then
                                        peso = peso/1000
                                     end if

                                     if(saldolib > 0) then
                                        saldolib = saldolib/1000
                                     end if

                                     if(saldo > 0) then
                                        saldo = saldo/1000
                                     end if

                                     if(dblPesoAlCorte > 0) then
                                        dblPesoAlCorte = dblPesoAlCorte/1000
                                     end if

                                     if(dblSaldoAlcorte > 0) then
                                        dblSaldoAlcorte = dblSaldoAlcorte/1000
                                     end if

                                     if(dblPesoAlCorte3 > 0) then
                                        dblPesoAlCorte3 = dblPesoAlCorte3/1000
                                     end if

                                     if(dblSaldoAlcorte3 > 0) then
                                        dblSaldoAlcorte3 = dblSaldoAlcorte3/1000
                                     end if


                                %>

                               <BR>
                                 <table  border="1" cellspacing="3" cellpadding="3">

                                  <tr bgcolor="#C0C0C0">
                                     <td bgcolor="#C0C0C0" >
                                       <font size="1" color="#993300" >
                                          <b>Buque:</b>
                                       </FONT>
                                     </td>
                                     <td  align="left" colspan="2" >
                                       <font size="1" color="#993300">
                                          <!-- <b> RESPONSE.Write(buque) </b> -->
                                          <b> <%RESPONSE.Write(StrRnombar)%> </b>

                                       </font>
                                     </td>
                                     <td bgcolor="#C0C0C0" >
                                      <font size="1" color="#993300" >
                                        <b>Producto</b>
                                      </FONT>
                                     </td>

                                     <td  align="left" colspan="3">
                                      <font size="1" color="#993300" >
                                          <!-- <b> RESPONSE.Write( mercancia ) </b> -->
                                          <b> <%RESPONSE.Write( Strrmercan )%> </b>

                                      </font>
                                     </td>

                                     <td bgcolor="#C0C0C0" colspan="2">
                                      <font size="1" color="#993300" >
                                          <b>TM PEDIMENTO</b>
                                      </FONT>
                                     </td>

                                     <td  align="left">
                                      <font size="1" color="#993300">
                                        <b>

                                            <%RESPONSE.Write( trim(cstr(formatnumber(peso,3))) ) %>
                                        </b>
                                      </font>
                                     </td>

                                     <td bgcolor="#C0C0C0" colspan="2" >
                                      <font size="1" color="#993300" >
                                          <b>TM TOTAL LIBERADO </b>
                                      </FONT>
                                     </td>

                                     <td  align="left">
                                      <font size="1" color="#993300">
                                        <b>
                                            <%RESPONSE.Write( trim(cstr(formatnumber(saldolib,3))) ) %>
                                        </b>
                                      </font>
                                     </td>

                                     <td bgcolor="#C0C0C0" colspan="2" >
                                      <font size="1" color="#993300" >
                                          <b>TM SALDO TOTAL POR LIBERAR </b>
                                      </FONT>
                                     </td>

                                     <td  align="left">
                                      <font size="1" color="#993300">
                                        <b>
                                            <%RESPONSE.Write( trim(cstr(formatnumber(saldo,3))) ) %>
                                        </b>
                                      </font>
                                     </td>

                                 </tr>
                                 <tr bgcolor="#C0C0C0" >
                                     <td  >
                                       <font size="1" color="#993300" >
                                         <b>Pedimento:</b>
                                       </FONT>
                                     </td>
                                     <td  align="left" colspan="2">
                                       <font size="1" color="#993300">
                                         <b> <% RESPONSE.Write( trim(patente)&"-"&trim(pedimento) )%> </b>
                                       </font>
                                     </td>
                                     <td >
                                       <font size="1" color="#993300" >
                                          <b>Cliente</b>
                                       </FONT>
                                     </td>
                                     <td  align="left" colspan="3">
                                       <font size="1" color="#993300" >
                                          <b><%RESPONSE.Write(cliente)%> </b>
                                       </font>
                                     </td>
                                     <td  colspan="2">
                                       <font size="1" color="#993300" >
                                          <b>UNIDADES ACOMULADAS AL CORTE:</b>
                                       </FONT>
                                     </td>
                                     <td  align="left">
                                       <font size="1" color="#993300" >
                                          <b><%RESPONSE.Write(intcontadorunidades3)%> </b>
                                       </font>
                                     </td>

                                     <td bgcolor="#C0C0C0" colspan="2">
                                      <font size="1" color="#993300" >
                                          <b> TM LIBERADO ACOMULADO AL CORTE</b>
                                      </FONT>
                                     </td>

                                     <td  align="left">
                                      <font size="1" color="#993300">
                                        <b>
                                            <%RESPONSE.Write( trim(cstr(formatnumber(dblPesoAlCorte3,3))) ) %>
                                        </b>
                                      </font>
                                     </td>

                                     <td bgcolor="#C0C0C0" colspan="2">
                                      <font size="1" color="#993300" >
                                          <b> TM POR LIBERAR AL CORTE</b>
                                      </FONT>
                                     </td>

                                     <td  align="left">
                                      <font size="1" color="#993300">
                                        <b>
                                            <%RESPONSE.Write( trim(cstr(formatnumber(dblSaldoAlcorte3,3))) ) %>
                                        </b>
                                      </font>
                                     </td>

                                 </tr>
                                 <tr >

                                     <td colspan="7">

                                     </td>

                                     <td bgcolor="#C0C0C0" bgcolor="#C0C0C0" colspan="2">
                                      <font size="1" color="#993300" >
                                          <b>UNIDADES AL CORTE</b>
                                      </FONT>
                                     </td>

                                     <td  align="left">
                                      <font size="1" color="#993300">
                                        <b>

                                            <%RESPONSE.Write( trim(cstr(formatnumber(intcontadorunidades,0))) ) %>
                                        </b>
                                      </font>
                                     </td>

                                     <td bgcolor="#C0C0C0" colspan="2" >
                                      <font size="1" color="#993300" >
                                          <b>TM LIBERADO AL CORTE</b>
                                      </FONT>
                                     </td>

                                     <td  align="left">
                                      <font size="1" color="#993300">
                                        <b>
                                            <%RESPONSE.Write( trim(cstr(formatnumber(dblPesoAlCorte,3))) ) %>
                                        </b>
                                      </font>
                                     </td>

                                     <td bgcolor="#C0C0C0" colspan="2" >
                                      <font size="1" color="#993300" >
                                          <b></b>
                                      </FONT>
                                     </td>

                                     <td bgcolor="#C0C0C0" align="left">
                                      <font size="1" color="#993300">
                                        <b>

                                        </b>
                                      </font>
                                     </td>

                                 </tr>

                                 </table>
                               <BR>





                                  <%	   '*******************************************************

                                  '--------------------------------------------------------------------------------- 
								  %>
                                    <table border="1" cellspacing="3" cellpadding="3">
                                          <!--
                                          <tr>
                                             <th bgcolor="#009999">&nbsp;</th>
                                             <th bgcolor="#009999">&nbsp;</th>
                                             <th bgcolor="#009999">&nbsp;</th>
                                             <th bgcolor="#009999" colspan="2"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Transporte</FONT></th>
                                             <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Inicios</FONT></th>
                                             <th bgcolor="#009999">&nbsp;</th>
                                          </tr>
                                          -->
                                          <tr bgcolor="#000000" >
                                             <!--
                                             <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Fecha</FONT></th>
                                             <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Partida</FONT></th>
                                             <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">No.ticket</FONT></th>
                                             <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Tipo</FONT></th>
                                             <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Placas</FONT></th>
                                             <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">FL</FONT></th>
                                             <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Peso Neto</FONT></th>
                                             -->

                                             <td width="50px" ><font size="1" color="#FFFFFF" > <b> LOAD             </b> </FONT></td>
											 <td width="50px" ><font size="1" color="#FFFFFF" > <b> RECINTO             </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> FOLIO TICKET     </b> </FONT></td> <!-- Ticket -->
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> FECHA PESAJE     </b> </FONT></td> <!-- Fecha Ticket -->
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> HORA PESAJE      </b> </FONT></td> <!-- Hora -->

                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> FECHA LIBERACIÓN </b> </FONT></td> <!-- Fecha LIBERACIÓN -->
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> HORA LIBERACIÓN  </b> </FONT></td> <!-- Hora LIBERACION-->

                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> PESO TARA(KG)    </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> PESO BRUTO(KG)   </b> </FONT></td> <!-- Hora -->
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> PESO NETO(KG)    </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> TRANSPORTE       </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> PLACA CAJA       </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> SELLO EMBARQUE   </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> TIPO DE EMBARQUE </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> BARCO            </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> PRODUCTO         </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> CLIENTE          </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> CONTRATO         </b> </FONT></td>
                                             <td width="50px"><font size="1" color="#FFFFFF" > <b> SELLO ESLINGA    </b> </FONT></td>

                                          </tr>

                                  <% '************************************partidas tce**********************

                                  saldolib = 0
                                  Set Connx = Server.CreateObject ("ADODB.Connection")
                                  Set RSPart = Server.CreateObject ("ADODB.RecordSet")
                                  Connx.Open MM_EXTRANET_STRING2
                                  'strSQL= " SELECT a.refe01 as ref,        " &_
                                  '        "        a.pedi01 as ped,        " &_
                                  '        "        a.nomcli01 as clie,     " &_
                                  '        "        a.patent01 as pat,      " &_
                                  '        "        a.nombar01,             " &_
                                  '        "        a.fecpag01,             " &_
                                  '        "        b.frac01 as frac,       " &_
                                  '        "        b.desc01 as merc,       " &_
                                  '        "        a.pesobr01 as pesob,    " &_
                                  '        "        c.consec01 as consec,   " &_
                                  '        "        c.tipveh as tipveh,     " &_
                                  '        "        c.nload,                " &_
                                  '        "        c.ticket ,              " &_
                                  '        "        c.fectick,              " &_
                                  '        "        c.placas ,              " &_
                                  '        "        c.pesoBruto,            " &_
                                  '        "        c.pesoTara,             " &_
                                  '        "        c.pesoNeto,             " &_
                                  '        "        TIME_FORMAT(c.horaPes,'%H:%i') as horaPes, " &_
                                  '        "        c.destino,              " &_
                                  '        "        transpor,               " &_
                                  '        "        sellos,                 " &_
                                  '        "        contrato,               " &_
                                  '        "        sellose,                " &_
                                  '        "        mercan01                " &_
                                  '        " FROM pedimentos a,fracciones b,tcepartidas c " &_
                                  '        " where a.refe01='"&referencia&"'  " &_
                                  '        "   and a.refe01=b.refe01         " &_
                                  '        "   and a.refe01=c.refe01         " &_
                                  '        "   AND ( c.FECTICK >= '"&FormatoFechaInv(strDateIni)&"' AND      " & _
                                  '        "         c.HORAPES >= '"&strHoraIni&"'                     ) AND " & _
                                  '        "       ( c.FECTICK <= '"&FormatoFechaInv(strDateFin)&"' AND      " & _
                                  '        "         c.HORAPES <= '"&strHoraFin&"'                     )      " &permi &_
                                  '        "   and c.consec01 > 0            " &_
                                  '        "   and c.partComp = 0            " & StrfilTipVeh & _
                                  '        "   and pesoaudita > 0  "


                                  strSQL= " SELECT a.refe01 as ref,        " &_
                                          "        a.pedi01 as ped,        " &_
                                          "        a.nomcli01 as clie,     " &_
                                          "        a.patent01 as pat,      " &_
                                          "        a.nombar01,             " &_
                                          "        a.fecpag01,             " &_
                                          "        b.frac01 as frac,       " &_
                                          "        b.desc01 as merc,       " &_
                                          "        a.pesobr01 as pesob,    " &_
                                          "        c.consec01 as consec,   " &_
                                          "        c.tipveh as tipveh,     " &_
                                          "        c.nload,                " &_
                                          "        c.ticket ,              " &_
                                          "        c.fectick,              " &_
                                          "        c.placas ,              " &_
                                          "        c.pesoBruto,            " &_
                                          "        c.pesoTara,             " &_
                                          "        c.pesoNeto,             " &_
                                          "        TIME_FORMAT(c.horaPes,'%H:%i') as horaPes, " &_
                                          "        c.destino,              " &_
                                          "        transpor,               " &_
                                          "        sellos,                 " &_
                                          "        contrato,               " &_
                                          "        sellose,                " &_
                                          "        mercan01,               " &_
                                          "        c.FecDoc,               " &_
										  "        c.recinto,               " &_
                                          "        TIME_FORMAT(c.horaDoc,'%H:%i') as horaDoc  " &_
                                          " FROM pedimentos a,fracciones b,tcepartidas c " &_
                                          " where a.refe01='"&referencia&"' " &_
                                          "   and a.refe01=b.refe01         " &_
                                          "   and a.refe01=c.refe01         " & strFilTipoFecha3 &permi &_
                                          "   and c.consec01 > 0            " &_
                                          "   and c.partComp = 0            " & StrfilTipVeh & _
                                          "   and pesoaudita > 0  "



                                  'Response.Write(strSQL)
                                  Set RSPart= Connx.Execute(strSQL)

                                  Do while not RSPart.Eof

                                      'fechapartida=RSPart("fectick")
                                      'partida=RSPart("consec")
                                      'ticket=RSPart("ticket")
                                      'tipoveh=RSPart("tipveh")
                                      'placas=RSPart("placas")
                                      'inicios=RSPart("inicio")

                                      pesoneto     = RSPart("pesoneto")
                                      strnload     = RSPart("nload")
                                      strticket    = RSPart("ticket")
                                      strfectick   = RSPart("fectick")
                                      strpesoBruto = RSPart("pesoBruto")
                                      strpesoTara  = RSPart("pesoTara")
                                      strpesoNeto  = RSPart("pesoNeto")
                                      strtranspor  = RSPart("transpor")
                                      strplacas    = RSPart("placas")
                                      strsellos    = RSPart("sellos")
                                      strhoraPes   = RSPart("horaPes")
                                      strdestino   = RSPart("destino")
                                      strcontrato  = RSPart("contrato")
                                      strsellose   = RSPart("sellose")
                                      if(strsellose = "") then
                                        strsellose = "SN"
                                      end if
                                      strmercan    = RSPart("mercan01")
                                      strnombar    = RSPart("nombar01")
									  strrecinto   = RSpart("recinto")
                                      strhoraDoc   = RSPart("horaDoc")
                                      strfecDOc    = RSPart("FecDoc")

                                      saldolib=saldolib+round(pesoneto)
                                      'select case tipoveh
                                      '     case "C"
                                      '          Tipotrans="Camion"
                                      '     case "T"
                                      '          Tipotrans="Tolva"
                                      '     case "F"
                                      '          Tipotrans="Furgon"
                                      '     case "G"
                                      '          Tipotrans="Gondola"
                                      'end select
                                      '---------------------------TERCER TABLA PARTIDAS------------------
									  %>
                                       <tr>
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strnload)%>     </font></td><!-- No. LOAD      -->
										 <td ><font size="1" color="#000000" > <%RESPONSE.Write(strrecinto)%>     </font></td><!-- No. LOAD      -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strticket)%>    </font></td><!-- No. FICHA     -->


                                         <% 
                                           '<td><font size="1" color="#000000" > <RESPONSE.Write(strfectick)>   </font></td><!-- FECHA         -->
                                           '<td><font size="1" color="#000000" > <RESPONSE.Write(strhoraPes)>   </font></td><!-- HoraPes       -->
                                           'strhoraDoc   = RSPart("horaDoc")
                                           'strfecDOc    = RSPart("FecDoc")
                                         %>

                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strfectick)%>  </font></td><!-- FECHA         -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strhoraPes)%>  </font></td><!-- HoraPes       -->

                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strfecDOc)%>   </font></td><!-- FECHA LIBERACIÓN -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strhoraDoc)%>  </font></td><!-- Hora LIBERACIÓN  -->

                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write( trim(cstr(formatnumber(strpesoTara,0)))  )%> </font></td><!-- P.TARA  -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write( trim(cstr(formatnumber(strpesoBruto,0))) )%> </font></td><!-- P.BRUTO -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write( trim(cstr(formatnumber(strpesoNeto,0)))  )%> </font></td><!-- P. NETO -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strtranspor)%>  </font></td><!-- TRANSPORTE       -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strplacas)%>    </font></td><!-- PLACA TRACTOR    -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strsellos)%>    </font></td><!-- SELLO EMBARQUE   -->
                                         <td align="Center" ><font size="1" color="#000000" > <%RESPONSE.Write("1")%>          </font></td><!-- TIPO DE EMBARQUE -->

                                         <!--
                                         <td ><font size="1" color="#000000" > RESPONSE.Write(strnombar)    </font></td>
                                         <td ><font size="1" color="#000000" > RESPONSE.Write(strmercan)    </font></td>
                                         -->

                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(StrRnombar)%>    </font></td><!-- BARCO            -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(Strrmercan)%>    </font></td><!-- PRODUCTO         -->

                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strdestino)%>   </font></td><!-- CLIENTE          -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strcontrato)%>  </font></td><!-- CONTRATO         -->
                                         <td ><font size="1" color="#000000" > <%RESPONSE.Write(strsellose)%>   </font></td><!-- SELLO ESLINGA    -->
                                       </tr>

                                      <%

                                      RSPart.MoveNext  ' de las partidas de tce
                                  Loop ' de las partidas de tce



                                      %>
                               <tr>
                                    <td colspan="8" align="right">
                                      <font size="1" color="#000000" >
                                        <b>Saldo Liberado ( KG ) </b>
                                      </font>
                                    </td>
                                    <td>
                                      <font size="1" color="#000000" >
                                        <% RESPONSE.Write( trim(cstr(formatnumber(saldolib,0))) ) %>
                                      </font>
                                    </td>
                               </tr>
                               <%
                                   saldolib = saldolib/1000
                               %>
                               <tr>
                                    <td colspan="8" align="right">
                                      <font size="1" color="#000000" >
                                        <b>Saldo Liberado ( TM ) </b>
                                      </font>
                                    </td>
                                    <td>
                                      <font size="1" color="#000000" >
                                        <% RESPONSE.Write( trim(cstr(formatnumber(saldolib,3))) ) %>
                                      </font>
                                    </td>
                               </tr>
                            </table>
                          <%'**********************************************************************

                                ' RESPONSE.Write("</tr>")
                                  ' Refe.MoveNext 'avanza referencia  ---->
                              Refe.MoveNext
                                   wend 'REFErencia
                           else
						     Response.Write(strsql & "1")
						  response.end()
                          %>
						
                          <tr>
                            <th colspan=12>
                              <font size="2" face="Arial">No se Encontro ningun registro con esos parametros</font>
                            </th>
                          </tr>
                          <table>

                            <% 
                              'end if
                              end if
                            %>
                          </form>




 <% 
    else
	response.write(strsql & "2")
	response.end()
    %>
	
    <table>
    <tr>
      <th colspan=12>
        <font size="2" face="Arial">No se Encontro ningun registro con esos parametros</font>
      </th>
    </tr>
    <table>
    <% 
    end if

end if


else
  response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>")
end if%>

</BODY>
</HTML>


