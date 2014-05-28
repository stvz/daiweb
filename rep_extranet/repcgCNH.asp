
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% On Error Resume Next %>
<%
  MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

  'Response.Write("<br>")
  'Response.Write("<br>")
  'Response.Write(MM_EXTRANET_STRING)
  'Response.Write("<br>")
  'Response.Write("<br>")

  Response.Buffer = TRUE
  'Response.Addheader "Content-Disposition", "attachment;"
  'Response.ContentType = "application/vnd.ms-excel"

  'Response.ContentType = "application/csv"
  Response.ContentType = "text/html"

  'Response.AddHeader "Content-Disposition", "attachment; filename=Cuenta_de_Gastos.txt;"
  'Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo

  Server.ScriptTimeOut=100000

  strPermisos = Request.Form("Permisos")
  permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
  permi2 = PermisoClientesTabla("B",Session("GAduana") ,strPermisos,"clie31")

  'Response.Write("permi")
  'Response.Write(permi)
  'Response.Write("<BR>permi2")
  'Response.Write(permi2)
  'Response.Write("<BR>")
  'response.write(permi2)
  'Response.End

  if not permi = "" then
    permi = "  and (" & permi & ") "
  end if

  if not permi2 = "" then
    permi2 = "  and (" & permi2 & ") "
  end if


  blnAplicaFiltro= false
  strFiltroCliente = ""
  strFiltroCliente = request.Form("txtCliente")
  strTipoUsuario    = request.Form("TipoUser")
  if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
     blnAplicaFiltro = true
  end if
  if blnAplicaFiltro then
     permi = " AND cvecli01 =" & strFiltroCliente
     permi2 = " AND B.clie31 =" & strFiltroCliente
  end if

  if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
     permi = ""
     permi2 = ""
  end if

      'Response.Write("permi")
      'Response.Write(permi)
      'Response.Write("<BR>permi2")
      'Response.Write(permi2)
      'Response.Write("<BR>strFiltroCliente")
      'Response.Write(strFiltroCliente)
      'Response.Write("<BR>blnAplicaFiltro")
      'Response.Write(blnAplicaFiltro)
      'Response.Write("<BR>strTipoUsuario")
      'Response.Write(strTipoUsuario)
      'Response.Write("<BR>MM_Cod_Admon")
      'Response.Write(MM_Cod_Admon)
      'Response.Write(strSQL)
      'Response.End



  '***********************************
  '  AplicaFiltro = false
  '  strFiltroCliente = ""
  '  strFiltroCliente = request.Form("txtCliente")
  '  if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
  '     blnAplicaFiltro = true
  '  end if
  '  if blnAplicaFiltro then
  '     permi2 = " AND B.clie31 =" & strFiltroCliente
  '  end if
  '  if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
  '     permi2 = ""
  '  end if
   '***********************************




  strDateIni=""
  strDateFin=""
  strTipoPedimento= ""
  strCodError = "0"

  strDateIni        = trim(request.Form("txtDateIni"))
  strDateFin        = trim(request.Form("txtDateFin"))
  strTipoPedimento  = trim(request.Form("rbnTipoDate"))
  strTipoPedimento2 = trim(request.Form("rbnTipoDate2"))

  strDateIni        = trim(request.Form("txtDateIni"))
  strDateFin        = trim(request.Form("txtDateFin"))
  strDateIni2       = trim(request.Form("txtDateIni2"))
  strDateFin2       = trim(request.Form("txtDateFin2"))

  strtxtCta         = trim(request.Form("txtCta"))
  strtxtRefer       = trim(request.Form("txtRefer"))
  strtxtPed         = trim(request.Form("txtPed"))

  strUsuario        = request.Form("user")

  strDescripcion    = trim(request.Form("txtDescripcion"))


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
      strTipoFiltro=trim(request.Form("TipoFiltro"))
      tmpTipo = ""

      if strTipoFiltro  = "xrangofec" then ' por rango de fechas
         strSQL= "  select B.cgas31,  " & _
                 "         A.refe31,  " & _
                 "         B.fech31,  " & _
                 "         B.suph31,  " & _
                 "         B.coad31,  " & _
                 "         B.csce31,  " & _
                 "         B.chon31,  " & _
                 "         B.piva31,  " & _
                 "         B.anti31,  " & _
                 "         B.sald31,  " & _
                 "         B.tota31,  " & _
                 "         (B.tota31 - (B.suph31 + B.chon31 + B.csce31 + B.caho31)) as iva " & _
                 "  from e31cgast as B " & _
                 "       inner join d31refer as A " & _
                 "       on  B.cgas31 = A.cgas31  " & _
                 "  where B.esta31  = 'I'    and  " & _
                 "        B.FECH31 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                 "        B.FECH31 <= '"&FormatoFechaInv(strDateFin)&"' " & permi2
      else
          if strTipoFiltro  = "xCG" then 'por una cuenta de gastos especifica
             strSQL= "  select B.cgas31,  " & _
                     "         A.refe31,  " & _
                     "         B.fech31,  " & _
                     "         B.suph31,  " & _
                     "         B.coad31,  " & _
                     "         B.csce31,  " & _
                     "         B.chon31,  " & _
                     "         B.piva31,  " & _
                     "         B.anti31,  " & _
                     "         B.sald31,  " & _
                     "         B.tota31,  " & _
                     "         (B.tota31 - (B.suph31 + B.chon31 + B.csce31 + B.caho31)) as iva " & _
                     "  from e31cgast as B " & _
                     "       inner join d31refer as A " & _
                     "       on  B.cgas31 = A.cgas31  " & _
                     "  where B.esta31  = 'I'    and  " & _
                     "  B.cgas31 = '"& strtxtCta & "'"    & permi2
          else
              if strTipoFiltro  = "xreferencia" then 'por una referencia especifica
                 strSQL= "  select B.cgas31,  " & _
                     "         A.refe31,  " & _
                     "         B.fech31,  " & _
                     "         B.suph31,  " & _
                     "         B.coad31,  " & _
                     "         B.csce31,  " & _
                     "         B.chon31,  " & _
                     "         B.piva31,  " & _
                     "         B.anti31,  " & _
                     "         B.sald31,  " & _
                     "         B.tota31,  " & _
                     "         (B.tota31 - (B.suph31 + B.chon31 + B.csce31 + B.caho31)) as iva " & _
                     "  from e31cgast as B " & _
                     "       inner join d31refer as A " & _
                     "       on  B.cgas31 = A.cgas31  " & _
                     "  where B.esta31  = 'I'    and  " & _
                     "  A.refe31 = '"& strtxtRefer & "'"    & permi2
					 
              else
                  if strTipoFiltro  = "xpedimento" then 'por un pedimento especifico
                       refaux = ""
                       if strTipoPedimento = "1" then
                         strSQLPed = " SELECT REFCIA01 FROM ssdagi01 WHERE NUMPED01='"& strtxtPed &"'"
                       end if
                       if strTipoPedimento  = "2" then
                         strSQLPed = " SELECT REFCIA01 FROM ssdage01 WHERE NUMPED01='"& strtxtPed &"'"
                       end if

                       Set rsBuscaPedAux = Server.CreateObject("ADODB.Recordset")
                       rsBuscaPedAux.ActiveConnection = MM_EXTRANET_STRING
                       rsBuscaPedAux.Source = strSQLPed
                       rsBuscaPedAux.CursorType = 0
                       rsBuscaPedAux.CursorLocation = 2
                       rsBuscaPedAux.LockType = 1
                       rsBuscaPedAux.Open()
                       if not rsBuscaPedAux.eof then
                           'While NOT rsBuscaPedAux.EOF
                                   refaux  = rsBuscaPedAux.Fields.Item("REFCIA01").Value
                           'rsBuscaPedAux.movenext
                           'Wend
                       end if
                       rsBuscaPedAux.close
                       set rsBuscaPedAux = Nothing

                       strSQL= "  Select B.cgas31,  " & _
                               "         A.refe31,  " & _
                               "         B.fech31,  " & _
                               "         B.suph31,  " & _
                               "         B.coad31,  " & _
                               "         B.csce31,  " & _
                               "         B.chon31,  " & _
                               "         B.piva31,  " & _
                               "         B.anti31,  " & _
                               "         B.sald31,  " & _
                               "         B.tota31,  " & _
                               "         (B.tota31 - (B.suph31 + B.chon31 + B.csce31 + B.caho31)) as iva " & _
                               "  From e31cgast as B " & _
                               "       inner join d31refer as A " & _
                               "       on  B.cgas31 = A.cgas31  " & _
                               "  Where B.esta31  = 'I'    and  " & _
                               "  A.refe31 = '"& refaux & "'"    & permi2
'strtxtCta
              'strtxtRefer
              'strtxtPed

                  end if
              end if
          end if
      end if




      'Response.Write("permi")
      'Response.Write(permi)
      'Response.Write("<BR>permi2")
      'Response.Write(permi2)
      'Response.Write("<BR>strFiltroCliente")
      'Response.Write(strFiltroCliente)
      'Response.Write("<BR>blnAplicaFiltro")
      'Response.Write(blnAplicaFiltro)
      'Response.Write("<BR>strTipoUsuario")
      'Response.Write(strTipoUsuario)
      'Response.Write("<BR>MM_Cod_Admon")
      'Response.Write(MM_Cod_Admon)
      ' Response.Write(strSQL)
      ' Response.End




      '******************************************************************
		
	   codflete = ""
       codnoflete = ""
       intcontflete = 1
       Set RsbuscaFlete = Server.CreateObject("ADODB.Recordset")
       RsbuscaFlete.ActiveConnection = MM_EXTRANET_STRING
       strSQLcodfle = " select clav21 " & _
                " from c21paghe "&_
                " where desc21 like '%FLETE%' "
       RsbuscaFlete.Source = strSQLcodfle
       RsbuscaFlete.CursorType = 0
       RsbuscaFlete.CursorLocation = 2
       RsbuscaFlete.LockType = 1
       RsbuscaFlete.Open()
       if not RsbuscaFlete.EOF then
           While NOT RsbuscaFlete.EOF
            if intcontflete = 1 then
              codflete   =  "conc21=" & RsbuscaFlete.Fields.Item("clav21").Value
              codnoflete =  "conc21<>" & RsbuscaFlete.Fields.Item("clav21").Value
            else
              codflete   =  codflete & " OR " & "conc21=" & RsbuscaFlete.Fields.Item("clav21").Value
              codnoflete =  codnoflete & " and " & "conc21<>" & RsbuscaFlete.Fields.Item("clav21").Value
            end if
            RsbuscaFlete.movenext
            intcontflete = intcontflete + 1
           Wend
       end if
       RsbuscaFlete.close
       set  RsbuscaFlete = nothing
       codflete = "(" & codflete & ")"

       'Response.Write(codflete)
       'Response.Write("<br><br><br>")
       'Response.Write(codnoflete)
       'Response.End
      '******************************************************************

      strRefcia    = ""
      strCgas      = ""
      strFecCgas   = ""
      strSuph      = ""
      strCoad      = ""
      strCsce      = ""
      strChon      = ""
      strIva       = ""
      strPiva      = ""
      strAnti      = ""
      strSald      = ""
      strCveCli    = ""
      strAduana    = ""
      strSeccion   = ""
      strPatente   = ""
      strNumped    = ""
      strEmpresa   = ""
      strNomcli    = ""
	  strRFCCli    = ""
      strtotal     = ""
      strimpuestosPed   = ""
      strFlete     = ""
      strOtros     = 0
      strivaOtros  = 0
	  strIVAFletes  = 0
      strnomArcPat = ""
      strnomArcCli = ""
      intContnomArcPat = 1
      intContnomArcCli = 1

	  ' Response.Write(strSQL)
      ' Response.End()
      if not trim(strSQL)="" then
             Set rsCGPrincipal = Server.CreateObject("ADODB.Recordset")
             rsCGPrincipal.ActiveConnection = MM_EXTRANET_STRING

             rsCGPrincipal.Source = strSQL
             rsCGPrincipal.CursorType = 0
             rsCGPrincipal.CursorLocation = 2
             rsCGPrincipal.LockType = 1
             rsCGPrincipal.Open()
			 
             if not rsCGPrincipal.eof then
                 While NOT rsCGPrincipal.EOF

                     strRefcia  = rsCGPrincipal.Fields.Item("refe31").Value
                     strCgas    = rsCGPrincipal.Fields.Item("cgas31").Value
                     'strCgas    = rsCGPrincipal.Fields.Item("cgas31").Value

                     ' if IsNumeric(strCgas) then
                       ' strCgas =  CLng(strCgas)
                     ' else
                       ' 'response.write(strCgas)
                       ' strCgas = Mid(strCgas,2,(len(strCgas) - 1))
                     ' end if

                     strFecCgas = rsCGPrincipal.Fields.Item("fech31").Value
                     strSuph    = rsCGPrincipal.Fields.Item("suph31").Value
                     strCoad    = rsCGPrincipal.Fields.Item("coad31").Value
                     strCsce    = rsCGPrincipal.Fields.Item("csce31").Value
                     strChon    = rsCGPrincipal.Fields.Item("chon31").Value
                     strPiva    = rsCGPrincipal.Fields.Item("piva31").Value
                     strAnti    = rsCGPrincipal.Fields.Item("anti31").Value
                     strSald    = rsCGPrincipal.Fields.Item("sald31").Value
                     strtotal   = rsCGPrincipal.Fields.Item("tota31").Value
                     strIva     = rsCGPrincipal.Fields.Item("iva").Value

                     '*******************************************************************
                     strOficina=left(trim(strRefcia),3)
						strSQL1 = " SELECT CVECLI01, NOMCLI01, RFCCLI01, CVEADU01, CVESEC01 ,PATENT01, NUMPED01 " & _
								  " FROM SSDAGI01 " & _
								  " WHERE REFCIA01 = '" & strRefcia & "' " & _
								  " UNION ALL " & _
								  " SELECT CVECLI01, NOMCLI01, RFCCLI01, CVEADU01, CVESEC01 ,PATENT01, NUMPED01 " & _
								  " FROM SSDAGE01 " & _
								  " WHERE REFCIA01 = '" & strRefcia & "' "
                     Set rsBuscaPed = Server.CreateObject("ADODB.Recordset")
                     rsBuscaPed.ActiveConnection = ODBC_POR_ADUANA(strOficina)
                     rsBuscaPed.Source = strSQL1
                     rsBuscaPed.CursorType = 0
                     rsBuscaPed.CursorLocation = 2
                     rsBuscaPed.LockType = 1
                     rsBuscaPed.Open()
                     if not rsBuscaPed.eof then
                         While NOT rsBuscaPed.EOF
                                 strCveCli  = rsBuscaPed.Fields.Item("CVECLI01").Value
								 strRFCCli  = rsBuscaPed.Fields.Item("RFCCLI01").Value
                                 strAduana  = rsBuscaPed.Fields.Item("CVEADU01").Value
                                 strSeccion = rsBuscaPed.Fields.Item("CVESEC01").Value
                                 strPatente = rsBuscaPed.Fields.Item("PATENT01").Value
                                 strNumped  = rsBuscaPed.Fields.Item("NUMPED01").Value
                                 strNomcli  = rsBuscaPed.Fields.Item("NOMCLI01").Value
                         rsBuscaPed.movenext
                         Wend
                     end if
                     rsBuscaPed.close
                     set rsBuscaPed = Nothing
                     '******************************************************************

                     Set RsValor1 = Server.CreateObject("ADODB.Recordset")
                     RsValor1.ActiveConnection = MM_EXTRANET_STRING
                     'strSQLph = " select refe21,"&_
                     '           "        sum(IF(conc21=1, (if(deha21 ='C',-1,1) ) *d.mont21,0 ))  as impuestos,"&_
                     '           "        sum(IF("& codflete &", (if(deha21 ='C',-1,1) ) *d.mont21, 0 ))  as fletes, "&_
                     '           "        sum(IF(conc21<>1 and " & codnoflete & ", (if(deha21 ='C',-1,1) ) *d.mont21,0)) as otros "&_
                     '           " from d21paghe as d, e21paghe as e  "&_
                     '           " where e.foli21 = d.foli21 and     "&_
                     '           "       e.fech21 = d.fech21 and  "&_
                     '           "       (e.esta21 = 'A'  or e.esta21='E' ) and   "&_
                     '           "       d.refe21 = '" &strRefcia& "' "&_
                     '           "       and  cgas21 = '" &pd(strCgas,7)& "' "&_
                     '           " group by refe21 "
					
					 '" sum(IF(conc21<>1 and " & codnoflete & " , (if(deha21 ='C',-1,1) ) *d.mont21/ (1+(piva21/100)  ) ,0)) as MontoSinIva "&_
					 If strRFCCli = "CME950209J18" then
						'Para CASE
						strSQLph = " select refe21, "&_
                                " sum(IF(conc21=1, (if(deha21 ='C',-1,1) ) *d.mont21,0 )) as impuestos, "&_
                                " sum(IF("& codflete &", (if(deha21 ='C',-1,1) ) *d.mont21, 0 )) as fletes,  "&_
                                " sum(IF(conc21<>1 and " & codnoflete & " , (if(deha21 ='C',-1,1) ) *d.mont21,0)) as otros , "&_
								" SUM(IF("& codflete &",(if(deha21 ='C',-1,1) * mfle21), 0)) * -0.04 AS 'reten', " &_
                                " ROUND(SUM(((piva21/100) * (d.mont21 / " &_
								" (1 + ((piva21 - IF("& codflete &", 4, 0)) / 100))))), 2) AS 'MontoSinIva' " &_
								" from d21paghe as d, e21paghe as e "&_
                                " where e.foli21 = d.foli21 "&_
                                " and year(e.fech21) = year(d.fech21) "&_
                                " and (e.esta21 = 'A' or e.esta21='E' ) "&_
                                " and d.refe21 = '" &strRefcia& "' "&_
                                " and cgas21 = '" &pd(strCgas,7)& "' "&_
                                " group by refe21 "
					 
					 else
						'Para CNH
						strSQLph = " select refe21, "&_
                                " sum(IF(conc21=1, (if(deha21 ='C',-1,1) ) *d.mont21,0 )) as impuestos, "&_
                                " sum(IF("& codflete &", (if(deha21 ='C',-1,1) ) *d.mont21, 0 )) as fletes,  "&_
                                " sum(IF(conc21<>1 and " & codnoflete & " , (if(deha21 ='C',-1,1) ) *d.mont21,0)) as otros , "&_
								" SUM(IF("& codflete &",(if(deha21 ='C',-1,1) * mfle21), 0)) * -0.04 AS 'reten', " &_
                                " ROUND(SUM(((if(conc21=63,16,piva21)/100) * (d.mont21 / " &_
								" (1 + ((if(conc21=63,16,piva21) - IF((conc21=3 OR conc21=7 OR conc21=19 OR conc21=74 OR conc21=104 OR conc21=110 OR conc21=145 OR conc21=174), -400000000, 0)) / 100))))), 2) AS 'MontoSinIva', " &_
								" ROUND(SUM(if((conc21=3 OR conc21=7 OR conc21=19 OR conc21=74 OR conc21=104 OR conc21=110 OR conc21=145 "&_
								" OR conc21=174),if(deha21 = 'C',-1,1)*(mfle21*(piva21/100)),0)),2) as IVAfletes "&_
								" from d21paghe as d, e21paghe as e "&_
                                " where e.foli21 = d.foli21 "&_
                                " and year(e.fech21) = year(d.fech21) "&_
                                " and (e.esta21 = 'A' or e.esta21='E' ) "&_
                                " and d.refe21 = '" &strRefcia& "' "&_
                                " and cgas21 = '" &pd(strCgas,7)& "' "&_
                                " group by refe21 "
								
					 end if				
                      'Response.Write(strSQLph & "<br><br>")
					  'Response.Write(codflete & "<br><br>")
                      'Response.End

                     RsValor1.Source = strSQLph
                     RsValor1.CursorType = 0
                     RsValor1.CursorLocation = 2
                     RsValor1.LockType = 1
                     RsValor1.Open()
                     if not RsValor1.EOF then
                         strimpuestosPed = RsValor1.Fields.Item("impuestos").Value
                         strFlete        = RsValor1.Fields.Item("fletes").Value
                         strOtros        = RsValor1.Fields.Item("otros").Value
                         strMontoSinIva  = RsValor1.Fields.Item("MontoSinIva").Value
                         'strivaOtros     = strOtros - Round(strMontoSinIva,2)
						 strivaOtros     = Round(strMontoSinIva,2)
						 strReten  = RsValor1.Fields.Item("reten").Value
						 strIVAFletes = RsValor1.Fields.Item("IVAfletes").Value
                     end if
                     RsValor1.close
                     set  RsValor1 = nothing
                     '******************************************************************


                     '******************************************************************
                     montofleteAux2 = 0
                     montoSvrComp   = 0
                     Set RsValor2 = Server.CreateObject("ADODB.Recordset")
                     RsValor2.ActiveConnection = MM_EXTRANET_STRING
                     strSQLsvr = " SELECT refe32,ttar32, dcrp32,mont32" &_
                                 " FROM d32rserv  " &_
                                 " WHERE REFE32 = '" &strRefcia& "'"

                     'Response.Write(strSQLsvr)
                     'Response.End

                     RsValor2.Source = strSQLsvr
                     RsValor2.CursorType = 0
                     RsValor2.CursorLocation = 2
                     RsValor2.LockType = 1
                     RsValor2.Open()
                     if not RsValor2.EOF then
                         While NOT RsValor2.EOF
                           if (RsValor2.Fields.Item("ttar32").Value = "00033") then
                              montofleteAux2  = montofleteAux2 + RsValor2.Fields.Item("mont32").Value
                           else
                              montoSvrComp  = montoSvrComp + RsValor2.Fields.Item("mont32").Value
                           end if
                           RsValor2.movenext
                         Wend
                     end if
                     RsValor2.close
                     set  RsValor2 = nothing
                     strFlete = strFlete + montofleteAux2
                     '******************************************************************

                         strEmpresa = ""
                         if InStr(strNomcli ,"CASE") > 0 then
                            strEmpresa = "CASE"
                         else
                            if InStr(strNomcli ,"CNH COMPONENTES") > 0 then
                                strEmpresa = "CNHCMP"
                            else
                               if InStr(strNomcli ,"CNH COMERCIAL") > 0 then
                                  strEmpresa = "CNHCOM"
                               else
                                   if InStr(strNomcli ,"CNH INDUSTRIAL") > 0 then
                                      strEmpresa = "CNHIND"
                                   else
                                       if InStr(strNomcli ,"CNH DE MEXICO") > 0 then
                                          strEmpresa = "CNHMEX"
                                       else
                                         if InStr(strNomcli ,"CNH SERVICIOS CORPORATIVOS") > 0 then
                                            strEmpresa = "CNHSER"
                                         else
                                           if InStr(strNomcli ,"NEW HOLLAND") > 0 then
                                              strEmpresa = "NHMEX"
                                           else
                                              strEmpresa = strNomcli
                                           end  if
                                         end  if
                                       end  if
                                   end  if
                               end  if
                            end  if
                         end  if

                         if intContnomArcPat = 1 then
                              strnomArcPat = strPatente
                         else
                            if strnomArcPat <> strPatente then
                              strnomArcPat = ""
                            end if
                         end if

                         if intContnomArcCli = 1 then
                            strnomArcCli = strEmpresa
                         else
                            if strnomArcCli <> strEmpresa then
                               strnomArcCli = ""
                            end if
                         end if

                         'strnomArcPat
                         'strnomArcCli

                         'Response.Write("<BR>")
						 
                         Response.Write( strEmpresa )
                         Response.Write(",")
                         Response.Write( strAduana )
                         Response.Write(",")
                         Response.Write( strSeccion )
                         Response.Write(",")
                         Response.Write( strPatente )
                         Response.Write(",")
                         Response.Write( strNumped )
                         Response.Write(",")
                         Response.Write( strRefcia )
                         Response.Write(",")
                         Response.Write( strCgas  )
                         Response.Write(",")
                         Response.Write( strFecCgas )
                         Response.Write(",")
                         Response.Write("PS" )
                         Response.Write(",")
                         Response.Write(strimpuestosPed) 'Impuestos del pedimento
                         Response.Write(",")
                         Response.Write("0" ) 'Franja amarilla
                         Response.Write(",")
                         Response.Write( strChon ) ' Honorarios
                         Response.Write(",")
                         Response.Write(strFlete) 'Fletes
                         Response.Write(",")
                         Response.Write(montoSvrComp ) 'Servicios complementarios
                         Response.Write(",")
                         Response.Write(strOtros) 'Otros
                         Response.Write(",")
                         Response.Write("0" ) 'valor cuenta USA
                         Response.Write(",")
                         Response.Write( strIva ) 'IVA
                         Response.Write(",")
                         Response.Write( strtotal ) 'total de cuenta de gastos
                         Response.Write(",")
                         Response.Write( "0" ) ' Factura americana
                         Response.Write(",")
                         Response.Write( "0" ) ' Cuenta americana
                         Response.Write(",")
                         Response.Write( "" )  'Moneda
                         Response.Write(",")
                         Response.Write( strivaOtros ) 'Iva Otros Gastos
                         Response.Write(",")
						 If strRFCCli = "CME950209J18" then
							Response.Write( "0" ) 'Iva Fletes para CASE
						 else
							Response.Write( strIVAFletes ) 'Iva Fletes para CNH
						 end if
                         Response.Write(",")
                         Response.Write( strReten ) 'Retención Fletes
                         Response.write vbNewLine

             '        'end if
                 rsCGPrincipal.movenext
                 intContnomArcPat = intContnomArcPat + 1
                 intContnomArcCli = intContnomArcCli + 1
                 Wend
             end if
             rsCGPrincipal.close
             set rsCGPrincipal = Nothing
             '*******************************************************************
      end if




      'Response.AddHeader "Content-Disposition", "attachment; filename=3210_COM_03-03-08.txt;"
      'Response.AddHeader "Content-Disposition", "attachment; filename=3210_COM_"&CStr(day(date()))&"-"&CStr(month(date()))&"-"&CStr(year(date()))&".csv;"
      Response.AddHeader "Content-Disposition", "attachment; filename="&strnomArcPat&"_"&strnomArcCli&"_"&formatofechaNum(date() )&".csv;"


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


<%
   Function pd(n, totalDigits)
      if totalDigits > len(n) then
          pd = String(totalDigits-len(n),"0") & n
      else
          pd = n
      end if
  End Function

  Function formatofechaNum(DFecha)

     'Response.End

     if isdate( DFecha ) then
        strDateaux = mid( CStr( YEAR(DFecha) ) ,3, 2)
        formatofechaNum = Pd(DAY( DFecha ),2) & "-" & Pd(Month( DFecha ),2) & "-" & strDateaux
     else
        formatofechaNum	= DFecha
     end if
  End Function
%>





