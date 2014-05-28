
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% On Error Resume Next %>
<%



 'Response.End

  MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

  'Response.Write("<br>")
  'Response.Write("<br>")
  'Response.Write(MM_EXTRANET_STRING)
  'Response.Write("<br>")
  'Response.Write("<br>")

  Response.Buffer = TRUE
  Response.ContentType = "application/csv"
  Response.AddHeader "Content-Disposition", "attachment; filename=Pedimentos.csv;"



  'Response.Addheader "Content-Disposition", "attachment;"
  'Response.ContentType = "application/vnd.ms-excel"

  ''Response.ContentType = "text/html"
  ''Response.AddHeader "Content-Disposition", "attachment; filename=" & archivo


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

  AplicaFiltro = false
  strFiltroCliente = ""

  strFiltroCliente = request.Form("txtCliente")
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


  strDateIni=""
  strDateFin=""
  strTipoPedimento= ""
  strCodError = "0"

  'strDateIni        = trim(request.Form("txtDateIni"))
  'strDateFin        = trim(request.Form("txtDateFin"))

  strTipoPedimento  = trim(request.Form("rbnTipoDate1"))
  strTipoPedimento2 = trim(request.Form("rbnTipoDate2"))
  strTipoPedimento3 = trim(request.Form("rbnTipoDate3"))

  strDateIni        = trim(request.Form("txtDateIni"))
  strDateFin        = trim(request.Form("txtDateFin"))
  strDateIni2       = trim(request.Form("txtDateIni2"))
  strDateFin2       = trim(request.Form("txtDateFin2"))

  strtxtCta         = trim(request.Form("txtCta"))
  strtxtRefer       = trim(request.Form("txtRefer"))
  strtxtPed         = trim(request.Form("txtPed"))

  strUsuario        = request.Form("user")
  strTipoUsuario    = request.Form("TipoUser")
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

      'response.write(strTipoFiltro)
      'response.write("<br>")
      'response.write(strTipoPedimento)
      'response.write("<br>")
      'response.write(strTipoPedimento2)
      'response.write("<br>")
      'response.write(strTipoPedimento3)

      if strTipoFiltro  = "xrangofec" then ' por rango de fechas

          '      'Paginas del Pedimento	 NONE; 0
          '      'Referencia           REFCIA01
          '      'Pedimento	          NUMPED01
          '      'Patente	            patent01
          '      'Aduana               CVEADU01
          '      'Sección	            CVESEC01
          '      'Tipo de Operación
          '      'Clave de Pedimento	  CVEPED01
          '      'Regimen	            REGIME01
          '      'Destino Origen	      zonaod01 as Destino_Origen,
          '      'Tipo de Cambio	      TIPCAM01
          '      'Peso Total           pesobr01
          '      'Transporte (E/S)	    cvemtr01
          '      'Transporte (Arribo)	cvemta01
          '      'Transporte (Salida)	cvemts01
          '      Valor Dolares
          '      Valor Aduana (valor factura)
          '      Valor Comercial
          '      Empresa
          '      'Valor Seguros   valseg01 as valorseguros
          '      'Seguro          segros01 as seguros,
          '      'Flete           fletes01 as fletes,
          '      'Embalajes       embala01 as embalajes
          '      'Incrementables   incble01 as otrosIncrementables,
          '      BLS
          '      'Fecha de Entrada   FECENT01
          '      'Fecha de pago      FECPAG01
          '      Concepto (IGI)
          '      Concepto (IVA)
          '      Concepto (DTA)
          '      Concepto (PRV)
          '      Concepto (Multas)
          '      Concepto (Recargos)
          '      concepto (Cuotas compensatorias)
          '      Concepto (Articulo 303)
          '      Concepto (RT)
          '      Concepto (Otros)
          '      Concepto (ISAN)
          '      Concepto (IEPS)
          '      FORMA DE PAGO  (IGI)
          '      FORMA DE PAGO  (IVA)
          '      FORMA DE PAGO  (DTA)
          '      FORMA DE PAGO  (PRV)
          '      FORMA DE PAGO  (Multas)
          '      FORMA DE PAGO  (Recargos)
          '      FORMA DE PAGO  (Cuotas compensatorias)
          '      FORMA DE PAGO (Articulo 303)
          '      FORMA DE PAGO (RT)
          '      FORMA DE PAGO (Otros)
          '      FORMA DE PAGO (ISAN)
          '      FORMA DE PAGO (IEPS)
          '      Importe Total
          '      Proveedor (TAX ID)
          '      Vinculacion
          '      No. Guia (Master)
          '      No. Guia (House)
          '      Talón Flete/No. Transporte
          '      Observaciones
          '      'Fecha Despacho Puerto   FDSP01 as despacho
          '      'Fecha de Embarque       fpro01 as embarque
          '      'Fecha de Previo         FPRE01 as Previo
          '      'Fecha Revalidacion      FREV01 as Revalidacion
          '      Transportista
          '      Consolidador
          '      Moneda
          '      R1 o A3
          '      Pedimento Original
          '      RFC

         if strTipoPedimento = "1" then
             strSQL= " SELECT REFCIA01," & _
                     "        NUMPED01," & _
                     "        patent01," & _
                     "        CVEADU01," & _
                     "        CVESEC01," & _
                     "        'IMP' AS TIPOPE01 ," & _
                     "        CVEPED01," & _
                     "        REGIME01," & _
                     "        zonaod01 as Destino_Origen," & _
                     "        TIPCAM01," & _
                     "        pesobr01," & _
                     "        cvemtr01," & _
                     "        cvemta01," & _
                     "        cvemts01," & _
                     "        CVECLI01," & _
                     "        valseg01 as valorseguros," & _
                     "        segros01 as seguros," & _
                     "        fletes01 as fletes," & _
                     "        embala01 as embalajes," & _
                     "        incble01 as otrosIncrementables," & _
                     "        FECENT01 as entrada," & _
                     "        FECPAG01," & _
                     "        CVEPRO01,  " & _
                     "        FDSP01 as despacho," & _
                     "        fpro01 as embarque," & _
                     "        FPRE01 as Previo," & _
                     "        FREV01 as Revalidacion," & _
                     "        cvepro01, " & _
                     "        irspro22 taxidProv, " & _
                     "        nomcli01, " & _
                     "        cvepfm01,  " & _
                     "        VALDOL01,  " & _
                     "        ROUND( VALMER01*TIPCAM01*factmo01) AS VALORCOMERCIAL " & _
                     " FROM SSDAGI01 inner join c01refer on refe01=refcia01 " & _
                     "               inner join ssprov22 on  cvepro01=cvepro22 " & _
                     " WHERE FIRMAE01 <> '' AND " & _
                     " FECPAG01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                     " FECPAG01 <= '"&FormatoFechaInv(strDateFin)&"' " & permi
         end if
         if strTipoPedimento  = "2" then
             strSQL= " SELECT REFCIA01," & _
                     "        NUMPED01," & _
                     "        patent01," & _
                     "        CVEADU01," & _
                     "        CVESEC01," & _
                     "        'EXP' AS TIPOPE01 ," & _
                     "        CVEPED01," & _
                     "        REGIME01," & _
                     "        zonaod01 as Destino_Origen," & _
                     "        TIPCAM01," & _
                     "        pesobr01," & _
                     "        cvemtr01," & _
                     "        cvemta01," & _
                     "        cvemts01," & _
                     "        CVECLI01," & _
                     "        valseg01 as valorseguros," & _
                     "        segros01 as seguros," & _
                     "        fletes01 as fletes," & _
                     "        embala01 as embalajes," & _
                     "        incble01 as otrosIncrementables," & _
                     "        FECPRE01 as entrada," & _
                     "        FECPAG01," & _
                     "        CVEPRO01,  " & _
                     "        FDSP01 as despacho," & _
                     "        fpro01 as embarque," & _
                     "        FPRE01 as Previo," & _
                     "        FREV01 as Revalidacion," & _
                     "        cvepro01, " & _
                     "        irspro22 taxidProv, " & _
                     "        nomcli01, " & _
                     "        cvepfm01,  " & _
                     "        VALDOL01,  " & _
                     "        0 AS VALORCOMERCIAL " & _
                     " FROM SSDAGE01 inner join c01refer on refe01=refcia01 " & _
                     "               inner join ssprov22 on  cvepro01=cvepro22 " & _
                     " WHERE FIRMAE01 <> '' AND " & _
                     " FECPAG01 >= '"&FormatoFechaInv(strDateIni)&"' AND " & _
                     " FECPAG01 <= '"&FormatoFechaInv(strDateFin)&"' " & permi


                     '"        ROUND( VALMER01*TIPCAM01*factmo01) AS VALORCOMERCIAL " & _
         end if
                        '--Valor Dolares
                        '--Valor Aduana (valor factura)
                        '--Valor Comercial
                        '--BLS
                        '--Concepto (IGI)
                        '--Concepto (IVA)
                        '--Concepto (DTA)
                        '--Concepto (PRV)
                        '--Concepto (Multas)
                        '--Concepto (Recargos)
                        '--concepto (Cuotas compensatorias)
                        '--Concepto (Articulo 303)
                        '--Concepto (RT)
                        '--Concepto (Otros)
                        '--Concepto (ISAN)
                        '--Concepto (IEPS)
                        '--FORMA DE PAGO  (IGI)
                        '--FORMA DE PAGO  (IVA)
                        '--FORMA DE PAGO  (DTA)
                        '--FORMA DE PAGO  (PRV)
                        '--FORMA DE PAGO  (Multas)
                        '--FORMA DE PAGO  (Recargos)
                        '--FORMA DE PAGO  (Cuotas compensatorias)
                        '--FORMA DE PAGO (Articulo 303)
                        '--FORMA DE PAGO (RT)
                        '--FORMA DE PAGO (Otros)
                        '--FORMA DE PAGO (ISAN)
                        '--FORMA DE PAGO (IEPS)
                        '--Importe Total
                        '--Proveedor (TAX ID)
                        '--Vinculacion
                        '--No. Guia (Master)
                        '--No. Guia (House)
                        '--Talón Flete/No. Transporte
                        '--Observaciones
                        '--Transportista
                        '--Consolidador
                        '--Moneda
                        '--R1 o A3
                        '--Pedimento Original
                        '--RFC

      else
          if strTipoFiltro  = "xreferencia" then 'por una referencia especifica
                   if strTipoPedimento2 = "1" then
                       strSQL= " SELECT REFCIA01," & _
                               "        NUMPED01," & _
                               "        patent01," & _
                               "        CVEADU01," & _
                               "        CVESEC01," & _
                               "        'IMP' AS TIPOPE01 ," & _
                               "        CVEPED01," & _
                               "        REGIME01," & _
                               "        zonaod01 as Destino_Origen," & _
                               "        TIPCAM01," & _
                               "        pesobr01," & _
                               "        cvemtr01," & _
                               "        cvemta01," & _
                               "        cvemts01," & _
                               "        CVECLI01," & _
                               "        valseg01 as valorseguros," & _
                               "        segros01 as seguros," & _
                               "        fletes01 as fletes," & _
                               "        embala01 as embalajes," & _
                               "        incble01 as otrosIncrementables," & _
                               "        FECENT01," & _
                               "        FECPAG01," & _
                               "        CVEPRO01,  " & _
                               "        FDSP01 as despacho," & _
                               "        fpro01 as embarque," & _
                               "        FPRE01 as Previo," & _
                               "        FREV01 as Revalidacion," & _
                               "        irspro22 taxidProv, " & _
                               "        nomcli01, " & _
                               "        cvepfm01,  " & _
                               "        VALDOL01,  " & _
                               "        ROUND( VALMER01*TIPCAM01*factmo01) AS VALORCOMERCIAL " & _
                               " FROM SSDAGI01 inner join c01refer on refe01=refcia01 " & _
                               "               inner join ssprov22 on  cvepro01=cvepro22 " & _
                               " WHERE FIRMAE01 <> '' AND REFCIA01='"& strtxtRefer & "'"
                   end if
                   if strTipoPedimento2  = "2" then
                       strSQL= " SELECT REFCIA01," & _
                               "        NUMPED01," & _
                               "        patent01," & _
                               "        CVEADU01," & _
                               "        CVESEC01," & _
                               "        'EXP' AS TIPOPE01 ," & _
                               "        CVEPED01," & _
                               "        REGIME01," & _
                               "        zonaod01 as Destino_Origen," & _
                               "        TIPCAM01," & _
                               "        pesobr01," & _
                               "        cvemtr01," & _
                               "        cvemta01," & _
                               "        cvemts01," & _
                               "        CVECLI01," & _
                               "        valseg01 as valorseguros," & _
                               "        segros01 as seguros," & _
                               "        fletes01 as fletes," & _
                               "        embala01 as embalajes," & _
                               "        incble01 as otrosIncrementables," & _
                               "        FECPRE01 as entrada," & _
                               "        FECPAG01," & _
                               "        CVEPRO01,  " & _
                               "        FDSP01 as despacho," & _
                               "        fpro01 as embarque," & _
                               "        FPRE01 as Previo," & _
                               "        FREV01 as Revalidacion," & _
                               "        irspro22 taxidProv, " & _
                               "        nomcli01, " & _
                               "        cvepfm01,  " & _
                               "        VALDOL01,  " & _
                               "        0 AS VALORCOMERCIAL " & _
                               " FROM SSDAGE01 inner join c01refer on refe01=refcia01 " & _
                               "               inner join ssprov22 on  cvepro01=cvepro22 " & _
                               " WHERE FIRMAE01 <> '' AND REFCIA01='"& strtxtRefer & "'"
                   end if
          else
              if strTipoFiltro  = "xpedimento" then 'por un pedimento especifico
                   if strTipoPedimento3 = "1" then
                       strSQL= " SELECT REFCIA01," & _
                               "        NUMPED01," & _
                               "        patent01," & _
                               "        CVEADU01," & _
                               "        CVESEC01," & _
                               "        'IMP' AS TIPOPE01 ," & _
                               "        CVEPED01," & _
                               "        REGIME01," & _
                               "        zonaod01 as Destino_Origen," & _
                               "        TIPCAM01," & _
                               "        pesobr01," & _
                               "        cvemtr01," & _
                               "        cvemta01," & _
                               "        cvemts01," & _
                               "        CVECLI01," & _
                               "        valseg01 as valorseguros," & _
                               "        segros01 as seguros," & _
                               "        fletes01 as fletes," & _
                               "        embala01 as embalajes," & _
                               "        incble01 as otrosIncrementables," & _
                               "        FECENT01," & _
                               "        FECPAG01," & _
                               "        CVEPRO01,  " & _
                               "        FDSP01 as despacho," & _
                               "        fpro01 as embarque," & _
                               "        FPRE01 as Previo," & _
                               "        FREV01 as Revalidacion," & _
                               "        irspro22 taxidProv, " & _
                               "        nomcli01, " & _
                               "        cvepfm01,  " & _
                               "        VALDOL01,  " & _
                               "        ROUND( VALMER01*TIPCAM01*factmo01) AS VALORCOMERCIAL " & _
                               " FROM SSDAGI01 inner join c01refer on refe01=refcia01 " & _
                               "               inner join ssprov22 on  cvepro01=cvepro22 " & _
                               " WHERE FIRMAE01 <> '' AND NUMPED01='"& strtxtPed &"'"
                   end if
                   if strTipoPedimento3  = "2" then
                       strSQL= " SELECT REFCIA01," & _
                               "        NUMPED01," & _
                               "        patent01," & _
                               "        CVEADU01," & _
                               "        CVESEC01," & _
                               "        'EXP' AS TIPOPE01 ," & _
                               "        CVEPED01," & _
                               "        REGIME01," & _
                               "        zonaod01 as Destino_Origen," & _
                               "        TIPCAM01," & _
                               "        pesobr01," & _
                               "        cvemtr01," & _
                               "        cvemta01," & _
                               "        cvemts01," & _
                               "        CVECLI01," & _
                               "        valseg01 as valorseguros," & _
                               "        segros01 as seguros," & _
                               "        fletes01 as fletes," & _
                               "        embala01 as embalajes," & _
                               "        incble01 as otrosIncrementables," & _
                               "        FECPRE01 as entrada," & _
                               "        FECPAG01," & _
                               "        CVEPRO01,  " & _
                               "        FDSP01 as despacho," & _
                               "        fpro01 as embarque," & _
                               "        FPRE01 as Previo," & _
                               "        FREV01 as Revalidacion," & _
                               "        irspro22 taxidProv, " & _
                               "        nomcli01, " & _
                               "        cvepfm01,  " & _
                               "        VALDOL01,  " & _
                               "        0 AS VALORCOMERCIAL " & _
                               " FROM SSDAGE01 inner join c01refer on refe01=refcia01 " & _
                               "               inner join ssprov22 on  cvepro01=cvepro22 " & _
                               " WHERE FIRMAE01 <> '' AND NUMPED01='"& strtxtPed &"'"
                   end if

              end if
          end if
      end if

      'Response.Write("permi")
      'Response.Write(permi)
      'Response.Write("<BR>permi2")
      'Response.Write(permi2)
      'Response.Write(strSQL)
      'Response.End

      '******************************************************************
       'codflete = ""
       'codnoflete = ""
       'intcontflete = 1
       'Set RsbuscaFlete = Server.CreateObject("ADODB.Recordset")
       'RsbuscaFlete.ActiveConnection = MM_EXTRANET_STRING
       'strSQLcodfle = " select clav21 " & _
       '         " from c21paghe "&_
       '         " where desc21 like '%FLETE%' "
       'RsbuscaFlete.Source = strSQLcodfle
       'RsbuscaFlete.CursorType = 0
       'RsbuscaFlete.CursorLocation = 2
       'RsbuscaFlete.LockType = 1
       'RsbuscaFlete.Open()
       'if not RsbuscaFlete.EOF then
       '    While NOT RsbuscaFlete.EOF
       '     if intcontflete = 1 then
       '       codflete   =  "conc21=" & RsbuscaFlete.Fields.Item("clav21").Value
       '       codnoflete =  "conc21<>" & RsbuscaFlete.Fields.Item("clav21").Value
       '     else
       '       codflete   =  codflete & " OR " & "conc21=" & RsbuscaFlete.Fields.Item("clav21").Value
       '       codnoflete =  codnoflete & " and " & "conc21<>" & RsbuscaFlete.Fields.Item("clav21").Value
       '     end if
       '     RsbuscaFlete.movenext
       '     intcontflete = intcontflete + 1
       '    Wend
       'end if
       'RsbuscaFlete.close
       'set  RsbuscaFlete = nothing
       'codflete = "(" & codflete & ")"

       'Response.Write(codflete)
       'Response.Write("<br><br><br>")
       'Response.Write(codnoflete)
       'Response.End
      '******************************************************************

      'else
      '   if strTipoFiltro  = "Descripcion" then
      '    if strTipoPedimento2  = "1" then
      '     tmpTipo = "IMPORTACION"
      '     strSQL = "SELECT distinct adusec01,ifnull(valdol01,0) as valdol01,tipopr01,ifnull(valmer01,0) as valmer01,ifnull(factmo01,0) as factmo01, ifnull(p_dta101,0) as p_dta101, ifnull(t_reca01,0) as t_reca01, ifnull(i_dta101,0) as i_dta101, cvecli01, refcia01, fecpag01, ifnull(valfac01,0) as valfac01, fletes01, segros01, ifnull(cvepvc01,'0') as cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecent01,anexol01 FROM ssdagi01,ssfrac02 WHERE refcia01 = refcia02 and d_mer102 like '%" & strDescripcion & "%' and fecpag01 >='"&FormatoFechaInv(strDateIni2)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
      '    ' Response.Write(strSQL)
      ' '    Response.End
      '   end if
      '   if strTipoPedimento2  = "2" then
      '     tmpTipo = "EXPORTACION"
      '     strSQL = "SELECT distinct adusec01,ifnull(valdol01,0) as valdol01,tipopr01, ifnull(factmo01,0) as factmo01, ifnull(p_dta101,0) AS p_dta101, ifnull(t_reca01,0) as t_reca01, ifnull(i_dta101,0) as i_dta101, cvecli01, refcia01, fecpag01, ifnull(valfac01,0) as valfac01, ifnull(fletes01,0) as fletes01, ifnull(segros01,0) as segros01, cvepvc01, tipcam01, patent01, numped01, totbul01, cveped01, cveadu01, desf0101, nompro01, cvepod01, nombar01, tipopr01, fecpre01,anexol01 FROM ssdage01,ssfrac02 WHERE refcia01 = refcia02 and d_mer102 like '%" & strDescripcion & "%' and fecpag01 >='"&FormatoFechaInv(strDateIni2)&"' AND fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & Permi & " and firmae01 !='' order by refcia01"
      '   end if
      '   end if
      'end if

      'strRefcia  = ""
      'strCgas    = ""
      'strFecCgas = ""
      'strSuph    = ""
      'strCoad    = ""
      'strCsce    = ""
      'strChon    = ""
      'strIva     = ""
      'strPiva    = ""
      'strAnti    = ""
      'strSald    = ""
      'strCveCli  = ""
      'strAduana  = ""
      'strSeccion = ""
      'strPatente = ""
      'strNumped  = ""
      'strEmpresa = ""
      'strNomcli  = ""
      'strtotal   = ""
      'strimpuestosPed   = ""
      'strFlete   = ""
      'strOtros   = ""

               STRREFCIARP               = ""
               STRNUMPEDRP               = ""
               STRpatentRP               = ""
               STRCVEADURP               = ""
               STRCVESECRP               = ""
               STRTIPOPERP               = ""
               STRCVEPEDRP               = ""
               STRREGIMERP               = ""
               STRDestino_ORP            = ""
               STRTIPCAMRP               = ""
               STRpesobrRP               = ""
               STRcvemtrRP               = ""
               STRcvemtaRP               = ""
               STRcvemtsRP               = ""
               STRCVECLIRP               = ""
               STRvalorsegRP             = ""
               STRsegurosRP              = ""
               STRfletesRP               = ""
               STRembalajesRP            = ""
               STRotrosIncRP             = ""
               STRFECENTRP               = ""
               STRFECPAGRP               = ""
               STRCVEPRORP               = ""
               STRdespachoRP             = ""
               STRembarqueRP             = ""
               STRPrevioRP               = ""
               STRRevalidacionRP         = ""

               STRValorDolaresRP         = ""
               STRValorAduanaRP          = ""
               STRValorComercialRP       = ""
               STRBLSRP                  = ""
               STRIGIRP                  = ""
               STRIVARP                  = ""
               STRDTARP                  = ""
               STRPRVRP                  = ""
               STRMultasRP               = ""
               STRRecargosRP             = ""
               STRCuotascompRP           = ""
               STRArticulo_303RP         = ""
               STRRTRP                   = ""
               STROtrosRP                = ""
               STRISANRP                 = ""
               STRIEPSRP                 = ""
               STRFPIGIRP                = ""
               STRFPIVARP                = ""
               STRFPDTARP                = ""
               STRFPPRVRP                = ""
               STRFPMultasRP             = ""
               STRFPRecargosRP           = ""
               STRFPCuotascompRP         = ""
               STRFPArticulo_303RP       = ""
               STRFPRTRP                 = ""
               STRFPOtrosRP              = ""
               STRFPISANRP               = ""
               STRFPIEPSRP               = ""
               STRImporteTotalRP         = ""
               STRPro_TAXIDRP            = ""
               STRPro_VinculacionRP      = ""
               STRGuiaMasterRP           = ""
               STRGuiaHouseRP            = ""
               STRTalonFleteTransporteRP = ""
               STRObservacionesRP        = ""
               STRTransportistaRP        = ""
               STRConsolidadorRP         = ""
               STRMonedaRP               = ""
               STRRectiR1_A3RP           = ""
               STRREctiPedOriRP          = ""
               STRRFCRP                  = ""
               strEmpresaPR              = ""

      'Response.End
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

                     STRREFCIARP               = rsCGPrincipal.Fields.Item("REFCIA01").Value
                     STRNUMPEDRP               = rsCGPrincipal.Fields.Item("NUMPED01").Value
                     STRpatentRP               = rsCGPrincipal.Fields.Item("patent01").Value
                     STRCVEADURP               = rsCGPrincipal.Fields.Item("CVEADU01").Value
                     STRCVESECRP               = rsCGPrincipal.Fields.Item("CVESEC01").Value
                     STRTIPOPERP               = rsCGPrincipal.Fields.Item("TIPOPE01").Value
                     STRCVEPEDRP               = rsCGPrincipal.Fields.Item("CVEPED01").Value
                     STRREGIMERP               = rsCGPrincipal.Fields.Item("REGIME01").Value
                     STRDestino_ORP            = rsCGPrincipal.Fields.Item("Destino_Origen").Value
                     STRTIPCAMRP               = rsCGPrincipal.Fields.Item("TIPCAM01").Value
                     STRpesobrRP               = rsCGPrincipal.Fields.Item("pesobr01").Value
                     STRcvemtrRP               = rsCGPrincipal.Fields.Item("cvemtr01").Value
                     STRcvemtaRP               = rsCGPrincipal.Fields.Item("cvemta01").Value
                     STRcvemtsRP               = rsCGPrincipal.Fields.Item("cvemts01").Value
                     STRCVECLIRP               = rsCGPrincipal.Fields.Item("nomcli01").Value
                     STRvalorsegRP             = rsCGPrincipal.Fields.Item("valorseguros").Value
                     STRsegurosRP              = rsCGPrincipal.Fields.Item("seguros").Value
                     STRfletesRP               = rsCGPrincipal.Fields.Item("fletes").Value
                     STRembalajesRP            = rsCGPrincipal.Fields.Item("embalajes").Value
                     STRotrosIncRP             = rsCGPrincipal.Fields.Item("otrosIncrementables").Value
                     STRFECENTRP               = rsCGPrincipal.Fields.Item("entrada").Value
                     STRFECPAGRP               = rsCGPrincipal.Fields.Item("FECPAG01").Value
                     STRCVEPRORP               = rsCGPrincipal.Fields.Item("CVEPRO01").Value
                     STRdespachoRP             = rsCGPrincipal.Fields.Item("despacho").Value
                     STRembarqueRP             = rsCGPrincipal.Fields.Item("embarque").Value
                     STRPrevioRP               = rsCGPrincipal.Fields.Item("Previo").Value
                     STRRevalidacionRP         = rsCGPrincipal.Fields.Item("Revalidacion").Value
                     STRPro_TAXIDRP            = rsCGPrincipal.Fields.Item("taxidProv").Value
                     STRMonedaRP               = rsCGPrincipal.Fields.Item("cvepfm01").Value
                     STRValorDolaresRP         = rsCGPrincipal.Fields.Item("VALDOL01").Value 'Valor Dolares
                     STRValorComercialRP       = rsCGPrincipal.Fields.Item("VALORCOMERCIAL").Value 'Valor Comercial

                     STRValorAduanaRP          = ""
                     STRIGIRP                  = ""
                     STRIVARP                  = ""
                     STRDTARP                  = ""
                     STRPRVRP                  = ""
                     STRMultasRP               = ""
                     STRRecargosRP             = ""
                     STRCuotascompRP           = ""
                     STRArticulo_303RP         = ""
                     STRRTRP                   = ""
                     STROtrosRP                = ""
                     STRISANRP                 = ""
                     STRIEPSRP                 = ""
                     STRFPIGIRP                = ""
                     STRFPIVARP                = ""
                     STRFPDTARP                = ""
                     STRFPPRVRP                = ""
                     STRFPMultasRP             = ""
                     STRFPRecargosRP           = ""
                     STRFPCuotascompRP         = ""
                     STRFPArticulo_303RP       = ""
                     STRFPRTRP                 = ""
                     STRFPOtrosRP              = ""
                     STRFPISANRP               = ""
                     STRFPIEPSRP               = ""
                     STRImporteTotalRP         = ""

                     STRBLSRP                  = ""
                     STRPro_VinculacionRP      = ""
                     strEmpresaPR              = ""

                     STRRectiR1_A3RP           = ""
                     STRREctiPedOriRP          = ""

                     'STRValorDolaresRP         = ""
                     'STRValorAduanaRP          = ""
                     'STRValorComercialRP       = ""
                     'STRBLSRP                  = ""

                     'STRPro_VinculacionRP       = ""
                     'STRGuiaMasterRP           = ""
                     'STRGuiaHouseRP            = ""
                     'STRTalonFleteTransporteRP = ""
                     'STRObservacionesRP        = ""
                     'STRTransportistaRP        = ""
                     'STRConsolidadorRP         = ""
                     'STRMonedaRP               = ""
                     'STRRectiR1_A3RP           = ""
                     'STRREctiPedOriRP          = ""
                     'STRRFCRP                  = ""

                     '*******************************************************************
                     strSQL1 = " SELECT refcia36 Referencia,                                               " & _
                               "        sum(IF( cveimp36='1'  , import36,  0 ) )  as dta,                  " & _
                               "        sum(IF( cveimp36='1'  , fpagoi36,  0 ) )  as formapagodta,         " & _
                               "        sum(IF( cveimp36='2'  , import36,  0 ) )  as CC,                   " & _
                               "        sum(IF( cveimp36='2'  , fpagoi36,  0 ) )  as formapagoCC,          " & _
                               "        sum(IF( cveimp36='12' , import36,  0 ) )  as articulo303,          " & _
                               "        sum(IF( cveimp36='12' , fpagoi36,  0 ) )  as formapagoarticulo303, " & _
                               "        sum(IF( cveimp36='13' , import36,  0 ) )  as RT,                   " & _
                               "        sum(IF( cveimp36='13' , fpagoi36,  0 ) )  as formapagoRT,          " & _
                               "        sum(IF( cveimp36='4'  , import36,  0 ) )  as ISAN,                 " & _
                               "        sum(IF( cveimp36='4'  , fpagoi36,  0 ) )  as formapagoISAN,        " & _
                               "        sum(IF( cveimp36='5'  , import36,  0 ) )  as IEPS,                 " & _
                               "        sum(IF( cveimp36='5'  , fpagoi36,  0 ) )  as formapagoIEPS,        " & _
                               "        sum(IF( cveimp36='3'  , import36,  0 ) )  as iva,                  " & _
                               "        sum(IF( cveimp36='3'  , fpagoi36,  0 ) )  as formapagoiva,         " & _
                               "        sum(IF( cveimp36='6'  , import36,  0 ) )  as igi,                  " & _
                               "        sum(IF( cveimp36='6'  , fpagoi36,  0 ) )  as formapagoigi,         " & _
                               "        sum(IF( cveimp36='15' , import36,  0 ) )  as prev,                 " & _
                               "        sum(IF( cveimp36='15' , fpagoi36,  0 ) )  as formapagoprev,        " & _
                               "        sum(IF( cveimp36='7'  , import36,  0 ) )  as recargo,              " & _
                               "        sum(IF( cveimp36='7'  , fpagoi36,  0 ) )  as formapagorecargo,     " & _
                               "        sum(IF( cveimp36='11' , import36,  0 ) )  as multa,                " & _
                               "        sum(IF( cveimp36='11' , fpagoi36,  0 ) )  as formapagomulta,       " & _
                               "        sum(IF( cveimp36<>'1' and cveimp36<>'2' and cveimp36<>'12' and cveimp36<>'13' and cveimp36<>'4' and cveimp36<>'5' and cveimp36<>'3' and cveimp36<>'6' and cveimp36<>'7' and cveimp36<>'11' and cveimp36<>'15', import36,   0 ) )  as otros,      " & _
                               "        sum(IF( cveimp36<>'1' and cveimp36<>'2' and cveimp36<>'12' and cveimp36<>'13' and cveimp36<>'4' and cveimp36<>'5' and cveimp36<>'3' and cveimp36<>'6' and cveimp36<>'7' and cveimp36<>'11' and cveimp36<>'15' , fpagoi36,  0 ) )  as formaotros, " & _
                               "        sum(import36) Monto " & _
                               " FROM SSCONT36              " & _
                               " WHERE REFCIA36 = '" & STRREFCIARP  & "' " & " group by refcia36"
                     'Response.Write(strSQL1)
                     'Response.End
                     Set rsCuadroLiq = Server.CreateObject("ADODB.Recordset")
                     rsCuadroLiq.ActiveConnection = MM_EXTRANET_STRING
                     rsCuadroLiq.Source = strSQL1
                     rsCuadroLiq.CursorType = 0
                     rsCuadroLiq.CursorLocation = 2
                     rsCuadroLiq.LockType = 1
                     rsCuadroLiq.Open()
                     if not rsCuadroLiq.eof then
                         'While NOT rsCuadroLiq.EOF
                                 STRIGIRP             = rsCuadroLiq.Fields.Item("igi").Value
                                 STRIVARP             = rsCuadroLiq.Fields.Item("iva").Value
                                 STRDTARP             = rsCuadroLiq.Fields.Item("dta").Value
                                 STRPRVRP             = rsCuadroLiq.Fields.Item("prev").Value
                                 STRMultasRP          = rsCuadroLiq.Fields.Item("multa").Value
                                 STRRecargosRP        = rsCuadroLiq.Fields.Item("recargo").Value
                                 STRCuotascompRP      = rsCuadroLiq.Fields.Item("CC").Value
                                 STRArticulo_303RP    = rsCuadroLiq.Fields.Item("articulo303").Value
                                 STRRTRP              = rsCuadroLiq.Fields.Item("RT").Value
                                 STROtrosRP           = rsCuadroLiq.Fields.Item("otros").Value
                                 STRISANRP            = rsCuadroLiq.Fields.Item("ISAN").Value
                                 STRIEPSRP            = rsCuadroLiq.Fields.Item("IEPS").Value
                                 STRFPIGIRP           = rsCuadroLiq.Fields.Item("formapagoigi").Value
                                 STRFPIVARP           = rsCuadroLiq.Fields.Item("formapagoiva").Value
                                 STRFPDTARP           = rsCuadroLiq.Fields.Item("formapagodta").Value
                                 STRFPPRVRP           = rsCuadroLiq.Fields.Item("formapagoprev").Value
                                 STRFPMultasRP        = rsCuadroLiq.Fields.Item("formapagomulta").Value
                                 STRFPRecargosRP      = rsCuadroLiq.Fields.Item("formapagorecargo").Value
                                 STRFPCuotascompRP    = rsCuadroLiq.Fields.Item("formapagoCC").Value
                                 STRFPArticulo_303RP  = rsCuadroLiq.Fields.Item("formapagoarticulo303").Value
                                 STRFPRTRP            = rsCuadroLiq.Fields.Item("formapagoRT").Value
                                 STRFPOtrosRP         = rsCuadroLiq.Fields.Item("formaotros").Value
                                 STRFPISANRP          = rsCuadroLiq.Fields.Item("formapagoISAN").Value
                                 STRFPIEPSRP          = rsCuadroLiq.Fields.Item("formapagoIEPS").Value
                                 STRImporteTotalRP    = rsCuadroLiq.Fields.Item("Monto").Value
                            'rsCuadroLiq.movenext
                         'Wend
                     end if
                     rsCuadroLiq.close
                     set rsCuadroLiq = Nothing
                     '******************************************************************
                             Set Recguia = Server.CreateObject("ADODB.Recordset")
                             Recguia.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT  IF( IDNGUI04=1,numgui04,'') AS guiaMaster,  " & _
                                          "         IF( IDNGUI04=2,numgui04,'') AS guiaHouse    " & _
                                          " from ssguia04  " & _
                                          " where refcia04='" & ltrim(STRREFCIARP)&"'"
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

                             STRGuiaMasterRP =  strGuiaMaster
                             STRGuiaHouseRP  =  strGuiaMasterHouse

                             if STRGuiaHouseRP = "" then
                                 STRBLSRP  =  STRGuiaMasterRP
                             else
                                 STRBLSRP  =  STRGuiaMasterRP & "; "& STRGuiaHouseRP
                             end if
                     '******************************************************************
                             Set RsVicul = Server.CreateObject("ADODB.Recordset")
                             RsVicul.ActiveConnection = MM_EXTRANET_STRING
                             strSqlPrv =  " select REFCIA39, " & _
                                          "        vincul39  " & _
                                          " from ssfact39    " & _
                                          " where refcia39  ='" & ltrim(STRREFCIARP)&"'" & "  AND CVEPRO39 = "& STRCVEPRORP
                             'Response.Write(strSqlPrv)
                             'Response.End
                             RsVicul.Source = strSqlPrv
                             RsVicul.CursorType = 0
                             RsVicul.CursorLocation = 2
                             RsVicul.LockType = 1
                             RsVicul.Open()
                             if not RsVicul.eof then
                                 'While NOT RsVicul.EOF
                                    if RsVicul.Fields.Item("vincul39").Value = 1 then
                                       STRPro_VinculacionRP = "YES"
                                    else
                                       STRPro_VinculacionRP = "NO"
                                    end if
                                 '   RsVicul.movenext
                                 'Wend
                             end if
                             RsVicul.close
                             set RsVicul = Nothing
                     '******************************************************************
                     if STRCVEPEDRP = "A3" OR STRCVEPEDRP = "R1" then
                                 Set Rsrecti = Server.CreateObject("ADODB.Recordset")
                                 Rsrecti.ActiveConnection = MM_EXTRANET_STRING
                                 strSqlrecti =  " SELECT pedorg06 " & _
                                                " FROM SSRECP06 " & _
                                                " WHERE REFCIA06 ='" & STRREFCIARP&"'"
                                 'Response.Write(strSqlrecti)
                                 'Response.End
                                 Rsrecti.Source = strSqlrecti
                                 Rsrecti.CursorType = 0
                                 Rsrecti.CursorLocation = 2
                                 Rsrecti.LockType = 1
                                 Rsrecti.Open()
                                 if not Rsrecti.eof then
                                     'While NOT Rsrecti.EOF
                                        STRRectiR1_A3RP  = "YES"
                                        STRREctiPedOriRP = Rsrecti.Fields.Item("pedorg06").Value
                                     '   Rsrecti.movenext
                                     'Wend
                                 end if
                                 Rsrecti.close
                                 set Rsrecti = Nothing
                     end if
                     '******************************************************************
                             'STRValorAduanaRP
                             Set RsValAdu = Server.CreateObject("ADODB.Recordset")
                             RsValAdu.ActiveConnection = MM_EXTRANET_STRING
                             strSqlValAdu =  " SELECT sum(vaduan02) as valoraduana " & _
                                          " FROM `ssfrac02`  " & _
                                          " where refcia02 ='" & ltrim(STRREFCIARP) & "'" & _
                                          " order by refcia02 "
                             'Response.Write(strSqlValAdu)
                             'Response.End
                             RsValAdu.Source = strSqlValAdu
                             RsValAdu.CursorType = 0
                             RsValAdu.CursorLocation = 2
                             RsValAdu.LockType = 1
                             RsValAdu.Open()
                             if not RsValAdu.eof then
                                 'While NOT RsValAdu.EOF
                                    STRValorAduanaRP = RsValAdu.Fields.Item("valoraduana").Value
                                 '   RsValAdu.movenext
                                 'Wend
                             end if
                             RsValAdu.close
                             set RsValAdu = Nothing
                     '******************************************************************
                             if InStr(STRCVECLIRP ,"CASE") > 0 then
                                strEmpresaPR = "CASE"
                             else
                                if InStr(STRCVECLIRP ,"CNH COMPONENTES") > 0 then
                                    strEmpresaPR = "CNHCMP"
                                else
                                   if InStr(STRCVECLIRP ,"CNH COMERCIAL") > 0 then
                                      strEmpresaPR = "CNHCOM"
                                   else
                                       if InStr(STRCVECLIRP ,"CNH INDUSTRIAL") > 0 then
                                          strEmpresaPR = "CNHIND"
                                       else
                                           if InStr(STRCVECLIRP ,"CNH DE MEXICO") > 0 then
                                              strEmpresaPR = "CNHMEX"
                                           else
                                             if InStr(STRCVECLIRP ,"CNH SERVICIOS CORPORATIVOS") > 0 then
                                                strEmpresaPR = "CNHSER"
                                             else
                                               if InStr(STRCVECLIRP ,"NEW HOLLAND") > 0 then
                                                  strEmpresaPR = "NHMEX"
                                               else
                                                  strEmpresaPR = STRCVECLIRP
                                               end  if
                                             end  if
                                           end  if
                                       end  if
                                   end  if
                                end  if
                             end  if
                     '******************************************************************



                     'SELECT VALDOL01,VALMER01, VALMER01*TIPCAM01 AS VALORCOMERCIAL
                     'FROM SSDAGI01
                     'WHERE REFCIA01 = 'DAI08-1233'

                      'SELECT refcia02,valdls02, vaduan02, vmerme02
                      'FROM `ssfrac02`
                      'where refcia02 = 'DAI08-1233'



                        '                     Set RsValor1 = Server.CreateObject("ADODB.Recordset")
                        '                     RsValor1.ActiveConnection = MM_EXTRANET_STRING
                        '                     'strSQL = " select refe21,"&_
                        '                     '        "        sum( IF( conc21=1  , d.mont21,   0 )   )  as impuestos,"&_
                        '                     '        "        sum( IF( ((conc21=3 OR conc21=7 OR conc21=19 OR conc21=74 OR conc21=104 OR conc21=110 OR conc21=145)), d.mont21,   0 )   )  as fletes, "&_
                        '                     '        "        sum( IF( conc21<>1 and conc21<>3 and conc21<>7 and conc21<>19 and conc21<>74 and conc21<>104 and conc21<>110 and conc21<>145  , d.mont21,   0 )   )  as otros "&_
                        '                     '        " from d21paghe as d, e21paghe as e  "&_
                        '                     '        " where e.foli21 = d.foli21 and     "&_
                        '                     '        "       e.fech21 = d.fech21 and  "&_
                        '                     '        "       (e.esta21 = 'A'  or e.esta21='E' ) and   "&_
                        '                     '        "       d.refe21 = 'DAI08-0956' "&_
                        '                     '        " group by refe21 "
                        '                     strSQLph =" select refe21,"&_
                        '                             "        sum(IF(conc21=1, d.mont21,0 ))  as impuestos,"&_
                        '                             "        sum(IF("& codflete &", d.mont21, 0 ))  as fletes, "&_
                        '                             "        sum(IF(conc21<>1 and " & codnoflete & ", d.mont21,0)) as otros "&_
                        '                             " from d21paghe as d, e21paghe as e  "&_
                        '                             " where e.foli21 = d.foli21 and     "&_
                        '                             "       e.fech21 = d.fech21 and  "&_
                        '                             "       (e.esta21 = 'A'  or e.esta21='E' ) and   "&_
                        '                             "       d.refe21 = '" &strRefcia& "' "&_
                        '                             " group by refe21 "
                        '                     'Response.Write(strSQLph)
                        '                     'Response.End
                        '                     RsValor1.Source = strSQLph
                        '                     RsValor1.CursorType = 0
                        '                     RsValor1.CursorLocation = 2
                        '                     RsValor1.LockType = 1
                        '                     RsValor1.Open()
                        '                     if not RsValor1.EOF then
                        '                         strimpuestosPed = RsValor1.Fields.Item("impuestos").Value
                        '                         strFlete        = RsValor1.Fields.Item("fletes").Value
                        '                         strOtros        = RsValor1.Fields.Item("otros").Value
                        '                     end if
                        '                     RsValor1.close
                        '                     set  RsValor1 = nothing
                        '                     '******************************************************************

                        '                     '******************************************************************
                        '                     montofleteAux2 = 0
                        '                     montoSvrComp   = 0
                        '                     Set RsValor2 = Server.CreateObject("ADODB.Recordset")
                        '                     RsValor2.ActiveConnection = MM_EXTRANET_STRING
                        '                     strSQLsvr = " SELECT refe32,ttar32, dcrp32,mont32" &_
                        '                                 " FROM d32rserv  " &_
                        '                                 " WHERE REFE32 = '" &strRefcia& "'"
                        '                     'Response.Write(strSQLsvr)
                        '                     'Response.End
                        '                     RsValor2.Source = strSQLsvr
                        '                     RsValor2.CursorType = 0
                        '                     RsValor2.CursorLocation = 2
                        '                     RsValor2.LockType = 1
                        '                     RsValor2.Open()
                        '                     if not RsValor1.EOF then
                        '                         While NOT RsValor2.EOF
                        '                           if (RsValor2.Fields.Item("ttar32").Value = "00033") then
                        '                              montofleteAux2  = montofleteAux2 + RsValor2.Fields.Item("mont32").Value
                        '                           else
                        '                              montoSvrComp  = montoSvrComp + RsValor2.Fields.Item("mont32").Value
                        '                           end if
                        '                           RsValor2.movenext
                        '                         Wend
                        '                     end if
                        '                     RsValor2.close
                        '                     set  RsValor2 = nothing
                        '
                        '                     strFlete = strFlete + montofleteAux2
                        '
                        '                     '******************************************************************







                        '                         'Response.Write("<BR>")
                        '                         Response.Write( strEmpresa )
                        '                         Response.Write(",")
                        '                         Response.Write( strAduana )
                        '                         Response.Write(",")
                        '                         Response.Write( strSeccion )
                        '                         Response.Write(",")
                        '                         Response.Write( strPatente )
                        '                         Response.Write(",")
                        '                         Response.Write( strNumped )
                        '                         Response.Write(",")
                        '                         Response.Write( strRefcia )
                        '                         Response.Write(",")
                        '                         Response.Write(  strCgas  )
                        '                         Response.Write(",")
                        '                         Response.Write( strFecCgas )
                        '                         Response.Write(",")
                        '                         Response.Write("PS" )
                        '                         Response.Write(",")
                        '                         Response.Write(strimpuestosPed) 'Impuestos del pedimento
                        '                         Response.Write(",")
                        '                         Response.Write("0" ) 'Franja amarilla
                        '                         Response.Write(",")
                        '                         Response.Write( strChon ) ' Honorarios
                        '                         Response.Write(",")
                        '                         Response.Write(strFlete) 'Fletes
                        '                         Response.Write(",")
                        '                         Response.Write(montoSvrComp ) 'Servicios complementarios
                        '                         Response.Write(",")
                        '                         Response.Write(strOtros) 'Otros
                        '                         Response.Write(",")
                        '                         Response.Write("0" ) 'valor cuenta USA
                        '                         Response.Write(",")
                        '                         Response.Write( strIva ) 'IVA
                        '                         Response.Write(",")
                        '                         Response.Write( strtotal )
                        '                         Response.Write(",")  'total de cuenta de gastos
                        '
                        '                         Response.Write( "0" )
                        '                         Response.Write(",")
                        '                         Response.Write( "0" )
                        '                         Response.Write(",")
                        '                         Response.Write( "" )
                        '                         Response.Write(",")
                        '                         Response.Write( "0" )
                        '                         Response.Write(",")
                        '                         Response.Write( "0" )
                        '                         Response.Write(",")
                        '                         Response.Write( "0" )

                                                 'Response.Write( strSuph )
                                                 'Response.Write(",")
                                                 'Response.Write( strCoad )
                                                 'Response.Write(",")
                                                 'Response.Write( strCsce )
                                                 'Response.Write(",")
                                                 'Response.Write( strAnti )
                                                 'Response.Write(",")
                                                 'Response.Write( strSald )

                         Response.Write("0") 'Paginas del Pedimento
                         Response.Write(",")
                         Response.Write(STRREFCIARP) 'Referencia
                         Response.Write(",")
                         Response.Write(STRNUMPEDRP) ' Pedimento
                         Response.Write(",")
                         Response.Write(STRpatentRP)  'Patente
                         Response.Write(",")
                         Response.Write(STRCVEADURP)  'Aduana
                         Response.Write(",")
                         Response.Write(STRCVESECRP)  'Sección
                         Response.Write(",")
                         Response.Write(STRTIPOPERP)  'Tipo de Operación
                         Response.Write(",")
                         Response.Write(STRCVEPEDRP)  'Clave de Pedimento
                         Response.Write(",")
                         Response.Write(STRREGIMERP)  'Regimen
                         Response.Write(",")
                         Response.Write(STRDestino_ORP) 'Destino Origen
                         Response.Write(",")
                         Response.Write(STRTIPCAMRP)    'Tipo de Cambio
                         Response.Write(",")
                         Response.Write(STRpesobrRP)    'Peso Total
                         Response.Write(",")
                         Response.Write(STRcvemtrRP)    'Transporte (E/S)
                         Response.Write(",")
                         Response.Write(STRcvemtaRP)    'Transporte (Arribo)
                         Response.Write(",")
                         Response.Write(STRcvemtsRP)    'Transporte (Salida)
                         Response.Write(",")
                         Response.Write(STRValorDolaresRP) 'Valor Dolares
                         Response.Write(",")
                         Response.Write(STRValorAduanaRP)  'Valor Aduana (valor factura)
                         Response.Write(",")
                         Response.Write(STRValorComercialRP) 'Valor Comercial
                         Response.Write(",")
                         Response.Write(strEmpresaPR)  'Empresa
                         Response.Write(",")
                         'Response.Write(STRCVECLIRP)
                         Response.Write(STRvalorsegRP) 'Valor Seguros
                         Response.Write(",")
                         Response.Write(STRsegurosRP)  'Seguro
                         Response.Write(",")
                         Response.Write(STRfletesRP)  'Flete
                         Response.Write(",")
                         Response.Write(STRembalajesRP) 'Embalajes
                         Response.Write(",")
                         Response.Write(STRotrosIncRP)  'Incrementables
                         Response.Write(",")
                         Response.Write(STRBLSRP) 'BLS
                         Response.Write(",")
                         Response.Write(STRFECENTRP) 'Fecha de Entrada
                         Response.Write(",")
                         Response.Write(STRFECPAGRP)  'Fecha de pago
                         Response.Write(",")
                         Response.Write(STRIGIRP)  'Concepto (IGI)
                         Response.Write(",")
                         Response.Write(STRIVARP)  'Concepto (IVA)
                         Response.Write(",")
                         Response.Write(STRDTARP)  'Concepto (DTA)
                         Response.Write(",")
                         Response.Write(STRPRVRP)  'Concepto (PRV)
                         Response.Write(",")
                         Response.Write(STRMultasRP) 'Concepto (Multas)
                         Response.Write(",")
                         Response.Write(STRRecargosRP) 'Concepto (Recargos)
                         Response.Write(",")
                         Response.Write(STRCuotascompRP) 'concepto (Cuotas compensatorias)
                         Response.Write(",")
                         Response.Write(STRArticulo_303RP) 'Concepto (Articulo 303)
                         Response.Write(",")
                         Response.Write(STRRTRP) 'Concepto (RT)
                         Response.Write(",")
                         Response.Write(STROtrosRP) 'Concepto (Otros)
                         Response.Write(",")
                         Response.Write(STRISANRP)  'Concepto (ISAN)
                         Response.Write(",")
                         Response.Write(STRIEPSRP)  'Concepto (IEPS)
                         Response.Write(",")
                         Response.Write(STRFPIGIRP) 'FORMA DE PAGO  (IGI)
                         Response.Write(",")
                         Response.Write(STRFPIVARP) 'FORMA DE PAGO  (IVA)
                         Response.Write(",")
                         Response.Write(STRFPDTARP) 'FORMA DE PAGO  (DTA)
                         Response.Write(",")
                         Response.Write(STRFPPRVRP) 'FORMA DE PAGO  (PRV)
                         Response.Write(",")
                         Response.Write(STRFPMultasRP) 'FORMA DE PAGO  (Multas)
                         Response.Write(",")
                         Response.Write(STRFPRecargosRP) 'FORMA DE PAGO  (Recargos)
                         Response.Write(",")
                         Response.Write(STRFPCuotascompRP) 'FORMA DE PAGO  (Cuotas compensatorias)
                         Response.Write(",")
                         Response.Write(STRFPArticulo_303RP) 'FORMA DE PAGO (Articulo 303)
                         Response.Write(",")
                         Response.Write(STRFPRTRP) 'FORMA DE PAGO (RT)
                         Response.Write(",")
                         Response.Write(STRFPOtrosRP) 'FORMA DE PAGO (Otros)
                         Response.Write(",")
                         Response.Write(STRFPISANRP)  'FORMA DE PAGO (ISAN)
                         Response.Write(",")
                         Response.Write(STRFPIEPSRP)  'FORMA DE PAGO (IEPS)
                         Response.Write(",")
                         Response.Write(STRImporteTotalRP) 'Importe Total
                         Response.Write(",")
                         Response.Write(STRPro_TAXIDRP) 'Proveedor (TAX ID)
                         Response.Write(",")
                         Response.Write(STRPro_VinculacionRP) 'Vinculacion
                         Response.Write(",")
                         Response.Write(STRGuiaMasterRP)  'No. Guia (Master)
                         Response.Write(",")
                         Response.Write(STRGuiaHouseRP) 'No. Guia (House)
                         Response.Write(",")
                         Response.Write(STRTalonFleteTransporteRP) 'Talón Flete/No. Transporte
                         Response.Write(",")
                         Response.Write(STRObservacionesRP) 'Observaciones
                         Response.Write(",")
                         Response.Write(STRdespachoRP) 'Fecha Despacho Puerto
                         Response.Write(",")
                         Response.Write(STRembarqueRP) 'Fecha de Embarque
                         Response.Write(",")
                         Response.Write(STRPrevioRP)   'Fecha de Previo
                         Response.Write(",")
                         Response.Write(STRRevalidacionRP) 'Fecha Revalidacion
                         Response.Write(",")
                         Response.Write(STRTransportistaRP)  'Transportista
                         Response.Write(",")
                         Response.Write(STRConsolidadorRP) 'Consolidador
                         Response.Write(",")
                         Response.Write(STRMonedaRP)  'Moneda
                         Response.Write(",")
                         Response.Write(STRRectiR1_A3RP)  'R1 o A3
                         Response.Write(",")
                         Response.Write(STRREctiPedOriRP) 'Pedimento Original
                         Response.Write(",")
                         Response.Write(STRRFCRP)  'RFC

                         'Response.Write(STRCVEPRORP)
                         'Response.Write(",")

                         Response.write vbNewLine

             '        'end if
                 rsCGPrincipal.movenext
                 Wend
             end if
             rsCGPrincipal.close
             set rsCGPrincipal = Nothing
             '*******************************************************************


              'Response.Write("<br>")
              'Response.Write("<br>")
              'Response.Write("<br>")
              'Response.Write("<br>")
              'Response.Write(permi)
              'Response.Write("<br>")
              'Response.Write(permi2)
              'Response.Write("<br>")
              'Response.Write("<br>")

              'Response.Write(strSQL)
              'Response.Write("<br>")
              'Response.Write("<br>")

              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc")

              'Response.write vbNewLine
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc")

              'Response.write vbNewLine
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc")

              'Response.write vbNewLine
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc")

              'Response.write vbNewLine
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc,")
              'Response.Write("Archivo svc")

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







