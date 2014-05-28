
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<%

    Function pd(n, totalDigits)
        if totalDigits > len(n) then
            pd = String(totalDigits-len(n),"0") & n
        else
            pd = n
        end if
    End Function


    Function formatofechaNum(DFecha)
       if isdate( DFecha ) then
          formatofechaNum = YEAR(DFecha) & Pd(Month( DFecha ),2) & Pd(DAY( DFecha ),2)
       else
          formatofechaNum	= DFecha
       end if
    End Function


    Function diasTrimFinSemana(DInicio, DFin)
         x_Dias = 0
         x_Dias = dateDiff("d", DInicio , DFin )


         if x_Dias > 0 then
           x_Con=1
           x_finSemana=0
           Do While (x_Con <= x_Dias)
              x_diasemana=WeekDay( DateAdd("d",x_Con,  DInicio ) )
              if x_diasemana=1 or x_diasemana=7 then
                 x_finSemana = x_finSemana +1
              end if
              x_Con = x_Con + 1
           loop
         x_Dias = x_Dias - x_finSemana ' Restamos los dias de fin de semana
         end if
         diasTrimFinSemana = x_Dias

    End Function


    Function SumarDiasSinFinSemana(DFecha,IntDayAdd)
         x_Dias = 0
         x_Dias = IntDayAdd
         if x_Dias > 0 then
           x_Con=1
           x_finSemana=0
           Do While (x_Con <= x_Dias)
              x_diasemana=WeekDay( DateAdd("d",x_Con,  DFecha ) )
              if x_diasemana=1 or x_diasemana=7 then
                 x_finSemana = x_finSemana +1
              end if
              x_Con = x_Con + 1
           loop
         x_Dias = x_Dias + x_finSemana ' sumamos los dias de fin de semana
         end if
         DNewFecha = DateAdd("d",x_Dias, DFecha  )

         numDia= WeekDay( DNewFecha )
         if numDia=1 then ' domingo
            DNewFecha = DateAdd("d",1, DNewFecha  )
         else
            if numDia=7 then ' Sabado
                DNewFecha = DateAdd("d",2, DNewFecha  )
            end if
         end if
         SumarDiasSinFinSemana =  DNewFecha

    End Function

    Function SumarDias(DFecha,IntDayAdd,intType)
      if isdate(DFecha) then
         if intType = 1 then ' dias Naturales
            SumarDias = DateAdd("d",IntDayAdd,  DFecha )
         else ' dias habiles
            'if intType = 2 then
              SumarDias = SumarDiasSinFinSemana(DFecha,IntDayAdd)
            'end if
         end if
      else
        SumarDias = DFecha
      end if
    End Function

    '--------------------------------------------------------------------------------------------------------------------------------
    'Funcion para escribir el encabezado del reporte en la cadena HTML
    function DespliegaEncabezado()
       strHTML = strHTML & " <br> "
       strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">GRUPO REYES KURI, S.C. </font></strong> <br> "
       strHTML = strHTML & "<strong><font color=""#969696"" size=""3"" face=""Arial, Helvetica, sans-serif""> TRACKING AEREO " & " </font></strong>"
       'strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
       strHTML = strHTML & "<table bordercolor=""#7D997D"" border=""1"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)


       contCamposplantilla = UBound(arrcampos,2) - 1
       strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)
       For intRow = 0 To contCamposplantilla
           strHTML = strHTML & "<td width=""120"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & arrcampos(2,intRow) & " </font></strong></td>" & chr(13) & chr(10)
       next
       strHTML = strHTML & "</tr>"& chr(13) & chr(10)

    end function




     '-------------------------------------------------------------------------------------------------------------------------------
    'Funcion para escribir el encabezado del reporte en la cadena HTML
    function agregarfilaHTML(COLOR,      REFERENCIA,      OTD2,     GUIA_MASTER,     GUIA_HOUSE   ,  P_OA ,   FORWARDER_AIR_LINE,     AEROPUERTO_SALIDA,     CUSTOM_OF_DISPATCH,    NOTIFICACION_GUIA,     FECHA_NOTIFICACION,     VESSEL,     IMPORT_DOCUMENT,     PROVEEDOR,     INVOICE,     MODEL,     DESCRIPCION_COMERCIAL,     DESCRIPTION_CODE,     QTY,     INCOTERMS,     SERIAL_NUMBER,     CERT_NOM,	    ORIGIN_ETD,     ETA_LAX,     ATA_CUSTOM,     RESQUEST_DUTIES,     FECHA_DE_REVALIDACION,     FECHA_DE_PREVIO,     ETA_CUSTOM_CLEARANCE ,    DATE_OF_CLEARANCE,     ETA_C_P,      ATA_CP ,    ETA_W_H,     ATA_WH,     TIMEOFDELIVERY,      REMARKS,     MODALIDAD ,     WEEK,     AMOUNTOFDUTIES,     NUM_INVOICECUSTOM,     DATEINVOICECUSTOM,     ADUDESPACHO, RMKATDORIGIN,    RMKATAPORT,    RMKDEPACHO,    RMKATDRAIL,    RMKCP,    ATASPL,    STATUS,    LASTRMK,    KPISTATUS )
             
       contCamposplantilla = UBound(arrcampos,2) - 1
       For intRow = 0 To contCamposplantilla

           if arrcampos(1,intRow) = "REFERENCIA"  then' Nombre del campo
             arrcampos(4,intRow) = REFERENCIA  ' titulo
           end if
           
           if arrcampos(1,intRow) = "ITTSNOTIFDATE"  then' Nombre del campo
             arrcampos(4,intRow) =  FECHA_NOTIFICACION ' titulo
           end if
           if arrcampos(1,intRow) = "BOFL_AWBM"  then' Nombre del campo
             arrcampos(4,intRow) = GUIA_MASTER  ' titulo
           end if
           if arrcampos(1,intRow) = "CONTAINER_AWBH"  then' Nombre del campo
             arrcampos(4,intRow) = GUIA_HOUSE  ' titulo
           end if
           
           if arrcampos(1,intRow) = "IMPORTDOCUMENT"  then' Nombre del campo
             arrcampos(4,intRow) = IMPORT_DOCUMENT  ' titulo
           end if
           
           if arrcampos(1,intRow) = "INVOICE"  then' Nombre del campo
             arrcampos(4,intRow) = INVOICE  ' titulo
           end if
           if arrcampos(1,intRow) = "MODEL"  then' Nombre del campo
             arrcampos(4,intRow) = MODEL  ' titulo
           end if
           if arrcampos(1,intRow) = "DESCRIPTION"  then' Nombre del campo
             arrcampos(4,intRow) = DESCRIPCION_COMERCIAL  ' titulo
           end if
           if arrcampos(1,intRow) = "DESCRIPTIONCODE"  then' Nombre del campo
             arrcampos(4,intRow) = DESCRIPTION_CODE  ' titulo
           end if
           if arrcampos(1,intRow) = "QTY"  then' Nombre del campo
             arrcampos(4,intRow) = QTY ' titulo
           end if
           
           if arrcampos(1,intRow) = "ETAPORT_LAX"  then' Nombre del campo
             arrcampos(4,intRow) = ETA_LAX  ' titulo
           end if
           
           if arrcampos(1,intRow) = "SERIALNUMBER"  then' Nombre del campo
             arrcampos(4,intRow) = SERIAL_NUMBER  ' titulo
           end if
           if arrcampos(1,intRow) = "DATEOFRELEASE"  then' Nombre del campo
             arrcampos(4,intRow) = FECHA_DE_REVALIDACION  ' titulo
           end if
           if arrcampos(1,intRow) = "AMOUNTOFDUTIES"  then' Nombre del campo
             arrcampos(4,intRow) = AMOUNTOFDUTIES  ' titulo
           end if
           if arrcampos(1,intRow) = "PREVIO"  then' Nombre del campo
             arrcampos(4,intRow) = FECHA_DE_PREVIO  ' titulo
           end if
           if arrcampos(1,intRow) = "DATEOFCLEARANCE"  then' Nombre del campo
             arrcampos(4,intRow) = DATE_OF_CLEARANCE  ' titulo
           end if
           if arrcampos(1,intRow) = "ETAWH"  then' Nombre del campo
             arrcampos(4,intRow) = ETA_W_H  ' titulo
           end if
           
           if arrcampos(1,intRow) = "HISTORIAL"  then' Nombre del campo
             arrcampos(4,intRow) = REMARKS  ' titulo
           end if
           
           if arrcampos(1,intRow) = "CUSTOMOFDISPATCH"  then' Nombre del campo
             arrcampos(4,intRow) = ADUDESPACHO  ' titulo
           end if
           
           if arrcampos(1,intRow) = "STATUS"  then' Nombre del campo
             arrcampos(4,intRow) = STATUS  ' titulo
           end if
           
           if arrcampos(1,intRow) = "KPISTATUS"  then' Nombre del campo
             arrcampos(4,intRow) = KPISTATUS  ' titulo
           end if
       next

       '*******************************************************************************************************


       if COLOR=1 then
          str_color = "#FFFFFF"
          str_fcolor = "#000000"
       else
         if COLOR=2 then ' AZUL DIFERENCIA A FAVOR AGENCIA
            str_color = "#FFFFFF"
            str_fcolor = "#0099FF"
         else
            if COLOR=3 then ' ROJO RETRASO
               str_color = "#FFFFFF"
               str_fcolor = "#FF0000"
            end if
         end if
       end if
      strColorNA = "#DCDCDC"

       if strTipoFiltro  = "BotonOtrosOpVivas" and ATA_WH <> "" and not isnull(ATA_WH) then
           str_color = "#FFFFCC"
          'str_color = "#99CCFF"
       end if


       if COLOR <> 2 and COLOR <> 3 then

           strHTML = strHTML& "<tr bgcolor= '"&str_color&"' align=""center"" >"& chr(13) & chr(10)
           For intRow = 0 To contCamposplantilla
               if arrcampos(4,intRow) = "N/A" then
                  strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " " & " </font></td>" & chr(13) & chr(10) '
               else
                  strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & arrcampos(4,intRow) & " </font></td>" & chr(13) & chr(10) '
               end if
           next
           strHTML = strHTML & "</tr>"& chr(13) & chr(10)

       else

             strHTML = strHTML& "<tr bgcolor= '"&str_color&"' align=""center"" >"& chr(13) & chr(10)
             For intRow = 0 To contCamposplantilla
                 if arrcampos(4,intRow) = "N/A" then
                    strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "  &" </font></strong></td>" & chr(13) & chr(10) '
                 else
                    strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & arrcampos(4,intRow)        &" </font></strong></td>" & chr(13) & chr(10) '
                 end if
             next
             strHTML = strHTML & "</tr>"& chr(13) & chr(10)


       end if

    end function

     '-------------------------------------------------------------------------------------------------------------------------------
%>

<%
    'TipoFiltro

     tempstrOficina = adu_ofi( Session("GAduana") )
     IF NOT enproceso(tempstrOficina) THEN



    MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
    MM_EXTRANET_STRING_STATUS = ODBC_POR_ADUANA(Session("GAduana")&"_STATUS")

    Response.Buffer = TRUE
    Response.Addheader "Content-Disposition", "attachment;filename=TRACKING_AEREO.xls"
    Response.ContentType = "application/vnd.ms-excel"
    Server.ScriptTimeOut=100000

    strUsuario     = request.Form("user")
    strTipoUsuario = request.Form("TipoUser")
    strPermisos    = Request.Form("Permisos")
    permi          = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01 ")
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
    strTipoFiltro         = trim(request.Form("TipoFiltro"))
    strDateIni            = trim(request.Form("txtDateIni"))
    strDateFin            = trim(request.Form("txtDateFin"))
    strTipoFecha          = trim(request.Form("txtTipoFecha"))
    strLinNav             = trim(request.Form("txtLinNav"))
    strModalidad          = trim(request.Form("txtMod"))
    strProv               = trim(request.Form("txtProv"))
    strfiltrosrestantes   = trim(request.Form("txtfiltrosrestantes"))
    strTipoOtrosFiltros   = trim(request.Form("txttipoOtrosFiltros"))

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


    strHTML = ""


    if strCodError = "0" then

    tmpTipo = ""
    strSQL  = ""

 
		if strTipoFiltro  = "BotonOtrosOpVivas" then  'Otros Flitros de captura libre
			strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
				 "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
				 "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
				 "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
				 "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
				 "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
				 "        SSDAGI01.PATENT01,                     " & _
				 "        CONCAT(SSDAGI01.PATENT01, CONCAT( '',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
				 "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
				 "        C01REFER.feta01    AS ETA_PORT,        " & _
				 "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
				 "        SSDAGI01.fecent01,                     " & _
				 "        C01REFER.frev01    AS REVALIDACION,    " & _
				 "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
				 "        C01REFER.fpre01    AS PREVIO,          " & _
				 "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
				 "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
				 "        SSDAGI01.desf0101  AS FACTURAS,        " & _
				 "        firmae01,                              " & _
				 "        frec01 as FecITTS,                     " & _
				 "        cvrexp01 as FORWARDER,                 " & _
				 "        cvela01  as  AIRLINE,                  " & _
				 "        feorig01,                              " & _
				 "        etalax01,                              " & _
				 "        cbuq01,                                " & _
				 "        CVEPED01,                              " & _
				 "        cveptoemb,                             " & _
				 "        ADUDES01                               " & _
				 "  FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01   " & _
				 "  WHERE adusec01=470 and modo01 = 'T' AND " & _
				 "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv     & _
				 "  GROUP BY REFCIA01 " & _
				 " ORDER BY ETA_PORT2, ETA_PORT  "

		end if
              

         if not trim(strSQL)="" then
            Set RsRep = Server.CreateObject("ADODB.Recordset")
            RsRep.ActiveConnection = MM_EXTRANET_STRING
            RsRep.Source = strSQL
            RsRep.CursorType = 0
            RsRep.CursorLocation = 2
            RsRep.LockType = 1
            RsRep.Open()

            banCargaRun = false

            'response.write(strSQL)
            'response.end

            if not RsRep.eof then
               DespliegaEncabezado()
               'response.end
               intColumna = 1
               While NOT RsRep.EOF AND not banCargaRun
                         ' ya tenemos los registros a nivel pedimento, ahora vamos por todos los campos a nivel pedimento restantes
                         StrRefer = RsRep.Fields.Item("REFERENCIA").Value
                         '**************************************************************************************************************

                          Bolbanrecti = True
                        
                    if Bolbanrecti = True then
                         ' GUIA
                         strGuiaMaster      = ""
                         strGuiaMasterHouse = ""
                         if StrRefer <> "" then
                             Set Recguia = Server.CreateObject("ADODB.Recordset")
                             Recguia.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT  IF( IDNGUI04=1,numgui04,'') AS guiaMaster,  " & _
                                          "         IF( IDNGUI04=2,numgui04,'') AS guiaHouse    " & _
                                          " from ssguia04  " & _
                                          " where refcia04='" & ltrim(StrRefer)&"'"
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
                                           strGuiaMaster      = strGuiaMaster & ", "& Recguia.Fields.Item("guiaMaster").Value
                                       end if
                                       intcountguia1= intcountguia1 + 1
                                    end if

                                    if Recguia.Fields.Item("guiaHouse").Value <> "" then
                                       if intcountguia2 = 1 then
                                           strGuiaMasterHouse = Recguia.Fields.Item("guiaHouse").Value
                                       else
                                           strGuiaMasterHouse = strGuiaMasterHouse & ", "& Recguia.Fields.Item("guiaHouse").Value
                                       end if
                                       intcountguia2= intcountguia2 + 1
                                    end if

                                 Recguia.movenext

                                 Wend
                             end if
                             Recguia.close
                             set Recguia = Nothing
                         end if

                         '**************************************************************************************************************
                         ' fORWARDWERS Y LINEAS AEREAS
                          strForwarder = RsRep.Fields.Item("FORWARDER").Value
                           if strForwarder <> "" then
                              Set RForwLA = Server.CreateObject("ADODB.Recordset")
                              RForwLA.ActiveConnection = MM_EXTRANET_STRING
                              strSqlSel =  " select CVEFOR01  " & _
                                           " from c01reexp        " & _
                                           " where cvrexP01='"&ltrim(strForwarder)&"' "
                              RForwLA.Source = strSqlSel
                              RForwLA.CursorType = 0
                              RForwLA.CursorLocation = 2
                              RForwLA.LockType = 1
                              RForwLA.Open()
                              if not RForwLA.eof then
                                  strForwarder = RForwLA.Fields.Item("CVEFOR01").Value
                              else
                                  strForwarder = ""
                              end if
                              RForwLA.close
                              set RForwLA = Nothing
                          end if


                         strForwardertmp = RsRep.Fields.Item("FORWARDER").Value

                         if strForwardertmp <> "" then
                             Set Rshipping_line = Server.CreateObject("ADODB.Recordset")
                             Rshipping_line.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT clifor01, " & _
                                          "        numdia01  " & _
                                          " FROM d01reexp    " & _
                                          " where cverex01 = '" & ltrim(strForwardertmp) & "' " & permi
                             Rshipping_line.Source = strSqlSel
                             Rshipping_line.CursorType = 0
                             Rshipping_line.CursorLocation = 2
                             Rshipping_line.LockType = 1
                             Rshipping_line.Open()
                             if not Rshipping_line.eof then
                                 strForwardertmp = Rshipping_line.Fields.Item("clifor01").Value
                                 StdEtdLoad = Rshipping_line.Fields.Item("numdia01").Value
                             else
                                 StdEtdLoad = 0
                                 strForwardertmp = ""
                             end if
                             Rshipping_line.close
                             set Rshipping_line = Nothing
                         end if
                         if strForwardertmp <> "" then
                           strForwarder = strForwardertmp
                         end if
                         '**************************************************************************************************************

                         strProveedor = RsRep.Fields.Item("PROVEEDOR").Value
                         if strProveedor <> "" then
                             Set RProv = Server.CreateObject("ADODB.Recordset")
                             RProv.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  "select nompro22,npscli22 from ssprov22 where cvepro22=" & ltrim(strProveedor)
                             RProv.Source = strSqlSel
                             RProv.CursorType = 0
                             RProv.CursorLocation = 2
                             RProv.LockType = 1
                             RProv.Open()
                             if not RProv.eof then
                                 strProveedor = RProv.Fields.Item("npscli22").Value
                             else
                                 strProveedor = ""
                             end if
                             RProv.close
                             set RProv = Nothing
                         end if
                       
                         strModalidad = "N/A"

                         ' Impuestos
                         strImpuestos = ""
                         if StrRefer <> "" then
                             Set RImpuestos = Server.CreateObject("ADODB.Recordset")
                             RImpuestos.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT SUM(import36) as Impuestos " & _
                                          " FROM sscont36         " & _
                                          " WHERE  REFCIA36 = '"&ltrim(StrRefer)&"'  AND " & _
                                          "        FPAGOI36 = 0 " & _
                                          " GROUP BY refcia36 "
                             RImpuestos.Source = strSqlSel
                             RImpuestos.CursorType = 0
                             RImpuestos.CursorLocation = 2
                             RImpuestos.LockType = 1
                             RImpuestos.Open()
                             if not RImpuestos.eof then
                                 strImpuestos = RImpuestos.Fields.Item("Impuestos").Value
                             else
                                 strImpuestos = ""
                             end if
                             RImpuestos.close
                             set RImpuestos = Nothing
                         end if
                         '**************************************************************************************************************

                         ' Cuentas de Gastos
                         strCuentaGastos = ""
                         strFecCuentaGastos = ""
                         if StrRefer <> "" then
                             Set RCuentaGastos = Server.CreateObject("ADODB.Recordset")
                             RCuentaGastos.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT e31cgast.cgas31,  " & _
                                          "        e31cgast.fech31   " & _
                                          " FROM d31refer,e31cgast   " & _
                                          " WHERE refe31  = '"&ltrim(StrRefer)&"' AND " & _
                                          "       d31refer.cgas31=e31cgast.cgas31 AND " & _
                                          "       esta31='I' "
                             RCuentaGastos.Source = strSqlSel
                             RCuentaGastos.CursorType = 0
                             RCuentaGastos.CursorLocation = 2
                             RCuentaGastos.LockType = 1
                             RCuentaGastos.Open()
                             if not RCuentaGastos.eof then
                               intcontemp = 1
                               While NOT RCuentaGastos.EOF
                                 if intcontemp = 1 then
                                    strCuentaGastos    = RCuentaGastos.Fields.Item("cgas31").Value
                                    strFecCuentaGastos = RCuentaGastos.Fields.Item("fech31").Value
                                 else
                                    strCuentaGastos    = strCuentaGastos &", "& RCuentaGastos.Fields.Item("cgas31").Value
                                    strFecCuentaGastos = strFecCuentaGastos&", "& RCuentaGastos.Fields.Item("fech31").Value
                                 end if
                                 intcontemp = intcontemp + 1
                                 RCuentaGastos.movenext
                               Wend
                             end if
                             RCuentaGastos.close
                             set RCuentaGastos = Nothing
                         end if
                         '**************************************************************************************************************

                         ' Incoterms
                         strIncoterms = ""

                         ' Vamos por las mercancias
                           strPO_Pedido = ""
                           strPO_PedidoNA = ""
                           strDescMerc  = ""
                           strModelo    = ""
                           strDescCode  = ""
                           Set RMercancias = Server.CreateObject("ADODB.Recordset")
                           RMercancias.ActiveConnection = MM_EXTRANET_STRING
                           strSqlSel = " Select  refe05,pedi05, desc05, cpro05,descod05 " & _
                                       " from d05artic  " & _
                                       " where refe05='" & ltrim(StrRefer) & "' "
                           RMercancias.Source = strSqlSel
                           RMercancias.CursorType = 0
                           RMercancias.CursorLocation = 2
                           RMercancias.LockType = 1
                           RMercancias.Open()
                           if not RMercancias.eof then
                           intcontemp = 1
                           intcontped = 1
                             While NOT RMercancias.EOF
                                 if RMercancias.Fields.Item("pedi05").Value <> ""  then
                                    if UCase(ltrim(RMercancias.Fields.Item("pedi05").Value)) <> "N/A" then
                                        if intcontped = 1 then
                                           strPO_Pedido  = RMercancias.Fields.Item("pedi05").Value
                                        else
                                           strPO_Pedido  = strPO_Pedido& ", "&RMercancias.Fields.Item("pedi05").Value
                                        end if
                                        intcontped = intcontped + 1
                                    else
                                       strPO_PedidoNA = "N/A"
                                    end if
                                 end if

                                 if intcontemp = 1 then
                                    strDescMerc   = RMercancias.Fields.Item("desc05").Value
                                    strModelo     = RMercancias.Fields.Item("cpro05").Value
                                    strDescCode   = RMercancias.Fields.Item("descod05").Value
                                 else
                                    strDescMerc   = strDescMerc & ", " & RMercancias.Fields.Item("desc05").Value
                                    strModelo     = strModelo & ", " & RMercancias.Fields.Item("cpro05").Value
                                    strDescCode   = strDescCode & ", " & RMercancias.Fields.Item("descod05").Value
                                 end if
                                 intcontemp = intcontemp + 1
                                 RMercancias.movenext
                             Wend

                             if strPO_Pedido = "" then
                              strPO_Pedido = strPO_PedidoNA

                             end if
                           end if


                         '**************************************************************************************************************

                         ' Cantidad de fracciones
                         strQTY = ""
                         if StrRefer <> "" then
                             Set Rfracciones = Server.CreateObject("ADODB.Recordset")
                             Rfracciones.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT SUM(CANCOM02) as QTY " & _
                                          " FROM ssfrac02         " & _
                                          " WHERE  REFCIA02 = '"&ltrim(StrRefer)&"' " & _
                                          " GROUP BY refcia02 "
                             Rfracciones.Source = strSqlSel
                             Rfracciones.CursorType = 0
                             Rfracciones.CursorLocation = 2
                             Rfracciones.LockType = 1
                             Rfracciones.Open()
                             if not Rfracciones.eof then
                                 strQTY = Rfracciones.Fields.Item("QTY").Value
                             else
                                 strQTY = ""
                             end if
                             Rfracciones.close
                             set Rfracciones = Nothing
                         end if
                         '**************************************************************************************************************

                         ' Fechas de documentos
                         'CNO -> CERTIFICADO NOM
                         'CNS -> CARTA CON NUMERO DE INSTRUCCIONES
                         'GUA -> GUIA AEREA
                         strCERTNOM  = ""
                         StrNUMSERIE = ""
                         strGUANotificacion = ""
                         strGUA      = ""
                         if StrRefer <> "" then
                             Set RFecDocu = Server.CreateObject("ADODB.Recordset")
                             RFecDocu.ActiveConnection = MM_EXTRANET_STRING
                               strSqlSel =  " SELECT C07DOCRE.CLAV07,  " & _
                                          "         C07DOCRE.FECH07, " & _
                                          "         C07DOCRE.ORIG07, " & _
                                          "         C07DOCOR.DESC07, " & _
                                          "         DISP07            " & _
                                          " FROM C07DOCRE LEFT JOIN C07DOCOR ON C07DOCRE.ORIG07=C07DOCOR.CLAV07 " & _
                                          " WHERE  C07DOCRE.REFE07 ='"&ltrim(StrRefer)&"' AND " & _
                                          "       (C07DOCRE.CLAV07='CNO' or " & _
                                          "        C07DOCRE.clav07='CNS' or " & _
                                          "        C07DOCRE.clav07='GUA'   )"
                             RFecDocu.Source = strSqlSel
                             RFecDocu.CursorType = 0
                             RFecDocu.CursorLocation = 2
                             RFecDocu.LockType = 1
                             RFecDocu.Open()
                             While NOT RFecDocu.EOF
                                 if RFecDocu.Fields.Item("CLAV07").Value <>"" and ltrim(RFecDocu.Fields.Item("CLAV07").Value) = "CNO"  then
                                      if RFecDocu.Fields.Item("DISP07").Value = "F"   then
                                         strCERTNOM  = "N/A"
                                      else
                                         strCERTNOM  = RFecDocu.Fields.Item("FECH07").Value
                                      end if
                                 else
                                    if RFecDocu.Fields.Item("CLAV07").Value <>"" and ltrim(RFecDocu.Fields.Item("CLAV07").Value) = "CNS"  then
                                         if RFecDocu.Fields.Item("DISP07").Value = "F"   then
                                            StrNUMSERIE = "N/A"
                                         else
                                            StrNUMSERIE = RFecDocu.Fields.Item("FECH07").Value
                                         end if
                                    else
                                       if RFecDocu.Fields.Item("CLAV07").Value <>"" and ltrim(RFecDocu.Fields.Item("CLAV07").Value) = "GUA"  then
                                          strGUA = RFecDocu.Fields.Item("FECH07").Value
                                          strGUANotificacion = RFecDocu.Fields.Item("DESC07").Value
                                       end if
                                    end if
                                 end if
                                 RFecDocu.movenext
                             Wend
                             RFecDocu.close
                             set RFecDocu = Nothing
                         end if
                         '**************************************************************************************************************

                         ' OBSERVACIONES
                         strObservaciones = ""
                         if StrRefer <> "" then
                             Set RObservEtapas = Server.CreateObject("ADODB.Recordset")
                             RObservEtapas.ActiveConnection = MM_EXTRANET_STRING_STATUS

                             strSQL = " SELECT (n_secuenc), " & _
                                      "        D.n_etapa,   " & _
                                      "        f_fecha,     " & _
                                      "        m_observ     " & _
                                      " FROM ETXPD as D     " & _
                                      " WHERE not(date_format(D.f_fecha,'%Y%m%d') = '00000000') and  " & _
                                      "       D.c_referencia = '"&ltrim(StrRefer)&"'" & _
                                      " ORDER BY N_ETAPA, N_SECUENC "

                             RObservEtapas.Source = strSQL
                             RObservEtapas.CursorType = 0
                             RObservEtapas.CursorLocation = 2
                             RObservEtapas.LockType = 1
                             RObservEtapas.Open()
                             intcontObs = 1
                             While NOT RObservEtapas.EOF
                                 strObsTemp = RObservEtapas.Fields.Item("m_observ").Value
                                 if strObsTemp <>"" and ltrim(strObsTemp) <> "" and InStr( strObservaciones, strObsTemp) = 0 then
                                    if intcontObs = 1 then
                                       strObservaciones  =RObservEtapas.Fields.Item("m_observ").Value
                                    else
                                       strObservaciones  = strObservaciones & " ; "& RObservEtapas.Fields.Item("m_observ").Value
                                    end if
                                    intcontObs = intcontObs + 1
                                 end if
                             RObservEtapas.movenext
                             Wend
                             RObservEtapas.close
                             set RObservEtapas = Nothing
                         end if
                         '**************************************************************************************************************

                         ' FACTURAS

                             StrINVOICE = ""
                             if StrRefer <> "" then
                                 Set RFactuRef = Server.CreateObject("ADODB.Recordset")
                                 RFactuRef.ActiveConnection = MM_EXTRANET_STRING
                                 strSQL = " SELECT NUMFAC39, FECFAC39 " & _
                                          " FROM SSFACT39 " & _
                                          " WHERE REFCIA39='" & ltrim(StrRefer) & "'"
                                 RFactuRef.Source = strSQL
                                 RFactuRef.CursorType = 0
                                 RFactuRef.CursorLocation = 2
                                 RFactuRef.LockType = 1
                                 RFactuRef.Open()
                                 intcontObs = 1
                                 While NOT RFactuRef.EOF
                                     StrINVOICETemp    = RFactuRef.Fields.Item("NUMFAC39").Value
                                     StrfECINVOICETemp = RFactuRef.Fields.Item("FECFAC39").Value
                                     if StrINVOICETemp <> "" and StrfECINVOICETemp <> "" then
                                        if intcontObs = 1 then
                                           StrINVOICE  = StrINVOICETemp
                                        else
                                           StrINVOICE  = StrINVOICE & "; "& StrINVOICETemp
                                        end if
                                        intcontObs = intcontObs + 1
                                     end if
                                 RFactuRef.movenext
                                 Wend

                                 RFactuRef.close
                                 set RFactuRef = Nothing
                             end if

                         ' Contenedores
                         strNumConte = ""
                         strATDRAIL  = ""
                         strETA_CP   = ""
                         strATAC_P   = ""
                         strETAW_H   = ""
                         '----------------
                         strFechaATAWH      = ""
                         strComentarioATAWH = ""
                         strHoraATAWH       = ""

                         if StrRefer <> "" then
                             Set RContenedores = Server.CreateObject("ADODB.Recordset")
                             RContenedores.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " select marc01, " & _
                                          "       fcTren01 as ATDRAIL, " & _
                                          "       feCont01 as ETA_CP,  " & _
                                          "       frCont01 as ATAC_P,  " & _
                                          "       feAlma01 as ETAW_H   " & _
                                          " from d01conte where refe01 = '" & ltrim(StrRefer) & "' "

                             RContenedores.Source = strSqlSel
                             RContenedores.CursorType = 0
                             RContenedores.CursorLocation = 2
                             RContenedores.LockType = 1
                             RContenedores.Open()
                             if not RContenedores.eof then
                               While NOT RContenedores.EOF
                                       strNumConte = RContenedores.Fields.Item("marc01").Value
                                       strATDRAIL  = RContenedores.Fields.Item("ATDRAIL").Value
                                       '*********************************************
                                         strFechaATAWH      = ""
                                         strComentarioATAWH = ""
                                         strHoraATAWH       = ""
                                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                         strSqlSel = " SELECT f_fecha,   " & _
                                                     "        t_hora,   " & _
                                                     "        m_observ  " & _
                                                     " FROM etxcoi, etaps " & _
                                                     " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                     "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and    " & _
                                                     "       ltrim(c_conte)      = '" & ltrim(strNumConte) & "' and " & _
                                                     "       d_abrev      = 'LLP'             " & _
                                                     " order by n_secuenc desc                  "
                                         RConteDetalle.Source = strSqlSel
                                         RConteDetalle.CursorType = 0
                                         RConteDetalle.CursorLocation = 2
                                         RConteDetalle.LockType = 1
                                         RConteDetalle.Open()
                                         if not RConteDetalle.eof then
                                             strFechaATAWH       = RConteDetalle.Fields.Item("f_fecha").Value
                                             strHoraATAWH        = RConteDetalle.Fields.Item("t_hora").Value
                                             strObsTemp = ""
                                             intcontObs = 1
                                             While NOT RConteDetalle.EOF
                                                 strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                 if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                    if intcontObs = 1 then
                                                       strComentarioATAWH  = RConteDetalle.Fields.Item("m_observ").Value
                                                    else
                                                       strComentarioATAWH  = strComentarioATAWH & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                    end if
                                                    intcontObs = intcontObs + 1
                                                 end if
                                             RConteDetalle.movenext
                                             Wend
                                         end if
                                         RConteDetalle.close
                                         set RConteDetalle = Nothing

                                         '*********************************************
                                         strATAC_P           = ""
                                         strComentarioATAC_P = ""
                                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                         strSqlSel = " SELECT f_fecha,  " & _
                                                     "        m_observ  " & _
                                                     " FROM etxcoi, etaps " & _
                                                     " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                     "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and    " & _
                                                     "       ltrim(c_conte)      = '" & ltrim(strNumConte) & "' and " & _
                                                     "       d_abrev      = 'CP'             " & _
                                                     " order by n_secuenc desc                  "
                                         RConteDetalle.Source = strSqlSel
                                         RConteDetalle.CursorType = 0
                                         RConteDetalle.CursorLocation = 2
                                         RConteDetalle.LockType = 1
                                         RConteDetalle.Open()
                                         if not RConteDetalle.eof then
                                             strATAC_P            = RConteDetalle.Fields.Item("f_fecha").Value
                                             strObsTemp = ""
                                             intcontObs = 1
                                             While NOT RConteDetalle.EOF
                                                 strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                 if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                    if intcontObs = 1 then
                                                       strComentarioATAC_P  = RConteDetalle.Fields.Item("m_observ").Value
                                                    else
                                                       strComentarioATAC_P  = strComentarioATAC_P & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                    end if
                                                    intcontObs = intcontObs + 1
                                                 end if
                                             RConteDetalle.movenext
                                             Wend
                                         end if
                                         RConteDetalle.close
                                         set RConteDetalle = Nothing

                                       '*********************************************
                                         strATDRAIL          = ""
                                         strComentarioATDRAIL = ""
                                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                         strSqlSel = " SELECT f_fecha,  " & _
                                                     "        m_observ  " & _
                                                     " FROM etxcoi, etaps " & _
                                                     " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                     "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and    " & _
                                                     "       ltrim(c_conte)      = '" & ltrim(strNumConte) & "' and " & _
                                                     "       d_abrev      = 'RAIL'            " & _
                                                     " order by n_secuenc desc                  "
                                         RConteDetalle.Source = strSqlSel
                                         RConteDetalle.CursorType = 0
                                         RConteDetalle.CursorLocation = 2
                                         RConteDetalle.LockType = 1
                                         RConteDetalle.Open()
                                         if not RConteDetalle.eof then
                                             strATDRAIL            = RConteDetalle.Fields.Item("f_fecha").Value
                                             strObsTemp = ""
                                             intcontObs = 1
                                             While NOT RConteDetalle.EOF
                                                 strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                 if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                    if intcontObs = 1 then
                                                       strComentarioATDRAIL  = RConteDetalle.Fields.Item("m_observ").Value
                                                    else
                                                       strComentarioATDRAIL  = strComentarioATDRAIL & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                    end if
                                                    intcontObs = intcontObs + 1
                                                 end if
                                             RConteDetalle.movenext
                                             Wend
                                         end if
                                         RConteDetalle.close
                                         set RConteDetalle = Nothing

                                       '*********************************************
                                         strATASPLTMP           = ""
                                         strTimeSLP             = ""
                                         strComentarioATASPLTMP = ""
                                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                         strSqlSel = " SELECT f_fecha,  " & _
                                                     "        t_hora,   " & _
                                                     "        m_observ  " & _
                                                     " FROM etxcoi, etaps " & _
                                                     " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                     "       c_referencia = '" & ltrim(StrRefer)    & "' and    " & _
                                                     "       c_conte      = '" & ltrim(strNumConte) & "' and " & _
                                                     "       d_abrev      = 'SPL'             " & _
                                                     " order by n_secuenc desc                  "
                                         RConteDetalle.Source = strSqlSel
                                         RConteDetalle.CursorType = 0
                                         RConteDetalle.CursorLocation = 2
                                         RConteDetalle.LockType = 1
                                         RConteDetalle.Open()
                                         if not RConteDetalle.eof then
                                             strATASPLTMP = RConteDetalle.Fields.Item("f_fecha").Value
                                             strTimeSLP   = RConteDetalle.Fields.Item("t_hora").Value
                                             strObsTemp = ""
                                             intcontObs = 1
                                             While NOT RConteDetalle.EOF
                                                 strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                 if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                    if intcontObs = 1 then
                                                       strComentarioATASPLTMP  = RConteDetalle.Fields.Item("m_observ").Value
                                                    else
                                                       strComentarioATASPLTMP  = strComentarioATASPLTMP & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                    end if
                                                    intcontObs = intcontObs + 1
                                                 end if
                                             RConteDetalle.movenext
                                             Wend

                                         end if
                                         RConteDetalle.close
                                         set RConteDetalle = Nothing
                                       '*********************************************

                                 RContenedores.movenext
                               Wend
                             else

                                          strATDRAIL  = ""
                                          '*********************************************
                                           strFechaATAWH      = ""
                                           strComentarioATAWH = ""
                                           strHoraATAWH       = ""
                                           Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                           RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                           strSqlSel = " SELECT f_fecha,   " & _
                                                       "        t_hora,   " & _
                                                       "        m_observ  " & _
                                                       " FROM etxcoi, etaps " & _
                                                       " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                       "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and " & _
                                                       "       d_abrev      = 'LLP'             " & _
                                                       " order by n_secuenc desc                  "
                                           RConteDetalle.Source = strSqlSel
                                           RConteDetalle.CursorType = 0
                                           RConteDetalle.CursorLocation = 2
                                           RConteDetalle.LockType = 1
                                           RConteDetalle.Open()
                                           if not RConteDetalle.eof then
                                               strFechaATAWH       = RConteDetalle.Fields.Item("f_fecha").Value
                                               strHoraATAWH        = RConteDetalle.Fields.Item("t_hora").Value
                                               strObsTemp = ""
                                               intcontObs = 1
                                               While NOT RConteDetalle.EOF
                                                   strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                   if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                      if intcontObs = 1 then
                                                         strComentarioATAWH  = RConteDetalle.Fields.Item("m_observ").Value
                                                      else
                                                         strComentarioATAWH  = strComentarioATAWH & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                      end if
                                                      intcontObs = intcontObs + 1
                                                   end if
                                               RConteDetalle.movenext
                                               Wend
                                           end if
                                           RConteDetalle.close
                                           set RConteDetalle = Nothing

                                           '*********************************************
                                           strATAC_P           = ""
                                           strComentarioATAC_P = ""
                                           Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                           RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                           strSqlSel = " SELECT f_fecha,  " & _
                                                       "        m_observ  " & _
                                                       " FROM etxcoi, etaps " & _
                                                       " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                       "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and    " & _
                                                       "       d_abrev      = 'CP'             " & _
                                                       " order by n_secuenc desc                  "
                                           RConteDetalle.Source = strSqlSel
                                           RConteDetalle.CursorType = 0
                                           RConteDetalle.CursorLocation = 2
                                           RConteDetalle.LockType = 1
                                           RConteDetalle.Open()
                                           if not RConteDetalle.eof then
                                               strATAC_P            = RConteDetalle.Fields.Item("f_fecha").Value
                                               strObsTemp = ""
                                               intcontObs = 1
                                               While NOT RConteDetalle.EOF
                                                   strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                   if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                      if intcontObs = 1 then
                                                         strComentarioATAC_P  = RConteDetalle.Fields.Item("m_observ").Value
                                                      else
                                                         strComentarioATAC_P  = strComentarioATAC_P & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                      end if
                                                      intcontObs = intcontObs + 1
                                                   end if
                                               RConteDetalle.movenext
                                               Wend
e
                                           end if
                                           RConteDetalle.close
                                           set RConteDetalle = Nothing
                                         '*********************************************

                                         '*********************************************
                                           strATDRAIL          = ""
                                           strComentarioATDRAIL = ""
                                           Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                           RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                           strSqlSel = " SELECT f_fecha,  " & _
                                                       "        m_observ  " & _
                                                       " FROM etxcoi, etaps " & _
                                                       " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                       "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and " & _
                                                       "       d_abrev      = 'RAIL'            " & _
                                                       " order by n_secuenc desc                  "
                                           RConteDetalle.Source = strSqlSel
                                           RConteDetalle.CursorType = 0
                                           RConteDetalle.CursorLocation = 2
                                           RConteDetalle.LockType = 1
                                           RConteDetalle.Open()
                                           if not RConteDetalle.eof then
                                               strATDRAIL            = RConteDetalle.Fields.Item("f_fecha").Value
                                               strObsTemp = ""
                                               intcontObs = 1
                                               While NOT RConteDetalle.EOF
                                                   strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                   if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                      if intcontObs = 1 then
                                                         strComentarioATDRAIL  = RConteDetalle.Fields.Item("m_observ").Value
                                                      else
                                                         strComentarioATDRAIL  = strComentarioATDRAIL & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                      end if
                                                      intcontObs = intcontObs + 1
                                                   end if
                                               RConteDetalle.movenext
                                               Wend
                                           end if
                                           RConteDetalle.close
                                           set RConteDetalle = Nothing
                                         '*********************************************

                                         '*********************************************
                                           strATASPLTMP           = ""
                                           strTimeSLP             = ""
                                           strComentarioATASPLTMP = ""
                                           Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                           RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                           strSqlSel = " SELECT f_fecha,  " & _
                                                       "        t_hora,   " & _
                                                       "        m_observ  " & _
                                                       " FROM etxcoi, etaps " & _
                                                       " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                       "       c_referencia = '" & ltrim(StrRefer)    & "' and " & _
                                                       "       d_abrev      = 'SPL'             " & _
                                                       " order by n_secuenc desc                  "
                                           RConteDetalle.Source = strSqlSel
                                           RConteDetalle.CursorType = 0
                                           RConteDetalle.CursorLocation = 2
                                           RConteDetalle.LockType = 1
                                           RConteDetalle.Open()
                                           if not RConteDetalle.eof then
                                               strATASPLTMP = RConteDetalle.Fields.Item("f_fecha").Value
                                               strTimeSLP   = RConteDetalle.Fields.Item("t_hora").Value
                                               strObsTemp = ""
                                               intcontObs = 1
                                               While NOT RConteDetalle.EOF
                                                   strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                   if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                      if intcontObs = 1 then
                                                         strComentarioATASPLTMP  = RConteDetalle.Fields.Item("m_observ").Value
                                                      else
                                                         strComentarioATASPLTMP  = strComentarioATASPLTMP & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                      end if
                                                      intcontObs = intcontObs + 1
                                                   end if
                                               RConteDetalle.movenext
                                               Wend

                                           end if
                                           RConteDetalle.close
                                           set RConteDetalle = Nothing
                                         '*********************************************

                             end if
                             RContenedores.close
                             set RContenedores = Nothing

                                       StrREFERENCIA = StrRefer
                                       StrPGUIA_MASTER	         = strGuiaMaster      'GUIA MASTER


                                       StrPFORWARDER_AIR_LINE    = strForwarder       'FORWARDER Y/O AIR  LINE

                                       if RsRep.Fields.Item("PORT_DISCHARGE").Value <> "" and RsRep.Fields.Item("PORT_DISCHARGE").Value = "200" then
                                         StrPGUIA_HOUSE	           = strNumConte 'GUIA HOUSE - para pantaco tomar el contenedor
                                       else
                                         StrPGUIA_HOUSE	           = strGuiaMasterHouse 'GUIA HOUSE
                                       end if


                                       strADUDESPACHO = ""
                                         StrAdutmp = RsRep.Fields.Item("PORT_DISCHARGE").Value
                                         if ltrim(StrAdutmp)="430" then
                                            strADUDESPACHO = StrAdutmp&"-VERACRUZ" 'aduana aduana de despacho
                                         else
                                           if ltrim(StrAdutmp)="160" then
                                              strADUDESPACHO = StrAdutmp&"-MANZANILLO" 'aduana aduana de despacho
                                           else
                                              if ltrim(StrAdutmp)="200" or ltrim(StrAdu)="202" or ltrim(StrAdu)="470" then
                                                 strADUDESPACHO = StrAdutmp&"-PANTACO" 'aduana aduana de despacho
                                              else
                                                 if ltrim(StrAdutmp)="380" or ltrim(StrAdu)="810" then
                                                    strADUDESPACHO = StrAdutmp&"-TAMPICO" 'aduana aduana de despacho
                                                 else
                                                    if ltrim(StrAdutmp)="510" then
                                                       strADUDESPACHO = StrAdutmp&"-LAZARO CARDENAS" 'aduana aduana de despacho
                                                    else
                                                       if ltrim(StrAdutmp)="470" then
                                                          strADUDESPACHO = StrAdutmp&"-AEROPUERTO" 'aduana aduana de despacho
                                                       end if
                                                    end if
                                                 end if
                                              end if
                                           end if
                                         end if

                                         StrCUSTOM_OF_DISPATCH = ""
                                         StrAdutmp = RsRep.Fields.Item("ADUDES01").Value
                                         if ltrim(StrAdutmp)="430" then
                                            StrCUSTOM_OF_DISPATCH = StrAdutmp&"-VERACRUZ" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                         else
                                           if ltrim(StrAdutmp)="160" then
                                              StrCUSTOM_OF_DISPATCH = StrAdutmp&"-MANZANILLO" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                           else
                                              if ltrim(StrAdutmp)="200" or ltrim(StrAdu)="202" or ltrim(StrAdu)="470" then
                                                 StrCUSTOM_OF_DISPATCH = StrAdutmp&"-PANTACO" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                              else
                                                 if ltrim(StrAdutmp)="380" or ltrim(StrAdu)="810" then
                                                    StrCUSTOM_OF_DISPATCH = StrAdutmp&"-TAMPICO" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                                 else
                                                    if ltrim(StrAdutmp)="510" then
                                                       StrCUSTOM_OF_DISPATCH = StrAdutmp&"-LAZARO CARDENAS" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                                    else
                                                       if ltrim(StrAdutmp)="470" then
                                                          StrCUSTOM_OF_DISPATCH = StrAdutmp&"-AEROPUERTO" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                                       end if
                                                    end if
                                                 end if
                                              end if
                                           end if
                                         end if


                                       if RsRep.Fields.Item("PORT_LOADING").Value <> "" then 'Puerto
                                          if RsRep.Fields.Item("VESSEL_LOADING").Value <> "" then 'Pais
                                             StrPAEROPUERTO_SALIDA   = RsRep.Fields.Item("PORT_LOADING").Value&","&RsRep.Fields.Item("VESSEL_LOADING").Value 'PORT OF LOADING
                                          else
                                             StrPAEROPUERTO_SALIDA   = RsRep.Fields.Item("PORT_LOADING").Value 'PORT OF LOADING
                                          end if
                                       else
                                          if RsRep.Fields.Item("VESSEL_LOADING").Value <> "" then
                                             StrPAEROPUERTO_SALIDA   = "" 'PORT OF LOADING
                                          else
                                             StrPAEROPUERTO_SALIDA   = "" 'PORT OF LOADING
                                          end if
                                       end if


                                       StrPNOTIFICACION_GUIA	   = strGUANotificacion 'NOTIFICACION DE GUIA
                                       StrPFECHA_NOTIFICACION = formatofechaNum(RsRep.Fields.Item("FecITTS").Value) 'ASIGNADO ITTS

                                       StrPVESSEL                = strvessel    'VESSEL
                                       StrPIMPORT_DOCUMENT	     = RsRep.Fields.Item("IMPORT_DOCUMENT").Value 'IMPORT DOCUMENT
                                       StrPPROVEEDOR	           = strProveedor 'PROVEEDOR
                                       StrPINVOICE	             = StrINVOICE
                                       StrPP_O                   = strPO_Pedido  'P/O
                                       StrPMODEL 	               = strModelo    'MODEL
                                       StrPDESCRIPCION_COMERCIAL = strDescMerc  'DESCRIPCION COMERCIAL
                                       StrPDESCRIPTION_CODE	     = strDescCode  'DESCRIPTION CODE
                                       StrPQTY	                 = strQTY       'QTY
                                       StrPINCOTERMS	           = strIncoterms 'INCOTERMS

                                       '*************************************************************
                                       '***                Vamos por los remarks                  ***
                                       '*************************************************************
                                       'variables para los Remarks
                                       rmkEtdLoad    = "" 'rmk para ETDLOAD
                                       rmkATAPORT    = "" 'rmk para ATAPORT
                                       rmkDSP        = "" 'rmk para DESPACHO

                                       diaRmkEtdLoad  = 0 'rmk para ETDLOAD
                                       diaRmkATAPORT  = 0 'rmk para ATAPORT
                                       diaRmkDSP      = 0 'rmk para DESPACHO

                                       tipoRmkEtdLoad  = 1 'rmk para ETDLOAD
                                       tipoRmkATAPORT  = 1 'rmk para ATAPORT
                                       tipoRmkDSP      = 1 'rmk para DESPACHO

                                       descRmkEtdLoad  = "" 'Descripcion del rmk para ETDLOAD
                                       descRmkATAPORT  = "" 'Descripcion del rmk para ATAPORT
                                       descRmkDSP      = "" 'Descripcion del rmk para DESPACHO

                                       strLastRMKtmp = "" ' El ultimo Remark en el que se encuentre

                                       Set RsRmk = Server.CreateObject("ADODB.Recordset")
                                       RsRmk.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                       strSqlrmk = " SELECT c_refer           as referencia, " & _
                                                   "        c_conte           as contenedor, " & _
                                                   "        c_desc            as remark,     " & _
                                                   "        c01rmrks.n_cvermk as claveint,   " & _
                                                   "        d_abrev           as etapa,      " & _
                                                   "        c_cvefor          as clavefor,   " & _
                                                   "        n_dias            as dias,       " & _
                                                   "        n_tipodia         as tipodia     " & _
                                                   " FROM d01rmrks, c01rmrks, etaps          " & _
                                                   " where d01rmrks.n_cvermk = c01rmrks.n_cvermk  " & _
                                                   "       and c01rmrks.n_etapa = etaps.n_etapa " & _
                                                   "       and status = 'A'  " & _
                                                   "       and c_refer = '" & ltrim(StrRefer)    & "' " & _
                                                   "       and c_conte = '" & ltrim(strNumConte) & "' "
                                       RsRmk.Source = strSqlrmk
                                       RsRmk.CursorType = 0
                                       RsRmk.CursorLocation = 2
                                       RsRmk.LockType = 1
                                       RsRmk.Open()
                                       if not RsRmk.eof then
                                           While NOT RsRmk.EOF
                                              if RsRmk.Fields.Item("etapa").Value="ETDLOAD" then ' RMK de salida de origen
                                                 if RsRmk.Fields.Item("dias").Value > diaRmkEtdLoad then
                                                    rmkEtdLoad     = RsRmk.Fields.Item("clavefor").Value
                                                    diaRmkEtdLoad  = RsRmk.Fields.Item("dias").Value
                                                    tipoRmkEtdLoad = RsRmk.Fields.Item("tipodia").Value
                                                    descRmkEtdLoad = RsRmk.Fields.Item("remark").Value
                                                 end if
                                              else
                                                  if RsRmk.Fields.Item("etapa").Value="ATAPORT" then ' RMK de llegada a puerto
                                                      if RsRmk.Fields.Item("dias").Value > diaRmkATAPORT then
                                                         rmkATAPORT     = RsRmk.Fields.Item("clavefor").Value
                                                         diaRmkATAPORT  = RsRmk.Fields.Item("dias").Value
                                                         tipoRmkATAPORT = RsRmk.Fields.Item("tipodia").Value
                                                         descRmkATAPORT = RsRmk.Fields.Item("remark").Value
                                                      end if
                                                  else
                                                      if RsRmk.Fields.Item("etapa").Value="DSP" then ' RMK de despacho
                                                         if RsRmk.Fields.Item("dias").Value > diaRmkDSP then
                                                            rmkDSP     = RsRmk.Fields.Item("clavefor").Value
                                                            diaRmkDSP  = RsRmk.Fields.Item("dias").Value
                                                            tipoRmkDSP = RsRmk.Fields.Item("tipodia").Value
                                                            descRmkDSP = RsRmk.Fields.Item("remark").Value
                                                         end if
                                                      end if
                                                  end if
                                              end if
                                           RsRmk.movenext
                                           Wend
                                       end if
                                       RsRmk.close
                                       set RsRmk = Nothing


                                       if rmkDSP <> "" then
                                         strLastRMKtmp = descRmkDSP
                                       else
                                          if rmkATAPORT <> "" then
                                             strLastRMKtmp = descRmkATAPORT
                                          else
                                             if rmkEtdLoad <> "" then
                                                strLastRMKtmp = descRmkEtdLoad
                                             end if
                                          end if
                                       end if

                                       '**************************************************************************************
                                       '**************************************************************************************

                                       if isdate( StrNUMSERIE ) then
                                          StrPSERIAL_NUMBER 	   = YEAR( StrNUMSERIE ) & Pd(Month( StrNUMSERIE ),2) & Pd(DAY( StrNUMSERIE ),2)  'SERIAL NUMBER
                                       else
                                          StrPSERIAL_NUMBER	     = StrNUMSERIE  'SERIAL NUMBER
                                       end if
                                       StrPCERT_NOM          = strCERTNOM         'CERT. NOM

                                       if isdate(RsRep.Fields.Item("feorig01").Value) then
                                          StrPORIGIN_ETD 	 = YEAR( RsRep.Fields.Item("feorig01").Value ) & Pd(Month(RsRep.Fields.Item("feorig01").Value ),2) & Pd(DAY(RsRep.Fields.Item("feorig01").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPORIGIN_ETD = RsRep.Fields.Item("feorig01").Value 'FECHA DE NOTIFICACION
                                       end if

                                       if StdEtdLoad > 0 then
                                          if RsRep.Fields.Item("feorig01").Value  <> ""  then
                                             StrPETA_LAX = formatofechaNum(DateAdd("d",diaRmkEtdLoad,  DateAdd("d",StdEtdLoad, RsRep.Fields.Item("feorig01").Value )) ) ' Calculamos ETA PORT apartir de la fecha de salida de origen
                                          end if
                                       else
                                          if isdate(RsRep.Fields.Item("etalax01").Value) then
                                             StrPETA_LAX 	 = YEAR( RsRep.Fields.Item("etalax01").Value ) & Pd(Month(RsRep.Fields.Item("etalax01").Value ),2) & Pd(DAY(RsRep.Fields.Item("etalax01").Value ),2)  'FECHA DE NOTIFICACION
                                          else
                                             StrPETA_LAX = formatofechaNum(RsRep.Fields.Item("etalax01").Value) 'FECHA DE NOTIFICACION
                                          end if
                                       end if
                                       '************************************************************************
                                       DFechEntAux = RsRep.Fields.Item("fecent01").Value
                                       if isdate(DFechEntAux) then
                                          if DFechEntAux > date() then
                                             DFechEntAux = ""
                                          end if
                                       end if
                                       '************************************************************************

                                       if isdate(DFechEntAux) then
                                         StrETA_CUSTOM_CLEARANCE = SumarDias( SumarDias( DFechEntAux , StdATAPORTDSP,tipoStdATAPORTDSP) , diaRmkATAPORT,tipoRmkATAPORT)
                                       else
                                         StrETA_CUSTOM_CLEARANCE = SumarDias( SumarDias( RsRep.Fields.Item("etalax01").Value , StdATAPORTDSP,tipoStdATAPORTDSP) , diaRmkATAPORT,tipoRmkATAPORT)
                                       end if

                                       StrColorfila = 1
                                       StrETA_C_P = "N/A"

                                       if isdate( StrETA_CUSTOM_CLEARANCE ) then
                                         if isdate( RsRep.Fields.Item("DATE_CUSTOM").Value ) then
                                           IndFila = DateDiff("d",StrETA_CUSTOM_CLEARANCE , RsRep.Fields.Item("DATE_CUSTOM").Value )
                                           if IndFila = 0 then
                                              StrColorfila = 1
                                              StrETA_W_H_AUX = SumarDias( SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                           else
                                              StrETA_W_H_AUX = SumarDias( SumarDias( RsRep.Fields.Item("DATE_CUSTOM").Value , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                              if IndFila < 0 then
                                                  StrColorfila = 2
                                              else
                                                  StrColorfila = 3
                                              end if
                                           end if
                                         else
                                            StrETA_W_H_AUX = SumarDias( SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                            IndFila = DateDiff("d", StrETA_CUSTOM_CLEARANCE , DATE() )
                                            if IndFila > 0 then
                                                 StrColorfila = 3
                                            end if
                                         end if
                                       else
                                          if isdate( RsRep.Fields.Item("DATE_CUSTOM").Value ) then
                                             StrETA_W_H_AUX = SumarDias( SumarDias( RsRep.Fields.Item("DATE_CUSTOM").Value , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                             'StrColorfila = 1
                                          else
                                             StrETA_W_H_AUX = SumarDias( SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                          end if
                                       end if

                                       if StrETA_W_H_AUX then
                                         if isdate(strFechaATAWH ) then
                                             IndFila = DateDiff("d",StrETA_W_H_AUX , strFechaATAWH )
                                             if IndFila <> 0 then
                                                if StrColorfila = 1 then
                                                   if IndFila < 0 then
                                                      StrColorfila = 2
                                                   else
                                                      StrColorfila = 3
                                                   end if
                                                end if
                                             end if
                                         else
                                            IndFila = DateDiff("d", StrETA_W_H_AUX , DATE() )
                                            if IndFila > 0 then
                                               StrColorfila = 3
                                            end if
                                         end if
                                       end if

                                       StrPETA_CUSTOM_CLEARANCE = formatofechaNum(StrETA_CUSTOM_CLEARANCE)
                                       StrPETA_C_P              = formatofechaNum(StrETA_C_P)
                                       if isdate( strETAW_H ) then
                                          StrPETA_W_H 	 = YEAR( strETAW_H ) & Pd(Month( strETAW_H ),2) & Pd(DAY( strETAW_H ),2)  'FECHA DE NOTIFICACION
                                          StrETA_W_H_AUX = StrPETA_WH
                                       else
                                          StrPETA_W_H              = formatofechaNum(StrETA_W_H_AUX)
                                       end if
                                       StrPATA_CP  = "N/A"       'ATA C./P.

                                       '********************************************************************************
                                       '********************************************************************************

                                       if isdate( DFechEntAux ) then
                                          StrPATA_CUSTOM = YEAR( DFechEntAux ) & Pd(Month( DFechEntAux ),2) & Pd(DAY( DFechEntAux ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPATA_CUSTOM = DFechEntAux 'FECHA DE NOTIFICACION
                                       end if

                                       if isdate( RsRep.Fields.Item("RESQUEST_DUTIES").Value ) then
                                          StrPRESQUEST_DUTIES 	 = YEAR( RsRep.Fields.Item("RESQUEST_DUTIES").Value ) & Pd(Month( RsRep.Fields.Item("RESQUEST_DUTIES").Value ),2) & Pd(DAY( RsRep.Fields.Item("RESQUEST_DUTIES").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPRESQUEST_DUTIES = RsRep.Fields.Item("RESQUEST_DUTIES").Value 'FECHA DE NOTIFICACION
                                       end if

                                       if isdate( RsRep.Fields.Item("REVALIDACION").Value ) then
                                          StrPFECHA_DE_REVALIDACION 	 = YEAR( RsRep.Fields.Item("REVALIDACION").Value ) & Pd(Month( RsRep.Fields.Item("REVALIDACION").Value ),2) & Pd(DAY( RsRep.Fields.Item("REVALIDACION").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPFECHA_DE_REVALIDACION = RsRep.Fields.Item("REVALIDACION").Value 'FECHA DE NOTIFICACION
                                       end if

                                       if isdate( RsRep.Fields.Item("PREVIO").Value ) then
                                          StrPFECHA_DE_PREVIO 	 = YEAR( RsRep.Fields.Item("PREVIO").Value ) & Pd(Month( RsRep.Fields.Item("PREVIO").Value ),2) & Pd(DAY( RsRep.Fields.Item("PREVIO").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPFECHA_DE_PREVIO = RsRep.Fields.Item("PREVIO").Value 'FECHA DE NOTIFICACION
                                       end if
                                       if isdate( RsRep.Fields.Item("DATE_CUSTOM").Value ) then
                                          StrPDATE_OF_CLEARANCE 	 = YEAR( RsRep.Fields.Item("DATE_CUSTOM").Value ) & Pd(Month( RsRep.Fields.Item("DATE_CUSTOM").Value ),2) & Pd(DAY( RsRep.Fields.Item("DATE_CUSTOM").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPDATE_OF_CLEARANCE = RsRep.Fields.Item("DATE_CUSTOM").Value 'FECHA DE NOTIFICACION
                                       end if
                                       if isdate( strFechaATAWH ) then
                                          StrPATA_WH 	 = YEAR( strFechaATAWH ) & Pd(Month( strFechaATAWH ),2) & Pd(DAY( strFechaATAWH ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPATA_WH = strFechaATAWH 'FECHA DE NOTIFICACION
                                       end if

                                       strATASPL            = strTimeSLP
                                       StrPTIMEOFDELIVERY   = strHoraATAWH  'TIME OF DELIVERY IN SEM

                                       if strComentarioATAWH <> "" AND ltrim(strComentarioATAWH) <> "" then
                                         strObservaciones = strObservaciones&" ; "& strComentarioATAWH
                                       end if
                                       if strComentarioATAC_P <> "" and ltrim(strComentarioATAC_P) <> "" then
                                         strObservaciones = strObservaciones&" ; "& strComentarioATAC_P
                                       end if
                                       if strComentarioETAW_H <> "" and ltrim(strComentarioETAW_H) <> "" then
                                         strObservaciones = strObservaciones&" ; "& strComentarioETAW_H
                                       end if
                                       if strComentarioATASPLTMP <> "" and ltrim(strComentarioATASPLTMP) <> "" then
                                         strObservaciones = strObservaciones&" ; "& strComentarioATASPLTMP
                                       end if

                                       StrPREMARKS = strObservaciones  'REMARKS

                                       if isdate(strFechaATAWH) then
                                          if not isempty(strFechaATAWH) then
                                             numeroDiasAnio = dateDiff("d",CDate("01/01/"&Datepart("yyyy", strFechaATAWH )), strFechaATAWH )
                                             numeroDiasAnio =    int(numeroDiasAnio/7)+1
                                           else
                                             numeroDiasAnio = 0
                                           end if
                                       else
                                          if isdate(StrETA_W_H_AUX) then
                                             numeroDiasAnio = dateDiff("d",CDate("01/01/"&Datepart("yyyy", StrETA_W_H_AUX )), StrETA_W_H_AUX )
                                             numeroDiasAnio =    int(numeroDiasAnio/7)+1
                                          else
                                             numeroDiasAnio = 0
                                          end if
                                       end if

                                       StrPMODALIDAD         = StrModalidad   'MODALIDAD
                                       StrPWEEK	             = numeroDiasAnio   'WEEK

                                       StrPAMOUNTOFDUTIES	       = strImpuestos 'AMOUNT OF DUTIES
                                       StrPNUM_INVOICECUSTOM	   = strCuentaGastos 'NUM. INVOICE CUSTOM
                                       if isdate( strFecCuentaGastos ) then
                                          StrPDATEINVOICECUSTOM 	 = YEAR( strFecCuentaGastos ) & Pd(Month( strFecCuentaGastos ),2) & Pd(DAY( strFecCuentaGastos ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPDATEINVOICECUSTOM = strFecCuentaGastos 'FECHA DE NOTIFICACION
                                       end if

                                       if isdate(strFechaATAWH) then
                                           if isdate(DFechEntAux) then
                                              intoTD = DiasTrimFinSemana( DFechEntAux ,strFechaATAWH )
                                           else
                                              'intoTD = 0
                                              if isdate( RsRep.Fields.Item("etalax01").Value ) then
                                                intoTD = DiasTrimFinSemana( RsRep.Fields.Item("etalax01").Value , strFechaATAWH )
                                              else
                                                intoTD = 0
                                              end if
                                           end if
                                       else
                                           if isdate(StrETA_W_H_AUX) then
                                              if isdate(DFechEntAux) then
                                                 intoTD = DiasTrimFinSemana( DFechEntAux , StrETA_W_H_AUX )
                                              else
                                                 'intoTD = 0
                                                 if isdate( RsRep.Fields.Item("etalax01").Value ) then
                                                   intoTD = DiasTrimFinSemana( RsRep.Fields.Item("etalax01").Value , StrETA_W_H_AUX )
                                                 else
                                                   intoTD = 0
                                                 end if
                                              end if
                                           else
                                              intoTD = 0
                                           end if
                                       end if

                                       StrPOTD2                  = intoTD 'OTD2

                                       strStatusTmp  = "" ' Exactamnete en donde se encuentra la mercancia
                                       strKPISTTmp  = "" ' Para saber si viene en tiempo o retrasado

                                       if intoTD <= 2 then
                                         strKPISTTmp = "ON TIME"
                                       else
                                         strKPISTTmp = "DELAY"
                                       end if

                                       if strFechaATAWH <> "" then
                                          strStatusTmp = "SEM"
                                       else
                                          if StrPDATE_OF_CLEARANCE <> "" then
                                             strStatusTmp = "ADUANA"
                                          else
                                             if DFechEntAux <> "" then
                                                strStatusTmp = "AEROPUERTO"
                                             else
                                                if RsRep.Fields.Item("feorig01").Value <> "" then
                                                  strStatusTmp = "TRANSITO AEREO"
                                                end if
                                             end if
                                          end if
                                       end if

                                       strRMKATDORIGIN = rmkEtdLoad
                                       strRMKATAPORT   = rmkATAPORT
                                       strRMKDEPACHO   = rmkDSP
                                       strRMKATDRAIL   = rmkRAIL
                                       strRMKCP        = rmkCP

                                       strSTATUS       = strStatusTmp
                                       strLASTRMK      = strLastRMKtmp
                                       strKPISTATUS    = strKPISTTmp

                                       agregarfilaHTML StrColorfila, StrREFERENCIA, StrPOTD2, StrPGUIA_MASTER, StrPGUIA_HOUSE, StrPP_O, StrPFORWARDER_AIR_LINE, StrPAEROPUERTO_SALIDA, StrCUSTOM_OF_DISPATCH, StrPNOTIFICACION_GUIA, StrPFECHA_NOTIFICACION, StrPVESSEL, StrPIMPORT_DOCUMENT, StrPPROVEEDOR, StrPINVOICE, StrPMODEL, StrPDESCRIPCION_COMERCIAL, StrPDESCRIPTION_CODE, StrPQTY, StrPINCOTERMS, StrPSERIAL_NUMBER, StrPCERT_NOM,	StrPORIGIN_ETD, StrPETA_LAX, StrPATA_CUSTOM, StrPRESQUEST_DUTIES, StrPFECHA_DE_REVALIDACION, StrPFECHA_DE_PREVIO, StrPETA_CUSTOM_CLEARANCE, StrPDATE_OF_CLEARANCE, StrPETA_C_P,  StrPATA_CP ,StrPETA_W_H, StrPATA_WH, StrPTIMEOFDELIVERY , StrPREMARKS, StrPMODALIDAD , StrPWEEK, StrPAMOUNTOFDUTIES, StrPNUM_INVOICECUSTOM, StrPDATEINVOICECUSTOM, strADUDESPACHO, strRMKATDORIGIN, strRMKATAPORT, strRMKDEPACHO, strRMKATDRAIL, strRMKCP, strATASPL, strSTATUS, strLASTRMK, strKPISTATUS
                             '*************************************************************




                         end if

                  end if 'if Bolbanrecti = True then

                     RsRep.movenext
                    'Response.Write(strHTML)
                    'Response.End

                    if enproceso( adu_ofi( Session("GAduana") ) ) then
                      banCargaRun=true
                    end if

               Wend

            strHTML = strHTML & "</table>"& chr(13) & chr(10)

            end if

            'response.Write(strHTML)
            'Response.End

            if banCargaRun = false then
                RsRep.close
                Set RsRep = Nothing
                'Se pinta todo el HTML formado
                response.Write(strHTML)
                if strHTML = "" then
                   strHTML = "NO EXISTEN REGISTROS"
                   response.Write(strHTML)
                end if
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
<%end if

ELSE
%>
     <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/carga_activa.asp" -->
<%

END IF

%>




