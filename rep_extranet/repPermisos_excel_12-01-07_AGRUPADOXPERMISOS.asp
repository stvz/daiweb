<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<% Language=VBScript %>

<%
    MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
    Response.Buffer = True
    Response.ContentType = "application/vnd.ms-excel"


    strUsuario = request.Form("user")
    strTipoUsuario = request.Form("TipoUser")

    strPermisos = Request.Form("Permisos")
    permi = PermisoClientes(Session("GAduana"),strPermisos,"clie01")

    if not permi = "" then
       permi = " (" & permi & ") "
    end if

    AplicaFiltro = false
    strFiltroCliente = ""
    strFiltroCliente = request.Form("txtCliente")
    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
      blnAplicaFiltro = true
    end if
    if blnAplicaFiltro then
      permi = "  clie01 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
      permi = ""
    end if

    Dim NumOrders, NumProds, r
    NumOrders = 300
    NumProds = 10

    Dim oSS
    Dim oOrdersSheet
    Dim oTotalsSheet
    Dim HojaMuestras
    Dim HojaVencidos
    Dim oRange
    Dim c
    Dim contadorPag()

    Set oSS = CreateObject("OWC10.Spreadsheet")
    Set c = oSS.Constants
    ' Eliminamos la ultima hoja por si acaso
    oSS.Worksheets(3).Delete
    oSS.Worksheets(2).Delete



       strSQL =" SELECT count(distinct ejec01) as ejecu " & _
               " FROM c01permi " & _
               " WHERE " &permi
'       response.write(strSQL)
'       response.end

       if not trim(strSQL)="" then
          Set RsRep = Server.CreateObject("ADODB.Recordset")
          RsRep.ActiveConnection = MM_EXTRANET_STRING
	   	    RsRep.Source = strSQL
		      RsRep.CursorType = 0

		      RsRep.CursorLocation = 2
  		    RsRep.LockType = 1
	  	    RsRep.Open()

          Intnumejec = 0

  	      if not RsRep.eof then
             ' Comienza el HTML, se pintan los titulos de las columnas
             Intnumejec = RsRep.Fields.Item("ejecu").Value

                strSQLEjec = " SELECT  distinct ejec01 as ejecutivo " & _
                             " FROM c01permi " & _
                             " WHERE " &permi
                Set RsRepEjec = Server.CreateObject("ADODB.Recordset")
                RsRepEjec.ActiveConnection = MM_EXTRANET_STRING
	   	          RsRepEjec.Source = strSQLEjec
		            RsRepEjec.CursorType = 0
		            RsRepEjec.CursorLocation = 2
  		          RsRepEjec.LockType = 1
	  	          RsRepEjec.Open()

'                Response.Write(strSQLEjec)
'                Response.End

                intcontEjec = 1
                if not RsRepEjec.eof then
                    While NOT RsRepEjec.EOF
                        oSS.Worksheets.add()
                        oSS.Worksheets(1).Name = LTRIM(RsRepEjec.Fields.Item("ejecutivo").Value)
                        intcontEjec = intcontEjec + 1
                        RsRepEjec.movenext
                    wend
                end if

                RsRepEjec.close
                Set RsRepEjec = Nothing

                   'Renombrar las hojas
                   'oSS.Worksheets(intcontEjec).Name     = "MUESTRAS_PRU"
                   oSS.Worksheets(intcontEjec).Name = "VENCIDOS"

'                   oSS.Worksheets(intcontEjec).Name     = "PERLA VAZQUEZ"
'                   oSS.Worksheets(intcontEjec + 1).Name = "ANTONIETA CAMPUZANO"



                'Creamos un vector para guardar el numero de fila de cada pagina

                ReDim Preserve contadorPag(intcontEjec)

                For intRow = 1 To Intnumejec + 1
                    contadorPag(intRow - 1) = 3

                    oSS.Worksheets(intRow).Activate
                    'HojaEjec1.Activate
                    With oSS.ActiveSheet
                       .Cells(2, 2).Value = "Producto"
                       .Cells(2, 2).Interior.ColorIndex = 15
                       .Cells(2, 2).Font.size = 8
                       '.Cells(2, 2).Interior.Color = 2

                       .Cells(2, 3).Value = "Cantidad"
                       .Cells(2, 3).Interior.ColorIndex = 15
                       .Cells(2, 3).Font.size = 8
                       '.Cells(2, 3).Interior.Color = 2

                       .Cells(2, 4).Value = "Utilizada"
                       .Cells(2, 4).Interior.ColorIndex = 15
                       .Cells(2, 4).Font.size = 8
                       '.Cells(2, 4).Interior.Color = 2

                       .Cells(2, 5).Value = "Saldo"
                       .Cells(2, 5).Interior.ColorIndex = 15
                       .Cells(2, 5).Font.size = 8
                       '.Cells(2, 5).Interior.Color = 2

                       .Cells(2, 6).Value = "Unidad de Medida"
                       .Cells(2, 6).Interior.ColorIndex = 15
                       .Cells(2, 6).Font.size = 8
                       '.Cells(2, 6).Interior.Color = 2

                       .Cells(2, 7).Value = "Registro Sanitario"
                       .Cells(2, 7).Interior.ColorIndex = 15
                       .Cells(2, 7).Font.size = 8
                       '.Cells(2, 7).Interior.Color = 2

                       .Cells(2, 8).Value = "Permiso No"
                       .Cells(2, 8).Interior.ColorIndex = 15
                       .Cells(2, 8).Font.size = 8
                       '.Cells(2, 8).Interior.Color = 2

                       .Cells(2, 9).Value = "Fecha de Salida"
                       .Cells(2, 9).Interior.ColorIndex = 15
                       .Cells(2, 9).Font.size = 8
                       '.Cells(2, 9).Interior.Color = 2

                       .Cells(2, 10).Value = "Fecha  de Vencimiento"
                       .Cells(2, 10).Interior.ColorIndex = 15
                       .Cells(2, 10).Font.size = 8
                       '.Cells(2, 10).Interior.Color = 2

                       .Cells(2, 11).Value = "Fabricante, Facturador y Proveedor"
                       .Cells(2, 11).Interior.ColorIndex = 15
                       .Cells(2, 11).Font.size = 8
                       '.Cells(2, 11).Interior.Color = 2

                       .Cells(2, 12).Value = "Observaciones"
                       .Cells(2, 12).Interior.ColorIndex = 15
                       .Cells(2, 12).Font.size = 8
                       '.Cells(2, 12).Interior.Color = 2
                    End With
                Next



               '-----------------------------------------------------------------
               ' VAMOS A TRAER LOS DATOS
               '-----------------------------------------------------------------
               '                strSQLPermi = " SELECT PRODUCT01       AS PRODUCTO,       " & _
               '                              "        CANT01          AS CANTIDAD,       " & _
               '                              "        UTIL01          AS UTILIZADA,      " & _
               '                              "        SALDO01         AS SALDO,          " & _
               '                              "        ltrim(descri31) as UnidadMedida,   " & _
               '                              "        PRESENT01,                         " & _
               '                              "        REGSAN01,                          " & _
               '                              "        PERMI01         AS PERMISO,        " & _
               '                              "        FECSAL01        AS FECHASALIDA,    " & _
               '                              "        FECVEN01        AS FECHAVENCI,     " & _
               '                              "        A.PAIPRO22      AS PAISFABRICANTE, " & _
               '                              "        B.PAIPRO22      AS PAISFACTURADOR, " & _
               '                              "        C.PAIPRO22      AS PAISPROVEEDOR,  " & _
               '                              "        CPROV01,   " & _
               '                              "        DIASVEN01, " & _
               '                              "        EJEC01,    " & _
               '                              "        OBSERV01,  " & _
               '                              "        CLIE01     " & _
               '                              " FROM c01permi  LEFT JOIN SSUMED31 ON CLAVEM31 = UMEDT01 " & _
               '                              "      LEFT JOIN SSPROV22  A  ON A.CVEPRO22 = FABRIC01    " & _
               '                              "      LEFT JOIN SSPROV22  B  ON B.CVEPRO22 = FACTUR01    " & _
               '                              "      LEFT JOIN SSPROV22  C  ON C.CVEPRO22 = PROV01      " & _
               '                              " WHERE " &permi

                strSQLPermi = " SELECT PRODUCT01       AS PRODUCTO,       " & _
                              "        CANT01          AS CANTIDAD,       " & _
                              "        UTIL01          AS UTILIZADA,      " & _
                              "        SALDO01         AS SALDO,          " & _
                              "        ltrim(descri31) as UnidadMedida,   " & _
                              "        PRESENT01,                         " & _
                              "        REGSAN01,                          " & _
                              "        PERMI01         AS PERMISO,        " & _
                              "        FECSAL01        AS FECHASALIDA,    " & _
                              "        FECVEN01        AS FECHAVENCI,     " & _
                              "        A.PAIPRO22      AS PAISFABRICANTE, " & _
                              "        B.PAIPRO22      AS PAISFACTURADOR, " & _
                              "        C.PAIPRO22      AS PAISPROVEEDOR,  " & _
                              "        CPROV01,                           " & _
                              "        DIASVEN01,                         " & _
                              "        EJEC01,                            " & _
                              "        OBSERV01,                          " & _
                              "        CLIE01,                            " & _
                              "        SUM(CANCOM02)   AS CANTIDADFAC     " & _
                              " FROM c01permi  LEFT JOIN SSUMED31 ON CLAVEM31   = UMEDT01   " & _
                              "      LEFT JOIN SSPROV22  A        ON A.CVEPRO22 = FABRIC01  " & _
                              "      LEFT JOIN SSPROV22  B        ON B.CVEPRO22 = FACTUR01  " & _
                              "      LEFT JOIN SSPROV22  C        ON C.CVEPRO22 = PROV01    " & _
                              "      LEFT JOIN SSIPAR12           ON NUMPER12   = PERMI01   " & _
                              "      LEFT JOIN SSFRAC02           ON REFCIA02   = REFCIA12   AND ORDFRA02  = ORDFRA12 " & _
                              " WHERE "&permi& _
                              " GROUP BY PERMI01 "


'                strSQLPermi = " SELECT PRODUCT01       AS PRODUCTO,                 " & _
'                              "        CANT01          AS CANTIDAD,                 " & _
'                              "        UTIL01          AS UTILIZADA,                " & _
'                              "        SALDO01         AS SALDO,                    " & _
'                              "        ltrim(descri31) as UnidadMedida,             " & _
'                              "        PRESENT01,                                   " & _
'                              "        REGSAN01,                                    " & _
'                              "        PERMI01         AS PERMISO,                  " & _
'                              "        FECSAL01        AS FECHASALIDA,              " & _
'                              "        FECVEN01        AS FECHAVENCI,               " & _
'                              "        A.PAIPRO22      AS PAISFABRICANTE,           " & _
'                              "        B.PAIPRO22      AS PAISFACTURADOR,           " & _
'                              "        C.PAIPRO22      AS PAISPROVEEDOR,            " & _
'                              "        CPROV01,                                     " & _
'                              "        DIASVEN01,                                   " & _
'                              "        EJEC01,                                      " & _
'                              "        OBSERV01,                                    " & _
'                              "        CLIE01,                                      " & _
'                              "        refcia02,                                    " & _
'                              "        CANCOM02   AS CANTIDADFAC                    " & _
'                              " FROM c01permi  LEFT JOIN SSUMED31 ON CLAVEM31   = UMEDT01   " & _
'                              "                LEFT JOIN SSPROV22  A        ON A.CVEPRO22 = FABRIC01  " & _
'                              "                LEFT JOIN SSPROV22  B        ON B.CVEPRO22 = FACTUR01  " & _
'                              "                LEFT JOIN SSPROV22  C        ON C.CVEPRO22 = PROV01    " & _
'                              "                LEFT JOIN SSIPAR12           ON NUMPER12   = PERMI01   " & _
'                              "                LEFT JOIN SSFRAC02           ON REFCIA02   = REFCIA12   AND ORDFRA02  = ORDFRA12 " & _
'                              " WHERE "&permi


                Set RsRepPermi = Server.CreateObject("ADODB.Recordset")
                RsRepPermi.ActiveConnection = MM_EXTRANET_STRING
	   	          RsRepPermi.Source = strSQLPermi
		            RsRepPermi.CursorType = 0
		            RsRepPermi.CursorLocation = 2
  		          RsRepPermi.LockType = 1
	  	          RsRepPermi.Open()
'                Response.Write(strSQLPermi)
'                Response.End

                'intUtilizada2 = 0

                if not RsRepPermi.eof then
                    While NOT RsRepPermi.EOF

                        intUtilizada2  = 0
                        intCantidad    = 0
                        strProducto    = RsRepPermi.Fields.Item("PRODUCTO").Value
                        intCantidad    = RsRepPermi.Fields.Item("CANTIDAD").Value
                        intUtilizada   = RsRepPermi.Fields.Item("UTILIZADA").Value
                        intSaldo       = RsRepPermi.Fields.Item("SALDO").Value
                        strUniMed      = RsRepPermi.Fields.Item("UnidadMedida").Value
                        strPresent     = RsRepPermi.Fields.Item("PRESENT01").Value
                        strregsan      = RsRepPermi.Fields.Item("REGSAN01").Value
                        strPermi       = RsRepPermi.Fields.Item("PERMISO").Value
                        dFechaSali     = RsRepPermi.Fields.Item("FECHASALIDA").Value
                        dFechaVenci    = RsRepPermi.Fields.Item("FECHAVENCI").Value
                        strPaisFabric  = RsRepPermi.Fields.Item("PAISFABRICANTE").Value
                        strPaisFactur  = RsRepPermi.Fields.Item("PAISFACTURADOR").Value
                        strPaisProv    = RsRepPermi.Fields.Item("PAISPROVEEDOR").Value
                        intDiasVenci   = RsRepPermi.Fields.Item("DIASVEN01").Value
                        strEjecu       = RsRepPermi.Fields.Item("EJEC01").Value
                        strObserv      = RsRepPermi.Fields.Item("OBSERV01").Value
                        intcvecli      = RsRepPermi.Fields.Item("CLIE01").Value
                        intUtilizada2  = RsRepPermi.Fields.Item("CANTIDADFAC").Value

                        if isnull(intUtilizada) then
                             intUtilizada = 0
                        end if

                        if isnull(intUtilizada2) then
                             intUtilizada2 = 0
                        end if

                        if isnull(intCantidad) then
                             intCantidad = 0
                        end if

                        if isnull(intSaldo) then
                             intSaldo = 0
                        end if

                        if( (intUtilizada) > (intUtilizada2) ) then
                           intUtilizada2=(intUtilizada)
                        end if

                        if intSaldo = 0 then
                           intSaldo = intCantidad - intUtilizada2
                        end if


                        IF isnull(dFechaVenci) then ' si no le capturaron la fecha de vencimiento
                                dFechaVenci = dFechaSali + 179
                        end if

                        if NOT isnull(strPresent) then
                          strUniMed = strUniMed & strPresent
                        end if


                        '--------------------------------------------
                        'PRIMERO CHEKAR VENCIDOS, DESPUES CHEKAR EJECUTIVOS Y LOS QUE QUEDEN SERAN MUESTRAS
                        '--------------------------------------------

                        '--------------------------------------------
                        'PARA CHEKAR VENCIDOS, PRIMERO BUSCAMOS POR FECHA DE VENCIMIENTO, DESPUES POR CANTIDAD UTILIZADA
                        '--------------------------------------------

                        dfechaactual =  date()
                        if dFechaVenci < dfechaactual then ' tiene una fecha vencida
                            oSS.Worksheets(intcontEjec).Activate
                            With oSS.ActiveSheet
                                .Cells(contadorPag(intcontEjec-1), 2).Value  = strProducto
                                .Cells(contadorPag(intcontEjec-1), 2).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 2).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 3).Value  = intCantidad
                                .Cells(contadorPag(intcontEjec-1), 3).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 3).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 4).Value  = intUtilizada2
                                .Cells(contadorPag(intcontEjec-1), 4).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 4).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 5).Value  = intSaldo
                                .Cells(contadorPag(intcontEjec-1), 5).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 5).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 6).Value  = strUniMed
                                .Cells(contadorPag(intcontEjec-1), 6).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 6).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 7).Value  = strregsan
                                .Cells(contadorPag(intcontEjec-1), 7).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 7).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 8).Value  = strPermi
                                .Cells(contadorPag(intcontEjec-1), 8).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 8).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 9).Value  =  FormatDateTime(dFechaSali, vbGeneralDate)
                                .Cells(contadorPag(intcontEjec-1), 9).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 9).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 10).Value = FormatDateTime(dFechaVenci, vbGeneralDate)
                                .Cells(contadorPag(intcontEjec-1), 10).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 10).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 11).Value = strPaisFabric&","&strPaisFactur&","&strPaisProv
                                .Cells(contadorPag(intcontEjec-1), 11).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 11).Font.size = 8
                                .Cells(contadorPag(intcontEjec-1), 12).Value = strObserv
                                .Cells(contadorPag(intcontEjec-1), 12).Interior.ColorIndex = 3
                                .Cells(contadorPag(intcontEjec-1), 12).Font.size = 8
                                contadorPag(intcontEjec-1) = contadorPag(intcontEjec-1) + 1
                               ' .Cells(contadorPag(intcontEjec), 19).Value = contadorPag(intcontEjec)
                            End With
                        else ' No esta vencido por fecha, verificar la cantidad utilizada para ver si no se ha vencido por cantidad
                             if (intUtilizada2)>=(intCantidad) then
                                  oSS.Worksheets(intcontEjec).Activate
                                  With oSS.ActiveSheet
                                      .Cells(contadorPag(intcontEjec-1), 2).Value  = strProducto
                                      .Cells(contadorPag(intcontEjec-1), 2).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 2).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 3).Value  = intCantidad
                                      .Cells(contadorPag(intcontEjec-1), 3).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 3).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 4).Value  = intUtilizada2
                                      .Cells(contadorPag(intcontEjec-1), 4).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 4).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 5).Value  = intSaldo
                                      .Cells(contadorPag(intcontEjec-1), 5).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 5).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 6).Value  = strUniMed
                                      .Cells(contadorPag(intcontEjec-1), 6).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 6).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 7).Value  = strregsan
                                      .Cells(contadorPag(intcontEjec-1), 7).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 7).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 8).Value  = strPermi
                                      .Cells(contadorPag(intcontEjec-1), 8).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 8).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 9).Value  = FormatDateTime(dFechaSali, vbGeneralDate)
                                      .Cells(contadorPag(intcontEjec-1), 9).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 9).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 10).Value = FormatDateTime(dFechaVenci, vbGeneralDate)
                                      .Cells(contadorPag(intcontEjec-1), 10).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 10).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 11).Value = strPaisFabric&","&strPaisFactur&","&strPaisProv
                                      .Cells(contadorPag(intcontEjec-1), 11).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 11).Font.size = 8
                                      .Cells(contadorPag(intcontEjec-1), 12).Value = strObserv
                                      .Cells(contadorPag(intcontEjec-1), 12).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 12).Font.size = 8

                                      'if isnumeric(intUtilizada2) then
                                      '   .Cells(contadorPag(intcontEjec), 13).Value = TypeName(intUtilizada2)
                                      '   .Cells(contadorPag(intcontEjec), 13).Value = TypeName(dFechaSali)
                                      '   .Cells(contadorPag(intcontEjec), 13).Value = TypeName(dFechaVenci)
                                      'end if
                                      'if isnumeric((intCantidad)) then
                                      '   .Cells(contadorPag(intcontEjec), 14).Value = TypeName(intCantidad)
                                      'end if

                                  End With
                                  contadorPag(intcontEjec-1) = contadorPag(intcontEjec-1) + 1
                             else ' NO ESTA VENCIDO, CHEKAR A QUE EJECUTIVO LE CORRESPONDE

                                 intbanEjec = 0
                                 For intRow = 1 To Intnumejec
                                     IF oSS.Worksheets(intRow).Name  = LTRIM(strEjecu) THEN
                                         intbanEjec = 1
                                         oSS.Worksheets(intRow).Activate
                                         With oSS.ActiveSheet
                                             .Cells(contadorPag(intRow - 1), 2).Value  = strProducto
                                             .Cells(contadorPag(intRow - 1), 2).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 3).Value  = intCantidad
                                             .Cells(contadorPag(intRow - 1), 3).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 4).Value  = intUtilizada2
                                             .Cells(contadorPag(intRow - 1), 4).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 5).Value  = intSaldo
                                             .Cells(contadorPag(intRow - 1), 5).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 6).Value  = strUniMed
                                             .Cells(contadorPag(intRow - 1), 6).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 7).Value  = strregsan
                                             .Cells(contadorPag(intRow - 1), 7).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 8).Value  = strPermi
                                             .Cells(contadorPag(intRow - 1), 8).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 9).Value  = FormatDateTime(dFechaSali, vbGeneralDate)
                                             .Cells(contadorPag(intRow - 1), 9).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 10).Value = FormatDateTime(dFechaVenci, vbGeneralDate)
                                             .Cells(contadorPag(intRow - 1), 10).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 11).Value = strPaisFabric&","&strPaisFactur&","&strPaisProv
                                             .Cells(contadorPag(intRow - 1), 11).Font.size = 8
                                             .Cells(contadorPag(intRow - 1), 12).Value = strObserv
                                             .Cells(contadorPag(intRow - 1), 12).Font.size = 8
                                             contadorPag(intRow - 1) = contadorPag(intRow - 1) + 1
                                         End With
                                     END IF
                                 Next
                                 if intbanEjec = 0 then ' Para la hoja de muestras
                                     oSS.Worksheets(intcontEjec).Activate
                                     With oSS.ActiveSheet
                                        .Cells(contadorPag(intcontEjec-1), 2).Value  = strProducto
                                        .Cells(contadorPag(intcontEjec-1), 2).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 2).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 3).Value  = intCantidad2
                                        .Cells(contadorPag(intcontEjec-1), 3).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 3).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 4).Value  = intUtilizada
                                        .Cells(contadorPag(intcontEjec-1), 4).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 4).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 5).Value  = intSaldo
                                        .Cells(contadorPag(intcontEjec-1), 5).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 5).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 6).Value  = strUniMed
                                        .Cells(contadorPag(intcontEjec-1), 6).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 6).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 7).Value  = strregsan
                                        .Cells(contadorPag(intcontEjec-1), 7).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 7).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 8).Value  = strPermi
                                        .Cells(contadorPag(intcontEjec-1), 8).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 8).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 9).Value  = FormatDateTime(dFechaSali, vbGeneralDate)
                                        .Cells(contadorPag(intcontEjec-1), 9).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 9).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 10).Value = FormatDateTime(dFechaVenci, vbGeneralDate)
                                        .Cells(contadorPag(intcontEjec-1), 10).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 10).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 11).Value = strPaisFabric&","&strPaisFactur&","&strPaisProv
                                        .Cells(contadorPag(intcontEjec-1), 11).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 11).Font.size = 8
                                        .Cells(contadorPag(intcontEjec-1), 12).Value = strObserv
                                        .Cells(contadorPag(intcontEjec-1), 12).Interior.ColorIndex = 3
                                        .Cells(contadorPag(intcontEjec-1), 12).Font.size = 8
                                     End With
                                     contadorPag(intcontEjec-1) = contadorPag(intcontEjec-1) + 1
                                 end if

                             end if
                        end if

                        RsRepPermi.movenext
                    wend
                end if





                RsRepPermi.close
                Set RsRepPermi = Nothing
          end if





          RsRep.close
          Set RsRep = Nothing

       end if

'    oSS.DisplayToolbar = False
    oSS.AutoFit = True
    'oOrdersSheet.Activate

    Response.Write oSS.XMLData
    'Response.Write oSS.CSVData




    Response.End





'          if Intnumejec > 0 then
'             select case Intnumejec
'                case 1
'                   'agregar hoja
'                   oSS.Worksheets.add()
'                   'Renombrar las hojas
'                   oSS.Worksheets(1).Name = "Ejecutivo1"
'                   oSS.Worksheets(2).Name = "Muestras"
'                   oSS.Worksheets(3).Name = "Vencidos"
'                case 2
'                   'agregar hojas
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   'Renombrar las hojas
'                   oSS.Worksheets(1).Name = "Ejecutivo1"
'                   oSS.Worksheets(2).Name = "Ejecutivo2"
'                   oSS.Worksheets(3).Name = "Muestras"
'                   oSS.Worksheets(4).Name = "Vencidos"
'                case 3
'                   'agregar hojas
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   'Renombrar las hojas
'                   oSS.Worksheets(1).Name = "Ejecutivo1"
'                   oSS.Worksheets(2).Name = "Ejecutivo2"
'                   oSS.Worksheets(3).Name = "Ejecutivo3"
'                   oSS.Worksheets(4).Name = "Muestras"
'                   oSS.Worksheets(5).Name = "Vencidos"
'                case 4
'                   'agregar hojas
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   'Renombrar las hojas
'                   oSS.Worksheets(1).Name = "Ejecutivo1"
'                   oSS.Worksheets(2).Name = "Ejecutivo2"
'                   oSS.Worksheets(3).Name = "Ejecutivo3"
'                   oSS.Worksheets(4).Name = "Ejecutivo4"
'                   oSS.Worksheets(5).Name = "Muestras"
'                   oSS.Worksheets(6).Name = "Vencidos"
'                case 5
'                   'agregar hojas
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   'Renombrar las hojas
'                   oSS.Worksheets(1).Name = "Ejecutivo1"
'                   oSS.Worksheets(2).Name = "Ejecutivo2"
'                   oSS.Worksheets(3).Name = "Ejecutivo3"
'                   oSS.Worksheets(4).Name = "Ejecutivo4"
'                   oSS.Worksheets(5).Name = "Ejecutivo5"
'                   oSS.Worksheets(6).Name = "Muestras"
'                   oSS.Worksheets(7).Name = "Vencidos"
'                case 6
'                   'agregar hojas
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   oSS.Worksheets.add()
'                   'Renombrar las hojas
'                   oSS.Worksheets(1).Name = "Ejecutivo1"
'                   oSS.Worksheets(2).Name = "Ejecutivo2"
'                   oSS.Worksheets(3).Name = "Ejecutivo3"
'                   oSS.Worksheets(4).Name = "Ejecutivo4"
'                   oSS.Worksheets(5).Name = "Ejecutivo5"
'                   oSS.Worksheets(6).Name = "Ejecutivo6"
'                   oSS.Worksheets(7).Name = "Muestras"
'                   oSS.Worksheets(8).Name = "Vencidos"
'             end select
'          end if


               '      strSQL =" SELECT PRODUCT01       AS PRODUCTO,       " & _
               '              "        CANT01          AS CANTIDAD,       " & _
               '              "        UTIL01          AS UTILIZADA,      " & _
               '              "        SALDO01         AS SALDO,          " & _
               '              "        ltrim(descri31) as UnidadMedida,   " & _
               '              "        PRESENT01,                         " & _
               '              "        REGSAN01,                          " & _
               '              "        PERMI01         AS PERMISO,        " & _
               '              "        FECSAL01        AS FECHASALIDA,    " & _
               '              "        FECVEN01        AS FECHAVENCI,     " & _
               '              "        A.PAIPRO22      AS PAISFABRICANTE, " & _
               '              "        B.PAIPRO22      AS PAISFACTURADOR, " & _
               '              "        C.PAIPRO22      AS PAISPROVEEDOR,  " & _
               '              "        CPROV01,   " & _
               '              "        DIASVEN01, " & _
               '              "        EJEC01,    " & _
               '              "        OBSERV01,  " & _
               '              "        CLIE01     " & _
               '              " FROM c01permi  LEFT JOIN SSUMED31 ON CLAVEM31 = UMEDT01 " & _
               '              "      LEFT JOIN SSPROV22  A  ON A.CVEPRO22 = FABRIC01    " & _
               '              "      LEFT JOIN SSPROV22  B  ON B.CVEPRO22 = FACTUR01    " & _
               '              "      LEFT JOIN SSPROV22  C  ON C.CVEPRO22 = PROV01      " & _
               '              " WHERE " &permi

               ' SELECT REFCIA12 ,ORDFRA12, NUMPER12, CANCOM02,U_MEDC02
               ' FROM SSIPAR12, SSFRAC02
               ' WHERE REFCIA02    = REFCIA12    AND
               '         ORDFRA02  = ORDFRA12  AND
               '          NUMPER12 <>''
               ' ORDER BY  NUMPER12

               '  SELECT NUMPER12,
               '                   SUM(CANCOM02) AS CANTIDAD
               '  FROM SSIPAR12,
               '               SSFRAC02,
               '               SSDAGI01
               '  WHERE NUMPER12='000778'  AND
               '                  CVEPED01 <> 'R1'       AND
               '                  REFCIA02    = REFCIA12    AND
               '                  ORDFRA02  = ORDFRA12  AND
               '                  REFCIA02    = REFCIA01
               '  GROUP BY NUMPER12
               '  UNION
               '  SELECT NUMPER12,
               '                   SUM(CANCOM02) AS CANTIDAD
               '  FROM SSIPAR12,
               '               SSFRAC02,
               '               SSDAGE01
               '  WHERE NUMPER12='000778'  AND
               '                  CVEPED01 <> 'R1'       AND
               '                  REFCIA02    = REFCIA12    AND
               '                  ORDFRA02  = ORDFRA12  AND
               '                  REFCIA02    = REFCIA01
               '  GROUP BY NUMPER12

          ' agregar una hoja
          ' oSS.Worksheets.add()
          ' oSS.Worksheets.add()

          'Rename Sheet1 to "Orders", rename Sheet2 to "Totals" and remove Sheet3
          '    Set HojaEjec1 = oSS.Worksheets(1)
          '    HojaEjec1.Name = "Ejecutivo1"
          '    Set HojaEjec2 = oSS.Worksheets(2)
          '    HojaEjec2.Name = "Ejecutivo2"
          '    Set HojaMuestras = oSS.Worksheets(3)
          '    HojaMuestras.Name = "Muestras"
          '    Set HojaVencidos = oSS.Worksheets(4)
          '    HojaVencidos.Name = "Vencidos"

        '=== Build the Second Worksheet (Totals) ===========================================

'      For intRow = 1 To 100
'         For intCol = 1 To 10
'            .Cells(intRow, intCol).Value = (intRow - intCol) / pintDivisor
'            If .Cells(intRow, intCol).Value Mod 3 = 0 Then
'               .Cells(intRow, intCol).Interior.Color = pstrColor
'            End If
'         Next
'         .Cells(intRow, 11).Value = "= I" & CStr(intRow) & "+J" & CStr(intRow)
'         If intRow Mod 2 = 0 Then .Cells(intRow, 11).Interior.Color = "LightGray"
'      Next
'      .Columns("A:D").AutoFilter

'    'Change the Column headings and hide row headings
'    HojaEjec1.Activate
'    oSS.Windows(1).ColumnHeadings(1).Caption  = "Producto"
'    oSS.Windows(1).ColumnHeadings(2).Caption  = "Cantidad"
'    oSS.Windows(1).ColumnHeadings(3).Caption  = "Utilizada"
'    oSS.Windows(1).ColumnHeadings(4).Caption  = "Saldo"
'    oSS.Windows(1).ColumnHeadings(5).Caption  = "Unidad de Medida"
'    oSS.Windows(1).ColumnHeadings(6).Caption  = "Registro Sanitario"
'    oSS.Windows(1).ColumnHeadings(7).Caption  = "Permiso No"
'    oSS.Windows(1).ColumnHeadings(8).Caption  = "Fecha de Salida"
'    oSS.Windows(1).ColumnHeadings(9).Caption  = "Fecha  de Vencimiento"
'    oSS.Windows(1).ColumnHeadings(10).Caption = "Fabricante, Facturador y Proveedor"
'    oSS.Windows(1).ColumnHeadings(11).Caption = "Observaciones"
'    oSS.Windows(1).DisplayRowHeadings = False

'    oSS.Windows(1).ColumnHeadings(1).Caption = "Product ID"
'    oSS.Windows(1).ColumnHeadings(2).Caption = "Total"
'    oSS.Windows(1).DisplayRowHeadings = False

'    'Add the product IDs to column 1
'    Dim aProductIDs
'    aProductIDs = GetProductIDs
'    oTotalsSheet.Range("A1:A" & NumProds).Value = aProductIDs
'    oTotalsSheet.Range("A1:A" & NumProds).HorizontalAlignment = c.xlHAlignCenter

'    'Add a formula to column 2 that computes totals per product from the Orders Sheet
'    oTotalsSheet.Range("B1:B" & NumProds).Formula = _
'        "=SUMIF(Ejecutivo1!B$2:B$" & NumOrders + 1 & ",A1,Ejecutivo1!F$2:F$" & NumOrders + 1 & ")"
'    oTotalsSheet.Range("B1:B" & NumProds).NumberFormat = "_(  $* #,##0.00   _)"

'    'Apply window settings for the Totals worksheet
'    oSS.Windows(1).ViewableRange = oTotalsSheet.UsedRange.Address

'-----------------------------------------------------------------------------------------------------------------------
    'oSS.Worksheets(3).Delete

    '===================== Build the First Worksheet (Orders) ===========================
    'Add headings to A1:F1 of the Orders worksheet and apply formatting
'    Set oRange = oOrdersSheet.Range("A1:F1")
'    oRange.Value = Array("Order Number", "Product ID", "Quantity", "Price", "Discount", "Total")
'    oRange.Font.Bold = True
'    oRange.Interior.Color = "Silver"
'    oRange.Borders(c.xlEdgeBottom).Weight = c.xlThick
'    oRange.HorizontalAlignment = c.xlHAlignCenter

'    'Apply formatting to the columns
'    oOrdersSheet.Range("A:A").ColumnWidth = 20
'    oOrdersSheet.Range("B:E").ColumnWidth = 15
'    oOrdersSheet.Range("F:F").ColumnWidth = 20
'    oOrdersSheet.Range("A2:E" & NumOrders + 1 _
'        ).HorizontalAlignment = c.xlHAlignCenter
'    oOrdersSheet.Range("D2:D" & NumOrders + 1).NumberFormat = "0.00"
'    oOrdersSheet.Range("E2:E" & NumOrders + 1).NumberFormat = "0 % "
'    oOrdersSheet.Range("F2:F" & NumOrders + 1).NumberFormat = "$ 0.00" '"_($* #,##0.00_)"

'    'Obtain the order information for the first five columns in the Orders worksheet
'    'and populate the worksheet with that data starting at row 2
'    Dim aOrderData
'    aOrderData = GetOrderInfo
'    oOrdersSheet.Range("A2:E" & NumOrders + 1).Value = aOrderData

'    'Add a formula to calculate the order total for each row and format the column
'    oOrdersSheet.Range("F2:F" & NumOrders + 1).Formula = "=C2*D2*(1-E2)"
'        oOrdersSheet.Range("F2:F" & NumOrders + 1).NumberFormat = "_(  $* #,##0.00   _)"

'    'Apply a border to the used rows
'    oOrdersSheet.UsedRange.Borders(c.xlInsideHorizontal).Weight = c.xlThin
'    oOrdersSheet.UsedRange.BorderAround , c.xlThin, 15

'    'Turn on AutoFilter and display an initial criteria where
'    'the Product ID (column 2) is equal to 5
'    oOrdersSheet.UsedRange.AutoFilter
'    oOrdersSheet.AutoFilter.Filters(2).Criteria.FilterFunction = c.ssFilterFunctionInclude
'    oOrdersSheet.AutoFilter.Filters(2).Criteria.Add "5"
'    oOrdersSheet.AutoFilter.Apply

    'Add a Subtotal at the end of the usedrange
'    oOrdersSheet.Range("F" & NumOrders + 3).Formula = "=SUBTOTAL(9, F2:F" & NumOrders + 1 & ")"

'    'Apply window settings for the Orders worksheet
'    oOrdersSheet.Activate   'Makes the Orders sheet active
'    oSS.Windows(1).ViewableRange = oOrdersSheet.UsedRange.Address
'    oSS.Windows(1).DisplayRowHeadings = False
'    oSS.Windows(1).DisplayColumnHeadings = False
'    oSS.Windows(1).FreezePanes = True
'    oSS.Windows(1).DisplayGridlines = False

    '=== Build the Second Worksheet (Totals) ===========================================

'    'Change the Column headings and hide row headings
'    oTotalsSheet.Activate
'    oSS.Windows(1).ColumnHeadings(1).Caption = "Product ID"
'    oSS.Windows(1).ColumnHeadings(2).Caption = "Total"
'    oSS.Windows(1).DisplayRowHeadings = False

'    'Add the product IDs to column 1
'    Dim aProductIDs
'    aProductIDs = GetProductIDs
'    oTotalsSheet.Range("A1:A" & NumProds).Value = aProductIDs
'    oTotalsSheet.Range("A1:A" & NumProds).HorizontalAlignment = c.xlHAlignCenter

'    'Add a formula to column 2 that computes totals per product from the Orders Sheet
'    oTotalsSheet.Range("B1:B" & NumProds).Formula = _
'        "=SUMIF(Ejecutivo1!B$2:B$" & NumOrders + 1 & ",A1,Ejecutivo1!F$2:F$" & NumOrders + 1 & ")"
'    oTotalsSheet.Range("B1:B" & NumProds).NumberFormat = "_(  $* #,##0.00   _)"

'    'Apply window settings for the Totals worksheet
'    oSS.Windows(1).ViewableRange = oTotalsSheet.UsedRange.Address

    '=== Setup for final presentation ==================================================


%>
