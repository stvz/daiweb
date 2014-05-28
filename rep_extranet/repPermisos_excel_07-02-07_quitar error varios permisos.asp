<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<% Language=VBScript %>

<%
    MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
    Response.Buffer = True
    Response.ContentType = "application/vnd.ms-excel"

    Server.ScriptTimeOut=100000


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



       strSQL =" SELECT count(distinct ltrim(ejec01) ) as ejecu " & _
               " FROM c01permi " & _
               " WHERE "&permi
'               " WHERE "&permi&" or (permi01='I802982006')"
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

                strSQLEjec = " SELECT  distinct ltrim(ejec01) as ejecutivo " & _
                             " FROM c01permi " & _
                             " WHERE "&permi
'                             " WHERE "&permi&" or (permi01='I802982006')"
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

                       .Cells(2, 13).Value = "Referencia"
                       .Cells(2, 13).Interior.ColorIndex = 15
                       .Cells(2, 13).Font.size = 8
                       '.Cells(2, 13).Interior.Color = 2

                    End With
                Next

                'Dim myarray()
                'ReDim myarray(1, 10)
                'ReDim  myarray(2, 10)

               '-----------------------------------------------------------------
               ' VAMOS A TRAER LOS DATOS
               '-----------------------------------------------------------------

                'strSQLPermi = " SELECT PRODUCT01       AS PRODUCTO,                 " & _
                '              "        CANT01          AS CANTIDAD,                 " & _
                '              "        UTIL01          AS UTILIZADA,                " & _
                '              "        SALDO01         AS SALDO,                    " & _
                '              "        ltrim(descri31) as UnidadMedida,             " & _
                '              "        PRESENT01,                                   " & _
                '              "        REGSAN01,                                    " & _
                '              "        PERMI01         AS PERMISO,                  " & _
                '              "        FECSAL01        AS FECHASALIDA,              " & _
                '              "        FECVEN01        AS FECHAVENCI,               " & _
                '              "        A.PAIPRO22      AS PAISFABRICANTE,           " & _
                '              "        B.PAIPRO22      AS PAISFACTURADOR,           " & _
                '              "        C.PAIPRO22      AS PAISPROVEEDOR,            " & _
                '              "        CPROV01,                                     " & _
                '              "        DIASVEN01,                                   " & _
                '              "        EJEC01,                                      " & _
                '              "        OBSERV01,                                    " & _
                '              "        CLIE01,                                      " & _
                '              "        refcia02,                                    " & _
                '              "        CANCOM02   AS CANTIDADFAC                    " & _
                '              " FROM c01permi  LEFT JOIN SSUMED31 ON CLAVEM31   = UMEDT01   " & _
                '              "                LEFT JOIN SSPROV22  A        ON A.CVEPRO22 = FABRIC01  " & _
                '              "                LEFT JOIN SSPROV22  B        ON B.CVEPRO22 = FACTUR01  " & _
                '              "                LEFT JOIN SSPROV22  C        ON C.CVEPRO22 = PROV01    " & _
                '              "                LEFT JOIN SSIPAR12           ON NUMPER12   = PERMI01 AND  PERMI01 is not null   " & _
                '              "                LEFT JOIN SSFRAC02           ON REFCIA02   = REFCIA12   AND ORDFRA02  = ORDFRA12 " & _
                '              " WHERE "&permi



             strSQLPermi =  " SELECT         PERMI01         AS PERMISO,                 " & _
                            "                FECSAL01        AS FECHASALIDA,             " & _
                            "                PRODUCT01       AS PRODUCTO,                " & _
                            "                CANT01          AS CANTIDAD,                " & _
                            "                OBSERV01,                                   " & _
                            "                UTIL01          AS UTILIZADA,               " & _
                            "                SALDO01         AS SALDO,                   " & _
                            "                ltrim(descri31) as UnidadMedida,            " & _
                            "                PRESENT01,                                  " & _
                            "                REGSAN01,                                   " & _
                            "                PERMI01         AS PERMISO,                 " & _
                            "                FECVEN01        AS FECHAVENCI,              " & _
                            "                A.PAIPRO22      AS PAISFABRICANTE,          " & _
                            "                B.PAIPRO22      AS PAISFACTURADOR,          " & _
                            "                C.PAIPRO22      AS PAISPROVEEDOR,           " & _
                            "                CPROV01,                                    " & _
                            "                DIASVEN01,                                  " & _
                            "                EJEC01,                                     " & _
                            "                CLIE01,                                     " & _
                            "                refcia02,                                   " & _
                            "                CANCOM02   AS CANTIDADFAC                   " & _
                            "   FROM c01permi  LEFT JOIN SSUMED31 ON CLAVEM31   = UMEDT01             " & _
                            "                  LEFT JOIN SSPROV22  A        ON A.CVEPRO22 = FABRIC01  " & _
                            "                  LEFT JOIN SSPROV22  B        ON B.CVEPRO22 = FACTUR01  " & _
                            "                  LEFT JOIN SSPROV22  C        ON C.CVEPRO22 = PROV01    " & _
                            "                  LEFT JOIN SSIPAR12           ON NUMPER12   = PERMI01   AND PERMI01 is not null  " & _
                            "                  LEFT JOIN SSFRAC02           ON REFCIA02   = REFCIA12  AND ORDFRA02  = ORDFRA12 " & _
                            " WHERE "&permi & _
                            " ORDER BY  PERMISO, FECHASALIDA, PRODUCTO, CANTIDAD, OBSERV01, EJEC01, cantidadfac "



'                              " WHERE "&permi&" or (permi01='I802982006')"
                              '+------------+
'                              | PERMISO    |
'                              +------------+
'                              | I904822006 |
'                              | I904602006 |
'                              | I904092006 |
'                              +------------+



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

                Dim detallepermiso()

                intcontpermi = 1
                strcompermi=""
                strObservaciones=""
                dblSumCantidadUtil=0

                if not RsRepPermi.eof then
                    'Vamos a leer el primer registro
                    ReDim detallepermiso(20,intcontpermi)
                      if not isnull(RsRepPermi.Fields.Item("PERMISO").Value) then
                          strcompermi                     = ltrim(RsRepPermi.Fields.Item("PERMISO").Value)
                          detallepermiso(0,intcontpermi-1)  = ltrim(RsRepPermi.Fields.Item("PERMISO").Value)
                      else
                          strcompermi                     = " "
                          detallepermiso(0,intcontpermi-1)  = (RsRepPermi.Fields.Item("PERMISO").Value)
                      end if
                      if not isnull(RsRepPermi.Fields.Item("PRODUCTO").Value) then
                         detallepermiso(1,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("PRODUCTO").Value)
                      else
                         detallepermiso(1,intcontpermi-1)  = (RsRepPermi.Fields.Item("PRODUCTO").Value)
                      end if
                      if not isnull(RsRepPermi.Fields.Item("CANTIDAD").Value) then
                         detallepermiso(2,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value )
                      else
                         detallepermiso(2,intcontpermi-1)  = (RsRepPermi.Fields.Item("CANTIDAD").Value )
                      end if
                      if not isnull(RsRepPermi.Fields.Item("UTILIZADA").Value) then
                         detallepermiso(3,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("UTILIZADA").Value )
                      else
                         detallepermiso(3,intcontpermi-1)  = (RsRepPermi.Fields.Item("UTILIZADA").Value )
                      end if
                      'if not isnull(RsRepPermi.Fields.Item("SALDO").Value) then
                      '   detallepermiso(4,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("SALDO").Value )
                      'else
                      '   detallepermiso(4,intcontpermi-1)  = (RsRepPermi.Fields.Item("SALDO").Value )
                      'end if
                      if not isnull(RsRepPermi.Fields.Item("UnidadMedida").Value) then
                         detallepermiso(5,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("UnidadMedida").Value )
                      else
                         detallepermiso(5,intcontpermi-1)  = (RsRepPermi.Fields.Item("UnidadMedida").Value )
                      end if
                      if not isnull(RsRepPermi.Fields.Item("PRESENT01").Value) then
                         detallepermiso(6,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("PRESENT01").Value )
                      else
                         detallepermiso(6,intcontpermi-1)  = (RsRepPermi.Fields.Item("PRESENT01").Value )
                      end if
                      if not isnull(RsRepPermi.Fields.Item("REGSAN01").Value) then
                         detallepermiso(7,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("REGSAN01").Value )
                      else
                         detallepermiso(7,intcontpermi-1)  = (RsRepPermi.Fields.Item("REGSAN01").Value )
                      end if
                      if not isnull(RsRepPermi.Fields.Item("PERMISO").Value) then
                         detallepermiso(8,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("PERMISO").Value )
                      else
                         detallepermiso(8,intcontpermi-1)  = (RsRepPermi.Fields.Item("PERMISO").Value )
                      end if
                      if not isnull(RsRepPermi.Fields.Item("FECHASALIDA").Value) then
                         detallepermiso(9,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("FECHASALIDA").Value)
                      else
                         detallepermiso(9,intcontpermi-1)  = (RsRepPermi.Fields.Item("FECHASALIDA").Value)
                      end if
                      if not isnull(RsRepPermi.Fields.Item("FECHAVENCI").Value) then
                         detallepermiso(10,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("FECHAVENCI").Value )
                      else
                         detallepermiso(10,intcontpermi-1) = (RsRepPermi.Fields.Item("FECHAVENCI").Value )
                      end if
                      if not isnull(RsRepPermi.Fields.Item("PAISFABRICANTE").Value) then
                         detallepermiso(11,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("PAISFABRICANTE").Value )
                      else
                         detallepermiso(11,intcontpermi-1) = " "
                      end if
                      if not isnull(RsRepPermi.Fields.Item("PAISFACTURADOR").Value) then
                         detallepermiso(12,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("PAISFACTURADOR").Value )
                      else
                         detallepermiso(12,intcontpermi-1) = " "
                      end if
                      if not isnull(RsRepPermi.Fields.Item("PAISPROVEEDOR").Value) then
                         detallepermiso(13,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("PAISPROVEEDOR").Value )
                      else
                         detallepermiso(13,intcontpermi-1) = " "
                      end if
                      if not isnull(RsRepPermi.Fields.Item("DIASVEN01").Value) then
                         detallepermiso(14,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("DIASVEN01").Value )
                      else
                         detallepermiso(14,intcontpermi-1) = (RsRepPermi.Fields.Item("DIASVEN01").Value )
                      end if
                      if not isnull(RsRepPermi.Fields.Item("EJEC01").Value) then
                         detallepermiso(15,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("EJEC01").Value )
                      else
                         detallepermiso(15,intcontpermi-1) = (RsRepPermi.Fields.Item("EJEC01").Value )
                      end if
                      if not isnull(RsRepPermi.Fields.Item("OBSERV01").Value) then
                         detallepermiso(16,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("OBSERV01").Value )
                         strObservaciones=CStr(RsRepPermi.Fields.Item("OBSERV01").Value)
                      else
                         detallepermiso(16,intcontpermi-1) = (RsRepPermi.Fields.Item("OBSERV01").Value )
                         strObservaciones= RsRepPermi.Fields.Item("OBSERV01").Value
                      end if

                      if not isnull(RsRepPermi.Fields.Item("CANTIDADFAC").Value) then
                         detallepermiso(18,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CANTIDADFAC").Value)
                         detallepermiso(19,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value) 'Cantidad descontada
                         detallepermiso(4,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value-RsRepPermi.Fields.Item("CANTIDADFAC").Value)' SALDO

                         dblSumCantidadUtil= dblSumCantidadUtil + RsRepPermi.Fields.Item("CANTIDADFAC").Value
                      else
                         detallepermiso(18,intcontpermi-1) = (RsRepPermi.Fields.Item("CANTIDADFAC").Value)                         ' CHECARLO
                         detallepermiso(19,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value) 'Cantidad descontada   ' CHECARLO
                         detallepermiso(4,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value)' SALDO                 ' CHECARLO
                      end if

                      detallepermiso(17,intcontpermi-1) = (RsRepPermi.Fields.Item("refcia02").Value ) 'VAMOS A UTILIZARLO PARA GUARDAR LA REFERENCIA

                      intcontpermi = intcontpermi + 1
                    RsRepPermi.movenext



                    While NOT RsRepPermi.EOF


                       quitarep = 0
                       'BUSCAMOS EL SI LA CATIDAD FACTURA YA ESTA EN EL VECTOR
                       for y=0 to ( UBound(detallepermiso,2) - 1)
                         if not isnull( detallepermiso(18,y) ) and not isnull(RsRepPermi.Fields.Item("CANTIDADFAC").Value) then
                           if (detallepermiso(18,y) ) = CStr(RsRepPermi.Fields.Item("CANTIDADFAC").Value) then
                              quitarep = 1
                           end if
                         end if
                       next

                         'Response.Write( "PRUEBA"  )
                         'Response.Write( detallepermiso(0,0) )
                         'Response.Write( detallepermiso(18,y)  )
                         'Response.End

                      if( detallepermiso(0,intcontpermi -2) = ltrim(RsRepPermi.Fields.Item("PERMISO").Value )     and detallepermiso(9,intcontpermi -2)  = CStr(RsRepPermi.Fields.Item("FECHASALIDA").Value ) and detallepermiso(1,intcontpermi -2)  = CStr(RsRepPermi.Fields.Item("PRODUCTO").Value )    and detallepermiso(2,intcontpermi -2) = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value ) and detallepermiso(15,intcontpermi-2) = CStr(RsRepPermi.Fields.Item("EJEC01").Value )  and detallepermiso(16,intcontpermi-2) = CStr(RsRepPermi.Fields.Item("OBSERV01").Value ) and detallepermiso(6,intcontpermi -2) = CStr(RsRepPermi.Fields.Item("PRESENT01").Value ) and detallepermiso(17,intcontpermi-2) = ltrim(RsRepPermi.Fields.Item("refcia02").Value )  ) or quitarep = 1 then

                      ELSE


                        if strcompermi= ltrim(RsRepPermi.Fields.Item("PERMISO").Value) and strObservaciones=ltrim(RsRepPermi.Fields.Item("OBSERV01").Value) then 'es el mismo permiso
                            'ReDim detallepermiso(intcontpermi,20)


                            ReDim PRESERVE detallepermiso(20,intcontpermi)
                            if not isnull(RsRepPermi.Fields.Item("PERMISO").Value) then
                               strcompermi                     = ltrim(RsRepPermi.Fields.Item("PERMISO").Value)
                               detallepermiso(0,intcontpermi-1)  = ltrim(RsRepPermi.Fields.Item("PERMISO").Value)
                            else
                               strcompermi                     = " "
                               detallepermiso(0,intcontpermi-1)  = (RsRepPermi.Fields.Item("PERMISO").Value)
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PRODUCTO").Value) then
                               detallepermiso(1,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("PRODUCTO").Value)
                            else
                               detallepermiso(1,intcontpermi-1)  = (RsRepPermi.Fields.Item("PRODUCTO").Value)
                            end if
                            if not isnull(RsRepPermi.Fields.Item("CANTIDAD").Value) then
                               detallepermiso(2,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value )
                               'detallepermiso(2,intcontpermi-1)  = detallepermiso(2,intcontpermi-2)
                            else
                               detallepermiso(2,intcontpermi-1)  = (RsRepPermi.Fields.Item("CANTIDAD").Value )
                            end if

                            'if not isnull(RsRepPermi.Fields.Item("UTILIZADA").Value) then
                            '   detallepermiso(3,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("UTILIZADA").Value )
                            'else
                            '   detallepermiso(3,intcontpermi-1)  = (RsRepPermi.Fields.Item("UTILIZADA").Value )
                            'end if

                            'if not isnull(RsRepPermi.Fields.Item("SALDO").Value) then
                            '   detallepermiso(4,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("SALDO").Value )
                            'else
                            '   detallepermiso(4,intcontpermi-1)  = (RsRepPermi.Fields.Item("SALDO").Value )
                            'end if

                            if not isnull(RsRepPermi.Fields.Item("UnidadMedida").Value) then
                               detallepermiso(5,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("UnidadMedida").Value )
                            else
                               detallepermiso(5,intcontpermi-1)  = (RsRepPermi.Fields.Item("UnidadMedida").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PRESENT01").Value) then
                               detallepermiso(6,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("PRESENT01").Value )
                            else
                               detallepermiso(6,intcontpermi-1)  = (RsRepPermi.Fields.Item("PRESENT01").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("REGSAN01").Value) then
                               detallepermiso(7,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("REGSAN01").Value )
                            else
                               detallepermiso(7,intcontpermi-1)  = (RsRepPermi.Fields.Item("REGSAN01").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PERMISO").Value) then
                               detallepermiso(8,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("PERMISO").Value )
                            else
                               detallepermiso(8,intcontpermi-1)  = (RsRepPermi.Fields.Item("PERMISO").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("FECHASALIDA").Value) then
                               detallepermiso(9,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("FECHASALIDA").Value)
                            else
                               detallepermiso(9,intcontpermi-1)  = (RsRepPermi.Fields.Item("FECHASALIDA").Value)
                            end if
                            if not isnull(RsRepPermi.Fields.Item("FECHAVENCI").Value) then
                               detallepermiso(10,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("FECHAVENCI").Value )
                            else
                               detallepermiso(10,intcontpermi-1) = (RsRepPermi.Fields.Item("FECHAVENCI").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PAISFABRICANTE").Value) then
                               detallepermiso(11,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("PAISFABRICANTE").Value )
                            else
                               detallepermiso(11,intcontpermi-1) = " "
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PAISFACTURADOR").Value) then
                               detallepermiso(12,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("PAISFACTURADOR").Value )
                            else
                               detallepermiso(12,intcontpermi-1) = " "
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PAISPROVEEDOR").Value) then
                               detallepermiso(13,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("PAISPROVEEDOR").Value )
                            else
                               detallepermiso(13,intcontpermi-1) = " "
                            end if
                            if not isnull(RsRepPermi.Fields.Item("DIASVEN01").Value) then
                               detallepermiso(14,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("DIASVEN01").Value )
                            else
                               detallepermiso(14,intcontpermi-1) = (RsRepPermi.Fields.Item("DIASVEN01").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("EJEC01").Value) then
                               detallepermiso(15,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("EJEC01").Value )
                            else
                               detallepermiso(15,intcontpermi-1) = (RsRepPermi.Fields.Item("EJEC01").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("OBSERV01").Value) then
                               detallepermiso(16,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("OBSERV01").Value )
                               strObservaciones  = CStr(RsRepPermi.Fields.Item("OBSERV01").Value)
                            else
                               detallepermiso(16,intcontpermi-1) = (RsRepPermi.Fields.Item("OBSERV01").Value )
                               strObservaciones=(RsRepPermi.Fields.Item("OBSERV01").Value)
                            end if
                            'if not isnull(RsRepPermi.Fields.Item("CLIE01").Value) then
                            '   detallepermiso(17,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CLIE01").Value )
                            'else
                            '   detallepermiso(17,intcontpermi-1) = (RsRepPermi.Fields.Item("CLIE01").Value )
                            'end if
                            if not isnull(RsRepPermi.Fields.Item("CANTIDADFAC").Value) then
                              detallepermiso(18,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CANTIDADFAC").Value)
                              detallepermiso(19,intcontpermi-1) = detallepermiso(4,intcontpermi-2) 'Cantidad descontada
                              detallepermiso(4,intcontpermi-1)  = CStr(detallepermiso(4,intcontpermi-2) - RsRepPermi.Fields.Item("CANTIDADFAC").Value ) ' SALDO

                              dblSumCantidadUtil= dblSumCantidadUtil + RsRepPermi.Fields.Item("CANTIDADFAC").Value
                            else
                              detallepermiso(18,intcontpermi-1) = (RsRepPermi.Fields.Item("CANTIDADFAC").Value)
                              detallepermiso(19,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value) 'Cantidad descontada   ' CHECARLO
                              'detallepermiso(19,intcontpermi-1) = CStr(0) 'Cantidad descontada   ' CHECARLO
                              detallepermiso(4,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value)' SALDO                 ' CHECARLO
                            end if

                            detallepermiso(17,intcontpermi-1) = (RsRepPermi.Fields.Item("refcia02").Value ) 'VAMOS A UTILIZARLO PARA GUARDAR LA REFERENCIA

                            'With oSS.ActiveSheet
                            '       .Cells(intcontperm+3, 1).Value = strcompermi
                            '       '.Cells(1, x+1).Value = detallepermiso(1,x)
                            'end with

                            intcontpermi = intcontpermi + 1
                        else ' cambiamos de permiso, hay que dibujar los registros que llevamos en el arreglo
                            '--------------------------------------------------------------------------------------------
                            ' Dibujar lo que hay almacenado en el vector
                            '--------------------------------------------------------------------------------------------

                             for x=0 to ( UBound(detallepermiso,2)-1 )
                                intUtilizada2  = 0
                                intCantidad    = 0

                                strProducto_aux    = detallepermiso(1,x) 'PRODUCTO
                                intCantidad_aux    = detallepermiso(19,x) 'CANTIDAD '
                                intUtilizada_aux   = detallepermiso(3,x) 'UTILIZADA
                                intSaldo_aux       = detallepermiso(4,x) 'SALDO
                                strUniMed_aux      = detallepermiso(5,x) 'UnidadMedida
                                strPresent_aux     = detallepermiso(6,x) 'PRESENT01
                                strregsan_aux      = detallepermiso(7,x) 'REGSAN01
                                strPermi_aux       = detallepermiso(8,x) 'PERMISO
                                strReferencia      = detallepermiso(17,x) 'PERMISO

                                if not isnull(detallepermiso(9,x)) then
                                    dFechaSali_aux  = Day(detallepermiso(9,x))&"/"&month(detallepermiso(9,x))&"/"&year(detallepermiso(9,x)) 'FECHASALIDA
                                end if

                                if not isnull(detallepermiso(10,X)) then
                                    FechaVenci_aux  = Day(detallepermiso(10,x))&"/"&month(detallepermiso(10,x))&"/"&year(detallepermiso(10,x)) 'FECHAVENCI
                                end if

                                'dFechaVenci_aux    = detallepermiso(x,10) 'FECHAVENCI
                                strPaisFabric_aux  = detallepermiso(11,x) 'PAISFABRICANTE
                                strPaisFactur_aux  = detallepermiso(12,x) 'PAISFACTURADOR
                                strPaisProv_aux    = detallepermiso(13,x) 'PAISPROVEEDOR
                                intDiasVenci_aux   = detallepermiso(14,x) 'DIASVEN01
                                strEjecu_aux       = detallepermiso(15,x) 'EJEC01
                                strObserv_aux      = detallepermiso(16,x) 'OBSERV01
                                intcvecli_aux      = detallepermiso(17,x) 'CLIE01
                                intUtilizada2_aux  = detallepermiso(18,x) 'CANTIDADFAC

'                                strProducto    = detallepermiso(x,1) 'PRODUCTO
'                                intCantidad    = detallepermiso(x,2) 'CANTIDAD
'                                intUtilizada   = detallepermiso(x,3) 'UTILIZADA
'                                intSaldo       = detallepermiso(x,4) 'SALDO
'                                strUniMed      = detallepermiso(x,5) 'UnidadMedida
'                                strPresent     = detallepermiso(x,6) 'PRESENT01
'                                strregsan      = detallepermiso(x,7) 'REGSAN01
'                                strPermi       = detallepermiso(x,8) 'PERMISO
'                                dFechaSali     = detallepermiso(x,9) 'FECHASALIDA
'                                dFechaVenci    = detallepermiso(x,10) 'FECHAVENCI
'                                strPaisFabric  = detallepermiso(x,11) 'PAISFABRICANTE
'                                strPaisFactur  = detallepermiso(x,12) 'PAISFACTURADOR
'                                strPaisProv    = detallepermiso(x,13) 'PAISPROVEEDOR
'                                intDiasVenci   = detallepermiso(x,14) 'DIASVEN01
'                                strEjecu       = detallepermiso(x,15) 'EJEC01
'                                strObserv      = detallepermiso(x,16) 'OBSERV01
'                                intcvecli      = detallepermiso(x,17) 'CLIE01
'                                intUtilizada2  = detallepermiso(x,18) 'CANTIDADFAC

                              if not isnull(strProducto_aux) then
                                   strProducto    = ltrim(strProducto_aux) 'PRODUCTO
                              end if

                                'With oSS.ActiveSheet
                                '   .Cells(1, x+1).Value = strcompermi
                                '   '.Cells(1, x+1).Value = detallepermiso(1,x)
                                'end with


                              if not isnull(strUniMed_aux) then
                                   strUniMed      = ltrim(strUniMed_aux) 'UnidadMedida
                              end if

                              if not isnull(strPresent_aux) then
                                   strPresent     = ltrim(strPresent_aux) 'PRESENT01
                              end if

                              if not isnull(strregsan_aux) then
                                   strregsan      = ltrim(strregsan_aux) 'REGSAN01
                              end if

                              if not isnull(strPermi_aux) then
                                   strPermi       = ltrim(strPermi_aux) 'PERMISO
                              end if

                              if not isnull(strPaisFabric_aux) then
                                   strPaisFabric  = ltrim(strPaisFabric_aux) 'PAISFABRICANTE
                              end if

                              if not isnull(strPaisFactur_aux) then
                                   strPaisFactur  = ltrim(strPaisFactur_aux) 'PAISFACTURADOR
                              end if

                              if not isnull(strPaisProv_aux) then
                                   strPaisProv    = ltrim(strPaisProv_aux) 'PAISPROVEEDOR
                              end if

                              if not isnull(strEjecu_aux) then
                                   strEjecu       = ltrim(strEjecu_aux) 'EJEC01
                              end if

                              if not isnull(strObserv_aux) then
                                   strObserv      = ltrim(strObserv_aux) 'OBSERV01
                              end if

                              if NOT isnull(strPresent) then
                                strUniMed = strUniMed & strPresent
                              end if
                              '------------------------------------
                              if isnull(intUtilizada_aux) then
                                   intUtilizada = 0
                              else
                                   intUtilizada = Cdbl(intUtilizada_aux)
                              end if

                              if isnull(intUtilizada2_aux) then
                                   intUtilizada2 = 0
                              else
                                   intUtilizada2 = Cdbl(intUtilizada2_aux)
                              end if

                              if isnull(intCantidad_aux) then
                                   intCantidad = 0
                              else
                                   intCantidad = Cdbl(intCantidad_aux)
                              end if

                              if isnull(intSaldo_aux) then
                                   intSaldo = 0
                              else
                                   intSaldo = Cdbl(intSaldo_aux)
                              end if

                              if( (intUtilizada) > (intUtilizada2) ) then
                                 intUtilizada = intUtilizada
                              end if

'                              if intSaldo = 0 then
'                                 intSaldo = intCantidad - intUtilizada2
'                              end if

'                                dFechaSali     = detallepermiso(x,9) 'FECHASALIDA
'                                dFechaVenci    = detallepermiso(x,10) 'FECHAVENCI
'                                intDiasVenci   = detallepermiso(x,14) 'DIASVEN01


                              if not isnull(dFechaSali_aux) then ' Fecha de salida
                                if  isdate(dFechaSali_aux) then
                                      dFechaSali = DateValue(dFechaSali_aux)
                                 end if
                              end if

                              if not isnull(dFechaVenci_aux) then ' si no le capturaron la fecha de vencimiento
                                 if  isdate(dFechaVenci_aux) then
                                      dFechaVenci = DateValue(dFechaVenci_aux)
                                 else
                                      dFechaVenci = dFechaSali + 179
                                 end if
                              else
                                      dFechaVenci = dFechaSali + 179
                              end if

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
                                      .Cells(contadorPag(intcontEjec-1), 13).Value = strReferencia
                                      .Cells(contadorPag(intcontEjec-1), 13).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 13).Font.size = 8

                                      contadorPag(intcontEjec-1) = contadorPag(intcontEjec-1) + 1
                                     ' .Cells(contadorPag(intcontEjec), 19).Value = contadorPag(intcontEjec)
                                  End With
                              else ' No esta vencido por fecha, verificar la cantidad utilizada para ver si no se ha vencido por cantidad
                                   if (dblSumCantidadUtil)>=(intCantidad) then
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
                                            .Cells(contadorPag(intcontEjec-1), 13).Value = strReferencia
                                            .Cells(contadorPag(intcontEjec-1), 13).Interior.ColorIndex = 3
                                            .Cells(contadorPag(intcontEjec-1), 13).Font.size = 8
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
                                                   .Cells(contadorPag(intRow - 1), 13).Value = strReferencia
                                                   .Cells(contadorPag(intRow - 1), 13).Font.size = 8
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
                                              .Cells(contadorPag(intcontEjec-1), 13).Value = strReferencia
                                              .Cells(contadorPag(intcontEjec-1), 13).Font.size = 8
                                           End With
                                           contadorPag(intcontEjec-1) = contadorPag(intcontEjec-1) + 1
                                       end if

                                   end if
                              end if

                             next


                            '----------------------------------------------
                            ' Redimensionar el vector
                            '----------------------------------------------
                            'ReDim detallepermiso(intcontpermi,20)
                            'strcompermi                     = RsRepPermi.Fields.Item("PERMISO").Value
                            'detallepermiso(intcontpermi,0)  = CStr( RsRepPermi.Fields.Item("PERMISO").Value )
                            'detallepermiso(intcontpermi,1)  = CStr(RsRepPermi.Fields.Item("PRODUCTO").Value )
                            'detallepermiso(intcontpermi,2)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value )
                            'detallepermiso(intcontpermi,3)  = CStr(RsRepPermi.Fields.Item("UTILIZADA").Value )
                            'detallepermiso(intcontpermi,4)  = CStr(RsRepPermi.Fields.Item("SALDO").Value )
                            'detallepermiso(intcontpermi,5)  = CStr(RsRepPermi.Fields.Item("UnidadMedida").Value )
                            'detallepermiso(intcontpermi,6)  = CStr(RsRepPermi.Fields.Item("PRESENT01").Value )
                            'detallepermiso(intcontpermi,7)  = CStr(RsRepPermi.Fields.Item("REGSAN01").Value )
                            'detallepermiso(intcontpermi,8)  = CStr(RsRepPermi.Fields.Item("PERMISO").Value )
                            'detallepermiso(intcontpermi,9)  = CStr(RsRepPermi.Fields.Item("FECHASALIDA").Value )
                            'detallepermiso(intcontpermi,10) = CStr(RsRepPermi.Fields.Item("FECHAVENCI").Value )
                            'detallepermiso(intcontpermi,11) = CStr(RsRepPermi.Fields.Item("PAISFABRICANTE").Value )
                            'detallepermiso(intcontpermi,12) = CStr(RsRepPermi.Fields.Item("PAISFACTURADOR").Value )
                            'detallepermiso(intcontpermi,13) = CStr(RsRepPermi.Fields.Item("PAISPROVEEDOR").Value )
                            'detallepermiso(intcontpermi,14) = CStr(RsRepPermi.Fields.Item("DIASVEN01").Value )
                            'detallepermiso(intcontpermi,15) = CStr(RsRepPermi.Fields.Item("EJEC01").Value )
                            'detallepermiso(intcontpermi,16) = CStr(RsRepPermi.Fields.Item("OBSERV01").Value )
                            'detallepermiso(intcontpermi,17) = CStr(RsRepPermi.Fields.Item("CLIE01").Value )
                            'detallepermiso(intcontpermi,18) = CStr(RsRepPermi.Fields.Item("CANTIDADFAC").Value )

                            intcontpermi = 1
                            dblSumCantidadUtil=0
                            ReDim detallepermiso(20,intcontpermi)
                            if not isnull(RsRepPermi.Fields.Item("PERMISO").Value) then
                               strcompermi                     = ltrim(RsRepPermi.Fields.Item("PERMISO").Value)
                               detallepermiso(0,intcontpermi-1)  = ltrim(RsRepPermi.Fields.Item("PERMISO").Value)
                            else
                               strcompermi                     = " "
                               detallepermiso(0,intcontpermi-1)  = (RsRepPermi.Fields.Item("PERMISO").Value)
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PRODUCTO").Value) then
                               detallepermiso(1,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("PRODUCTO").Value)
                            else
                               detallepermiso(1,intcontpermi-1)  = (RsRepPermi.Fields.Item("PRODUCTO").Value)
                            end if
                            if not isnull(RsRepPermi.Fields.Item("CANTIDAD").Value) then
                               detallepermiso(2,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value )
                            else
                               detallepermiso(2,intcontpermi-1)  = (RsRepPermi.Fields.Item("CANTIDAD").Value )
                               'detallepermiso(2,intcontpermi-1)  = CStr(0)
                            end if
                            if not isnull(RsRepPermi.Fields.Item("UTILIZADA").Value) then
                               detallepermiso(3,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("UTILIZADA").Value )
                            else
                               detallepermiso(3,intcontpermi-1)  = (RsRepPermi.Fields.Item("UTILIZADA").Value )
                            end if
                            'if not isnull(RsRepPermi.Fields.Item("SALDO").Value) then
                            '   detallepermiso(4,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("SALDO").Value )
                            'else
                            '   detallepermiso(4,intcontpermi-1)  = (RsRepPermi.Fields.Item("SALDO").Value )
                            'end if
                            if not isnull(RsRepPermi.Fields.Item("UnidadMedida").Value) then
                               detallepermiso(5,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("UnidadMedida").Value )
                            else
                               detallepermiso(5,intcontpermi-1)  = (RsRepPermi.Fields.Item("UnidadMedida").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PRESENT01").Value) then
                               detallepermiso(6,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("PRESENT01").Value )
                            else
                               detallepermiso(6,intcontpermi-1)  = (RsRepPermi.Fields.Item("PRESENT01").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("REGSAN01").Value) then
                               detallepermiso(7,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("REGSAN01").Value )
                            else
                               detallepermiso(7,intcontpermi-1)  = (RsRepPermi.Fields.Item("REGSAN01").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PERMISO").Value) then
                               detallepermiso(8,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("PERMISO").Value )
                            else
                               detallepermiso(8,intcontpermi-1)  = (RsRepPermi.Fields.Item("PERMISO").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("FECHASALIDA").Value) then
                               detallepermiso(9,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("FECHASALIDA").Value)
                            else
                               detallepermiso(9,intcontpermi-1)  = (RsRepPermi.Fields.Item("FECHASALIDA").Value)
                            end if
                            if not isnull(RsRepPermi.Fields.Item("FECHAVENCI").Value) then
                               detallepermiso(10,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("FECHAVENCI").Value )
                            else
                               detallepermiso(10,intcontpermi-1) = (RsRepPermi.Fields.Item("FECHAVENCI").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PAISFABRICANTE").Value) then
                               detallepermiso(11,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("PAISFABRICANTE").Value )
                            else
                               detallepermiso(11,intcontpermi-1) = " "
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PAISFACTURADOR").Value) then
                               detallepermiso(12,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("PAISFACTURADOR").Value )
                            else
                               detallepermiso(12,intcontpermi-1) = " "
                            end if
                            if not isnull(RsRepPermi.Fields.Item("PAISPROVEEDOR").Value) then
                               detallepermiso(13,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("PAISPROVEEDOR").Value )
                            else
                               detallepermiso(13,intcontpermi-1) = " "
                            end if
                            if not isnull(RsRepPermi.Fields.Item("DIASVEN01").Value) then
                               detallepermiso(14,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("DIASVEN01").Value )
                            else
                               detallepermiso(14,intcontpermi-1) = (RsRepPermi.Fields.Item("DIASVEN01").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("EJEC01").Value) then
                               detallepermiso(15,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("EJEC01").Value )
                            else
                               detallepermiso(15,intcontpermi-1) = (RsRepPermi.Fields.Item("EJEC01").Value )
                            end if
                            if not isnull(RsRepPermi.Fields.Item("OBSERV01").Value) then
                               detallepermiso(16,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("OBSERV01").Value )
                               strObservaciones = CStr(RsRepPermi.Fields.Item("OBSERV01").Value)
                            else
                               detallepermiso(16,intcontpermi-1) = (RsRepPermi.Fields.Item("OBSERV01").Value )
                               strObservaciones = (RsRepPermi.Fields.Item("OBSERV01").Value)
                            end if
                            'if not isnull(RsRepPermi.Fields.Item("CLIE01").Value) then
                            '   detallepermiso(17,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CLIE01").Value )
                            'else
                            '   detallepermiso(17,intcontpermi-1) = (RsRepPermi.Fields.Item("CLIE01").Value )
                            'end if
                            if not isnull(RsRepPermi.Fields.Item("CANTIDADFAC").Value) then
                              detallepermiso(18,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CANTIDADFAC").Value)
                              detallepermiso(19,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value) 'Cantidad descontada
                              detallepermiso(4,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value-RsRepPermi.Fields.Item("CANTIDADFAC").Value)' SALDO

                              dblSumCantidadUtil= dblSumCantidadUtil + RsRepPermi.Fields.Item("CANTIDADFAC").Value
                            else
                              detallepermiso(18,intcontpermi-1) = (RsRepPermi.Fields.Item("CANTIDADFAC").Value)
                              detallepermiso(19,intcontpermi-1) = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value) 'Cantidad descontada   ' CHECARLO
                              'detallepermiso(19,intcontpermi-1) = CStr(0) 'Cantidad descontada   ' CHECARLO
                              'detallepermiso(2,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value)
                              detallepermiso(4,intcontpermi-1)  = CStr(RsRepPermi.Fields.Item("CANTIDAD").Value)' SALDO                 ' CHECARLO
                            end if

                            detallepermiso(17,intcontpermi-1) = (RsRepPermi.Fields.Item("refcia02").Value ) 'VAMOS A UTILIZARLO PARA GUARDAR LA REFERENCIA

                            intcontpermi = intcontpermi + 1
                        end if



                     end if    'fin


                        RsRepPermi.movenext
                    wend




                                                ' Dibujar lo que hay almacenado en el vector
                            '--------------------------------------------------------------------------------------------

                             for x=0 to ( UBound(detallepermiso,2)-1 )
                                intUtilizada2  = 0
                                intCantidad    = 0

                                strProducto_aux    = detallepermiso(1,x) 'PRODUCTO
                                intCantidad_aux    = detallepermiso(2,x) 'CANTIDAD
                                'intUtilizada_aux   = detallepermiso(3,x) 'UTILIZADA
                                intUtilizada_aux   = detallepermiso(19,x) 'UTILIZADA
                                intSaldo_aux       = detallepermiso(4,x) 'SALDO
                                strUniMed_aux      = detallepermiso(5,x) 'UnidadMedida
                                strPresent_aux     = detallepermiso(6,x) 'PRESENT01
                                strregsan_aux      = detallepermiso(7,x) 'REGSAN01
                                strPermi_aux       = detallepermiso(8,x) 'PERMISO

                                if not isnull(detallepermiso(9,x)) then
                                    dFechaSali_aux  = Day(detallepermiso(9,x))&"/"&month(detallepermiso(9,x))&"/"&year(detallepermiso(9,x)) 'FECHASALIDA
                                end if

                                if not isnull(detallepermiso(10,X)) then
                                    FechaVenci_aux  = Day(detallepermiso(10,x))&"/"&month(detallepermiso(10,x))&"/"&year(detallepermiso(10,x)) 'FECHAVENCI
                                end if

                                'dFechaVenci_aux    = detallepermiso(x,10) 'FECHAVENCI
                                strPaisFabric_aux  = detallepermiso(11,x) 'PAISFABRICANTE
                                strPaisFactur_aux  = detallepermiso(12,x) 'PAISFACTURADOR
                                strPaisProv_aux    = detallepermiso(13,x) 'PAISPROVEEDOR
                                intDiasVenci_aux   = detallepermiso(14,x) 'DIASVEN01
                                strEjecu_aux       = detallepermiso(15,x) 'EJEC01
                                strObserv_aux      = detallepermiso(16,x) 'OBSERV01
                                intcvecli_aux      = detallepermiso(17,x) 'CLIE01
                                intUtilizada2_aux  = detallepermiso(18,x) 'CANTIDADFAC
                                strReferencia      = detallepermiso(17,x) 'PERMISO

                              if not isnull(strProducto_aux) then
                                   strProducto    = ltrim(strProducto_aux) 'PRODUCTO
                              end if

                                'With oSS.ActiveSheet
                                '   .Cells(1, x+1).Value = strcompermi
                                '   '.Cells(1, x+1).Value = detallepermiso(1,x)
                                'end with


                              if not isnull(strUniMed_aux) then
                                   strUniMed      = ltrim(strUniMed_aux) 'UnidadMedida
                              end if

                              if not isnull(strPresent_aux) then
                                   strPresent     = ltrim(strPresent_aux) 'PRESENT01
                              end if

                              if not isnull(strregsan_aux) then
                                   strregsan      = ltrim(strregsan_aux) 'REGSAN01
                              end if

                              if not isnull(strPermi_aux) then
                                   strPermi       = ltrim(strPermi_aux) 'PERMISO
                              end if

                              if not isnull(strPaisFabric_aux) then
                                   strPaisFabric  = ltrim(strPaisFabric_aux) 'PAISFABRICANTE
                              end if

                              if not isnull(strPaisFactur_aux) then
                                   strPaisFactur  = ltrim(strPaisFactur_aux) 'PAISFACTURADOR
                              end if

                              if not isnull(strPaisProv_aux) then
                                   strPaisProv    = ltrim(strPaisProv_aux) 'PAISPROVEEDOR
                              end if

                              if not isnull(strEjecu_aux) then
                                   strEjecu       = ltrim(strEjecu_aux) 'EJEC01
                              end if

                              if not isnull(strObserv_aux) then
                                   strObserv      = ltrim(strObserv_aux) 'OBSERV01
                              end if

                              if NOT isnull(strPresent) then
                                strUniMed = strUniMed & strPresent
                              end if

                              if isnull(intUtilizada_aux) then
                                   intUtilizada = 0
                              else
                                   intUtilizada = Cdbl(intUtilizada_aux)
                              end if

                              if isnull(intUtilizada2_aux) then
                                   intUtilizada2 = 0
                              else
                                   intUtilizada2 = Cdbl(intUtilizada2_aux)
                              end if

                              if isnull(intCantidad_aux) then
                                   intCantidad = 0
                              else
                                   intCantidad = Cdbl(intCantidad_aux)
                              end if

                              if isnull(intSaldo_aux) then
                                   intSaldo = 0
                              else
                                   intSaldo = Cdbl(intSaldo_aux)
                              end if

                              if( (intUtilizada) > (intUtilizada2) ) then
                                 intUtilizada = intUtilizada
                              end if

                              if intSaldo = 0 then
                                 intSaldo = intCantidad - intUtilizada2
                              end if

'                                dFechaSali     = detallepermiso(x,9) 'FECHASALIDA
'                                dFechaVenci    = detallepermiso(x,10) 'FECHAVENCI
'                                intDiasVenci   = detallepermiso(x,14) 'DIASVEN01


                              if not isnull(dFechaSali_aux) then ' Fecha de salida
                                if  isdate(dFechaSali_aux) then
                                      dFechaSali = DateValue(dFechaSali_aux)
                                 end if
                              end if

                              if not isnull(dFechaVenci_aux) then ' si no le capturaron la fecha de vencimiento
                                 if  isdate(dFechaVenci_aux) then
                                      dFechaVenci = DateValue(dFechaVenci_aux)
                                 else
                                      dFechaVenci = dFechaSali + 179
                                 end if
                              else
                                      dFechaVenci = dFechaSali + 179
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
                                      .Cells(contadorPag(intcontEjec-1), 13).Value = strReferencia
                                      .Cells(contadorPag(intcontEjec-1), 12).Interior.ColorIndex = 3
                                      .Cells(contadorPag(intcontEjec-1), 13).Font.size = 8
                                      contadorPag(intcontEjec-1) = contadorPag(intcontEjec-1) + 1
                                     ' .Cells(contadorPag(intcontEjec), 19).Value = contadorPag(intcontEjec)
                                  End With
                              else ' No esta vencido por fecha, verificar la cantidad utilizada para ver si no se ha vencido por cantidad
                                   'if (intUtilizada2)>=(intCantidad) then
                                   if (dblSumCantidadUtil)>=(intCantidad) then
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
                                            .Cells(contadorPag(intcontEjec-1), 13).Value = strReferencia
                                            .Cells(contadorPag(intcontEjec-1), 12).Interior.ColorIndex = 3
                                            .Cells(contadorPag(intcontEjec-1), 13).Font.size = 8
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
                                                   .Cells(contadorPag(intRow - 1), 13).Value = strReferencia
                                                   .Cells(contadorPag(intRow - 1), 13).Font.size = 8
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
                                              .Cells(contadorPag(intcontEjec-1), 13).Value = strReferencia
                                              .Cells(contadorPag(intcontEjec-1), 13).Font.size = 8
                                           End With
                                           contadorPag(intcontEjec-1) = contadorPag(intcontEjec-1) + 1
                                       end if

                                   end if
                              end if

                             next
'--------------------------------------------------------------------------------------------------------------------------------------------




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
