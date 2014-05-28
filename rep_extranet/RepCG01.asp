<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% 


On error resume next
 Response.Buffer = TRUE
 Response.Addheader "Content-Disposition", "attachment;"
 Response.ContentType = "application/vnd.ms-excel"

 strHTML = ""
 strTipoUsuario = Session("GTipoUsuario")
 fechaini = trim(request.Form("txtDateIni"))
 fechafin = trim(request.Form("txtDateFin"))
 strTipoOperaciones = request.Form("rbnTipoDate")

 dim Rsio,Rsio2,Rsio3,Rsio4

strPermisos = Request.Form("Permisos")

strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")


 if not fechaini="" and not fechafin="" then


    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    strDateIni = tmpAnioIni & "/" &tmpMesIni & "/"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    strDateFin = tmpAnioFin & "/" &tmpMesFin & "/"& tmpDiaFin

    if strTipoOperaciones = 1 then
	     strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE OPERACIONES DE IMPORTACION</p></font></strong>"
    else
	     strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE OPERACIONES DE EXPORTACION</p></font></strong>"
    end if

	  strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & fechaini & " Al " & fechafin & "</p></font></strong>"
    strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	  strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
	  strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
	  strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CuentaGastos</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">cve_pedto" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">fecha_pago</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">proveedor</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">pais_prov</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">fact_prov</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">val_merc_mn</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">val_aduana</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">dta</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">adv</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">iva</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">preval</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">tot_impuesto</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">regimen</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ph_sin_iva</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">tasa_cero</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ph_iva15</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ph_iva10</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">p_iva</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">i_iva</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">hon_mas_sce</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">adicionales_hon</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">base_para_iva</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">iva_cgastos</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">total_cgastos</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">mercancia</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">anticipo</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">peso_bruto</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">division</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">tipo_oper</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Orden</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IPC</td>" & chr(13) & chr(10)

    strHTML = strHTML & "</tr>"& chr(13) & chr(10)
'  MM_EXTRANET_STRING_TEMP = ""
  'for x=4 to 4
  '  if x=1 then
  '     aduana="VER"
  '  end if
  '  if x=2 then
  '     aduana="MAN"
  '  end if
  '  if x=3 then
  '     aduana="TAM"
  '  end if
  '  if x=4 then
  '     aduana="MEX"
  '  end if
  '  if x=5 then
  '     aduana="LZR"
  '  end if
    'if x=6 then
    '   aduana="LAR"
    'end if

    MM_EXTRANET_STRING_TEMP = ODBC_POR_ADUANA(Session("GAduana"))

   permi = PermisoClientes(Session("GAduana"),strPermisos,"clie31")
   permi2 = PermisoClientesTabla("E",Session("GAduana"),strPermisos,"clie31")



    if not permi = "" then
       permi = "  and (" & permi & ") "
       permi2 = "  and (" & permi2 & ") "
    end if

    AplicaFiltro = false
    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
       blnAplicaFiltro = true
    end if
    if blnAplicaFiltro then
       permi = " AND clie31 =" & strFiltroCliente
       permi2 = " AND E.clie31 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
       permi = ""
       permi2 = ""
    end if

    ' if not permi= "" then
      set Rsio = server.CreateObject("ADODB.Recordset")
      Rsio.ActiveConnection =    MM_EXTRANET_STRING_TEMP

      strSQL = "SELECT E.caho31,E.cgas31,E.fech31,E.suph31,E.chon31,E.csce31,E.piva31,E.tota31,E.anti31,E.sald31,E.clie31,E.paho31,E.coad31,D.refe31 FROM D31REFER as D LEFT JOIN E31CGAST as E ON D.CGAS31 = E.CGAS31 WHERE E.esta31 = 'I' and (E.fech31 >= '" & strDateIni & "' and E.fech31 <= '" & strDateFin & "')  " & permi2 & " order by E.cgas31 "
      Rsio.Source= strSQL
      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()

      While NOT Rsio.EOF
              intTotaldeIPC = 0
              dblTotalFac = 0
              strPedimento =""
              strCvePed =""
              strFechaPago = ""
              strProveedor_Cliente= ""
              strPaisProveedor=""
              strFacturaProveedor=""
              strValorMercancia =0
              strValorAduana =0
              intDTA =0
              intADV =0
              intIVA =0
              intPRV =0
              intTotalImpuesto = 0
              intImpuesto=0
              strRegimen=""
              Ph=0
              ph_iva= 0
              strTipoMercancia = 0
              strTipoOperacion = ""
              dblPesoBruto = 0
              set RsDatos = server.CreateObject("ADODB.Recordset")
              RsDatos.ActiveConnection =    MM_EXTRANET_STRING_TEMP
              strSQL = "SELECT 'EXPO' as tipo,cvepvc01,nompro01,desdoc01,cveped01,fecpag01,desf0101,numped01,pesobr01,sum(vmerme02 * factmo01 * tipcam01) as valormercancia,sum(vaduan02) as valoraduana,sum(i_adv102 + i_adv202 + i_adv302) as ADV,sum(i_iva102 + i_iva202 + i_iva302) as IVA,sum(dtafpp02)   as V_DTA FROM ssdage01,ssfrac02 where refcia01 = refcia02 and refcia01 = '" & Rsio.Fields.Item("refe31").Value & "' group by refcia02 UNION SELECT 'IMPO' as tipo,cvepvc01,nompro01,desdoc01,cveped01,fecpag01,desf0101,numped01,pesobr01,sum(vmerme02 * factmo01 * tipcam01) as valormercancia,sum(vaduan02) as valoraduana,sum(i_adv102 + i_adv202 + i_adv302) as ADV,sum(i_iva102 + i_iva202 + i_iva302) as IVA,sum(dtafpp02) as V_DTA FROM ssdagi01,ssfrac02 where refcia01 = refcia02 and refcia01 = '" & Rsio.Fields.Item("refe31").Value  & "'  group by refcia02  "

              RsDatos.Source= strSQL

              RsDatos.CursorType = 0
              RsDatos.CursorLocation = 2
              RsDatos.LockType = 1
              RsDatos.Open()
              intTotalImpuesto = 0
              IF not RsDatos.eof then
                 strPedimento =RsDatos.Fields.Item("numped01").Value
                 strCvePed =RsDatos.Fields.Item("cveped01").Value
                 strFechaPago = RsDatos.Fields.Item("fecpag01").Value
                 strProveedor_Cliente= RsDatos.Fields.Item("nompro01").Value
                 strPaisProveedor=""
                 strFacturaProveedor=RsDatos.Fields.Item("desf0101").Value
                 strValorMercancia =cdbl(RsDatos.Fields.Item("valormercancia").Value)
                 strValorAduana =cdbl(RsDatos.Fields.Item("valoraduana").Value)
                 intDTA =cdbl(RsDatos.Fields.Item("V_DTA").Value)
                 intADV =cdbl(RsDatos.Fields.Item("ADV").Value)
                 intIVA =cdbl(RsDatos.Fields.Item("IVA").Value)
                 strTipoOperacion = RsDatos.Fields.Item("tipo").Value
                 dblPesoBruto = RsDatos.Fields.Item("pesobr01").Value
                 strRegimen = RsDatos.Fields.Item("desdoc01").Value
                 strPaisProveedor = RsDatos.Fields.Item("cvepvc01").Value

              end if
              RsDatos.close
              set RsDatos = Nothing

              ' SOLO VAMOS POR LA PREVALIDACION
              set RsDatos = server.CreateObject("ADODB.Recordset")
              RsDatos.ActiveConnection =    MM_EXTRANET_STRING_TEMP
              strSQL = "SELECT cveimp36 AS impuesto,SUM(import36) Monto FROM sscont36 WHERE refcia36= '" & Rsio.Fields.Item("refe31").Value & "' AND cveimp36='15' GROUP BY impuesto"
              RsDatos.Source= strSQL
              RsDatos.CursorType = 0
              RsDatos.CursorLocation = 2
              RsDatos.LockType = 1
              RsDatos.Open()
              intPRV = 0
              IF not RsDatos.eof then
                 intPRV  = RsDatos.Fields.Item("monto").Value
              end if
              RsDatos.close
              set RsDatos = Nothing




              ' por cada cuenta repetirla segun las mercancias que tenga y prorratear
               set RsMercancias = server.CreateObject("ADODB.Recordset")
               RsMercancias.ActiveConnection =    MM_EXTRANET_STRING_TEMP
               strSQL = "SELECT * FROM d05artic where refe05 = '" & Rsio.Fields.Item("refe31").Value & "'"
               RsMercancias.Source= strSQL
               RsMercancias.CursorType = 0
               RsMercancias.CursorLocation = 2
               RsMercancias.LockType = 1
               RsMercancias.Open()

               dblTotalFac = 0
               if not RsMercancias.eof then
                  set RsCuenta = server.CreateObject("ADODB.Recordset")
                  RsCuenta.ActiveConnection =    MM_EXTRANET_STRING_TEMP
                  strSQL = "SELECT count(refe05) as cuenta,SUM(vafa05) as ValorFac FROM d05artic where refe05 = '" & Rsio.Fields.Item("refe31").Value & "'"

                  RsCuenta.Source= strSQL
                  RsCuenta.CursorType = 0
                  RsCuenta.CursorLocation = 2
                  RsCuenta.LockType = 1
                  RsCuenta.Open()
                  if not RsCuenta.eof then
                     intTotaldeIPC = clng(RsCuenta.Fields.Item("cuenta").Value)
                     dblTotalFac = cdbl(RsCuenta.Fields.Item("ValorFac").Value)
                  end if
                  RsCuenta.close
                  set RsCuenta = Nothing


                While NOT RsMercancias.EOF
                 strHTML = strHTML&"<tr>" & chr(13) & chr(10)

                 intFactor = cdbl((cdbl(RsMercancias.fields("vafa05").value)  * 100 / dblTotalFac ) / 100)
				
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("refe31").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("cgas31").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strPedimento&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strCvePed&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFechaPago&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strProveedor_Cliente&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strPaisProveedor&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFacturaProveedor&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(strValorMercancia*intFactor,2)&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(strValorAduana*intFactor,2)&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& formatnumber(intDTA * intFactor,2) &"</font></td>" & chr(13) & chr(10)
                 intTotalImpuesto = intDTA * intFactor
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& formatnumber(intADV * intFactor,2)&"</font></td>" & chr(13) & chr(10)
                 intTotalImpuesto = intTotalImpuesto + ( intADV * intFactor)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(intIVA*intFactor,2) &"</font></td>" & chr(13) & chr(10)
                 intTotalImpuesto = intTotalImpuesto + ( intIVA * intFactor)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & formatnumber(cdbl(intPRV) / cdbl(intTotaldeIPC),2) & "</font></td>" & chr(13) & chr(10)
                 intTotalImpuesto = cdbl(intTotalImpuesto) + (cdbl(intPRV) / cdbl(intTotaldeIPC))

                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(intTotalImpuesto,2)&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strRegimen &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& formatnumber(cdbl( Rsio.fields("suph31").value/1.15)*intFactor,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& formatnumber(cdbl(( Rsio.fields("suph31").value/1.15) * 0.15) *intFactor,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)

                 SumaIva = 0
                 IvaProrrateado=0
		             SumaIva = cdbl(Rsio.fields("chon31").value) + cdbl(Rsio.fields("csce31").value + cdbl(Rsio.fields("caho31").value))
                 SumaIva = SumaIva * (cdbl(Rsio.fields("piva31").value)/100)

                 IvaProrrateado= (SumaIva * intFactor)
                 SumaServHon = 0
                 SumaServHon = cdbl(Rsio.fields("chon31").value) + cdbl(Rsio.fields("csce31").value)
                 SumaServHon = SumaServHon * intFactor
                 dblBaseIVA = 0
                 dblBaseIVA = 0

                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(SumaServHon,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber((cdbl(Rsio.fields("tota31").value) / (cdbl(Rsio.fields("piva31").value) / 100 + 1) * intFactor) ,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber((IvaProrrateado),2)&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(cdbl(Rsio.fields("tota31").value) *intFactor ,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsMercancias.Fields.Item("desc05").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(dblPesoBruto * intFactor,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strTipoOperacion&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.fields("fech31").value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsMercancias.Fields.Item("pedi05").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsMercancias.Fields.Item("item05").Value&"</font></td>" & chr(13) & chr(10)


                 strHTML = strHTML & "</tr>"& chr(13) & chr(10)
                 Response.Write(strHTML)
                 strHTML = ""
                 RsMercancias.MoveNext()
                Wend
            else

                 strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("refe31").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("cgas31").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strPedimento&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strCvePed&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFechaPago&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strProveedor_Cliente&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strPaisProveedor&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strFacturaProveedor&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(strValorMercancia*intFactor,2)&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(strValorAduana*intFactor,2)&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& formatnumber(intDTA ,2) &"</font></td>" & chr(13) & chr(10)
                 intTotalImpuesto = intDTA
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& formatnumber(intADV,2)&"</font></td>" & chr(13) & chr(10)
                 intTotalImpuesto = intTotalImpuesto + ( intADV)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(intIVA,2) &"</font></td>" & chr(13) & chr(10)
                 intTotalImpuesto = intTotalImpuesto + ( intIVA)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(intPRV,2)&"</font></td>" & chr(13) & chr(10)
                 intTotalImpuesto = cdbl(intTotalImpuesto) + cdbl(intPRV)

                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(intTotalImpuesto,2)&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& strRegimen &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& formatnumber(cdbl( Rsio.fields("suph31").value/1.15),2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& formatnumber(cdbl(( Rsio.fields("suph31").value/1.15) * 0.15),2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)

                 SumaIva = 0
                 IvaProrrateado=0
		             SumaIva = cdbl(Rsio.fields("chon31").value) + cdbl(Rsio.fields("csce31").value + cdbl(Rsio.fields("caho31").value))
                 SumaIva = SumaIva * (cdbl(Rsio.fields("piva31").value)/100)

                 IvaProrrateado= (SumaIva)
                 SumaServHon = 0
                 SumaServHon = cdbl(Rsio.fields("chon31").value) + cdbl(Rsio.fields("csce31").value)
                 SumaServHon = SumaServHon
                 dblBaseIVA = 0
                 dblBaseIVA = 0

                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(SumaServHon,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber((cdbl(Rsio.fields("tota31").value) / (cdbl(Rsio.fields("piva31").value) / 100 + 1) * intFactor) ,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber((IvaProrrateado),2)&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(cdbl(Rsio.fields("tota31").value) ,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&formatnumber(dblPesoBruto,2) &"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strTipoOperacion&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.fields("fech31").value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">NC</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">NC</font></td>" & chr(13) & chr(10)


                 strHTML = strHTML & "</tr>"& chr(13) & chr(10)
                 Response.Write(strHTML)
                 strHTML = ""


            end if

               RsMercancias.close
               set RsMercancias = Nothing

          Rsio.MoveNext()
      Wend

   Rsio.Close()
   Set Rsio = Nothing
  ' end if
 ' next
   strHTML = strHTML & "</table>"& chr(13) & chr(10)
   response.Write(strHTML)

 end if


%>