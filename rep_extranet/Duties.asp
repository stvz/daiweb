<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
 Server.ScriptTimeOut=200
 MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

 strHTML = ""

 fechaini = trim(request.Form("txtDateIni"))
 fechafin = trim(request.Form("txtDateFin"))

 dim Rsio,Rsio2,Rsio3,Rsio4

 strPermisos = Request.Form("Permisos")
 permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

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


if not fechaini="" and not fechafin="" then
  ' if ChecarReferencias(fechaini,fechafin)<> 0 then
      ok=GenerarReporte(fechaini,fechafin)
      response.Write(strHTML)
  ' else
   %>
   <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>El reporte aún esta incompleto...</p></font></strong>
   <%
   'end if

end if

Function GenerarReporte(fechi,fechf)

    Response.Buffer = TRUE
    ' Response.Addheader "Content-Disposition", "attachment;"
    ' Response.ContentType = "application/vnd.ms-excel"

    fechaini = fechi
    fechafin = fechf

    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    strDateIni = tmpAnioIni & "/" &tmpMesIni & "/"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    strDateFin = tmpAnioFin & "/" &tmpMesFin & "/"& tmpDiaFin


      set Rsio = server.CreateObject("ADODB.Recordset")
      Rsio.ActiveConnection = MM_EXTRANET_STRING

      'strSQL = "select refcia01,patent01,cvepod01,cvepvc01,numped01,fecpag01,adusec01,firmae01 from ssdagi01 where rfccli01='IFF610526PQ6' and (fecpag01 >='"&strDateIni&"'and fecpag01 <='"&strDateFin&"') and firmae01<>'' and cveped01<>'R1'"
      strSQL = "select refcia01,patent01,cvepod01,cvepvc01,numped01,fecpag01,adusec01,firmae01 from ssdagi01 where (fecpag01 >='"&strDateIni&"'and fecpag01 <='"&strDateFin&"') and firmae01<>'' and cveped01<>'R1' " &  permi & ""
	  ' Response.Write(strSQL)
	  ' Response.End()
      Rsio.Source= strSQL
      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()

      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DUTIES</p></font></strong>"
	    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & fechaini & " Al " & fechafin & "</p></font></strong>"
      strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	    strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Numero pedimento" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Proveedor</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Factura</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de pago</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais de origen</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais de procedencia</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor aduana</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fraccion</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripcion</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Peso</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Unidad</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IPC</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Orden de compra</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DTA</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tasa ADV</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">ADV</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA<</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Total</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Puerto</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Si TLC</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">NTLC Razon</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">NPS Razon</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observaciones</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">orden fraccion</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Apendice</td>" & chr(13) & chr(10)
      strHTML = strHTML & "</tr>"& chr(13) & chr(10)

      While NOT Rsio.EOF

          set Rsio2 = server.CreateObject("ADODB.Recordset")
          Rsio2.ActiveConnection = MM_EXTRANET_STRING


          strSQL2 = "select apen01 from c01refer where refe01='" &Rsio.Fields.Item("refcia01").Value & "'"

          Rsio2.Source= strSQL2
          Rsio2.CursorType = 0
          Rsio2.CursorLocation = 2
          Rsio2.LockType = 1
          Rsio2.Open()

             While NOT Rsio2.EOF

                 set Rsio3 = server.CreateObject("ADODB.Recordset")
                 Rsio3.ActiveConnection = MM_EXTRANET_STRING

                 strSQL3 = "Select d05artic.refe05,d05artic.agru05,ssprov22.nompro22,d05artic.fact05,d05artic.frac05,d05artic.desc05,d05artic.cata05,d05artic.umta05,d05artic.item05,d05artic.pedi05,vafa05 From	d05artic,ssprov22 where d05artic.refe05='" &Rsio.Fields.Item("refcia01").Value & "' and d05artic.prov05=ssprov22.cvepro22"

                 Rsio3.Source= strSQL3
                 Rsio3.CursorType = 0
                 Rsio3.CursorLocation = 2
                 Rsio3.LockType = 1
                 Rsio3.Open()

                 While NOT Rsio3.EOF

                     set Rsio4 = server.CreateObject("ADODB.Recordset")
                     Rsio4.ActiveConnection = MM_EXTRANET_STRING

                     strSQL4 = "Select ssfrac02.vaduan02 * '" &Rsio3.Fields.Item("vafa05").Value & "' / ssfrac02.vmerme02 as Val_adu,ssfrac02.dtafpp02 * '" &Rsio3.Fields.Item("vafa05").Value & "' / ssfrac02.vmerme02	as DTA,ssfrac02.tasadv02 as Tasa_ADV,(ssfrac02.i_adv102 + ssfrac02.i_adv202 + ssfrac02.i_adv302) * '" &Rsio3.Fields.Item("vafa05").Value & "' / ssfrac02.vmerme02	as ADV,(ssfrac02.i_iva102 + ssfrac02.i_iva202 + ssfrac02.i_iva302) * '" &Rsio3.Fields.Item("vafa05").Value & "' / ssfrac02.vmerme02	as IVA,(ssfrac02.dtafpp02 * '" &Rsio3.Fields.Item("vafa05").Value & "' / ssfrac02.vmerme02)+(ssfrac02.i_adv102 + ssfrac02.i_adv202 + ssfrac02.i_adv302) * '" &Rsio3.Fields.Item("vafa05").Value & "' / ssfrac02.vmerme02 + (ssfrac02.i_iva102 + ssfrac02.i_iva202 + ssfrac02.i_iva302) * '" &Rsio3.Fields.Item("vafa05").Value & "' / ssfrac02.vmerme02 as Total,ssfrac02.ntlc02 as NTLC_Razon,ssfrac02.nps02 as NPS_Razon,ssfrac02.ordfra02 From	ssfrac02 where ssfrac02.refcia02='" &Rsio3.Fields.Item("refe05").Value & "' and ssfrac02.ordfra02='" &Rsio3.Fields.Item("agru05").Value & "'"

                     Rsio4.Source= strSQL4
                     Rsio4.CursorType = 0
                     Rsio4.CursorLocation = 2
                     Rsio4.LockType = 1
                     Rsio4.Open()

                    While NOT Rsio4.EOF

                        set Rsio5 = server.CreateObject("ADODB.Recordset")
                        Rsio5.ActiveConnection = MM_EXTRANET_STRING

                        strSQL5 = "select cveide12 from ssipar12 where ssipar12.refcia12='" &Rsio3.Fields.Item("refe05").Value & "' and ssipar12.ordfra12='" &Rsio4.Fields.Item("ordfra02").Value & "' group by ssipar12.refcia12"

                        Rsio5.Source= strSQL5
                        Rsio5.CursorType = 0
                        Rsio5.CursorLocation = 2
                        Rsio5.LockType = 1
                        Rsio5.Open()

                           While NOT Rsio5.EOF


                                strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio3.Fields.Item("refe05").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("patent01").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("numped01").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio3.Fields.Item("nompro22").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio3.Fields.Item("fact05").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fecpag01").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("cvepod01").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("cvepvc01").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio4.Fields.Item("Val_adu").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio3.Fields.Item("frac05").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio3.Fields.Item("desc05").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio3.Fields.Item("cata05").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio3.Fields.Item("umta05").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio3.Fields.Item("item05").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio3.Fields.Item("pedi05").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio4.Fields.Item("DTA").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio4.Fields.Item("Tasa_ADV").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio4.Fields.Item("ADV").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio4.Fields.Item("IVA").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio4.Fields.Item("Total").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("adusec01").Value&"</font></td>" & chr(13) & chr(10)
                                if trim(Rsio5.Fields.Item("cveide12").Value)="TL" then
                                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">SI</font></td>" & chr(13) & chr(10)
                                else
                                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">NO</font></td>" & chr(13) & chr(10)
                                end if
                                if trim(Rsio5.Fields.Item("cveide12").Value)="PS" then
                                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">SI</font></td>" & chr(13) & chr(10)
                                else
                                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">NO</font></td>" & chr(13) & chr(10)
                                end if

                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio4.Fields.Item("NPS_Razon").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio4.Fields.Item("ordfra02").Value&"</font></td>" & chr(13) & chr(10)
                                strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio2.Fields.Item("apen01").Value&"</font></td>" & chr(13) & chr(10)
                                Response.Write( strHTML)
                                strHTML = ""
                              Rsio5.MoveNext()
                           Wend

                       Rsio4.MoveNext()
                    Wend

                    Rsio3.MoveNext()
                 Wend

               Rsio2.MoveNext()
             Wend

          Rsio.MoveNext()
      Wend

   Rsio.Close()
   Set Rsio = Nothing

   strHTML = strHTML & "</tr>"& chr(13) & chr(10)
   strHTML = strHTML & "</td>"& chr(13) & chr(10)
   strHTML = strHTML & "</table>"& chr(13) & chr(10)
   strHTML = strHTML & "</table>"& chr(13) & chr(10)

   GenerarReporte=true

End Function



Function ChecarReferencias(fechi,fechf)

    fechaini = fechi
    fechafin = fechf
    cont=0

    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    strDateIni = tmpAnioIni & "/" &tmpMesIni & "/"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    strDateFin = tmpAnioFin & "/" &tmpMesFin & "/"& tmpDiaFin


      set Rsio = server.CreateObject("ADODB.Recordset")
      Rsio.ActiveConnection = MM_EXTRANET_STRING

      'strSQL = "Select distinct refcia01 as REF_PEDTO from ssdagi01 Where rfccli01='IFF610526PQ6' and (fecpag01 >='"&strDateIni&"'and fecpag01 <='"&strDateFin&"') and firmae01<>'' and cveped01<>'R1'"
      strSQL = "Select distinct refcia01 as REF_PEDTO from ssdagi01 Where (fecpag01 >='"&strDateIni&"'and fecpag01 <='"&strDateFin&"') and firmae01<>'' and cveped01<>'R1' " &  permi & ""

      Rsio.Source= strSQL
      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()

      While NOT Rsio.EOF

          set Rsio2 = server.CreateObject("ADODB.Recordset")
          Rsio2.ActiveConnection = MM_EXTRANET_STRING

          strSQL2 = "Select distinct d05artic.refe05 as REF_MERCANCIAS From	d05artic,ssprov22 where d05artic.refe05='" &Rsio.Fields.Item("REF_PEDTO").Value & "' and d05artic.prov05=ssprov22.cvepro22 "

          Rsio2.Source= strSQL2
          Rsio2.CursorType = 0
          Rsio2.CursorLocation = 2
          Rsio2.LockType = 1
          Rsio2.Open()

          if not Rsio2.EOF then
              While NOT Rsio2.EOF

                Rsio2.MoveNext()
              Wend
          else
             cont=cont+1
          end if

          Rsio2.Close()
          Set Rsio2 = Nothing

          Rsio.MoveNext()
      Wend

   Rsio.Close()
   Set Rsio = Nothing

ChecarReferencias=cont

End Function
%>
