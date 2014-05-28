<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
 Server.ScriptTimeOut=200
 MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
 'MM_EXTRANET_STRING = ODBC_POR_ADUANA("VER")
 MM_EXTRANET_STRING_VER = ODBC_POR_ADUANA("VER")

 Response.Buffer = TRUE
 Response.Addheader "Content-Disposition", "attachment;"
 Response.ContentType = "application/vnd.ms-excel"

 strHTML = ""

 'clavecli = trim(request.Form("txtcvecli"))
 fechaini = trim(request.Form("txtDateIni"))
 fechafin = trim(request.Form("txtDateFin"))
 toper = trim(request.Form("rbnTipoDate"))

 'clavecli = trim("2054")
 'fechaini = trim("01/01/2005")
 'fechafin = trim("31/01/2005")
 'toper = trim("2")
 dim Rsio,Rsio2,Rsio3,Rsio4,Rsio5,Rsio6


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


if not Session("GAduana")="MAN" then


 if not fechaini="" and not fechafin="" then


    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    strDateIni = tmpAnioIni & "/" &tmpMesIni & "/"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    strDateFin = tmpAnioFin & "/" &tmpMesFin & "/"& tmpDiaFin

   if toper="1" then

      set Rsio = server.CreateObject("ADODB.Recordset")
      Rsio.ActiveConnection = MM_EXTRANET_STRING

      'strSQL = "select refcia01,rcli01,feta01,fdoc01,frev01,fdsp01,obser01 from ssdagi01,c01refer where (cveped01 <> 'A3' and cveped01 <> 'R1' and cveped01 <> 'F4') and cvecli01 = "&(clavecli)&" and (fecpag01 >='"&strDateIni&"' and fecpag01 <='"&strDateFin&"') and firmae01<>'' and refe01=refcia01"
      strSQL = "select refcia01,rcli01,ifnull(feta01,0000-00-00) feta01,ifnull(fdoc01,0000-00-00) fdoc01,ifnull(frev01,0000-00-00) frev01,ifnull(fdsp01,0000-00-00) fdsp01,obser01 from ssdagi01,c01refer where (cveped01 <> 'A3' and cveped01 <> 'R1' and cveped01 <> 'F4') " &  permi & " and (fecpag01 >='"&strDateIni&"' and fecpag01 <='"&strDateFin&"') and firmae01<>'' and refe01=refcia01"


      Rsio.Source= strSQL

      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()



	    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE INDICADORES</p></font></strong>"
	    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Importación</p></font></strong>"
      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & fechaini & " Al " & fechafin & "</p></font></strong>"
      strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	    strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Interno</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Doctos.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Reval.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Desp.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dias</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observ.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.ETA" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Destino" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Transporte" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Factur.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dias Fac.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Demoras</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Almacenaje</td>" & chr(13) & chr(10)
      strHTML = strHTML & "</tr>"& chr(13) & chr(10)

      cuenta = 0
      While NOT Rsio.EOF
          cuenta = cuenta + 1
          dias=0
          xFechaCuenta="  /  /    "
          xAlmacenaje=0
          xDemoras=0
          NumDias=0

        if cuenta = 21 then

         Response.Write(strHTML)

         Response.End
         end if

          if not Rsio.Fields.Item("fdsp01").Value="" and not Rsio.Fields.Item("frev01").Value=""  and not Rsio.Fields.Item("fdsp01").Value="0000-00-00" and not Rsio.Fields.Item("frev01").Value="0000-00-00" then

             dias=DateDiff("d",Rsio.Fields.Item("frev01").Value,Rsio.Fields.Item("fdsp01").Value)-QuitaSabadoDomingo(Rsio.Fields.Item("fdsp01").Value,Rsio.Fields.Item("frev01").Value)-QuitaDiasFestivos(Rsio.Fields.Item("fdsp01").Value,Rsio.Fields.Item("frev01").Value)-1
             if dias<1 then
                dias=0
             end if
          end if

          set Rsio2 = server.CreateObject("ADODB.Recordset")
          Rsio2.ActiveConnection = MM_EXTRANET_STRING

          strSQL2 = "select fech31 from e31cgast,d31refer where d31refer.cgas31 = e31cgast.cgas31 and d31refer.refe31 = '" &Rsio.Fields.Item("refcia01").Value & "'"

          Rsio2.Source= strSQL2

          Rsio2.CursorType = 0
          Rsio2.CursorLocation = 2
          Rsio2.LockType = 1
          Rsio2.Open()

          if not Rsio2.EOF then
             xFechaCuenta=Rsio2.Fields.Item("fech31").Value
          else
             xFechaCuenta="  /  /    "

          end if


          set Rsio3 = server.CreateObject("ADODB.Recordset")
          Rsio3.ActiveConnection = MM_EXTRANET_STRING

          strSQL3 = "select refe21 referencia,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Monto from d21paghe,e21paghe,c21paghe where YEAR(e21paghe.fech21)=YEAR(d21paghe.fech21) and e21paghe.foli21=d21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 and conc21 = clav21 AND (conc21=4) AND refe21 = '" &Rsio.Fields.Item("refcia01").Value & "'  and tpag21=1 group by refe21,desc21 order by refe21"

          Rsio3.Source= strSQL3
          Rsio3.CursorType = 0
          Rsio3.CursorLocation = 2
          Rsio3.LockType = 1
          Rsio3.Open()

          if not Rsio3.EOF then
             xAlmacenaje=Rsio3.Fields.Item("Monto").Value
          else
             xAlmacenaje=0
          end if

          set Rsio4 = server.CreateObject("ADODB.Recordset")
          Rsio4.ActiveConnection = MM_EXTRANET_STRING

          strSQL4 = "select refe21 referencia,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Monto from d21paghe,e21paghe,c21paghe where YEAR(e21paghe.fech21)=YEAR(d21paghe.fech21) and e21paghe.foli21=d21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 and conc21 = clav21 AND (conc21=11) AND refe21 = '" &Rsio.Fields.Item("refcia01").Value & "'  and tpag21=1 group by refe21,desc21 order by refe21"

          Rsio4.Source= strSQL4
          Rsio4.CursorType = 0
          Rsio4.CursorLocation = 2
          Rsio4.LockType = 1
          Rsio4.Open()

          if not Rsio4.EOF then
             xDemoras=Rsio4.Fields.Item("Monto").Value
          else
             xDemoras=0
          end if



          if not xFechaCuenta="  /  /    " then
             if not Rsio.Fields.Item("fdsp01").Value = "0000-00-00"   then
              NumDias=DateDiff("d",Rsio.Fields.Item("fdsp01").Value,xFechaCuenta)-QuitaSabadoDomingo(xFechaCuenta,Rsio.Fields.Item("fdsp01").Value)-QuitaDiasFestivos(Rsio.Fields.Item("fdsp01").Value,xFechaCuenta)-5

       	      if NumDias<0 then
		             NumDias=0
	            end if
              else
                 NumDias=0
              end if
          else
              NumDias=0
          end if


         strHTML = strHTML&"<tr>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("refcia01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("rcli01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fdoc01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("frev01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fdsp01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dias&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("obser01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("feta01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">""</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">""</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xFechaCuenta&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&NumDias&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xDemoras&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xAlmacenaje&"</font></td>" & chr(13) & chr(10)

         Rsio.MoveNext()
      Wend
    Rsio.close
    set Rsio = Nothing
   else

      set Rsio = server.CreateObject("ADODB.Recordset")
      Rsio.ActiveConnection = MM_EXTRANET_STRING

      'strSQL = "select refcia01,desf0101,feta01,cvepod01,cvemta01,fdoc01,finmer10,fdsp01,obser01 from ssdage01,c01refer,e10art23 where cveped01 <> 'BB' and cvecli01 = "&(clavecli)&" and (fecpag01>='"&strDateIni&"'and fecpag01<='"&strDateFin&"') and firmae01<>'' and refe01 = refcia01 and refe10 = refcia01"
      strSQL = "select refcia01,desf0101,feta01,cvepod01,cvemta01,fdoc01,finmer10,fdsp01,obser01 from ssdage01,c01refer,e10art23 where cveped01 <> 'BB' " &  permi & " and (fecpag01>='"&strDateIni&"'and fecpag01<='"&strDateFin&"') and firmae01<>'' and refe01 = refcia01 and refe10 = refcia01"

      Rsio.Source= strSQL
      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()


      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE INDICADORES</p></font></strong>"
	    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Exportación</p></font></strong>"
      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & fechaini & " Al " & fechafin & "</p></font></strong>"
      strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	    strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Facturas</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Doctos.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.ingreso</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Desp.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dias</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observ.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.ETA" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Destino" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Transporte" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Factur.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dias Fac.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Demoras</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Almacenaje</td>" & chr(13) & chr(10)
      strHTML = strHTML & "</tr>"& chr(13) & chr(10)

      While NOT Rsio.EOF

          dias=0
          xFechaCuenta="  /  /    "
          xAlmacenaje=0
          xDemoras=0
          NumDias=0
          xPais=""
          xTrans=""

          if not Rsio.Fields.Item("fdsp01").Value="" and not Rsio.Fields.Item("finmer10").Value="" then
             dias=DateDiff("d",Rsio.Fields.Item("finmer10").Value,Rsio.Fields.Item("fdsp01").Value)-QuitaSabadoDomingo(Rsio.Fields.Item("fdsp01").Value,Rsio.Fields.Item("finmer10").Value)-QuitaDiasFestivos(Rsio.Fields.Item("fdsp01").Value,Rsio.Fields.Item("finmer10").Value)-1
             if dias<1 then
                dias=0
             end if
          end if

          set Rsio2 = server.CreateObject("ADODB.Recordset")
          Rsio2.ActiveConnection = MM_EXTRANET_STRING


          strSQL2 = "select fech31 from e31cgast ,d31refer  where d31refer.cgas31 = e31cgast.cgas31 and d31refer.refe31 = '" &Rsio.Fields.Item("refcia01").Value & "'"

          Rsio2.Source= strSQL2
          Rsio2.CursorType = 0
          Rsio2.CursorLocation = 2
          Rsio2.LockType = 1
          Rsio2.Open()

          if not Rsio2.EOF then
             xFechaCuenta=Rsio2.Fields.Item("fech31").Value
          else
             xFechaCuenta="  /  /    "
          end if

          set Rsio3 = server.CreateObject("ADODB.Recordset")
          Rsio3.ActiveConnection = MM_EXTRANET_STRING

          strSQL3 = "select refe21 referencia,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Monto from d21paghe,e21paghe,c21paghe where YEAR(e21paghe.fech21)=YEAR(d21paghe.fech21) and e21paghe.foli21=d21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 and conc21 = clav21 and (conc21=4) AND refe21 = '" &Rsio.Fields.Item("refcia01").Value & "' and tpag21=1 group by refe21,desc21 order by refe21"

          Rsio3.Source= strSQL3
          Rsio3.CursorType = 0
          Rsio3.CursorLocation = 2
          Rsio3.LockType = 1
          Rsio3.Open()

          if not Rsio3.EOF then
             xAlmacenaje=Rsio3.Fields.Item("Monto").Value
          else
             xAlmacenaje=0
          end if

          set Rsio4 = server.CreateObject("ADODB.Recordset")
          Rsio4.ActiveConnection = MM_EXTRANET_STRING

          strSQL4 = "select refe21 referencia,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Monto from d21paghe,e21paghe,c21paghe where YEAR(e21paghe.fech21)=YEAR(d21paghe.fech21) and e21paghe.foli21=d21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 and conc21 = clav21 AND (conc21=11) AND refe21 = '" &Rsio.Fields.Item("refcia01").Value & "'  and tpag21=1 group by refe21,desc21 order by refe21"

          Rsio4.Source= strSQL4
          Rsio4.CursorType = 0
          Rsio4.CursorLocation = 2
          Rsio4.LockType = 1
          Rsio4.Open()

          if not Rsio4.EOF then
             xDemoras=Rsio4.Fields.Item("Monto").Value
          else
             xDemoras=0
          end if


          set Rsio5 = server.CreateObject("ADODB.Recordset")
          Rsio5.ActiveConnection = MM_EXTRANET_STRING

          strSQL5 = "select cvepai19,nompai19 from sspais19 where cvepai19 = '" &Rsio.Fields.Item("cvepod01").Value & "' "

          Rsio5.Source= strSQL5
          Rsio5.CursorType = 0
          Rsio5.CursorLocation = 2
          Rsio5.LockType = 1
          Rsio5.Open()

          if not Rsio5.EOF then
             xPais=Rsio5.Fields.Item("nompai19").Value
          else
             xPais=""
          end if

          set Rsio6 = server.CreateObject("ADODB.Recordset")
          Rsio6.ActiveConnection = MM_EXTRANET_STRING

          strSQL6 = "select clavet30,descri30 from ssmtra30 where clavet30 = '" &Rsio.Fields.Item("cvemta01").Value & "' "

          Rsio6.Source= strSQL6
          Rsio6.CursorType = 0
          Rsio6.CursorLocation = 2
          Rsio6.LockType = 1
          Rsio6.Open()

          if not Rsio6.EOF then
             xTrans=Rsio6.Fields.Item("descri30").Value
          else
             xTrans=""
          end if

          if not xFechaCuenta="  /  /    " then
              if not Rsio.Fields.Item("fdsp01").Value = "  /  /    " then
              NumDias=DateDiff("d",Rsio.Fields.Item("fdsp01").Value,xFechaCuenta)-QuitaSabadoDomingo(xFechaCuenta,Rsio.Fields.Item("fdsp01").Value)-QuitaDiasFestivos(Rsio.Fields.Item("fdsp01").Value,xFechaCuenta)-5
	            if NumDias<0 then
		             NumDias=0
	            end if
              else
              NumDias=0
              end if
          else
              NumDias=0
          end if


         strHTML = strHTML&"<tr>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("refcia01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("desf0101").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fdoc01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("finmer10").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fdsp01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dias&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("obser01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("feta01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xPais&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xTrans&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xFechaCuenta&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&NumDias&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xDemoras&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xAlmacenaje&"</font></td>" & chr(13) & chr(10)

         Rsio.MoveNext()
      Wend

   end if

   Rsio.Close()
   Set Rsio = Nothing

 end if

else
  ' para Manzanillo

  if not fechaini="" and not fechafin="" then


    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    strDateIni = tmpAnioIni & "/" &tmpMesIni & "/"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    strDateFin = tmpAnioFin & "/" &tmpMesFin & "/"& tmpDiaFin

   if toper="1" then

      set Rsio = server.CreateObject("ADODB.Recordset")
      Rsio.ActiveConnection = MM_EXTRANET_STRING

      strSQL = "select refcia01,rcli01,feta01,fdoc01,frev01,fdsp01,obser01 from ssdagi01,c01refer where (cveped01 <> 'A3' and cveped01 <> 'R1' and cveped01 <> 'F4') and cvecli01 = "&(clavecli)&" and (fecpag01 >='"&strDateIni&"' and fecpag01 <='"&strDateFin&"') and firmae01<>'' and refe01=refcia01"
     strSQL = "select refcia01,rcli01,ifnull(feta01,0000-00-00) feta01,ifnull(fdoc01,0000-00-00) fdoc01,ifnull(frev01,0000-00-00) frev01,ifnull(fdsp01,0000-00-00) fdsp01,obser01  from ssdagi01,c01refer where (cveped01 <> 'A3' and cveped01 <> 'R1' and cveped01 <> 'F4') " &  permi & " and (fecpag01 >='"&strDateIni&"' and fecpag01 <='"&strDateFin&"') and firmae01<>'' and refe01=refcia01"

      Rsio.Source= strSQL
      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()


	    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE INDICADORES</p></font></strong>"
	    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Importación</p></font></strong>"
      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & fechaini & " Al " & fechafin & "</p></font></strong>"
      strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	    strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Interno</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Doctos.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Reval.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Desp.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dias</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observ.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.ETA" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Destino" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Transporte" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Factur.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dias Fac.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Demoras</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Almacenaje</td>" & chr(13) & chr(10)
      strHTML = strHTML & "</tr>"& chr(13) & chr(10)

      While NOT Rsio.EOF

          dias=0
          xFechaCuenta="  /  /    "
          xCuenta = ""
          xAlmacenaje=0
          xDemoras=0
          NumDias=0


          set Rsio2 = server.CreateObject("ADODB.Recordset")
          Rsio2.ActiveConnection = MM_EXTRANET_STRING_VER

          strSQL2 = "select fech31,e31cgast.cgas31 from e31cgast,d31refer where d31refer.cgas31 = e31cgast.cgas31 and d31refer.refe31 = '" &Rsio.Fields.Item("refcia01").Value & "'"

          Rsio2.Source= strSQL2
          Rsio2.CursorType = 0
          Rsio2.CursorLocation = 2
          Rsio2.LockType = 1
          Rsio2.Open()

          if not Rsio2.EOF then
             xFechaCuenta= Rsio2.Fields.Item("fech31").Value
             xCuenta= Rsio2.Fields.Item("cgas31").Value
          else
             xFechaCuenta="  /  /    "
          end if

         Rsio2.close
         set Rsio2 = Nothing



         strHTML = strHTML&"<tr>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("refcia01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("rcli01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fdoc01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("frev01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fdsp01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& xCuenta&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xFechaCuenta&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)

         strHTML = strHTML&"</tr>" & chr(13) & chr(10)
         Rsio.MoveNext()
      Wend
      Rsio.close
      set Rsio = Nothing

 strHTML = strHTML & "</table>"& chr(13) & chr(10)
 Response.Write( strHTML)
Response.End
   else

      set Rsio = server.CreateObject("ADODB.Recordset")
      Rsio.ActiveConnection = MM_EXTRANET_STRING
      'strSQL = "select refcia01,desf0101,feta01,cvepod01,cvemta01,fdoc01,finmer10,fdsp01,obser01 from ssdage01,c01refer,e10art23 where cveped01 <> 'BB' and cvecli01 = "&(clavecli)&" and (fecpag01>='"&strDateIni&"'and fecpag01<='"&strDateFin&"') and firmae01<>'' and refe01 = refcia01 and refe10 = refcia01"
      strSQL = "select refcia01,desf0101,feta01,cvepod01,cvemta01,fdoc01,fdsp01,obser01 from ssdage01,c01refer where cveped01 <> 'BB' " &  permi & " and (fecpag01>='"&strDateIni&"'and fecpag01<='"&strDateFin&"') and firmae01<>'' and refe01 = refcia01 "
      Rsio.Source= strSQL
      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()


      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE INDICADORES</p></font></strong>"
	    strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p></p></font></strong>"
      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Exportación</p></font></strong>"
      strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & fechaini & " Al " & fechafin & "</p></font></strong>"
      strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	    strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
	    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Facturas</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Doctos.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.ingreso</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Desp.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dias</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observ.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.ETA" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Destino" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Transporte" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">F.Factur.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dias Fac.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Demoras</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Almacenaje</td>" & chr(13) & chr(10)
      strHTML = strHTML & "</tr>"& chr(13) & chr(10)

      While NOT Rsio.EOF

          dias=0
          xFechaCuenta="  /  /    "
          xAlmacenaje=0
          xDemoras=0
          NumDias=0

          ' if not Rsio.Fields.Item("fdsp01").Value="" and not Rsio.Fields.Item("finmer10").Value="" then
          if not Rsio.Fields.Item("fdsp01").Value=""  then
             ' dias=DateDiff("d",Rsio.Fields.Item("finmer10").Value,Rsio.Fields.Item("fdsp01").Value)-QuitaSabadoDomingo(Rsio.Fields.Item("fdsp01").Value,Rsio.Fields.Item("finmer10").Value)-QuitaDiasFestivos(Rsio.Fields.Item("fdsp01").Value,Rsio.Fields.Item("finmer10").Value)-1
             dias=DateDiff("d",Rsio.Fields.Item("fdsp01").Value - 2,Rsio.Fields.Item("fdsp01").Value)-QuitaSabadoDomingo(Rsio.Fields.Item("fdsp01").Value,Rsio.Fields.Item("fdsp01").Value- 2)-QuitaDiasFestivos(Rsio.Fields.Item("fdsp01").Value,Rsio.Fields.Item("fdsp01").Value - 2)-1
             if dias<1 then
                dias=0
             end if
          end if

          set Rsio2 = server.CreateObject("ADODB.Recordset")
          Rsio2.ActiveConnection = MM_EXTRANET_STRING_VER


          strSQL2 = "select fech31 from e31cgast ,d31refer  where d31refer.cgas31 = e31cgast.cgas31 and d31refer.refe31 = '" &Rsio.Fields.Item("refcia01").Value & "'"

          Rsio2.Source= strSQL2
          Rsio2.CursorType = 0
          Rsio2.CursorLocation = 2
          Rsio2.LockType = 1
          Rsio2.Open()

          if not Rsio2.EOF then
             xFechaCuenta=Rsio2.Fields.Item("fech31").Value
          else
             xFechaCuenta="  /  /    "
          end if

          set Rsio3 = server.CreateObject("ADODB.Recordset")
          Rsio3.ActiveConnection = MM_EXTRANET_STRING_VER

          strSQL3 = "select refe21 referencia,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Monto from d21paghe,e21paghe,c21paghe where YEAR(e21paghe.fech21)=YEAR(d21paghe.fech21) and e21paghe.foli21=d21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 and conc21 = clav21 and (conc21=4) AND refe21 = '" &Rsio.Fields.Item("refcia01").Value & "' and tpag21=1 group by refe21,desc21 order by refe21"

          Rsio3.Source= strSQL3
          Rsio3.CursorType = 0
          Rsio3.CursorLocation = 2
          Rsio3.LockType = 1
          Rsio3.Open()

          if not Rsio3.EOF then
             xAlmacenaje=Rsio3.Fields.Item("Monto").Value
          else
             xAlmacenaje=0
          end if

          set Rsio4 = server.CreateObject("ADODB.Recordset")
          Rsio4.ActiveConnection = MM_EXTRANET_STRING_VER

          strSQL4 = "select refe21 referencia,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Monto from d21paghe,e21paghe,c21paghe where YEAR(e21paghe.fech21)=YEAR(d21paghe.fech21) and e21paghe.foli21=d21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 and conc21 = clav21 AND (conc21=11) AND refe21 = '" &Rsio.Fields.Item("refcia01").Value & "'  and tpag21=1 group by refe21,desc21 order by refe21"

          Rsio4.Source= strSQL4
          Rsio4.CursorType = 0
          Rsio4.CursorLocation = 2
          Rsio4.LockType = 1
          Rsio4.Open()

          if not Rsio4.EOF then
             xDemoras=Rsio4.Fields.Item("Monto").Value
          else
             xDemoras=0
          end if

          set Rsio5 = server.CreateObject("ADODB.Recordset")
          Rsio5.ActiveConnection = MM_EXTRANET_STRING

          strSQL5 = "select cvepai19,nompai19 from sspais19 where cvepai19 = '" &Rsio.Fields.Item("cvepod01").Value & "' "

          Rsio5.Source= strSQL5
          Rsio5.CursorType = 0
          Rsio5.CursorLocation = 2
          Rsio5.LockType = 1
          Rsio5.Open()

          if not Rsio5.EOF then
             xPais=Rsio5.Fields.Item("nompai19").Value
          else
             xPais=""
          end if

          set Rsio6 = server.CreateObject("ADODB.Recordset")
          Rsio6.ActiveConnection = MM_EXTRANET_STRING

          strSQL6 = "select clavet30,descri30 from ssmtra30 where clavet30 = '" &Rsio.Fields.Item("cvemta01").Value & "' "

          Rsio6.Source= strSQL6
          Rsio6.CursorType = 0
          Rsio6.CursorLocation = 2
          Rsio6.LockType = 1
          Rsio6.Open()

          if not Rsio6.EOF then
             xTrans=Rsio6.Fields.Item("descri30").Value
          else
             xTrans=""
          end if

          if not xFechaCuenta="  /  /    " then
             if not Rsio.Fields.Item("fdsp01").Value = "  /  /    " then
             NumDias=DateDiff("d",Rsio.Fields.Item("fdsp01").Value,xFechaCuenta)-QuitaSabadoDomingo(xFechaCuenta,Rsio.Fields.Item("fdsp01").Value)-QuitaDiasFestivos(Rsio.Fields.Item("fdsp01").Value,xFechaCuenta)-5
	           if NumDias<0 then
		            NumDias=0
	           end if
             else
              NumDias=0
            end if
          else
              NumDias=0
          end if


         strHTML = strHTML&"<tr>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("refcia01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("desf0101").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fdoc01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fdsp01").Value -2 &"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("fdsp01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dias&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("obser01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("feta01").Value&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xPais&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xTrans&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xFechaCuenta&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&NumDias&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xDemoras&"</font></td>" & chr(13) & chr(10)
         strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&xAlmacenaje&"</font></td>" & chr(13) & chr(10)

         Rsio.MoveNext()
      Wend

   end if

   Rsio.Close()
   Set Rsio = Nothing

 end if


end if

 strHTML = strHTML & "</tr>"& chr(13) & chr(10)
 strHTML = strHTML & "</td>"& chr(13) & chr(10)
 strHTML = strHTML & "</table>"& chr(13) & chr(10)
 strHTML = strHTML & "</table>"& chr(13) & chr(10)
 response.Write(strHTML)

Function QuitaSabadoDomingo(fechi,fechf)


    sabdoms=0
    xdias=0
    xdias=DateDiff("d",fechf,fechi)

    for x=1 to xdias
	     if Weekday(DateAdd("d",-x,fechi))=1 or Weekday(DateAdd("d",-x,fechi))=7 then
		      sabdoms = sabdoms + 1
	     end if
    next
    QuitaSabadoDomingo=sabdoms

End Function



Function QuitaDiasFestivos(fechi,fechf)

    diasFestivos=0
    xdias=0
    xdias=DateDiff("d",fechf,fechi)

    for x=0 to xdias

    revisafecha ="  /  /    "
    revisafecha = DateAdd("d",x,fechf)

    if (Weekday(revisafecha) => 2 and Weekday(revisafecha) <=6) then

       intmes= MONTH(DateAdd("d",x,fechf))
       intDia = DAY(DateAdd("d",x,fechf))

       select CASE intmes
       CASE 1
       if intDia = 1 then
          diasFestivos = diasFestivos + 1
       end if
       CASE 2
       if intDia = 5 then
          diasFestivos = diasFestivos + 1
       end if
       CASE 3
       if intDia = 21 then
          diasFestivos = diasFestivos + 1
       end if
       CASE 5
       if intDia = 1 then
          diasFestivos = diasFestivos + 1
       end if
       CASE 9
       if intDia = 16 then
          diasFestivos = diasFestivos + 1
       end if
       CASE 11
       if intDia = 2 then
          diasFestivos = diasFestivos + 1
       end if
       if intDia = 20 then
          diasFestivos = diasFestivos + 1
       end if
       CASE 12
       if intDia = 25 then
          diasFestivos = diasFestivos + 1
       end if
       END select

   end if
   next

   QuitaDiasFestivos=diasFestivos

End Function

%>
