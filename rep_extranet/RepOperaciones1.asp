<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%


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
	  strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CvePedimento</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha Pago</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Facturas</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pais Orige/Destino</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pies</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Arribo</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Permisos</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">No.Contenedores</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Contenedores</td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Incoterm</td>" & chr(13) & chr(10)
    if strTipoOperaciones = 1 then
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Proveedor</td>" & chr(13) & chr(10)
    ELSE
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cliente</td>" & chr(13) & chr(10)
    END IF
    strHTML = strHTML & "</tr>"& chr(13) & chr(10)
  MM_EXTRANET_STRING_TEMP = ""
  for x=1 to 6
    if x=1 then
       aduana="VER"
    end if
    if x=2 then
       aduana="MAN"
    end if
    if x=3 then
       aduana="TAM"
    end if
    if x=4 then
       aduana="MEX"
    end if
    if x=5 then
       aduana="LZR"
    end if
    if x=6 then
       aduana="TOL"
    end if

    MM_EXTRANET_STRING_TEMP = ODBC_POR_ADUANA(aduana)

    permi = PermisoClientes(aduana,strPermisos,"cvecli01")


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

     if not permi= "" then
	 'Set OBJdbConnection = Server.CreateObject("ADODB.Connection")
      set Rsio = server.CreateObject("ADODB.Recordset")
	  	'response.write(" INICIO:"& MM_EXTRANET_STRING_TEMP & ":MM_EXTRANET_STRING_TEMP, aduana: "& aduana &"<BR>" ) 
	'response.end()
      Rsio.ActiveConnection =    MM_EXTRANET_STRING_TEMP

      if strTipoOperaciones = 1 then
         strSQL = "select distinct refcia01 Referencia,cveped01, numped01 Pedimento, adusec01 aduana, fecpag01 FechaPago, desf0101 Facturas, nompai19 PaisDestino, ifnull(piesco40,0)  Pies, descri30 Arribo,NOMPRO01 from  ssmtra30,sspais19,ssdagi01  left join sscont40  on refcia01=refcia40 where cvepod01=cvepai19 and cvemta01  = clavet30  and firmae01<>'' and  fecpag01>='"&strDateIni&"' and  fecpag01<='"&strDateFin&"' and cveped01<>'R1' " &  permi & " order by fecpag01"
        IF aduana="LAR" THEN
            strSQL = "select distinct refcia01 Referencia,cveped01, numped01 Pedimento, adusec01 aduana, fecpag01 FechaPago, desf0101 Facturas, nompai19 PaisDestino, 0 as Pies, descri30 Arribo from  ssmtra30,sspais19,ssdagi01  where cvepod01=cvepai19 and cvemta01  = clavet30  and firmae01<>'' and  fecpag01>='"&strDateIni&"' and  fecpag01<='"&strDateFin&"' and cveped01<>'R1' " &  permi & " order by fecpag01"

        END IF

      else
         strSQL = "select distinct refcia01 Referencia, numped01 Pedimento, adusec01 aduana, cveped01 , fecpag01 FechaPago, desf0101 Facturas, nompai19 PaisDestino, piesco40 Pies, descri30 Arribo ,nomPRO01 From  ssmtra30,sspais19,ssdage01  left join sscont40  on refcia01=refcia40 where cvepod01=cvepai19 and cvemta01  = clavet30  and firmae01<>'' and  fecpag01>='"&strDateIni&"' and  fecpag01<='"&strDateFin&"' and cveped01<>'R1' " &  permi & " order by fecpag01"
      end if

      Rsio.Source= strSQL
      Rsio.CursorType = 0
      Rsio.CursorLocation = 2
      Rsio.LockType = 1
      Rsio.Open()

      While NOT Rsio.EOF


                 strHTML = strHTML&"<tr>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Referencia").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Pedimento").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("cveped01").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("aduana").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("FechaPago").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Facturas").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("PaisDestino").Value&"</font></td>" & chr(13) & chr(10)
                 if not trim(Rsio.Fields.Item("Pies").Value) = "" then
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Pies").Value&"</font></td>" & chr(13) & chr(10)
                 else
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)
                 end if
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("Arribo").Value&"</font></td>" & chr(13) & chr(10)

                strSQL= "select REFCIA12,ordfra12,cveide12,tipoid12,numper12 from ssipar12 where refcia12='" &  Rsio.Fields.Item("Referencia").Value & "'"
                Set RsIdent = Server.CreateObject("ADODB.Recordset")
			          RsIdent.ActiveConnection =  MM_EXTRANET_STRING_TEMP
			          RsIdent.Source = strSQL
			          RsIdent.CursorType = 0
			          RsIdent.CursorLocation = 2
			          RsIdent.LockType = 1
			          RsIdent.Open()
                strIdentificadores=""
                strNumPermisos =""
			          if not RsIdent.eof then
                  While not RsIdent.eof
                   strIdentificadores = strIdentificadores  & "  " & RsIdent.Fields.Item("cveide12").Value  & "  "
                   if RsIdent.Fields.Item("tipoid12").Value = "1" then
                      strNumPermisos = strNumPermisos & " -  " & RsIdent.Fields.Item("numper12").Value  & "  "
                   end if
                   RsIdent.movenext
                 wend
               end if
               RsIdent.close
               set RsIdent = Nothing
                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strNumPermisos&"</font></td>" & chr(13) & chr(10)

                 set Rscont = server.CreateObject("ADODB.Recordset")
                 Rscont.ActiveConnection =    MM_EXTRANET_STRING_TEMP
                 strSQL = "select numcon40 from  sscont40 where refcia40 = '"&Rsio.Fields.Item("Referencia").Value&"' "
                 Rscont.Source= strSQL
                 Rscont.CursorType = 0
                 Rscont.CursorLocation = 2
                 Rscont.LockType = 1
                 strContenedores=""
                 intContenedores=0
                 Rscont.Open()
                 if not Rscont.eof then
                   While NOT Rscont.EOF
                    strContenedores=strContenedores & " " & trim(rscont.Fields.Item("numcon40").Value)
                    intContenedores=  intContenedores + 1
                    Rscont.MoveNext()
                   Wend
                   Rscont.close
                   set Rscont = nothing

                   set RsFac = server.CreateObject("ADODB.Recordset")
                 RsFac.ActiveConnection =    MM_EXTRANET_STRING_TEMP
                 strSQL = "select * from  ssfact39 where refcia39 = '"&Rsio.Fields.Item("Referencia").Value&"' "
                 RsFac.Source= strSQL
                 RsFac.CursorType = 0
                 RsFac.CursorLocation = 2
                 RsFac.LockType = 1
                 strIncoterm=""
                 RsFac.Open()
                 if not RsFac.eof then

                    strIncoterm=trim(RsFac.Fields.Item("terfac39").Value)
                  end if
                   Rsfac.close
                   set Rsfac = nothing

                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&intContenedores&"</font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strContenedores&"</font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strIncoterm&"</font></td>" & chr(13) & chr(10)
                 else
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)
                   strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
                     Rscont.close
                 set Rscont = Nothing
                 end if


                 strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&Rsio.Fields.Item("NOMPRO01").Value&"</font></td>" & chr(13) & chr(10)
                 strHTML = strHTML & "</tr>"& chr(13) & chr(10)
                 Response.Write(strHTML)
                 strHTML = ""
          Rsio.MoveNext()
      Wend

   Rsio.Close()
   Set Rsio = Nothing
   end if
 next


   strHTML = strHTML & "</table>"& chr(13) & chr(10)
   response.Write(strHTML)

 end if


%>