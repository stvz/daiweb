<%
Function FormatoFechaInv(Fecha)
'Permite darle formato a la Fecha
  tmpstrFecha = cdate(Fecha)
  tmpstrDia =cstr(datepart("d",tmpstrFecha))
  tmpstrMes = cstr(datepart("m",tmpstrFecha))
  tmpstrAnio = cstr(datepart("yyyy",tmpstrFecha))
  FormatoFechaInv = tmpstrAnio & "/" & tmpstrMes & "/" & tmpstrDia
End Function %>


<%

if request.Form("btnGenerar") = "OK" then
  Server.ScriptTimeOut=200000
  Response.Buffer = TRUE
  Response.Addheader "Content-Disposition", "attachment;"
  Response.ContentType = "application/vnd.ms-excel"
  strHTML = ""

  IPHost="10.66.1.5"
  IPHostLAR="200.67.110.202"
  strHTML= ""
  strOficina=""
  strOficina=request.Form("selectOficina")
 ' strTipoOperacion=""
 ' strTipoOperacion =request.Form("selTipo")
  strCheck1=""
  strCheck1=request.Form("buscaexacto")

  strDescMercan=""
  strDescMercan=request.Form("BuscaPor")
  strRfcClien=""
  strRfcClien=request.Form("rfcCliente")

  'strFechaIniTemp=request.Form("textFechaIni")
  'strFechaFinTemp=request.Form("textFechaFin")
  'strFechaIni=FormatoFechaInv(strFechaIniTemp)
  'strFechaFin=FormatoFechaInv(strFechaFinTemp)


    if strOficina <> "todas" then
        y=0
     else
        y=5
     end if


    for x=0 to y
     strOficina_temp = ""
      if strOficina = "todas" then
        select case x
        case 0
          strOficina_temp = "rku"
         case 1
          strOficina_temp = "ceg"
          case 2
          strOficina_temp = "sap"
          case 3
          strOficina_temp = "lzr"
          case 4
          strOficina_temp = "zgl"
          case 5
          strOficina_temp = "dai"
          end select
      else
        strOficina_temp = strOficina
      end if


     if strOficina_temp ="zgl" then
        MM_EXTRANET_STRING = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER="& IPHostLAR &"; DATABASE=zgl_extranet; UID=EXTRANET; PWD=zgl_admin; OPTION=16427"
     else
        MM_EXTRANET_STRING = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER="& IPHost &"; DATABASE="&strOficina_temp &"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
     end if

     Set RsStatus = Server.CreateObject("ADODB.Recordset")
     RsStatus.ActiveConnection = MM_EXTRANET_STRING

     if strCheck1 <> "" AND  trim(strDescMercan) <> "" then
        if strRfcClien = "" then
           strxRfC =""
        else
           strxRfC =" and rfccli01 = '" & strRfcClien&   "'  "
        end if


        if strTipoOperacion = "1" then
           strSQL="select fecpag01,d05artic.prov05,ssdagi01.cvecli01,ssdagi01.nomcli01,ssdagi01.refcia01,ssdagi01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdagi01,d05artic,ssfrac02 where ssdagi01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdagi01.firmae01)<>'' and ssfrac02.d_mer102 like '%" & strDescMercan & "%' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "' " & strxRfC &  " order by refcia01, agru05, frac05 asc"
        end if
        if strTipoOperacion = "2" then
           strSQL="select fecpag01,d05artic.prov05,ssdage01.cvecli01,ssdage01.nomcli01,ssdage01.refcia01,ssdage01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdage01,d05artic,ssfrac02 where ssdage01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdage01.firmae01)<>'' and ssfrac02.d_mer102 like '%" & strDescMercan & "%' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "' " & strxRfC &  " order by refcia01, agru05, frac05 asc"
        end if
        if strTipoOperacion = "3" then
           strSQL="select fecpag01,d05artic.prov05,ssdagi01.cvecli01,ssdagi01.nomcli01,ssdagi01.refcia01,ssdagi01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdagi01,d05artic,ssfrac02 where ssdagi01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdagi01.firmae01)<>'' and ssfrac02.d_mer102 like '%" & strDescMercan & "%' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "' " & strxRfC &  " union select d05artic.prov05,ssdage01.cvecli01,ssdage01.nomcli01,ssdage01.refcia01,ssdage01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdage01,d05artic,ssfrac02 where ssdage01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdage01.firmae01)<>'' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "'" & strxRfC
        end if


     end if

      if strCheck1 <> "" AND  trim(strDescMercan)  = "" then

        if strRfcClien = "" then
           strxRfC =""
        else
           strxRfC =" and rfccli01 = '" & strRfcClien&   "'  "
        end if


        if strTipoOperacion = "1" then
           strSQL="select fecpag01,d05artic.prov05,ssdagi01.cvecli01,ssdagi01.nomcli01,ssdagi01.refcia01,ssdagi01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdagi01,d05artic,ssfrac02 where ssdagi01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdagi01.firmae01)<>'' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "' " & strxRfC &  " order by refcia01, agru05, frac05 asc"
        end if
        if strTipoOperacion = "2" then
           strSQL="select fecpag01,d05artic.prov05,ssdage01.cvecli01,ssdage01.nomcli01,ssdage01.refcia01,ssdage01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdage01,d05artic,ssfrac02 where ssdage01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdage01.firmae01)<>'' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "' " & strxRfC &  " order by refcia01, agru05, frac05 asc"
        end if
        if strTipoOperacion = "3" then
           strSQL="select fecpag01,d05artic.prov05,ssdagi01.cvecli01,ssdagi01.nomcli01,ssdagi01.refcia01,ssdagi01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdagi01,d05artic,ssfrac02 where ssdagi01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdagi01.firmae01)<>'' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "' " & strxRfC &  " union select d05artic.prov05,ssdage01.cvecli01,ssdage01.nomcli01,ssdage01.refcia01,ssdage01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdage01,d05artic,ssfrac02 where ssdage01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdage01.firmae01)<>'' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "'" & strxRfC
        end if


     end if



     if strCheck1 = "" then

        if strTipoOperacion = "1" then
           strSQL="select fecpag01,d05artic.prov05,ssdagi01.cvecli01,ssdagi01.nomcli01,ssdagi01.refcia01,ssdagi01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdagi01,d05artic,ssfrac02 where ssdagi01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdagi01.firmae01)<>'' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "' order by refcia01, agru05, frac05 asc"
        end if
        if strTipoOperacion = "2" then
           strSQL="select fecpag01,d05artic.prov05,ssdage01.cvecli01,ssdage01.nomcli01,ssdage01.refcia01,ssdage01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdage01,d05artic,ssfrac02 where ssdage01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdage01.firmae01)<>'' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "' order by refcia01, agru05, frac05 asc"
        end if
        if strTipoOperacion = "3" then
           strSQL="select fecpag01,d05artic.prov05,ssdagi01.cvecli01,ssdagi01.nomcli01,ssdagi01.refcia01,ssdagi01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdagi01,d05artic,ssfrac02 where ssdagi01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdagi01.firmae01)<>'' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "' union select d05artic.prov05,ssdage01.cvecli01,ssdage01.nomcli01,ssdage01.refcia01,ssdage01.numped01,d05artic.cpro05,d05artic.item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,ssfrac02.fraarn02 as frac_ped,ssfrac02.ordfra02 as orden_frac,if(d05artic.frac05=ssfrac02.fraarn02,'I','D') as Dif_Frac,ssfrac02.d_mer102 as desc_ped from ssdage01,d05artic,ssfrac02 where ssdage01.refcia01=d05artic.refe05 and d05artic.refe05=ssfrac02.refcia02 and d05artic.agru05=ssfrac02.ordfra02 and trim(ssdage01.firmae01)<>'' and Fecpag01>='" & strFechaIni & "' and Fecpag01<='" & strFechaFin & "'"
        end if

     end if




     RsStatus.Source = strSQL
'Response.Write(strSQL)
'Response.End

     RsStatus.CursorType = 0
     RsStatus.CursorLocation = 2
     RsStatus.LockType = 1
     RsStatus.Open()


     if not RsStatus.eof  then


        if strTipoOperacion = "1" then
           strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Oficina  " & ucase(strOficina) &"  Tipo de operación Importación</p></font></strong>" & chr(13) & chr(10)
        end if
        if strTipoOperacion = "2" then
           strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Oficina  " & ucase(strOficina) &"  Tipo de operación Exportación</p></font></strong>"  & chr(13) & chr(10)
        end if
        if strTipoOperacion = "3" then
           strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Oficina  " & ucase(strOficina) &"  Tipo de operación Importación y Exportación</p></font></strong>" & chr(13) & chr(10)
        end if


        strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>Del " & strFechaIniTemp & " Al " & strFechaFinTemp & "</p></font></strong>"  & chr(13) & chr(10)
        strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	      strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
	      strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Clave cliente</td>" & chr(13) & chr(10)
	      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Nombre cliente</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Clave producto</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Item</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dif CvePro CveCli</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fraccion mercancia</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Orden mercancia</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descrip mercancia</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fraccion pedimento</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Orden fraccion</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Dif Fraccion</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descrip pedimento</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">CveProveedor</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripcion Proveedor</td>" & chr(13) & chr(10)
        strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FechaPago</td>" & chr(13) & chr(10)

        strHTML = strHTML & "</tr>"& chr(13) & chr(10)

        While NOT RsStatus.EOF

           strHTML = strHTML&"<tr>" & chr(13) & chr(10)
           strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("cvecli01").Value&"</font></td>" & chr(13) & chr(10)
           strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("nomcli01").Value&"</font></td>" & chr(13) & chr(10)
           strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("refcia01").Value&"</font></td>" & chr(13) & chr(10)
           strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("numped01").Value&"</font></td>" & chr(13) & chr(10)
           'Response.Write(RsStatus.Fields.Item("orden_frac").Value)

          ' Set RsProd = Server.CreateObject("ADODB.Recordset")

           'RsProd.ActiveConnection = MM_EXTRANET_STRING



          'strSQL="select cpro05,item05,if(d05artic.cpro05=d05artic.item05,'I','D') as Dif_CvePro_CveCli,d05artic.frac05 as frac_mcia,d05artic.agru05 as orden_mcia,d05artic.desc05 as Desc_mcia,if(d05artic.frac05=" & RsStatus.Fields.Item("frac_ped").Value & ",'I','D') as Dif_Frac,prov05 from d05artic where d05artic.refe05 ='" & RsStatus.Fields.Item("refcia01").Value & "'"
          'RsProd.Source = strSQL




           'RsProd.CursorType = 0
           'RsProd.CursorLocation = 2
           'RsProd.LockType = 1
           'RsProd.Open()
           'if not RsProd.eof then

             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("cpro05").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("item05").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("Dif_CvePro_CveCli").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("frac_mcia").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("orden_mcia").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("Desc_mcia").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("frac_ped").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("orden_frac").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("Dif_Frac").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("desc_ped").Value&"</font></td>" & chr(13) & chr(10)
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("prov05").Value&"</font></td>" & chr(13) & chr(10)


             Set RsProv = Server.CreateObject("ADODB.Recordset")
             RsProv.ActiveConnection = MM_EXTRANET_STRING
             strSQLProv="select * from ssprov22 where cvepro22 =" & RsStatus.Fields.Item("prov05").Value
             RsProv.Source = strSQLProv
             RsProv.CursorType = 0
             RsProv.CursorLocation = 2
             RsProv.LockType = 1
             RsProv.Open()

             if not RsProv.eof then
               strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsProv.Fields.Item("nompro22").Value & "</font></td>" & chr(13) & chr(10)
             else
               strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""></font></td>" & chr(13) & chr(10)
             end if

             RsProv.close
             set RsProv = nothing
             strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsStatus.Fields.Item("fecpag01").Value&"</font></td>" & chr(13) & chr(10)

           Response.Write(strHTML)
           strHTML = ""
           RsStatus.movenext
        wend

        strHTML = strHTML & "</tr>"& chr(13) & chr(10)
        strHTML = strHTML & "</td>"& chr(13) & chr(10)
        strHTML = strHTML & "</table>"& chr(13) & chr(10)
        response.Write(strHTML)
        strHTML=""
     else
       strHTML = strHTML & "No se encontró ningún dato para los parámetros que introdujo..." &strOficina_temp&"<br>" & chr(13) & chr(10)
       response.Write(strHTML)
       strHTML=""
     end if



    next

    strHTML = strHTML & "</tr>"& chr(13) & chr(10)
    strHTML = strHTML & "</td>"& chr(13) & chr(10)
    strHTML = strHTML & "</table>"& chr(13) & chr(10)
    response.Write(strHTML)

  end if

  RsStatus.close
  set RsStatus = nothing
%>
