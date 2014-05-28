<% Server.ScriptTimeout=1500 %>

<%
'strTipoUsuario = request.Form("TipoUser")
'strPermisos = Request.Form("Permisos")
'permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
    'permi          = PermisoClientes(Session("GAduana"),strPermisos,"cliE01")
'    if not permi = "" then
'      permi = "  and (" & permi & ") "
'    end if
'    AplicaFiltro = false
'    strFiltroCliente = ""
'    strFiltroCliente = request.Form("txtCliente")
'    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
'       blnAplicaFiltro = true
'    end if
'    if blnAplicaFiltro then
'       permi = " AND cvecli01 =" & strFiltroCliente
'    end if
'    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
'       permi = ""
'    end if

%>

<%if  Session("GAduana") <> "" then%>
<%
usu="carlosmg"
pass="123456"
serv="localhost"

            Fechas=0
            ' es arcaico pero no enocntre otra forma de separar jejej XD
            espacio="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            'optiene las variable de la caratula

              VAdu1=Request.Form("Aduana1")
              VAdu2=Request.Form("Aduana2")
              VAdu3=Request.Form("Aduana3")
              VAdu4=Request.Form("Aduana4")
              VAdu5=Request.Form("Aduana5")
              VAdu6=Request.Form("Aduana6")

            'Response.Write(VAdu1)
            'Response.Write(VAdu2)
            'Response.Write(VAdu3)
            'Response.Write(VAdu4)
            'Response.Write(VAdu5)
            ' es la opcion del tipo de reporte
            'if request.form("tipRep") = 2 then
            tipRep=2
            if tipRep = 2 then
                Response.Buffer = True
               'Response.Addheader "Content-Disposition", "attachment;"
               Response.ContentType = "application/vnd.ms-excel"

               Set oSS = CreateObject("OWC10.Spreadsheet")
               Set c = oSS.Constants
            end if

            ' si es de la semana pasada o un rango de fechas.
            if Request.Form("fecha")=0 then
              STRFecha="WEEK(fecpag01,1) = "&DatePart("ww",(DateAdd("ww",-1,date())),2,1)&" and year(fecpag01)=year(now())  "
              STRFechaA="WEEK(fecpag01,1) <="&DatePart("ww",(DateAdd("ww",-1,date())),2,1)&" and year(fecpag01)=year(now()) "

              TITFECH=" REPORTE DE OPERACIONES DE LA SEMANA : "&DatePart("ww",(DateAdd("ww",-1,date())),2,1)

            else
                STRFINI=Request.Form("FINI")
                STRFFIN=Request.Form("FFIN")
                if isdate(STRFINI) and isdate(STRFFIN) AND DateDiff("d",STRFINI,STRFFIN)>=0 then
                  DiaI = cstr(datepart("d",STRFINI))
                  Mesi = cstr(datepart("m",STRFINI))
                  AnioI = cstr(datepart("yyyy",STRFINI))
                  ISTRFINI = Anioi&"/"&Mesi&"/"&Diai

                  DiaF = cstr(datepart("d",STRFFIN))
                  MesF = cstr(datepart("m",STRFFIN))
                  AnioF = cstr(datepart("yyyy",STRFFIN))
                  FSTRFFIN = AnioF&"/"&MesF&"/"&DiaF
                else
                  Response.Write("VERIFIQUE SUS FECHAS")
                  Response.End
                end if

              STRFecha="fecpag01>='"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"' "
              STRFechaA=" fecpag01<='"&FSTRFFIN&"' "
              'Response.Write(ISTRFINI)
              'Response.Write(FSTRFFIN)

              TITFECH=" REPORTE DE OPERACIONES DEL  "&FormatDateTime(ISTRFINI)&"  AL "&FormatDateTime(FSTRFFIN)

            end if

            'Response.Write(STRFecha)
            'Response.Write(STRFechaA)
            'Response.End
            'MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE="&base_datos&"; UID="&usu&"; PWD="&pass&"; OPTION=16427"

'=============================VALIDA K ESTE ACTUALIZADA BD=================================================================	
 OFI1=Request.Form("Aduana1")
 OFI2=Request.Form("Aduana2")
 OFI3=Request.Form("Aduana3")
 OFI4=Request.Form("Aduana4")
 OFI5=Request.Form("Aduana5")
 OFI6=Request.Form("Aduana6")
 
 for cj=1 to 6
  Select Case  cj
     Case "1":
	    if ofi6="rku" then
        bd="rku_extranet"
		yea=2001
		ban=1
		else
		ban=0
		end if
     Case "2":
	    if ofi2="dai" then
        bd="dai_extranet"
		yea=2001
		ban=1
		else
		ban=0
		end if
     Case "3":
	    if ofi4="sap" then
        bd="sap_extranet"
		yea=2001
		ban=1
		else
		ban=0
		end if
     Case "4":
	    if ofi1="ceg" then
        bd="ceg_extranet"
		yea=2001
		ban=1
		else
		ban=0
		end if
     Case "5":
	    if ofi3="lzr" then
        bd="lzr_extranet"
		yea=2005
		ban=1
		else
		ban=0
		end if
     Case "6":
	    if ofi5="tol" then
        bd="tol_extranet"
		yea=2008
		ban=1
		else
		ban=0
		end if
  End Select
  
  ofix= left(bd, len(bd)-9) 'solo los 3 caracteres de la izq 
if  BAN=1 THEN
   MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE="&bd&"; UID="&usu&"; PWD="&pass&"; OPTION=16427"
   Set Rst = Server.CreateObject("ADODB.Recordset")
   Rst.ActiveConnection = MM_STRING 
   squery="select count(refcia01) as opSem, cvecli01, nomcli01, patent01, fecpag01 from ssdagi01 " &_
   " where firmae01<>'' and cveped01<>'R1' and "& STRFecha&" group by patent01,cvecli01 union all " &_
   " select count(refcia01) as opSem, cvecli01, nomcli01, patent01, fecpag01 from ssdage01 " &_
   " where firmae01<>'' and cveped01<>'R1' and "& STRFecha&" group by patent01,cvecli01 order by fecpag01 desc "
   'response.Write(squery) 
   'response.End()
   Rst.Source= squery
   Rst.CursorType = 0
   Rst.CursorLocation = 2
   Rst.LockType = 1
   Rst.Open()
   if not Rst.eof then
      fechapago =  Rst.Fields.Item("fecpag01").Value
   end if
   Rst.close
   set Rst = nothing 
   'fecha=date()
   ahora = now()
   diasemana = weekday(ahora)
   diasemanapago = weekday(fechapago)
   fechaformu=Request.Form("fecha")
   if fechaformu=0 and (diasemanapago=6 or diasemanapago=7) then  'aki si es dia vie-sab y escogio el check de la semana pasada
    fpag=fechapago	
    'response.Write("La Base de Datos de: "&ofix&" esta actualizada ") 
    'response.Write("Ultima Fecha de Pago:"&fpag&"<br>") 
 else
   if fechaformu=0 and diasemanapago=5 and (ofi5="tol" or ofi3="lzr") then  'solo k sea lzr y tol y fecpago=jue ya ha pasado
    fpag=fechapago
	'response.Write("La Base de Datos de: "&ofix&" esta actualizada ") 
    'response.Write("Ultima Fecha de Pago:"&fpag&"<br>")
    else
	if fechaformu=0 then
    fpag=fechapago 
    jx=weekday(fechapago)
	youday=weekdayname(jx)
    response.Write("La Base de Datos de: "&ofix&" <strong>No</strong> esta Actualizada ") 
    response.Write("Ultima Fecha de Pago: "&youday&" "&fpag&"<br>") 
    response.End()
    end if
 end if 
end if 
'========================para checar ke este bien los datos de la carga completa==============================================
  if fechaformu<>0 then
  STRFecha1="WEEK(fecpag01,1) = "&DatePart("ww",(DateAdd("ww",-1,date())),2,1)&" and year(fecpag01)=year(now())  "
    Set Rst2 = Server.CreateObject("ADODB.Recordset")
   Rst2.ActiveConnection = MM_STRING 
   squery2=" select refcia01 as opSem, cvecli01, nomcli01, patent01, fecpag01 from ssdagi01 " &_
   " where firmae01<>'' and cveped01<>'R1' and year(fecpag01)="&yea&"  union all " &_
   " select refcia01 as opSem, cvecli01, nomcli01, patent01, fecpag01 " &_
   " from ssdage01 where firmae01<>'' and cveped01<>'R1' and  year(fecpag01)="&yea&" " &_
   " order by fecpag01 "
   'response.Write(squery2) 
   'response.End()
   Rst2.Source= squery2
   Rst2.CursorType = 0
   Rst2.CursorLocation = 2
   Rst2.LockType = 1
   Rst2.Open()
   if not Rst2.eof then
      fpago =  Rst2.Fields.Item("fecpag01").Value
   end if
   Rst2.close
   set Rst2 = nothing 
   an= year(fpago)
   if yea=an then
   'response.Write("La Base de Datos de la oficina: "&ofix&" esta <strong>Actualizada</strong> CARGA COMPLETA ") 
   'response.Write("Año:"&an&"<br>") 
   else
   response.Write("La Base de Datos de la oficina: "&ofix&" <strong>No</strong> esta actualizada CARGA COMPLETA ") 
   response.Write("Año:"&an&"<br>") 
   response.End()
   end if
   
   'ofix= left(bd, len(bd)-9)
   MMT= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=intranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
   Set Rst3 = Server.CreateObject("ADODB.Recordset")
   Rst3.ActiveConnection = MMT
   squery3="SELECT * FROM registro_monitor where ofic00='"&ofix&"' " 
   Rst3.Source= squery3
   Rst3.CursorType = 0
   Rst3.CursorLocation = 2
   Rst3.LockType = 1
   Rst3.Open()
   if not Rst3.eof then
      fmonitor =  Rst3.Fields.Item("fecha_hora_act").Value
   end if
   Rst3.close
   set Rst3 = nothing
   fmonitor2= left(fmonitor, len(fmonitor)-11)
   hoy=date()
   hoi=CStr(hoy) 'para ke se pueda comparar x los tipos de datos  
   if fmonitor2=hoi then
   'response.Write("La Base de Datos de: "&ofix&" esta actualizada ") 
   'response.Write("Ultima Actualizacion:"&fmonitor&"<br>") 
   else
   response.Write("La Base de Datos de: "&ofix&" <strong>No</strong> esta actualizada ") 
   response.Write("Ultima Actualizacion:"&fmonitor&"<br>") 
   response.End()
   end if
  end if ' if de fechaformu<>0 
 END IF	'este del IF de BAN=1
next 
'==========================================================================================================================

            ish=0
            for index= 1 to 6
			
'============================PARA BUSCAR SI SE ESTA ACTUALIZANDO ALGUNA BD=================================================	

 STROFI="'"&OFI1&"','"&OFI2&"','"&OFI3&"','"&OFI4&"','"&OFI5&"','"&OFI6&"')"
 MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=intranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
 Set RsRJ = Server.CreateObject("ADODB.Recordset")
 RsRJ.ActiveConnection = MM_STRING
'squery="SELECT count(m_bandera) as tot FROM ban_extranet where  m_bandera='NA' "
 'RsRJ.ActiveConnection = MM_STRING 
 squery="SELECT * FROM ban_extranet WHERE M_BANDERA='A' AND C_OFICINA IN ( "&STROFI& " "
'response.Write(squery)
'response.End()
 RsRJ.Source= squery
 RsRJ.CursorType = 0
 RsRJ.CursorLocation = 2
 RsRJ.LockType = 1
 RsRJ.Open()
 if not RsRJ.eof then
      OFICINA_ACT =  RsRJ.Fields.Item("C_OFICINA").Value
 end if
 RsRJ.close
 set RsRJ = nothing 
'response.Write(OFICINA_ACT)
'response.End()
 if OFICINA_ACT="CEG" or OFICINA_ACT="DAI" or OFICINA_ACT="LZR" or OFICINA_ACT="SAP" or OFICINA_ACT="TOL" or OFICINA_ACT="RKU" then
  response.Write("Se esta actualizando la Base de Datos de la oficina: "&OFICINA_ACT&" espere un momento e intente de nuevo")
  response.end 
 else
'==================================================
            ir=0
                 strOficina= ""
                 MM_STRING  = ""
                 ok="NOok"
                 ii=0
            'Response.Write(index)
                 IF index=1  and Vadu1 = "ceg" THEN
                    MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=ceg_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
                    Vofi="CEG - COMERCIO EXTERIOR DEL GOLFO "
                    ok="ok"
                    'Response.Write(index)
                    ish=ish+1
                    ir=1
                    nombhoja="oCEGSheet"
                    'Set oOrdersSheet = oSS.Worksheets(ish)
                    Set nombhoja = oSS.Worksheets(ish)
                    nombhoja.Name = "CEG"
                    'oOrdersSheet.Name = "CEG"
                    'arma = arma & arma(Query,MM_STRING,Vrfc,Vfracc,Vdescrip,Vcod,Vofi)
                 END IF
                 IF index=2 and Vadu2 = "dai" THEN
                    MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=dai_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
                    Vofi="DAI - DESPACHOS AEREOS INTEGRADOS"
                    ok="ok"
                    'Response.Write(index)
                    ish=ish+1
                    ir=1
                    nombhoja="oDAISheet"
                    Set nombhoja = oSS.Worksheets(ish)
                    nombhoja.Name = "DAI"
                    'arma = arma & arma(Query,MM_STRING,Vrfc,Vfracc,Vdescrip,Vcod,Vofi)
                 END IF
                 IF index=3 and Vadu3 = "lzr" THEN
                    MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=lzr_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
                    Vofi="LAZARO"
                    ok="ok"
                    'Response.Write(index)
                    ish=ish+1
                    ir=1
                    nombhoja="oLZRSheet"
                    Set nombhoja = oSS.Worksheets(ish)
                    nombhoja.Name = "LZR"
                    'arma= arma(Query,MM_STRING,Vrfc,Vfracc,Vdescrip,Vcod,Vofi)
                 END IF
                 IF index=4 and Vadu4 = "sap" THEN
                    MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=sap_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
                    Vofi="SERVICIOS ADUANALES DEL PACIFICO"
                    ok="ok"
                    'Response.Write(index)
                    ish=ish+1
                    ir=1
                    nombhoja="oSAPSheet"
                    'Set nombhoja = oSS.Worksheets(ish)
                    Set nombhoja = oSS.Worksheets.add()
                    nombhoja.Name = "SAP"
                    'arma = arma & arma(Query,MM_STRING,Vrfc,Vfracc,Vdescrip,Vcod,Vofi)
                 END IF
                 IF index=5 and Vadu5 = "tol" THEN
                    MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=tol_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
                    Vofi="TOLUCA "
                    ok="ok"
                    'Response.Write(index)
                    ish=ish+1
                    ir=1
                    nombhoja="oRKUSheet"
                    'Set nombhoja = oSS.Worksheets(ish)
                    Set nombhoja = oSS.Worksheets.add()
                    nombhoja.Name = "TOL"
                    'arma = arma & arma(Query,MM_STRING,Vrfc,Vfracc,Vdescrip,Vcod,Vofi)
                 END IF
                 IF index=6 and Vadu6 = "rku" THEN
                    MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=rku_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
                    Vofi="RKU - GRUPO REYES KURI "
                    ok="ok"
                    'Response.Write(index)
                    ish=ish+1
                    ir=1
                    nombhoja="oRKUSheet"
                    'Set nombhoja = oSS.Worksheets(ish)
                    Set nombhoja = oSS.Worksheets.add()
                    nombhoja.Name = "RKU"
                    'arma = arma & arma(Query,MM_STRING,Vrfc,Vfracc,Vdescrip,Vcod,Vofi)
                 END IF


                 '''''''

                if ok="ok" then
                      ' inicia el proceso
                      Esum3210=0
                      Esum3921=0
                      Esum3931=0
                      Esum3933=0
                      Isum3210=0
                      Isum3921=0
                      Isum3931=0
                      Isum3933=0
					  Esum3945=0
                      Isum3945=0

                      fmax="01/01/1900"
                      fmin="31-12-3000"

                      for it=1 to 2

                      sumopSem=0
                      sumpvAdu=0
                      sumpopAnu=0
                      sumpACvAdu=0

                        if it=1 then
                          titulo=" EXPORTACION"
                          Tabla="ssdage01"

                          'Set oRange = oOrdersSheet.Range("A1:G1")
                          Set oRange = nombhoja.Range("A1:G1")
                          oRange.Value = Array(" ", "",TITFECH , "", "", "","")
                          oRange.Font.Bold = true
                          oRange.Font.color = "#FFFFFF"
                          oRange.Interior.Color = "#336699"
                          'oRange.Borders(c.xlEdgeBottom).Weight = c.xlThick
                          oRange.HorizontalAlignment = c.xlHAlignCenter

                          'Set oRange = oOrdersSheet.Range("A"&ir&":G"&ir)
                          ir=ir+1
                          Set oRange = nombhoja.Range("A"&ir&":G"&ir)
                          oRange.Value = Array(" ", "",titulo ,"SEM.", "SEM.", "ACUM.","ACUM.")
                          oRange.Font.Bold = true
                          oRange.Font.color = "#FFFFFF"
                          oRange.Interior.Color = "#336699"
                          'oRange.Borders(c.xlEdgeBottom).Weight = c.xlThick
                          oRange.HorizontalAlignment = c.xlHAlignCenter

                          'Set oRange = oOrdersSheet.Range("A"&ir&":G"&ir)
                          ir=ir+1
                          Set oRange = nombhoja.Range("A"&ir&":G"&ir)
                          oRange.Value = Array(" PATENTE", "CVE.CLI", "NOMBRE DEL CLIENTE", "No.OP'S", "V.ADU", "NoOP'S","V.ADU")
                          oRange.Font.Bold = true
                          oRange.Font.color = "#FFFFFF"
                          oRange.Interior.Color = "#336699"
                          'oRange.Borders(c.xlEdgeBottom).Weight = c.xlThick
                          oRange.HorizontalAlignment = c.xlHAlignCenter

                          'Apply formatting to the columns  ANCHO DE COLUMNAS
                          nombhoja.Range("A:B").ColumnWidth = 10
                          nombhoja.Range("C:C").ColumnWidth = 40
                          nombhoja.Range("D:G").ColumnWidth = 10
                        else
                          ir=ir+3
                          titulo=" IMPORTACION"
                          Tabla="ssdagi01"

                          'Set oRange = oOrdersSheet.Range("A1:G1")
                          Set oRange = nombhoja.Range("A"&ir&":G"&ir)
                          oRange.Value = Array(" ", "",TITFECH , "", "", "","")
                          oRange.Font.Bold = true
                          oRange.Font.color = "#FFFFFF"
                          oRange.Interior.Color = "#336699"
                          'oRange.Borders(c.xlEdgeBottom).Weight = c.xlThick
                          oRange.HorizontalAlignment = c.xlHAlignCenter

                          'Set oRange = oOrdersSheet.Range("A"&ir&":G"&ir)
                          ir=ir+1
                          Set oRange = nombhoja.Range("A"&ir&":G"&ir)
                          oRange.Value = Array(" ", "",titulo , "SEM.", "SEM.", "ACUM.","ACUM.")
                          oRange.Font.Bold = true
                          oRange.Font.color = "#FFFFFF"
                          oRange.Interior.Color = "#336699"
                          'oRange.Borders(c.xlEdgeBottom).Weight = c.xlThick
                          oRange.HorizontalAlignment = c.xlHAlignCenter

                          'Set oRange = oOrdersSheet.Range("A"&ir&":G"&ir)
                          ir=ir+1
                          Set oRange = nombhoja.Range("A"&ir&":G"&ir)
                          oRange.Value = Array(" PATENTE", "CVE.CLI", "NOMBRE DEL CLIENTE", "No.OP'S", "V.ADU", "NoOP'S","V.ADU")
                          oRange.Font.Bold = true
                          oRange.Font.color = "#FFFFFF"
                          oRange.Interior.Color = "#336699"
                          'oRange.Borders(c.xlEdgeBottom).Weight = c.xlThick
                          oRange.HorizontalAlignment = c.xlHAlignCenter

                          'Apply formatting to the columns  ANCHO DE COLUMNAS
                          'nombhoja.Range("A:B").ColumnWidth = 10
                          'nombhoja.Range("C:C").ColumnWidth = 40
                          'nombhoja.Range("D:G").ColumnWidth = 10
                        end if
                      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            sqlCVEPAT="select count(refcia01) as opSem, cvecli01, nomcli01, patent01, fecpag01 "&_
                                      "from "&tabla&" "&_
                                      "where firmae01<>'' and cveped01<>'R1' and "&STRFecha&"  "&_
                                      "group by patent01,cvecli01 "&_
                                      "order by patent01,opSem desc "
                      'Response.Write(sqlCVEPAT)
                      'Response.End
                             set RSCVEPAT = server.CreateObject("ADODB.Recordset")
                             RSCVEPAT.ActiveConnection = MM_STRING
                             RSCVEPAT.Source= sqlCVEPAT
                             RSCVEPAT.CursorType = 0
                             RSCVEPAT.CursorLocation = 2
                             RSCVEPAT.LockType = 1
                             RSCVEPAT.Open()

                        while not RSCVEPAT.eof
                            pvAdu=0
                            popAnu=0
                            pACvAdu=0

                        'strPatente=RSCVEPAT.fields.item("patent01").value
                        'strCvecli=RSCVEPAT.fields.item("cvecli01").value

                        ' aki sumo los valores total para cada patente dependiendo del tipo de OPERACION
                          if it= 1 then
                              Select Case (RSCVEPAT.fields.item("patent01").value)
                                 Case "3210":
                                     Esum3210=Esum3210+RSCVEPAT.fields.item("opSem").value
                                 Case "3921":
                                     Esum3921=Esum3921+RSCVEPAT.fields.item("opSem").value
                                 Case "3931":
                                     Esum3931=Esum3931+RSCVEPAT.fields.item("opSem").value
                                 Case "3933":
                                     Esum3933=Esum3933+RSCVEPAT.fields.item("opSem").value
   							     Case "3945":
                                     Esum3945=Esum3945+RSCVEPAT.fields.item("opSem").value
                              End Select
                           else
                              Select Case (RSCVEPAT.fields.item("patent01").value)
                                 Case "3210":
                                     Isum3210=Isum3210+RSCVEPAT.fields.item("opSem").value
                                 Case "3921":
                                     Isum3921=Isum3921+RSCVEPAT.fields.item("opSem").value
                                 Case "3931":
                                     Isum3931=Isum3931+RSCVEPAT.fields.item("opSem").value
                                 Case "3933":
                                     Isum3933=Isum3933+RSCVEPAT.fields.item("opSem").value
								 Case "3945":
                                     Isum3945=Isum3945+RSCVEPAT.fields.item("opSem").value
                              End Select
                           end if


                          'para obtener las  Fechas inicial y final
                          if RSCVEPAT.fields.item("fecpag01").value > fmax then
                              fmax=RSCVEPAT.fields.item("fecpag01").value
                          end if
                          if RSCVEPAT.fields.item("fecpag01").value < fmin then
                              fmin=RSCVEPAT.fields.item("fecpag01").value
                          end if
                          'end if


                        sqlVaduSem= " select cvecli01,nomcli01, sum(vaduan02) as vAdu  "&_
                                    "from "&tabla&" join ssfrac02 on refcia01=refcia02  "&_
                                    "where firmae01<>'' and cveped01<>'R1' and  "&_
                                    ""&STRFecha&" and  "&_
                                    "patent01='"&RSCVEPAT.fields.item("patent01").value&"' and  "&_
                                    "cvecli01='"&RSCVEPAT.fields.item("cvecli01").value&"'  "&_
                                    "group by patent01,cvecli01  "&_
                                    "order by patent01,cvecli01"
                      'Response.Write(sqlVaduSem)
                      'Response.End
                        sqlNOpAnual="select count(refcia01) as opAnu, cvecli01, nomcli01, patent01 "&_
                                    "from "&tabla&" "&_
                                    "where firmae01<>'' and cveped01<>'R1' and "&_
                                    "fecpag01>='"&YEAR(STRFFIN)&"-01-01' and "&STRFechaA&" and  "&_
                                    "patent01='"&RSCVEPAT.fields.item("patent01").value&"' and  "&_
                                    "cvecli01='"&RSCVEPAT.fields.item("cvecli01").value&"'  "&_
                                    "group by patent01,cvecli01 "&_
                                    "order by patent01,cvecli01 "

                      'Response.Write(sqlNOpAnual)
                        sqlVaduAnual= " select cvecli01,nomcli01, sum(vaduan02) as ACvAdu "&_
                                    "from "&tabla&" join ssfrac02 on refcia01=refcia02 "&_
                                    "where firmae01<>'' and cveped01<>'R1' and "&_
                                    "fecpag01>='"&YEAR(STRFFIN)&"-01-01' and "&STRFechaA&" and "&_
                                    "patent01='"&RSCVEPAT.fields.item("patent01").value&"' and "&_
                                    "cvecli01='"&RSCVEPAT.fields.item("cvecli01").value&"' "&_
                                    "group by patent01,cvecli01 "&_
                                    "order by patent01,cvecli01"
                      'Response.Write(sqlVaduAnual)

                      '********************************************************************************************
                        'operaciones anuales de otros clientes
                        sqlOpAnualOtros= " select count(refcia01) as opAnu, cvecli01, nomcli01, patent01 "&_
                                         " from "&tabla&" "&_
                                         " where firmae01<>'' and cveped01<>'R1' and "&_
                                         "       fecpag01>='"&YEAR(STRFFIN)&"-01-01' and "&STRFechaA&" "&_
                                         " group by patent01,cvecli01 "&_
                                         " order by patent01,cvecli01 "
                      'Response.Write(sqlOpAnualOtros)
                        'Valor aduana anuales de otros clientes
                        sqlVaduAnualOtros= " select cvecli01,nomcli01, sum(vaduan02) as ACvAdu "&_
                                           " from "&tabla&" join ssfrac02 on refcia01=refcia02 "&_
                                           " where firmae01<>'' and cveped01<>'R1' and "&_
                                           "       fecpag01>='"&YEAR(STRFFIN)&"-01-01' and "&STRFechaA&" "&_
                                           " group by patent01,cvecli01 "&_
                                           " order by patent01,cvecli01"
                      'Response.Write(sqlVaduAnualOtros)
                      '********************************************************************************************

                      ' valor en aduana a la semana
                             set RSVaduSem = server.CreateObject("ADODB.Recordset")
                             RSVaduSem.ActiveConnection = MM_STRING
                             RSVaduSem.Source= sqlVaduSem
                             RSVaduSem.CursorType = 0
                             RSVaduSem.CursorLocation = 2
                             RSVaduSem.LockType = 1
                             RSVaduSem.Open()

                                if not RSVaduSem.eof then
                                    pvAdu= RSVaduSem.fields.item("vAdu").value
                                else
                                    pvAdu=0
                                end if

                             RSVaduSem.close
                             set RSVaduSem = nothing

                      ' numero de operacion al año
                             set RSNOpAnual = server.CreateObject("ADODB.Recordset")
                             RSNOpAnual.ActiveConnection = MM_STRING
                             RSNOpAnual.Source= sqlNOpAnual
                             RSNOpAnual.CursorType = 0
                             RSNOpAnual.CursorLocation = 2
                             RSNOpAnual.LockType = 1
                             RSNOpAnual.Open()

                                if not RSNOpAnual.eof then
                                    popAnu= RSNOpAnual.fields.item("opAnu").value
                                else
                                    popAnu=0
                                end if

                             RSNOpAnual.close
                             set RSNOpAnual = nothing

                      ' valor el aduana anual
                             set RSVaduAnual = server.CreateObject("ADODB.Recordset")
                             RSVaduAnual.ActiveConnection = MM_STRING
                             RSVaduAnual.Source= sqlVaduAnual
                             RSVaduAnual.CursorType = 0
                             RSVaduAnual.CursorLocation = 2
                             RSVaduAnual.LockType = 1
                             RSVaduAnual.Open()

                                if not RSVaduAnual.eof then
                                    pACvAdu= RSVaduAnual.fields.item("ACvAdu").value
                                else
                                    pACvAdu=0
                                end if

                             RSVaduAnual.close
                             set RSVaduAnual = nothing


                      sumopSem=sumopSem+RSCVEPAT.fields.item("opSem").value
                      sumpvAdu=sumpvAdu+pvAdu
                      sumpopAnu=sumpopAnu+popAnu
                      sumpACvAdu=sumpACvAdu+pACvAdu

                      '''''''''''''''''''''''''''''''''''''''''''
                      '''''''''' AGREGA LOS VALORES A LA CELDA
                      '''''''''''''''''''''''''''''''''''''''''''
                          ir=ir+1   ' este es el indice de la fila

                                  nombhoja.Range("A"&ir).Value = RSCVEPAT.fields.item("patent01").value
                                  nombhoja.Range("B"&ir).Value = RSCVEPAT.fields.item("cvecli01").value
                                  nombhoja.Range("C"&ir).Value = RSCVEPAT.fields.item("nomcli01").value
                                  nombhoja.Range("D"&ir).Value = RSCVEPAT.fields.item("opSem").value
                                  nombhoja.Range("E"&ir).Value = pvAdu
                                  nombhoja.Range("F"&ir).Value = popAnu
                                  nombhoja.Range("G"&ir).Value = pACvAdu



                          RSCVEPAT.movenext()
                        wend

                        RSCVEPAT.close
                        set RSCVEPAT = nothing
                      'IEsum3210=IEsum3210+sum3210
                      'IEsum3921=IEsum3921+sum3921
                      'IEsum3931=IEsum3931+sum3931

                        '********************************************************************************************
                        '***** Antes de poner el total ponemos el acomulado anual de los clientes que no han tenido operaciones esta semana
                        '********************************************************************************************
                        'operaciones anuales de otros clientes
                        sqlOpAnualOtros= " select count(refcia01) as opAnu "&_
                                         " from "&tabla&" "&_
                                         " where firmae01<>'' and cveped01<>'R1' and "&_
                                         "       fecpag01>='"&YEAR(STRFFIN)&"-01-01' and "&STRFechaA&" "&_
                                         " order by patent01,cvecli01 "
                        'Response.Write(sqlOpAnualOtros)

                        ' numero de operacion al año
                        popAnuOtros = ""
                        set RSNOpAnualOtros = server.CreateObject("ADODB.Recordset")
                        RSNOpAnualOtros.ActiveConnection = MM_STRING
                        RSNOpAnualOtros.Source= sqlOpAnualOtros
                        RSNOpAnualOtros.CursorType = 0
                        RSNOpAnualOtros.CursorLocation = 2
                        RSNOpAnualOtros.LockType = 1
                        RSNOpAnualOtros.Open()

                        if not RSNOpAnualOtros.eof then
                            'while not RSNOpAnualOtros.eof
                               popAnuOtros = RSNOpAnualOtros.fields.item("opAnu").value
                            '   RSNOpAnualOtros.movenext
                            'wend
                        else
                           popAnuOtros = 0
                        end if

                        RSNOpAnualOtros.close
                        set RSNOpAnualOtros = nothing

                        'Response.Write(popAnuOtros)
                        'Response.Write(sqlOpAnualOtros)
                        'Response.End

                        'Valor aduana anuales de otros clientes
                        sqlVaduAnualOtros= " select cvecli01,nomcli01, sum(vaduan02) as ACvAdu "&_
                                           " from "&tabla&" join ssfrac02 on refcia01=refcia02 "&_
                                           " where firmae01<>'' and cveped01<>'R1' and "&_
                                           "       fecpag01>='"&YEAR(STRFFIN)&"-01-01' and "&STRFechaA&" "&_
                                           " group by patent01,cvecli01 "&_
                                           " order by patent01,cvecli01"
                        'Response.Write(sqlVaduAnualOtros)
                        '********************************************************************************************

                        ir=ir+1   ' este es el indice de la fila
                        ' imprime los totales de no de operaciones y valor en aduna.
                                  nombhoja.Range("A"&ir).Value = ""
                                  nombhoja.Range("B"&ir).Value = ""
                                  nombhoja.Range("C"&ir).Value = "OTROS"
                                  nombhoja.Range("D"&ir).Value = "0"
                                  nombhoja.Range("E"&ir).Value = "0"
                                  'nombhoja.Range("F"&ir).Value = popAnuOtros
                                  nombhoja.Range("F"&ir).Value = ""
                                  nombhoja.Range("G"&ir).Value = "0"
                        '********************************************************************************************


                        ir=ir+1   ' este es el indice de la fila
                        ' imprime los totales de no de operaciones y valor en aduna.
                                  nombhoja.Range("A"&ir).Value = ""
                                  nombhoja.Range("B"&ir).Value = "TOTALES"
                                  nombhoja.Range("C"&ir).Value = ":"
                                  nombhoja.Range("D"&ir).Value = sumopSem
                                  nombhoja.Range("E"&ir).Value = sumpvAdu
                                  nombhoja.Range("F"&ir).Value = sumpopAnu
                                  nombhoja.Range("G"&ir).Value = sumpACvAdu



                        next

                        ' imprime los totales por patente..........
                        ir=ir+3   ' este es el indice de la fila

                          Set oRange = nombhoja.Range("d"&ir&":G"&ir)
                          oRange.Value = Array( "PATENTE", "EXPORT", "IMPORT","EXP+IMP")
                          oRange.Font.Bold = true
                          oRange.Font.color = "#FFFFFF"
                          oRange.Interior.Color = "#336699"
                          'oRange.Borders(c.xlEdgeBottom).Weight = c.xlThick
                          oRange.HorizontalAlignment = c.xlHAlignCenter

                                  'nombhoja.Range("A"&ir).Value = "..."
                                  'nombhoja.Range("B"&ir).Value = "..."
                                  'nombhoja.Range("C"&ir).Value = "..."
                                  'nombhoja.Range("D"&ir).Value = "PATENTE"
                                  'nombhoja.Range("E"&ir).Value = "EXPORTACION"
                                  'nombhoja.Range("F"&ir).Value = "IMPORTACION"
                                  'nombhoja.Range("G"&ir).Value = "EXP + IMP"

                        ir=ir+1   ' este es el indice de la fila
                                  nombhoja.Range("A"&ir).Value = ""
                                  nombhoja.Range("B"&ir).Value = ""
                                  nombhoja.Range("C"&ir).Value = ""
                                  nombhoja.Range("D"&ir).Value = "3210"
                                  nombhoja.Range("E"&ir).Value = Esum3210
                                  nombhoja.Range("F"&ir).Value = Isum3210
                                  nombhoja.Range("G"&ir).Value = Esum3210+Isum3210
                        ir=ir+1   ' este es el indice de la fila
                                  nombhoja.Range("A"&ir).Value = ""
                                  nombhoja.Range("B"&ir).Value = ""
                                  nombhoja.Range("C"&ir).Value = ""
                                  nombhoja.Range("D"&ir).Value = "3921"
                                  nombhoja.Range("E"&ir).Value = Esum3921
                                  nombhoja.Range("F"&ir).Value = Isum3921
                                  nombhoja.Range("G"&ir).Value = Esum3921+Isum3921
                        ir=ir+1   ' este es el indice de la fila
                                  nombhoja.Range("A"&ir).Value = ""
                                  nombhoja.Range("B"&ir).Value = ""
                                  nombhoja.Range("C"&ir).Value = ""
                                  nombhoja.Range("D"&ir).Value = "3931"
                                  nombhoja.Range("E"&ir).Value = Esum3931
                                  nombhoja.Range("F"&ir).Value = Isum3931
                                  nombhoja.Range("G"&ir).Value = Esum3931+Isum3931
                        ir=ir+1   ' este es el indice de la fila
                                  nombhoja.Range("A"&ir).Value = ""
                                  nombhoja.Range("B"&ir).Value = ""
                                  nombhoja.Range("C"&ir).Value = ""
                                  nombhoja.Range("D"&ir).Value = "3933"
                                  nombhoja.Range("E"&ir).Value = Esum3933
                                  nombhoja.Range("F"&ir).Value = Isum3933
                                  nombhoja.Range("G"&ir).Value = Esum3933+Isum3933
					    ir=ir+1   ' este es el indice de la fila
                                  nombhoja.Range("A"&ir).Value = ""
                                  nombhoja.Range("B"&ir).Value = ""
                                  nombhoja.Range("C"&ir).Value = ""
                                  nombhoja.Range("D"&ir).Value = "3945"
                                  nombhoja.Range("E"&ir).Value = Esum3945
                                  nombhoja.Range("F"&ir).Value = Isum3945
                                  nombhoja.Range("G"&ir).Value = Esum3945+Isum3945			  

              end if
              if ok="ok" then
                    nombhoja.Activate   'Makes the Orders sheet active
                    oSS.Windows(1).ViewableRange = nombhoja.UsedRange.Address
              end if
     		 end if 'termina el if de para checar bandera de cargas actualizandose
            next

          Response.Write oSS.XMLData
          Response.End
%>
<%else
  response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>")
end if%>