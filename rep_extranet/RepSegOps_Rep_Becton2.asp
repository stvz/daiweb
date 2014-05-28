<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% Server.ScriptTimeout=1500 %>
<HTML>
<HEAD>
<TITLE>:: REPORTE DE SEGUIMIENTO DE OPERACIONES=.... ::</TITLE>
</HEAD>
<BODY>
<%
strTipoUsuario = request.Form("TipoUser")
strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
    'permi          = PermisoClientes(Session("GAduana"),strPermisos,"cliE01")
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

'strFiltroCliente = request.Form("cve")
'response.Write("VArible combo" )
'response.Write(strFiltroCliente )
'response.write("     ")
'response.Write(MM_Cod_Admon)
'response.Write("VArible permiso" )
'RESPONSE.Write(permi)
'response.End()


                        'response.write(permi)
                        'Response.End
                        'if PermisoMenu(strMenu,",03-") = "PERMITIDO" or strTipoUsuario = MM_Cod_Admon then
                        '    permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

    %>
<%''''''''''''''''''''''''if  Session("GUsuario") <> "" then
if  Session("GAduana") <> "" then %>
    <table width="778"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="center">
          <%
          dim cve,mov,fi,ff,tabla,sql1,sql2,sql3,refe

		  oficina_adu=GAduana

		   jnxadu=Session("GAduana")

       'Response.Write(jnxadu)
       'Response.End

           select case jnxadu
             case "VER"
                  strOficina="rku"
             case "MEX"
                  strOficina="dai"
             case "MAN"
                  strOficina="sap"
             case "GUA"
                  strOficina="rku"
             case "TAM"
                  strOficina="ceg"
             case "LAR"
                  strOficina="LAR"
             case "LZR"
                  strOficina="lzr"
             case "TOL"
                  strOficina="tol"
           end select

          cve=request.form("cve")
          mov=request.form("mov")
          fi=trim(request.form("fi"))
          ff=trim(request.form("ff"))
          Vrfc=Request.Form("rfcCliente")
          Vckcve=Request.Form("ckcve")
          'Vclave=Request.Form("cveCliente")
          'aduan=request.form("aduana")
          if isdate(fi) and isdate(ff) then
                '---------------------------
                  DiaI = cstr(datepart("d",fi))
                  Mesi = cstr(datepart("m",fi))
                  AnioI = cstr(datepart("yyyy",fi))
                  DateI = Anioi&"/"&Mesi&"/"&Diai

                  DiaF = cstr(datepart("d",ff))
                  MesF = cstr(datepart("m",ff))
                  AnioF = cstr(datepart("yyyy",ff))
                  DateF = AnioF&"/"&MesF&"/"&DiaF

                 if not isdate(DateI) then
                      fec="1"
                 end if
                 if not  isdate(DateF) then
                       fec="1"
                 end if
                 if fec<>"1" then
                    '---------------------------
                    'Antes de nada hay que instanciar el objeto Connection
                    'Set Conn = Server.CreateObject("ADODB.Connection")
                    'Una vez instanciado Connection lo podemos abrir y le asignamos la base de datos donde vamos a efectuar las operaciones
                    'Conn.Open "Mibase"
                    '''''''''''''
                    MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
                    Set oConn = Server.CreateObject ("ADODB.Connection")
                    Set RS = Server.CreateObject ("ADODB.RecordSet")
                    Set RS2 = Server.CreateObject ("ADODB.RecordSet")
                    Set RS3 = Server.CreateObject ("ADODB.RecordSet")
                    Set RS4 = Server.CreateObject ("ADODB.RecordSet")
                    oConn.Open MM_EXTRANET_STRING
                    '''''''''''''
                    'Response.Write(MM_EXTRANET_STRING)
                      'Response.End
                      Diferencia = DateDiff("D", DateI,DateF)
                      if Diferencia>="0" then
                      '------------
                      if request.form("tipRep") = "2" then
                         Response.Addheader "Content-Disposition", "attachment;"
                         Response.ContentType = "application/vnd.ms-excel"
                      end if
                                    if mov="i" then
                                      if Vckcve="0" then
  sql1="SELECT trim(concat(concat(concat(concat(concat(adusec01),'-'),patent01),'-'),numped01)) as Importa, " &_
  " feta01,fdoc01,frev01, fpre01, fdsp01,tipo01, refe01,clie01, nomcli01,refcia01,totbul01, fecpag01, Obser01, " &_
 "paiori02,d_mer102,numped01,if(fdoc01>feta01,if(fdoc01>frev01,fdoc01,frev01),if(feta01>frev01,feta01,frev01)) as fini, " &_
       "fecent01 FROM ssfrac02,c01refer inner join ssdagi01 on refe01= refcia01 and  "&_
       " firmae01<>'"&""&"' and rfccli01='"&Vrfc&"' and tipo01='1' and refcia02=refcia01 and " &_
       "fdsp01>='"&DateI&"' and fdsp01<='"&DateF&"' and cveped01<>'"&"R1"&"' " & " ORDER BY clie01, REFE01"
                                          '   Response.Write("impo ck=0")
                                         '    Response.Write(sql1)
                                      else
                                        'Response.Write("impo ck=1")
 sql1="SELECT trim(concat(concat(concat(concat(concat(adusec01),'-'),patent01),'-'),numped01)) as Importa, " &_
 " feta01,fdoc01,frev01, fpre01, fdsp01,tipo01, refe01,clie01, nomcli01, refcia01,totbul01, fecpag01, Obser01," &_
 " numped01,paiori02,d_mer102, if(fdoc01>feta01,if(fdoc01>frev01,fdoc01,frev01),if(feta01>frev01,feta01,frev01)) as fini, " &_
" fecent01 FROM ssfrac02, c01refer inner join ssdagi01 on refe01= refcia01 and firmae01<>'"&""&"'  " & permi & " and tipo01='1' "&_
" and fdsp01>='"&DateI&"' and refcia02=refcia01 and fdsp01<='"&DateF&"' and cveped01<>'"&"R1"&"' " & " ORDER BY clie01, REFE01 "
                                        'Response.Write(sql1)
                                      end if
                                    else
                                      if Vckcve="0" then
                                        'Response.Write("EXpo ck=0")
  sql1="SELECT trim(concat(concat(concat(concat(concat(adusec01),'-'),patent01),'-'),numped01)) as Exporta, " &_
   " feta01,fdoc01,frev01, fpre01, fdsp01,tipo01, refe01,clie01, nomcli01, refcia01,totbul01, fecpag01, Obser01, " &_
"numped01,paiori02,desf0101, frec01 as fini " & _
 "FROM ssfrac02,c01refer inner join ssdage01 on refe01= refcia01 and firmae01<>'"&""&"' and rfccli01='"&Vrfc&"' and tipo01='2' " &_
 " and refcia02=refcia01 and fdsp01>='"&DateI&"' and fdsp01<='"&DateF&"' and cveped01<>'"&"R1"&"' " & " ORDER BY clie01, REFE01"
                                        '     Response.Write(sql1)
                                      else
                                       ' Response.Write("EXpo ck=1")
  sql1="SELECT trim(concat(concat(concat(concat(concat(adusec01),'-'),patent01),'-'),numped01)) as Exporta, " &_
      " feta01,fdoc01,frev01, fpre01, fdsp01,tipo01, refe01,clie01, nomcli01, refcia01,totbul01, fecpag01, Obser01, " &_
  " numped01,paiori02,desf0101,frec01 as fini " &_
  "FROM ssfrac02,c01refer inner join ssdage01 on refe01= refcia01 and firmae01<>'"&""&"'  " & permi & " and tipo01='2'  and " &_
  " fdsp01>='"&DateI&"' and refcia02=refcia01 and fdsp01<='"&DateF&"' and cveped01<>'"&"R1"&"' " & " ORDER BY clie01, REFE01"
                                      '       Response.Write(sql1)
                                      end if

                                    end if
                                    ' Response.Write(sql1)
                                    ' Response.End
                                    'Ejecutamos la orden
                                    set RS=oConn.Execute(sql1)

                                    'Mostramos los registros
                    %>
                                    <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif">
                                    <p align="left">REPORTE DE SEGUIMIENTO DE OPERACIONES</p>
                                    </font>
                                    </strong>
                                    <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif">
                                    <p></p>
                                    </font>
                                    </strong>
                                    <%IF mov="i" THEN%>
                                      <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif">
                                      <p align="left">:: IMPORTACI&Oacute;N ::</p>
                                      </font>
                                      </strong>
                                      <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif">
                                      <p></p>
                                      </font>
                                      </strong>
                                    <%ELSE%>
                                      <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif">
                                      <p align="left">:: EXPORTACI&Oacute;N ::</p>
                                      </font>
                                      </strong>
                                      <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif">
                                      <p></p>
                                      </font>
                                      </strong>
                                    <%END IF%>
                                    <strong>
                                    <font color="#000066" size="2" face="Arial, Helvetica, sans-serif">
                                    <p align="left">FECHA INICIAL: <%=fi %>    FECHA FINAL: <%=ff %></p>
                                    </font>
                                    </strong>
                                    <table align="center"   Width="1000" bordercolor="#C1C1C1" border="2" align="center" cellpadding="0" cellspacing="0">
                                    <tr bgcolor="#006699" class="boton">
                                    <%if mov="i" then%>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Ind</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Referencia</td>
				<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Pedimento</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Clientes</td>
										 <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Ejecutivo</td>
										  <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Descrip. Mercancia</td>
										   <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Pais Origen</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Contenedores</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Bultos</td>
			          <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">ETA</td>       
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Documentos</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Revalidacion</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Previo</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">PagPedto</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Despacho</td>
										 <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Entrada</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">IndDsp</td>
								        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">IndDsp2</td>
										<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Observaciones Trafico</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">VacioCont</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Fecha C.Gastos</td>
										<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">C.Gastos</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">IndCGast</td>
									    <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Fec.Acuse Rec</td>
										<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">IndAcuse</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Observaciones Administrador</td>
                                    <%else%>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Ind</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Referencia</td>
			          <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Pedimento</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Clientes</td>
										<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Ejecutivo</td>
										  <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">No. Factura</td>
										   <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Pais Destino</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Bultos</td>

                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Documentos</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Fecha Entrada</td>

                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">PagPedto</td>

                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Despacho</td>
								        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">IndDsp</td>
										<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">VacioCont</td>
										<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Observaciones Trafico</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Fecha C.Gastos</td>
										<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">C.Gastos</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">IndCGast</td>				
										<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Fec.Acuse Rec</td>
										<td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">IndAcuse</td>	
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Observaciones Administracion</td>
										
                                    <%end if%>
                                    </tr>
                                    <%
                                    index=0
                                    Tmax=0
                                    Tmin=1000
                                    Tsum=0
                                    TMDsp=0
                                    TmaxCG=0
                                    TminCG=1000
                                    TsumCG=0
                                    TMCG=0

									Dim diaz(2)
                                    Dim masctual(2)
                                    Dim TempArray(800)
                                    Dim numetapas(50)
                                    Dim TempArray2(50)

                                    Do While Not RS.Eof
                                        index=index+1
                                    '---------convertir el formato de fechas para la realizar la resta-----
                                        if mov="i" then
                                        '***estas fechas son para el indice de despacho
					     if strOficina="sap" then
						 xDFini = RS("fini")
						 if isdate(xDFini) then
                                                   Diai = cstr(datepart("d",RS("fini")))
                                                   Mesi = cstr(datepart("m",RS("fini")))
                                                  Anioi = cstr(datepart("yyyy",RS("fini")))
                                                  DateIni = Diai & "/" &Mesi & "/"& Anioi
                                                 end if
					     end if

                                            xDFrev = RS("fini")
                                            if isdate(xDFrev) then
                                              Diai = cstr(datepart("d",RS("fini")))
                                              Mesi = cstr(datepart("m",RS("fini")))
                                              Anioi = cstr(datepart("yyyy",RS("fini")))
                                              DateRev = Diai & "/" &Mesi & "/"& Anioi
                                            end if
											
											Dateinidupont= RS("fini")
                                            if isdate(Dateinidupont) then
                                              Diai = cstr(datepart("d",RS("fini")))
                                              Mesi = cstr(datepart("m",RS("fini")))
                                              Anioi = cstr(datepart("yyyy",RS("fini")))
                                              DateRev = Diai & "/" &Mesi & "/"& Anioi
                                            end if

                                            xDFdsp=RS("fdsp01")
                                            if isdate(xDFdsp) then
                                              DiaF = cstr(datepart("d",RS("fdsp01")))
                                              MesF = cstr(datepart("m",RS("fdsp01")))
                                              AnioF = cstr(datepart("yyyy",RS("fdsp01")))
                                              DateDsp = DiaF & "/" &MesF & "/"& AnioF
                                            end if
											
                      '--------------------------MODIFICACION PARA UNILEVER FECHA ENTRADA VS DESPACHO INDICADOR2----------											
									if mov="i" then		
									 xDFrev2 = RS("fecent01")
                                            if isdate(xDFrev2) then
                                              Diai = cstr(datepart("d",RS("fecent01")))
                                              Mesi = cstr(datepart("m",RS("fecent01")))
                                              Anioi = cstr(datepart("yyyy",RS("fecent01")))
                                              DateEnt = Diai & "/" &Mesi & "/"& Anioi
                                            end if

                                            xDFdsp=RS("fdsp01")
                                            if isdate(xDFdsp) then
                                              DiaF = cstr(datepart("d",RS("fdsp01")))
                                              MesF = cstr(datepart("m",RS("fdsp01")))
                                              AnioF = cstr(datepart("yyyy",RS("fdsp01")))
                                              DateDsp = DiaF & "/" &MesF & "/"& AnioF
                                            end if
											
											TimeDsp2=DateDiff("d",DateEnt,DateDsp)
									 		
									
					 '=================----aki se saca la dif entre la fec.entrada y despacho------------------
                                            x=0
                                            t=0.0
                                            Do While (x<=TimeDsp2)
                                              x=x+1
                                              diasemana=WeekDay(DateAdd("d",x,DateEnt))
                                              if diasemana=1 then
                                                t=t+1
                                              end if
                                              if diasemana=7 then
                                                t=t+.5
                                              end if
                                            loop
                                            '----------------------
                                            TimeDsp2=0 'para asegurarno k no tenga ningun valor
                                            TimeDsp2=x-t' estos son los dias del Ind.Desp
                                            'if TimeDsp2>Tmax then
'                                              Tmax=TimeDsp2
'                                            end if
'                                            if TimeDsp2<Tmin then
'                                              Tmin=TimeDsp2
'                                            end if
'                                            Tsum=Tsum+TimeDsp2
                                        end if
									
					  '-==========================================================================================00					

	                  
                                                 TimeDsp=DateDiff("d",Dateinidupont,DateDsp)
				  				   
                                      '----aki se saca la dif entre Revalidacion y despacho------------------
                                            x=WeekDay(Dateinidupont)
											
                                            t=0.0
                                            Do While (x<=TimeDsp)
                                            
                                              diasemana=WeekDay(DateAdd("d",x,Dateinidupont))
											  
                                              if diasemana=1 then
                                                t=t+1
                                              end if
                                           '   if diasemana=7 then
'											    t=t-1
						'response.Write(x&":"&dateinidupont&",t="&t&",Referncia:"&rs("refcia01")&",timedsp:"&TimeDsp)
											'response.Write("<br>")

'    											response.Write(Dateinidupont&diasemana&"t="&t&","&WeekDay(DateAdd("d",7,Dateinidupont)))
'												response.Write("<br>")
'	     										'response.End()
'                                                
'                                              end if
											  if diasemana=7 then
                                                t=t+.5
                                              end if
                                              x=x+1
											loop
											'response.Write(x&":"&dateinidupont&",t="&t&",Referncia:"&rs("refcia01"))
'											response.Write("<br>")
'											response.Write(t)
											'response.End()
                                            '----------------------
                                            'TimeDsp=0 'para asegurarno k no tenga ningun valor
                                            'TimeDsp=x-t' estos son los dias del Ind.Desp
											TimeDsp=TimeDsp-t
                                            if TimeDsp>Tmax then
                                              Tmax=TimeDsp
                                            end if
                                            if TimeDsp<Tmin then
                                              Tmin=TimeDsp
                                            end if
                                            Tsum=Tsum+TimeDsp
                                        end if
				'----------------------------------------------------------PARA EXPO

                                       if mov<>"i" then
                                        '***estas fechas son para el indice de despacho
										  'TimeDsp=0
						xDFini = RS("fini")
						 if isdate(xDFini) then
                                                Diai = cstr(datepart("d",RS("fini")))
                                                Mesi = cstr(datepart("m",RS("fini")))
                                                Anioi = cstr(datepart("yyyy",RS("fini")))
                                                DateIni = Diai & "/" &Mesi & "/"& Anioi
                                              end if
					      
					      
					   xDFdsp=RS("fdsp01")
                                            if isdate(xDFdsp) then
                                              DiaF = cstr(datepart("d",RS("fdsp01")))
                                              MesF = cstr(datepart("m",RS("fdsp01")))
                                              AnioF = cstr(datepart("yyyy",RS("fdsp01")))
                                              DateDsp = DiaF & "/" &MesF & "/"& AnioF
                                            end if
											
											Dateinidupont= RS("fini")
                                            if isdate(Dateinidupont) then
                                              Diai = cstr(datepart("d",RS("fini")))
                                              Mesi = cstr(datepart("m",RS("fini")))
                                              Anioi = cstr(datepart("yyyy",RS("fini")))
                                              DateRev = Diai & "/" &Mesi & "/"& Anioi
                                            end if

                                               TimeDsp=DateDiff("d",Dateinidupont,DateDsp)
				  				
                                      '----aki se saca la dif entre Revalidacion y despacho------------------
                                            x=1 'WeekDay(Dateinidupont)
											
                                            t=0.0
                                            Do While (x<=TimeDsp)
                                            
                                              diasemana=WeekDay(DateAdd("d",x,Dateinidupont))
											  
                                              if diasemana=1 then
                                                t=t+1
                                              end if
                                           '   if diasemana=7 then
'											    t=t-1
'												
'    											response.Write(Dateinidupont&diasemana&"t="&t&","&WeekDay(DateAdd("d",7,Dateinidupont)))
'												response.Write("<br>")
'	     										'response.End()
'                                                
'                                              end if
											  if diasemana=7 then
                                                t=t+.5
                                              end if
                                              x=x+1
											loop
										'	response.Write(x&":"&dateinidupont&",t="&t&",Referncia:"&rs("refcia01")&",timedsp:"&TimeDsp)
											'response.Write("<br>")
'											response.Write(t)
											'response.End()
                                            '----------------------
                                            'TimeDsp=0 'para asegurarno k no tenga ningun valor
                                            'TimeDsp=x-t' estos son los dias del Ind.Desp
											TimeDsp=TimeDsp-t
                                            if TimeDsp>Tmax then
                                              Tmax=TimeDsp
                                            end if
                                            if TimeDsp<Tmin then
                                              Tmin=TimeDsp
                                            end if
                                            Tsum=Tsum+TimeDsp
                                        end if
					
					'-----------------------------------------------------------------------------------------------------------------------
		
                                    '--------------------------------------
                                    refe=RS("refcia01")
                                    obs=RS("Obser01")
									'fact=RS("desf0101")
                                    '+++aki sacamos las observaciones de las referencias*****************************************
                                    'if obs="0" or (obs>="A" and obs<="Z") then
                                          sql4 = "select c11desc from c11obser where c11clave='"&obs&"'"
                                          set RS4=oConn.Execute(sql4)
                                     '     desc=RS
                                    'else
                                    '     sql4 = "select c11desc from c11obser where c11clave='"&obs&"' "
                                    ''      set RS4=oConn.Execute(sql4)
                                    '      RS4(c11desc)="."
                                    'end if
                                    'Response.Write(sql4)
                                    'Response.End
                              '''''''''''''''
                                    sql2="select REFE01 as referencia,MARC01 as contenedor,FCARTA01 as fecha " & _
                                          "from d01conte WHERE refe01='"&refe&"'"
                                    SET RS2=oConn.Execute(sql2)
                                    IF NOT RS2.EOF THEN 'esta es en la tabla d01conte en la cual se encuentra los dos, si esta vacion, se lanza a la k sigue
                                        if not RS2.eof then
                                           xFCV = ""
                                           xCONT = ""
                                           xCount = 1
                                           while not RS2.eof
                                           if xCount=1 then
                                                xFCV =  RS2("fecha").value
                                                xCONT=  RS2("contenedor").value
                                           else
                                                xFCV  = xFCV & "," & RS2("fecha").value
                                                xCONT= xCONT & ","& RS2("contenedor").value
                                                '''''''''''''''
                                           end if
                                           xCount = xCount + 1
                                           RS2.MoveNext
                                           wend
                                       else
                                          xFCV = ""
                                          xCONT = ""
                                       end if
                                    ELSE ' aki solo busca los contenedores sin carta de vacio....lo k pondra solo sera numcon40 como contenedor
                                        sql2="select refcia40 as referencia, numcon40 as contenedor, '' as fecha " & _
                                           " from SSCONT40 WHERE refCIA40='"&refe&"'"
                                        SET RS2=oConn.Execute(sql2)
                                        if not RS2.eof then
                                           'xFCV = ""
                                           xCONT = ""
                                           xCount = 1
                                           while not RS2.eof
                                           if xCount=1 then
                                                'xFCV =  RS2("fecha").value
                                                xCONT=  RS2("contenedor").value

                                           else
                                                'xFCV  = xFCV & "," & RS2("fecha").value
                                                xCONT= xCONT & ","& RS2("contenedor").value
                                                '''''''''''''''
                                           end if
                                           xCount = xCount + 1
                                           RS2.MoveNext
                                           wend
                                        else
                                          'xFCV = ""
                                          xCONT = ""
                                        end if
                                    END IF
                              '''''''''''''''
                                    sql3="select d31refer.cgas31 as cg, fech31,frec31 "  & _
                                         "from d31refer inner join e31cgast on d31refer.cgas31=e31cgast.cgas31 and refe31='"&refe&"' AND esta31='I'"
                                         'Response.Write(sql3)
                                        ' Response.End
                                       set RS3=oConn.Execute(sql3)
                                       if not RS3.eof then
                                           TimeCG=0
										   TimeAcuse=0
                                           xFCG = ""
										   xFCG2 = ""
										   xFacuse = ""
                                           xCount=1
                                           while not RS3.EOf
                                           if xCount=1 then
                                                xFCG =  RS3("fech31").value
												xFCG2 = RS3("cg").value
												xFacuse= RS3("frec31").value
                                          '--------AKI HACE EL CALCULO CON LA CUENTA DE GASTOS POR SER LA PRIMERA fecha---------
                                               IF mov="i" or  mov<>"i" THEN
                                                    xDFcg = RS3("fech31")
                                                    if isdate(xDFcg) then
                                                      Diaf = cstr(datepart("d",RS3("fech31")))
                                                      Mesf = cstr(datepart("m",RS3("fech31")))
                                                      Aniof = cstr(datepart("yyyy",RS3("fech31")))
                                                      DateCG = Diaf & "/" &Mesf & "/"& Aniof
                                                    end if
													'-----AGREGA APARTE SOLO (fecha_cuenta_gasto-fecha acuse de recibo)-------
													 xFacuse= RS3("frec31")
													 
													 if isdate(xFacuse) then
                                                      Diaf = cstr(datepart("d",RS3("frec31")))
                                                      Mesf = cstr(datepart("m",RS3("frec31")))
                                                      Aniof = cstr(datepart("yyyy",RS3("frec31")))
                                                      DateCGAcuse = Diaf & "/" &Mesf & "/"& Aniof
                                                    end if
													  if xFacuse<>"" then
                                                      TimeAcuse=DateDiff("d",DateCG,DateCGAcuse)  
													  else
													  TimeAcuse=0
													  end if
													'--------------------------------------------------------------
                                                    TimeCG=DateDiff("d",DateDsp,DateCG)
                                                '----aki se saca la dif entre despacho y Cuenta de Gastos------------------
                                                    xx=0
                                                    tt=0
                                                    Do While (xx<TimeCG)
                                                      xx=xx+1
                                                      diasemanacg=WeekDay(DateAdd("d",xx,DateDsp))
                                                      if diasemanacg=1 then
                                                        tt=tt+1
                                                      end if
                                                      if diasemanacg=7 then
                                                        tt=tt+.5
                                                      end if
                                                    loop
                                                    '----------------------
                                                    TimeCG=xx-tt ' estos son los dias del Ind.CGastos
                                                    if TimeCG>TmaxCG then
                                                      TmaxCG=TimeCG
                                                    end if
                                                    if TimeCG<TminCG then
                                                      TminCG=TimeCG
                                                    end if
                                                    TsumCG=TsumCG+TimeCG
                                                END IF
                                           '-------hasta aki es lo del ind de cuenta de gastos----------
                                           else
                                                xFCG  = xFCG & "," & RS3("fech31").value
                                                xFCG2 = xFCG2 & "," & RS3("cg").value
												xFacuse = xFacuse & "," & RS3("frec31").value
                                           end if
                                           xCount = xCount + 1
                                           RS3.MoveNext
                                           wend
                                       else
                                           xFCG=""
										   xFCG2=""
										   xFacuse=""
                                       end if
                                     'Response.Write(sql3)
                                     'Response.End

'********************************MODIFICACION OBSERVACIONES JUEVES 2 JULIO 09***********************************


IPHost ="localhost"
iti=0
referencia=RS("refcia01")
'strOficina="rku"

'response.Write(GAduana)

MM_EXTRANET_STRING2 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="& IPHost &"; DATABASE="&strOficina &"_status; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
        Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RsObser2 = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING2
	    STRSQL="SELECT a.n_secuenc,a.n_etapa as et,a.c_referencia,a.m_observ,b.d_nombre,a.f_fecha, " &_
	       "TO_DAYS(CURDATE())-TO_DAYS(a.f_fecha) as Dtranew " &_
      "FROM etxpd a,etaps b where a.c_referencia='"&referencia&"' and a.n_etapa=b.n_etapa " &_
	  "order by a.n_etapa,a.f_fecha desc "
	    Set RsObser2= Connx.Execute(strSQL)
		'response.Write(strSQl)
		'response.End()
		Do while not RsObser2.Eof
            netapa=RsObser2("et")
	        numetapas(iti)=netapa
		    iti = iti+1
           RsObser2.MoveNext  ' de las Observaciones
         Loop ' de las Observaciones


'========elimina repetidos VECTOR si funciona probado=============================
For iy = LBound(numetapas) To UBound(numetapas)
          'Asignamos al array temporal el valor del otro array
          TempArray2(iy) = numetapas(iy)
   Next

    For x7 = 0 To UBound(numetapas)
        z = 0
        For y = 0 To UBound(numetapas)
            'Si el elemento del array es igual al array temporal
            If numetapas(x7) = TempArray2(z) And y <> x7 Then
                'Entonces Eliminamos el valor duplicado
                numetapas(y) = ""
                Nduplicado = Nduplicado + 1
            End If
            z = z + 1
        Next
    Next

   'For ipx = 0 to UBound(numetapas)
     'Response.write(numetapas(ipx)&"Posicion"&ipx&"--")
    'next
	'------------------ordenar de mayor a meno
	For ix = 0 To ubound(numetapas)
           For j = 0 To ubound(numetapas)-1
               If numetapas(j) < numetapas(j + 1) Then
                  aux = numetapas(j)
			      numetapas(j) = numetapas(j + 1)
                  numetapas(j + 1) = aux
               End If
            Next
        Next

'***********************************************************************************
observacion=""
observacionCG=""
observa2=""
observat=""

For k = 0 to UBound(numetapas)
	 if  numetapas(k)<>" " then

	 '****************modificacion 11/06/09
             ix2=1
			 MM_EXTRANET_STRING2 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="& IPHost &"; DATABASE="&strOficina &"_status;             UID=EXTRANET; PWD=rku_admin; OPTION=16427"
             Set Conx = Server.CreateObject ("ADODB.Connection")
             Set RSobs = Server.CreateObject ("ADODB.RecordSet")
             Conx.Open MM_EXTRANET_STRING2
	  SQLx="SELECT a.n_secuenc as a,a.n_etapa as b,a.c_referencia as c,a.m_observ as d1,b.d_nombre as e1,a.f_fecha as f, " &_
	   " TO_DAYS(CURDATE())-TO_DAYS(a.f_fecha) as Dtranew " &_
      " FROM etxpd a,etaps b where a.c_referencia='"&referencia&"' and a.n_etapa='"&numetapas(k)&"' and a.n_etapa=b.n_etapa " &_
	  "order by a.n_etapa,a.f_fecha desc "
	         Set RSobs= Conx.Execute(SQLx)
            'response.Write(sqlx)
            'response.end
            Do while not RSobs.Eof
	           observat= RsObs("d1")
			   if not isnull(observa) then
			     if ix2=1 then
			       observa2=observat
				 else
 			       observa2=observat+" "+observa2
				 end if
			   end if
			   ix2=ix2+1
	        RSobs.MoveNext
            Loop
			'response.Write(observa2)

            '****************************


      MM_EXTRANET_STRING2 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="& IPHost &"; DATABASE="&strOficina &"_status; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
      Set RsObser = Server.CreateObject("ADODB.Recordset")
      RsObser.ActiveConnection = MM_EXTRANET_STRING2
      STRSQL="SELECT a.n_secuenc as a,a.n_etapa as b,a.c_referencia as c,a.m_observ as d1,b.d_nombre as e1,a.f_fecha as f, " &_
	       " TO_DAYS(CURDATE())-TO_DAYS(a.f_fecha) as Dtranew " &_
      " FROM etxpd a,etaps b where a.c_referencia='"&referencia&"' and a.n_etapa='"&numetapas(k)&"' and a.n_etapa=b.n_etapa " &_
	  "order by a.n_etapa,a.f_fecha desc "
      RsObser.Source = strSQL
      RsObser.CursorType = 0
      RsObser.CursorLocation = 2
      RsObser.LockType = 1
      RsObser.Open()
        if not RsObser.eof then
            numetapa = RsObser.Fields.Item("b").value 
	        observa  = RsObser.Fields.Item("d1").value
			nombetapa= RsObser.Fields.Item("e1").value
			fechetapa= RsObser.Fields.Item("f").value
			if not isnull(fechetapa) then
			fechetapa2=Cstr(fechetapa)
			end if
			diaetapa=RsObser.Fields.Item("Dtranew").value
            if ix=1 then
			   'observacion=observacion
			   'complemento=comple
			else
                        '---modificacion Rodrigo lópez  07/07/2010 que las observaciones  de CGs vayan en la ultima columna. 
                        If numetapa <> 11 then
                            observacion=nombetapa+" "+fechetapa2+" "+observa2+" "+observacion
                        end if
                        If numetapa = 11 then
                            observacionCG =  nombetapa+" "+fechetapa2+" "+ observa2+" "+ observacionCG
                        end if
		            'observacion=observacion+","+observa+","+nombetapa
		         end if
		        ix=ix+1


         end if
       RsObser.close
       set RsObser = nothing

      end if  ' if  numetapas(k)<>" " then
	 next
	'********************************************************************FACTURAS EXPO 
	   ixf=0 
        Set ConnxF = Server.CreateObject ("ADODB.Connection")
        Set Rsfact = Server.CreateObject ("ADODB.RecordSet")
        ConnxF.Open MM_EXTRANET_STRING
	    STRSQL="select  * from ssfact39 where refcia39='"&referencia&"' " 
	    Set Rsfact= ConnxF.Execute(strSQL)
		Do while not Rsfact.Eof
            numfac=Rsfact("numfac39")
	          if ixf=0 then
			   facturas=numfac  
			else
  		       facturas=numfac+", "+facturas
		    end if
		    ixf=ixf+1

           Rsfact.MoveNext  ' de las EXPO
         Loop ' de las EXPO

	 
	 
	 
	 
	 
	 
	 
'**********************************************************************************************
                                       %>
                                    <tr>
                                    <%if mov="i" then%>


                                         <td align="center"><font size="1" face="Arial"><%Response.Write(index) %></font></td>
                                         <td align="center"><font size="1" face="Arial"><%=RS("refcia01") %></font></td>
								         <td align="center"><font size="1" face="Arial"><%=RS("importa") %></font></td>
                                         <td align="center" nowrap><font size="1" face="Arial"><%=RS("clie01") %></font></td>
										 <td align="center" nowrap><font size="1" face="Arial">&nbsp;</font></td>
										 <%if RS("d_mer102") <>"" then %>
										    <td align="center" nowrap><font size="1" face="Arial"><%=RS("d_mer102") %></font></td>
										 <%else%>
										   <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
										  <%if RS("paiori02") <>"" then %>
										    <td align="center" nowrap><font size="1" face="Arial"><%=RS("paiori02") %></font></td>
										 <%else%>
										   <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>

                                        <%if xCONT <>"" then %>
                                          <td align="center"><font size="1" face="Arial"><%=Response.Write( xCONT ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if RS("totbul01") <>"" then %>
                                            <td align="center" nowrap><font size="1" face="Arial"><%=RS("totbul01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
										 <%if RS("feta01") <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("feta01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if RS("fdoc01") <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("fdoc01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if RS("frev01") <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("frev01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if RS("fpre01") <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("fpre01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if RS("fecpag01") <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("fecpag01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if RS("fdsp01") <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("fdsp01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
										 <%if RS("fecent01") <>"" then %>
                                            <td align="center" nowrap><font size="1" face="Arial"><%=RS("fecent01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if TimeDsp <>"" then %>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial"><%=Response.Write( TimeDsp ) %></font></td>
                                         <%else%>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
										 <%if TimeDsp2 <>"" then %>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial"><%=Response.Write( TimeDsp2 ) %></font></td>
                                         <%else%>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>

										 <%if observacion <>"" then %>
                                           <td align="left" nowrap><font size="1" face="Arial"><%RESPONSE.Write(observacion)%></font></td>
                                         <%else%>
                                          <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if xFCV <>"" then %>
                                          <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCV ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">----</font></td>
                                         <%end if%>
                                         <%if xFCG <>"" then %>
                                          <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCG ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">----</font></td>
                                         <%end if%>
										 <%if xFCG <>"" then %>
                                         <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCG2 ) %></font></td>                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">----</font></td>
                                         <%end if%>
                                         <%if xFCG <>"" then %>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial"><%=Response.Write( TimeCG ) %></font></td>
                                         <%else%>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
										 <%if xFacuse <>"" then %>
                                         <td align="center" ><font size="1" face="Arial"><%=Response.Write(xFacuse)%></font></td>
                                         <%else%>
                                            <td align="center" ><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
										  <%if TimeAcuse <>"" then %>
  			                                 <td align="center" bgcolor="#0099CC"><font size="1" face="Arial"><%=Response.Write(TimeAcuse) %></font></td>
                                         <%else%>
                                          <td align="center" bgcolor="#0099CC"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
										 <%if observacionCG <>"" then %>
                                           <td align="left" nowrap><font size="1" face="Arial"><%=Response.Write(observacionCG)%></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         
							<%else%>
										 
                                         <td align="center"><font size="1" face="Arial"><%Response.Write(index) %></font></td>
                                         <td align="center"><font size="1" face="Arial"><%=RS("refcia01") %></font></td>
										 <td align="center"><font size="1" face="Arial"><%=RS("Exporta") %></font></td>
                                         <td align="center" nowrap><font size="1" face="Arial"><%=RS("clie01") %></font></td>
                                         <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
										 <%if facturas <>"" then %>
								         <td align="center" nowrap><font size="1" face="Arial"><%Response.Write(facturas)%></font></td>
										 <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
										 <%if RS("paiori02") <>"" then %>
										 <td align="center" nowrap><font size="1" face="Arial"><%=RS("paiori02") %></font></td>
										 <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if RS("totbul01") <>"" then %>
                                            <td align="center" nowrap><font size="1" face="Arial"><%=RS("totbul01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>


                                         <%if RS("fdoc01")  <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("fdoc01")  %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>


                                         <%if RS("fini") <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("fini") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>

                                       
                                         <%if RS("fecpag01") <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("fecpag01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if RS("fdsp01") <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=RS("fdsp01") %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
										 
										 
				                         <%if TimeDsp <>"" then %>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial"><%=Response.Write( TimeDsp ) %></font></td>
                                         <%else%>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
									     
										                                 
                                         <%if xFCV <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCV ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">----</font></td>
                                         <%end if%>
										 <%if observacion <>"" then %>
                                             <td align="left" nowrap><font size="1" face="Arial"><%RESPONSE.Write(observacion)%></font></td>
                                          <%else%>
                                          <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if xFCG <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCG ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">----</font></td>
                                         <%end if%>
										 <%if xFCG <>"" then %>
                                       <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCG2 ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">----</font></td>
                                         <%end if%>
                                         <%if xFCG <>"" then %>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial"><%=Response.Write( TimeCG ) %></font></td>
                                         <%else%>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
				
				                         <%if xFacuse <>"" then %>
                                            <td align="center" ><font size="1" face="Arial"><%=Response.Write(xFacuse)%></font></td>
                                         <%else%>
                                            <td align="center" ><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
				                         <%if TimeAcuse <>"" then %>
  			                                <td align="center" bgcolor="#0099CC"><font size="1" face="Arial"><%=Response.Write(TimeAcuse) %></font></td>
                                         <%else%>
                                           <td align="center" bgcolor="#0099CC"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
					 
	                                     <%if observacionCG <>"" then %>
                                           <td align="left" nowrap><font size="1" face="Arial"><%=Response.Write(observacionCG)%></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">--</font></td>
                                         <%end if%>
                           <%end if%>



                                    </tr>
                                    <%''''''''''''''
                                    RS.MoveNext
                                    Loop
                                    if Tsum <>0 then
                                      TMDsp=Tsum/index
                                    else
                                      TMDsp=0
                                    end if
                                    if Tsum <>0 then
                                      TMCG=TsumCG/index
                                    else
                                      'TMDsp=0
                                      TMCG=0
                                    end if
                                    'Cerramos el sistema de conexion
                                    oConn.Close
                                    %></table>
                                    <%if mov="i" then%>
                                            <br></br>
                                            <table align="center" bordercolor="#C1C1C1" border="2" align="center" cellpadding="0" cellspacing="0">
                                                <tr>
                                                  <td bgcolor="#006699" align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Media Tiempo Despacho:</td>
                                                  </td>
                                                  <%
                                                  TMDsp = ROUND(TMDsp,2)
                                                  TMCG=ROUND(TMCG,2)
                                                  %>
                                                  <td><%Response.Write(TMDsp)%>
                                                  </td>
                                                  <td bgcolor="#006699" align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Media Tiempo C.Gast.:</td>
                                                  </td>
                                                  <td><%Response.Write(TMCG)%>
                                                  </td>
                                                </tr>
                                                <tr>
                                                  <td bgcolor="#006699" align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Tiempo Maximo:</td>
                                                  </td>
                                                  <td><%Response.Write(Tmax)%>
                                                  </td>
                                                  <td bgcolor="#006699" align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Tiempo Maximo:</td>
                                                  </td>
                                                  <td><%Response.Write(TmaxCG)%>
                                                  </td>
                                                </tr>
                                                <tr>
                                                  <td bgcolor="#006699" align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Tiempo Minimo:</td>
                                                  </td>
                                                  <td><%Response.Write(Tmin)%>
                                                  </td>
                                                  <td bgcolor="#006699" align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Tiempo Minimo:</td>
                                                  </td>
                                                  <td><%Response.Write(TminCG)%>
                                                  </td>
                                                </tr>
                                              </table>
                                    <%
                                    'Response.Write(sql1)
                                     'Response.Write(sql2)
                                     'Response.Write(sql3)
                                    end if%>
                      <%else
                          'Response.Redirect "consulta.asp"
                      %>
                          <br><div align="center"><br></br><br></br><br></br><fieldset> :-: VERIFIQUE SU RANGO DE FECHAS :-: </fieldset></div>
                      <%
                      end if
                  else
                    %>
                     <br><div align="center"><br></br><br></br><br></br><fieldset> :+: VERIFIQUE SU RANGO DE FECHAS :+: </fieldset></div>
                    <%
                  end if
          else%>
            <br><div align="center"><br></br><br></br><br></br><fieldset> :*: VERIFIQUE SUS FECHAS :*: </fieldset></div>
        <%end if%>
        </td>
      </tr>
    </table>
    <%'else
      'strMenjError = "No tiene Autorización para visualizar este reporte"
%>
<%else
  response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>")
end if%>
<!--table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
    <tr>
    <td><%'=(strMenjError)%></td>
    </tr>
  </table-->
<%'end if %>
</BODY>
</HTML>