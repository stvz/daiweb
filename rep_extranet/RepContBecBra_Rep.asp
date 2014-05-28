<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->


<HTML>
<HEAD>
<TITLE>:: REPORTE DE CONTENEDORES EXPORTACION ::</TITLE>
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

if  Session("GAduana") <> "" then
'Response.Write(Session("GAduana"))
'Response.End%>
<table width="778"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">
      <%
  	  'cve=request.form("cve")
      fi=trim(request.form("fi"))
      ff=trim(request.form("ff"))
	    tipRep=request.form("tipRep")
      'rfc="BME8307014Z0"
      'fi="2007-09-01"
      'ff="2007-09-14"
      'tipRep="2"
    'Response.Write(rfc)
    'Response.Write(fi)
    'Response.Write(ff)
    'response.Write(tipRep)
    'Response.end
      if isdate(fi) and isdate(ff) then
            '---------------------------
              DiaI = cstr(datepart("d",fi))
              Mesi = cstr(datepart("m",fi))
              AnioI = cstr(datepart("yyyy",fi))
              DateI = Anioi & "/" &Mesi & "/"& Diai

              DiaF = cstr(datepart("d",ff))
              MesF = cstr(datepart("m",ff))
              AnioF = cstr(datepart("yyyy",ff))
              DateF = AnioF & "/" &MesF & "/"& DiaF

             if not isdate(DateI) then
                  fec="1"
             end if
             if not  isdate(DateF) then
                   fec="1"
             end if
             if fec<>"1" then
                '---------------------------
                MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
                Set oConn = Server.CreateObject ("ADODB.Connection")
                'set RS0=server.CreateObject("ADODB.RecordSet")
                Set RS = Server.CreateObject ("ADODB.RecordSet")
                Set RS2=Server.CreateObject ("ADODB.RecordSet")
                Set RS3=Server.CreateObject ("ADODB.RecordSet")
                Set RS4=Server.CreateObject ("ADODB.RecordSet")
                'oConn.Open "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE=_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
                oConn.Open MM_EXTRANET_STRING
                'oConn.Open "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE=rku_extranet; UID=carlosmg; PWD=123456; OPTION=16427"
                '''''''''''''
                'Response.Write(MM_EXTRANET_STRING)
                'Response.End

                diferencia=datediff("d",DateI,DateF)
                  if diferencia>="0" then
                  '------------
                  'if request.form("tipRep") = "2" then
                  if tipRep="2" then
                     Response.Addheader "Content-Disposition", "attachment;"
                     Response.ContentType = "application/vnd.ms-excel"
                  end if
                  '-----------
                  'sql0="SELECT rfccli18 from ssclie18, ssdage01 where cvecli18<>'' " &permi&  " "
                  'set RS0=oConn.Execute(sql0)
                                sql1="SELECT refcia01,cvepod01,nombar01,numped01,fecpag01,cveadu01,cvecli01,firmae01,nomcli01,rfccli01, refcia40,numcon40 "&_
                                    "FROM ssdage01 left join sscont40 on refcia40=refcia01 "&_
                                    "where fecpag01>='"&DateI&"' and fecpag01<='"&DateF&"' " &permi& " and CvePed01<>'R1' and firmae01<>'' order by refcia01"
'Response.Write(sql1)
'Response.End
                                set RS=oConn.Execute(sql1)
                                %>
                                <%if not RS.EOF then%>
                                <font color="#336699" size="3" class="boton">
                                <p align="left"><b>&nbsp;Cliente:&nbsp;<%=RS.Fields("cvecli01").value%>&nbsp;&nbsp;&nbsp;RFC:&nbsp;<%=RS.Fields("rfccli01").value%>&nbsp;&nbsp; FECHA INICIAL:&nbsp;<%=fi %>&nbsp;&nbsp;&nbsp; FECHA FINAL:&nbsp;<%=ff %></b></p>
                                </font>
                                <%else%>
                                <font color="#336699" size="3" class="boton">
                                <p align="left"><b>&nbsp;NO HAY MOVIMIENTOS ENTRE ESTAS FECHAS &nbsp;&nbsp; FECHA INICIAL:&nbsp;<%=fi %>&nbsp;&nbsp;&nbsp; FECHA FINAL:&nbsp;<%=ff %></b></p>
                                </font>
                                <%end if%>
                                <table align="center"  border="1" Width="1000"cellpadding="0" cellspacing="0" >
                                <tr bgcolor="#336699" class="boton">
                                    <th ><font color="#FFFFFF" size="2">i</font></th>
                                    <th ><font color="#FFFFFF" size="2">Referencia</font></th>
                                    <th ><font color="#FFFFFF" size="2">Destino</font></th>
                                    <th ><font color="#FFFFFF" size="2"> Buque </font></th>
                                    <th ><font color="#FFFFFF" size="2">Contenedor</font></th>
                                    <th ><font color="#FFFFFF" size="2">FllegadaCont</font></th>
                                    <th ><font color="#FFFFFF" size="2">Booking</font></th>
                                    <th ><font color="#FFFFFF" size="2">FRealDesp</font></th>
                                    <th ><font color="#FFFFFF" size="2">Naviera</font></th>
                                    <th ><font color="#FFFFFF" size="2">Pedimento</font></th>
                                    <th ><font color="#FFFFFF" size="2">FEntregaPediNavi</font></th>
                                    <th ><font color="#FFFFFF" size="2">Aduana</font></th>
                                    <th ><font color="#FFFFFF" size="2">obs</font></th>

                                </tr>
                                <%
                                index=0
                                Do While Not RS.Eof
                                refe=""
                                CONTE=""
                                    index=index+1
                                  refe=RS.Fields("refcia01").value
'                                  sql2="select refcia04,numgui04 from ssguia04 where refcia04='"&refe&"'"
                                  sql2="select refe09,guia09 from d09conoc where refe09='"&refe&"'"
                                  'sql3="select refe01,fdsp01,cbuq01 from c01refer where refe01='"&refe&"'"
                                  sql3="select refe01,fdsp01,cbuq01,navi06, cve01, nom01 from (c01refer inner join c06barco on cbuq01=clav06) inner join c55navie on navi06=cve01 where refe01='"&refe&"'"
                                  'sql4="select refe01,marc01,feacon01 from d01conte where refe01='"&refe&"' AND marc01='"&CONTE&"' "

                                  set RS2=oConn.Execute(sql2)
                                  if not RS2.eof then
                                      pBooking=RS2("guia09")
                                  else
                                      sql2="select refcia04,numgui04 from ssguia04 where refcia04='"&refe&"'"
                                      set RS2=oConn.Execute(sql2)
                                      if not RS2.eof then
                                        pBooking=RS2("numgui04")
                                      else
                                        pBooking="&nbsp;"
                                      end if
                                  end if

                                  set RS3=oConn.Execute(sql3)
'C=rs3.Fields("refe01").value
                                  'set RS4=oConn.Execute(sql4)
                                  'Response.Write(sql2)
                                  'Response.End
                                   %>
                                <tr>
                                     <td align="center"><font size="1" face="Arial"><%Response.Write(index) %></font></td>
                                     <%if RS("refcia01") <>"" then %>
                                     <td align="center"><font size="1" face="Arial"><%=RS("refcia01") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
                                     <%if RS("cvepod01") <>"" then %>
                                        <td align="center" nowrap><font size="1" face="Arial"><%=RS("cvepod01") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
                                     <%if RS("nombar01") <>"" then %>
                                        <td align="center"><font size="1" face="Arial"><%=RS("nombar01") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
                                     <%if RS("numcon40") <>"" then
                                            CONTE=rs.Fields("numcon40").value%>
                                        <td align="center"><font size="1" face="Arial"><%=RS("numcon40") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
<%%>
                                     <%
                                     'Response.Write(Session("GAduana"))
                                     if Session("GAduana") = "VER" then
                                     sql4="select refe01,marc01,feacon01 from d01conte where refe01='"&refe&"' AND REPLACE(marc01,'-','')=REPLACE('"&CONTE&"','-','')"
				     'if REFE="RKU08-03420" then
					'Response.Write(sql4)
					'RESPONSE.END
				     'end if				
                                     else
                                     sql4="select refe01,marc01,' / / ' as feacon01 from d01conte where refe01='"&refe&"' AND marc01='"&CONTE&"' "
                                     end if
                                     set RS4=oConn.Execute(sql4)
                                    if NOT RS4.EOF THEN'and RS4("feacon01")<>"" THEN not isnull(RS4("feacon01")) then and IF RS4("feacon01")<>"" THEN%>
                                        <td align="center"><font size="1" face="Arial"><%=RS4("feacon01") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
                                      <%if not RS2.eof then %>
                                        <!--td align="center"><font size="1" face="Arial"><%'=RS2("numgui04") %></font></td-->
                                        <td align="center"><font size="1" face="Arial"><%=pBooking %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
                                        <%if not RS3.EOF and not isnull(RS3("fdsp01"))  then 'RS3("fdsp01")<>"" and isdate(RS3("fdsp01")) then
                                     '    if RS3.fields("fdsp01").value<>"" and isdate(RS3.fields("fdsp01").value) then                                        %>
                                        <td align="center"><font size="1" face="Arial"><%=RS3("fdsp01") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
                                     <%If not RS3.eof and not isnull(RS3("nom01")) then %>
                                        <td align="center"><font size="1" face="Arial"><%=RS3("nom01") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
                                     <%if RS("numped01") <>"" then %>
                                        <td align="center"><font size="1" face="Arial"><%=RS("numped01") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
                                     <%if not RS3.EOF and not isnull(RS3("fdsp01")) then 'RS3("fdsp01") then  %>
                                        <td align="center"><font size="1" face="Arial"><%=RS3("fdsp01") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
<%'Response.Write(sql3)
'Response.Write(c)
'Response.End%>
                                     <%if RS("cveadu01") <>"" then %>
                                        <td align="center" nowrap><font size="1" face="Arial"><%=RS("cveadu01") %></font></td>
                                     <%else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                     <%end if%>
                                      <%'if RS("rfccli01") <>"" then %>
                                        <!--td align="center" nowrap><font size="1" face="Arial"><%=RS("rfccli01") %></font></td>
                                     <%'else%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td-->
                                     <%'end if%>
                                        <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                </tr>
                                <%
                                RS.MoveNext
                                Loop
								%>
                              </table>
                  <!--br><div align="center"><br></br><a href="" class="boton">:: CONSULTAR OTRO CLIENTE ::</a><br></div-->
                  <%else
                      'Response.Redirect "consulta.asp"
                  %>
                      <br><div align="center"><br></br><br></br><br></br><fieldset><a href="" class="boton"> :: VERIFIQUE SU RANGO DE FECHAS :: </a></fieldset></div>
                  <%
                  end if
              else
                %>
                 <br><div align="center"><br></br><br></br><br></br><fieldset><a href="" class="boton"> :: VERIFIQUE SU RANGO DE FECHAS :: </a></fieldset></div>
                <%
              end if
      else%>
        <br><div align="center"><br></br><br></br><br></br><fieldset><a href="" class="boton"> :: VERIFIQUE SUS FECHAS :: </a></fieldset></div>
    <%end if%>
    </td>
  </tr>
</table>
<%else
  response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>")
end if%>
</BODY>
</HTML>