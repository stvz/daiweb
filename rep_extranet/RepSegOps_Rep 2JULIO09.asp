<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% Server.ScriptTimeout=1500 %>
<HTML>
<HEAD>
<TITLE>:: REPORTE DE SEGUIMIENTO DE OPERACIONES.... ::</TITLE>
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
                                        sql1="SELECT fdoc01,frev01, fpre01, fdsp01,tipo01, refe01,clie01, nomcli01, refcia01,totbul01, fecpag01, Obser01 " & _
                                             " FROM c01refer inner join ssdagi01 on refe01= refcia01 and firmae01<>'"&""&"' and rfccli01='"&Vrfc&"' and tipo01='1'  and fdsp01>='"&DateI&"' and fdsp01<='"&DateF&"' and cveped01<>'"&"R1"&"' " & " ORDER BY clie01, REFE01"
                                          '   Response.Write("impo ck=0")
                                         '    Response.Write(sql1)
                                      else
                                        'Response.Write("impo ck=1")
                                        sql1="SELECT fdoc01,frev01, fpre01, fdsp01,tipo01, refe01,clie01, nomcli01, refcia01,totbul01, fecpag01, Obser01 " & _
                                             " FROM c01refer inner join ssdagi01 on refe01= refcia01 and firmae01<>'"&""&"'  " & permi & " and tipo01='1'  and fdsp01>='"&DateI&"' and fdsp01<='"&DateF&"' and cveped01<>'"&"R1"&"' " & " ORDER BY clie01, REFE01"
                                        'Response.Write(sql1)
                                      end if
                                    else
                                      if Vckcve="0" then
                                        'Response.Write("EXpo ck=0")
                                        sql1="SELECT fdoc01,frev01, fpre01, fdsp01,tipo01, refe01,clie01, nomcli01, refcia01,totbul01, fecpag01, Obser01 " & _
                                             " FROM c01refer inner join ssdage01 on refe01= refcia01 and firmae01<>'"&""&"' and rfccli01='"&Vrfc&"' and tipo01='2'  and fdsp01>='"&DateI&"' and fdsp01<='"&DateF&"' and cveped01<>'"&"R1"&"' " & " ORDER BY clie01, REFE01"
                                        '     Response.Write(sql1)
                                      else
                                       ' Response.Write("EXpo ck=1")
                                        sql1="SELECT fdoc01,frev01, fpre01, fdsp01,tipo01, refe01,clie01, nomcli01, refcia01,totbul01, fecpag01, Obser01 " & _
                                             " FROM c01refer inner join ssdage01 on refe01= refcia01 and firmae01<>'"&""&"'  " & permi & " and tipo01='2'  and fdsp01>='"&DateI&"' and fdsp01<='"&DateF&"' and cveped01<>'"&"R1"&"' " & " ORDER BY clie01, REFE01"
                                      '       Response.Write(sql1)
                                      end if

                                    end if
                                     'Response.Write(sql1)
                                     'Response.End
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
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Clientes</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Contenedores</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Bultos</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Documentos</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Revalidacion</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Previo</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">PagPedto</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Despacho</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">IndDsp</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">VacioCont</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">C.Gastos</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">IndCGast</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Observaciones</td>
                                    <%else%>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Ind</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Referencia</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Clientes</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Bultos</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">PagPedto</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Despacho</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">VacioCont</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">C.Gastos</td>
                                        <td width="100" nowrap><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Observaciones</td>
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
                                    Do While Not RS.Eof
                                        index=index+1
                                    '---------convertir el formato de fechas para la realizar la resta-----
                                        if mov="i" then
                                        '***estas fechas son para el indice de despacho
                                            xDFrev = RS("frev01")
                                            if isdate(xDFrev) then
                                              Diai = cstr(datepart("d",RS("frev01")))
                                              Mesi = cstr(datepart("m",RS("frev01")))
                                              Anioi = cstr(datepart("yyyy",RS("frev01")))
                                              DateRev = Diai & "/" &Mesi & "/"& Anioi
                                            end if

                                            xDFdsp=RS("fdsp01")
                                            if isdate(xDFdsp) then
                                              DiaF = cstr(datepart("d",RS("fdsp01")))
                                              MesF = cstr(datepart("m",RS("fdsp01")))
                                              AnioF = cstr(datepart("yyyy",RS("fdsp01")))
                                              DateDsp = DiaF & "/" &MesF & "/"& AnioF
                                            end if

                                            TimeDsp=DateDiff("d",DateRev,DateDsp)
                                      '----aki se saca la dif entre Revalidacion y despacho------------------
                                            x=0
                                            t=0
                                            Do While (x<=TimeDsp)
                                              x=x+1
                                              diasemana=WeekDay(DateAdd("d",x,DateDoc))
                                              if diasemana=1 then
                                                t=t+1
                                              end if
                                              if diasemana=7 then
                                                t=t+.5
                                              end if
                                            loop
                                            '----------------------
                                            TimeDsp=0 'para asegurarno k no tenga ningun valor
                                            TimeDsp=x-t' estos son los dias del Ind.Desp
                                            if TimeDsp>Tmax then
                                              Tmax=TimeDsp
                                            end if
                                            if TimeDsp<Tmin then
                                              Tmin=TimeDsp
                                            end if
                                            Tsum=Tsum+TimeDsp
                                        end if
                                    '--------------------------------------
                                    refe=RS("refcia01")
                                    obs=RS("Obser01")
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
                                    sql3="select d31refer.cgas31 as cg, fech31 "  & _
                                         "from d31refer inner join e31cgast on d31refer.cgas31=e31cgast.cgas31 and refe31='"&refe&"' AND esta31='I'"
                                         'Response.Write(sql3)
                                        ' Response.End
                                       set RS3=oConn.Execute(sql3)
                                       if not RS3.eof then
                                           TimeCG=0
                                           xFCG = ""
                                           xCount=1
                                           while not RS3.EOf
                                           if xCount=1 then
                                                xFCG =  RS3("fech31").value
                                               '--------AKI HACE EL CALCULO CON LA CUENTA DE GASTOS POR SER LA PRIMERA fecha---------
                                               IF mov="i" THEN
                                                    xDFcg = RS3("fech31")
                                                    if isdate(xDFcg) then
                                                      Diaf = cstr(datepart("d",RS3("fech31")))
                                                      Mesf = cstr(datepart("m",RS3("fech31")))
                                                      Aniof = cstr(datepart("yyyy",RS3("fech31")))
                                                      DateCG = Diaf & "/" &Mesf & "/"& Aniof
                                                    end if
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
                                           end if
                                           xCount = xCount + 1
                                           RS3.MoveNext
                                           wend
                                       else
                                           xFCG=""
                                       end if
                                     'Response.Write(sql3)
                                     'Response.End
                                       %>
                                    <tr>
                                    <%if mov="i" then%>
                                         <td align="center"><font size="1" face="Arial"><%Response.Write(index) %></font></td>
                                         <td align="center"><font size="1" face="Arial"><%=RS("refcia01") %></font></td>
                                         <td align="center" nowrap><font size="1" face="Arial"><%=RS("nomcli01") %></font></td>
                                         <!--td align="center"><font size="1" face="Arial"><%=RS("Obser01") %></font></td-->
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
                                         <%if TimeDsp <>"" then %>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial"><%=Response.Write( TimeDsp ) %></font></td>
                                         <%else%>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if xFCV <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCV ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if xFCG <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCG ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if xFCG <>"" then %>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial"><%=Response.Write( TimeCG ) %></font></td>
                                         <%else%>
                                            <td align="center" bgcolor="#0099CC"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if not RS4.eof then %>
                                            <td align="center" nowrap><font size="1" face="Arial"><%=RS4("c11desc")%></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>

                                    <%else%>
                                         <td align="center"><font size="1" face="Arial"><%Response.Write(index) %></font></td>
                                         <td align="center"><font size="1" face="Arial"><%=RS("refcia01") %></font></td>
                                         <td align="center" nowrap><font size="1" face="Arial"><%=RS("nomcli01") %></font></td>

                                         <%if RS("totbul01") <>"" then %>
                                            <td align="center" nowrap><font size="1" face="Arial"><%=RS("totbul01") %></font></td>
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

                                         <%if xFCV <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCV ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if xFCG <>"" then %>
                                            <td align="center"><font size="1" face="Arial"><%=Response.Write( xFCG ) %></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
                                         <%end if%>
                                         <%if not RS4.eof then %>
                                            <td align="center" nowrap><font size="1" face="Arial"><%=RS4("c11desc")%></font></td>
                                         <%else%>
                                            <td align="center"><font size="1" face="Arial">&nbsp;</font></td>
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