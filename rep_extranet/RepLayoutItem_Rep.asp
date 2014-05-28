<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% Server.ScriptTimeout=1500 %>
<HTML>
<HEAD>
<TITLE>:: LAYOUT DE OPERACIONES  ::</TITLE>
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
    strFiltroCliente = request.Form("txt_cvecli")
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

          <%
          Server.ScriptTimeOut =2000
          'texto="Bienvenidos a tutores.org!!"
          'response.write texto
          'response.write "<br>"
          'response.write(Replace(texto, "a tut", ".org"))
          'Response.End
          %>


          <%

          serv="10.66.1.5"
          base_datos= "rku_extranet"
          usu="carlosmg"
          pass="123456"

          'STRCKCVE= Request.Form("CKCVE")
          STRCKCVE= "1"
            strcvecli= request.Form("txt_cvecli")
            'strrfccli= Request.Form("txt_rfccli")
            strrfccli= "AME920102SS4"

          STROPRF=request.FORM("OPRF")
            STRCVEREFE=request.FORM("cverefe")
            STRFINI=request.FORM("FINI")
            STRFFIN=request.FORM("FFIN")

          strselTipo= Request.Form("selTipo")

          strtipRep= request.form("tipRep")

          if strselTipo="1" then
            operacion="IMPORTACION"
            tabla="ssdagi01"
            fecha="fecent01"
          else
          operacion="EXPORTACION"
            tabla="ssdage01"
            fecha="fecpre01"
          end if

          if strtipRep = "2" then
             Response.Addheader "Content-Disposition", "attachment;"
             Response.ContentType = "application/vnd.ms-excel"
          end if

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
              msjFechas=":: VERIFIQUE SUS FECHAS ::"
              %>
              <table width="50%"  border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
                <tr>
                  <td bgcolor="#DDDDDD" width="100%" height="30" align="center" ><font color="#336699"><strong><%=(msjFechas)%></td>
                </tr>
              </table>
              <%
                'Response.Write("<table width=""50%"" border=""1"" align=""center"" cellpadding=""0"" cellspacing=""7"" class=""titulosconsultas""><tr><td width=""18%"" height=""30"" class=""OpcPedimento"">:: VERIFIQUE SUS FECHAS ::</td></tr></table>")
                Response.End
          end if



          'option=no se bien como funcion pero siempre va el numero 16427
          'conention = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER="&serv&"; DATABASE="&base_datos&"; UID="&usu&"; PWD="&pass&"; OPTION=16427"
          conention= ODBC_POR_ADUANA(Session("GAduana"))
          'Response.Write(conention)

          'sqlReferencias= "select refcia01,nomcli01,tipopr01,adusec01,refcia01,patent01,numped01,fecpag01,fecent01,cveped01,regime01,pesobr01,tipcam01,cvepod01,cvepvc01,incble01,((i_dta101 + i_dta201) * factmo01)as AdvPagDLL,( (i_dta101 + i_dta201) * factmo01 * tipcam01)as AdvPagMX, cvepro01, factmo01,substring(desf0101,2,1) as CatidadFact "&_
          '                "from ssdagi01 "&_
          '                "where refcia01='RKU07-03603'"
          '
          '     set RSReferencias = server.CreateObject("ADODB.Recordset")
          '     RSReferencias.ActiveConnection = conention
          '     RSReferencias.Source= sqlReferencias
          '     RSReferencias.CursorType = 0
          '     RSReferencias.CursorLocation = 2
          '     RSReferencias.LockType = 1
          '     RSReferencias.Open()

          'while not RSReferencias.eof
          %>
          <table border="0">
              <tr><td COLSPAN="63"  ><font ><strong>:: REPORTE DE MERCANCIAS::</td></tr><BR>
              <tr><td COLSPAN="63"><font size="2"><strong><%=operacion%></td></tr><BR>
              <tr><td COLSPAN="63"><font size="2"><strong> DE <%=STRFINI%>  A <%=STRFFIN%></td></tr><BR>
          </table>
          <table align="center"  border="1" Width="100%">
          <!--TR class="boton">
                <td>    <font size="1" face="Arial">  1</td>
                <td>    <font size="1" face="Arial">  2</td>
                <td>    <font size="1" face="Arial">  3</td>
                <td>    <font size="1" face="Arial">  4</td>
                <td>    <font size="1" face="Arial">  5</td>
                <td>    <font size="1" face="Arial">  6</td>
                <td>    <font size="1" face="Arial">  7</td>
                <td>    <font size="1" face="Arial">  8</td>
                <td>    <font size="1" face="Arial">  9</td>
                <td>    <font size="1" face="Arial">  10</td>
                <td>    <font size="1" face="Arial">  11</td>
                <td>    <font size="1" face="Arial">  12</td>
                <td>    <font size="1" face="Arial">  13</td>
                <td>    <font size="1" face="Arial">  14</td>
                <td>    <font size="1" face="Arial">  15</td>
                <td>    <font size="1" face="Arial">  16</td>
                <td>    <font size="1" face="Arial">  17</td>
                <td>    <font size="1" face="Arial">  18</td>
                <td>    <font size="1" face="Arial">  19</td>
                <td>    <font size="1" face="Arial">  20</td>
                <td>    <font size="1" face="Arial">  21</td>
                <td>    <font size="1" face="Arial">  22</td>
                <td>    <font size="1" face="Arial">  23</td>
                <td>    <font size="1" face="Arial">  24</td>
                <td>    <font size="1" face="Arial">  25</td>
                <td>    <font size="1" face="Arial">  26</td>
                <td>    <font size="1" face="Arial">  27</td>
                <td>    <font size="1" face="Arial">  28</td>
                <td>    <font size="1" face="Arial">  29</td>
                <td>    <font size="1" face="Arial">  30</td>
                <td>    <font size="1" face="Arial">  31</td>
                <td>    <font size="1" face="Arial">  32</td>
                <td>    <font size="1" face="Arial">  33</td>
                <td>    <font size="1" face="Arial">  34</td>
                <td>    <font size="1" face="Arial">  35</td>
                <td>    <font size="1" face="Arial">  36</td>
                <td>    <font size="1" face="Arial">  37</td>
                <td>    <font size="1" face="Arial">  38</td>
                <td>    <font size="1" face="Arial">  39</td>
                <td>    <font size="1" face="Arial">  40</td>
                <td>    <font size="1" face="Arial">  41</td>
                <td>    <font size="1" face="Arial">  42</td>
                <td>    <font size="1" face="Arial">  43</td>
                <td>    <font size="1" face="Arial">  44</td>
                <td>    <font size="1" face="Arial">  45</td>
                <td>    <font size="1" face="Arial">  46</td>
                <td>    <font size="1" face="Arial">  47</td>
                <td>    <font size="1" face="Arial">  48</td>
                <td>    <font size="1" face="Arial">  49</td>
                <td>    <font size="1" face="Arial">  50</td>
                <td>    <font size="1" face="Arial">  51</td>
                <td>    <font size="1" face="Arial">  52</td>
                <td>    <font size="1" face="Arial">  53</td>
                <td>    <font size="1" face="Arial">  54</td>
                <td>    <font size="1" face="Arial">  55</td>
                <td>    <font size="1" face="Arial">  56</td>
                <td>    <font size="1" face="Arial">  57</td>
                <td>    <font size="1" face="Arial">  58</td>
                <td>    <font size="1" face="Arial">  59</td>
                <td>    <font size="1" face="Arial">  60</td>
                <td>    <font size="1" face="Arial">  61</td>
                <td>    <font size="1" face="Arial">  62</td>
                <td>    <font size="1" face="Arial">  63</td>

                <td>    <font size="1" face="Arial">    </td>
            </TR-->

            <tr class="boton" bgcolor="#336699">
              <td><font color="#FFFFFF" ><strong>Line Item</strong></font></td>

              <td><font color="#FFFFFF" ><strong>Cliente</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Tipo_Op</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Aduana</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Referencia</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Patente</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Pedimento</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Fecha de Pago</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Fecha Ent</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Clave</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Regimen</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Peso (kgs)</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Tipo Cambio</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Origen</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Procedencia</strong></font></td>
              <td><font color="#FFFFFF" ><strong>NAFTA_EUR</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Part Ped</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Part Fact</strong></font></td>
              <td><font color="#FFFFFF" ><strong>No Parte</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Descripcion</strong></font></td>
              <td><font color="#FFFFFF" ><strong>CantFacs</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Uni Fact</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Cant Tarifa</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Uni Comercial</strong></font></td>
              <td><font color="#FFFFFF" ><strong>No Factura</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Fech Fact</strong></font></td>
              <td><font color="#FFFFFF" ><strong>No PO Fact</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Tip Moneda Ext</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Factor Conv Moneda</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Incrementables</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Factor</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Fraccion</strong></font></td>
              <td><font color="#FFFFFF" ><strong>AdvMXPP</strong></font></td>
              <td><font color="#FFFFFF" ><strong>AdvUSDPP</strong></font></td>
              <td><font color="#FFFFFF" ><strong>FormaPAdv</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Tipo Tasa</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Tasa Pedimento</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Permisos</strong></font></td>
              <td><font color="#FFFFFF" ><strong>VCPedimentoMXP</strong></font></td>
              <td><font color="#FFFFFF" ><strong>VCPedimentoUSD</strong></font></td>
              <td><font color="#FFFFFF" ><strong>VAPedimentoUSD</strong></font></td>
              <td><font color="#FFFFFF" ><strong>VAPedimentoMXP</strong></font></td>
              <td><font color="#FFFFFF" ><strong>AdvMXPNPedimento</strong></font></td>
              <td><font color="#FFFFFF" ><strong>AdvUSDPedimento</strong></font></td>
              <td><font color="#FFFFFF" ><strong>DTAMXP</strong></font></td>
              <td><font color="#FFFFFF" ><strong>DTAUSD</strong></font></td>
              <td><font color="#FFFFFF" ><strong>IVAMXP</strong></font></td>
              <td><font color="#FFFFFF" ><strong>IVAUSD</strong></font></td>
              <td><font color="#FFFFFF" ><strong>FormapagoIVA</strong></font></td>
              <td><font color="#FFFFFF" ><strong>PrevalidacionMXP</strong></font></td>
              <td><font color="#FFFFFF" ><strong>ValCNNParteMX</strong></font></td>
              <td><font color="#FFFFFF" ><strong>ValCNNParteUSD</strong></font></td>
              <td><font color="#FFFFFF" ><strong>ValANparteMX</strong></font></td>
              <td><font color="#FFFFFF" ><strong>VaANParteUSD</strong></font></td>
              <td><font color="#FFFFFF" ><strong>AdvNParteMX</strong></font></td>
              <td><font color="#FFFFFF" ><strong>AdvNParteUSD</strong></font></td>
              <td><font color="#FFFFFF" ><strong>ValUnitarioNP</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Proveedor</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Codigo</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Viculacion</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Valoracion</strong></font></td>
              <td><font color="#FFFFFF" ><strong>ValDolFactura</strong></font></td>
              <td><font color="#FFFFFF" ><strong>Incoterm</strong></font></td>
              <!--td><font color="#FFFFFF" ><strong>("")</strong></font></td-->
          </tr>

          <%

          '*-*-*-*-*-*-*-*
          'sqlgra="select refcia01,nomcli01 from "&tabla&" where rfccli01='"&rfccliente&"' and firmae01<>"" and cveped01<>'R1' and fecpag01>="&ff&" and fecpag01<="&ff&""
          sqlgra="select refcia01,nomcli01 from "&tabla&" where firmae01<>'' and cveped01<>'R1' "

          IF STRCKCVE="1" THEN
              sqlgra= sqlgra &"" & permi
              'sqlgra= sqlgra &"AND cvecli01='"&strcvecli&"'"
          ELSE
              sqlgra= sqlgra &"AND rfccli01='"&strrfccli&"'"
          END IF

          IF STROPRF="1" THEN
              sqlgra= sqlgra &" AND REFCIA01='"&STRCVEREFE&"'"
          ELSE
              sqlgra= sqlgra &" AND fecpag01>='"&ISTRFINI&"' AND fecpag01<='"&FSTRFFIN&"'"
          END IF
          'Response.Write(PERMI)
          'Response.Write(sqlgra)
          'Response.End
          '*-*-*-*-*-*-*-*

               set rsgralrefes = server.CreateObject("ADODB.Recordset")
               rsgralrefes.ActiveConnection = conention
               rsgralrefes.Source= sqlgra
               rsgralrefes.CursorType = 0
               rsgralrefes.CursorLocation = 2
               rsgralrefes.LockType = 1
               rsgralrefes.Open()
          indgral=0

              while not rsgralrefes.eof

                  indgral=indgral + 1

                  if (indgral mod 2)=0 then
                      colfila="#ffffff"
                  else
                      colfila="#ffffff"
                      'colfila="#FFFFCC"
                  end if

                  refe=rsgralrefes.fields.ITEM("REFCIA01").VALUE
                    'sqlcamp="select nomcli01,tipopr01,adusec01,refcia01,patent01,numped01,fecpag01,"&fecha&"  as fech_ent_pre ,cveped01,regime01,pesobr01,tipcam01,cvepod01, "&_
                            '"cvepvc01,incble01,(((i_dta101 + i_dta201))) as DTAMXP,( (i_dta101 + i_dta201)/ tipcam01)as DTAUSD, cvepro01,cvecli01 "&_
                            '"factmo01, substring(desf0101,2,1) as CatidadFact,valfac01, refcia02,ordfra02, fraarn02, ((i_adv102 + i_adv202)) as AdvMXP,"&_
                            '"  (((i_adv102 + i_adv202) / tipcam01 )) as AdvDLL,(p_adv102 + p_adv202) as FormaPAdv, tasadv02,( (vaduan02 )) as VAPedimentoMXPx, "&_
                            '" (vaduan02 / tipcam01 ) as VAPedimentoUSDx, ((i_iva102 + i_iva202)) as IVAMXP,(((i_iva102 + i_iva202) / tipcam01 )) as IVADLL, "&_
                            '" valdls02, (p_iva102 + p_iva202) as FPagIVA, vincul02, metval02,((vmerme02) ) as CVPedimentoMXP,(vmerme02 / tipcam01) as CVPedimentoUSD,preuni02, "&_
                            '" refe05,frac05, pped05,item05, pfac05,desc05,umta05,caco05,umco05,fact05,pedi05,(incble01/vafa05) as factor,agru05, "&_
                            '" (((vafa05 * factmo01) * tipcam01)) as ValCNNParteMX, ((vafa05 * factmo01 )) as ValCNNParteUSD, refcia12,cveide12, ordfra12, "&_
                            '" if(cveide12='TL' and comide12='EMU','EUR',if (cveide12='TL' and (comide12='USA' or comide12='CAN'),'NAFTA','')) as NAFTA_EUR, "&_
                            '" refcia39,numfac39,fecfac39,monfac39,facmon39,terfac39,cvepro39,valdls39,valmex39 ,refcia36,cveimp36,import36 ,nompro22,irspro22 "&_
                            '" from ( ((((ssdagi01 join ssfrac02 on refcia01=refcia02) "&_
                            '"       join d05artic on ( refcia02=refe05 and fraarn02=frac05 and ordfra02=agru05 )) "&_
                            '"         join ssipar12 on (refcia02=refcia12 and ordfra02= ordfra12)) "&_
                            '"           join ssfact39 on (refe05=refcia39 and fact05=numfac39)) "&_
                            '"             join sscont36 on (refcia01=refcia36 and cveimp36=15) ) "&_
                            '"              join ssprov22 on (prov05=cvepro22) "&_
                            '"where refcia01='"&refe&"' "&_
                            '"group by ordfra02,agru05,item05,fact05 "&_
                            '"order by ordfra02"

                    sqlcamp="select nomcli01,tipopr01,adusec01,refcia01,patent01,numped01,fecpag01,"&fecha&"  as fech_ent_pre ,cveped01,regime01,pesobr01,tipcam01,cvepod01, "&_
                            "cvepvc01,incble01,(((i_dta101 + i_dta201))) as DTAMXP,( (i_dta101 + i_dta201)/ tipcam01)as DTAUSD, cvepro01,cvecli01 "&_
                            "factmo01, substring(desf0101,2,1) as CatidadFact,valfac01, refcia02,ordfra02, fraarn02, ((i_adv102 + i_adv202)) as AdvMXP,"&_
                            "  (((i_adv102 + i_adv202) / tipcam01 )) as AdvDLL,(p_adv102 + p_adv202) as FormaPAdv, tasadv02,( (vaduan02 )) as VAPedimentoMXPx, "&_
                            " (vaduan02 / tipcam01 ) as VAPedimentoUSDx, ((i_iva102 + i_iva202)) as IVAMXP,(((i_iva102 + i_iva202) / tipcam01 )) as IVADLL, "&_
                            " tt_adv02,  "&_
                            " valdls02, (p_iva102 + p_iva202) as FPagIVA, vincul02, metval02,((vmerme02) ) as CVPedimentoMXP,(vmerme02 / tipcam01) as CVPedimentoUSD,preuni02, "&_
                            " refe05,frac05, pped05,item05, pfac05,desc05,umta05,caco05,umco05,fact05,pedi05,(incble01/vafa05) as factor,agru05, "&_
                            " (((vafa05 * factmo01) * tipcam01)) as ValCNNParteMX, ((vafa05 * factmo01 )) as ValCNNParteUSD,  "&_
                            " refcia39,numfac39,fecfac39,monfac39,facmon39,terfac39,cvepro39,valdls39,valmex39 ,refcia36,cveimp36,import36 ,nompro22,irspro22 "&_
                            " from ( (((("&tabla&" join ssfrac02 on refcia01=refcia02) "&_
                            "       join d05artic on ( refcia02=refe05 and fraarn02=frac05 and ordfra02=agru05 )) "&_
                            "           join ssfact39 on (refe05=refcia39 and fact05=numfac39)) "&_
                            "             join sscont36 on (refcia01=refcia36 and cveimp36=15) ) "&_
                            "              join ssprov22 on (prov05=cvepro22)) "&_
                            "where refcia01='"&refe&"' "&_
                            "group by ordfra02,agru05,item05,fact05 "&_
                            "order by ordfra02"

                    'Response.Write(sqlcamp)
                    'Response.End
                         set RScamps = server.CreateObject("ADODB.Recordset")
                         RScamps.ActiveConnection = conention
                         RScamps.Source= sqlcamp
                         RScamps.CursorType = 0
                         RScamps.CursorLocation = 2
                         RScamps.LockType = 1
                         RScamps.Open()
                    ind=0
                        AdvMXPNPedimento = 0
                        AdvUSDPNPedimento = 0
                        IVAMXP = 0
                        IVAUSD = 0
                        NAFTA_EUR=""
                        while not RScamps.eof
                          permis = ""

                        ValANparteMX = 0
                        ValANparteUSD = 0
                        AdvNParteMX = 0
                        AdvNParteUSD = 0

                        ind=ind + 1
                    'Select Case (RScamps.fields.item("cvecli01").value)
                    '   Case :
                           'Sentencias
                    '       ...
                    '   Case else:
                           'Sentencias
                    '       ...
                    'End Select

                      if ind=1 then

                            'facmon=Replace(RScamps.fields.item("factmo01").value,",",".")
                            facmon=Replace(RScamps.fields.item("facmon39").value,",",".")
                            tipcam=Replace(RScamps.fields.item("tipcam01").value,",",".")
                            'IVAMXP=Replace(RScamps.fields.item("IVAMXP").value,",",".")
                            'IVAUSD=Replace(RScamps.fields.item("IVADLL").value,",",".")
                            'Response.Write(" facmon39 =  "&xfacmon39)
                            'Response.Write(" facmon01 =  "&facmon)
                            'Response.Write("  tipcam01=  "&tipcam)

                          sqlIVAyAdv="select sum(( i_adv102+i_adv202 ) / "&tipcam&" ) as AdvUSDPNPedimento, "&_
                                "(sum( i_adv102+i_adv202) )as AdvMXPNPedimento, "&_
                                "(sum( i_iva102+i_iva202 )) as IVAMXP,"&_
                                "((sum( i_iva102+i_iva202 )) / "&tipcam&" ) as IVAUSD, "&_
                                "(sum(vaduan02 )) as VAPedimentoMXP, "&_
                                "(sum(vaduan02) / "&tipcam&") as VAPedimentoUSD, "&_
                                "(sum(vmerme02) ) as CVPedimentoMXP,"&_
                                "(sum(vmerme02 / "&tipcam&")) as CVPedimentoUSD "&_
                                "from ssfrac02 "&_
                                "where refcia02='"&refe&"' "&_
                                "group by refcia02"
          '          Response.Write(sqlIVAyAdv)
                    'Response.End
                         set RSIVAyAdv = server.CreateObject("ADODB.Recordset")
                         RSIVAyAdv.ActiveConnection = conention
                         RSIVAyAdv.Source= sqlIVAyAdv
                         RSIVAyAdv.CursorType = 0
                         RSIVAyAdv.CursorLocation = 2
                         RSIVAyAdv.LockType = 1
                         RSIVAyAdv.Open()

                         if not RSIVAyAdv.eof then
                                AdvUSDPNPedimento=replace( RSIVAyAdv.fields.item("AdvUSDPNPedimento").value,",",".")
                                AdvMXPNPedimento=replace( RSIVAyAdv.fields.item("AdvMXPNPedimento").value,",",".")
                                IVAMXP=replace( RSIVAyAdv.fields.item("IVAMXP").value,",",".")
                                VAPedimentoMXP=replace( RSIVAyAdv.fields.item("VAPedimentoMXP").value,",",".")
                                VAPedimentoUSD=replace( RSIVAyAdv.fields.item("VAPedimentoUSD").value,",",".")
                                CVPedimentoMXP=replace( RSIVAyAdv.fields.item("CVPedimentoMXP").value,",",".")
                                CVPedimentoUSD=replace( RSIVAyAdv.fields.item("CVPedimentoUSD").value,",",".")
                         end if

                          RSIVAyAdv.close
                         set RSIVAyAdv = nothing
                         'numero de facturas
                          sqlNoFacs="select distinct  count(numfac39) as NoFacs from ssfact39 where refcia39='"&refe&"' group by  refcia39"
          '          Response.Write(sqlNoFacs)
                    'Response.End
                         set RSNoFacs = server.CreateObject("ADODB.Recordset")
                         RSNoFacs.ActiveConnection = conention
                         RSNoFacs.Source= sqlNoFacs
                         RSNoFacs.CursorType = 0
                         RSNoFacs.CursorLocation = 2
                         RSNoFacs.LockType = 1
                         RSNoFacs.Open()

                         if not RSNoFacs.eof then
                                NoFacs=replace( RSNoFacs.fields.item("NoFacs").value,",",".")
                         end if
                         RSNoFacs.close
                         set RSNoFacs = nothing
                         'numero de facturas
                    end if


                    '-----------------
                    vAdvFra=0
                    vAduaFra=0
                    '-----------------
                    xordfrac=RScamps.fields.item("ordfra02").value

                      'if xordfrac<> actualordfra02 then
                          actualordfra02=xordfrac
'/ "&tipcam&"
                          sqlsumAduYAdv="select refcia02,fraarn02,ordfra02,agru05, item05,"&_
                                        " (sum( i_adv102 + i_adv202 )) as Advfra,"&_
                                        " ((sum( i_adv102 + i_adv202 ))            ) as AdvfraUSD,"&_
                                        " (vaduan02 ) as Adu,"&_
                                        " (vaduan02 / "&tipcam&") as AduUSD,"&_
                                        " ((sum((vafa05 * "&facmon&") * "&tipcam&"))) as Sumavafac,"&_
                                        " ((sum(vafa05) * "&facmon&")) as SumavafacUSD, "&_
                                        " (( i_iva102+i_iva202 )) as vIVAMXP,"&_
                                        " ((( i_iva102+i_iva202 )) / "&tipcam&" ) as vIVAUSD "&_
                                        " from ssfrac02 join d05artic on (refcia02=refe05 and ordfra02=agru05 ) "&_
                                        " where refcia02=refe05 and ordfra02=agru05 and refcia02='"&refe&"' and ordfra02="&actualordfra02&" "&_
                                        " group by ordfra02 "&_
                                        " order by agru05"

          '                sqlsumAduYAdv="select max(ordfra02),refcia02,fraarn02,ordfra02,agru05,"&_
          '                              " sum(( i_adv102+i_adv202 )) as Advfra,"&_
          '                              " ((sum( i_adv102+i_adv202 ))/ "&tipcam&") as AdvfraUSD,"&_
          '                              " vaduan02 as Adu,(vaduan02 * "&tipcam&") as AduUSD, item05,"&_
          '                              " (((sum((vafa05 * "&facmon&")/ "&tipcam&")))) as Sumavafac,"&_
          '                              " ((sum((vafa05)) * "&facmon&")) as SumavafacUSD "&_
          '                              " from ssfrac02 join d05artic on (refcia02=refe05 and ordfra02=agru05 ) "&_
          '                              " where refcia02=refe05 and ordfra02=agru05 and refcia02='"&refe&"' and ordfra02="&actualordfra02&" "&_
          '                              " group by ordfra02 "&_
          '                              " order by agru05"

          '          Response.Write(sqlsumAduYAdv)
                         set sumAduYAdv = server.CreateObject("ADODB.Recordset")
                         sumAduYAdv.ActiveConnection = conention
                         sumAduYAdv.Source= sqlsumAduYAdv
                         sumAduYAdv.CursorType = 0
                         sumAduYAdv.CursorLocation = 2
                         sumAduYAdv.LockType = 1
                         sumAduYAdv.Open()

                         if not sumAduYAdv.eof then
                            vAdvFra=replace(sumAduYAdv.fields.item("Advfra").value,",",".")
                            vAdvFraUSD=replace(sumAduYAdv.fields.item("AdvfraUSD").value,",",".")

                            vVaduanFra=replace(sumAduYAdv.fields.item("Adu").value,",",".")
                            vVaduanFraUSD=replace(sumAduYAdv.fields.item("AduUSD").value,",",".")

                            vVafaFra=replace(sumAduYAdv.fields.item("Sumavafac").value,",",".")
                            vVafaFraUSD=replace(sumAduYAdv.fields.item("SumavafacUSD").value,",",".")

                            vIVAMXP=replace(sumAduYAdv.fields.item("vIVAMXP").value,",",".")
                            vIVAUSD=replace(sumAduYAdv.fields.item("vIVAUSD").value,",",".")
                         end if

                         sumAduYAdv.close
                         set sumAduYAdv = nothing

                      'end if ************ aki kdada pendiente el verificar pòr que no sale el advaloren por item05, y asi poder dividirlo ya ke con la multiplicaciones sale un valor grande.
                          'sqlvalorItem="select (((vafa05 * "&facmon&") * vaduan02)/ "&vAduaFra&" ) as vItemAdvMxpPedimento from ssfrac02 join d05artic on (refcia02=refe05 and fraarn02=frac05 and ordfra02=agru05 and refcia02='"&refe&"' and item05='"&RScamps.fields.item("item05").value&"' )"
                          sqlvalorItem="select (((vafa05 * "&facmon&" * "&tipcam&") *  vaduan02)/("&vVaduanFra&") ) as vItemADUMxpPedimento,"&_
                                       "(((((vafa05* "&facmon&" ) * (vaduan02/"&tipcam&"))/ "&vVaduanFraUSD&") ))  as vItemADUUSDpPedimento, "&_
                                       "((((vafa05)  * (i_adv102+i_adv202 ))/ ("&vAdvFra&")) * "&tipcam&" * "&facmon&") as vItemAdvMxpPedimento,  "&_
                                       "((((vafa05 ) * (((i_adv102+i_adv202 ))))/ "&vAdvFraUSD&" ) * "&facmon&") as vItemAdvUSDPedimento "&_
                                       "from ssfrac02 join d05artic on (refcia02=refe05 and fraarn02=frac05 and ordfra02=agru05 and "&_
                                       "refcia02='"&refe&"' and item05='"&RScamps.fields.item("item05").value&"' and agru05='"&RScamps.fields.item("agru05").value&"' ) "
          '            Response.Write(sqlvalorItem)
          '            Response.End
                         set Vprorrateados = server.CreateObject("ADODB.Recordset")
                         Vprorrateados.ActiveConnection = conention
                         Vprorrateados.Source= sqlvalorItem
                         Vprorrateados.CursorType = 0
                         Vprorrateados.CursorLocation = 2
                         Vprorrateados.LockType = 1
                         Vprorrateados.Open()

                         if not Vprorrateados.eof then
                            if isnull(Vprorrateados.fields.item("vItemADUMxpPedimento").value) then
                                ValANparteMX=0
                                ValANparteUSD=0
                            else
                                ValANparteMX=replace(Vprorrateados.fields.item("vItemADUMxpPedimento").value,",",".")
                                ValANparteUSD=replace(Vprorrateados.fields.item("vItemADUUSDpPedimento").value,",",".")
                            end if
                            if Isnull(Vprorrateados.fields.item("vItemAdvMxpPedimento").value) then
                                AdvNParteMX = 0
                                AdvNParteUSD = 0
                            else
                                AdvNParteMX=replace(Vprorrateados.fields.item("vItemAdvMxpPedimento").value,",",".")
                                AdvNParteUSD=replace(Vprorrateados.fields.item("vItemAdvUSDPedimento").value,",",".")
                            end if
                         end if

                         Vprorrateados.close
                         set Vprorrateados = nothing


                      permis=""
                      'TipoTasa=""
                    'TipoTasa y NAFTA_EUR
                      if ind=1 then
                         sqlTipoTasaNafta="select  if(cveide12='TL' and comide12='EMU','EUR',if (cveide12='TL' and (comide12='USA' or comide12='CAN'),'NAFTA','')) as NAFTA_EUR, "&_
                                      "cveide12 as TipoTasa "&_
                                      "from ssipar12 where refcia12='"&refe&"' and comide12<>'' and cveide12='TL' and ordfra12="&RScamps.fields.item("ordfra02").value&" "
                                      ' and ordfra12= "&RScamps.fields.item("ordfra02").value&"

                         set RSsqlTipoTasaNafta = server.CreateObject("ADODB.Recordset")
                         RSsqlTipoTasaNafta.ActiveConnection = conention
                         RSsqlTipoTasaNafta.Source= sqlTipoTasaNafta
                         RSsqlTipoTasaNafta.CursorType = 0
                         RSsqlTipoTasaNafta.CursorLocation = 2
                         RSsqlTipoTasaNafta.LockType = 1
                         RSsqlTipoTasaNafta.Open()

                            'NAFTA_EUR=RSsqlTipoTasaNafta.fields.item("NAFTA_EUR").value
                            'TipoTasa= RSsqlTipoTasaNafta.fields.item("TipoTasa").value

                         if RSsqlTipoTasaNafta.eof  then
                            NAFTA_EUR = "&nbsp;"
                         else
                            NAFTA_EUR=RSsqlTipoTasaNafta.fields.item("NAFTA_EUR").value
                         end if
                         'if TipoTasa="" then
                          '  TipoTasa = "&nbsp;"
                         'end if

                         RSsqlTipoTasaNafta.close
                         set RSsqlTipoTasaNafta= nothing
                      end if
                    'TipoTasa y NAFTA_EUR
                    '9999999
                          sqlpermisos="select  if(cveide12='TL' and comide12='EMU','EUR',if (cveide12='TL' and (comide12='USA' or comide12='CAN'),'NAFTA','')) as NAFTA_EUR, "&_
                                      "cveide12 as TipoTasa, concat_ws(',',cveide12,comide12) as cvepermis  "&_
                                      "from ssipar12 where refcia12='"&refe&"' and ordfra12= "&RScamps.fields.item("ordfra02").value&" "&_
                                      " group by cveide12"

                         set RSPermiso = server.CreateObject("ADODB.Recordset")
                         RSPermiso.ActiveConnection = conention
                         RSPermiso.Source= sqlpermisos
                         RSPermiso.CursorType = 0
                         RSPermiso.CursorLocation = 2
                         RSPermiso.LockType = 1
                         RSPermiso.Open()
                         pp=0
                         while not RSPermiso.eof
                            if pp=0 then
                                'NAFTA_EUR=RSPermiso.fields.item("NAFTA_EUR").value
                                permis= RSPermiso.fields.item("cvepermis").value
                                'TipoTasa= RSPermiso.fields.item("TipoTasa").value
                                pp=1
                            else
                                permis= permis &" , "& RSPermiso.fields.item("cvepermis").value
                            end if
                          RSPermiso.movenext
                         wend
                         RSPermiso.close
                         set RSPermiso= nothing

                         if permis="" then
                            permis = "&nbsp;"
                         end if
                    '9999999

                        %>
                    <tr bgcolor="<%=colfila%>">
                     <td ><font size="1" face="Arial"><%=ind%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("nomcli01").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("tipopr01").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("adusec01").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("refcia01").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("patent01").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("numped01").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("fecpag01").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("fech_ent_pre").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("cveped01").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("regime01").value%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("pesobr01").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("tipcam01").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("cvepod01").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("cvepvc01").value%></td>
                        <td><font size="1" face="Arial"><%=NAFTA_EUR%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("pped05").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("pfac05").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("item05").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("desc05").value%></td>
                        <!--td><font size="1" face="Arial"><%'=RScamps.fields.item("CatidadFact").value%></td-->
                        <td><font size="1" face="Arial"><%=NoFacs%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("umta05").value%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("caco05").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("umco05").value%></td>
                        <!--td><font size="1" face="Arial"><%'=RScamps.fields.item("fact05").value%></td><-->
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("numfac39").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("fecfac39").value%></td>
                        <td><font size="1" face="Arial"><%if isnull(RScamps.fields.item("pedi05").value) then Response.Write("&nbsp;") else Response.Write(RScamps.fields.item("pedi05").value) end if%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("monfac39").value%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("facmon39").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("incble01").value%></td>
                        <td><font size="1" face="Arial"><%if isnull(RScamps.fields.item("factor").value) then Response.Write(0) else Response.Write(Replace(RScamps.fields.item("factor").value,",",".")) end if%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("fraarn02").value%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("AdvMXP").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("AdvDLL").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("FormaPAdv").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("tt_adv02").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("tasadv02").value%></td>
                        <td><font size="1" face="Arial"><%=replace(permis,", ,",",")%></td>
                        <!--td><font size="1" face="Arial"><%'=replace(RScamps.fields.item("CVPedimentoMXP").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%'=replace(RScamps.fields.item("CVPedimentoUSD").value,",",".")%></td-->
                        <td><font size="1" face="Arial"><%=replace(CVPedimentoMXP,",",".")%></td>
                        <td><font size="1" face="Arial"><%=replace(CVPedimentoUSD,",",".")%></td>
                        <!--td><font size="1" face="Arial"><%'=replace(RScamps.fields.item("VAPedimentoUSD").value,",",".")%></td-->
                        <!--td><font size="1" face="Arial"><%'=replace(RScamps.fields.item("VAPedimentoMXP").value,",",".")%></td-->
                        <td><font size="1" face="Arial"><%=replace(VAPedimentoUSD,",",".")%></td>
                        <td><font size="1" face="Arial"><%=replace(VAPedimentoMXP,",",".")%></td>
                        <!--td><font size="1" face="Arial"><%'=RScamps.fields.item("import36").value%></tr-->
                        <!--td><font size="1" face="Arial"><%'=vAdvFra%></td-->
                        <td><font size="1" face="Arial"><%=AdvMXPNPedimento%></td>
                        <!--td><font size="1" face="Arial"><%'=vAdvFraUSD%></td-->
                        <td><font size="1" face="Arial"><%=AdvUSDPNPedimento%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("DTAMXP").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("DTAUSD").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=IVAMXP%></td>
                        <td><font size="1" face="Arial"><%=IVAUSD%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("FPagIVA").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("import36").value%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("ValCNNParteMX").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("ValCNNParteUSD").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=ValANparteMX%></td>
                        <td><font size="1" face="Arial"><%=ValANparteUSD%></td>
                        <td><font size="1" face="Arial"><%=AdvNParteMX%></td>
                        <td><font size="1" face="Arial"><%=AdvNParteUSD%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("preuni02").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("nompro22").value%></td>
                        <td><font size="1" face="Arial"><%=CSTR(RScamps.fields.item("irspro22").value)%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("vincul02").value%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("metval02").value%></td>
                        <td><font size="1" face="Arial"><%=replace(RScamps.fields.item("valdls02").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%=RScamps.fields.item("terfac39").value%></td>

                        <!--td><font size="1" face="Arial"><%'=RScamps.fields.item("frac05").value%></td-->
                        <!--td><font size="1" face="Arial"><%'=RScamps.fields.item("ordfra02").value%></td-->
                        <!--td><font size="1" face="Arial"><%'=RScamps.fields.item("item05").value%></td>

                        <td><font size="1" face="Arial"><%'=RScamps.fields.item("ordfra12").value%></td>
                        <td><font size="1" face="Arial"><%'=RScamps.fields.item("valmex39").value%></td>
                        <td><font size="1" face="Arial"><%'=RScamps.fields.item("valfac01").value%></td-->
                    </tr>
                    <%
                      RScamps.movenext()
                      wend

                    RScamps.close
                    set RScamps = nothing

            rsgralrefes.movenext()
            wend

          rsgralrefes.close
          set rsgralrefes = nothing
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