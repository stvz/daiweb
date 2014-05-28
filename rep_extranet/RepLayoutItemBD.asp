



<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<%



'Server.ScriptTimeout=1500
Server.ScriptTimeOut=100000

%>
<HTML>
<HEAD>
<TITLE>:: LAYOUT DE OPERACIONES BD ::</TITLE>
</HEAD>
<BODY>



<%

    strTipoUsuario = request.Form("TipoUser")
    strPermisos    = Request.Form("Permisos")
    permi          = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
    oficina_zego   = Session("GAduana")

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
          MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))


          'Response.Write(conention)
          'Response.End

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
              <tr><td COLSPAN="63"  ><font ><strong>:: REPORTE DE MERCANCIAS BD ::</td></tr><BR>
              <tr><td COLSPAN="63"><font size="2"><strong><%=operacion%></td></tr><BR>
              <tr><td COLSPAN="63"><font size="2"><strong> DE <%=STRFINI%>  A <%=STRFFIN%></td></tr><BR>
          </table>
          <table align="center"  border="1" Width="100%">
          <TR class="boton">

            <tr class="boton" bgcolor="#336699">
                <td><font color="#FFFFFF" ><strong>  Referencia        </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Tipo              </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Cliente           </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Aduana            </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Pedimento         </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Fecha de Pago     </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Fecha Ent         </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Clave             </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Regimen           </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Tipo Cambio       </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Valor comercial   </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Valor aduana      </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  IGI               </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  DTA               </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  IVA               </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Prevalidacion     </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  codigo parte      </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  mercancias        </strong></font></td> <!-- Descripcion mercancia-->
                <td><font color="#FFFFFF" ><strong>  factura           </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Fecha factura     </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  fraccion          </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Desc fraccion     </strong></font></td> <!-- Descripción de la fraccion-->
                <td><font color="#FFFFFF" ><strong>  Proveedor         </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Tax ID            </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Viculacion        </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Valoracion        </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  ValDolFactura     </strong></font></td>
                <td><font color="#FFFFFF" ><strong>  Incoterm          </strong></font></td>
          </tr>

          <%

          '*-*-*-*-*-*-*-*
          'sqlgra="select refcia01,nomcli01 from "&tabla&" where rfccli01='"&rfccliente&"' and firmae01<>"" and cveped01<>'R1' and fecpag01>="&ff&" and fecpag01<="&ff&""

          'sqlgra=" select refcia01, " & _
          '       " nomcli01         " & _
          '       " from "&tabla&"   " & _
          '       " where firmae01<>'' and " & _
          '       " cveped01<>'R1' "

          sqlgra = " SELECT refcia01,  nomcli01" & _
                   " FROM  "&tabla&",  ssiped11  " & _
                   " WHERE refcia11 = refcia01 " & _
                   "       AND cveide11 = 'IM' " & _
                   "       AND firmae01 <> ''  "



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

          sqlgra= sqlgra & " GROUP BY REFCIA01 "



          'sqlgra = sqlgra & " AND REFCIA01 IN ('RKU07-03634A','RKU07-04491A','RKU07-06154A','RKU07-09518A','RKU08-00177A','RKU08-00309A','RKU08-01388A','RKU08-01681A','RKU08-04731A','RKU08-05553','RKU08-05739','RKU08-05740','RKU08-05807','RKU08-05905','RKU08-05906','RKU08-05907','RKU08-05908','RKU08-05909','RKU08-05910','RKU08-05911','RKU08-06012','RKU08-06361','RKU08-06362','RKU08-06378','RKU08-06464','RKU08-06508','RKU08-06567')     "
          'sqlgra = sqlgra & " AND REFCIA01 IN ('RKU07-09885A')     "
          'sqlgra = sqlgra & " AND REFCIA01 IN ('SAP07-7018-2','SAP08-0972','SAP08-0972-1','SAP08-0972-2') "

          'sqlgra = sqlgra & " AND REFCIA01 IN ('DAI08-2458-1','DAI08-3247-1','DAI08-3541-1','DAI08-3543-1','DAI08-3700-1','DAI08-3969-1','DAI08-3994-1','DAI08-4059-1','DAI08-4059-2','DAI08-5109','DAI08-5396-1','DAI08-5441','DAI08-6002-1','DAI08-6586','DAI08-6733-1'," & _
          '         "'DAI08-6735-1','DAI08-6858','DAI08-7207','DAI08-7245','DAI08-7257-1','DAI08-7416','DAI08-7417','DAI08-7418','DAI08-7419','DAI08-7420','DAI08-7421'," & _
          '         "'DAI08-7438','DAI08-7441','DAI08-7442','DAI08-7477','DAI08-7478','DAI08-7479','DAI08-7480','DAI08-7481','DAI08-7482','DAI08-7506','DAI08-7515','DAI08-7535','DAI08-7536','DAI08-7537'," & _
          '         "'DAI08-7538','DAI08-7539','DAI08-7540','DAI08-7541','DAI08-7542','DAI08-7543','DAI08-7544','DAI08-7561','DAI08-7562','DAI08-7563','DAI08-7564','DAI08-7565','DAI08-7566'," & _
          '         "'DAI08-7575','DAI08-7576','DAI08-7577','DAI08-7601','DAI08-7608','DAI08-7652','DAI08-7653','DAI08-7654','DAI08-7655','DAI08-7656','DAI08-7657','DAI08-7662'," & _
          '         "'DAI08-7672','DAI08-7673','DAI08-7674','DAI08-7676','DAI08-7697','DAI08-7700','DAI08-7701','DAI08-7702','DAI08-7703','DAI08-7704','DAI08-7728','DAI08-7735'," & _
          '         "'DAI08-7738','DAI08-7739','DAI08-7749','DAI08-7751','DAI08-7752','DAI08-7755','DAI08-7756','DAI08-7757','DAI08-7761','DAI08-7762','DAI08-7764','DAI08-7765'," & _
          '         "'DAI08-7766','DAI08-7767','DAI08-7769','DAI08-7812','DAI08-7816','DAI08-7850','DAI08-7852','DAI08-7861','DAI08-7862','DAI08-7863','DAI08-7864'," & _
          '         "'DAI08-7866','DAI08-7867','DAI08-7868','DAI08-7869','DAI08-7870','DAI08-7893','DAI08-7894','DAI08-7895','DAI08-7896','DAI08-7897','DAI08-7898'," & _
          '         "'DAI08-7905','DAI08-7919','DAI08-7923','DAI08-7924','DAI08-7925','DAI08-7926','DAI08-7958','DAI08-7959','DAI08-7960','DAI08-7961','DAI08-7962'," & _
          '         "'DAI08-7963','DAI08-7980','DAI08-8009','DAI08-8010','DAI08-8011','DAI08-8012','DAI08-8013','DAI08-8014','DAI08-8015','DAI08-8021','DAI08-8036'," & _
          '         "'DAI08-8048','DAI08-8049','DAI08-8050','DAI08-8058','DAI08-8088','DAI08-8089','DAI08-8090','DAI08-8091','DAI08-8092','DAI08-8093','DAI08-8100'," & _
          '         "'DAI08-8107','DAI08-8120','DAI08-8121','DAI08-8122','DAI08-8123','DAI08-8124','DAI08-8125','DAI08-8126','DAI08-8166','DAI08-8167','DAI08-8202'," & _
          '         "'DAI08-8203','DAI08-8204','DAI08-8205','DAI08-8206','DAI08-8233','DAI08-8235','DAI08-8236','DAI08-8237','DAI08-8239','DAI08-8240','DAI08-8241','DAI08-8268'," & _
          '         "'DAI08-8269','DAI08-8316','DAI08-8317','DAI08-8318','DAI08-8319','DAI08-8351','DAI08-8353','DAI08-8399')"

          'sqlgra = sqlgra & " AND REFCIA01 IN ('DAI08-7564')"
          'sqlgra = sqlgra & " AND REFCIA01 IN ('DAI08-8772-1', 'DAI08-8529-1', 'DAI08-8466-1', 'DAI08-7416-1', 'DAI08-6564-1', 'DAI08-8391-1')"

          ' sqlgra = sqlgra & " AND REFCIA01 IN ('DAI08-1860', 'DAI08-5212', 'DAI08-5814', 'DAI08-6251', 'DAI08-6558', 'DAI08-7561', 'DAI08-7561', 'DAI08-7919', 'DAI08-9398') "

'sqlgra = sqlgra & " AND REFCIA01 IN ('DAI08-10251A','DAI08-10381A','DAI08-10633','DAI08-10639','DAI08-10640','DAI08-10641','DAI08-10658','DAI08-10659'," & _
'                  "'DAI08-10660','DAI08-10683','DAI08-10693','DAI08-10694','DAI08-10695','DAI08-10702','DAI08-10723'," & _
'                  "'DAI08-10724','DAI08-10725','DAI08-10726','DAI08-10727','DAI08-10728','DAI08-10729','DAI08-10730'," & _
'                  "'DAI08-10731','DAI08-10732','DAI08-10733','DAI08-10734','DAI08-10735','DAI08-10736','DAI08-10737'," & _
'                  "'DAI08-10754','DAI08-10755','DAI08-10766','DAI08-10771','DAI08-10783','DAI08-10786','DAI08-10787'," & _
'                  "'DAI08-10808','DAI08-10818','DAI08-10819','DAI08-10820','DAI08-10821','DAI08-10822','DAI08-10843'," & _
'                  "'DAI08-10849','DAI08-10852','DAI08-10853','DAI08-10854','DAI08-10855','DAI08-10856','DAI08-10857'," & _
'                  "'DAI08-10865','DAI08-10889','DAI08-10901','DAI08-10914','DAI08-10915','DAI08-10916','DAI08-10917'," & _
'                  "'DAI08-10918','DAI08-10927','DAI08-10946','DAI08-10947','DAI08-10948','DAI08-10985','DAI08-10986'," & _
'                  "'DAI08-10991','DAI08-10992','DAI08-11010','DAI08-11011','DAI08-11012','DAI08-11013','DAI08-11030'," & _
'                  "'DAI08-11031','DAI08-11032','DAI08-11033','DAI08-11044','DAI08-11068','DAI08-11070','DAI08-11071'," & _
'                  "'DAI08-11072','DAI08-11081','DAI08-11089','DAI08-11103','DAI08-11104','DAI08-11105','DAI08-11106'," & _
'                  "'DAI08-11107','DAI08-11108','DAI08-11109','DAI08-11110','DAI08-11111','DAI08-11112','DAI08-11113'," & _
'                  "'DAI08-11114','DAI08-11115','DAI08-11116','DAI08-11117','DAI08-11131','DAI08-11142','DAI08-11146'," & _
'                  "'DAI08-11156','DAI08-11157','DAI08-11171','DAI08-11173','DAI08-11173A','DAI08-11179','DAI08-11180'," & _
'                  "'DAI08-11181','DAI08-11182','DAI08-11183','DAI08-11184','DAI08-11185','DAI08-11186','DAI08-11187'," & _
'                  "'DAI08-11196','DAI08-11198','DAI08-11199','DAI08-11203','DAI08-3543-1','DAI08-4180-1','DAI08-4401-1'," & _
'                  "'DAI08-6641-1','DAI08-6733-1','DAI08-6735-1','DAI08-6811-1','DAI08-6905-1','DAI08-6919-1','DAI08-7257-1'," & _
'                  "'DAI08-7416','DAI08-9317-1','DAI08-9479-1','DAI09-0241-1','DAI09-1326' )"



                  'sqlgra = sqlgra & " AND REFCIA01 IN ('RKU08-00309A', 'RKU08-01388A', 'RKU08-01681A', 'RKU08-04731A' ) "

                  'sqlgra = sqlgra & " AND REFCIA01 IN ('DAI08-0659-1','DAI08-3543-1','DAI08-4180-1','DAI08-4401-1','DAI08-6641-1','DAI08-6733-1'," & _
                  '                                    "'DAI08-6735-1','DAI08-6811-1','DAI08-6919-1','DAI08-6905-1','DAI08-7257-1','DAI08-7416'," & _
                  '                                    "'DAI08-9317-1','DAI08-9479-1','DAI09-0241-1','DAI09-1326'  ) "










          'Response.Write(PERMI)
          'Response.Write(sqlgra)
          'Response.Write("<br>")
          'Response.Write(conention)
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

               'Response.Write(sqlgra)
               'Response.End

              while not rsgralrefes.eof
                  refe=""
                  tipcam=0
                  facmon=0
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


                  sqlcamp = "  select nomcli01,                                       " & _
                            "         tipopr01,                                       " & _
                            "         adusec01,                                       " & _
                            "         cveadu01,                                       " & _
                            "         refcia01,                                       " & _
                            "         patent01,                                       " & _
                            "         numped01,                                       " & _
                            "         fecpag01,                                       " & _
                                      fecha & " as fech_ent_pre ,                      " & _
                            "         cveped01,                                       " & _
                            "         regime01,                                       " & _
                            "         pesobr01,                                       " & _
                            "         tipcam01,                                       " & _
                            "         cvepod01,                                       " & _
                            "         cvepvc01,                                       " & _
                            "         paiori02,                                       " & _
                            "         paiscv02,                                       " & _
                            "        (segros01+fletes01+embala01+incble01) as incble01,  "&_
                            "         (((i_dta101 + i_dta201))) as DTAMXP,            " & _
                            "         ( (i_dta101 + i_dta201)/ tipcam01)as DTAUSD,    " & _
                            "         cvepro01,                                       " & _
                            "         cvecli01,                                       " & _
                            "         factmo01,                                       " & _
                            "         substring(desf0101,2,1) as CatidadFact,         " & _
                            "         valfac01,                                       " & _
                            "         refcia02,                                       " & _
                            "         ordfra02,                                       " & _
                            "         fraarn02,                                       " & _
                            "         substring(fraarn02,1,6) as fraccion06,          " & _
                            "         ((i_adv102 + i_adv202)) as AdvMXP,              " & _
                            "         (((i_adv102 + i_adv202) / tipcam01 )) as AdvDLL," & _
                            "         (p_adv102 + p_adv202) as FormaPAdv,             " & _
                            "         tasadv02,( (vaduan02 )) as VAPedimentoMXPx,     " & _
                            "         (vaduan02 / tipcam01 ) as VAPedimentoUSDx,      " & _
                            "         ((i_iva102 + i_iva202)) as IVAMXP,              " & _
                            "         (((i_iva102 + i_iva202) / tipcam01 )) as IVADLL," & _
                            "         tt_adv02,                                       " & _
                            "         valdls02,                                       " & _
                            "         (p_iva102 + p_iva202) as FPagIVA,               " & _
                            "         vincul02,                                       " & _
                            "         metval02,                                       " & _
                            "         ((vmerme02)*factmo01*tipcam01 ) as CVPedimentoMXP,                " & _
                            "         (vmerme02*factmo01) as CVPedimentoUSD,        " & _
                            "         preuni02,                                       " & _
                            "         refe05,                                         " & _
                            "         frac05,                                         " & _
                            "         pped05,                                         " & _
                            "         item05,                                         " & _
                            "         pfac05,                                         " & _
                            "         desc05,                                         " & _
                            "         d_mer102,                                       " & _
                            "         umta05,                                         " & _
                            "         caco05,                                         " & _
                            "         umco05,                                         " & _
                            "         fact05,                                         " & _
                            "         cancom02,                                       " & _
                            "         u_medc02,                                       " & _
                            "         cantar02,                                       " & _
                            "         u_medt02,                                       " & _
                            "         vmerme02,                                       " & _
                            "         pedi05,                                         " & _
                            "         ((segros01+valseg01+fletes01+embala01+incble01)/vafa05) as factor, " & _
                            "         agru05,                                         " & _
                            "         (((vafa05 * factmo01) * tipcam01)) as ValCNNParteMX, " & _
                            "         ((vafa05 * factmo01 )) as ValCNNParteUSD,            " & _
                            "         vafa05  as ValCNNParteME,                            " & _
                            "         refcia39,                                       " & _
                            "         numfac39,                                       " & _
                            "         if( fecfac39 >'1900-01-01', fecfac39, '0000-00-00') as  fecfac39,                                       " & _
                            "         monfac39,                                       " & _
                            "         facmon39,                                       " & _
                            "         terfac39,                                       " & _
                            "         cvepro39,                                       " & _
                            "         valdls39,                                       " & _
                            "         valmex39 ,                                      " & _
                            "         refcia36,                                       " & _
                            "         cveimp36,                                       " & _
                            "         import36 ,                                      " & _
                            "         nompro22,                                       " & _
                            "         irspro22,                                       " & _
                            "         npscli22,                                       " & _
                            "         valdol01,                                       " & _
                            "         cata05                                          " & _
                            " from ( (((("&tabla&" inner join ssfrac02 on refcia01=refcia02) "&_
                            "       left join d05artic on ( refcia01=refe05 and fraarn02=frac05 and ordfra02=agru05 )) "&_
                            "           left join ssfact39 on (refcia01=refcia39 and fact05=numfac39)) "&_
                            "             left join sscont36 on (refcia01=refcia36 and cveimp36=15) ) "&_
                            "              left join ssprov22 on (prov05=cvepro22)) "&_
                            " where refcia01='"&refe&"' "&_
                            " group by ordfra02,agru05,item05,fact05,prov05, pped05, pfac05,caco05,vafa05"&_
                            " order by ordfra02"


                            '  "        (segros01+valseg01+fletes01+embala01+incble01) as incble01,  "&_
                            '  ""         repcli01,                                       " & _


                            '" group by ordfra02,agru05,item05"&_


                    'sqlcamp="select nomcli01,tipopr01,adusec01,refcia01,patent01,numped01,fecpag01,"&fecha&"  as fech_ent_pre ,cveped01,regime01,pesobr01,tipcam01,cvepod01, "&_
                    '        "cvepvc01,incble01,(((i_dta101 + i_dta201))) as DTAMXP,( (i_dta101 + i_dta201)/ tipcam01)as DTAUSD, cvepro01,cvecli01 "&_
                    '        "factmo01, substring(desf0101,2,1) as CatidadFact,valfac01, refcia02,ordfra02, fraarn02,substring(fraarn02,1,6) as fraccion06, ((i_adv102 + i_adv202)) as AdvMXP,"&_
                    '        "  (((i_adv102 + i_adv202) / tipcam01 )) as AdvDLL,(p_adv102 + p_adv202) as FormaPAdv, tasadv02,( (vaduan02 )) as VAPedimentoMXPx, "&_
                    '        " (vaduan02 / tipcam01 ) as VAPedimentoUSDx, ((i_iva102 + i_iva202)) as IVAMXP,(((i_iva102 + i_iva202) / tipcam01 )) as IVADLL, "&_
                    '        " tt_adv02,  "&_
                    '        " valdls02, (p_iva102 + p_iva202) as FPagIVA, vincul02, metval02,((vmerme02) ) as CVPedimentoMXP,(vmerme02 / tipcam01) as CVPedimentoUSD,preuni02, "&_
                    '        " refe05,frac05, pped05,item05, pfac05,desc05,umta05,caco05,umco05,fact05,pedi05,(incble01/vafa05) as factor,agru05, "&_
                    '        " (((vafa05 * factmo01) * tipcam01)) as ValCNNParteMX, ((vafa05 * factmo01 )) as ValCNNParteUSD,  "&_
                    '        " refcia39,numfac39,fecfac39,monfac39,facmon39,terfac39,cvepro39,valdls39,valmex39 ,refcia36,cveimp36,import36 ,nompro22,irspro22, npscli22 "&_
                    '        " from ( (((("&tabla&" left join ssfrac02 on refcia01=refcia02) "&_
                    '        "       left join d05artic on ( refcia01=refe05 and fraarn02=frac05 and ordfra02=agru05 )) "&_
                    '        "           left join ssfact39 on (refcia01=refcia39 and fact05=numfac39)) "&_
                    '        "             left join sscont36 on (refcia01=refcia36 and cveimp36=15) ) "&_
                    '        "              left join ssprov22 on (prov05=cvepro22)) "&_
                    '        "where refcia01='"&refe&"' "&_
                    '        "group by ordfra02,agru05,item05,fact05 "&_
                    '        "order by ordfra02"

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

                         'Response.Write(sqlcamp)
                         'Response.End

                        while not RScamps.eof
                          permis = ""

                          ValANparteMX = 0
                          ValANparteUSD = 0
                          AdvNParteMX = 0
                          AdvNParteUSD = 0
                          'Strcontacto = CampoCliente(RScamps.fields.item("cvecli01").Value,"repcli18")
                          Strcontacto = ""
                          StrRepLegCli = CampoCliente(RScamps.fields.item("cvecli01").Value,"REPLEG18")


                         '******************************************************************************
                         '  EJECUTIVO DE TRAFICO
                         '******************************************************************************
                             StrEjecTraf = ""
                             sqlEjecTraf = " SELECT distinct REFE01, repcli01,NOMB40 "&_
                                           " FROM C01REFER INNER JOIN c40repleg ON repcli01=CVEREP40 "&_
                                           " WHERE REFE01 = '"&refe&"' "

                             ' Response.Write(sqlIVAyAdv)
                             ' Response.End
                             set RSEjecTraf = server.CreateObject("ADODB.Recordset")
                             RSEjecTraf.ActiveConnection = conention
                             RSEjecTraf.Source= sqlEjecTraf
                             RSEjecTraf.CursorType = 0
                             RSEjecTraf.CursorLocation = 2
                             RSEjecTraf.LockType = 1
                             RSEjecTraf.Open()
                             if not RSEjecTraf.eof then
                                    StrEjecTraf = RSEjecTraf.fields.item("NOMB40").value
                             end if
                             RSEjecTraf.close
                             set RSEjecTraf = nothing
                         '******************************************************************************
                          if StrEjecTraf <> "" then
                             Strcontacto = StrEjecTraf
                          else
                              if StrRepLegCli <> "" then
                                 Strcontacto = StrRepLegCli
                              else
                                 Strcontacto = ""
                              end if
                          end if


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

                            facmon=Replace(RScamps.fields.item("factmo01").value,",",".")
                            'facmon=Replace(RScamps.fields.item("facmon39").value,",",".")
                            tipcam=Replace(RScamps.fields.item("tipcam01").value,",",".")
                            'IVAMXP=Replace(RScamps.fields.item("IVAMXP").value,",",".")
                            'IVAUSD=Replace(RScamps.fields.item("IVADLL").value,",",".")
                            'Response.Write(" facmon39 =  "&xfacmon39)
                            'Response.Write(" facmon01 =  "&facmon)
                            'Response.Write("  tipcam01=  "&tipcam)

                          sqlIVAyAdv = " select sum(( i_adv102+i_adv202 ) / "&tipcam&" ) as AdvUSDPNPedimento, "&_
                                       "        (sum( i_adv102+i_adv202) )as AdvMXPNPedimento, "&_
                                       "        (sum( i_iva102+i_iva202 )) as IVAMXP,"&_
                                       "        ((sum( i_iva102+i_iva202 )) / "&tipcam&" ) as IVAUSD, "&_
                                       "        (sum(vaduan02 )) as VAPedimentoMXP, "&_
                                       "        (sum(vaduan02) / "&tipcam&") as VAPedimentoUSD, "&_
                                       "        (sum(vmerme02)*"&facmon&"*"&tipcam&") as CVPedimentoMXP,"&_
                                       "        (sum(vmerme02)*"&facmon&") as CVPedimentoUSD "&_
                                       " from ssfrac02 "&_
                                       " where refcia02='"&refe&"' "&_
                                       " group by refcia02"
                                       '  * factmo01 * tipcam01
                                       '"        (sum(vmerme02 / "&tipcam&")) as CVPedimentoUSD "&_
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

                                factorIncble = VAPedimentoMXP/CVPedimentoMXP
                         end if

                         RSIVAyAdv.close
                         set RSIVAyAdv = nothing

                         vdolPedimento = VAPedimentoMXP / tipcam
                         if not trim(RScamps.fields.item("valdol01").value) = "" and not trim(RScamps.fields.item("valdol01").value) = "0"  then
                             vdolPedimento = RScamps.fields.item("valdol01").value
                         end if



                         'numero de facturas
                         'sqlNoFacs=" select distinct count(numfac39) as NoFacs " & _
                         '          " from ssfact39 " & _
                         '          " where refcia39='"&refe&"' " & _
                         '          " group by  refcia39"

                         sqlNoFacs=" select numfac39, " & _
                                   " if( fecfac39 >'1900-01-01', fecfac39, '0000-00-00') as  fecfac39 " & _
                                   " from ssfact39 " & _
                                   " where refcia39='"&refe&"' "

                          'Response.Write(sqlNoFacs)
                          'Response.End
                         set RSNoFacs = server.CreateObject("ADODB.Recordset")
                         RSNoFacs.ActiveConnection = conention
                         RSNoFacs.Source= sqlNoFacs
                         RSNoFacs.CursorType = 0
                         RSNoFacs.CursorLocation = 2
                         RSNoFacs.LockType = 1
                         RSNoFacs.Open()
                         NoFacs = 0
                         factLar = ""
                         FecfactLar = ""
                         if not RSNoFacs.eof then
                           while not RSNoFacs.eof

                               'Response.Write( RSNoFacs.fields.item("fecfac39").value )

                                'NoFacs=replace( RSNoFacs.fields.item("NoFacs").value,",",".")
                              if NoFacs = 0 then
                                factLar    = RSNoFacs.fields.item("numfac39").value
                                'if RSNoFacs.fields.item("fecfac39").value <> "0000-00-00" then
                                   FecfactLar = RSNoFacs.fields.item("fecfac39").value
                                'end if
                              else
                                factLar    = factLar    & " ; "& RSNoFacs.fields.item("numfac39").value
                                'if RSNoFacs.fields.item("fecfac39").value <> "0000-00-00" then
                                  FecfactLar = FecfactLar & " ; "& RSNoFacs.fields.item("fecfac39").value
                                'end if
                              end if
                                NoFacs = NoFacs + 1
                                RSNoFacs.movenext()
                           wend
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
                                        " (( i_adv102 + i_adv202 )) as Advfra,"&_
                                        " ((( i_adv102 + i_adv202 ))            ) as AdvfraUSD, "&_
                                        " (vaduan02 ) as Adu,"&_
                                        " (vaduan02 / "&tipcam&") as AduUSD,"&_
                                        " ((((vmerme02 * "&facmon&") * "&tipcam&"))) as Sumavafac,"&_
                                        " (((vmerme02) * "&facmon&")) as SumavafacUSD, "&_
                                        " ((vmerme02) ) as SumavafacME, "&_
                                        " (( i_iva102+i_iva202 )) as vIVAMXP,"&_
                                        " ((( i_iva102+i_iva202 )) / "&tipcam&" ) as vIVAUSD "&_
                                        " from ssfrac02 left join d05artic on (refcia02=refe05 and ordfra02=agru05 ) "&_
                                        " where refcia02='"&refe&"' and ordfra02="&actualordfra02&" "&_
                                        " group by ordfra02 "&_
                                        " order by agru05"

                          'sqlsumAduYAdv="select refcia02,fraarn02,ordfra02,agru05, item05,"&_
                          '              " (sum( i_adv102 + i_adv202 )) as Advfra,"&_
                          '              " ((sum( i_adv102 + i_adv202 ))            ) as AdvfraUSD, "&_
                          '              " (vaduan02 ) as Adu,"&_
                          '              " (vaduan02 / "&tipcam&") as AduUSD,"&_
                          '              " ((sum((vmerme02 * "&facmon&") * "&tipcam&"))) as Sumavafac,"&_
                          '              " ((sum(vmerme02) * "&facmon&")) as SumavafacUSD, "&_
                          '              " (sum(vmerme02) ) as SumavafacME, "&_
                          '              " (( i_iva102+i_iva202 )) as vIVAMXP,"&_
                          '              " ((( i_iva102+i_iva202 )) / "&tipcam&" ) as vIVAUSD "&_
                          '              " from ssfrac02 left join d05artic on (refcia02=refe05 and ordfra02=agru05 ) "&_
                          '              " where refcia02='"&refe&"' and ordfra02="&actualordfra02&" "&_
                          '              " group by ordfra02 "&_
                          '              " order by agru05"

                    'Response.Write(sqlsumAduYAdv)
                    'Response.End

                         set sumAduYAdv = server.CreateObject("ADODB.Recordset")
                         sumAduYAdv.ActiveConnection = conention
                         sumAduYAdv.Source= sqlsumAduYAdv
                         sumAduYAdv.CursorType = 0
                         sumAduYAdv.CursorLocation = 2
                         sumAduYAdv.LockType = 1
                         sumAduYAdv.Open()

                         vAdvFra       = 1
                         vAdvFraUSD    = 1
                         vVaduanFra    = 1
                         vVaduanFraUSD = 1
                         vVafaFra      = 1
                         vVafaFraUSD   = 1
                         vIVAMXP       = 1
                         vIVAUSD       = 1

                         if not sumAduYAdv.eof then
                            'vAdvFra=replace(sumAduYAdv.fields.item("Advfra").value,",",".")
                            'vAdvFraUSD=replace(sumAduYAdv.fields.item("AdvfraUSD").value,",",".")
                            'vVaduanFra=replace(sumAduYAdv.fields.item("Adu").value,",",".")
                            'vVaduanFraUSD=replace(sumAduYAdv.fields.item("AduUSD").value,",",".")
                            'vVafaFra=replace(sumAduYAdv.fields.item("Sumavafac").value,",",".")
                            'vVafaFraUSD=replace(sumAduYAdv.fields.item("SumavafacUSD").value,",",".")
                            'vIVAMXP=replace(sumAduYAdv.fields.item("vIVAMXP").value,",",".")
                            'vIVAUSD=replace(sumAduYAdv.fields.item("vIVAUSD").value,",",".")

                            vAdvFra       =  sumAduYAdv.fields.item("Advfra").value
                            vAdvFraUSD    =  sumAduYAdv.fields.item("AdvfraUSD").value
                            vVaduanFra    =  sumAduYAdv.fields.item("Adu").value
                            vVaduanFraUSD =  sumAduYAdv.fields.item("AduUSD").value
                            vVafaFra      =  sumAduYAdv.fields.item("Sumavafac").value
                            vVafaFraUSD   =  sumAduYAdv.fields.item("SumavafacUSD").value
                            vVafaFraME    =  sumAduYAdv.fields.item("SumavafacME").value
                            vIVAMXP       =  sumAduYAdv.fields.item("vIVAMXP").value
                            vIVAUSD       =  sumAduYAdv.fields.item("vIVAUSD").value
                         end if

                         sumAduYAdv.close
                         set sumAduYAdv = nothing


                         'Response.Write("sumAduYAdv")
                         'Response.End


                      'end if ************ aki kdada pendiente el verificar pòr que no sale el advaloren por item05, y asi poder dividirlo ya ke con la multiplicaciones sale un valor grande.
                          'sqlvalorItem="select (((vafa05 * "&facmon&") * vaduan02)/ "&vAduaFra&" ) as vItemAdvMxpPedimento from ssfrac02 join d05artic on (refcia02=refe05 and fraarn02=frac05 and ordfra02=agru05 and refcia02='"&refe&"' and item05='"&RScamps.fields.item("item05").value&"' )"
                          'sqlvalorItem="select (((vafa05 ) *  vaduan02)/("&vVafaFra&") ) as vItemADUMxpPedimento,"&_
                          '             "(((((vafa05) * (vaduan02))/ "&vVafaFra&") )/ "&tipcam&")  as vItemADUUSDpPedimento, "&_
                          '             "((((vafa05)  * (i_adv102+i_adv202 ))/ ("&vVafaFra&"))                           ) as vItemAdvMxpPedimento,  "&_
                          '             "((((vafa05) * (((i_adv102+i_adv202 ))))/ "&vVafaFra&" ) / "&tipcam&"          ) as vItemAdvUSDPedimento "&_
                          '             "from ssfrac02 join d05artic on (refcia02=refe05 and fraarn02=frac05 and ordfra02=agru05 and "&_
                          '             "refcia02='"&refe&"' and item05='"&RScamps.fields.item("item05").value&"' and agru05='"&RScamps.fields.item("agru05").value&"' ) "


                          sqlvalorItem=" Select ((((vafa05 ) *  vaduan02)/("&vVafaFraME&") )   ) as vItemADUMxpPedimento,"&_
                                       "        (((((vafa05) * (vaduan02))/ "&vVafaFraME&") ) / "&tipcam&" )  as vItemADUUSDpPedimento, "&_
                                       "        ((((vafa05)  * (i_adv102+i_adv202 ))/ ("&vVafaFraME&"))   ) as vItemAdvMxpPedimento,  "&_
                                       "        ((((vafa05)  * (((i_adv102+i_adv202 ))))/ "&vVafaFraME&" ) / "&tipcam&"          ) as vItemAdvUSDPedimento "&_
                                       " From ssfrac02 left join d05artic on (refcia02=refe05 and fraarn02=frac05 and ordfra02=agru05 )  "&_
                                       " Where refcia02='"&refe&"' and item05='"&RScamps.fields.item("item05").value&"' and agru05='"&RScamps.fields.item("agru05").value&"'  "

                          'sqlvalorItem=" Select ((((vafa05 ) *  vaduan02)/("&vVafaFra&") ) * "&facmon&" * "&tipcam&") as vItemADUMxpPedimento,"&_
                          '             "        (((((vafa05) * (vaduan02))/ "&vVafaFra&") ) * "&facmon&")  as vItemADUUSDpPedimento, "&_
                          '             "        ((((vafa05)  * (i_adv102+i_adv202 ))/ ("&vVafaFra&"))   ) as vItemAdvMxpPedimento,  "&_
                          '             "        ((((vafa05)  * (((i_adv102+i_adv202 ))))/ "&vVafaFra&" ) / "&tipcam&"          ) as vItemAdvUSDPedimento "&_
                          '             " From ssfrac02 left join d05artic on (refcia02=refe05 and fraarn02=frac05 and ordfra02=agru05 )  "&_
                          '             " Where refcia02='"&refe&"' and item05='"&RScamps.fields.item("item05").value&"' and agru05='"&RScamps.fields.item("agru05").value&"'  "

                      'Response.Write("sqlvalorItem=")
                      'Response.Write(sqlvalorItem)
                      'Response.End

                         set Vprorrateados = server.CreateObject("ADODB.Recordset")
                         Vprorrateados.ActiveConnection = conention
                         Vprorrateados.Source= sqlvalorItem
                         Vprorrateados.CursorType = 0
                         Vprorrateados.CursorLocation = 2
                         Vprorrateados.LockType = 1
                         Vprorrateados.Open()

                         'Response.Write("sumAduYAdv")
                         'Response.End

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

                         sqlpermisos= " select  if(cveide12='TL' and comide12='EMU','EUR',if (cveide12='TL' and (comide12='USA' or comide12='CAN'),'NAFTA','')) as NAFTA_EUR, "&_
                                      " cveide12 as TipoTasa, " &_
                                      " concat_ws(',',cveide12,comide12) as cvepermis  "&_
                                      " from ssipar12 " &_
                                      " where refcia12='"&refe&"' and ordfra12= "&RScamps.fields.item("ordfra02").value&" "&_
                                      " and   cveide12 <> 'MA'  "&_
                                      " group by cveide12"

                         'Response.Write(sqlpermisos)

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

                         'if permis="" then
                         '   permis = "&nbsp;"
                         'end if
                    '9999999

                     StrTipoTasaIgi = ""

                     'if isNull(permis) or permis="" or isEmpty(permis) then
                     if Len(Trim(permis)) = 0 then
                       StrTipoTasaIgi = ""
                     else
                         if InStr(permis,"TL") > 0 then
                           StrTipoTasaIgi = "TL"
                         else
                             if InStr(permis,"PS") > 0 then
                               StrTipoTasaIgi = "PS"
                             else
                               StrTipoTasaIgi = "NORMAL"
                             end if
                         end if
                     end if





                        %>
                    <tr bgcolor="<%=colfila%>">

                        <!--StrEjecTraf
                        Strcontacto -->
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("refcia01").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("tipopr01").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("nomcli01").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("adusec01").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("cveadu01").value  & "-"&RScamps.fields.item("patent01").value&"-"&RScamps.fields.item("numped01").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("fecpag01").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("fech_ent_pre").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("cveped01").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("regime01").value%></td>
                        <td><font size="1" face="Arial"><%= replace(RScamps.fields.item("tipcam01").value,",",".")%></td>
                        <td><font size="1" face="Arial"><%= CVPedimentoUSD %></td>
                        <td><font size="1" face="Arial"><%= VAPedimentoMXP %></td>
                        <td><font size="1" face="Arial"><%= AdvNParteMX%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("DTAMXP").value %></td>
                        <td><font size="1" face="Arial"><%= IVAMXP%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("import36").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("item05").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("desc05").value %></td> <!-- Descripcion mercancia-->
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("numfac39").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("fecfac39").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("fraarn02").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("d_mer102").value %></td> <!-- Descripción de la fraccion-->
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("nompro22").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("irspro22").value %></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("vincul02").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("metval02").value%></td>
                        <td><font size="1" face="Arial"><%= vdolPedimento %></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("terfac39").value%></td>






                        <!--Contacto -->
                        <!--
                        <td><font size="1" face="Arial"><%= Strcontacto%> </td>
                        <td><font size="1" face="Arial"><%= ind%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("patent01").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("paiori02").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("paiscv02").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("pped05").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("pfac05").value%></td>
                        <td><font size="1" face="Arial"><%  if isnull(RScamps.fields.item("pedi05").value) then Response.Write("&nbsp;") else Response.Write(RScamps.fields.item("pedi05").value) end if%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("caco05").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("umco05").value %></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("cantar02").value %></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("u_medt02").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("pesobr01").value %></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("monfac39").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("facmon39").value %></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("incble01").value%></td>
                        <td><font size="1" face="Arial"><%= factorIncble %></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("fraccion06").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("FormaPAdv").value%></td>
                        <td><font size="1" face="Arial"><%= StrTipoTasaIgi%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("tasadv02").value&"%"%></td>
                        <td><font size="1" face="Arial"><%= permis %></td>
                        <td><font size="1" face="Arial"><%= CVPedimentoMXP%></td>
                        <td><font size="1" face="Arial"><%= VAPedimentoUSD%></td>
                        <td><font size="1" face="Arial"><%= AdvMXPNPedimento%></td>
                        <td><font size="1" face="Arial"><%= AdvUSDPNPedimento%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("DTAUSD").value %></td>
                        <td><font size="1" face="Arial"><%= (IVAMXP/ tipcam)%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("FPagIVA").value%></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("ValCNNParteMX").value  %></td>
                        <td><font size="1" face="Arial"><%= RScamps.fields.item("ValCNNParteUSD").value %></td>
                        <td><font size="1" face="Arial"><%= ValANparteMX%></td>
                        <td><font size="1" face="Arial"><%= ValANparteUSD%></td>
                        <td><font size="1" face="Arial"><%= AdvNParteUSD%></td>
                        -->


                    </tr>
                    <%
                      RScamps.movenext()
                      'Response.Flush


                      'Response.Write("primer fraccion")
                      'Response.End

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



