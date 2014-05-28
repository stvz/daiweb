 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<% 
    'serv="10.66.1.5"
    'base_datos="dai_extranet"
    'usu="carlosmg"
    'pass="123456"
    'MM_STRING = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER="&serv&"; DATABASE="&base_datos&"; UID="&usu&"; PWD="&pass&"; OPTION=16427"

    MM_STRING = ODBC_POR_ADUANA(Session("GAduana"))
    'MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

    'Response.Write(Session("GAduana"))
    'Response.End


    Response.Buffer = TRUE
    Response.Addheader "Content-Disposition", "attachment;filename=HoneyWell.xls"
    Response.ContentType = "application/vnd.ms-excel"
    Server.ScriptTimeOut=100000

    strUsuario     = request.Form("user")
    strTipoUsuario = request.Form("TipoUser")
    strPermisos    = Request.Form("Permisos")
    permi          = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
    permi2         = PermisoClientesTabla("B",Session("GAduana") ,strPermisos,"clie31")

    if not permi2 = "" then
      permi2 = "  and (" & permi2 & ") "
    end if

    AplicaFiltro = false
    strFiltroCliente = ""
    strFiltroCliente = request.Form("txtCliente")
    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
       blnAplicaFiltro = true
    end if
    if blnAplicaFiltro then
       permi2 = " AND B.clie31 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
       permi2 = ""
    end if

    if not permi = "" then
      permi = "  and (" & permi & ") "
    end if

    AplicaFiltro = false
    'strFiltroCliente = ""
    'strFiltroCliente = request.Form("txtCliente")
    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
       blnAplicaFiltro = true
    end if
    if blnAplicaFiltro then
       permi = " AND cvecli01 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
       permi = ""
    end if

    '*****************************************************************************************************************************
    strDateIni       = ""
    strDateFin       = ""
    strTipoPedimento = ""
    strCodError      = "0"

    '*******************************************************
    strTipoFiltro         = trim(request.Form("TipoFiltro"))
    strDateIni            = trim(request.Form("txtDateIni"))
    strDateFin            = trim(request.Form("txtDateFin"))
    strTipoOperacion      = trim(request.Form("rbnTipoDate"))

    'strLinNav             = trim(request.Form("txtLinNav"))
    'strModalidad          = trim(request.Form("txtMod"))
    'strProv               = trim(request.Form("txtProv"))
    'strfiltrosrestantes   = trim(request.Form("txtfiltrosrestantes"))
    'strTipoOtrosFiltros   = trim(request.Form("txttipoOtrosFiltros"))

    'strTipoConte = trim(request.Form("txttipoConte"))
    '*******************************************************

    if not isdate(strDateIni) then
      strCodError = "5"
    end if
    if not isdate(strDateFin) then
      strCodError = "6"
    end if
    if strDateIni="" or strDateFin="" then
      strCodError = "1"
    end if

    ' Response.Write(strTipoFiltro)
    ' strHTML = ""

    if strCodError = "0" then

    tempstrOficina = adu_ofi( Session("GAduana") )
    'Response.Write(tempstrOficina)
    'Response.Write( enproceso(tempstrOficina) )


    IF NOT enproceso(tempstrOficina) THEN

    '******************************************************************************************************

    '+----------+-----------------------------------------------+--------------+
    '| cvecli18 | nomcli18       DAI                            | rfccli18     |
    '+----------+-----------------------------------------------+--------------+
    '|     3297 | HONEYWELL OPERATIVA MEXICO S. DE R.L. DE C.V. | HOM050106JIA |
    '|     3309 | HONEYWELL S.A. DE C.V.                        | HON641119JI7 |
    '+----------+-----------------------------------------------+--------------+

    ISTRFINI = strDateIni
    FSTRFFIN = strDateFin

    'CCLIENTE=" 3297 or cvecli01=3309 "
    'Response.Write(CCLIENTE)
    'Response.Write(permi)
    'Response.Write(permi2)
    'Response.End

  'base = "MEX"
  base = Session("GAduana")

  Select Case (base)
   Case "VER":
       c_fleteTerr="15"
       c_man="2"
       c_fleteAereo="35"
       c_desconsolidado="85"
       c_maniobrasyalmacenajes="230"
   Case "MEX":
       c_fleteTerr="7"
       c_man="127"
       c_fleteAereo="3"
       c_desconsolidado="6"
       c_maniobrasyalmacenajes="2"
   Case "MAN":
       c_fle="5"
       c_man="2"
   Case "TAM":
      c_fle="15"
       c_man="2"
   Case "LZR":
       c_fle="15"
       c_man="2"
   Case "GUA":
       c_fle="15"
       c_man="2"
   Case "LAR":
       c_fle="15"
       c_man="2"
   End Select

  'Response.Write( strTipoOperacion )
  'Response.end

  top="i"
  if strTipoOperacion = 1 then
    tabla       = "ssdagi01"
    num_con_f_m = c_fleteTerr
    fecEntAdu   = "fecent01"
    pTipOpe     = "Importación"
  else
    tabla       = "ssdage01"
    num_con_f_m = c_fleteTerr
    fecEntAdu   = "fecpre01"
    pTipOpe     = "Exportación"
  end if


'sqlReferencias= " Select refcia01,numped01,(i_dta101+i_dta201) as dta, cvepro01, cvepod01,cvepvc01, "&_
'                "       tipcam01,fecpag01,"&fecEntAdu&" as fechaEA, firmae01,desf0101,pesobr01,valmer01,segros01,fletes01, sum(fletes01+segros01) as fleyseg,cveped01,"&_
'                "       desdoc01,cveadu01,cvepfm01,cvecli01,"&_
'                "       fech31,d.cgas31 as cg "&_
'                " FROM "&tabla&" LEFT JOIN d31refer AS d on d.refe31=refcia01 left join  e31cgast as e  on e.cgas31=d.cgas31  " & _
'                " where firmae01<>'' and  cveped01<>'R1' and fecpag01 >= '"&FormatoFechaInv(ISTRFINI)&"' and fecpag01 <= '"&FormatoFechaInv(FSTRFFIN)&"'  "&_
'                " and ( cvecli01="&CCLIENTE&" )"&_
'                " group by refcia01 "

sqlReferencias =  " SELECT DATE_FORMAT(fecpag01, '%m/%d/%Y') AS 'fecpag01',  " &_
                  "        refcia01,  " &_
                  "        fraarn02,  " &_
                  "        PTOEMB01 AS PORT_LOADING, " &_
                  "        REGBAR01,  " &_
                  "        ordfra02,  " &_
                  "        CANCOM02,  " &_
                  "        VADUAN02,  " &_
                  "        VMERME02,  " &_
                  "        FACTMO01,  " &_
                  "        VMERME02 * FACTMO01 AS VALORDOLARES, " &_
                  "        PAIORI02,  " &_
                  "        RCLI01,    " &_
                           fecEntAdu&" as fechentrada, "&_
                  "        observ02, "&_
                  "        d_mer102, "&_
                  "        d_mer202  "&_
                  " FROM "&tabla&" LEFT JOIN C01REFER ON REFCIA01 = REFE01   "&_
                  "                LEFT JOIN SSFRAC02 ON REFCIA01 = REFCIA02 "&_
                  " WHERE firmae01<>''       "&_
                  "       AND fecpag01 >= '"&FormatoFechaInv(ISTRFINI)&"' "&_
                  "       AND fecpag01 <= '"&FormatoFechaInv(FSTRFFIN)&"' " & permi &_
                  " group by refcia01,FRAARN02,ordfra02 "

                  '"       AND cveped01<>'R1' "&_
'Response.Write(sqlReferencias)
'Response.End

     MM_STRING = ODBC_POR_ADUANA(Session("GAduana"))
     set RSReferencias = server.CreateObject("ADODB.Recordset")
     RSReferencias.ActiveConnection = MM_STRING
     RSReferencias.Source= sqlReferencias
     RSReferencias.CursorType = 0
     RSReferencias.CursorLocation = 2
     RSReferencias.LockType = 1
     RSReferencias.Open()

%>

 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Layout Honeywell Fracciones de <%=pTipOpe%> </p></font></strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p></p></font></strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Del <%=ISTRFINI%> al <%=FSTRFFIN%> </p></font></strong>

<table border="1">
  <tr bgcolor="#CCFFCC">
      <th> Control            </th>
      <th> Tarriff Code       </th>
      <th> Shipping point     </th>
      <th> Quantity           </th>
      <th> Part number        </th>
      <th> Total EXW          </th>
      <th> Currency           </th>
      <th> Origin             </th>
      <th> LOB                </th>
      <th> Ready for delivery </th>
      <th> Descripción        </th>
      <!-- <th> Observaciones      </th> -->
  </tr>
<%



if not RSReferencias.eof then

while not RSReferencias.eof

     ' claveclientePed = RSReferencias.fields.item("cvecli01").value
     ' sqlDatosCliente=" SELECT cvecli18,nomcli18,rfccli18,domcli18,ciucli18 "&_
     '                 " FROM ssclie18 "&_
     '                 " where cvecli18="& claveclientePed & " "
     ' 'Response.Write(sqlDatosCliente)
     ' 'Response.End

     ' set RSDatosCliente = server.CreateObject("ADODB.Recordset")
     ' RSDatosCliente.ActiveConnection = MM_STRING
     ' RSDatosCliente.Source= sqlDatosCliente
     ' RSDatosCliente.CursorType = 0
     ' RSDatosCliente.CursorLocation = 2
     ' RSDatosCliente.LockType = 1
     ' RSDatosCliente.Open()
     '      if not RSDatosCliente.eof then
     '            'pTotalFleteLocal=RSFleLocalIVARet4.fields.item("TOTALFLETELOCAL").value
     '            pnombre=RSDatosCliente.fields.item("nomcli18").value
     '            prfc=RSDatosCliente.fields.item("rfccli18").value
     '            pcd =RSDatosCliente.fields.item("ciucli18").value
     '      else
     '            'pTotalFleteLocal=0
     '            pnombre=" - - - "
     '            prfc=" - - - "
     '            pcd=" - - - "
     '      end if
     ' RSDatosCliente.close
     ' set RSDatosCliente = nothing

    'Response.End
    ' inicializo variables
    ptipcam=0
    pdta=0
    pprev=0
    padv_igi=0
    sumaT=0

    idx=idx +1

         ' MERCANCIAS
         pordenc = ""
         ItemMer = ""
         facItem = ""

         sqlordencomp= " SELECT cpro05,item05, pedi05, fact05 " & _
                       " FROM D05ARTIC " & _
                       " WHERE refe05='" & RSReferencias.fields.item("refcia01").value & "' " & _
                       " AND frac05 = '" & RSReferencias.fields.item("fraarn02").value & "' "

         'Response.Write(sqlordencomp)
         'Response.End

         set RSordencomp = server.CreateObject("ADODB.Recordset")
         RSordencomp.ActiveConnection = MM_STRING
         RSordencomp.Source= sqlordencomp
         RSordencomp.CursorType = 0
         RSordencomp.CursorLocation = 2
         RSordencomp.LockType = 1
         RSordencomp.Open()
         poc=0

         while not RSordencomp.eof
          if RSordencomp.fields.item("item05").value<>"" then

               factemp = RSordencomp.fields.item("fact05").value

                       ' facturas comerciales
                       sqlNumPartFact = " SELECT MONFAC39 " & _
                                        " FROM SSFACT39   " & _
                                        " WHERE REFCIA39 = '"&RSReferencias.fields.item("refcia01").value&"' " & _
                                        "   AND NUMFAC39 = '"&factemp&"' "
                       'Response.Write(sqlNumPartFact)
                       'Response.End

                       set RSNumPartFact = server.CreateObject("ADODB.Recordset")
                       RSNumPartFact.ActiveConnection = MM_STRING
                       RSNumPartFact.Source= sqlNumPartFact
                       RSNumPartFact.CursorType = 0
                       RSNumPartFact.CursorLocation = 2
                       RSNumPartFact.LockType = 1
                       RSNumPartFact.Open()
                       'pnpf=0
                       pNumPartFact=""

                       if not RSNumPartFact.eof then
                         'if RSNumPartFact.fields.item("NumPartFact").value<>"" and RSNumPartFact.fields.item("fact05").value<>"" then
                          'while not RSNumPartFact.eof
                             'if pnpf=0 then
                                pNumPartFact = RSNumPartFact.fields.item("MONFAC39").value
                                '&" de "&RSNumPartFact.fields.item("fact05").value
                                'pnpf=321
                             'else
                             '   pNumPartFact=pNumPartFact&" , "&RSNumPartFact.fields.item("NumPartFact").value&" de "&RSNumPartFact.fields.item("fact05").value
                             'end if
                           'RSNumPartFact.movenext()
                          'wend
                        else
                            pNumPartFact=""
                        end if
                         'else
                         '  pNumPartFact=" - - - "
                         'end if
                       RSNumPartFact.close
                       set RSNumPartFact = nothing

               if poc=0 then
                  pordenc = RSordencomp.fields.item("pedi05").value
                  ItemMer = RSordencomp.fields.item("item05").value
                  'facItem = RSordencomp.fields.item("fact05").value
                  facItem = pNumPartFact
                  poc=321
               else
                  pordenc = pordenc & " , "& RSordencomp.fields.item("pedi05").value
                  ItemMer = ItemMer & " , "& RSordencomp.fields.item("item05").value
                  'facItem = facItem & " , "& RSordencomp.fields.item("fact05").value
                  facItem = facItem & " , "& pNumPartFact
               end if
          else
              pordenc = " "
              ItemMer = " "
              facItem = " "
          end if
           RSordencomp.movenext()
         wend
          RSordencomp.close
          set RSordencomp = nothing





%>



  <tr>
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("refcia01").value%>     </th> <!-- Control  -->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("fraarn02").value%>     </th> <!-- Tarriff Code  -->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("PORT_LOADING").value%> </th> <!-- Shipping point  -->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("CANCOM02").value%>     </th> <!-- Quantity -->
      <th><font size="1" face="Arial"><%=ItemMer%>                                         </th> <!--Part number -->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("VALORDOLARES").value%> </th> <!-- Total EXW -->
      <th><font size="1" face="Arial"><%=facItem%>                                         </th> <!-- Currency -->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("PAIORI02").value%>     </th> <!-- Origin  -->
      <!-- <th><font size="1" face="Arial"><pordenc;RSReferencias.fields.item("RCLI01").value -->       </th> <!-- LOB  -->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("RCLI01").value%>       </th> <!-- LOB  -->

      <!--<th><font size="1" face="Arial"><RSReferencias.fields.item("fechentrada").value>  </th>--> <!-- Ready for delivery -->
      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("fecpag01").value%>  </th> <!-- Ready for delivery -->


      <th><font size="1" face="Arial"><%=RSReferencias.fields.item("d_mer102").value%>     </th> <!-- Descripción -->

      <!-- <th><font size="1" face="Arial">RSReferencias.fields.item("observ02").value     </th> --> <!-- Observaciones  -->
  </tr>

<%

'Response.Write("aqui")
'Response.End

  RSReferencias.movenext()

wend

RSReferencias.close
set RSReferencias = nothing


else
%>

<tr>
  <th colspan=12>
    <font size="1" face="Arial">No se Encontro ningun registro con esos parametros
  </th>
</tr>

<%

end if

%>
</table>

<%


             'Carga en Proceso
             else

                strHTML = "<table>"
                strHTML = strHTML &  "<tr bgcolor='#1B5296'>"
                strHTML = strHTML &  "      <td colspan='4' class='textForm2'><div align='right'></div></td> "
                strHTML = strHTML &  " </tr> "
                strHTML = strHTML &  " <tr>  "
                strHTML = strHTML &  "    <td colspan='4'><div align='center'></div></td> "
                strHTML = strHTML &  " </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td width='250' rowspan='4' align='center'><img src='http://rkzego.no-ip.org/PortalMySQL/Extranet/ext-Images/computadora_animo.jpg' width='150' height='157'></td> "
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=4 COLOR=red>Espere un momento...</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=5 COLOR=red>La Base de Datos se esta Actualizando</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=5 COLOR=red>Genere este Reporte unos minutos mas tarde</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=3 COLOR=red>estamos trabajando para brindarle un mejor servicio</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='4'><div align='center'></div></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr bgcolor='#1B5296'><td colspan='4'></td></tr>"
                strHTML = strHTML &  "  </table>"
                response.Write(strHTML)
            end if


 else
      select case strCodError
        case "1"
         strMenjError = "Campo en Blanco Especifique!.."
      case "5","6"
         strMenjError = "Fechas Erroneas, Verifique!"
      case "7"
         strMenjError = "Registros No Encontrados!"
      end select
%>
    <table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
      <tr>
      <td><%=strMenjError%></td>
      </tr>
    </table>
    <br>
<%
end if
%>

