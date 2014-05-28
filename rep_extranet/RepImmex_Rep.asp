<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

<%
'serv="10.66.1.5"
'base_datos="rku_extranet"
'usu="carlosmg"
'pass="123456"
'base="VER"
'tipRep=1

'c_man=0
'c_fle=0

STRCVECLI=Request.Form("txtCliente")
STRFINI=Request.Form("FINI")
STRFFIN=Request.Form("FFIN")
strSeltipo=Request.Form("seltipo")
strpimmex=Request.Form("immex")
strTipRep=Request.Form("tipRep")
c_man=Request.Form("c_man")
c_fle=Request.Form("c_fle")

 MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
 'MM_EXTRANET_STRING = ODBC_POR_ADUANA("VER")
 MM_EXTRANET_STRING_VER = ODBC_POR_ADUANA("VER")

if strTipRep = 2 then
   Response.Addheader "Content-Disposition", "attachment;"
   Response.ContentType = "application/vnd.ms-excel"
end if

'strSeltipo="i"

SRTADUAN=Request.Form("Aduana")
SRTA=Request.Form("vddAduana")

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

'Response.Write("STRADUAN")
'Response.Write(STRADUAN)
'Response.Write("stra")
'Response.Write(STRA)
'Response.Write("client")
'Response.Write(STRCVECLI)
'Response.Write("PERMI")
'Response.Write(PERMI)


  'Select Case (STRADUAN)
   'Case "VER":
    '   c_fle="15"
     '  c_man="2"
   'Case "MEX":
    '   c_fle="7"
     '  c_man="127"
   'Case "MAN":
    '  c_fle="5"
     '  c_man="2"
   'Case "TAM":
    '  c_fle="15"
     '  c_man="2"
   'Case "LZR":
    '   c_fle="15"
     '  c_man="2"
   'Case "GUA":
    '   c_fle="15"
      ' c_man="2"
   'End Select


if  Session("GAduana") <> "" then

    if strSeltipo="i" then

        num_con_f_m = c_man
        tabla="ssdagi01"
        STRECNABEZADO="<th bgcolor=""#DDDDDD"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >indx</th>"&_
                    "<th bgcolor=""#DDDDDD"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >  - Referencia -</th>"&_
                    "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" > Aduana</th>"&_
                    "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" > Fech Emb</th>"&_
                    "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" > Fech Recp SLP</th>"&_
                    "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" > Tip Trans</th>"&_
                    "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" > Cant Ped</th>"&_
                    "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Val Aduana(USD)</th>"&_
                    "<th bgcolor=""#CCFFCC"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Tip Camb</th>"&_
                    "<th bgcolor=""#CCFFCC"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >DTA</th>"&_
                    "<th bgcolor=""#CCFFCC"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Prev</th>"&_
                    "<th bgcolor=""#CCFFCC"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >IGI</th>"&_
                    "<th bgcolor=""#FFFF66"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Hono LDO.</th>"&_
                    "<th bgcolor=""#FFFF66"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >OverTime LDO</th>"&_
                    "<th bgcolor=""#FFFF66"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >OtherExpenses LDO</th>"&_
                    "<th bgcolor=""#FFFF66"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Importe Transfer</th>"&_
                    "<th bgcolor=""#99CCFF"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Hono's AA</th>"&_
                    "<th bgcolor=""#99CCFF"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >ServComp AA</th>"&_
                    "<th bgcolor=""#99CCFF"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Otros AA</th>"&_
                    "<th bgcolor=""#99CCFF"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Maniobras p. Mex</th>"&_
                    "<th bgcolor=""#66CC00"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Flete SLP</th>"&_
                    "<th bgcolor=""#FFCC33"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Total Gtos</th>"

        cLDO="<th align=""center""><font size=""1"" face=""Arial"">No Disp.</th><th align=""center""><font size=""1"" face=""Arial"">No Disp.</th><th align=""center""><font size=""1"" face=""Arial"">No Disp.</th><th align=""center""><font size=""1"" face=""Arial""> No Disp.</th>"

        if strpimmex=1 then
            Titulo=" IMPORTACION IMMEX"

            sqlimmex=" and cveped01='A2' "
            c_fm=""
        else
            Titulo=" IMPORTACION REGULAR "
            sqlimmex=" and cveped01<>'A2' "
        end if
    else
        Titulo=" EXPORTACION "
        tabla="ssdage01"
        num_con_f_m = c_fle
        cLDO="<th align=""center""><font size=""1"" face=""Arial"">No Disp.</th><th align=""center""><font size=""1"" face=""Arial"">No Disp.</th><th align=""center""><font size=""1"" face=""Arial"">No Disp.</th>"

        STRECNABEZADO="<th bgcolor=""#DDDDDD""  width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >indx</th>"&_
        "<th bgcolor=""#DDDDDD"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" > - Referencia -</th>"&_
        "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Aduana</th>"&_
        "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Fech Emb</th>"&_
        "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Fech Recp SLP</th>"&_
        "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Tip Trans</th>"&_
        "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Cant Ped</th>"&_
        "<th bgcolor=""#FFFF99"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Val Aduana(USD)</th>"&_
        "<th bgcolor=""#CCFFCC"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Tip Camb</th>"&_
        "<th bgcolor=""#CCFFCC"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >DTA</th>"&_
        "<th bgcolor=""#CCFFCC"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Prev</th>"&_
        "<th bgcolor=""#FFFF66"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Hono LDO.</th>"&_
        "<th bgcolor=""#FFFF66"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >OverTime LDO</th>"&_
        "<th bgcolor=""#FFFF66"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >OtherExpenses LDO</th>"&_
        "<th bgcolor=""#99CCFF"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Hono's AA</th>"&_
        "<th bgcolor=""#99CCFF"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >ServComp AA</th>"&_
        "<th bgcolor=""#99CCFF"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Otros AA</th>"&_
           "<th bgcolor=""#66CC00"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Flete X Expo Dest.Fin.</th>"&_
        "<th bgcolor=""#66CC00"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Flete X Expo Laredo </th>"&_
        "<th bgcolor=""#FFCC33""  width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Total Gtos</th>"

        '"<th bgcolor=""#99CCFF"" width=""100"" nowrap><strong><font size=""2"" face=""Arial, Helvetica, sans-serif"" >Maniobras p. Mex</th>"&_
    end if

'Response.Write(num_con_f_m)

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

'Response.Write(MM_EXTRANET_STRING)
'RESPONSE.END

    '"DRIVER={MySQL ODBC 3.51 Driver}; SERVER="&serv&"; DATABASE="&base_datos&"; UID="&usu&"; PWD="&pass&"; OPTION=16427"

    conentionRefe = MM_EXTRANET_STRING

    slqReferencias="select cvecli01,refcia01,cveadu01, tipcam01,(i_dta101 + i_dta201) as dta,cvemtr01,descri30, "&_
              " fdsp01,cveimp36,import36 as prev,sum(i_adv102 + i_adv202) as adv_igi, SUM(vaduan02) AS vadu "&_
              " from "&tabla&",c01refer,ssmtra30,sscont36,ssfrac02 "&_
              " where firmae01<>'' and  cveped01<>'R1' "&sqlimmex&" and "&_
              "     fecpag01>='"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"' "&_
              "   " &permi& "    and  "&_
              "    refcia01=refe01 and "&_
              "    refcia01=refcia36 and "&_
              "    refcia01=refcia02 and "&_
              "    cveimp36=15 and "&_
              "    cvemtr01=clavet30  "&_
              "group by refcia01 "&_
              "order by refcia01"
    'Response.Write(slqReferencias)
    'Response.end
         set RSReferencias = server.CreateObject("ADODB.Recordset")
         RSReferencias.ActiveConnection = conentionRefe
         RSReferencias.Source= slqReferencias
         RSReferencias.CursorType = 0
         RSReferencias.CursorLocation = 2
         RSReferencias.LockType = 1
         RSReferencias.Open()

    idx= 0


    %>
    <title><%=Titulo%></title>
      <p><strong><font size="2" face="Arial, Helvetica, sans-serif"><%=Titulo%></p>
      <p><strong><font size="2" face="Arial, Helvetica, sans-serif"> Fecha Inicial: <%=ISTRFINI%>   Fecha final: <%=FSTRFFIN%>
      <%'=STRCVECLI%>

    <table align="center"  border="1">
      <tr><strong><font size="2" face="Arial, Helvetica, sans-serif"><%=STRECNABEZADO%></tr>
    <%
    while not RSReferencias.eof
      idx=idx +1
      xrefer=RSReferencias.fields.item("refcia01").value
        ptipcam=0
        phonorarios=0
        pservicios=0
        po_cargos=0
        pmaniobras=0
        pfacmo=0
        'Response.Write("xrefer")
        'Response.end

        sqlHonoYserv="select refe31, fech31 as fcg, "&_
                      " (e31cgast.cgas31) as cg,fech31, sum(if (esta31='I',(chon31+caho31), 0 ) ) as honorarios, "&_
                      " sum( if (esta31='I', (csce31), 0 )) as servicios, "&_
                      " sum(if(esta31='I', (chon31+caho31+csce31), 0)) as ingresos,"&_
                      " sum(if (esta31='I',(suph31), 0)) as o_cargos "&_
                      "from d31refer join e31cgast on d31refer.cgas31=e31cgast.cgas31 "&_
                      "where d31refer.refe31='"&xrefer&"' "&_
                      "group by refe31 "&_
                      "order by refe31"
    'Response.Write(sqlHonoYserv)

         set RSHonoYserv = server.CreateObject("ADODB.Recordset")
         RSHonoYserv.ActiveConnection = conentionRefe
         RSHonoYserv.Source= sqlHonoYserv
         RSHonoYserv.CursorType = 0
         RSHonoYserv.CursorLocation = 2
         RSHonoYserv.LockType = 1
         RSHonoYserv.Open()
           if not RSHonoYserv.eof then
              if RSHonoYserv.fields.item("honorarios").value <>"" then
                    phonorarios=RSHonoYserv.fields.item("honorarios").value
              else
                    phonorarios=0
              end if
              if RSHonoYserv.fields.item("servicios").value <>"" then
                    pservicios=RSHonoYserv.fields.item("servicios").value
              else
                    pservicios=0
              end if
              if RSHonoYserv.fields.item("o_cargos").value <>"" then
                    po_cargos=RSHonoYserv.fields.item("o_cargos").value
              else
                    po_cargos=0
              end if
            else
                phonorarios=0
                po_cargos=0
                pservicios=0
            end if
         RSHonoYserv.close
        set RSHonoYserv = nothing

        sqlManiobras=" select  sum(if(d.esta21=0,d.mont21,0)) as maniobras_flete "&_
                      " from d21paghe as d join e21paghe as e on  (e.foli21=d.foli21 and e.fech21=d.fech21 and  e.conc21="&num_con_f_m&" )  "&_
                      " where d.refe21= '"&xrefer&"' "&_
                      " group by refe21 "&_
                      " order by refe21 "
    'Response.Write(sqlManiobras)
    'Response.End
         set RSManiobras = server.CreateObject("ADODB.Recordset")
         RSManiobras.ActiveConnection = conentionRefe
         RSManiobras.Source= sqlManiobras
         RSManiobras.CursorType = 0
         RSManiobras.CursorLocation = 2
         RSManiobras.LockType = 1
         RSManiobras.Open()

              if not RSManiobras.eof then
                    pmaniobras_flete=RSManiobras.fields.item("maniobras_flete").value
              else
                    pmaniobras_flete=0
              end if


         RSManiobras.close
        set RSManiobras = nothing

    ptipcam=RSReferencias.fields.item("tipcam01").value
    pdta=RSReferencias.fields.item("dta").value
    pprev=RSReferencias.fields.item("prev").value
    padv_igi=RSReferencias.fields.item("adv_igi").value

    sumaT=(pdta/ptipcam) + (pprev/ptipcam) + (padv_igi/ptipcam) + (phonorarios/ptipcam) + (pservicios/ptipcam) + (po_cargos/ptipcam) + (pmaniobras_flete/ptipcam)

    'FORMATNUMBER(preciomaximo,0,-1,0,-1)     = formatnumber(registros.Fields("Cantidad").Value,2) %

    %>
      <tr>
        <td align="center"><font size="1" face="Arial"><%=idx%></td>
        <td align="center"><font size="1" face="Arial"><%=RSReferencias.fields.item("refcia01").value%></td>
        <td align="center"><font size="1" face="Arial"><%=RSReferencias.fields.item("cveadu01").value%></td>
        <td align="center"><font size="1" face="Arial"><%=RSReferencias.fields.item("fdsp01").value%></td>
        <td align="center"><font size="1" face="Arial">No Disp.</td>
        <td align="center"><font size="1" face="Arial"><%=RSReferencias.fields.item("descri30").value%></td>
        <td align="center"><font size="1" face="Arial">1</td>
        <td align="center"><font size="1" face="Arial"><%=replace(RSReferencias.fields.item("vadu").value,",",".")%></td>
        <td align="center"><font size="1" face="Arial"><%=replace(ptipcam,",",".")%></td>
        <td align="center"><font size="1" face="Arial"><%=Round(replace(((pdta)/ptipcam),",","."),2)%></td>
        <td align="center"><font size="1" face="Arial"><%=Round(replace(((pprev)/ptipcam),",","."),2)%></td>
       <% if strSeltipo="i" then %>
          <td align="center"><font size="1" face="Arial"><%=Round(replace(((padv_igi)/ptipcam),",","."),2)%></td>
       <% end if %>
        <%=cLDO%>
        <td align="center"><font size="1" face="Arial"><%=Round(replace((phonorarios/ptipcam),",","."),2)%></td>
        <td align="center"><font size="1" face="Arial"><%=Round(replace((pservicios/ptipcam),",","."),2)%></td>
        <td align="center"><font size="1" face="Arial"><%=Round(replace((po_cargos/ptipcam),",","."),2)%></td>
       <%' if strSeltipo="i" then %>
          <td align="center"><font size="1" face="Arial"><%=Round(replace((pmaniobras_flete/ptipcam),",","."),2)%></td>
       <%' end if %>
        <td align="center"><font size="1" face="Arial">No Disp.</td>
        <td align="center"><font size="1" face="Arial"><%=Round(replace(sumaT,",","."),2)%></td>
      </tr>

    <%
      RSReferencias.movenext()
    wend

    RSReferencias.close
    set RSReferencias = nothing
    %>
    </TABLE>
    <%
    if idx=0 then
        %>
        <P><strong><font size="2" face="Arial, Helvetica, sans-serif">NO SE REGISTRARON OPERACIONES DE <%=TITULO%> DURANTE EL PERIODO...</P>
        <%
    end if

    %>

<%else
  response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>")
end if%>