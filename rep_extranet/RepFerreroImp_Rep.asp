
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% Server.ScriptTimeout=1500 %>
<%
serv="localhost"
base_datos="rku_extranet"
usu="EXTRANET"
pass="rku_admin"
base="VER"


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


    %>
<HTML>
<HEAD>
<TITLE>:: REPORTE GASTOS, IMPUESTOS Y HONORARIOS ::</TITLE>
</HEAD>
<BODY>
<%if  Session("GAduana") <> "" then %>

<%
if request.form("fecha")=0 then

      Diaf = cstr(datepart("d",now()))
      Mesf = cstr(datepart("m",now()))
      Aniof = cstr(datepart("yyyy",now()))
      FSTRFFIN = Aniof&"/"&Mesf&"/"&Diaf

      STRFINI=Dateadd("d",-8,now())
      Diai = cstr(datepart("d",STRFINI))
      Mesi = cstr(datepart("m",STRFINI))
      Anioi = cstr(datepart("yyyy",STRFINI))
      ISTRFINI = Anioi&"/"&Mesi&"/"&Diai
else
  STRFINI=request.form("FINI")
  STRFFIN=request.form("FFIN")

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
end if
'CCLIENTE=request.form("txt_cvecli")


if request.form("tipRep") = 2 then
   Response.Addheader "Content-Disposition", "attachment;"
   Response.ContentType = "application/vnd.ms-excel"
end if

'option=no se bien como funcion pero siempre va el numero 16427
'MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
MM_EXTRANET_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE="&base_datos&"; UID="&usu&"; PWD="&pass&"; OPTION=16427"

     'sqlcuentas="select refcia01,numped01,cgas31 FROM ssdagi01, e31cgast  where firmae01<>'' and  cveped01<>'R1' and fech31 >= '2008-01-01' and fech31 <= '2008-04-12' and esta31 = 'I' and cvecli01= 1928 "

    sqlcuentas=" Select refcia01,numped01,fech31, cvecli01 "&_
                " FROM  (e31cgast as e inner join d31refer as d on e.cgas31=d.cgas31) "&_
                " inner join ssdagi01 on d.refe31=refcia01 "&_
                " where firmae01<>'' and  cveped01<>'R1' and fech31>='"&ISTRFINI&"' and fech31 <='"&FSTRFFIN&"' and chon31>0 "&_
                " and esta31 = 'I' "&permi&" "&_
                "group by refcia01 "
                '
    'Response.Write(sqlcuentas)
    'Response.Write(MM_EXTRANET_STRING)

'and cvecli01="&CCLIENTE&"
     set RSRcgastos = server.CreateObject("ADODB.Recordset")
     RSRcgastos.ActiveConnection = MM_EXTRANET_STRING
     RSRcgastos.Source= sqlcuentas
     RSRcgastos.CursorType = 0
     RSRcgastos.CursorLocation = 2
     RSRcgastos.LockType = 1
     RSRcgastos.Open()
'Response.Write()
'Response.Write(sqlcuentas)
'Response.End   99CC99

  %>
<P><strong>CLAVE DE CLIENTE: <%=replace(permi,"AND cvecli01 ="," , ")%></P>
<p><strong>REPORTE GENERADO DEL <%=ucase(FormatDateTime(ISTRFINI,1))%> AL <%=ucase(FormatDateTime(FSTRFFIN,1))%> </p>
<table BORDER="0" cellspacing="2" >
  <TR bgcolor="#006699">
    <th><font color="#FFFFFF">REFERENCIA</th><th><font color="#FFFFFF">PEDIMENTO</th><th><font color="#FFFFFF">FACTURA</th>
    <th><font color="#FFFFFF">F.DESPACHO</th><th><font color="#FFFFFF">CONTENEDOR</th><th><font color="#FFFFFF">C.GASTOS</th>
    <th><font color="#FFFFFF">T.IMPUESTOS</th><!--th><font color="#FFFFFF">T.GASTOS</th-->
    <th><font color="#FFFFFF">GTOS+IVAyVALID.</th><th><font color="#FFFFFF">(-)VALI</th>
    <th><font color="#FFFFFF">GTOS+IVA</th><th><font color="#FFFFFF">GTOS-IVA</th><th>
    <font color="#FFFFFF">T.HONORARIOS</th></font>
    <th><font color="#FFFFFF">PROVEEDOR</th></font>
  </TR>
<%
'Response.end
indx=0
while not RSRcgastos.eof
indx=indx+1
if (indx mod 2) = 0then
  ccolor="#FFFFCC"
else
  ccolor="#FFFFFF"
end if

pimpuestos=0
pfdsp=""
PFACTURA=""
PCONTENEDOR=""
PCUENTA=""
pgastos=0
phonorarios=0
pfecha=""
pgastIVA=0
pvalidacion=0

referencia=RSRcgastos.fields.item("refcia01").value
're=RSRcgastos.fields.item("cuentag").value
'pfecha=RSRcgastos.fields.item("fech31").value



 ' -----------------------------------------------------------------------------------------------
 ' Pedimento
 ' -----------------------------------------------------------------------------------------------
     'slqimpuestos="select cvecli01,refcia01,cveadu01, tipcam01, fdsp01,cvepro01,cveimp36, sum( (i_dta101 + i_dta201) + import36 + (i_adv102 + i_adv202) ) as impuestos, cvemtr01,descri30, SUM(vaduan02) AS vadu,cvepro22,nompro22 "&_
     '     " from ssdagi01,c01refer,ssmtra30,sscont36,ssfrac02,ssprov22 "&_
     '     " where refcia01= '"&referencia&"'  and refcia01=refe01 and "&_
     '     "    refcia01=refcia36 and "&_
     '     "    refcia01=refcia02 and "&_
     '     "    cveimp36=15 and "&_
     '     "    cvemtr01=clavet30 and  "&_
     '     "    cvepro01=cvepro22  "&_
     '     "group by refcia01 "&_
     '     "order by refcia01"


     slqPedRef= "select refcia01, " & _
                   "    fdsp01,   " & _
                   "    nompro22  " & _
                   " from ssdagi01   " & _
                   " inner join c01refer on refcia01=refe01   " & _
                   " left join ssprov22  on cvepro01=cvepro22 " & _
                   " where refcia01= '"&referencia&"'      "

     'Response.Write(slqimpuestos)
     'Response.end

     set RSPedRef = server.CreateObject("ADODB.Recordset")
     RSPedRef.ActiveConnection = MM_EXTRANET_STRING
     RSPedRef.Source= slqPedRef
     RSPedRef.CursorType = 0
     RSPedRef.CursorLocation = 2
     RSPedRef.LockType = 1
     RSPedRef.Open()

          if not RSPedRef.eof then
                'pimpuestos = RSPedRef.fields.item("impuestos").value
                pfdsp      = RSPedRef.fields.item("fdsp01").value
                pprov      = RSPedRef.fields.item("nompro22").value
          else
                'pimpuestos=0
                pfdsp="---"
                pprov="---"
          end if

          'if isnull(RSimpuestos.fields.item("fdsp01").value) then
           ' pfdsp="--"
          'else
            'pfdsp=RSimpuestos.fields.item("fdsp01").value
          'end if
     RSPedRef.close
    set RSPedRef = nothing


 ' -----------------------------------------------------------------------------------------------
  'IMPUESTOS

 ' -----------------------------------------------------------------------------------------------
     'slqimpuestos="select cvecli01,refcia01,cveadu01, tipcam01, fdsp01,cvepro01,cveimp36, sum( (i_dta101 + i_dta201) + import36 + (i_adv102 + i_adv202) ) as impuestos, cvemtr01,descri30, SUM(vaduan02) AS vadu,cvepro22,nompro22 "&_
     '     " from ssdagi01,c01refer,ssmtra30,sscont36,ssfrac02,ssprov22 "&_
     '     " where refcia01= '"&referencia&"'  and refcia01=refe01 and "&_
     '     "    refcia01=refcia36 and "&_
     '     "    refcia01=refcia02 and "&_
     '     "    cveimp36=15 and "&_
     '     "    cvemtr01=clavet30 and  "&_
     '     "    cvepro01=cvepro22  "&_
     '     "group by refcia01 "&_
     '     "order by refcia01"

     slqimpuestos = " select refcia36, " & _
                    "        SUM(import36)  as impuestos " & _
                    " from sscont36 " & _
                    " where refcia36= '"&referencia&"' " & _
                    " group by refcia36 "

     'Response.Write(slqimpuestos)
     'Response.end

     set RSimpuestos = server.CreateObject("ADODB.Recordset")
     RSimpuestos.ActiveConnection = MM_EXTRANET_STRING
     RSimpuestos.Source= slqimpuestos
     RSimpuestos.CursorType = 0
     RSimpuestos.CursorLocation = 2
     RSimpuestos.LockType = 1
     RSimpuestos.Open()
          if not RSimpuestos.eof then
                pimpuestos=RSimpuestos.fields.item("impuestos").value
                'pfdsp=RSimpuestos.fields.item("fdsp01").value
                'pprov=RSimpuestos.fields.item("nompro22").value
          else
                pimpuestos=0
                'pfdsp="---"
                'pprov="---"
          end if
     RSimpuestos.close
    set RSimpuestos = nothing
' -----------------------------------------------------------------------------------------------




' factura comercial
slqFACTURA="select NUMFAC39,REFCIA39 FROM ssfact39 WHERE refcia39='"&referencia&"'    "&_
          "order by refcia39"
'Response.Write(slqReferencias)
'Response.end
     set RSFACTURA = server.CreateObject("ADODB.Recordset")
     RSFACTURA.ActiveConnection = MM_EXTRANET_STRING
     RSFACTURA.Source= slqFACTURA
     RSFACTURA.CursorType = 0
     RSFACTURA.CursorLocation = 2
     RSFACTURA.LockType = 1
     RSFACTURA.Open()

          IFAC=0
          while  NOT RSFACTURA.eof
              IF IFAC=0 THEN
                PFACTURA=RSFACTURA.fields.item("NUMFAC39").value
                IFAC=16545
              ELSE
                PFACTURA=PFACTURA &"  , "& RSFACTURA.fields.item("NUMFAC39").value
              END IF
          RSFACTURA.movenext()
          wend

     RSFACTURA.close
    set RSFACTURA = nothing




' CUENTA DE gastos
slqCUENTA=" select  A.cgas31 AS CG,FECH31, REFE31 "&_
          "FROM d31refer AS A JOIN  e31cgast AS B ON A.CGAS31=B.CGAS31 "&_
          " WHERE REFE31='"&referencia&"' AND fech31>='"&ISTRFINI&"' AND fech31<='"&FSTRFFIN&"' AND esta31 = 'I' and chon31>0  "&_
          " order by CG"
'Response.Write(slqReferencias)
'Response.end
     set RSCUENTA = server.CreateObject("ADODB.Recordset")
     RSCUENTA.ActiveConnection = MM_EXTRANET_STRING
     RSCUENTA.Source= slqCUENTA
     RSCUENTA.CursorType = 0
     RSCUENTA.CursorLocation = 2
     RSCUENTA.LockType = 1
     RSCUENTA.Open()

          IC=0
          while  NOT RSCUENTA.eof
              IF IC=0 THEN
                PCUENTA=RSCUENTA.fields.item("CG").value
 '               pfecha=RSCUENTA.fields.item("FECH31").value
                IC=16545
              ELSE
                PCUENTA=PCUENTA &" , "& RSCUENTA.fields.item("CG").value
 '               pfecha=pfecha &" , "& RSCUENTA.fields.item("FECH31").value
              END IF
          RSCUENTA.movenext()
          wend
          IF PCUENTA="" THEN PCUENTA="..." END IF
  '        IF pfecha="" THEN pfecha="..." END IF
     RSCUENTA.close
    set RSCUENTA = nothing

' CUENTA DE CONTENEDORES
slqCONTENEDOR="select NUMCON40,REFCIA40 FROM SSCONT40 WHERE REFCIA40='"&referencia&"'   "&_
          "order by refcia40"
'Response.Write(slqReferencias)
'Response.end
     set RSCONTENEDOR = server.CreateObject("ADODB.Recordset")
     RSCONTENEDOR.ActiveConnection = MM_EXTRANET_STRING
     RSCONTENEDOR.Source= slqCONTENEDOR
     RSCONTENEDOR.CursorType = 0
     RSCONTENEDOR.CursorLocation = 2
     RSCONTENEDOR.LockType = 1
     RSCONTENEDOR.Open()
          ICT=0
          while  NOT RSCONTENEDOR.eof
              IF ICT=0 THEN
                PCONTENEDOR=RSCONTENEDOR.fields.item("NUMCON40").value
                ICT=16545
              ELSE
                PCONTENEDOR=PCONTENEDOR &"  , "& RSCONTENEDOR.fields.item("NUMCON40").value
              END IF
          RSCONTENEDOR.movenext()
          wend

     RSCONTENEDOR.close
    set RSCONTENEDOR = nothing




'sqlgast= " SELECT A.refe21 B.CONC21, desc21, SUM(If(B.DEHA21='A', A.MONT21/(1+(B.piva21/100)) ,A.MONT21/(1+(B.piva21/100)) *-1)) as  gastos "&_
'    "FROM D21PAGHE as A, E21PAGHE as B, C21PAGHE as C "&_
'    "WHERE ( B.FOLI21 = A.FOLI21  "&_
'    "AND YEAR(B.FECH21) = YEAR(A.FECH21) and refe21='"&referencia&"'  "&_
'    "AND B.TMOV21 = A.TMOV21 )   "&_
'    "AND B.CONC21 = C.CLAV21   "&_
'    "AND B.ESTA21 <> 'S' AND B.tpag21=1 AND B.TMOV21='P' and A.TMOV21='P'  GROUP BY A.refe21  "

 sqlgast="    SELECT A.refe21, B.CONC21, SUM(If (B.DEHA21='A' ,  A.MONT21 / ( 1+(B.piva21/100) ) , ( A.MONT21 / (1+(B.piva21/100) )  *-1)  ) )  as  gastos, "&_
         "   SUM(If (B.DEHA21='A',A.MONT21,((A.MONT21)*-1)  ) )  as  gastIVA,  "&_
   " SUM(If (B.DEHA21='A' and B.CONC21=218,A.MONT21,IF (B.DEHA21='C' and B.CONC21=218, ((A.MONT21)*-1) ,0) ))  as  validacion  "&_
         "   FROM D21PAGHE as A, E21PAGHE as B, C21PAGHE as C  "&_
         "   WHERE ( B.FOLI21 = A.FOLI21  AND year(B.FECH21) = year(A.FECH21) AND B.TMOV21 = A.TMOV21 )  "&_
         "           AND B.CONC21 = C.CLAV21  "&_
         "           AND B.ESTA21<> 'S' "&_
         "           AND B.tpag21=1 "&_
         "           AND B.TMOV21='P' "&_
         "           AND A.TMOV21='P' "&_
         "           AND A.refe21='"&referencia&"' and A.CGAS21='"&PCUENTA&"' "&_
        "GROUP BY A.refe21"

'Response.Write(sqlgast)
'Response.End
     set RSgast = server.CreateObject("ADODB.Recordset")
     RSgast.ActiveConnection = MM_EXTRANET_STRING
     RSgast.Source= sqlgast
     RSgast.CursorType = 0
     RSgast.CursorLocation = 2
     RSgast.LockType = 1
     RSgast.Open()

          if not RSgast.eof then
                pgastos=RSgast.fields.item("gastos").value
                pgastIVA=RSgast.fields.item("gastIVA").value
                pvalidacion=RSgast.fields.item("validacion").value
          else
                pgastos=0
                pgastIVA=0
                pvalidacion=0
          end if


     RSgast.close
    set RSgast = nothing


'

sqlHonoYserv="select refe31, fech31 as fcg, "&_
                  " (e31cgast.cgas31) as cg,fech31, sum(if (esta31='I',(chon31+caho31), 0 ) ) as honorarios, "&_
                  " sum( if (esta31='I', (csce31), 0 )) as servicios, "&_
                  " sum(if(esta31='I', (chon31+caho31+csce31), 0)) as ingresos,"&_
                  " sum(if (esta31='I',(suph31), 0)) as o_cargos "&_
                  "from d31refer join e31cgast on d31refer.cgas31=e31cgast.cgas31 "&_
                  "where d31refer.refe31='"&referencia&"' "&_
                  " AND fech31>='"&ISTRFINI&"' AND FECH31<='"&FSTRFFIN&"' and chon31>0   "&_
                  "group by refe31 "&_
                  "order by refe31"
'Response.Write(sqlHonoYserv)

     set RSHonoYserv = server.CreateObject("ADODB.Recordset")
     RSHonoYserv.ActiveConnection = MM_EXTRANET_STRING
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

'RSRcgastos.movenext()
  %>
  <TR bgcolor="<%=ccolor%>">
    <th><font color="#000000" size="2"><%=referencia%></th>
    <th><font color="#000000" size="2"><%=RSRcgastos.fields.item("numped01").value%></th>
    <th><font color="#000000" size="2"><%=PFACTURA%></th>
    <th><font color="#000000" size="2"><%=pfdsp%></th>
    <th><font color="#000000" size="2"><%=PCONTENEDOR%></th>
    <th><font color="#000000" size="2"><%=PCUENTA%></th>
    <th><font color="#000000" size="2"><%=replace(pimpuestos,",",".")%></th>
    <!--th><font color="#000000" size="2"><'%'=replace(pgastos,",",".")%></th-->
    <th><font color="#000000" size="2"><%=replace(pgastIVA,",",".")%></th>
    <th><font color="#000000" size="2"><%=replace(pvalidacion,",",".")%></th>
    <th><font color="#000000" size="2"><%=replace(pgastIVA-pvalidacion,",",".")%></th>
    <th><font color="#000000" size="2"><%=replace( ROUND(((pgastIVA-pvalidacion)/1.15),2 ),",",".")%></th>
    <th><font color="#000000" size="2"><%=replace(phonorarios,",",".")%></th>
    <th><font color="#000000" size="2"><%=replace(pprov,",",".")%></th>
  </TR>
<%
RSRcgastos.movenext()
wend

RSRcgastos.close
set RSRcgastos = nothing

'////////////////////////// cuenta de gastos COMPLEMENTARIAs
%>
</table>
  <br></br>
<%    sqlcuentasCOMP=" Select refcia01,numped01,fech31, cvecli01, nompro01 "&_
                " FROM  (e31cgast as e inner join d31refer as d on e.cgas31=d.cgas31) "&_
                " inner join ssdagi01 on d.refe31=refcia01 "&_
                " where firmae01<>'' and  cveped01<>'R1' and fech31>='"&ISTRFINI&"' and fech31 <='"&FSTRFFIN&"' "&_
                " and esta31 = 'I' and chon31<=0 "&permi&" "&_
                "group by refcia01 "
                '
'    Response.Write(sqlcuentasCOMP)
'    Response.Write(MM_EXTRANET_STRING)

'and cvecli01="&CCLIENTE&"
     set RSRcgastosCOMP = server.CreateObject("ADODB.Recordset")
     RSRcgastosCOMP.ActiveConnection = MM_EXTRANET_STRING
     RSRcgastosCOMP.Source= sqlcuentasCOMP
     RSRcgastosCOMP.CursorType = 0
     RSRcgastosCOMP.CursorLocation = 2
     RSRcgastosCOMP.LockType = 1
     RSRcgastosCOMP.Open()

indx=0


while not RSRcgastosCOMP.eof
indx=indx+1

if (indx mod 2) = 0then
  ccolor="#FFFFCC"
else
  ccolor="#FFFFFF"
end if
if indx=1 then
%>
<table BORDER="0" cellspacing="2" >
  <tr bgcolor="#006699">
      <TH><font color="#FFFFFF">REFERENCIA</TH>
      <TH><font color="#FFFFFF"> - </TH>
      <TH><font color="#FFFFFF"> - </TH>
      <TH><font color="#FFFFFF"> - </TH>
      <TH><font color="#FFFFFF"> - </TH>
      <TH><font color="#FFFFFF">C.GATOS</TH>
      <TH><font color="#FFFFFF"> - </TH>
      <!--TH><font color="#FFFFFF">GASTOS</TH-->
      <th><font color="#FFFFFF">GTOS+IVAyVALID.</th><th><font color="#FFFFFF">(-)VALI</th>
      <th><font color="#FFFFFF">GTOS+IVA</th><th><font color="#FFFFFF">GTOS-IVA</th>
      <TH><font color="#FFFFFF"> - </TH>
      <TH><font color="#FFFFFF">NOM.PROV.</TH>
     </tr>
<%
end if

pimpuestos=0
pfdsp=""
PFACTURA=""
PCONTENEDOR=""
PCUENTACOMP=""
pgastosCOMP=0
phonorarios=0
pfecha=""
pgastIVACOMP=0
pvalidacionCOMP=0
      '//////////////////////////////////////////////////////////////
      ' CUENTA DE GASTOS COMPLEMENTARIA
      slqCUENTACOMP=" select  A.cgas31 AS CGCOMP,FECH31, REFE31 "&_
                "FROM d31refer AS A JOIN  e31cgast AS B ON A.CGAS31=B.CGAS31 "&_
                " WHERE REFE31='"&RSRcgastosCOMP.fields.item("REFCIA01").value&"' AND fech31>='"&ISTRFINI&"' AND fech31<='"&FSTRFFIN&"' AND esta31 = 'I' and chon31<=0  "&_
                " order by CGCOMP"
'      Response.Write(slqCUENTACOMP)
'      Response.end
           set RSCUENTACOMP = server.CreateObject("ADODB.Recordset")
           RSCUENTACOMP.ActiveConnection = MM_EXTRANET_STRING
           RSCUENTACOMP.Source= slqCUENTACOMP
           RSCUENTACOMP.CursorType = 0
           RSCUENTACOMP.CursorLocation = 2
           RSCUENTACOMP.LockType = 1
           RSCUENTACOMP.Open()

                IC=0
                PCUENTACOMP=""
                SQLCUENTACOMP=""
                while  NOT RSCUENTACOMP.eof
                    IF IC=0 THEN
                      PCUENTACOMP=RSCUENTACOMP.fields.item("CGCOMP").value
                      SQLCUENTACOMP=RSCUENTACOMP.fields.item("CGCOMP").value
       '               pfecha=RSCUENTA.fields.item("FECH31").value
                      IC=16545
                    ELSE
                      PCUENTACOMP=PCUENTACOMP &" , "& RSCUENTACOMP.fields.item("CGCOMP").value
                      SQLCUENTACOMP=SQLCUENTACOMP &" or A.CGAS21='"& RSCUENTACOMP.fields.item("CGCOMP").value&"'"
       '               pfecha=pfecha &" , "& RSCUENTA.fields.item("FECH31").value
                    END IF
                RSCUENTACOMP.movenext()
                wend
                IF PCUENTACOMP="" THEN PCUENTACOMP="..." END IF
                'IF SQLCUENTACOMP="" THEN PCUENTACOMP="..." END IF
        '        IF pfecha="" THEN pfecha="..." END IF
           RSCUENTACOMP.close
          set RSCUENTACOMP = nothing


      ' GATOS DE LA CUENTA DE GASTOS COMPLEMENTARIA
       sqlgastCOMP="    SELECT A.refe21, B.CONC21, SUM(If (B.DEHA21='A' ,  A.MONT21 / ( 1+(B.piva21/100) ) , ( A.MONT21 / (1+(B.piva21/100) )  *-1)  ) )  as  gastos, "&_
               "   SUM(If (B.DEHA21='A',A.MONT21,((A.MONT21)*-1)  ) )  as  gastIVA,  "&_
 "   SUM(If (B.DEHA21='A' and B.CONC21=218,A.MONT21,IF (B.DEHA21='C' and B.CONC21=218, ((A.MONT21)*-1) ,0) ))  as  validacion  "&_
               "   FROM D21PAGHE as A, E21PAGHE as B, C21PAGHE as C  "&_
               "   WHERE ( B.FOLI21 = A.FOLI21  AND year(B.FECH21) = year(A.FECH21) AND B.TMOV21 = A.TMOV21 )  "&_
               "           AND B.CONC21 = C.CLAV21  "&_
               "           AND B.ESTA21<> 'S' "&_
               "           AND B.tpag21=1 "&_
               "           AND B.TMOV21='P' "&_
               "           AND A.TMOV21='P' "&_
               "           AND A.refe21='"&RSRcgastosCOMP.fields.item("REFCIA01").value&"' and A.CGAS21='"&SQLCUENTACOMP&"' "&_
              "GROUP BY A.refe21"

      'Response.Write(sqlgastCOMP)
      'Response.Write(sqlgast)
      'Response.End
           set RSgastCOMP = server.CreateObject("ADODB.Recordset")
           RSgastCOMP.ActiveConnection = MM_EXTRANET_STRING
           RSgastCOMP.Source= sqlgastCOMP
           RSgastCOMP.CursorType = 0
           RSgastCOMP.CursorLocation = 2
           RSgastCOMP.LockType = 1
           RSgastCOMP.Open()

                if not RSgastCOMP.eof then
                      pgastosCOMP=RSgastCOMP.fields.item("gastos").value
                      pgastIVACOMP=RSgastCOMP.fields.item("gastIVA").value
                      pvalidacionCOMP=RSgastCOMP.fields.item("validacion").value
                else
                      pgastosCOMP=0
                      pgastIVACOMP=0
                      pvalidacionCOMP=0
                end if


           RSgastCOMP.close
          set RSgastCOMP = nothing

      ''//////////////////////////////////////////////////////////////

      %>
     <tr bgcolor="<%=ccolor%>">
          <th><font color="#000000" size="2"><%=RSRcgastosCOMP.fields.item("REFCIA01").value%></th>
	<th><font color="#000000" size="2"><%=" - "%></th>
	<th><font color="#000000" size="2"><%=" - "%></th>
	<th><font color="#000000" size="2"><%=" - "%></th>
	<th><font color="#000000" size="2"><%=" - "%></th>
	<th><font color="#000000" size="2"><%=replace(PCUENTACOMP,",",".")%></th>
	<th><font color="#000000" size="2"><%=" - "%></th>
          <!--th><font color="#000000" size="2"><'%'=replace(pgastosCOMP,",",".")%></th-->
	<th><font color="#000000" size="2"><%=replace(pgastIVACOMP,",",".")%></th>
          <th><font color="#000000" size="2"><%=replace(pvalidacionCOMP,",",".")%></th>
          <th><font color="#000000" size="2"><%=replace(pgastIVACOMP-pvalidacionCOMP,",",".")%></th>
          <th><font color="#000000" size="2"><%=replace(ROUND(((pgastIVACOMP-pvalidacionCOMP)/1.15),2) ,",",".")%></th>
	<th><font color="#000000" size="2"><%=" - "%></th>
        <th><font color="#000000" size="2"><%=RSRcgastosCOMP.fields.item("nompro01").value%></th>
     </tr>
          <%
      RSRcgastosCOMP.movenext()
      wend

      RSRcgastosCOMP.close
      set RSRcgastosCOMP = nothing
%>
</table>
<%else
  response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>")
end if%>
</BODY>
</HTML>