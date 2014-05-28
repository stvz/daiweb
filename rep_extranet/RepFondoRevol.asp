<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
Server.ScriptTimeOut=100000

 Response.Buffer = TRUE
 Response.Addheader "Content-Disposition", "attachment;"
 Response.ContentType = "application/vnd.ms-excel"


 strHTML = ""

 strDate=trim(request.Form("txtDateFin"))


 if not strDate="" then


   tmpDiaFin = cstr(datepart("d",strDate))
   tmpMesFin = cstr(datepart("m",strDate))
   tmpAnioFin = cstr(datepart("yyyy",strDate))
   strDateFin = tmpAnioFin & "/" &tmpMesFin & "/"& tmpDiaFin


   dim con,Rsio,Rsio2


strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"ccli11")

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
   permi = " AND ccli11 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
end if



permi2 = PermisoClientes(Session("GAduana"),strPermisos,"clie01")

if not permi2 = "" then
  permi2 = "  and (" & permi2 & ") "
end if

if blnAplicaFiltro then
   permi2 = " AND clie01 =" & strFiltroCliente
end if

if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi2 = ""
end if


'Response.Write("POR EL MOMENTO NO PUEDE ACCESAR A ESTE REPORTE")

   set Rsio = server.CreateObject("ADODB.Recordset")
   Rsio.ActiveConnection = MM_EXTRANET_STRING

    ' strSQL = "select ccli11 as cliente,refe11 as Factura,sum(if(conc11='FA1'or conc11='SCA' or conc11='DEV' or conc11='CAR' or conc11='CF2' ,mont11,if(conc11='LIQ' or conc11='CF1' or conc11='SCR' or conc11='ABO' or conc11='BOH' or conc11='FA2',mont11 * -1,0))) as Saldo from d11movim,e11movim where d11movim.fech11<='"&strDateFin&"' and cont11<>'C' and e11movim.foli11=d11movim.foli11 and (ccli11=2438 or ccli11=2214 or ccli11=2219 or ccli11=2222 or ccli11=2291 or ccli11=2351 or ccli11=2566) group by refe11 HAVING  round(sum( if(conc11='FA1'or conc11='SCA' or conc11='DEV' or conc11='CAR' or conc11='CF2' ,mont11,if(conc11='LIQ' or conc11='CF1' or conc11='SCR' or conc11='ABO' or conc11='BOH' or conc11='FA2',mont11 * -1,0)))) <> 0 "
   strSQL = "select ccli11 as cliente,refe11 as Factura,sum(if(conc11='FA1'or conc11='SCA' or conc11='DEV' or conc11='CAR' or conc11='CF2' ,mont11,if(conc11='LIQ' or conc11='CF1' or conc11='SCR' or conc11='ABO' or conc11='BOH' or conc11='FA2',mont11 * -1,0))) as saldo  from d11movim left join e11movim on e11movim.foli11=d11movim.foli11  where d11movim.fech11<='" & strDateFin & "' and cont11<>'C'  " &  permi & " group by refe11 HAVING  round(sum( if(conc11='FA1'or conc11='SCA' or conc11='DEV' or conc11='CAR' or conc11='CF2' ,mont11,if(conc11='LIQ' or conc11='CF1' or conc11='SCR' or conc11='ABO' or conc11='BOH' or conc11='FA2',mont11 * -1,0)))) > 0 "
   Rsio.Source= strSQL
   Rsio.CursorType = 0
   Rsio.CursorLocation = 2
   Rsio.LockType = 1
   Rsio.Open()

%>
	<p><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>REPORTE DE FONDO REVOLVENTE AL </strong></font> </p>
  <table border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
	 <tr align="center" bordercolor="#000000" bgcolor="#006699">
   <td width="100" nowrap bgcolor="#0097DF"><div align="center"><font size="2"><strong><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Tipo OPeracion</font></strong></font></div></td>
	 <td width="100" nowrap bgcolor="#0097DF"><div align="center"><font size="2"><strong><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">No. De factura</font></strong></font></div></td>
   <td width="100" nowrap bgcolor="#0097DF"><div align="center"><font size="2"><strong><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Fecha Factura</font></strong></font></div></td>
	 <td width="90" nowrap bgcolor="#0097DF"><div align="center"><font size="2"><strong><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Saldo</font></strong></font></div></td>
   <td width="90" nowrap bgcolor="#0097DF"><div align="center"><font size="2"><strong><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Referencia </font></strong></font></div></td>
   <td width="90" nowrap bgcolor="#0097DF"><div align="center"><font size="2"><strong><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Cuenta de gastos</font></strong></font></div></td>
   <td width="90" nowrap bgcolor="#0097DF"><div align="center"><font size="2"><strong><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Suma</font></strong></font></div></td>
   <td width="220" nowrap bgcolor="#0097DF"><div align="center"><font size="2"><strong><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Concepto</font></strong></font></div></td>
   <td width="110" nowrap bgcolor="#0097DF"><div align="center"><font size="2"><strong><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Total</font></strong></font></div></td>
   </tr>
<%

dblSumaFondo = 2000000
'IMPRIMIMOS EL TOTAL DEL FONDO

%>

        <tr>
         <td colspan="8" nowrap bordercolor="#000000" bgcolor="#00FFFF"><font size="3" face="Arial, Helvetica, sans-serif"><strong>MONTO BASE DEL FONDO</strong></font></td>
         <td width="90" nowrap bordercolor="#000000"><font size="3" face="Arial, Helvetica, sans-serif"><strong>
           <% = formatnumber(dblSumaFondo,2)%> </strong>
         </font> </td>
    </tr>

<%

  dblSumaCuentasSaldoMayor0 = 0
  strfactura = ""

  While NOT Rsio.EOF
     set Rsio2 = server.CreateObject("ADODB.Recordset")
     Rsio2.ActiveConnection = MM_EXTRANET_STRING
   '  strSQL2 = "select refe21 as Referencia,cgas21 as Cgastos,e21paghe.conc21,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Suma FROM d21paghe LEFT JOIN  e21paghe ON YEAR(d21paghe.fech21) =  YEAR(e21paghe.FECH21) AND d21paghe.foli21 =  e21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 WHERE e21paghe.esta21 <> 'S' and e21paghe.fech21<='"&strDateFin&"' and cgas21='" &trim(Rsio.Fields.Item("Factura").Value) & "' and e21paghe.tpag21 <> 3 group by cgas21,conc21,refe21"
     strSQL2 = "select refe21 as Referencia,cgas21 as Cgastos,e21paghe.conc21,desc21,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Suma FROM d21paghe, e21paghe,c21paghe WHERE YEAR(d21paghe.fech21) =  YEAR(e21paghe.FECH21) AND d21paghe.foli21 =  e21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 and e21paghe.conc21 = c21paghe.clav21 and e21paghe.esta21 <> 'S' and e21paghe.fech21<='"&strDateFin&"' and cgas21='" &trim(Rsio.Fields.Item("Factura").Value) & "' and e21paghe.tpag21 <> 3 group by cgas21,conc21,refe21"

     if Session("GAduana") = "LAR" THEN
       strSQL2 = "select refe21 as Referencia,cgas21 as Cgastos,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Suma FROM d21paghe LEFT JOIN  e21paghe ON YEAR(d21paghe.fech21) =  YEAR(e21paghe.FECH21) AND d21paghe.foli21 =  e21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 where d21paghe.fech21<='"&strDateFin&"' and trim(cgas21)=trim('" &Rsio.Fields.Item("Factura").Value & "')  group by refe21,cgas21"
     end if
     Rsio2.Source= strSQL2
     Rsio2.CursorType = 0
     Rsio2.CursorLocation = 2
     Rsio2.LockType = 1
     Rsio2.Open()

   'Saldo=0
   'Saldo=cdbl(Rsio.Fields.Item("Saldo").Value)

    strFecha = ""
    strTipo = ""
    if not Rsio2.EOF then
       While NOT Rsio2.EOF
         Suma=0
         Suma=cdbl(Rsio2.Fields.Item("Suma").Value)
         set Rsio3 = server.CreateObject("ADODB.Recordset")
         Rsio3.ActiveConnection = MM_EXTRANET_STRING
         strTipo = ""
         strFecha = ""
         strSQL3 = "select c01refer.tipo01,e31cgast.fech31 from c01refer inner join  d31refer,e31cgast on (c01refer.refe01 = d31refer.refe31 and d31refer.cgas31 = e31cgast.cgas31)  where c01refer.refe01 ='" & Rsio2.Fields.Item("Referencia").Value & "' and d31refer.cgas31 ='" &  Rsio.Fields.Item("Factura").Value & "' "
         Rsio3.Source= strSQL3
         Rsio3.CursorType = 0
         Rsio3.CursorLocation = 2
         Rsio3.LockType = 1
         Rsio3.Open()
        if not Rsio3.EOF then
          strFecha = Rsio3.Fields.Item("fech31").Value
         if Rsio3.Fields.Item("tipo01").Value = "1" then
         strTipo = "IMPORTACION"
          else
         strTipo = ""
        end if
       end if
       Rsio3.close
       set Rsio3 = Nothing

%>
         <tr bordercolor="#000000">
         <td width="100" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=strTipo %></font></td>
         <td width="100" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rsio.Fields.Item("Factura").Value%></font></td>

         <td width="100" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=strFecha%></font></td>
         <%if strfactura = "" or trim(strfactura) <> trim(Rsio.Fields.Item("Factura").Value) then%>
            <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rsio.Fields.Item("Saldo").Value%></font></td>
         <%else%>
            <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">0</font></td>
         <%end if%>
          <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rsio2.Fields.Item("Referencia").Value%></font></td>

          <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rsio2.Fields.Item("Cgastos").Value%></font></td>

          <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=cdbl(Rsio2.Fields.Item("Suma").Value)%></font></td>
          <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rsio2.Fields.Item("desc21").Value%></font></td>
		  <td>&nbsp;</td>
         </tr>

        <%

         dblSumaCuentasSaldoMayor0 =  dblSumaCuentasSaldoMayor0 + cdbl(Rsio2.Fields.Item("Suma").Value)
         strfactura =  Rsio.Fields.Item("Factura").Value
         Rsio2.MoveNext()
       Wend
       Rsio2.Close()
       Set Rsio2 = Nothing

 else%>
       <tr bordercolor="#000000">
       <td width="100" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=strTipo %></font></td>
       <td width="100" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rsio.Fields.Item("Factura").Value%></font></td>
       <td width="100" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=strFecha%></font></td>
       <%if strfactura = "" or  trim(strfactura) <> trim(Rsio.Fields.Item("Factura").Value) then%>
         <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=cdbl(Rsio.Fields.Item("Saldo").Value)%></font></td>
        <% else%>
          <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">0</font></td>
        <% end if%>

         <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">0</font></td>
         <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">0</font></td>
         <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">0</font></td>
	     <td>&nbsp;</td>
    </tr>
         <%strfactura =  Rsio.Fields.Item("Factura").Value
 end if
 strfactura =  Rsio.Fields.Item("Factura").Value
 Rsio.MoveNext()
Wend
Rsio.Close()
Set Rsio = Nothing
end if
'IMPRIMIMOS EL TOTAL
%>


         <tr bordercolor="#000000">
         <td colspan="8" nowrap bgcolor="#00FFFF"><strong><font color="#000000" size="3" face="Arial, Helvetica, sans-serif">TOTAL DE PAGOS HECHOS DE LAS CUENTAS DE GASTOS NO LIQUIDADAS</font><font color="#000000" size="3" face="Arial, Helvetica, sans-serif"></font></strong></td>
         <td width="90" nowrap><font color="#000000" size="3" face="Arial, Helvetica, sans-serif"><strong><%= dblSumaCuentasSaldoMayor0%></strong></font> </td>
         </tr>

<%

   dblSumaPagosHechosNoFacturados = 0
' AQUI LE AGREGAMOS LOS PAGOS HECHOS NO FACTURADOS
   X=1
    While NOT X > 3

     BDx = ""
     if cint(X)=1 then
        BDx = "MAN"
     end if
     if cint(X)=2 then
       BDx = "MEX"
     end if
     if cint(X)=3 then
        BDx = "VER"
     end if
     if cint(X)=4 then
        BDx = "TAM"
     end if

 MM_EXTRANET_STRINGX = ODBC_POR_ADUANA(BDx)

   set Rsio = server.CreateObject("ADODB.Recordset")
   Rsio.ActiveConnection = MM_EXTRANET_STRINGX

   permi3 = PermisoClientes(BDx,strPermisos,"clie01")

   if not permi3 = "" then
      permi3 = "  and (" & permi3 & ") "
   end if
   if blnAplicaFiltro then
     permi3 = " AND clie01 =" & strFiltroCliente
   end if
   if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
     permi3 = ""
   end if




   strSQL3 = "select refe21 as Referencia,e21paghe.conc21,desc21,sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)) as Suma FROM d21paghe, e21paghe,c21paghe,c01refer WHERE YEAR(d21paghe.fech21) =  YEAR(e21paghe.FECH21) AND d21paghe.foli21 =  e21paghe.foli21 and e21paghe.tmov21=d21paghe.tmov21 and e21paghe.conc21 = c21paghe.clav21 and d21paghe.refe21 = c01refer.refe01 " & permi3 & " and e21paghe.esta21 <> 'S' and e21paghe.fech21<='"&strDateFin&"' and cgas21='' and e21paghe.tpag21 <> 3 group by conc21,refe21  having sum(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1))   > 0.1"
   'Response.Write(strSQL3)
   'Response.End
   Rsio.Source= strSQL3
   Rsio.CursorType = 0
   Rsio.CursorLocation = 2
   Rsio.LockType = 1
   Rsio.Open()

  if not Rsio.EOF then
   While NOT Rsio.EOF%>
        <tr bordercolor="#000000">
         <td width="100" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=strTipo %></font></td>
         <td width="100" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"></font></td>
         <td width="100" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"></font></td>
         <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">0</font></td>
         <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rsio.Fields.Item("Referencia").Value%></font></td>
         <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"></font></td>
         <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rsio.Fields.Item("Suma").Value%></font></td>
         <td width="90" nowrap><font color="#000000" size="1" face="Arial, Helvetica, sans-serif"><%=Rsio.Fields.Item("desc21").Value%></font></td>
         <%dblSumaPagosHechosNoFacturados = dblSumaPagosHechosNoFacturados + cdbl(Rsio.Fields.Item("Suma").Value)%>
		 <td>&nbsp;</td>
    </tr>
         <%
    Rsio.MoveNext()
   Wend
  end if
  Rsio.Close()
  Set Rsio = Nothing

    X= X + 1
  WEND

  'IMPRIMIMOS EL TOTAL%>
       <tr bordercolor="#000000">
       <td colspan="8" nowrap bgcolor="#00FFFF"><strong><font color="#000000" size="3" face="Arial, Helvetica, sans-serif">TOTAL PAGOS HECHOS NO FACTURADOS</font><font color="#000000" size="3" face="Arial, Helvetica, sans-serif"></font></strong></td>
       <td width="90" nowrap><font color="#000000" size="3" face="Arial, Helvetica, sans-serif"><strong><%=formatnumber(dblSumaPagosHechosNoFacturados,2)%></strong></font></td>
       </tr>
<%
dblDisponible = dblSumaFondo  - dblSumaPagosHechosNoFacturados - dblSumaCuentasSaldoMayor0
'IMPRIMIMOS EL GRAN TOTAL%>
         <tr bordercolor="#000000">
         <td colspan="8" nowrap bgcolor="#FFFF00"><strong><font color="#000000" size="3" face="Arial, Helvetica, sans-serif">DISPONIBLE EN EL FONDO </font><font color="#000000" size="3" face="Arial, Helvetica, sans-serif"></font></strong></td>
         <td width="90" nowrap bgcolor="#FFFF00"><font color="#000000" size="3" face="Arial, Helvetica, sans-serif"><strong><%=formatnumber(dblDisponible,2) %></strong></font></td>
         </tr>
</table>
