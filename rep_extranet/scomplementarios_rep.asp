
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
 
 <style type="text/css">
.style20 {color: #FFFFFF}
 </style>

<%
    Response.Buffer = TRUE
    Response.Addheader "Content-Disposition", "attachment;filename=ServCompl.xls"
    Response.ContentType = "application/vnd.ms-excel"
    Server.ScriptTimeOut=100000

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
     ' Response.End
    end if

strTipoUsuario = request.Form("TipoUser")
strPermisos    = Request.Form("Permisos")
permi  = PermisoClientes(Session("GAduana"),strPermisos,"c.clie31")
'Response.Write(permi)
'Response.End

  if not permi = "" then
     permi = "  and (" & permi & ") "
  end if


cvecli=request.Form("txtCliente")

if cvecli="Todos" then
nomb=cvecli
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")

if not strFiltroCliente= "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if

if blnAplicaFiltro then
   permi = " AND c.clie31 =" & strFiltroCliente
end if

if  strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
  permi = ""
end if


'Response.Write(permi)
'Response.End


if cvecli<>"Todos" then
'MM_EXTRANET_STRING = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER="& IPHost &"; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
  MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
  Set RsRevisa = Server.CreateObject("ADODB.Recordset")
  RsRevisa.ActiveConnection = MM_EXTRANET_STRING
  strSQL=  "SELECT cvecli18,NOMCLI18 AS NOMBRE FROM SSCLIE18 where cvecli18='"&cvecli&"' "
  RsRevisa.Source = strSQL
  RsRevisa.CursorType = 0
  RsRevisa.CursorLocation = 2
  RsRevisa.LockType = 1
  RsRevisa.Open()
  if not RsRevisa.eof then
      nomb =  RsRevisa.Fields.Item("nombre").Value
	  end if
  RsRevisa.close
  set RsRevisa = nothing
end if

'pa probar
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
 Set Conn = Server.CreateObject ("ADODB.Connection")
   Set RS = Server.CreateObject ("ADODB.RecordSet")
   Conn.Open MM_EXTRANET_STRING
  strSQL="select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
" c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
" a.fdsp01 as fechadespacho from c01refer a,d31refer b,e31cgast c " &_
" where a.refe01=b.refe31 and b.cgas31=c.cgas31  and c.esta31='I' " &_
" and c.fech31>='"&ISTRFINI&"' and c.fech31<='"&FSTRFFIN&"' "&permi&" " &_
"order by refe01 "

  'response.Write(strsql)
  'response.End()


  Set RS= Conn.Execute(strSQL)


 While (NOT  RS.EOF)
 referenciaj=RS("referencia")
 'if referenciaj=" " then
    'response.Write("No Tiene Referencias en estas fechas")
 'end if
     'PARA NUMERO DE CONCEPTOS DE SERVIOS COMPLEMENTARIOS DE C/REFE

   Set Conn2 = Server.CreateObject ("ADODB.Connection")
   Set RS2 = Server.CreateObject ("ADODB.RecordSet")
   Conn2.Open MM_EXTRANET_STRING
   sql="select count(dcrp32) AS NSC from d32rserv where refe32='"&RS("referencia")&"' "
   Set RS2 = Conn2.Execute(sql)
  ' response.Write(sql)
   'response.End()
    xj=RS2("NSC")
    totx=totx+xj
   RS.MoveNext
   Wend
'borara lo de aabajoelconcepto sinofuncina SI FUNCIONO ====================================================================
   if totx=0 then


  Set RsRevisa = Server.CreateObject("ADODB.Recordset")
  RsRevisa.ActiveConnection = MM_EXTRANET_STRING
  strSQL="select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
" c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
" a.fdsp01 as fechadespacho from c01refer a,d31refer b,e31cgast c " &_
" where a.refe01=b.refe31 and b.cgas31=c.cgas31  and c.esta31='I' " &_
" and c.fech31>='"&ISTRFINI&"' and c.fech31<='"&FSTRFFIN &"' "&permi&" " &_
"order by refe01 "
'RESPONSE.Write(strSQL)
'RESPONSE.End()
  RsRevisa.Source = strSQL
  RsRevisa.CursorType = 0
  RsRevisa.CursorLocation = 2
  RsRevisa.LockType = 1
  RsRevisa.Open()
  'REFI= RsRevisa.Fields.Item("REFERENCIA").Value
    %>
	 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Reporte de Servicios Complementarios </p></font></strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p></p></font></strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Del <%=STRFINI%> al <%=STRFFIN%> </p></font></strong>
      <table class="style9" align="left" >
      <tr bgcolor="#006699" >
	       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">REFERENCIA</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">FECHACG</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">CVECLI</FONT></td>
		   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">FECHA DESPACHO</FONT></td>
		   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">NO.PEDI</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">CG</FONT></td>
    	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">HONORARIOS</FONT></td>
		   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">PH</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">SC</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">ANTICIPO</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">SALDO</FONT></td>
		   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">TOTAL</FONT></td>
          <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">VALORADU</FONT></td>
	<td bgcolor="#FFFFFF"><span class="style20">llllllllllllllllllllllllllllllllllllllllllllllllllll</span></td>
    </tr>
	<% if not RsRevisa.eof then
	While (NOT  RsRevisa.EOF)
	'valor aduanal
	Set RsVad = Server.CreateObject("ADODB.Recordset")
          RsVad.ActiveConnection =  MM_EXTRANET_STRING
          strSQL="select sum(Vaduan02)as Val_Aduana from ssfrac02 where REFCIA02='"&RsRevisa.Fields.Item("REFERENCIA").Value&"' "
		  RsVad.Source = strSQL
          RsVad.CursorType = 0
          RsVad.CursorLocation = 2
          RsVad.LockType = 1
          RsVad.Open()
		  Vad=""
         if not RsVad.eof then
		    Vad = RsVad.Fields.Item("Val_Aduana").Value
		 end if
		  RsVad.close
		  set RsVad = nothing
	'NUMERO PEDIMENTO
	Set RsNoPD = Server.CreateObject("ADODB.Recordset")
          RsNoPD.ActiveConnection =  MM_EXTRANET_STRING
          strSQL="select NUMPED01 AS NoPD FROM SSDAGI01 where refcia01='"&RsRevisa.Fields.Item("REFERENCIA").Value&"' "
		  RsNoPD.Source = strSQL
          RsNoPD.CursorType = 0
          RsNoPD.CursorLocation = 2
          RsNoPD.LockType = 1
          RsNoPD.Open()
		  NoPD=""
         if not RsNoPD.eof then
		    NoPD = RsNoPD.Fields.Item("NoPD").Value
		 end if
		  RsNoPD.close
		  set RsNoPD = nothing
	%>
	   <tr>
	     <td><%=RsRevisa.Fields.Item("REFERENCIA").Value%></td>
 	     <td><%=RsRevisa.Fields.Item("FECHACG").Value%></td>
	     <td><%=RsRevisa.Fields.Item("CVECLIENTE").Value%></td>
		 <%if RsRevisa.Fields.Item("fechadespacho").Value<>" " then %>
         <td align="center"><%=RsRevisa.Fields.Item("fechadespacho").Value%></td>
		 <%else%>
		 <td>-</td>
		 <%end if%>
		 <%if NoPD<>" " then %>
         <td><%response.Write(NoPD)%></td>
		 <%else%>
		 <td>---</td>
		 <%end if%>
		 <td><%=RsRevisa.Fields.Item("CG").Value%></td>
		 <td align="center"><%=RsRevisa.Fields.Item("HONORARIOS").Value%></td>
 	     <td><%=RsRevisa.Fields.Item("PH").Value%></td>
	     <td><%=RsRevisa.Fields.Item("SC").Value%></td>
	     <td><%=RsRevisa.Fields.Item("ANTICIPO").Value%></td>
		 <td><%=RsRevisa.Fields.Item("SALDO").Value%></td>
 	     <td><%=RsRevisa.Fields.Item("TOTAL").Value%></td>
         <%if Vad<>" " then %>
         <td align="center"><%response.Write(vad)%></td>
		 <%else%>
		 <td>---</td>
		 <%end if%>
		 <%if RsRevisa.Fields.Item("SC").Value>0 then %>
         <td><font size="1" face="Arial, Helvetica, sans-serif">No se Registro el Concepto</font></td>
		 <%end if%>
	   </tr>

<%  RsRevisa.MoveNext()
    Wend
	RsRevisa.close
	set  RsRevisa = nothing
	else
%>
<tr>
  <th colspan=12>
    <font size="1" face="Arial">No se Encontro ningun registro con esos parametros</font>
  </th>
</tr>
<%
end if 'del if Not revisa.eof	%>
</table>
<% else   'aky si existe un servicio complementario


  Set Conn = Server.CreateObject ("ADODB.Connection")
  Set REF8 = Server.CreateObject ("ADODB.RecordSet")
  Conn.Open MM_EXTRANET_STRING
  SQL="select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
  " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
  " a.fdsp01 as fechadespacho from c01refer a,d31refer b,e31cgast c " &_
  " where a.refe01=b.refe31 and b.cgas31=c.cgas31  and c.esta31='I' " &_
  " and c.fech31>='"&ISTRFINI&"' and c.fech31<='"&FSTRFFIN&"' "&permi&" " &_
  " order by refe01 "
'RESPONSE.Write(strSQL)
'RESPONSE.End()
  Set REF8= Conn.Execute(SQL)
      referenciaj=REF8("referencia")
       %>
	 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Reporte de Servicios Complementarios </p></font></strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p></p></font></strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Del <%=STRFINI%> al <%=STRFFIN%> </p></font></strong>
	<table class="style9" align="left" >
     <tr bgcolor="#006699" >
	       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">REFERENCIA</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">FECHACG</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">CVECLI</FONT></td>
		   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">FECHA DESPACHO</FONT></td>
		   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">NO.PEDI</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">CG</FONT></td>
    	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">HONORARIOS</FONT></td>
		   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">PH</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">SC</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">ANTICIPO</FONT></td>
           <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">SALDO</FONT></td>
		   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">TOTAL</FONT></td>
          <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">VALORADU</FONT></td>
 <%  i=0
  Dim famname(2000)
  Dim conceptos(2000)
  Dim TempArray(2000)
  Dim XMontos(2000)
  Dim Encabezado(2000)
  Dim encabezadoONE(2000)
  Dim NOMBRECONCEPTO(2000)
  
   While (NOT  REF8.EOF)

      Set RsStatus = Server.CreateObject("ADODB.Recordset")
          RsStatus.ActiveConnection =  MM_EXTRANET_STRING
		  strSQL="select count(dcrp32) AS NSC from d32rserv where refe32='"&REF8("referencia")&"' "
		  RsStatus.Source = strSQL
          RsStatus.CursorType = 0
          RsStatus.CursorLocation = 2
          RsStatus.LockType = 1
          RsStatus.Open()
		  Contador=""
         if not RsStatus.eof then
		    Contador = RsStatus.Fields.Item("NSC").Value
		 end if
		  RsStatus.close
		  set RsStatus = nothing

    If Contador<>0  Then
	  Set Connx = Server.CreateObject ("ADODB.Connection")
      Set RSx = Server.CreateObject ("ADODB.RecordSet")
      Connx.Open MM_EXTRANET_STRING
	  SQLx="select refe32 as refe,dcrp32 as concepto, mont32 as Monto,ttar32 as idconcepto from d32rserv " &_
      " where refe32='"&REF8("referencia")&"' "
	  Set RSx= Conn.Execute(SQLx)
      concepto=RSx("concepto")
	  idconcepto=RSx("idconcepto")
      Montos=RSx("Monto")

      Do while not RSx.Eof
        concepto=RSx("concepto")
	    idconcepto=RSx("idconcepto")
		Montos=RSx("Monto")
	    famname(i)=idconcepto
		conceptos(i)=idconcepto
		XMontos(i)=Montos
		NOMBRECONCEPTO(i)=concepto
		i = i+1
          RSx.MoveNext
       Loop
	   REF8.MoveNext	 'avenzamos la referencia
      'wend 'REFErencia
    ELSE
     REF8.MoveNext
	  'wend 'REFErencia
   end if   'contador sino tiene servicio avanzamos la referencia

 wend 'REFErencia


'========elimina repe VECTOR si funciona probado=============================
For iy = LBound(NOMBRECONCEPTO) To UBound(NOMBRECONCEPTO)
          'Asignamos al array temporal el valor del otro array
          TempArray(iy) = NOMBRECONCEPTO(iy)
   Next

    For x7 = 0 To UBound(NOMBRECONCEPTO)
        z = 0
        For y = 0 To UBound(NOMBRECONCEPTO)
            'Si el elemento del array es igual al array temporal
            If NOMBRECONCEPTO(x7) = TempArray(z) And y <> x7 Then
                'Entonces Eliminamos el valor duplicado
                NOMBRECONCEPTO(y) = ""
                Nduplicado = Nduplicado + 1
            End If
            z = z + 1
        Next
    Next
'======MUESTRA LO KE TIENE Y AGREGA AL ENCABEZADO LOS CONCEPTOS si esta vacio el vector avanza

	  ' For noce = 0 to ubound(conceptos)
   ' Response.write(conceptos(noce)&"Posicion"&noce&"--")
   ' next
   
    acpk=0
	      norepeti2k=0
          for acpk=0 to UBound(NOMBRECONCEPTO)
	         if NOMBRECONCEPTO(acpk)<>"" then
		       norepeti2k=norepeti2k+1
			   
		     end if
		  next	
    aaa7=0
          yyyy7=0
	      while  aaa7<norepeti2k
		     For yyyy7 = 0 to UBound(NOMBRECONCEPTO)
			 if NOMBRECONCEPTO(yyyy7)<>"" then
			    encabezadoONE(aaa7)=NOMBRECONCEPTO(yyyy7)
		        aaa7=aaa7+1
			 end if
             Next
		  Wend

'MUESTRA EL CONTENIDO DEL VECTOR
'For noce = 0 to ubound(NOMBRECONCEPTO)   
'Response.write(NOMBRECONCEPTO(noce)&"Posicion"&noce&"--")
 'next
	
   
	
	For k = 0 to norepeti2k-1
	 if  NOMBRECONCEPTO(k)<>" " then
	 Set RsConceptos = Server.CreateObject("ADODB.Recordset")
          RsConceptos.ActiveConnection =  MM_EXTRANET_STRING
		  strSQL="select *  from d32rserv  WHERE dcrp32 like'"&encabezadoONE(k)&"' "
		  RsConceptos.Source = strSQL
          RsConceptos.CursorType = 0
          RsConceptos.CursorLocation = 2
          RsConceptos.LockType = 1
          RsConceptos.Open()
		  Conceptoz=""
         if not RsConceptos.eof then
		    Conceptoz = RsConceptos.Fields.Item("dcrp32").Value %>
<td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%response.Write(conceptoz)%></FONT></td>  		 <% end if
		  RsConceptos.close
		  set RsConceptos = nothing
	    'Response.write(conceptos(i)&"Posicion"&i&"--")
     end if
	 next %>


  <td bgcolor="#FFFFFF"><span class="style20">llllllllllllllllllllllllllllllllllllllllllllllllllll</span></td>	 
<%

'======AHORA AGREGAMOS FILAS x REFERENCIAS====================================================


  Set Conn = Server.CreateObject ("ADODB.Connection")
  Set REF9 = Server.CreateObject ("ADODB.RecordSet")
  Conn.Open MM_EXTRANET_STRING
  SQL="select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
  " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
  " a.fdsp01 as fechadespacho from c01refer a,d31refer b,e31cgast c " &_
  " where a.refe01=b.refe31 and b.cgas31=c.cgas31  and c.esta31='I' " &_
  " and c.fech31>='"&ISTRFINI&"' and c.fech31<='"&FSTRFFIN&"' "&permi&" " &_
  " order by refe01 "
'RESPONSE.Write(strSQL)
'RESPONSE.End()
  Set REF9= Conn.Execute(SQL)
      referenciaj=REF9("referencia")
          %>
  </tr>
	<%   While (NOT  REF9.EOF)
	    jc=0
		'valor aduanal
	Set RsVad = Server.CreateObject("ADODB.Recordset")
          RsVad.ActiveConnection =  MM_EXTRANET_STRING
		  strSQL="select sum(Vaduan02) as Val_Aduana from ssfrac02 where REFCIA02='"&REF9("referencia")&"' "
		  RsVad.Source = strSQL
          RsVad.CursorType = 0
          RsVad.CursorLocation = 2
          RsVad.LockType = 1
          RsVad.Open()
		  Vad=""
         if not RsVad.eof then
		    Vad = RsVad.Fields.Item("Val_Aduana").Value
		 end if
		  RsVad.close
		  set RsVad = nothing
		'NUMERO PEDIMENTO
	Set RsNoPD = Server.CreateObject("ADODB.Recordset")
          RsNoPD.ActiveConnection =  MM_EXTRANET_STRING
          strSQL="select NUMPED01 AS NoPD FROM SSDAGI01 where refcia01='"&REF9("referencia")&"' "
		  RsNoPD.Source = strSQL
          RsNoPD.CursorType = 0
          RsNoPD.CursorLocation = 2
          RsNoPD.LockType = 1
          RsNoPD.Open()
		  NoPD=""
         if not RsNoPD.eof then
		    NoPD = RsNoPD.Fields.Item("NoPD").Value
		 end if
		  RsNoPD.close
		  set RsNoPD = nothing

	%>
	   <tr>
	     <td><%=REF9("REFERENCIA").Value%></td>
 	     <td><%=REF9("FECHACG").Value%></td>
	     <td><%=REF9("CVECLIENTE").Value%></td>
		 <%if REF9("fechadespacho").Value<>" " then %>
         <td align="center"><%=REF9("fechadespacho").Value%></td>
		 <%else%>
		 <td align="center">--</td>
		 <%end if%>
		 <%if NoPD<>" " then %>
         <td><%response.Write(NoPD)%></td>
		 <%else%>
		 <td align="center">--</td>
		 <%end if%>
		 <td><%=REF9("CG").Value%></td>
		 <td align="center"><%=REF9("HONORARIOS").Value%></td>
 	     <td align="center"><%=REF9("PH").Value%></td>
	     <td><%=REF9("SC").Value%></td>
	     <td align="center"><%=REF9("ANTICIPO").Value%></td>
		 <td align="center"><%=REF9("SALDO").Value%></td>
 	     <td align="center"><%=REF9("TOTAL").Value%></td>
		 <%if Vad<>" " then %>
         <td><%response.Write(vad)%></td>
		 <%else%>
		 <td align="center">--</td>
		 <%end if%>

	<%   Set RsStatus = Server.CreateObject("ADODB.Recordset")
          RsStatus.ActiveConnection =  MM_EXTRANET_STRING
		  strSQL="select count(dcrp32) AS NSC from d32rserv where refe32='"&REF9("referencia")&"' "
		  RsStatus.Source = strSQL
          RsStatus.CursorType = 0
          RsStatus.CursorLocation = 2
          RsStatus.LockType = 1
          RsStatus.Open()
		  Contador=""
         if not RsStatus.eof then
		    Contador = RsStatus.Fields.Item("NSC").Value
		 end if
		  RsStatus.close
		  set RsStatus = nothing
     ' Cuenta los llenos en vector fam y conceptos
	      norepetidosx=0
          for abc=0 to UBound(famname)
	        if famname(abc)<>"" then
		       norepetidosx=norepetidosx+1
		       end if
		  next
		  'response.Write(norepetidosx)

		  acp=0
	      norepeti2=0
          for acp=0 to UBound(NOMBRECONCEPTO)
	         if NOMBRECONCEPTO(acp)<>"" then
		       norepeti2=norepeti2+1
		     end if
		  next

		  xxx7=0
          yyy7=0
	      while  xxx7<norepeti2
		     For yyy7 = 0 to UBound(NOMBRECONCEPTO)
			 if NOMBRECONCEPTO(yyy7)<>"" then
			    encabezado(xxx7)=NOMBRECONCEPTO(yyy7)
		        xxx7=xxx7+1
			 end if
             Next
		  Wend


		  'For ipx = 0 to UBound(conceptos)
   ' Response.write(conceptos(ipx)&"Posicion"&ipx&"--")
    'next
		 ' response.end()

		 xfi=0
	      nonull=0
          for xfi=0 to UBound(encabezado)
	         if NOMBRECONCEPTO(xfi)<>"" then
		       nonull=nonull+1
		     end if
		  next

		  'For ipx2 = 0 to norepeti2-1
   ' Response.write(encabezado(ipx2)&"Posicion"&ipx2&"--")
   ' next
'	response.Write(norepeti2)
		  'response.end()
'========================================================
       If (Contador<>0) and (REF9("SC").Value>0)  Then
	  for jc=0 to nonull-1
Set RsSCom = Server.CreateObject("ADODB.Recordset")
          RsSCom.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="select refe32 as refe,dcrp32 as concepto, mont32 as Monto,ttar32 as idconcepto from d32rserv " &_
        " where dcrp32 like '"&encabezado(jc)&"' and refe32='"&REF9("REFERENCIA").Value&"' "
		  RsSCom.Source = strSQL
          RsSCom.CursorType = 0
          RsSCom.CursorLocation = 2
          RsSCom.LockType = 1
          RsSCom.Open()
		  Dinero=""
		  Nconcepto=""
         if not RsSCom.eof then
		    Dinero = RsSCom.Fields.Item("Monto").Value
			Nconcepto = RsSCom.Fields.Item("idconcepto").Value
			else
                 Dinero=0
               Nconcepto=0
          end if
		  RsSCom.close
		  set RsSCom = nothing

		  if dinero=0 then%>
		     <td align="center">0</td>
		  <%else%>
		  	 <td align="center"><%response.Write(Dinero)%></td>
	<%    end if
	  next
	  REF9.MoveNext  %>
	  </tr>

   <%ELSE
     a=0
	 norepetidos=0
     for a=0 to UBound(NOMBRECONCEPTO)
	    if NOMBRECONCEPTO(a)<>"" then
		   norepetidos=norepetidos+1
	    end if
	 next
     ' response.write(norepetidos)
	 ' response.End()
       jxy=1
		do while jxy<=norepetidos %>
		   <td align="center">0</td>
		    <%jxy=jxy+1
	    loop 
		'aky si tiene SC pero no concepto
		 if (REF9("SC").Value>0) then%>
		<td width="230">
		 <div > <font size="1" face="Arial, Helvetica, sans-serif">No se Registro el Concepto</font></div></td>
		  <%end if%>
	   </tr>
	    <% REF9.MoveNext %>

  <% end if
  wend %>
 <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
  <td bgcolor="#006699" align="center"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">TOTAL</font></td>
<% 'aki viene pa sacar montos
v=0
totalM=0
for v=0 to nonull-1
totalM=0
Set Conn = Server.CreateObject ("ADODB.Connection")
   Set RS7 = Server.CreateObject ("ADODB.RecordSet")
   Conn.Open MM_EXTRANET_STRING
  strSQL="select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
" c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
" a.fdsp01 as fechadespacho from c01refer a,d31refer b,e31cgast c " &_
" where a.refe01=b.refe31 and b.cgas31=c.cgas31  and c.esta31='I' " &_
" and c.fech31>='"&ISTRFINI&"' and c.fech31<='"&FSTRFFIN&"' "&permi&" " &_
"order by refe01 "
  Set RS7= Conn.Execute(strSQL)
  'response.Write(strsql)
  'response.End()

 While (NOT  RS7.EOF)
 referenciaj=RS7("referencia")
 'if referenciaj=" " then
    'response.Write("No Tiene Referencias en estas fechas")
 'end if
 if RS7("SC").Value>0 then 
 Set RsSCom = Server.CreateObject("ADODB.Recordset")
          RsSCom.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="select refe32 as refe,dcrp32 as concepto, mont32 as Monto,ttar32 as idconcepto from d32rserv " &_
        " where dcrp32 like'"&encabezado(v)&"' and refe32='"&RS7("referencia")&"' "
		  RsSCom.Source = strSQL
          RsSCom.CursorType = 0
          RsSCom.CursorLocation = 2
          RsSCom.LockType = 1
          RsSCom.Open()
		  Dinero2=""
		  Nconcepto2=""
         if not RsSCom.eof then
		    Dinero2 = RsSCom.Fields.Item("Monto").Value
			Nconcepto2 = RsSCom.Fields.Item("idconcepto").Value
			else
                 Dinero2=0
               Nconcepto2=0
          end if
		  RsSCom.close
		  set RsSCom = nothing


    totalM=totalM+Dinero2
end if	 
   RS7.MoveNext
   Wend %>

   <td align="center"> <%response.Write(totalM)%></td>
<%next	%>
</table>
 <%


'======================nO BORRAR LO DE ABAJO ESTA BIEN ASI=================
end if 'del totx
'end if 'del movimiento %>


