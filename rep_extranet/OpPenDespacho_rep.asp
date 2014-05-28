 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
 
 <style type="text/css">
.style20 {color: #FFFFFF}
 </style>

<%
    Response.Buffer = TRUE
    Response.Addheader "Content-Disposition", "attachment;filename=OpPenDespacho.xls"
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
permi  = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
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
   permi = " AND cvecli01 =" & strFiltroCliente
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

'Busca referencias en ssdagi y ssdage
  MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
  Set Conn = Server.CreateObject ("ADODB.Connection")
  Set REFE = Server.CreateObject ("ADODB.RecordSet")
  Conn.Open MM_EXTRANET_STRING
  SQL="select refcia01,Numped01,nomcli01,cveped01,fecpag01,'Importacion' as tipoper,fecent01 from ssdagi01 where  " &_
  "fecpag01>= '"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"'  and firmae01='' "&permi&"  union all " &_
  "select refcia01,Numped01,nomcli01,cveped01,fecpag01,'Exportacion' as tipoper,fecpre01 from ssdage01 where   " &_
  "fecpag01>= '"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"'  and firmae01='' "&permi&" order by fecpag01 "
  Set REFE= Conn.Execute(SQL)
  'response.Write(SQL)
  'response.end%>
  
  <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>Reporte de Operaciones Pendientes de Despacho</p></font></strong>
  <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p></p></font></strong>
  <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Del <%=STRFINI%> al <%=STRFFIN%> </p></font></    strong>
 <table class="style9" align="left" >
   <tr bgcolor="#FFFFFF" >
     <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">TIPO PDTO</FONT></td>
     <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">REFERENCIA</FONT></td>
     <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">PEDIMENTO</FONT></td>
     <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">CLIENTE</FONT></td>
     <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">CLAVE</FONT></td>
     <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">FECHA DE ENTRADA</FONT></td>
     <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">DIAS TRANSCURRIDOS</FONT></td>
     <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">COMENTARIOS</FONT></td>
   </tr>  
   <%
   	if not REFE.eof then 
      While (NOT  REFE.EOF)
       referencia=REFE("refcia01")
	   Set Connx = Server.CreateObject ("ADODB.Connection")
       Set RSd = Server.CreateObject ("ADODB.RecordSet")
       Connx.Open MM_EXTRANET_STRING
	   SQL="select refcia01,Numped01,nomcli01,cveped01,fecpag01,'Importacion' as tipoper,REFE01,CGAS01,MODO01,feta01, " &_
	   "fecent01 as fent from ssdagi01,C01REFER  where refcia01='"&REFE("refcia01")&"' and  " &_
       "fecpag01>= '"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"' and firmae01=''  " &_
	   " and REFE01=refcia01 and CGAS01='' and modo01='T'  "&permi&"  union all " &_
       "select refcia01,Numped01,nomcli01,cveped01,fecpag01,'Exportacion' as tipoper,REFE01,CGAS01,MODO01,feta01, " &_
	   "fecpre01 as fent from ssdage01,C01REFER where refcia01='"&REFE("refcia01")&"' and  " &_
        "fecpag01>= '"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"' and firmae01='' and REFE01=refcia01 and " &_
		" CGAS01='' and modo01='T'  "&permi&" order by fecpag01 "
	    Set RSd= Connx.Execute(SQL)
		'response.Write(sql)
		'response.End()
            refcg=" "
	        tipo=" "
            fentrada=" "
			fechapresenta=" "
		
	    Do while not RSd.Eof
            refcg=RSd("REFE01")
	        tipo=RSd("modo01")
            fentrada=RSd("feta01")
			referencia=RSd("refcia01")
		    tipo_oper=RSd("tipoper")
		    npedimento=RSd("Numped01") 
		    nombreclie=RSd("nomcli01") 
		    cvepedimento=RSd("cveped01")  
            fechapago=RSd("fecpag01")
			if tipo_oper="Exportacion" then
               fechapresenta=RSd("fent")
			end if 
'================================================================================================        
		    fechita=IsDate(fentrada)
		    'response.Write(fechita)
            fechactual=date()
			dias=DateDiff("d", fentrada, fechactual)
			if tipo_oper="Exportacion" then
			   if fechapresenta<>"" or month(fechapresenta)=0 then
			      dias2=DateDiff("d", fechapresenta, fechactual)
			   else
			      dias2=DateDiff("d", fechapago, fechactual) 
			   end if	  
			end if%>
		 <tr>
		  <td><%RESPONSE.Write(tipo_oper)%></td>
		 <td><%RESPONSE.Write(referencia)%></td>
		 <td align="center"><%RESPONSE.Write(npedimento)%></td>	
		 <td><%RESPONSE.Write(nombreclie)%></td>
		 <td align="center"><%RESPONSE.Write(cvepedimento)%></td>
		 <%if  tipo_oper="Exportacion" then 
		     if fechapresenta<>"" or month(fechapresenta)=0 then%>
			   <td align="center"><%RESPONSE.Write(fechapresenta)%></td>
			   <%else%>
			   <td align="center"><%RESPONSE.Write(fechapago)%></td>
			 <%end if%>  
		    <td align="center"><%RESPONSE.Write(dias2)%></td>
		 <%else
		    if tipo_oper="Importacion" then		 
		       if (cvepedimento="CT" or cvepedimento="A3") then 
		          dias=DateDiff("d", fechapago, fechactual)%>
		          <td align="center"><%RESPONSE.Write(fechapago)%></td>
		       <%else
		       if fentrada<>"" or month(fentrada)=0 then%>
		          <td align="center"><%RESPONSE.Write(fentrada)%></td>
		       <%else
		          dias=DateDiff("d", fechapago, fechactual)%>
		          <td align="center"><%RESPONSE.Write(fechapago)%></td>
		       <%end if
		       end if%>
 		     <td align="center"><%RESPONSE.Write(dias)%></td>
		    <%end if
		 end if%>
		 <td align="center"></td>
		</tr>
		 
		<% RSd.MoveNext  ' de la CUENTA DE GASTOS
          Loop ' de la CUENTA DE GASTOS
	    Refe.MoveNext 'avanza referencia  --->
   wend 'REFErencia
 else
%>
<tr>
  <th colspan=8>
    <font size="1" face="Arial">No se Encontro ningun registro con esos parametros
  </th>
</tr>
<table>
<%end if%>
