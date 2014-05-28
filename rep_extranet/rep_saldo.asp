<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% Server.ScriptTimeout=1500 %>
<HTML>
<HEAD>
<TITLE>:: REPORTE DE SALDOS.... ::</TITLE>
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

    %>
<%''''''''''''''''''''''''if  Session("GUsuario") <> "" then
if  Session("GAduana") <> "" then %>
    <table width="778"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="center">
          <%
          dim cve,mov,fi,ff,tabla,sql1,sql2,sql3,refe

		  oficina_adu=GAduana

		   jnxadu=Session("GAduana")

       'Response.Write(jnxadu)
       'Response.End

           select case jnxadu
             case "VER"
                  strOficina="rku"
             case "MEX"
                  strOficina="dai"
             case "MAN"
                  strOficina="sap"
             case "GUA"
                  strOficina="rku"
             case "TAM"
                  strOficina="ceg"
             case "LAR"
                  strOficina="LAR"
             case "LZR"
                  strOficina="lzr"
           end select

      buque= request.Form("txtBuq")
	  fet=request.Form("txtFeta")
	  OFI=strOficina
      
	   if request.form("rbnTipoDate") = "2" then
          Response.Addheader "Content-Disposition", "attachment;"
          Response.ContentType = "application/vnd.ms-excel"
       end if
		  
      DiaI = cstr(datepart("d",fet))
      Mesi = cstr(datepart("m",fet))
      AnioI = cstr(datepart("yyyy",fet))
      feta = Anioi&"/"&Mesi&"/"&Diai  

       Dim vsumpped(1)
       Dim vsumsaldl(1)
       Dim vsumsaldo(1)
	   
'************************************ENCABEZADO TCE**********************
 MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

MM_EXTRANET_STRING2 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=rku_cpsimples; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
          if buque<>"" then
		  Set RSf = Server.CreateObject("ADODB.Recordset")
          RSf.ActiveConnection =  MM_EXTRANET_STRING2
          STRSQL= " select rku_extranet.ssdagi01.nombar01 as buque,rku_cpsimples.pedimentos.refe01 as refe, " &_
          " rku_extranet.c01refer.feta01 as feta,rku_cpsimples.pedimentos.pesobr01 as peso, " &_
          " rku_cpsimples.pedimentos.nomcli01 as clie,rku_cpsimples.pedimentos.pedi01 as pedi, " &_
		  " rku_cpsimples.pedimentos.fecpag01 as fecpag " &_
          " from rku_cpsimples.pedimentos, rku_extranet.ssdagi01, rku_extranet.c01refer " &_
          " where rku_extranet.ssdagi01.nombar01 like '%"&buque&"%'  " &_
		  " and rku_extranet.ssdagi01.refcia01=rku_extranet.c01refer.refe01 and " &_
		  " rku_extranet.ssdagi01.refcia01=rku_cpsimples.pedimentos.refe01 "
		 	'response.Write(strsql)
			'response.end
			RSf.Source = strSQL
            RSf.CursorType = 0
            RSf.CursorLocation = 2
            RSf.LockType = 1
            RSf.Open()
		    if not RSf.eof then
		  	   referencia=RSf.Fields.Item("refe").Value
		       pedimento= RSf.Fields.Item("pedi").Value
		       cliente=RSf.Fields.Item("clie").Value
		       buque= RSf.Fields.Item("buque").Value
		       feta=RSf.Fields.Item("feta").Value
		       peso=RSf.Fields.Item("peso").Value
			   fecpago=RSf.Fields.Item("fecpag").Value
		       vence=fecpago+60
		     end if
		     RSf.close
		     set RSf = nothing		
			 %>
	         <BR>
  <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>Reporte de Saldos TCE </p></font></strong>
   <table align="left" >
       <table width="854"  border="1" cellspacing="3" cellpadding="3">
	   <tr>
  <td width="12%" bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Saldos del B/M:</b></FONT></td>
	   <td Width=88% align="left"><font size="-1"><%RESPONSE.Write(buque)%></font></td>
	   </tr>
	   <tr>
       <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Orden:</b></FONT></td>
	   <td Width=88% align="left"><font size="-1"><%RESPONSE.Write("-")%></font></td>
	   </tr>
	   <tr>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Fecha ETA:</b></FONT></td>
	   <td Width=88% align="left"><font size="-1"><%RESPONSE.Write(feta)%></font></td>
        </tr>
	   </table>
	     <br> 
		<table width="854" border="1" cellspacing="3" cellpadding="3">
        <tr>
          <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Cliente</FONT></th>
          <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Pedimento</FONT></th>
          <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Saldo Ped.</FONT></th>
          <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Saldo Lib.</FONT></th>
          <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Saldo</FONT></th>
          <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Producto</FONT></th>
          <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Vencimiento</FONT></th>
        </tr>	
			
	 	
	       
			 
<%			 
'*********************************************************************************

  Set Conn = Server.CreateObject ("ADODB.Connection")
  Set REFE = Server.CreateObject ("ADODB.RecordSet")
  Conn.Open MM_EXTRANET_STRING  
  
  Set Conn = Server.CreateObject ("ADODB.Connection")
  Set REFE = Server.CreateObject ("ADODB.RecordSet")
  Conn.Open MM_EXTRANET_STRING2 

  DiaI = cstr(datepart("d",fet))
  Mesi = cstr(datepart("m",fet))
  AnioI = cstr(datepart("yyyy",fet))
  feta = Anioi&"/"&Mesi&"/"&Diai  

 STRSQL= " select rku_extranet.ssdagi01.nombar01 as buque,rku_cpsimples.pedimentos.refe01 as refe, " &_ 
 " rku_extranet.c01refer.feta01 as feta,rku_cpsimples.pedimentos.pesobr01 as peso, " &_
 " rku_cpsimples.pedimentos.nomcli01 as clie,rku_cpsimples.pedimentos.pedi01 as pedi, " &_
 " rku_cpsimples.pedimentos.fecpag01 as fecpag " &_
 " from rku_cpsimples.pedimentos, rku_extranet.ssdagi01, rku_extranet.c01refer " &_
 " where rku_extranet.ssdagi01.nombar01 like '%"&buque&"%' and rku_extranet.c01refer.feta01='"&feta&"' and " &_
 " rku_extranet.ssdagi01.refcia01=rku_extranet.c01refer.refe01 and " &_
 " rku_extranet.ssdagi01.refcia01=rku_cpsimples.pedimentos.refe01 " 
 
	
	     Set REFE= Conn.Execute(strSQL)
         'response.Write(strsql)
         'response.end()

        sumpped=0
        sumsaldolibre=0
		sumsaldotot=0
        
        if not REFE.eof then 
        While (NOT  REFE.EOF) 
        sumpesoneto=0
	    referencia=REFE("refe")  	
	    pedimento=REFE("pedi")
        cliente=REFE("clie") 
		buque=REFE("buque")
		feta=REFE("feta")
		nomcli=REFE("clie")
		peso=REFE("peso")
		vence=fecpago+60
		sumpped=sumpped+peso
		
		'************************************Mercancia TCE**********************
		  Set RSf = Server.CreateObject("ADODB.Recordset")
          RSf.ActiveConnection =  MM_EXTRANET_STRING2
		  strSQL="select * from fracciones where refe01='"&referencia&"'  " 
		  RSf.Source = strSQL
          RSf.CursorType = 0
          RSf.CursorLocation = 2
          RSf.LockType = 1
          RSf.Open()
		  if not RSf.eof then
		    frac=RSf.Fields.Item("frac01").Value
		    mercancia = RSf.Fields.Item("desc01").Value
		 end if
		  RSf.close
		  set RSf = nothing		

	'************************************partidas tce**********************
	    saldolib=0
		saldo=0

		Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RSPart = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING2
		strSQL=" select * from  tcepartidas where refe01='"&referencia&"' and frac01='"&frac&"'  "
	    Set RSPart= Connx.Execute(strSQL)
		if not RSPart.eof then 			
        Do while not RSPart.Eof
            partida=RSPart("parfra01")
			pesoneto=RSPart("pesoneto")
			saldolib=saldolib+round(pesoneto)
            saldo=peso-saldolib
		    sumpesoneto=sumpesoneto+round(pesoneto)
			sumsaldo=sumsaldo+saldo
			
			RSPart.MoveNext  ' de las partidas de tce 
			Loop ' de las partidas de tce	
		  else 
		    saldo=peso
			
		  end if
		
			
        '------------------------------------para la segunda tablas    %>
	  <br>
	      <tr>
          <td><font size="-1"><%RESPONSE.Write(cliente)%></font></td>
          <td><font size="-1"><%RESPONSE.Write(pedimento)%></font></td>
          <td><font size="-1"><%RESPONSE.Write(peso)%></font></td>
          <td><font size="-1"><%RESPONSE.Write(sumpesoneto)%></font></td>
          <td><font size="-1"><%RESPONSE.Write(saldo)%></font></td>
          <td><font size="-1"><%RESPONSE.Write(mercancia)%></font></td>
          <td><font size="-1"><%RESPONSE.Write(vence)%></font></td>
        </tr>
        
     
	  <%

	        
		' else ' del if not RSPart.eof then partidastce 
'			pesoneto=0
'			saldolib=0+round(pesoneto)
'            saldo=peso-saldolib
'         end if  
'**********************************************************************
   '	RESPONSE.Write("</tr>")
        ' Refe.MoveNext 'avanza referencia  ---->
         Refe.MoveNext 
		 sumpped=sumpped
		 sumsaldolibre=sumsaldolibre+sumpesoneto
		 sumsaldotot=sumsaldotot+saldo
    	 wend 'REFErencia
       
	   %>
          
		  <tr>
          <td colspan="2" align="right"><b><font size="2" face="Arial" COLOR="black">Saldos Totales</font><b></td>
          <td bgcolor="#C0C0C0"><b><font size="-1"><%RESPONSE.Write(sumpped)%></font><b></td>
		  <td bgcolor="#C0C0C0"><b><font size="-1"><%RESPONSE.Write(sumsaldolibre)%></font><b></td>
		  <td bgcolor="#C0C0C0"><b><font size="-1"><%RESPONSE.Write(sumsaldotot)%></font><b></td>
        </tr>
		</table>	  
		<%
     
 else
%>
<tr>
  <th colspan=12>
    <font size="1" face="Arial">No se Encontro ningun registro con esos parametros </font>
  </th>
</tr>
<table>
<%'end if
end if
else
response.write("<p class=""TextPequeAzul"">INGRESE EL NOMBRE DEL BUQUE</p>")
end if' fin del if de buque<>""
%>
</form>  



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