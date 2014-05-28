<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% Server.ScriptTimeout=1500 %>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<HTML>
<HEAD>
<TITLE>:: REPORTE DE PESOS .... ::</TITLE>
</HEAD>
<BODY>
<% 
IPHost = "localhost"
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
       permi = " AND a.cvecli01 =" & strFiltroCliente
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

         ' pedimento=request.form("txtPed")
          referencia=request.form("txtRef")
		  if referencia<>"" then
	        opp="and"
	      else
	        opp="or"
     	  end if
        
          if request.form("rbnTipoDate") = "2" then
             Response.Addheader "Content-Disposition", "attachment;filename=Rep_pesos.xls"
             Response.ContentType = "application/vnd.ms-excel"
          end if
		  
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

MM_EXTRANET_STRING2 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="& IPHost &"; DATABASE=rku_cpsimples; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
  Set Conn = Server.CreateObject ("ADODB.Connection")
  Set REFE = Server.CreateObject ("ADODB.RecordSet")
  Conn.Open MM_EXTRANET_STRING  
  
  Set Conn = Server.CreateObject ("ADODB.Connection")
  Set REFE = Server.CreateObject ("ADODB.RecordSet")
  Conn.Open MM_EXTRANET_STRING2 

 STRSQL= " SELECT a.refe01 as ref,a.pedi01 as ped,a.nomcli01 as clie,a.patent01 as pat,a.fecpag01,b.frac01 as frac, " &_
         " b.desc01 as merc,a.fecpag01 as fecpag,a.pesobr01 as pesob,b.partida01 as partida, a.cvecli01  " &_
		 " FROM pedimentos a,fracciones b" &_
         " where a.refe01='"&referencia&"' and a.refe01=b.refe01 "&permi&" "

  Set REFE= Conn.Execute(strSQL)
 'response.Write(strsql)
 'response.end()
  %>
  <br>		
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>Reporte de Pesos </p></font></strong>
 <table align="left" >
   
	 <%

       if not REFE.eof then 
       While (NOT  REFE.EOF) 

        referencia=REFE("ref")  	
	    pedimento=REFE("ped")
        cliente=REFE("clie") 
		patente=REFE("pat")
		mercancia=REFE("merc")
		fecpago=REFE("fecpag")
		peso=REFE("pesob")
		vence=fecpago+60
		frac=REFE("frac")
		saldolib=0
		
		tmpeso=trim(cstr(formatnumber(peso,0)))

		'******************************buque
		
		Set RsRevisa = Server.CreateObject("ADODB.Recordset")
        RsRevisa.ActiveConnection = MM_EXTRANET_STRING
        strSQL=" Select a.refe01,a.feta01 as eta,b.nombar01 as buque from c01refer a,ssdagi01 b " &_ 
        " where b.refcia01='"&referencia&"' and b.refcia01=a.refe01 "
		RsRevisa.Source = strSQL
        RsRevisa.CursorType = 0
        RsRevisa.CursorLocation = 2
        RsRevisa.LockType = 1
        RsRevisa.Open()
		'response.Write(strSQL)
		'response.End()
        if not RsRevisa.eof then
	       buque  =  RsRevisa.Fields.Item("buque").Value
	       fechaeta =  RsRevisa.Fields.Item("eta").Value
        end if
        RsRevisa.close
        set RsRevisa = nothing  
				
'		Set RsRevisa = Server.CreateObject("ADODB.Recordset")
'        RsRevisa.ActiveConnection = MM_EXTRANET_STRING
'        strSQL=" select a.refe01,a.feta01 as eta ,a.cbuq01,a.ptoemb01 as PUERTO,b.nomb06 AS BUQUE " &_
'        " from c01refer a,c06barco b where a.refe01='"&referencia&"' and a.cbuq01=b.clav06  "
'		RsRevisa.Source = strSQL
'        RsRevisa.CursorType = 0
'        RsRevisa.CursorLocation = 2
'        RsRevisa.LockType = 1
'        RsRevisa.Open()
'		'response.Write(strSQL)
'		'response.End()
'        if not RsRevisa.eof then
'	       strPuerto =  RsRevisa.Fields.Item("puerto").Value
'	       buque  =  RsRevisa.Fields.Item("buque").Value
'	       fechaeta =  RsRevisa.Fields.Item("eta").Value
'        end if
'        RsRevisa.close
'        set RsRevisa = nothing  
				
		'*****************************FACTURA PROVEEDOR*************************************
	      Set RSf = Server.CreateObject("ADODB.Recordset")
          RSf.ActiveConnection =  MM_EXTRANET_STRING
		  strSQL="SELECT count(numfac39) as nfac FROM ssfact39 WHERE  refcia39='"&referencia&"' "
		  RSf.Source = strSQL
          RSf.CursorType = 0
          RSf.CursorLocation = 2
          RSf.LockType = 1
          RSf.Open()
		  if not RSf.eof then
		    Contador_fact = RSf.Fields.Item("nfac").Value
		  end if
		  RSf.close
		  set RSf = nothing	
	
		  Set Conx = Server.CreateObject ("ADODB.Connection")
          Set RSfact = Server.CreateObject ("ADODB.RecordSet")
          Conx.Open MM_EXTRANET_STRING
	      SQLx="select * from ssfact39 where refcia39='"&referencia&"'  "
	      Set RSfact= Conx.Execute(SQLx)
          'response.Write(sqlx)
          'response.end
          Do while not RSfact.Eof
			 nfact= RSfact("numfac39")
			 if Contador_fact>1 then
			    numfactura=numfactura&"-"&nfact
			 else
			    numfactura=nfact
			 end if	
		 RSfact.MoveNext
         Loop
	 
		'***********************************************************************************
		
		'*****************************DATOS FILAS*****************************	
			%>
	   <BR>
       <table width="854"  border="1" cellspacing="3" cellpadding="3">
	   <tr>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Pedimento:</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(pedimento)%></font></td>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Patente</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(patente)%></font></td>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Vence</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(vence)%></font></td>
	   </tr>
	   <tr>
       <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Buque:</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(buque)%></font></td>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Regla</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write("12474")%></font></td>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Fact Prov:</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(numfactura)%></font></td>
	   </tr>
	   <tr>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Cliente:</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(cliente)%></font></td>
 <td Width=12% bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Fecha Mod.</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(fecpago)%></font></td>
	   </tr>
	   <tr>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Mercancia</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(mercancia)%></font></td>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Fecha ETA</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(fechaeta)%></font></td>
	   </tr>
       </table>
	   <BR>
	 	
        
	
<%	'------------------------------------fin de la 2 tabla ---------------

    '************************************partidas tce**********************
		Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RSPart = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING2
		strSQL=" select * from  tcepartidas where refe01='"&referencia&"' and frac01='"&frac&"'  "
	    Set RSPart= Connx.Execute(strSQL)
		if not RSPart.eof then 					
        Do while not RSPart.Eof
            partida=RSPart("parfra01")
			placas=RSPart("placas")
			pesoneto=RSPart("pesoneto")
			tipoveh=RSPart("tipveh")
			
			select case tipoveh
             case "C"
                  Tipotrans="Camion"
             case "T"
                  Tipotrans="Tolva"
             case "F"
                  Tipotrans="Furgon"
             case "G"
                  Tipotrans="Gondola"
           end select
			
			saldolib=saldolib+round(pesoneto)
            saldo=peso-saldolib
			tot_trailer=round(saldo/30000)
			tot_tolva=round(saldo/90000)
               
	     RSPart.MoveNext  ' de las partidas de tce
         Loop ' de las partidas de tce	
		 else
		   saldo=peso-saldolib
		   tot_trailer=round(saldo/30000)
		   tot_tolva=round(saldo/90000)
		 end if

       '------------------------------------para la segunda tablas    
	   %>
       <table width="854"  border="1" cellspacing="3" cellpadding="3">
       <tr>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Ton.Ped:</b></FONT></td>
	   <td Width=30% align="left"><font size="-1"><%RESPONSE.Write(tmpeso)%></font></td>
   <td bgcolor="#009999" align="center" colspan="2"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>No.copias Simples</b></FONT></td>
	   </tr>
	   <tr>
       <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Ton.Liberadas:</b></FONT></td>
       <%tmpsaldolib=trim(cstr(formatnumber(saldolib,0)))%>
	   <td Width=30% align="left" ><font size="-1"><%RESPONSE.Write(tmpsaldolib)%></font></td>
       <td Width=30% align="center"><font size="-1"><%RESPONSE.Write(tot_trailer)%></font></td>	  
	   <td Width=30% align="center"><font size="-1"><%RESPONSE.Write(tot_tolva)%></font></td>
	   </tr>
	   <tr>
	   <td bgcolor="#009999" ><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Saldo:</b></FONT></td>
	   <td align="left"><font size="-1">
	   <%tmpsaldo=trim(cstr(formatnumber(saldo,0)))%>
	     <%RESPONSE.Write(tmpsaldo)%>
	   </font></td>
 <td bgcolor="#009999" align="center"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Trailer</b></FONT></td>
 <td bgcolor="#009999" align="center"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Tolva</b></FONT></td>
	   </tr>
	   </table>
	   <br>

<%	   '*******************************************************
  
'--------------------------------------------------------------------------------- 
%>
	<table width="854" border="1" cellspacing="3" cellpadding="3">
         <tr>
           <th bgcolor="#009999">&nbsp;</th>
           <th bgcolor="#009999">&nbsp;</th>
           <th bgcolor="#009999">&nbsp;</th>
        <th bgcolor="#009999" colspan="2"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Transporte</FONT></th>
           <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Inicios</FONT></th>
           <th bgcolor="#009999">&nbsp;</th>
         </tr>
         <tr>
           <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Fecha</FONT></th>
           <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Partida</FONT></th>
           <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">No.ticket</FONT></th>
           <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Tipo</FONT></th>
           <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Placas</FONT></th>
           <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">FL</FONT></th>
           <th bgcolor="#009999"><font size="2" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Peso Neto</FONT></th>         
		 </tr>

<% '************************************partidas tce**********************
        saldolib=0			
		Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RSPart = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING2
		strSQL= " SELECT a.refe01 as ref,a.pedi01 as ped,a.nomcli01 as clie,a.patent01 as pat,a.fecpag01,b.frac01 as frac, " &_
          " b.desc01 as merc,a.fecpag01 as fecpag,a.pesobr01 as pesob, " &_ 
		  " c.consec01 as consec,c.refe01,c.fectick as fectick,c.ticket as ticket,c.tipveh as tipveh, " &_
		  " c.placas as placas,c.inicio as inicio,c.pesoneto as pesoneto  " &_
          " FROM pedimentos a,fracciones b,tcepartidas c " &_
          " where a.pedi01='"&pedimento&"' and a.refe01='"&referencia&"' and a.refe01=b.refe01 and a.refe01=c.refe01  " &_
		  " and c.consec01>0 "
	    Set RSPart= Connx.Execute(strSQL)
						
        Do while not RSPart.Eof
		    fechapartida=RSPart("fectick")
            partida=RSPart("consec")
		    ticket=RSPart("ticket") 			
			tipoveh=RSPart("tipveh")
			placas=RSPart("placas")
			inicios=RSPart("inicio")
			pesoneto=RSPart("pesoneto")
			saldolib=saldolib+round(pesoneto)
			
			select case tipoveh
             case "C"
                  Tipotrans="Camion"
             case "T"
                  Tipotrans="Tolva"
             case "F"
                  Tipotrans="Furgon"
             case "G"
                  Tipotrans="Gondola"
           end select
			
    '---------------------------TERCER TABLA PARTIDAS------------------
	%>

         <tr>
           <td><%RESPONSE.Write(fechapartida)%></td>
           <td><%RESPONSE.Write(partida)%></td>
           <td><%RESPONSE.Write(ticket)%></td>
           <td><%RESPONSE.Write(Tipotrans)%></td>
           <td><%RESPONSE.Write(placas)%></td>
           <td><%RESPONSE.Write(inicios)%></td>
           <td><%RESPONSE.Write(pesoneto)%></td>
         </tr>
		 
    

<%    RSPart.MoveNext  ' de las partidas de tce
         Loop ' de las partidas de tce	 
		 %>
		 <tr>
          <td colspan="6" align="right">Saldo Liberado</td> 
		   <td><%RESPONSE.Write(saldolib)%></td>
		 </tr>
	</table>  	 
<%'**********************************************************************

      ' RESPONSE.Write("</tr>")
        ' Refe.MoveNext 'avanza referencia  ---->
		Refe.MoveNext 
			   wend 'REFErencia
 else
%>
<tr>
  <th colspan=12>
    <font size="2" face="Arial">No se Encontro ningun registro con esos parametros</font>
  </th>
</tr>
<table>
<%'end if
end if 
 %>
</form>  


<%else
  response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>")
end if%>
<!--table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
    <tr>
    <td><%'=(strMenjError)
	%></td>
    </tr>
  </table-->
<%'end if 
%>
</BODY>
</HTML>