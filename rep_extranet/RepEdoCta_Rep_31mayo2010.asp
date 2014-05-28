<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% Server.ScriptTimeout=1500 %>
<HTML>
<HEAD>
<TITLE>:: REPORTE ESTADO DE CUENTA.... ::</TITLE>
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

          cve=request.form("cve")
          mov=request.form("mov")
          fi=trim(request.form("fi"))
          ff=trim(request.form("ff"))
          Vrfc=Request.Form("rfcCliente")
          Vckcve=Request.Form("ckcve")
          'Vclave=Request.Form("cveCliente")
          'aduan=request.form("aduana")
          if isdate(fi) and isdate(ff) then
                '---------------------------
                  DiaI = cstr(datepart("d",fi))
                  Mesi = cstr(datepart("m",fi))
                  AnioI = cstr(datepart("yyyy",fi))
                  DateI = Anioi&"/"&Mesi&"/"&Diai

                  DiaF = cstr(datepart("d",ff))
                  MesF = cstr(datepart("m",ff))
                  AnioF = cstr(datepart("yyyy",ff))
                  DateF = AnioF&"/"&MesF&"/"&DiaF

                 if not isdate(DateI) then
                      fec="1"
                 end if
                 if not  isdate(DateF) then
                       fec="1"
                 end if
		  END IF
                ' if fec<>"1" then
                    '---------------------------
                    'Antes de nada hay que instanciar el objeto Connection
                    'Set Conn = Server.CreateObject("ADODB.Connection")
                    'Una vez instanciado Connection lo podemos abrir y le asignamos la base de datos donde vamos a efectuar las operaciones
                    'Conn.Open "Mibase"
                    '''''''''''''
                    MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
                    Set oConn = Server.CreateObject ("ADODB.Connection")
                    Set RS = Server.CreateObject ("ADODB.RecordSet")
                    Set RS2 = Server.CreateObject ("ADODB.RecordSet")
                    Set RS3 = Server.CreateObject ("ADODB.RecordSet")
					Set Conn = Server.CreateObject ("ADODB.Connection")
                    Set REFE = Server.CreateObject ("ADODB.RecordSet")
                    oConn.Open MM_EXTRANET_STRING
                    '''''''''''''
                    'Response.Write(MM_EXTRANET_STRING)
                      'Response.End
                      Diferencia = DateDiff("D", DateI,DateF)
                      'if Diferencia>="0" then
                      '------------
                      if request.form("tipRep") = "2" then
                         Response.Addheader "Content-Disposition", "attachment;"
                         Response.ContentType = "application/vnd.ms-excel"
                      end if
                                    if mov="i" then
                                      if Vckcve="0" then
									  
			
sql1= " Select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
        " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
		" trim(concat(concat(concat(concat(concat(x.adusec01),'-'),x.patent01),'-'),x.numped01)) as pedimento, " &_
        " a.fdsp01 as fechadespacho,x.nomcli01,x.numped01, x.fecpag01 from ssdagi01 x,c01refer a,d31refer b,e31cgast c " &_
        " where x.rfccli01='"&Vrfc&"' and x.firmae01 is not null and x.firmae01 <> '' " &_
        " and x.refcia01=a.refe01 and  a.refe01=b.refe31 and b.cgas31=c.cgas31  and c.esta31='I'  and  c.esta31<>'C'   " &_
        " and c.fech31>='"&DateI&"' and c.fech31<='"&DateF&"'  "						  
                                          '   Response.Write("impo ck=0")
                                         '    Response.Write(sql1)
                                      else
									   'Response.Write("impo ck=1")
sql1= " Select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
        " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
		" trim(concat(concat(concat(concat(concat(x.adusec01),'-'),x.patent01),'-'),x.numped01)) as pedimento, " &_
        " a.fdsp01 as fechadespacho,x.nomcli01,x.numped01, x.fecpag01 from ssdagi01 x,c01refer a,d31refer b,e31cgast c " &_
        " where  x.firmae01 is not null and x.firmae01 <> '' " &_
        " and x.refcia01=a.refe01 and  a.refe01=b.refe31 and b.cgas31=c.cgas31  and c.esta31='I'  and  c.esta31<>'C'   " &_
        " and c.fech31>='"&DateI&"' and c.fech31<='"&DateF&"' " & permi & " "			  
							  
                                      end if
                                    else
                                      if Vckcve="0" then
                                        'Response.Write("EXpo ck=0")

sql1= " Select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
        " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
		" trim(concat(concat(concat(concat(concat(x.adusec01),'-'),x.patent01),'-'),x.numped01)) as pedimento, " &_
        " a.fdsp01 as fechadespacho,x.nomcli01,x.numped01, x.fecpag01 from ssdage01 x,c01refer a,d31refer b,e31cgast c " &_
        " where x.rfccli01='"&Vrfc&"' and x.firmae01 is not null and x.firmae01 <> '' " &_
        " and x.refcia01=a.refe01 and  a.refe01=b.refe31 and b.cgas31=c.cgas31  and c.esta31='I'  and  c.esta31<>'C'   " &_
        " and  c.fech31>='"&DateI&"' and c.fech31<='"&DateF&"'  "		                                        '     Response.Write(sql1)
                                      else
                                       ' Response.Write("EXpo ck=1")
sql1= " Select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
        " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
		" trim(concat(concat(concat(concat(concat(x.adusec01),'-'),x.patent01),'-'),x.numped01)) as pedimento, " &_
        " a.fdsp01 as fechadespacho,x.nomcli01,x.numped01, x.fecpag01 from ssdage01 x,c01refer a,d31refer b,e31cgast c " &_
        " where  x.firmae01 is not null and x.firmae01 <> '' " &_
        " and x.refcia01=a.refe01 and  a.refe01=b.refe31 and b.cgas31=c.cgas31  and c.esta31='I'  and  c.esta31<>'C'   " &_
        " and c.fech31>='"&DateI&"' and c.fech31<='"&DateF&"' " & permi & " "			  
                                      '       Response.Write(sql1)
                                      end if

                                    end if
                                     'Response.Write(sql1)
                                     'Response.End
                                    'Ejecutamos la orden
                                    set REFE=oConn.Execute(sql1)
									 'Response.Write(sql1)
                                    ' Response.End

                                    'Mostramos los registros
                    
                                   If strTipo=1 then 
    tipoper="Importacion"
 else	
    tipoper="Exportacion"
 end if  
  nomclie=REFE("nomcli01")
  %>		

<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>GRUPO REYES KURI, S.C.</p></font></strong>
<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>Cliente: <%=nomclie%></p></font></strong>
<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p></p></font></strong>
<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>Estado de Cuenta Del <%=DateI%> al <%=DateF%></p></font></strong>

   <table align="left" >
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>REFERENCIA</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>PEDIMENTO</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>FECHA</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>CUENTA</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>ANTICIPO</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>GASTOS</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>LIQUIDACIONES</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>SALDO</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>F.RECEPCION</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>VENCIDO</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>PLAZO</b></FONT></td>
  </tr> 
	 <%

          index=1
	      tempo=index
          if not REFE.eof then 
          While (NOT  REFE.EOF) 
	        referencia=REFE("referencia")
            pedimento=REFE("numped01")
		    fecpago=REFE("fecpag01")
		    cg=REFE("cg")
            anticipo1=REFE("ANTICIPO") 
			FechaCG=REFE("FechaCG")
		
'*************************************************ANTICIPOS*********************************************

       Set RSant = Server.CreateObject("ADODB.Recordset")
           RSant.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="Select IF (sum(mont11)>0,sum(mont11),0) as total_anticipos " &_
           " from d11movim where refe11='"&cg&"' and conc11='FA2' "
		   RSant.Source = strSQL
           RSant.CursorType = 0
           RSant.CursorLocation = 2
           RSant.LockType = 1
           RSant.Open()
		  if not RSant.eof then
		    anticipo2 = RSant.Fields.Item("total_anticipos").Value
		  end if
		   RSant.close
		  set RSant = nothing	
		  
		    Set RSant = Server.CreateObject("ADODB.Recordset")
           RSant.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="Select IF (sum(mont11)>0,sum(mont11),0) as total_anticipos_can " &_
           " from d11movim where refe11='"&cg&"' and conc11='CF2' "
		   RSant.Source = strSQL
           RSant.CursorType = 0
           RSant.CursorLocation = 2
           RSant.LockType = 1
           RSant.Open()
		  if not RSant.eof then
		    anticipo_cancelado = RSant.Fields.Item("total_anticipos_can").Value
		  end if
		   RSant.close
		  set RSant = nothing	
		  anticipo_tmp=anticipo2-anticipo_cancelado
		  if anticipo_tmp<>0 then
		     anticipo_final=anticipo_tmp
		  end if
		  IF anticipo_final=anticipo1 THEN
		     ANTICIPO=anticipo2
		  ELSE
		     ANTICIPO=anticipo2
			' ANTICIPO="REVISAR"  ' CUAL ES MAS CONFIABLE EL DE D11MOVIM O E31CGAST
		  END IF 
		  	

		
'*************************************************GASTOS*********************************************

       Set RSgastos = Server.CreateObject("ADODB.Recordset")
           RSgastos.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="SELECT sum(mont11) as totalgastos FROM e31cgast a,d11movim b " &_
           "where a.cgas31='"&cg&"' and a.cgas31=b.refe11  and conc11='FA1' "
		   RSgastos.Source = strSQL
           RSgastos.CursorType = 0
           RSgastos.CursorLocation = 2
           RSgastos.LockType = 1
           RSgastos.Open()
		  if not RSgastos.eof then
		    gastos = RSgastos.Fields.Item("totalgastos").Value
		  end if
		  RSgastos.close
		  set RSgastos = nothing	 
		
		   Set RSgastos2 = Server.CreateObject("ADODB.Recordset")
           RSgastos2.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="SELECT IF (sum(mont11)>0,sum(mont11),0) AS total_can "&_
           "FROM e31cgast a,d11movim b where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and conc11='CF1' " 
		   RSgastos2.Source = strSQL
           RSgastos2.CursorType = 0
           RSgastos2.CursorLocation = 2
           RSgastos2.LockType = 1
           RSgastos2.Open()
		  if not RSgastos2.eof then
		     gastos_cancelados = RSgastos2.Fields.Item("total_can").Value
		  end if
		  RSgastos2.close
		  set RSgastos2 = nothing	
		  
		  gastos_tmp=gastos-gastos_cancelados
		  if gastos_tmp=gastos then
		     gastos=gastos
		  else
		     gastos=gastos_tmp
 		  end if
		  
		  gastos_SinNotas=gastos
		  'para SCA Y SCR
		  Set RSgastos3 = Server.CreateObject("ADODB.Recordset")
           RSgastos3.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="SELECT IF (sum(mont11)>0,sum(mont11),0) AS total_notacargo " &_
           "FROM e31cgast a,d11movim b where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and conc11='SCA'  " 
		   RSgastos3.Source = strSQL
           RSgastos3.CursorType = 0
           RSgastos3.CursorLocation = 2
           RSgastos3.LockType = 1
           RSgastos3.Open()
		  if not RSgastos3.eof then
		     gastos_SCA = RSgastos3.Fields.Item("total_notacargo").Value
		  end if
		  RSgastos3.close
		  set RSgastos3 = nothing	
		  
		   Set RSgastos4 = Server.CreateObject("ADODB.Recordset")
           RSgastos4.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="SELECT IF (sum(mont11)>0,sum(mont11),0) AS total_notacredito " &_
           "FROM e31cgast a,d11movim b where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and conc11='SCR'  " 
		   RSgastos4.Source = strSQL
           RSgastos4.CursorType = 0
           RSgastos4.CursorLocation = 2
           RSgastos4.LockType = 1
           RSgastos4.Open()
		  if not RSgastos4.eof then
		     gastos_SCR = RSgastos4.Fields.Item("total_notacredito").Value
		  end if
		  RSgastos4.close
		  set RSgastos4 = nothing	
		  
		  if gastos_SCA>0 then
		     ban_gastos_sca=1
		  else	 
		     ban_gastos_sca=0
		  end if	 
		  
		   if gastos_SCR>0 then
		     ban_gastos_scr=1
		  else	 
		     ban_gastos_scr=0
		  end if	 

		  gastos=gastos+gastos_SCA
		  gastos=gastos-gastos_SCR
		  
		  Set RSgastos5 = Server.CreateObject("ADODB.Recordset")
           RSgastos5.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="SELECT COUNT(CONC11) AS NNOTAS FROM e31cgast a,d11movim b " &_
           " where a.cgas31='"&cg&"' and a.cgas31=b.refe11  and (conc11='SCA' OR conc11='SCR') "
		   RSgastos5.Source = strSQL
           RSgastos5.CursorType = 0
           RSgastos5.CursorLocation = 2
           RSgastos5.LockType = 1
           RSgastos5.Open()
		  if not RSgastos5.eof then
		     nnotas = RSgastos5.Fields.Item("NNOTAS").Value
		  end if
		  RSgastos5.close
		  set RSgastos5 = nothing	
		  
		  

'***********************************LIQUIDACIONES*************************************************************

          Set Conx = Server.CreateObject ("ADODB.Connection")
          Set rsliq = Server.CreateObject ("ADODB.RecordSet")
          Conx.Open MM_EXTRANET_STRING
	      SQLx=" SELECT IF (sum(mont11)>0,sum(mont11),0) as total_liq  FROM e31cgast a,d11movim b " &_ 
          " where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and b.conc11='LIQ' and " &_
		  " fech11>='"&DateI&"'  and fech11<='"&DateF&"' "
	      Set rsliq= Conx.Execute(SQLx)
          'response.Write(sqlx)
          'response.end
          Do while not rsliq.Eof
             liquidaciones = rsliq("total_liq")
		  rsliq.MoveNext
          Loop
'************************************DEVOLUCIONES********************************************************
           total_dev=0
		   Set RSdev = Server.CreateObject("ADODB.Recordset")
           RSdev.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL= "SELECT IF (sum(mont11)>0,sum(mont11),0) AS total_dev " &_
           " FROM e31cgast a,d11movim b where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and conc11='DEV' " 
		   RSdev.Source = strSQL
           RSdev.CursorType = 0
           RSdev.CursorLocation = 2
           RSdev.LockType = 1
           RSdev.Open()
		  if not RSdev.eof then
		     total_dev = RSdev.Fields.Item("total_dev").Value
		  end if
		  RSdev.close
		  set RSdev = nothing		  
		  
		  
		  

'******************************************PLAZO**********************************************************
                ahora = now()
                fecha_actual=FormatDateTime(ahora,2)
                   
                xDFpag=FechaCG
                if isdate(xDFpag) then
                   DiaF = cstr(datepart("d",FechaCG))
                   MesF = cstr(datepart("m",FechaCG))
                   AnioF = cstr(datepart("yyyy",FechaCG))
                   DateFechaCG = DiaF & "/" &MesF & "/"& AnioF
                end if
				TimePlazo=DateDiff("d",DateFechaCG,fecha_actual)

'*****************************************SALDOS*************************************************************
       IF ANTICIPO<>"REVISAR" THEN 
	      if total_dev=0 then
           'SALDOS=formatnumber((ANTICIPO-(GASTOS-LIQUIDACIONES)),2)
		  SALDOS=formatnumber((ANTICIPO+LIQUIDACIONES-(GASTOS)),2)
		  saldos=formatnumber((saldos-total_dev),2)
		  else
          'SALDOS=formatnumber((ANTICIPO-(GASTOS-LIQUIDACIONES)),2)
		  SALDOS=formatnumber((ANTICIPO+LIQUIDACIONES-(GASTOS)),2)
  		  saldos=formatnumber((saldos-total_dev),2)
		  'response.Write("CUENTA:"&CG&"--"&anticipo&"-"&gastos&"-"&liquidaciones&"-"&total_dev)
		  'response.Write("<br>")
		  end if 
       ELSE
	      SALDOS="REVISAR"
	   END IF 
'*****************************************VENCIDO*****************************************************************	   
	   IF TimePlazo>70 THEN
	      if saldos>1 then
	         vencido=saldos
		     else
			 vencido=saldos*(-1)  
			 end if
	   else	   
	      vencido="--"
	   END IF    
	   		  
		  
'*****************************DATOS FILAS*****************************
		  if SALDOS<>0 then 
		  'response.Write("referencia:"&referencia&" SAldo"&SALDOS&"<br>")
		  'quite saldos %>
	     <tr>
         <td><font size="-1"><%RESPONSE.Write(referencia)%></font></td>
         <td align="center"><font size="-1"><%RESPONSE.Write(pedimento)%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(fecpago)%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(cg)%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(anticipo)%></font></td>
 		 <td align="center"><font size="-1"><%RESPONSE.Write(gastos_SinNotas)%></font></td>
  		 <td align="center"><font size="-1"><%RESPONSE.Write(liquidaciones)%></font></td>
		 <% if (ban_gastos_sca=0 and ban_gastos_scr=0) then 
		       if saldos>1 then %>
   		 <td align="center"><font size="-1"><%RESPONSE.Write(saldos)%></font></td>
		 <%    else
		       saldos=saldos*(-1)  %>
			   <td align="center"><font size="-1"><%RESPONSE.Write(saldos)%></font></td>
			   <%end if %>
		 <% else  %>
   		 <td align="center"><font size="-1"><%RESPONSE.Write("--")%></font></td>
		 <% end if%>
		 <td align="center"><font size="-1"><%RESPONSE.Write(FechaCG)%></font></td>
		 <% if (ban_gastos_sca=0 and ban_gastos_scr=0) then %>
	     <td align="center"><font size="-1"><%RESPONSE.Write(vencido)%></font></td>
 		 <% else  %>
   		 <td align="center"><font size="-1"><%RESPONSE.Write("--")%></font></td>
		 <% end if%>
		 <td align="center"><font size="-1"><%RESPONSE.Write(TimePlazo)%></font></td>
 		 
		 <% 		 
		  

'**********************************************************************
   	      RESPONSE.Write("</tr>")
          
'*********************************para las notas de credito************************************		 
          cn=1
		  num_nota=1
          if  (ban_gastos_sca=1 or ban_gastos_scr=1) then
		  Set Conx = Server.CreateObject ("ADODB.Connection")
          Set rsliq = Server.CreateObject ("ADODB.RecordSet")
          Conx.Open MM_EXTRANET_STRING
	      SQLx=" SELECT mont11,fech11,foli11,refe11,conc11 FROM e31cgast a,d11movim b " &_
		  " where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and (conc11='SCA' or conc11='SCR') "
	      Set rsliq= Conx.Execute(SQLx)
          'response.Write(sqlx)
          'response.end
		  Do while not rsliq.Eof
		    
		     fecha_nota = rsliq("fech11")
		     gto_nota = rsliq("mont11")
			 concepto=rsliq("conc11")
		     if  conc11="SCR" then  
                 ban_gastos_scr=1
			 else
			   if conc11="SCA" then  
			      ban_gastos_sca=1	  
			   end if	   
			 end if
		 '******************************************PLAZO**********************************************************
                ahora = now()
                fecha_actual=FormatDateTime(ahora,2)
                   
                xDFnota=fecha_nota
                if isdate(xDFnota) then
                   DiaF = cstr(datepart("d",fecha_nota))
                   MesF = cstr(datepart("m",fecha_nota))
                   AnioF = cstr(datepart("yyyy",fecha_nota))
                   DateFechaNota = DiaF & "/" &MesF & "/"& AnioF
                end if
				TimePlazo2=DateDiff("d",DateFechaNota,fecha_actual)
          
		  '*****************************************SALDOS adicionando las NOTAS**************************************************
          IF ANTICIPO<>"REVISAR" THEN
		     if concepto="SCA" then  'notas de cargo se suman  
		        gasto_notas=gastos_SinNotas+gto_nota
			 else 
			    if concepto="SCR" then
				   gasto_notas=gastos_SinNotas-gto_nota
				end if
			 end if 
			 
			 'if ban_gastos_sca=1 then  'notas de cargo se suman  
'		        gasto_notas=gastos_SinNotas+gto_nota
'			 else 
'			    if ban_gastos_scr=1 then
'				   gasto_notas=gastos_SinNotas-gto_nota
'				end if
'			 end if 
		    ' response.Write("referencia:"&referencia&" Gastos conNota"&gasto_notas&" Bandera credito"&ban_gastos_scr&" Ban Cargo"&ban_gastos_sca&"<br>")
	         if total_dev=0 then
			    SALDOS_connota=formatnumber((ANTICIPO+LIQUIDACIONES-(gasto_notas)),2)
		        saldos_connota=formatnumber((SALDOS_connota-total_dev),2)
		     else
			    SALDOS_connota=formatnumber((ANTICIPO+LIQUIDACIONES-(gasto_notas)),2)
  		        saldos_connota=formatnumber((SALDOS_connota-total_dev),2)
		     end if 
          ELSE
	        SALDOS_connota="REVISAR"
	      END IF 
		  
		  '************************************************************************************************************
		  
		  ''while cn<=nnotas
		    if gto_nota<>0 then%>
		 <tr>
         <td><font size="-1"><%RESPONSE.Write(referencia)%></font></td>
         <td align="center"><font size="-1"><%RESPONSE.Write("")%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(fecpago)%></font></td>
		 <% if concepto="SCA" then %> 
		 <td align="center"><font size="-1"><%RESPONSE.Write(cg&" Nota Cargo")%></font></td>
		 <% else
		      if concepto="SCR" then%>
		 <td align="center"><font size="-1"><%RESPONSE.Write(cg&" Nota Credito")%></font></td>		 
		 <%   end if
		    end if%>
		 <td align="center"><font size="-1"><%RESPONSE.Write("--")%></font></td>
		 <% if ban_gastos_sca=1 then %> 
 		 <td align="center"><font size="-1"><%RESPONSE.Write(gto_nota)%></font></td>
		 <% else %>
		 <td align="center"><font size="-1"><%RESPONSE.Write("-"&gto_nota)%></font></td>
		 <% end if%>
  		 <td align="center"><font size="-1"><%RESPONSE.Write("--")%></font></td>
		 
		  <% if num_nota=nnotas then %> 
 		    <% if SALDOS>1 then%>
		      <td align="center"><font size="-1"><%RESPONSE.Write(SALDOS)%></font></td>
		    <%else 'para eliminar el negativo
		     SALDOS=SALDOS*(-1)%>
   		       <td align="center"><font size="-1"><%RESPONSE.Write(SALDOS)%></font></td>
		    <%end if%>	
		 <% else %>
		 <td align="center"><font size="-1"><%RESPONSE.Write("--")%></font></td>
		 <% end if%>
		 
		 <td align="center"><font size="-1"><%RESPONSE.Write(fecha_nota)%></font></td>
		
		 <% if num_nota=nnotas then %> 
 		    <% if SALDOS>1 then%>
		      <td align="center"><font size="-1"><%RESPONSE.Write(SALDOS)%></font></td>
		    <%else 'para eliminar el negativo
		     SALDOS=SALDOS*(-1)%>
   		       <td align="center"><font size="-1"><%RESPONSE.Write(SALDOS)%></font></td>
		    <%end if%>	
		 <% else %>
		 <td align="center"><font size="-1"><%RESPONSE.Write("--")%></font></td>
		 <% end if%>
		 
 		 <td align="center"><font size="-1"><%RESPONSE.Write(TimePlazo2)%></font></td>
 		<!--<td align="center"><font size="-1">'RESPONSE.Write(total_dev)</font></td>-->
		</tr>
<%         ''cn=cn+1
           ''wend
		   end if
		    rsliq.MoveNext
					  num_nota=num_nota+1
		  'response.Write("NumeroNota:"&num_nota&"<br>")
          Loop
		  end if 
         end if  'del if SALDOS<>0 then
'**********************************************************************************************		 
      Refe.MoveNext 'avanza referencia  ---->
		  index=index+1
      wend 'REFErencia
  else
%>
<tr>
  <th colspan=12>
    <font size="1" face="Arial">No se Encontro ningun registro con esos parametros
  </th>
</tr>
<table>


    <%'else
      'strMenjError = "No tiene Autorización para visualizar este reporte"
	  END IF
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