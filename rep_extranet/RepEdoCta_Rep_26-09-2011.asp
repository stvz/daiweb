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
          'if isdate(fi) and isdate(ff) then
		   if isdate(ff) then
                '---------------------------
'                  DiaI = cstr(datepart("d",fi))
'                  Mesi = cstr(datepart("m",fi))
'                  AnioI = cstr(datepart("yyyy",fi))
'                  DateI = Anioi&"/"&Mesi&"/"&Diai

                  DiaF = cstr(datepart("d",ff))
                  MesF = cstr(datepart("m",ff))
                  AnioF = cstr(datepart("yyyy",ff))
                  DateF = AnioF&"/"&MesF&"/"&DiaF

                 'if not isdate(DateI) then
                      'fec="1"
                 'end if
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
'*******************************************************************************************
if Vckcve="0" then

          Set RSclicve= Server.CreateObject("ADODB.Recordset")
           RSclicve.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL= "select count(cvecli18) as totcvecli from ssclie18  where rfccli18='"&Vrfc&"'  " 
		   RSclicve.Source = strSQL
           RSclicve.CursorType = 0
           RSclicve.CursorLocation = 2
           RSclicve.LockType = 1
           RSclicve.Open()
		  if not RSclicve.eof then
		     total_cves= RSclicve.Fields.Item("totcvecli").Value
		  end if
		  RSclicve.close
		  set RSclicve = nothing		 

          ixc=1
		  Set Conx = Server.CreateObject ("ADODB.Connection")
          Set rsclie = Server.CreateObject ("ADODB.RecordSet")
          Conx.Open MM_EXTRANET_STRING
	      SQLx=" select * from ssclie18  where rfccli18='"&Vrfc&"'  "
	      Set rsclie= Conx.Execute(SQLx)
          'response.Write(sqlx)
          'response.end
		  Do while not rsclie.Eof
		  cvecli18= rsclie("cvecli18")
		   if ixc=1 then
			   cvecliente=cvecli18
			else
		       cvecliente=cvecliente&","&cvecli18  
		  end if   
	      ixc=ixc+1
		  rsclie.MoveNext
		  Loop
		  cvecliente2="("&cvecliente&")"
          'response.Write(cvecliente2)
		  'response.End()
		  end if 'fin del if Vckcve="0" then
'************************************************************************************************
					                                      
                                      if Vckcve="0" then
									  
''sql1= " Select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
''      " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
''      " trim(concat(concat(concat(concat(concat(x.adusec01),'-'),x.patent01),'-'),x.numped01)) as pedimento, " &_
''      " a.fdsp01 as fechadespacho,x.nomcli01,x.numped01, x.fecpag01 from ssdagi01 x,c01refer a,d31refer b,e31cgast c, " &_
''	  " d11movim d  where x.rfccli01='"&Vrfc&"' and x.firmae01<> '' and x.firmae01 <> '' " &_
''      " and x.refcia01=a.refe01 and  a.refe01=b.refe31 and c.cgas31=d.refe11 and b.cgas31=c.cgas31  and c.esta31='I'  " &_
''	  " and  c.esta31<>'C' and d.fech11<='"&DateF&"'  group by referencia " &_
''      " union all " &_
''      " Select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
''      " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
''      " trim(concat(concat(concat(concat(concat(x.adusec01),'-'),x.patent01),'-'),x.numped01)) as pedimento, " &_
''      " a.fdsp01 as fechadespacho,x.nomcli01,x.numped01, x.fecpag01 from ssdage01 x,c01refer a,d31refer b,e31cgast c, " &_
''	  " d11movim d  where x.rfccli01='"&Vrfc&"' and x.firmae01 <> ''  and x.firmae01 <> '' " &_
''      " and x.refcia01=a.refe01 and  a.refe01=b.refe31 and c.cgas31=d.refe11 and b.cgas31=c.cgas31  and c.esta31='I'  " &_
''	  " and  c.esta31<>'C' and d.fech11<='"&DateF&"'  group by referencia " 

sql1= " Select i.refcia01 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
      " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
	  " trim(concat(concat(concat(concat(concat(i.adusec01),'-'),i.patent01),'-'),i.numped01)) as pedimento, " &_
	  " ref.fdsp01 as fechadespacho,i.nomcli01,i.numped01, i.fecpag01,'Importacion' as Tipoper  " &_
      " FROM ssdagi01 as i " &_
      " inner join ssclie18 as cli on cli.cvecli18 = i.cvecli01 " &_
      " inner join c01refer as ref on ref.refe01 = i.refcia01 " &_
      " inner join d31refer as r on r.refe31 = i.refcia01 " &_
      " inner join e31cgast as c on c.cgas31 = r.cgas31 " &_
      " left join D11MOVIM AS d on  d.cgas11 = r.cgas31   " &_
      " left join E11Movim AS e ON d.Foli11 = e.Foli11 " &_ 
      " where i.rfccli01='"&Vrfc&"'  " &_
	  " and d.fech11<='"&DateF&"' and d.ccli11 in  " &cvecliente2& " group by cg " &_
	  " union all " &_
	  " Select ex.refcia01 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
      " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
	  " trim(concat(concat(concat(concat(concat(ex.adusec01),'-'),ex.patent01),'-'),ex.numped01)) as pedimento, " &_
	  " ref.fdsp01 as fechadespacho,ex.nomcli01,ex.numped01, ex.fecpag01,'Exportacion' as Tipoper  " &_
      " FROM ssdage01 as ex " &_
      " inner join ssclie18 as cli on cli.cvecli18 = ex.cvecli01 " &_
      " inner join c01refer as ref on ref.refe01 = ex.refcia01 " &_
      " inner join d31refer as r on r.refe31 = ex.refcia01 " &_
      " inner join e31cgast as c on c.cgas31 = r.cgas31 " &_
      " left join D11MOVIM AS d on  d.cgas11 = r.cgas31  " &_
      " left join E11Movim AS e ON d.Foli11 = e.Foli11 " &_ 
      " where ex.rfccli01='"&Vrfc&"'  " &_
	  " and d.fech11<='"&DateF&"' and d.ccli11 in  " &cvecliente2& " group by cg order by cg " 
''	
	
                                         '   Response.Write("impo ck=0")
                                         '   Response.Write(sql1)
                                      else
									   'Response.Write("impo ck=1")
									   
'sql1= " Select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
'      " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
'      " trim(concat(concat(concat(concat(concat(x.adusec01),'-'),x.patent01),'-'),x.numped01)) as pedimento, " &_
'      " a.fdsp01 as fechadespacho,x.nomcli01,x.numped01, x.fecpag01 from ssdagi01 x,c01refer a,d31refer b,e31cgast c, " &_
'	  " d11movim d  where x.rfccli01='"&Vrfc&"' and x.firmae01<> '' and x.firmae01 <> '' " &_
'      " and x.refcia01=a.refe01 and  a.refe01=b.refe31 and c.cgas31=d.refe11 and b.cgas31=c.cgas31  and c.esta31='I'  " &_
'	  " and  c.esta31<>'C' and d.fech11<='"&DateF&"' " & permi & "  group by referencia " &_
'      " union all " &_
'      " Select b.refe31 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
'      " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
'      " trim(concat(concat(concat(concat(concat(x.adusec01),'-'),x.patent01),'-'),x.numped01)) as pedimento, " &_
'      " a.fdsp01 as fechadespacho,x.nomcli01,x.numped01, x.fecpag01 from ssdage01 x,c01refer a,d31refer b,e31cgast c, " &_
'	  " d11movim d  where  x.firmae01 <> ''  and x.firmae01 <> '' " &_
'      " and x.refcia01=a.refe01 and  a.refe01=b.refe31 and c.cgas31=d.refe11 and b.cgas31=c.cgas31  and c.esta31='I'  " &_
'	  " and  c.esta31<>'C' and d.fech11<='"&DateF&"' " & permi & "  group by referencia " 
	  
  sql1= " Select i.refcia01 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
      " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
	  " trim(concat(concat(concat(concat(concat(i.adusec01),'-'),i.patent01),'-'),i.numped01)) as pedimento, " &_
	  " ref.fdsp01 as fechadespacho,i.nomcli01,i.numped01, i.fecpag01,'Importacion' as Tipoper  " &_
      " FROM ssdagi01 as i " &_
      " inner join ssclie18 as cli on cli.cvecli18 = i.cvecli01 " &_
      " inner join c01refer as ref on ref.refe01 = i.refcia01 " &_
      " inner join d31refer as r on r.refe31 = i.refcia01 " &_
      " inner join e31cgast as c on c.cgas31 = r.cgas31 " &_
      " left join D11MOVIM AS d on  d.cgas11 = r.cgas31 and d.ccli11 = cli.cvecli18  " &_
      " left join E11Movim AS e ON d.Foli11 = e.Foli11 " &_ 
      " where  d.fech11<='"&DateF&"'  " & permi & " and d.ccli11="&strFiltroCliente&" group by cg " &_
	  " union all " &_
	  " Select ex.refcia01 as referencia,c.fech31 as FechaCG,c.clie31 as CVECLIENTE, c.CGAS31 AS cg, c.chon31 as HONORARIOS, " &_
      " c.SUPH31 AS PH, c.csce31 as SC,c.anti31 as ANTICIPO, c.sald31 as Saldo, c.tota31 as TOTAL, " &_
	  " trim(concat(concat(concat(concat(concat(ex.adusec01),'-'),ex.patent01),'-'),ex.numped01)) as pedimento, " &_
	  " ref.fdsp01 as fechadespacho,ex.nomcli01,ex.numped01, ex.fecpag01,'Exportacion' as Tipoper  " &_
      " FROM ssdage01 as ex " &_
      " inner join ssclie18 as cli on cli.cvecli18 = ex.cvecli01 " &_
      " inner join c01refer as ref on ref.refe01 = ex.refcia01 " &_
      " inner join d31refer as r on r.refe31 = ex.refcia01 " &_
      " inner join e31cgast as c on c.cgas31 = r.cgas31 " &_
      " left join D11MOVIM AS d on  d.cgas11 = r.cgas31 and d.ccli11 = cli.cvecli18  " &_
      " left join E11Movim AS e ON d.Foli11 = e.Foli11 " &_ 
      " where  d.fech11<='"&DateF&"' " & permi & " and d.ccli11="&strFiltroCliente&" group by cg order by cg" 
	  
                                    end if
                                     'Response.Write(sql1)
                                     'Response.End
                                    'Ejecutamos la orden
                                    set REFE=oConn.Execute(sql1)
									'Response.Write(sql1)
                                    'Response.End
Dim saldoscuenta(20000)
Dim saldoscuenta2(20000)
Dim cgrepetidas(20000)
Dim saldosvencidos(20000)
                                    'Mostramos los registros
                    
                                   If strTipo=1 then 
    tipoper="Importacion"
 else	
    tipoper="Exportacion"
 end if  
  if not REFE.eof then 
  nomclie=REFE("nomcli01")
  end if
  %>		

<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>GRUPO REYES KURI, S.C.</p></font></strong>
<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>Cliente: <%=nomclie%></p></font></strong>
<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p></p></font></strong>
<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>Estado de Cuenta Al <%=DateF%></p></font></strong>

   <table align="left" >
   <tr>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>REFERENCIA</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>PEDIMENTO</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>TIPO</b></FONT></td>
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
		  TOTAL_SALDOS=0
		  SUMA_SALDO=0
	      tempo=index
		  isaldos=0
		   ntotal2=0
		   ntotal4=0
          if not REFE.eof then 
          While (NOT  REFE.EOF) 
		  
	        referencia=REFE("referencia")
            pedimento=REFE("pedimento")
		    fecpago=REFE("fecpag01")
		    cg=REFE("cg")
            anticipo1=REFE("ANTICIPO") 
			FechaCG=REFE("FechaCG")
			Tipo_operacion=REFE("Tipoper")
		
'*************************************************ANTICIPOS*********************************************

       Set RSant = Server.CreateObject("ADODB.Recordset")
           RSant.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="Select IF (sum(mont11)>0,sum(mont11),0) as total_anticipos " &_
           " from d11movim where refe11='"&cg&"' and conc11='FA2' and fech11<='"&DateF&"' "
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
           " from d11movim where refe11='"&cg&"' and conc11='CF2' and fech11<='"&DateF&"' "
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
			 else
			 anticipo_final=anticipo_tmp
		  end if
		  'IF anticipo_final=anticipo1 THEN
		     'ANTICIPO=anticipo2
		 ' ELSE
		     'ANTICIPO=anticipo2
			' ANTICIPO="REVISAR"  ' CUAL ES MAS CONFIABLE EL DE D11MOVIM O E31CGAST
		 ' END IF 
		  	

		
'*************************************************GASTOS*********************************************

       Set RSgastos = Server.CreateObject("ADODB.Recordset")
           RSgastos.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="SELECT sum(mont11) as totalgastos FROM e31cgast a,d11movim b " &_
           "where a.cgas31='"&cg&"' and a.cgas31=b.refe11  and conc11='FA1' and fech11<='"&DateF&"' "
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
		  gastos_fa1=gastos
		  
		   Set RSgastos2 = Server.CreateObject("ADODB.Recordset")
           RSgastos2.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="SELECT IF (sum(mont11)>0,sum(mont11),0) AS total_can "&_
           "FROM e31cgast a,d11movim b where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and conc11='CF1' and fech11<='"&DateF&"' " 
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
		   gastos_CF1=gastos_cancelados
		  
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
           "FROM e31cgast a,d11movim b where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and conc11='SCA' and fech11<='"&DateF&"' " 
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
           "FROM e31cgast a,d11movim b where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and conc11='SCR' and fech11<='"&DateF&"' " 
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
           " where a.cgas31='"&cg&"' and a.cgas31=b.refe11  and (conc11='SCA' OR conc11='SCR') and fech11<='"&DateF&"' "
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
		  " fech11<='"&DateF&"' "
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
           " FROM e31cgast a,d11movim b where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and conc11='DEV' and fech11<='"&DateF&"' " 
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
		  devoluciones_temp=total_dev	  
		  
		  
		  

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

          if isnull(anticipo_tmp) then
             anticipo_tmp=0
		  end if 	
		  if isnull(LIQUIDACIONES) then
             LIQUIDACIONES=0
		  end if 
		  if isnull(GASTOS) then
             GASTOS=0
		  end if 
		  if isnull(total_dev) then
             total_dev=0
		  end if 

      ' IF ANTICIPO<>"REVISAR" THEN 
	      if total_dev=0 then
           'SALDOS=formatnumber((ANTICIPO-(GASTOS-LIQUIDACIONES)),2)
		  SALDOS=formatnumber((anticipo_tmp+LIQUIDACIONES-(GASTOS)),2)
		  saldos=formatnumber((saldos-total_dev),2)
		  else
          'SALDOS=formatnumber((ANTICIPO-(GASTOS-LIQUIDACIONES)),2)
		  SALDOS=formatnumber((anticipo_tmp+LIQUIDACIONES-(GASTOS)),2)
  		  saldos=formatnumber((saldos-total_dev),2)
		  'response.Write("CUENTA:"&CG&"--"&anticipo&"-"&gastos&"-"&liquidaciones&"-"&total_dev)
		  'response.Write("<br>")
		  end if 
     '  ELSE
	     ' SALDOS="REVISAR"
	  ' END IF 
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
	   		  
		   'TOTAL_SALDOS=(TOTAL_SALDOS)-(SALDOS)
		   'saldoscuenta(isaldos)=saldos
'****************************************NUEVA MODIFICACION****************************************************************		
' SACAR LOS SALDOS Y ELIMINAR LOS SALDO ACREEDORES(-)
          
		  
		  SUMA_FA2_ANTICIPOS=anticipo2-anticipo_cancelado
		  SUMA_FA1_GASTOS=gastos_fa1-gastos_cf1
		  
		   if isnull(SUMA_FA1_GASTOS) then
             SUMA_FA1_GASTOS=0
		  end if 
          if isnull(SUMA_FA2_ANTICIPOS) then
             SUMA_FA2_ANTICIPOS=0
		  end if 	
		  if isnull(LIQUIDACIONES) then
             LIQUIDACIONES=0
		  end if 
		  if isnull(GASTOS) then
             GASTOS=0
		  end if 
		  if isnull(devoluciones_temp) then
             devoluciones_temp=0
		  end if 
		  if isnull(gastos_SCA) then
             gastos_SCA=0
		  end if 
		  if isnull(gastos_SCR) then
             gastos_SCR=0
		  end if 


saldo_antgto=formatnumber(SUMA_FA1_GASTOS-SUMA_FA2_ANTICIPOS,2)
saldo_temporal=formatnumber((saldo_antgto)+(devoluciones_temp+gastos_SCA)-(LIQUIDACIONES+gastos_SCR),2)

SUMA_SALDO=formatnumber((saldo_temporal),2)

'***********************************************para la misma cuentas de gastos q tienen 2 referencias*****************

 Set RScgastos6 = Server.CreateObject("ADODB.Recordset")
           RScgastos6.ActiveConnection =  MM_EXTRANET_STRING
		   strSQL="select COUNT(cgas31) as nocuentasg FROM d31refer where cgas31='"&cg&"' " 
		   RScgastos6.Source = strSQL
           RScgastos6.CursorType = 0
           RScgastos6.CursorLocation = 2
           RScgastos6.LockType = 1
           RScgastos6.Open()
		  if not RScgastos6.eof then
		     nocuentascg = RScgastos6.Fields.Item("nocuentasg").Value
		  end if
		  RScgastos6.close
		  set RScgastos6 = nothing	
		  		   
		  if nocuentascg>1 then
		     ban_gastos_ncg=1
			 cgrepetidas(isaldos)=cg
		 else	 
		     ban_gastos_ncg=0
			 cgrepetidas(isaldos)=0
		 end if	
		 
		 if isaldos>0 then
		    if cg=cgrepetidas(isaldos-1) then
			   bander=1 
			else 
			   bander=0
			end if   
		 end if   
'***************SUMA DE TOTAL VENCIDO******************************************************

       IF TimePlazo>70 THEN
	      if saldo_temporal>1 then
	           saldo_vencido=saldo_temporal
		     else
			   saldo_vencido=saldo_temporal  
			 end if
	      else	   
	         saldo_vencido=0
	   END IF    		 
		 
'*****************************DATOS FILAS**********************************************************************************
		  if SALDOS<>0 and saldo_temporal>0 and bander=0  then
     	  'response.Write("referencia:"&referencia&" SAldo"&SALDOS&"<br>")
		  'quite saldos %>
	     <tr>
         <td><font size="-1"><%RESPONSE.Write(referencia)%></font></td>
         <td align="center"><font size="-1"><%RESPONSE.Write(pedimento)%></font></td>
         <td align="center"><font size="-1"><%RESPONSE.Write(Tipo_operacion)%></font></td> 		 
		 <td align="center"><font size="-1"><%RESPONSE.Write(fecpago)%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(cg)%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(anticipo_tmp)%></font></td>
 		 <td align="center"><font size="-1"><%RESPONSE.Write(gastos_SinNotas)%></font></td>
  		 <td align="center"><font size="-1"><%RESPONSE.Write(liquidaciones)%></font></td>
		 <% if (ban_gastos_sca=0 and ban_gastos_scr=0) then 
		        TOTAL_SALDOS=(TOTAL_SALDOS)-(SALDOS)
				saldoscuenta(isaldos)=saldos
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
		  
		  <% if (ban_gastos_sca=0 and ban_gastos_scr=0) then %>
 		<!--  <td align="center"><font size="-1"><%'RESPONSE.Write(saldo_temporal)%></font></td>-->
		 <!-- <td align="center"><font size="-1"><%'RESPONSE.Write(SUMA_SALDO)%></font></td>-->
		   <!--<td align="center"><font size="-1"><%'RESPONSE.Write(saldo_antgto)%></font></td>-->
		   <% saldoscuenta2(isaldos)=saldo_temporal 
		      saldosvencidos(isaldos)=saldo_vencido %>
		  <% else  %>
   		  <td align="center"><font size="-1"><%RESPONSE.Write("--")%></font></td>
		  <%end if%>
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
		  " where a.cgas31='"&cg&"' and a.cgas31=b.refe11 and (conc11='SCA' or conc11='SCR') and fech11<='"&DateF&"' "
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
			    SALDOS_connota=formatnumber((anticipo_tmp+LIQUIDACIONES-(gasto_notas)),2)
		        saldos_connota=formatnumber((SALDOS_connota-total_dev),2)
		     else
			    SALDOS_connota=formatnumber((anticipo_tmp+LIQUIDACIONES-(gasto_notas)),2)
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
		 <td align="center"><font size="-1"><%RESPONSE.Write("")%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(fecpago)%></font></td>
		 <% if concepto="SCA" then %> 
		 <td align="center"><font size="-1"><%RESPONSE.Write(cg)%></font></td>
		 <% else
		      if concepto="SCR" then%>
		 <td align="center"><font size="-1"><%RESPONSE.Write(cg)%></font></td>		 
		 <%   end if
		    end if%>
		 <td align="center"><font size="-1"><%RESPONSE.Write("--")%></font></td>
		 <% if ban_gastos_sca=1 then %> 
 		 <td align="center"><font size="-1"><%RESPONSE.Write(gto_nota)%></font></td>
		 <% else %>
		 <td align="center"><font size="-1"><%RESPONSE.Write("-"&gto_nota)%></font></td>
		 <% end if%>
  		 <td align="center"><font size="-1"><%RESPONSE.Write("--")%></font></td>
		 
		  <% if num_nota=nnotas then 
		        TOTAL_SALDOS=(TOTAL_SALDOS)-(SALDOS)
				saldoscuenta(isaldos)=saldos%> 
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
		 
		<!-- <td align="center"><font size="-1"><%'RESPONSE.Write(saldo_temporal)%></font></td>
		  <td align="center"><font size="-1"><%'RESPONSE.Write(SUMA_SALDO)%></font></td>
		  <td align="center"><font size="-1"><%'RESPONSE.Write(saldo_antgto)%></font></td>-->
		  <% saldoscuenta2(isaldos)=saldo_temporal 
		     saldosvencidos(isaldos)=saldo_vencido %>
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
	  isaldos=isaldos+1
      wend 'REFErencia
'************************************************************************************************		  
     'muestra el vector
     For noce = 0 to ubound(saldoscuenta)   
         'Response.write(saldoscuenta(noce)&" Posicion:"&noce&"--"&"<br>")
		 ntotal=ntotal+saldoscuenta(noce)
     next
	 
	 For noce = 0 to ubound(saldoscuenta2)   
	      if isnull(saldoscuenta2(noce)) then
             saldoscuenta2(noce)=0
		  end if 
		  if isnull(ntotal2) then
             ntotal2=0
		  end if 
		  'Response.write(saldoscuenta2(noce)&" Posicion:"&noce&"--"&"<br>")
		 ntotal2=(ntotal2+saldoscuenta2(noce))
     next
	 
	 '**********************************************muestra el vector 2
     'For noce = 0 to ubound(saldoscuenta)   
'         'Response.write(saldoscuenta(noce)&" Posicion:"&noce&"--"&"<br>")
'		 ntotal=ntotal+saldoscuenta(noce)
'     next
	 
	 For nocetu = 0 to ubound(saldosvencidos)   
	      if isnull( saldosvencidos(nocetu)) then
              saldosvencidos(nocetu)=0
		  end if 
		  if isnull(ntotal4) then
             ntotal4=0
		  end if 
		  'Response.write(saldoscuenta2(noce)&" Posicion:"&noce&"--"&"<br>")
		 ntotal4=(ntotal4+saldosvencidos(nocetu))
     next
		  
		  
'************************************************************************************************	
	  
	  %>
	  
	  <tr>
		 <td colspan="7" align="right"><font size="-1"><%RESPONSE.Write("Saldo:")%></font></td>
		 <td><font size="-1"><%'RESPONSE.Write(ntotal)%></font></td>
 		 <td align="center"><font size="-1"><%RESPONSE.Write(ntotal2)%></font></td>
		  <td align="right"><font size="-1"><%RESPONSE.Write("Saldo Vencido:")%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(ntotal4)%></font></td>
		 
		 
	  </tr>	 
	  
<%	  
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