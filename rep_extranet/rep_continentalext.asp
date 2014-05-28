<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
 
 <style type="text/css">
.style20 {color: #FFFFFF}
 </style>

<%
    Response.Buffer = TRUE
    Response.Addheader "Content-Disposition", "attachment;filename=RepContinental.xls"
    Response.ContentType = "application/vnd.ms-excel"
    Server.ScriptTimeOut=100000

 STRFINI=request.form("txtDateIni")
 STRFFIN=request.form("txtDateFin")
 strTipo=request.Form("rbnTipoDate")

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

'empieza el query del reporte
  MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
  Set Conn = Server.CreateObject ("ADODB.Connection")
  Set REFE = Server.CreateObject ("ADODB.RecordSet")
  Conn.Open MM_EXTRANET_STRING  

If strTipo=1 then 'importacion
STRSQL= "select ptoemb01,nomcli01,refcia01,feta01,totbul01,fdsp01,numped01,fecpag01,desf0101,valdol01,ctomar01,pesobr01,tipPed01 " &_
        " from ssdagi01 , c01refer  " &_
		" where refcia01=refe01 and fecpag01>='"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"'   and firmae01<>'' and cveped01 <>'R1' " &_
		" "&permi&"  "
else
STRSQL= "select ptoemb01,nomcli01,refcia01,feta01,totbul01,fdsp01,numped01,fecpag01,desf0101,valdol01,ctomar01,valfac01, " &_
        " tipcam01,pesobr01,tipPed01 from ssdage01 , c01refer  " &_
		" where refcia01=refe01 and fecpag01>='"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"'   and firmae01<>'' and cveped01 <>'R1' " &_
		" "&permi&"  "
end if  
Set REFE= Conn.Execute(strSQL)
  'response.Write(strsql)
  'response.end()
  If strTipo=1 then 
    tipoper="Importacion"
 else	
    tipoper="Exportacion"
 end if  
  %>		

 
 <table align="left" >
     <tr><td></td>
     <td bgcolor="#F3F781" colspan="22"><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>REPORTE DE EMBARQUES 
  &nbsp;&nbsp;DEL <%=STRFINI%> AL <%=STRFFIN%></b></FONT></td> </tr>
      <tr><td></td><td bgcolor="#F3F781" colspan="22"><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>CLIENTE: CONTINENTAL AUTOMOTIVE GUADALAJARA MEXICO SA DE CV</b></FONT></td>  </tr>
  <tr><td></td><td bgcolor="#F3F781" colspan="22"><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>GUADALAJARA, JALISCO</b></FONT></td></tr>
   
																																																					
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Num</b></FONT></td>
       <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Puerto</b></FONT></td>
	   <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Cliente</b></FONT></td>
	   <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Referencia</b></FONT></td>
	   <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Pedimento</b></FONT></td>
	   <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Fecha Pago</b></FONT></td>
	   <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Fecha Despacho</b></FONT></td>
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Proveedor</b></FONT></td>
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>No.Factura</b></FONT></td>
	 <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>No. De Parte</b></FONT></td>
	 <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>PALLETS</b></FONT></td>
	 <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Peso Bruto</b></FONT></td>
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Valor Comercial</b></FONT></td>		     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Valor Factura usd</b></FONT></td>		     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Valor Aduana</b></FONT></td>		     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Incoterm</b></FONT></td>		     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Buque</b></FONT></td>		     	     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>B/L.</b></FONT></td>  
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>ETA</b></FONT></td>
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Mercancia</b></FONT></td>		
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>DTA</b></FONT></td>
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>IVA</b></FONT></td>
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>ADV</b></FONT></td>		
     <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>PRV</b></FONT></td>	
    <td bgcolor="#FEC98D" ><font color="#060300" size="2" face="Arial, Helvetica, sans-serif"><b>Tipo Ped.</b></FONT></td>	     
  	    </tr> 
	 <%
	'Dim puert(30)   'fracciones y otros
        Dim buq(30)
        Dim doc(30)
		Dim vadu(20)
        Dim frac(200)
     	Dim merca(400)
        Dim ig(200)
        Dim iv(200)
        Dim tax(200)
		Dim prov(200)
		Dim mm(200)
		Dim ico(200)
		Dim icot(200)
		Dim npar(200)
		Dim npart(400)
		Dim nompr(200)
		Dim nomp(200)
		

       index=1
	   tempo=index
       if not REFE.eof then 
       While (NOT  REFE.EOF) 
	   
	    i=0
		i2=0
		i3=0
		i4=0
		i5=0
		i6=0
		i7=0
        sumigi=0
	    sumiva=0
	    puerto=REFE("ptoemb01")
		CLIENTE=REFE("nomcli01")
        referencia=REFE("refcia01") 
	    ETA=REFE("Feta01")
		Bultos=REFE("totbul01")
		Peso=Round(REFE("pesobr01"))
		Fdespacho=REFE("fdsp01")
		NOPED=REFE("NUMPED01") 		
	    FECHAPAG=REFE("fecpag01")
		nfacturas=REFE("desf0101")
        valdolfac=REFE("valdol01")
        BL=REFE("ctomar01")
	TipoPed=REFE("tipPed01")
		if strTipo=2 then
		   valfact=cdbl(REFE("valfac01"))
           tipcambio=cdbl(REFE("tipcam01")) 
		   valcomer=Round(tipcambio*valfact)
		end if   
		
        '****************************** BL
		if BL="" then 
		Set RsRevisa = Server.CreateObject("ADODB.Recordset")
        RsRevisa.ActiveConnection = MM_EXTRANET_STRING
        strSQL=" select * from d09conoc where refe09='"&REFE("refcia01")&"' "
		RsRevisa.Source = strSQL
        RsRevisa.CursorType = 0
        RsRevisa.CursorLocation = 2
        RsRevisa.LockType = 1
        RsRevisa.Open()
		'response.Write(strSQL)
		'response.End()
        if not RsRevisa.eof then
		   BL =  RsRevisa.Fields.Item("guia09").Value
		   if BL="" then
		      BL = RsRevisa.Fields.Item("fgui09").Value
		   end if
	    end if
        RsRevisa.close
        set RsRevisa = nothing  
        end if  
		'******************************buque 

		Set RsRevisa = Server.CreateObject("ADODB.Recordset")
        RsRevisa.ActiveConnection = MM_EXTRANET_STRING
        strSQL=" select a.refe01 as Referencia ,a.cbuq01,a.ptoemb01 as PUERTO,b.nomb06 AS BUQUE,d.fdoc01 as Doc_Recibidos " &_
        " from c01refer a,c06barco b,c01refer d where a.refe01='"&REFE("refcia01")&"' and a.cbuq01=b.clav06 and a.refe01=d.refe01 "
		RsRevisa.Source = strSQL
        RsRevisa.CursorType = 0
        RsRevisa.CursorLocation = 2
        RsRevisa.LockType = 1
        RsRevisa.Open()
		'response.Write(strSQL)
		'response.End()
        if not RsRevisa.eof then
	       strPuerto =  RsRevisa.Fields.Item("puerto").Value
	       strBuque  =  RsRevisa.Fields.Item("buque").Value
	       strDoctos =  RsRevisa.Fields.Item("doc_recibidos").Value
        end if
        RsRevisa.close
        set RsRevisa = nothing  

		'******************************valor aduanal
        Set RSf = Server.CreateObject("ADODB.Recordset")
            RSf.ActiveConnection =  MM_EXTRANET_STRING
            strSQL="SELECT count(refcia02) as nf,sum(Vaduan02) as valadu FROM SSFRAC02 " &_
		    " WHERE REFCIA02='"&REFE("refcia01")&"' "
		    RSf.Source = strSQL
            RSf.CursorType = 0
            RSf.CursorLocation = 2
            RSf.LockType = 1
            RSf.Open()
		  if not RSf.eof then
		    Contador = RSf.Fields.Item("nf").Value
			valoraduana= RSf.Fields.Item("valadu").Value
		  end if
		  RSf.close
		  set RSf = nothing		
		  
		  	'******************************valor comercial
			If strTipo=1 then 'importacion
			
               Set RSvalco = Server.CreateObject("ADODB.Recordset")
               RSvalco.ActiveConnection =  MM_EXTRANET_STRING
               strSQL="select (valmer01) as vcom from ssdagi01 where refcia01='"&REFE("refcia01")&"' " 
		       RSvalco.Source = strSQL
               RSvalco.CursorType = 0
               RSvalco.CursorLocation = 2
               RSvalco.LockType = 1
               RSvalco.Open()
		        if not RSvalco.eof then
		           vcomm = RSvalco.Fields.Item("vcom").Value
		        end if
		       RSvalco.close
		       set RSvalco = nothing
			   		
		     else
			 
			  Set RSvalco = Server.CreateObject("ADODB.Recordset")
               RSvalco.ActiveConnection =  MM_EXTRANET_STRING
               strSQL="select (tipcam01*valfac01) as vcom2 from ssdage01 where refcia01='"&REFE("refcia01")&"' "
		       RSvalco.Source = strSQL
               RSvalco.CursorType = 0
               RSvalco.CursorLocation = 2
               RSvalco.LockType = 1
               RSvalco.Open()
		        if not RSvalco.eof then
		           vcomm = RSvalco.Fields.Item("vcom2").Value
		        end if
		       RSvalco.close
		       set RSvalco = nothing	
			   end if
		  '******************************d05artic
		  Set RSf = Server.CreateObject("ADODB.Recordset")
            RSf.ActiveConnection =  MM_EXTRANET_STRING
            strSQL="SELECT count(refe05) as numrf5 FROM d05artic where refe05='"&REFE("refcia01")&"' "
		    RSf.Source = strSQL
            RSf.CursorType = 0
            RSf.CursorLocation = 2
            RSf.LockType = 1
            RSf.Open()
		  if not RSf.eof then
		    Countpart = RSf.Fields.Item("numrf5").Value
		  end if
		  RSf.close
		  set RSf = nothing		
		  
		  
		Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RSart = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING
		strSQL=" SELECT * FROM d05artic where refe05='"&REFE("refcia01")&"' "
	    Set RSart= Connx.Execute(strSQL)
        Do while not RSart.Eof
		    if RSart("item05")<>"" then
            nparte=RSart("item05")
			else
			nparte="-"
			end if
            npart(i6)=nparte
            i6 = i6+1
	     RSart.MoveNext  ' de las fracciones
         Loop ' de las fracciones

'************************************************************* impuestos DTA PRV

        Set ConnxF = Server.CreateObject ("ADODB.Connection")
        Set Rimp = Server.CreateObject ("ADODB.RecordSet")
        ConnxF.Open MM_EXTRANET_STRING
	    sqL="select refcia36,cveimp36,import36 from sscont36 where refcia36='"&REFE("refcia01")&"' "
	    Set Rimp= ConnxF.Execute(sqL)
		Do while not Rimp.Eof
           cveimpu=Rimp.Fields.Item("cveimp36").Value 
             select case (cveimpu)
			 case "1" : dta=Rimp.Fields.Item("import36").Value
			 case "15": prv=Rimp.Fields.Item("import36").Value
			 end select 
			 if dta="" then
		        dta=0
		     end if
		     if prv="" then
		        prv=0
		     end if	                 
           Rimp.MoveNext  ' de las EXPO
         Loop ' de las EXPO

'*****************************************************************
'	   Set Rimp = Server.CreateObject("ADODB.Recordset")
'           Rimp.ActiveConnection =  MM_EXTRANET_STRING
'		   sqL="select refcia36,cveimp36,import36 from sscont36 where refcia36='"&REFE("refcia01")&"' "
'		   Rimp.Source = sqL
'           Rimp.CursorType = 0
'           Rimp.CursorLocation = 2
'           Rimp.LockType = 1
'           Rimp.Open()
'		  if not Rimp.eof then
'		     cveimpu=Rimp.Fields.Item("cveimp36").Value 
'             select case (cveimpu)
'			 case "1" : dta=Rimp.Fields.Item("import36").Value
'			 case "15": prv=Rimp.Fields.Item("import36").Value
'			 end select                  
'		  end if
'		  Rimp.close
'		  set Rimp = nothing
'          if dta="" then
'		     dta=0
'		  end if
'		  if prv="" then
'		     prv=0
'		  end if	   

'************************************************************** mercancias
	      cont=0
		  Set RSf = Server.CreateObject("ADODB.Recordset")
          RSf.ActiveConnection =  MM_EXTRANET_STRING
		  strSQL="select count(desc05) as nmer  from d05artic where refe05='"&REFE("refcia01")&"'  " 
		  RSf.Source = strSQL
          RSf.CursorType = 0
          RSf.CursorLocation = 2
          RSf.LockType = 1
          RSf.Open()
		  if not RSf.eof then
		    nme = RSf.Fields.Item("nmer").Value
		 end if
		  RSf.close
		  set RSf = nothing		
				
		Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RSfrac = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING
		strSQL=" SELECT refe05,fraarn02,d_mer102,ifnull(i_adv102,0) as i_adv102, ifnull(I_iva102,0) as I_iva102 ,desc05 " &_
       " FROM SSFRAC02,d05artic WHERE REFCIA02='"&REFE("refcia01")&"' and refe05=REFCIA02 and frac05=fraarn02 and agru05=ordfra02 "
	    
		'response.Write(strSQL)
		'response.End()
		
		Set RSfrac= Connx.Execute(strSQL)
						
        Do while not RSfrac.Eof
            nfraccion=RSfrac("fraarn02")
			mercancia=RSfrac("d_mer102")

            merca(i3)=mercancia
		
          i3 = i3+1
	     RSfrac.MoveNext  ' de las fracciones
         tempo=index+1
         Loop ' de las fracciones

   	
		'********************************************para sacar iva, igi y dta de sscont36
	  Set Rsscont = Server.CreateObject("ADODB.Recordset")
          Rsscont.ActiveConnection =  MM_EXTRANET_STRING
		  sqL="select refcia36, import36 as iva from sscont36 where refcia36='"&REFE("refcia01")&"' and cveimp36=3 "
		  Rsscont.Source = sqL
          Rsscont.CursorType = 0
          Rsscont.CursorLocation = 2
          Rsscont.LockType = 1
          Rsscont.Open()
		  if not Rsscont.eof then
	        iva = Rsscont.Fields.Item("iva").Value
		  else
		    iva=0
		  end if
		  Rsscont.close
		  set Rsscont = nothing
		  
	  Set Rsscont2 = Server.CreateObject("ADODB.Recordset")
          Rsscont2.ActiveConnection =  MM_EXTRANET_STRING
		  sqL="select refcia36, import36 as igi from sscont36 where refcia36='"&REFE("refcia01")&"' and cveimp36=6 "
		  Rsscont2.Source = sqL
          Rsscont2.CursorType = 0
          Rsscont2.CursorLocation = 2
          Rsscont2.LockType = 1
          Rsscont2.Open()
		  if not Rsscont2.eof then
		     igi = Rsscont2.Fields.Item("igi").Value
			else
			 igi=0
		  end if
		  Rsscont2.close
		  set Rsscont2 = nothing  
         
	 Set Rsscont2 = Server.CreateObject("ADODB.Recordset")
          Rsscont2.ActiveConnection =  MM_EXTRANET_STRING
		  sqL="select refcia36, import36 as dta from sscont36 where refcia36='"&REFE("refcia01")&"' and cveimp36=1 "
		  Rsscont2.Source = sqL
          Rsscont2.CursorType = 0
          Rsscont2.CursorLocation = 2
          Rsscont2.LockType = 1
          Rsscont2.Open()
		  if not Rsscont2.eof then
		     dta = Rsscont2.Fields.Item("dta").Value
			else
			 dta=0
		  end if
		  Rsscont2.close
		  set Rsscont2 = nothing  
          
   	 '***********************************facturas

	   Set Rfac = Server.CreateObject("ADODB.Recordset")
          Rfac.ActiveConnection =  MM_EXTRANET_STRING
		  sqL="select refcia39,numfac39,idfisc39,nompro39,dompro39,terfac39 from ssfact39 " &_
          "where refcia39='"&REFE("refcia01")&"' "
			  Rfac.Source = sqL
          Rfac.CursorType = 0
          Rfac.CursorLocation = 2
          Rfac.LockType = 1
          Rfac.Open()
		  if not Rfac.eof then
		    taxi = Rfac.Fields.Item("idfisc39").Value
		    'nompro = Rfac.Fields.Item("nompro39").Value
		    dom = Rfac.Fields.Item("dompro39").Value
			'icoterm = Rfac.Fields.Item("terfac39").Value
			nfactura= Rfac.Fields.Item("numfac39").Value
			cveprov= Rfac.Fields.Item("nompro39").Value
		 end if
		  Rfac.close
		  set Rfac = nothing
   
   		  Set RSf = Server.CreateObject("ADODB.Recordset")
          RSf.ActiveConnection =  MM_EXTRANET_STRING
		  strSQL="select count(terfac39) as nico from ssfact39 where refcia39='"&REFE("refcia01")&"'  " 
		  RSf.Source = strSQL
          RSf.CursorType = 0
          RSf.CursorLocation = 2
          RSf.LockType = 1
          RSf.Open()
		  if not RSf.eof then
		    nmico = RSf.Fields.Item("nico").Value
		 end if
		  RSf.close
		  set RSf = nothing	
		  
		  	  Set RSf = Server.CreateObject("ADODB.Recordset")
          RSf.ActiveConnection =  MM_EXTRANET_STRING
		  strSQL="select count(nompro39) as nipro from ssfact39 where refcia39='"&REFE("refcia01")&"'  " 
		  RSf.Source = strSQL
          RSf.CursorType = 0
          RSf.CursorLocation = 2
          RSf.LockType = 1
          RSf.Open()
		  if not RSf.eof then
		    nopro = RSf.Fields.Item("nipro").Value
		 end if
		  RSf.close
		  set RSf = nothing
   
   
   	    Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RSfactu = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING
		strSQL=" select refcia39,numfac39,idfisc39,nompro39,dompro39,terfac39 from ssfact39 " &_
          "where refcia39='"&REFE("refcia01")&"' "
	    Set RSfactu= Connx.Execute(strSQL)
        Do while not RSfactu.Eof
		    if RSfactu("terfac39")<>"" then
            icoterm=RSfactu("terfac39")
			else
			icoterm="-"
			end if
            icot(i4)=icoterm
          i4 = i4+1
	     RSfactu.MoveNext  ' del icoterm
         Loop ' del icoterm		 

	    '------------------para sacar todos lo provedores en un renglon
		Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RSProv = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING
		strSQL=" select refcia39,numfac39,nompro39 from ssfact39 " &_
          "where refcia39='"&REFE("refcia01")&"' "
	    Set RSProv= Connx.Execute(strSQL)
        Do while not RSProv.Eof
		    if RSProv("nompro39")<>"" then
            nompro=RSProv("nompro39")
			else
			nompro="-"
			end if
            nompr(i7)=nompro
            i7 = i7+1
	     RSProv.MoveNext  ' de los proveedores
         Loop ' de los proveedores
   
'*****************************DATOS FILAS*****************************
	     x=0 
		 xj=0
		 xj2=0%>
		 
		 <tr>
         <td><font size="-1"><%RESPONSE.Write(index)%></font></td>
		 <td><font size="-1"><%RESPONSE.Write(puerto)%></font></td>
	 	 <td><font size="-1"><%RESPONSE.Write(cliente)%></font></td>
         <td><font size="-1"><%RESPONSE.Write(referencia)%></font></td>
		 <TD align="center" ><font size="-1"><%RESPONSE.Write(noped)%>&nbsp;</font></TD>
	     <td align="center" ><font size="-1"><%RESPONSE.Write(fechapag)%></font></td>	
		 <TD align="center" ><font size="-1"><%RESPONSE.Write(fdespacho)%></font></TD>
		 <%
		 for xj2=0 to nopro-1
		 if xj2=0 then
		  nomp(0)=nompr(0)
		 else
   		  nomp(0)= nomp(0)&","&nompr(xj2)
     	 end if
		 next
		 %>
		  <td><font size="-1"><%RESPONSE.Write(nomp(0))%></font></td>
		  <%if strTipo=1 then%> 
		 <td><font size="-1"><%RESPONSE.Write(nfacturas)%></font></td>
		 <%else%>
		 <td><font size="-1"><%RESPONSE.Write(nfactura)%></font></td>
		 <%end if%>
		 <%
		   for xj2=0 to Countpart-1
		      if xj2=0 then
		         npar(0)=npart(0)
		      else
   		         npar(0)= npar(0)&","&npart(xj2)
     	      end if
		   next
		 %>
		 <td><font size="-1"><%response.Write(npar(0))%></td>
		 <td><font size="-1"><%RESPONSE.Write(bultos)%></font></td>
	     <td><font size="-1"><%RESPONSE.Write(peso)%></font></td>
		 <td><font size="-1"><%response.write(vcomm)%></td>
 		 <td><font size="-1"><%RESPONSE.Write(valdolfac)%></font></td>
	 	 <td><font size="-1"><%RESPONSE.Write(valoraduana)%></font></td>
   	     <%  
		 for xj=0 to nme-1
		 if xj=0 then
		  mm(0)=merca(0)
		 else
		  mm(0) = mm(0)&","&merca(xj)
     	 end if
		 next
		 for xj2=0 to nmico-1
		 if xj2=0 then
		  ico(0)=icot(0)
		 else
   		  ico(0)= ico(0)&","&icot(xj2)
     	 end if
		 next
		 %>
		 <td><font size="-1"><%RESPONSE.Write(ico(0))%></font></td>
		 <td><font size="-1"><%RESPONSE.Write(strBuque)%></font></td>
 		 <td><font size="-1"><%RESPONSE.Write(BL)%></font></td>
		 <td><font size="-1"><%RESPONSE.Write(eta)%></font></td>
		 <td><font size="-1"><%RESPONSE.Write(mm(0))%></font></td>
		 <td><font size="-1"><%response.Write(dta)%></td>
 		 <td align="center" ><font size="-1"><%RESPONSE.Write(iva)%></font></td>
		 <td align="center" ><font size="-1"><%RESPONSE.Write(igi)%></font></td>
	     <td><font size="-1"><%response.Write(prv)%></td>
	         <td><font size="-1"><%response.Write(TipoPed)%></td>

<%'**********************************************************************
   	RESPONSE.Write("</tr>")
         Refe.MoveNext 'avanza referencia  ---->
		 index=index+1
   wend 'REFErencia
 else
%>
<tr>
  <th colspan=12>
    <font size="1" face="Arial"> No se Encontro ningun registro con esos parametros
  </th>
</tr>
<table>

<%end if%>

