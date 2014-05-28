<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

 <style type="text/css">
.style20 {color: #FFFFFF}
 </style>

<%
    Response.Buffer = TRUE
    'Response.Addheader "Content-Disposition", "attachment;filename=RepContinental.xls"
    'Response.ContentType = "application/vnd.ms-excel"
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

'***********aqui empiezan los querys del reporte
 MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
 Set Conn = Server.CreateObject ("ADODB.Connection")
 Set REFE = Server.CreateObject ("ADODB.RecordSet")
 Conn.Open MM_EXTRANET_STRING

If strTipo=1 then 'importacion --chekar si fecent01 aparece al = modo01='T' xeso no lo pues en expo
STRSQL= " select refcia01,year(fecpag01) as Anio,patent01,numped01,feta01,fecpag01,cveped01,'1' as Tipo,embala01, " &_
        " tipcam01,cveadu01,valmer01,factmo01,fletes01,segros01,otros01,cvepod01,fecent01,regime01,nombar01,tipPed01,fdsp01 " &_
        " from ssdagi01 , c01refer  " &_
		" where refcia01=refe01 and fecpag01>='"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"' and cveped01 <>'R1' " &_
		" and firmae01<>'' and modo01='T' "&permi&"  "
else
STRSQL= " select refcia01,year(fecpag01) as Anio,patent01,numped01,feta01,fecpag01,cveped01,'2' as Tipo,embala01, " &_
        " tipcam01,valfac01,cveadu01,factmo01,fletes01,segros01,otros01,cvepod01, " &_
        " pesobr01,regime01,nombar01,tipPed01,fdsp01 from ssdage01 , c01refer  " &_
	" where refcia01=refe01 and fecpag01>='"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"'  and cveped01 <>'R1' and firmae01<>'' " &_
		" "&permi&" "
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
<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p>Reporte de GETRAC de <%=tipoper%> </p></font></  strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p></p></font></strong>
 <strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif"><p> Del <%=STRFINI%> al <%=STRFFIN%> </p></font></  strong>
   <table align="left" border="1" >
   <tr>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>No de Pedimento</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Fecha de Cruce</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>No de Patente</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Aduana de Entrada</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Clave de pedimento</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Tipo de Cambio</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Valor en Aduanas</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Valor Comercial</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>No de Factura</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Proveedor</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>IGI (FP)</b></FONT></td>
 	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>IVA</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>DTA (FP)</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>PREV (FP)</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>CC (FP)</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Fracc. Arancelaria</b></FONT></td>
       <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Origenes</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>No de Parte</b></FONT></td>
	   <td bgcolor="#006699" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Descripción de la Mercancía</b></FONT></td>
  </tr>
	 <%
       Dim npart(200)
   

       index=1
	   tempo=index
       if not REFE.eof then
       While (NOT  REFE.EOF)
	    i6=0
        identificador=" "
	    referencia=REFE("refcia01")
	    ano=REFE("anio")
		patente=REFE("patent01")
        pedimento=REFE("numped01")
		ETA=REFE("Feta01")
		fecpago=REFE("fecpag01")
		clave=REFE("cveped01")
		tipo=REFE("tipo")
		tcam=REFE("tipcam01")
		adu=REFE("cveadu01")
		transportista=REFE("nombar01")
		fdespacho=REFE("fdsp01")
		if strTipo=1 then
		  if REFE("fecent01")<>"" then
		   fecent=REFE("fecent01")
		  else
		   fecent=""
		  end if
		end if
		fletes=REFE("fletes01")
		seguros=REFE("segros01")
	    embalaje=REFE("embala01")
		otros=REFE("otros01")
		tipoped=REFE("tipPed01")
		if strTipo=2 then
		 regimen=REFE("regime01")
		 else
		 regimen=REFE("regime01")
		end if
		Paiss=REFE("cvepod01")
		'if strTipo=2 then
		   'valfact=cdbl(REFE("valfac01"))
           'tipcambio=cdbl(REFE("tipcam01"))
		   'valcomer=Round(tipcambio*valfact)
		'end if

		'******************************IDENTIFICADORES
		ix=1
		Set Conn = Server.CreateObject ("ADODB.Connection")
        Set RsIDE = Server.CreateObject ("ADODB.RecordSet")
        Conn.Open MM_EXTRANET_STRING
     	strSQL=" SELECT * FROM ssiped11 where refcia11='"&REFE("refcia01")&"' "
	    Set RsIDE= Conn.Execute(strSQL)
        Do while not RsIDE.Eof
		    iden=RsIDE("deside11")
			comple=RsIDE("comide11")
		  if ix=1 then
			   identificador=iden
			   complemento=comple
			else
		       identificador=identificador+","+iden
			   complemento=complemento+","+comple
		  end if
		  ix=ix+1
		 RsIDE.MoveNext  ' de ssiped11
       Loop ' de ssiped11

         '****************************************GUIAS*************************************

          Set Conx = Server.CreateObject ("ADODB.Connection")
          Set RSGUIA = Server.CreateObject ("ADODB.RecordSet")
          Conx.Open MM_EXTRANET_STRING
	      SQLx="SELECT * FROM Ssguia04 WHERE REFCIA04='"&REFE("refcia01")&"' and numgui04<>'' "
	      Set RSGUIA= Conx.Execute(SQLx)
          'response.Write(sqlx)
          'response.end
          Do while not RSGUIA.Eof
            idngui04 = RSGUIA("idngui04")
			numgui04= RSGUIA("numgui04")
            if  idngui04=1 then
	            gmaster= numgui04
		     else
		      if  idngui04=2 then
                  gmaster2= numgui04
		      end if
		     end if
 	      RSGUIA.MoveNext
          Loop

		'******************************d05artic

		Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RsD05art = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING
     	strSQL=" select * from d05artic where refe05='"&REFE("refcia01")&"' group by agru05,fact05 "
	    Set RsD05art= Connx.Execute(strSQL)
        'Do while not RsD05art.Eof
		if not RsD05art.eof then
        while not RsD05art.Eof
		   strfrac=RsD05art("frac05")
		   strfact=RsD05art("fact05")
	       strAgru=RsD05art("agru05")
		   strmate=RsD05art("desc05")
		   strparte=RsD05art("item05")

      '************************************************************** fracciones

	     Set RSfracc = Server.CreateObject("ADODB.Recordset")
            RSfracc.ActiveConnection =  MM_EXTRANET_STRING
            strSQL="SELECT fraarn02,d_mer102,ifnull(tasadv02,0) as tasadv02, ifnull(I_iva102,0) as I_iva102, " &_
			" u_medc02,cantar02,u_medt02,cancom02,paiori02,paiscv02,i_cc0102  " &_
		    " FROM SSFRAC02 WHERE REFCIA02='"&REFE("refcia01")&"' and ordfra02='"&RsD05art("agru05")&"' and  " &_
		    " fraarn02='"&RsD05art("frac05")&"'"
		    RSfracc.Source = strSQL
            RSfracc.CursorType = 0
            RSfracc.CursorLocation = 2
            RSfracc.LockType = 1
            RSfracc.Open()
		  if not RSfracc.eof then
		    nfraccion = RSfracc.Fields.Item("fraarn02").Value
			mercancia= RSfracc.Fields.Item("d_mer102").Value
			igi = RSfracc.Fields.Item("tasadv02").Value
			ivaimpuesto= RSfracc.Fields.Item("I_iva102").Value
			umc=RSfracc.Fields.Item("u_medc02").Value
			cantarifa=RSfracc.Fields.Item("cantar02").Value
			umt=RSfracc.Fields.Item("u_medt02").Value
			cancom=RSfracc.Fields.Item("cancom02").Value
			paisOD=RSfracc.Fields.Item("paiori02").Value
			paisCV=RSfracc.Fields.Item("paiscv02").Value
			cc=RSfracc.Fields.Item("i_cc0102").Value  'ccompensatoria 
		  end if
		  RSfracc.close
		  set RSfracc = nothing

   	 '***********************************facturas

	   Set Rfac = Server.CreateObject("ADODB.Recordset")
          Rfac.ActiveConnection =  MM_EXTRANET_STRING
		  sqL="select * " &_
              "from ssfact39 where refcia39='"&REFE("refcia01")&"' and numfac39='"&RsD05art("fact05")&"' "
			  Rfac.Source = sqL
          Rfac.CursorType = 0
          Rfac.CursorLocation = 2
          Rfac.LockType = 1
          Rfac.Open()
		  if not Rfac.eof then
		    nompro = Rfac.Fields.Item("nompro39").Value
			nfactura= Rfac.Fields.Item("numfac39").Value
			nomprov= Rfac.Fields.Item("nompro39").Value
			monfact= Rfac.Fields.Item("monfac39").Value
			valor_usd=Rfac.Fields.Item("valdls39").value
			valor_me=Rfac.Fields.Item("valmex39").value
			if Rfac.Fields.Item("refcia39").value="CEG09-0549" and nfactura="06500052044616" then
			fechafac="2009-02-19"
			else
			fechafac=Rfac.Fields.Item("fecfac39").value
			end if
			cvepro=Rfac.Fields.Item("cvepro39").value
			vinculacion=Rfac.Fields.Item("vincul39").value
			factomon=Rfac.Fields.Item("facmon39").value
			if vinculacion=1 then
			   vinc="si"
			else
			   vinc="no"
			end if
		 end if
		  Rfac.close
		  set Rfac = nothing

		  '************************************************************** ssprov22

	     Set RSprov = Server.CreateObject("ADODB.Recordset")
            RSprov.ActiveConnection =  MM_EXTRANET_STRING
            strSQL="select paipro22 from ssprov22 where cvepro22='"&cvepro&"' "
		    RSprov.Source = strSQL
            RSprov.CursorType = 0
            RSprov.CursorLocation = 2
            RSprov.LockType = 1
            RSprov.Open()
		  if not RSprov.eof then
		    paiprov = RSprov.Fields.Item("paipro22").Value
		  end if
		  RSprov.close
		  set RSprov = nothing
		  
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
		  
		  '************************************************************* impuestos DTA PRV
	   Set Rimp = Server.CreateObject("ADODB.Recordset")
           Rimp.ActiveConnection =  MM_EXTRANET_STRING
		   sqL="select refcia36,cveimp36,import36 from sscont36 where refcia36='"&REFE("refcia01")&"' "
		   Rimp.Source = sqL
           Rimp.CursorType = 0
           Rimp.CursorLocation = 2
           Rimp.LockType = 1
           Rimp.Open()
		  if not Rimp.eof then
		     cveimpu=Rimp.Fields.Item("cveimp36").Value 
             select case (cveimpu)
			 case "1" : dta=Rimp.Fields.Item("import36").Value
			 case "15": prv=Rimp.Fields.Item("import36").Value
			 end select                  
		  end if
		  Rimp.close
		  set Rimp = nothing
          if dta="" then
		     dta=0
		  end if
		  if prv="" then
		     prv=0
		  end if	 
		  
		'********************************************************No de Parte**********************
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
		  
		  
'	    Set Connx = Server.CreateObject ("ADODB.Connection")
'        Set RSart = Server.CreateObject ("ADODB.RecordSet")
'        Connx.Open MM_EXTRANET_STRING
'		strSQL=" SELECT * FROM d05artic where refe05='"&REFE("refcia01")&"' "
'	    Set RSart= Connx.Execute(strSQL)
'        Do while not RSart.Eof
'		    if RSart("item05")<>"" then
'            nparte=RSart("item05")
'			else
'			nparte="-"
'			end if
'            npart(i6)=nparte
'            i6 = i6+1
'	     RSart.MoveNext  ' de las fracciones
'         Loop ' de las fracciones  
		 
		'******************************valor comercial 'importacion
			
               Set RSvalco = Server.CreateObject("ADODB.Recordset")
               RSvalco.ActiveConnection =  MM_EXTRANET_STRING
               strSQL="select (valmer01 * factmo01 *  tipcam01) as vcom from ssdagi01 where refcia01='"&REFE("refcia01")&"' " 
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


'*****************************DATOS FILAS*****************************
 
		 x=0
		 %>
	     <tr>
         <td align="center"><font size="-1"><%RESPONSE.Write(pedimento)%>&nbsp;</font></td>
         <td align="center"><font size="-1"><%RESPONSE.Write(fdespacho)%></font></td>
	 	 <td align="center"><font size="-1"><%RESPONSE.Write(patente)%></font></td>
		 <TD align="center"><font size="-1"><%RESPONSE.Write(adu)%></font></TD>
		 <td align="center"><font size="-1"><%RESPONSE.Write(clave)%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(tcam)%></font></td>
 		 <td align="center"><font size="-1"><%RESPONSE.Write(valoraduana)%></font></td>
  		 <td align="center"><font size="-1"><%RESPONSE.Write(vcomm)%></font></td>
  		 <td align="center"><font size="-1"><%RESPONSE.Write(nfactura)%></font></td>
		 <td><font size="-1"><%RESPONSE.Write(nomprov)%></font></td>
         <td><font size="-1"><%RESPONSE.Write(igi)%></font></td>		 
		 <td><font size="-1"><%RESPONSE.Write(iva)%></font></td>
		 <td><font size="-1"><%RESPONSE.Write(dta)%></font></td>
		 <td><font size="-1"><%RESPONSE.Write(prv)%></font></td>
 		 <td><font size="-1"><%RESPONSE.Write(cc)%></font></td>
  		 <td><font size="-1"><%RESPONSE.Write(nfraccion)%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(paisOD)%></font></td>
		 <td align="center"><font size="-1"><%RESPONSE.Write(strparte)%></font></td>
		 <td><font size="-1"><%RESPONSE.Write(mercancia)%></font></td>	 
		 
		 
	 <% '********d05artic
	   RsD05art.MoveNext  ' de d05artic
	   tempo=index+1
       wend ' de d05artic
	   else
	       banderita=1
         if baderita<>1 then
             'response.Write("Se encontro la referencia:"&myrefe)
              %>
              <!--
                <td colspan="15">
                  <font size="-1">  </font>
                </td>
              -->

        <%else
          banderita=banderita+1
         ' response.Write(banderita)
         end if 'del if baderita=1 then

	  end if 'del if d05artic.eof
     %>
<%'**********************************************************************
   	RESPONSE.Write("</tr>")
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
</table>
<%end if
'end if
'end if'del movimiento %>

</table>
</form>