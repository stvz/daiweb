<!-- #include virtual="/PortalMySQL/Extranet/ext-Asp/DAO/CuentaGastosSEMDAO.asp" -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
	Server.ScriptTimeout=15000000
	strTipoUsuario = request.Form("TipoUser")
	strPermisos = Request.Form("Permisos")
	permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
	fileName = "C0000009.FAC"
	
	'ContentType especifica el tipo de MIME de este encabezado.	
	Response.ContentType = "text/plain"
	'El método AddHeader agrega un encabezado HTML con un valor específico.
    'Content-disposition obliga al explorador a descargar.
	Response.AddHeader "content-disposition", "attachment; filename=""" & fileName & """"
	
	'call stpcg('dai','D000549',2,'2011-01-01', '2011-03-31', 'SEM950215S98');
	ofi="dai"
	cg=""
	strFechaI="2011-03-01"
	strFechaF="2011-03-31"
	sRFC="SEM950215S98"
	cont=0
	nPeds=1
	'http://10.66.1.9/portalmysql/extranet/ext-asp/reportes/CGSamsung.asp
  	  		
	Set oRST1= Nothing
	set oRST2= Nothing
	set oRST3= Nothing
	set oRST4= Nothing
	'---------------------------------------------------------------------------------->>Obtenemos las Cuentas de Gastos
	set oCG = New cCGSEM					
	oCG.getCGByFech ofi,sRFC,strFechaI,strFechaF,oRST1
	'---------------------------------------------------------------------------------->>
	'set confile = createObject("scripting.filesystemobject") 
			
	While NOT oRST1.EOF			
		nPeds= verTotPedms(oRST1.fields("CG").value)
		s910="910|"&oRST1.fields("Patente").value&"|"&oRST1.fields("Aduana").value&"|"&oRST1.fields("Pedimento").value&"|"&oRST1.fields("CGastos").value&"|"&oRST1.fields("FechCG").value&"|"&oRST1.fields("RFC").value&"|"&oRST1.fields("IMPORTE").value&"|"&oRST1.fields("IVA").value&"||"&oRST1.fields("Subtotal").value&"|"&oRST1.fields("ANTICIPO").value&"|"&oRST1.fields("TOTAL_CG").value&"|"&oRST1.fields("Honorarios").value&"|"&oRST1.fields("Financiamiento").value&"|"&oRST1.fields("ServComp").value&"|"&oRST1.fields("TotalPH").value&"||"&oRST1.fields("AdicHon").value&"|"&oRST1.fields("Proform").value&"|"&oRST1.fields("Cancela").value&"|" 
		response.write(s910&chr(13)&chr(10))
		if(CInt(nPeds)>1) Then
			genera911 oRST1.fields("CG").value,sRFC,oRST1.fields("Ref").value
		End If		
		genera912 oRST1.fields("CG").value,sRFC		
		oRST1.movenext()		
	Wend

	function  verTotPedms(cg)
		oCG.getTotCG cg,oRST2
		if not(oRST2.eof) then
			verTotPedms =oRST2.fields("totCG").value
		else
			verTotPedms = 0
		end if		
	end function
	
	sub genera911(cg,rfc,refer)
		oCG.getCGPed cg,rfc,oRST4
		if not(oRST4.eof) then
			while not oRST4.EOF
				if(refer<>oRST4.fields("Referencia").value)then
					s911="911|"&cg&"|"&oRST4.fields("Patente").value&"|"&oRST4.fields("Aduana").value&"|"&oRST4.fields("Pedim").value&"|"
					response.write(s911&chr(13)&chr(10))
				End If
				oRST4.movenext()
			wend			
		end if	
	End Sub
	
	sub genera912(cg,rfc)
		contad=1
		oCG.getPHbyCG cg,rfc,oRST3
		if not(oRST3.eof) then
			while not oRST3.EOF
				s912="912|"&cg&"|"&contad&"|"&oRST3.fields("RFC").value&"|"&oRST3.fields("Factura").value&"|"&oRST3.fields("Prov").value&"|"&oRST3.fields("Concepto").value&"|"&oRST3.fields("PH").value&"|"&oRST3.fields("IVA_PH").value&"|"
				response.write(s912&chr(13)&chr(10))				
				oRST3.movenext()
				contad=contad+1
			wend			
		end if		
	end sub 
	
	
	

%>