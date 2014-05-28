<% 
Server.ScriptTimeout=1500
%>
<HTML>
<HEAD>
<script language="JavaScript">
	function bloquear(){
		cliente = document.getElementById("txtCliente").value;
		//alert(cliente);
		PorRFC = document.RepSegOps_.ckcve[0].checked;
		if (PorRFC == false && cliente != "Todos"){
			//alert("cliente = " + cliente + "filtro = " + PorRFC);
			//alert("radios " + document.RepSegOps_.ckcve[0].checked);
			//document.RepSegOps_.multioficina.checked = false;
			//document.RepSegOps_.multioficina.enabled = false;
			document.getElementById("multioficina").disabled = true;
			document.getElementById("multioficina").checked = false;
		}
		else
		{
			//alert("PORRFC = FALSE");
			document.getElementById("multioficina").disabled = false;
			//document.RepSegOps_.multioficina.enabled = True;
		}
		
	}
</script>
<TITLE>:: CATALOGO DE PRODUCTOS CORPORATIVO ::</TITLE>
</HEAD>
<BODY>
<%
if PermisoMenu(strMenu,",03-") = "PERMITIDO" or strTipoUsuario = MM_Cod_Admon then
    permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
    %>
    <table width="50%" border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
      <tr>
        <td>:: CATALOGO DE PRODUCTOS ::</td>
      </tr>
    </table>
    <p align="center" class="titulosconsultas">Elige tu opci&oacute;n:</p>
    <%
    '''''''''''''
    Set oConn = Server.CreateObject ("ADODB.Connection")
    Set RS = Server.CreateObject ("ADODB.RecordSet")
     oConn.Open MM_EXTRANET_STRING
    %>
     <FORM name="RepSegOps_" id="RepSegOps_" METHOD="POST" ACTION="/PortalMySQL/Extranet/ext-Asp/reportes/ext_opc_repcatprod.asp" target="_Blank">
          <input name="Permisos"   type="hidden" value="<%=strPermisos%>">
          <input name="Contenido"  type="hidden" value="<%=strContenido%>">
          <input name="User"       type="hidden" value="<%=strUsuario%>">
          <input name="TipoUser"   type="hidden" value="<%=strTipoUsuario%>">
          <input name="Aduana"     type="hidden" value="<%=strAduana%>">
    <%
    if strTipoUsuario = MM_Cod_Admon or strTipoUsuario = MM_Cod_Ejecutivo_Grupo or strTipoUsuario=MM_Cod_Cliente_Division then
strSQL = FiltroUsuario(strTipoUsuario)
strSQLRFC = FiltroUsuarioRFC(strTipoUsuario)
strSQL = replace(strSQL, "division18", "nomcli18")
strSQLRFC = replace(strSQLRFC, "division18", "nomcli18")
' Response.Write("strSQL<br>" & strSQL & "<br>")
' Response.Write("strSQLRFC<br>" & strSQLRFC & "<br>")


%>
<br>
<table width="98%" border="1" align="center" cellpadding="0" cellspacing="7" bordercolor="#000066">
    <TR bordercolor="#FFFFFF">
      <td width="21%"  class="OpcPedimento">
		<INPUT TYPE="RADIO" NAME="ckcve" ID="ckcve" value="RFC" onClick="bloquear()" checked>RFC  DEL CLIENTE <%=TipoCliente%>:
	  </td>
    <td width="64%" class="TextNormalAzul">
      <%
        Set RsClieRFC = Server.CreateObject("ADODB.Recordset")
        RsClieRFC.ActiveConnection = MM_EXTRANET_STRING
        RsClieRFC.Source = strSQLRFC
        RsClieRFC.CursorType = 0
        RsClieRFC.CursorLocation = 2
        RsClieRFC.LockType = 1
        RsClieRFC.Open()
      %>
        <select name="rfcCliente" class="inputs">
		  <option value="Todos">Todos</option>
          <% 
		  Do Until RsClieRFC.EOF = True
		  %>
          <option value="<%=RsClieRFC("rfccli18")%>"><%=mid((RsClieRFC("nombre")),1,40) &"   RFC:  " & RsClieRFC("rfccli18")%></option>
          <%
		  RsClieRFC.MoveNext()
          Loop
          RsClieRFC.close
          set RsClieRFC = nothing
          %>
        </select>
	</td>

    </TR>

    <TR bordercolor="#FFFFFF">
      <td width="21%"  class="OpcPedimento">
		<INPUT TYPE="RADIO" NAME="ckcve" ID="ckcve" onClick="bloquear()" value="CLAVE">CLAVE DEL CLIENTE <%=TipoCliente%>:
	  </td>
    <td width="64%" class="TextNormalAzul">
	  <%
  		Set RsClie = Server.CreateObject("ADODB.Recordset")
        RsClie.ActiveConnection = MM_EXTRANET_STRING
        RsClie.Source = strSQL
        RsClie.CursorType = 0
        RsClie.CursorLocation = 2
        RsClie.LockType = 1
        RsClie.Open()
	  %>
        <select name="txtCliente" ID="txtCliente" class="inputs" onChange="bloquear()">
			<option value="Todos">Todos</option>
			<%
			Do Until RsClie.EOF = True
			%>
			   <option value="<%=RsClie("cvecli18")%>"><%=mid((RsClie("nombre")),1,40) &"   Clave:  " & RsClie("cvecli18")%></option>
			  <% RsClie.MoveNext()
			Loop
			RsClie.close
			set RsClie = nothing %>
        </select>
	</td>
    </TR>
  </table>
<%end if%>

<%
    strFechaIniTemp= request.Form("textFechaIni")
    strFechaFinTemp = request.Form("textFechaFin")

%>
<br>

<table width="98%" border="1" align="center" cellpadding="0" cellspacing="7" bordercolor="#000066">
	<tr bordercolor="#FFFFFF">
		<td width="15%" height="30" class="OpcPedimento">BUSCAR POR</td>
		<td width="14%" class="TextNormalAzul">
			<input name="BuscaPor" type="radio" value="codprod" checked>Código de Producto</input>
	    </td>
		<td width="14%" class="TextNormalAzul">
			<input name="BuscaPor" type="radio" value="descprod">Descripción del Producto</input>
	    </td>
		<td width="14%" class="TextNormalAzul">
			<input name="BuscaPor" type="radio" value="fraccion">Fracción</input>
        </td>
	</tr>	
	<tr bordercolor="#FFFFFF">
	    <td width="14%" class="TextNormalAzul">
			<input name="buscaexacto" id="buscaexacto" type="checkbox" value="T" >Palabras Exactas</input>
	    </td>	
        <td><span class="TextNormalAzul">Valor a buscar:</span>
		     <input name="TextoCaptura" type="text" size="90" maxlength="90" class="Formularios2">
        </td>	
    </tr> 		
  </table>
  
  <table width="98%" border="1" align="center" cellpadding="0" cellspacing="7" bordercolor="#000066">
    <tr bordercolor="#FFFFFF">
		<td width="18%" height="30" class="OpcPedimento">Aduana:</td>
		<td class="TextNormalAzul"colspan="4" >&nbsp Tipo: 
			<input name="selectOficina" type="radio" value="a" checked>Todas 
			&nbsp &nbsp <input name="selectOficina" type="radio" value="r">Veracruz 
			&nbsp &nbsp <input name="selectOficina" type="radio" value="d"> M&eacute;xico
			&nbsp &nbsp <input name="selectOficina" type="radio" value="s"> Manzanillo
			&nbsp &nbsp <input name="selectOficina" type="radio" value="l"> Lazaro Cardenas
			&nbsp &nbsp <input name="selectOficina" type="radio" value="t"> Toluca
			&nbsp &nbsp <input name="selectOficina" type="radio" value="c"> Altamira
		</td>
    </tr>
  </table>
	
   <br>
  <table width="98%" border="1" align="center" cellpadding="0" cellspacing="7" bordercolor="#000066">
	<tr bordercolor="#FFFFFF">
		<td width="15%" height="30" class="OpcPedimento">Tipo Reporte:</td>
		<td width="14%" class="TextNormalAzul">
			<input name="TipRep" type="radio" value="html" checked>HTML</input>
	  </td>
		<td width="14%" class="TextNormalAzul">
			<input name="TipRep" type="radio" value="excel">EXCEL</input>
	  </td>

	</tr>
  </table>
  <table width="98%" border="1" align="center" cellpadding="0" cellspacing="7" bordercolor="#000066">
    <tr bordercolor="#FFFFFF">
      <td width="100%">
	  	<center>
         <input name="btnGenerar" type="button" OnClick="javascript:Envia('OK')" class="botonesVerde" id="btnGenerar" value="Consultar">
		</center>
      </td>
    </tr>
     </FORM>
     </div>
<%else
	strMenjError = "No tiene Autorización para visualizar este reporte"
%>
<table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
	  <tr>
		<td><%=(strMenjError)%></td>
	  </tr>
	</table>
<%end if %>

</BODY>
</HTML>
