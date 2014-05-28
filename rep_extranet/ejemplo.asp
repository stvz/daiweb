<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<HTML>
<HEAD>
<TITLE>Ejemplo objeto Recordset</TITLE>
</HEAD>
<BODY>
<%Set objConexion=Server.CreateObject("ADODB.Connection")
objConexion.Open "DSN=FuenteBD;UID=pepe;PWD=xxx"
Set objRecordset=objConexion.Execute("SELECT * FROM Usuarios")%>
<table border="1" align="center">
<tr>
<th>DNI</th>
<th>Nombre</th>
<th>Domicilio</th>
<th>Código Postal</th>
</tr>
<%while not objRecordset.EOF%>
<tr>
<td><%=objRecordset("DNI")%></td>
<td><%=objRecordset("Nombre")%></td>
<td><%=objRecordset("Domicilio")%></td>
<td align="right"><%=objRecordset("Codigo_Postal")%></td>
</tr>
<%objRecordset.MoveNext
Wend
objRecordset.Close
Set objRecordset=Nothing
objConexion.Close
Set objConexion=Nothing%>
</table>
</body>
</html>