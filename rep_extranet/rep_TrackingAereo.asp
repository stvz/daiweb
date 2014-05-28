<META HTTP-EQUIV="Content-Type" CONTENT="text/html"; charset="utf-8">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html"; charset="utf-8">
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<%On Error Resume Next
'																										 		'
'																												'
' ---------------------------------------        TRACKING AEREO       ------------------------------------------------	'
'	AL CREAR UNA NUEVA EXTRANET, HABRÍA QUE HACERLO POR OFICINA, PARA IMPORTACION SOLAMENTE																											'
' 																												'

Response.Buffer = TRUE
response.Charset = "utf-8"
Response.Addheader "Content-Disposition", "attachment;"'
Response.ContentType = "application/vnd.ms-excel"

dim strTipoUsuario,fechaini,fechafin,oficina

strTipoUsuario = Session("GTipoUsuario")
fechaini = trim(request.Form("txtDateIni"))
fechafin = trim(request.Form("txtDateFin"))

oficina = "RKU"

if not fechaini="" and not fechafin="" then

    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    finicio = tmpAnioIni & "-" &tmpMesIni & "-"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    ffinal = tmpAnioFin & "-" &tmpMesFin & "-"& tmpDiaFin

	strFiltroCliente = ""
	strFiltroCliente = request.Form("txtCliente")
	if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
	   blnAplicaFiltro = true
	end if
	if blnAplicaFiltro then
	   permi = " AND i.cvecli01 =" & strFiltroCliente
	end if
	if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
	   permi = ""
	end if

	dim bgcolor,strHTML
	bgcolor="#FFFFFF"
	strHTML = ""
	
	Server.ScriptTimeOut=10000000
%>
<title> Reporte Tracking Aéreo.. </title>
<style>
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
@page
	{margin:.98in .79in .98in .79in;
	mso-header-margin:0in;
	mso-footer-margin:0in;}
tr
	{mso-height-source:auto;}
col
	{mso-width-source:auto;}
br
	{mso-data-placement:same-cell;}
.style21
	{mso-number-format:"_\(* \#\,\#\#0_\)\;_\(* \\\(\#\,\#\#0\\\)\;_\(* \0022-\0022_\)\;_\(\@_\)";
	mso-style-name:"Millares \[0\]_Hoja1";}
.style18
	{mso-number-format:"_-\0022$\0022* \#\,\#\#0\.00_-\;\\-\0022$\0022* \#\,\#\#0\.00_-\;_-\0022$\0022* \0022-\0022??_-\;_-\@_-";
	mso-style-name:Moneda;
	mso-style-id:4;}
.style0
	{mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	white-space:nowrap;
	mso-rotate:0;
	mso-background-source:auto;
	mso-pattern:auto;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial;
	mso-generic-font-family:auto;
	mso-font-charset:0;
	border:none;
	mso-protection:locked visible;
	mso-style-name:Normal;
	mso-style-id:0;}
.style22
	{mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	white-space:nowrap;
	mso-rotate:0;
	mso-background-source:auto;
	mso-pattern:auto;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial, sans-serif;
	mso-font-charset:0;
	border:none;
	mso-protection:locked visible;
	mso-style-name:Normal_Hoja1;}
.style20
	{mso-number-format:0%;
	mso-style-name:Porcentual;
	mso-style-id:5;}
td
	{mso-style-parent:style0;
	padding:0px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Arial;
	mso-generic-font-family:auto;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border:none;
	mso-background-source:auto;
	mso-pattern:auto;
	mso-protection:locked visible;
	white-space:nowrap;
	mso-rotate:0;}
.xl26
	{mso-style-parent:style0;
	text-align:center;}
.xl27
	{mso-style-parent:style0;
	text-align:center;
	background:white;
	mso-pattern:auto none;}
.xl28
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#0066FF;
	mso-pattern:auto none;
	white-space:normal;}
.xl29
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#003366;
	mso-pattern:auto none;
	white-space:normal;}
.xl30
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#003366;
	mso-pattern:auto none;
	white-space:normal;}
.xl31
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#003366;
	mso-pattern:auto none;
	white-space:normal;}
.xl32
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#003366;
	mso-pattern:auto none;
	white-space:normal;}
.xl33
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xrosa{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl34
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl35
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:0;
	text-align:center;
	vertical-align:middle;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl36
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl37
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl38
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl39
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:Standard;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl40
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl41
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:0;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl42
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:0;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#333399;
	mso-pattern:auto none;
	white-space:normal;}
.xl43
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:navy;
	mso-pattern:auto none;
	white-space:normal;}
.xl44
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:0;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:green;
	mso-pattern:auto none;
	white-space:normal;}
.xl45
	{mso-style-parent:style21;
	color:blue;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:yellow;
	mso-pattern:auto none;
	white-space:normal;}
.xl46
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:0;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#FF9900;
	mso-pattern:auto none;
	white-space:normal;}
.xl47
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:red;
	mso-pattern:auto none;
	white-space:normal;}
.xl48
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:blue;
	mso-pattern:auto none;
	white-space:normal;}
.xl49
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#99CC00;
	mso-pattern:auto none;
	white-space:normal;}
.xl50
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#993366;
	mso-pattern:auto none;
	white-space:normal;}
.xl51
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#33CCCC;
	mso-pattern:auto none;
	white-space:normal;}
.xl52
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#0066CC;
	mso-pattern:auto none;
	white-space:normal;}
.xl53
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	text-decoration:underline;
	text-underline-style:single;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"Medium Date";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#0066CC;
	mso-pattern:auto none;
	white-space:normal;}
.xl54
	{mso-style-parent:style18;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"_-\0022$\0022* \#\,\#\#0\.00_-\;\\-\0022$\0022* \#\,\#\#0\.00_-\;_-\0022$\0022* \0022-\0022??_-\;_-\@_-";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:purple;
	mso-pattern:auto none;
	white-space:normal;}
.xl55
	{mso-style-parent:style18;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"_-\0022$\0022* \#\,\#\#0\.00_-\;\\-\0022$\0022* \#\,\#\#0\.00_-\;_-\0022$\0022* \0022-\0022??_-\;_-\@_-";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#339966;
	mso-pattern:auto none;
	white-space:normal;}
.xl56
	{mso-style-parent:style21;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"_\(* \#\,\#\#0_\)\;_\(* \\\(\#\,\#\#0\\\)\;_\(* \0022-\0022_\)\;_\(\@_\)";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#666699;
	mso-pattern:auto none;
	white-space:normal;}
.xl57
	{mso-style-parent:style20;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:0%;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#666699;
	mso-pattern:auto none;
	white-space:normal;}
.xl58
	{mso-style-parent:style18;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"_-\0022$\0022* \#\,\#\#0\.00_-\;\\-\0022$\0022* \#\,\#\#0\.00_-\;_-\0022$\0022* \0022-\0022??_-\;_-\@_-";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#666699;
	mso-pattern:auto none;
	white-space:normal;}
.xl59
	{mso-style-parent:style18;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"_-\0022$\0022* \#\,\#\#0\.00_-\;\\-\0022$\0022* \#\,\#\#0\.00_-\;_-\0022$\0022* \0022-\0022??_-\;_-\@_-";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	background:red;
	mso-pattern:auto none;
	white-space:normal;}
.xl60
	{mso-style-parent:style18;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"_-\0022$\0022* \#\,\#\#0\.00_-\;\\-\0022$\0022* \#\,\#\#0\.00_-\;_-\0022$\0022* \0022-\0022??_-\;_-\@_-";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:#993300;
	mso-pattern:auto none;
	white-space:normal;}
.xl61
	{mso-style-parent:style18;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"_-\0022$\0022* \#\,\#\#0\.00_-\;\\-\0022$\0022* \#\,\#\#0\.00_-\;_-\0022$\0022* \0022-\0022??_-\;_-\@_-";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:#993300;
	mso-pattern:auto none;
	white-space:normal;}
.xl62
	{mso-style-parent:style18;
	color:white;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"_-\0022$\0022* \#\,\#\#0\.00_-\;\\-\0022$\0022* \#\,\#\#0\.00_-\;_-\0022$\0022* \0022-\0022??_-\;_-\@_-";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:#666699;
	mso-pattern:auto none;
	white-space:normal;}
.xl63
	{mso-style-parent:style21;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"0\.0%";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	background:yellow;
	mso-pattern:auto none;
	white-space:normal;}
.xl64
	{mso-style-parent:style21;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"0\.0%";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:yellow;
	mso-pattern:auto none;
	white-space:normal;}
.xl65
	{mso-style-parent:style18;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"_-\0022$\0022* \#\,\#\#0\.00_-\;\\-\0022$\0022* \#\,\#\#0\.00_-\;_-\0022$\0022* \0022-\0022??_-\;_-\@_-";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:yellow;
	mso-pattern:auto none;
	white-space:normal;}
.xl66
	{mso-style-parent:style21;
	font-weight:700;
	font-family:"Arial Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:"0\.0%";
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	background:yellow;
	mso-pattern:auto none;
	white-space:normal;}
.xl67
	{mso-style-parent:style22;
	font-family:Arial, sans-serif;
	mso-font-charset:0;
	text-align:center;
	background:white;
	mso-pattern:auto none;}
.xl68
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border:1.0pt solid silver;
	background:white;
	mso-pattern:auto none;}
.xl69
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:1.0pt solid silver;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:white;
	mso-pattern:auto none;}
.xl70
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:1.0pt solid silver;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:white;
	mso-pattern:auto none;}
.xl71
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	mso-number-format:"Short Date";
	text-align:center;
	border-top:1.0pt solid silver;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:white;
	mso-pattern:auto none;}
.xl72
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:1.0pt solid silver;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl73
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl74
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl75
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	mso-number-format:"Short Date";
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl76
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;}
.xl77
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:1.0pt solid silver;
	background:white;
	mso-pattern:auto none;}
.xl78
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:white;
	mso-pattern:auto none;}
.xl79
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:white;
	mso-pattern:auto none;}
.xl80
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	mso-number-format:"Short Date";
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:white;
	mso-pattern:auto none;}
.xl81
	{mso-style-parent:style0;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:white;
	mso-pattern:auto none;}
.xl82
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:red;
	mso-pattern:auto none;}
.xl83
	{mso-style-parent:style0;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;}
.xl84
	{mso-style-parent:style0;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl85
	{mso-style-parent:style0;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:none;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl86
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:1.0pt solid silver;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl87
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:red;
	mso-pattern:auto none;}
.xl88
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl89
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align:justify;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:white;
	mso-pattern:auto none;}
.xl90
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	text-align:center;
	vertical-align:justify;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl91
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	mso-number-format:"Short Date";
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl92
	{color:#000066;
	font-size:15pt;
	font-family:Arial, Helvetica, sans-serif;
	mso-font-charset:0;
	text-align:left;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	mso-pattern:auto none;}
.xl93
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:#CCFFFF;
	mso-pattern:auto none;}
.xl98
	{mso-style-parent:style0;
	color:black;
	font-size:7.5pt;
	font-family:Verdana, sans-serif;
	mso-font-charset:0;
	mso-number-format:"\@";
	text-align:center;
	border-top:none;
	border-right:1.0pt solid silver;
	border-bottom:1.0pt solid silver;
	border-left:none;
	background:white;
	mso-pattern:auto none;}

-->
</style>
</head>
<body>
<strong><font color="#000066" size="3" face="Arial, Helvetica, sans-serif">GRUPO REYES KURI, S.C. </font></strong><br>
<strong><font color="#969696" size="3" face="Arial, Helvetica, sans-serif"> TRACKING AEREO </font></strong>
	 
<table border=0 cellpadding=0 style='border-collapse: collapse;table-layout:fixed;width:5000pt'>
	 <col span = 33 style='mso-width-source:userset;mso-width-alt:5000;'>
	 <tr>
	<tr class=xl27>
<% genera_registros %>
</table>
</body>
</html>
<%

end if

sub genera_registros()
	dim c
	c=chr(34)

%>
<!-- Genera Encabezados -->
  <td class=xl28 id="_x0000_s1025" x:autofilter="all" width=100 style='height:39.0pt;width:94pt'>REFERENCIA</td>
  <td class=xl28 id="_x0000_s1026" x:autofilter="all" width=100  style='width:75pt'>ITTS/NOTIF DATE</td>
  <td class=xl28 id="_x0000_s1027" x:autofilter="all" width=100  style='width:75pt'>B. OF L. / AW. B. M. </td>
  <td class=xl28 id="_x0000_s1028" x:autofilter="all" width=100  style='width:75pt'>CONTAINER/ AW. B. H. </td>
  <td class=xl28 id="_x0000_s1029" x:autofilter="all" width=100  style='width:475pt'>CUSTOM OF DISPATCH </td>
  <td class=xl28 id="_x0000_s1030" x:autofilter="all" width=100  style='width:75pt'>IMPORT DOCUMENT</td>
  <td class=xl28 id="_x0000_s1031" x:autofilter="all" width=100  style='width:159pt'>INVOICE</td>
  <td class=xl28 id="_x0000_s1032" x:autofilter="all" width=300  style='width:159pt'>DESCRIPTION CODE</td>
  <td class=xl28 id="_x0000_s1033" x:autofilter="all" width=100  style='width:75pt'>MODEL</td>
  <td class=xl28 id="_x0000_s1034" x:autofilter="all" width=100  style='width:75pt'>DESCRIPTION</td> 
  <td class=xl28 id="_x0000_s1035" x:autofilter="all" width=100  style='width:75pt'>QUANTYTI</td>
  <td class=xl28 id="_x0000_s1036" x:autofilter="all" width=100  style='width:75pt'>ETA PORT/LAX </td>
  <td class=xl28 id="_x0000_s1037" x:autofilter="all" width=100  style='width:75pt'>SERIAL NUMBER </td> 
  <td class=xl28 id="_x0000_s1038" x:autofilter="all" width=100  style='width:75pt'>DATE OF RELEASE</td> 
  <td class=xl28 id="_x0000_s1039" x:autofilter="all" width=100  style='width:75pt'>AMOUNT OF DUTIES</td>
  <td class=xl28 id="_x0000_s1040" x:autofilter="all" width=100  style='width:75pt'>PREVIO</td>
  <td class=xl28 id="_x0000_s1041" x:autofilter="all" width=100  style='width:75pt'>DATE OF CLEARANCE</td>
  <td class=xl28 id="_x0000_s1042" x:autofilter="all" width=100  style='width:75pt'>ETA W/H</td>
  <td class=xl28 id="_x0000_s1043" x:autofilter="all" width=100  style='width:75pt'>STATUS</td>
  <td class=xl28 id="_x0000_s1044" x:autofilter="all" width=100  style='width:75pt'>KPI STATUS</td>
  <td class=xl28 id="_x0000_s1045" x:autofilter="all" width=100  style='width:75pt'>HISTORIAL</td>
  <td class=xl28 id="_x0000_s1046" x:autofilter="all" width=100  style='width:75pt'>DESCR RESULT DEL PREVIO</td>
  <td class=xl28 id="_x0000_s1047" x:autofilter="all" width=100  style='width:75pt'>PEDIMENTO</td>
  <td class=xl28 id="_x0000_s1048" x:autofilter="all" width=100  style='width:75pt'>CANTIDAD PARTIDAS</td>
  <td class=xl28 id="_x0000_s1049" x:autofilter="all" width=100  style='width:75pt'>VALOR MERCANCIAS</td>
  <td class=xl28 id="_x0000_s1050" x:autofilter="all" width=100  style='width:75pt'>FRACCION ARANCELARIA</td>
  <td class=xl28 id="_x0000_s1051" x:autofilter="all" width=100  style='width:75pt'>TASA IGI PROSEC SECTOR IIa</td>
  <td class=xl28 id="_x0000_s1052" x:autofilter="all" width=100  style='width:75pt'>FECHA RECIBIO</td>
  <td class=xl28 id="_x0000_s1053" x:autofilter="all" width=100  style='width:75pt'>HORA RECIBIO</td>
  <td class=xl28 id="_x0000_s1054" x:autofilter="all" width=100  style='width:75pt'>PERSONAL RECIBIO</td>
  <td class=xl28 id="_x0000_s1055" x:autofilter="all" width=100  style='width:75pt'>FECHA DE FACTURACION</td>
  <td class=xl28 id="_x0000_s1056" x:autofilter="all" width=100  style='width:75pt'>FECHA DE ENVIO DE FACTURA</td>
  <td class=xl28 id="_x0000_s1057" x:autofilter="all" width=100  style='width:75pt'>FECHA DE RECEPCION DE FACTURA</td>
 </tr>
<%
sqlAct=""

For i = 0 to 1
	Select Case i
		Case 0
			aduanaTmp = "dai"
		Case 1
			aduanaTmp = "tol"
	End Select'#ar.desc05 

	sqlAct= sqlAct & "select i.refcia01, r.frec01, r.adudes01, " &_
	  "f.numfac39 as 'Factura',     "&_
	  "i.adusec01 as 'AduanaSec',     "&_
	  "concat_ws(' ' ,i.patent01, i.numped01) as 'NumPedimento',    "&_
	  "i.numped01,    "&_
	  "r.feorig01 as 'F.BL',    "&_
	  " (SELECT  group_concat(distinct numgui04) from "&aduanaTmp&"_extranet.ssguia04 where refcia04=i.refcia01 and IDNGUI04 = 1 group by refcia04) as 'guiaMaster',"&_
	  " (SELECT  group_concat(distinct numgui04) from "&aduanaTmp&"_extranet.ssguia04 where refcia04=i.refcia01 and IDNGUI04 = 2 group by refcia04) as 'guiaHouse',"&_
	  "r.frev01 as 'F. ArriboAduana',     "&_
	  "r.fpre01 as 'F. Previo',     "&_
	  "r.fdsp01 as 'F.Desaduanamiento',    "&_
	  "i.fecent01,    "&_
	  "(SELECT SUM(fr.cancom02) FROM "&aduanaTmp&"_extranet.ssfrac02 fr WHERE fr.refcia02 = i.refcia01 and fr.fraarn02 = ar.frac05 and fr.ordfra02 = ar.agru05 ) as 'QUANTYTI',    "&_
	  "(SELECT numdia01 FROM "&aduanaTmp&"_extranet.d01reexp d WHERE d.cverex01 = r.cvrexp01 limit 1) as 'NUMDIASREEXP',    "&_
	  "r.etalax01,     "&_
	  " (SELECT  max(ordfra02) from "&aduanaTmp&"_extranet.ssfrac02 where refcia02=i.refcia01) as 'CANTPART',"&_
	  " (SELECT  SUM(vaduan02) from "&aduanaTmp&"_extranet.ssfrac02 where refcia02=i.refcia01) as 'Valor Aduana',"&_
	   "ar.descod05 as 'DESCRIPTION_CODE' ,"&_
	   "ar.cpro05 as 'MODEL' ,"&_
	  "i.cveped01 as 'cveped',    "&_
	   "replace(replace(replace(replace(ar.desc05,'\n',''),'\r',''),'\a',''),'\t','') as 'DESCRIPTION',    "&_
	   "ar.frac05 as 'far05'    "&_
		"from "&aduanaTmp&"_extranet.ssdagi01 as i     "&_
		"left join "&aduanaTmp&"_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01     "&_
		"left join "&aduanaTmp&"_extranet.c01refer as r on r.refe01 = i.refcia01    "&_
			"left join "&aduanaTmp&"_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01     "&_
			 " left join "&aduanaTmp&"_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01    "
		'Hay que cambiarlo para que se haga por oficina
		'Se Enlaza con el RFC del cliente porque busca en todas las oficinas y la cve del cte puede ser la de otro cliente en otra oficina				 
	sqlAct=sqlAct & "where cc.rfccli18 in ('SEM950215S98') "& Permi & " and i.firmae01 is not null and i.firmae01 <> '' and i.cveped01<>'R1'and  i.fecpag01 >='"&finicio&"' and i.fecpag01 <= '"&ffinal&"' "&_
	"group by i.refcia01,f.numfac39,ar.item05,ar.pfac05 "
		
	  
		if (i<>1) then
				sqlAct= sqlAct& " UNION ALL " & chr(13) & chr(10)
		end if
 Next 
	'response.write(sqlAct)
	'response.end()
	
	'Conexion a la base de datos
	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	
	dim refAnt, refActual, primerCiclo
	dim ref,refAux
	dim cambio
	cambio = 1
	primerCiclo = 0
	
	while not act2.eof
		response.Write("<tr align="&c&"center"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")
		ref = act2.fields("refcia01").value
		'esto cambia el color de fondo de la fila
		if (ref <> refAux)then
			if cambio = 1 then 
				cambio = 2
			else 
				if cambio = 2 then
					cambio = 1
				end if
			end if
		end if
		 
		if cambio = 1 then
			bgcolor="#D7ECF4"
		else
			bgcolor="#FFFFFF"
		end if 
			
		'comienza la impresion

		genera_html "d",act2.fields("refcia01").value,"center"  'Referencia 
		genera_html "d",formatofechaNum(act2.fields("frec01").value),"center"  'ITTSNOTIFDATE
		genera_html "d",act2.fields("guiaMaster").value,"center"  'GUIA MASTER
		if (act2.fields("AduanaSec").value = 200) then
			genera_html "d",retornaCampoDetalleCont(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"
		else 
			genera_html "d",act2.fields("guiaHouse").value,"center"
		end if
		StrCUSTOM_OF_DISPATCH = ""
		StrAdutmp = act2.Fields.Item("adudes01").Value
		if ltrim(StrAdutmp)="430" then
			StrCUSTOM_OF_DISPATCH = StrAdutmp&"-VERACRUZ" 'aduana de destino (en la que llega la mercancia directo de Origen)
		else
			if ltrim(StrAdutmp)="160" then
			  StrCUSTOM_OF_DISPATCH = StrAdutmp&"-MANZANILLO" 'aduana de destino (en la que llega la mercancia directo de Origen)
			else
				if ltrim(StrAdutmp)="200" or ltrim(StrAdutmp)="202" then
					StrCUSTOM_OF_DISPATCH = StrAdutmp&"-PANTACO" 'aduana de destino (en la que llega la mercancia directo de Origen)
				else
					if ltrim(StrAdutmp)="380" or ltrim(StrAdutmp)="810" then
						StrCUSTOM_OF_DISPATCH = StrAdutmp&"-TAMPICO" 'aduana de destino (en la que llega la mercancia directo de Origen)
					else
						if ltrim(StrAdutmp)="510" then
							StrCUSTOM_OF_DISPATCH = StrAdutmp&"-LAZARO CARDENAS" 'aduana de destino (en la que llega la mercancia directo de Origen)
						else
							if ltrim(StrAdutmp)="470" then
								StrCUSTOM_OF_DISPATCH = StrAdutmp&"-AEROPUERTO" 'aduana de destino (en la que llega la mercancia directo de Origen)
							else
								if ltrim(StrAdutmp)="650" then
									StrCUSTOM_OF_DISPATCH = StrAdutmp&"-TOLUCA" 'aduana de destino (en la que llega la mercancia directo de Origen)
								end if
							end if
						end if
					end if
				end if
			end if
		end if
		genera_html "d",StrCUSTOM_OF_DISPATCH,"center" 
		genera_html "d",act2.fields("NumPedimento").value,"center"
		genera_html "d",act2.fields("Factura").value,"center"  'Factura
		genera_html "d",act2.fields("DESCRIPTION_CODE").value,"center"  'DESCRIPCIO_CODE
		genera_html "d",act2.fields("MODEL").value,"center"  'DESCRIPCIO_CODE
		genera_html "d",act2.fields("DESCRIPTION").value,"center" 'DESCRIPTION
		genera_html "d",act2.fields("QUANTYTI").value,"center"  'QUANTYTI
		genera_html "d",formatofechaNum( act2.fields("etalax01").value ),"center"  'ETA PORT/LAX
		genera_html "d",retornaSerie(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"  'SERIAL NUMBER
		genera_html "d",formatofechaNum( act2.Fields.Item("F. ArriboAduana").Value ),"center"  'DATE OF RELEASE 
		' Para no repetir
		refActual = act2.fields("refcia01").value
		primerCiclo = primerCiclo + 1
		if refActual = refAnt then
			genera_html "d","","center"  'AMOUNT OF DUTIES 
		else
			genera_html "d",AMOUNTOFDUTIES(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"  'AMOUNT OF DUTIES 
		end if
		genera_html "d",formatofechaNum( act2.Fields.Item("F. Previo").Value ),"center"  'FECHA DE PREVIO
		genera_html "d",formatofechaNum( act2.Fields.Item("F.Desaduanamiento").Value ),"center"  'DATE OF CLEARANCE
		genera_html "d",ETA_W_H(act2.fields("fecent01").value,act2.fields("etalax01").value,act2.fields("F.Desaduanamiento").value,act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"  'ETA W/H
		genera_html "d",STATUS(act2.fields("fecent01").value,act2.fields("F.BL").value,act2.fields("F.Desaduanamiento").value,act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"'STATUS
		genera_html "d",KPISTATUS(act2.fields("fecent01").value,act2.fields("etalax01").value,act2.fields("F.Desaduanamiento").value,act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"  'KPI STATUS
		genera_html "d",HISTORIAL(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"  'HISTORIAL
		genera_html "d",DESCRIPCION_RESULTADO_PREVIO(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"  'DESCR RESULT DEL PREVIO
		genera_html "d",act2.Fields.Item("numped01").Value,"center" 'PEDIMENTO
		genera_html "d",act2.Fields.Item("CANTPART").Value,"center" 'CANTIDAD DE PARTIDAS
		if refActual = refAnt then
			genera_html "d","","center"  'VALOR DE  LAS MERCANCIAS
		else
			genera_html "d",act2.Fields.Item("Valor Aduana").Value,"center" 'VALOR DE  LAS MERCANCIAS
		end if
		genera_html "d",act2.Fields.Item("far05").Value,"center" '	FRACCION ARANCELARIA
		genera_html "d",regresa_tasa_IGI_PROSEC(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center" 'TASA IGI PROSEC SECTOR IIa

		dim rsRecibe
		rsRecibe = REGRESA_RECIBE(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3))
		dim Fecha,Recibe,Hora
		if(isarray(rsRecibe))then
			for y = 0 to Ubound(rsRecibe,2)
				if y = 0 then
					Fecha = trim(rsRecibe(0,y))
					Recibe = trim(rsRecibe(1,y)) 
					Hora = trim(rsRecibe(2,y))
				else
					Fecha = Fecha & trim(rsRecibe(0,y))
					Recibe = Recibe & trim(rsRecibe(1,y)) 
					Hora = Hora & trim(rsRecibe(2,y)) 
				end if
			next
		end if
		genera_html "d",formatofechaNum(Fecha),"center" ''FECHA SE RECIBE EN ALMACEN
		genera_htmltexto "d",Hora,"center" ''FECHA SE RECIBE EN ALMACEN
		genera_html "d",Recibe,"center" ''FECHA SE RECIBE EN ALMACEN
		
		
		genera_html "d",regresa_fecha_cuenta_gastos(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"  'FECHA DE FACTURACION
		genera_html "d",FECHA_ENVIO_FACTURA(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"  'FECHA DE ENVIO DE FACTURA
		genera_html "d",regresa_fecha_recepcion_cuenta_gastos(act2.fields("refcia01").value,mid(act2.fields("refcia01").value,1,3)),"center"  'FECHA DE RECEPCION DE FACTURA
		
		refAnt = refActual

		response.Write("</tr>")
		act2.movenext()
	wend

end sub

sub genera_html(tipo,valor,alineacion)
	if(tipo = "e")then
		response.Write("<td width="&c&"100"&c&" align="&c&alineacion&c&" nowrap bgcolor="&c&"#CCFF99"&c&"><div align="&c&alineacion&c&"><strong><em><font size="&c&"2"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></em></strong></div></td>")
	else 
		if bgcolor ="#D7ECF4" then
			response.Write("<td align="&c&alineacion&c&" class=xl73>"&valor&"</td>")
		else
			response.Write("<td align="&c&alineacion&c&" class=xl78>"&valor&"</td>")
		end if
	end if
end sub

sub genera_htmltexto(tipo,valor,alineacion)
	if(tipo = "e")then
		response.Write("<td width="&c&"100"&c&" align="&c&alineacion&c&" nowrap bgcolor="&c&"#CCFF99"&c&"><div align="&c&alineacion&c&"><strong><em><font size="&c&"2"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></em></strong></div></td>")
	else 
		if bgcolor ="#D7ECF4" then
			response.Write("<td align="&c&alineacion&c&" class=xl93>"&valor&"</td>")
		else
			response.Write("<td align="&c&alineacion&c&" class=xl98>"&valor&"</td>")
		end if
	end if
end sub

Function pd(n, totalDigits)
        if totalDigits > len(n) then
            pd = String(totalDigits-len(n),"0") & n
        else
            pd = n
        end if
End Function
	
Function formatofechaNum(DFecha)
       if isdate( DFecha ) then
          formatofechaNum = YEAR(DFecha) & Pd(Month( DFecha ),2) & Pd(DAY( DFecha ),2)
       else
          formatofechaNum	= DFecha
       end if
End Function

'--********************************************************************************----
function retornaCampoCtaGastos(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
 
 
 sqlAct = "select r."& campo &" as campo from "&oficina&"_extranet.e31cgast as cta " &_
 " inner join  "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
 " where  r.refe31 = '"& referencia &"' and cta.esta31 <> 'C' "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCampoCtaGastos = valor
 else
  retornaCampoCtaGastos =valor
 end if
end function

function retornaCampoDetalleCont(referencia,oficina)
	dim valor
	valor=""
 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
 
	sqlAct="SELECT D.refe01 ,group_concat(D.marc01 separator ' - ' ) as campo FROM "&oficina&"_EXTRANET.d01conte AS D WHERE D.refe01 ='"&referencia&"' "

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	
	if not(act2.eof) then
		valor = act2.fields("campo").value
	end if
	retornaCampoDetalleCont =valor
end function

function retornaSerie(referencia,oficina)
	if oficina="ALC" then oficina="LZR"
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	
 strCERTNOM  = ""
 StrNUMSERIE = ""
 if referencia <> "" then
	 Set RFecDocu = Server.CreateObject("ADODB.Recordset")
	 RFecDocu.ActiveConnection = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	 strSqlSel =  " SELECT distinct de.CLAV07,  " & _
				  "         de.FECH07, " & _
				  "         de.ORIG07, " & _
				  "         DISP07            " & _
				  " FROM "& oficina &"_extranet.C07DOCRE de " & _
				  " WHERE  de.REFE07 ='"&ltrim(referencia)&"' AND " & _
				  "       (de.CLAV07='CNO' or " & _
				  "        de.clav07='CNS' )"
	 'Response.Write(strSqlSel)
	 'Response.End
	 RFecDocu.Source = strSqlSel
	 RFecDocu.CursorType = 0
	 RFecDocu.CursorLocation = 2
	 RFecDocu.LockType = 1
	 RFecDocu.Open()
	 While NOT RFecDocu.EOF
		 if RFecDocu.Fields.Item("CLAV07").Value <>"" and ltrim(RFecDocu.Fields.Item("CLAV07").Value) = "CNO"  then
			  if RFecDocu.Fields.Item("DISP07").Value = "F"   then
				 strCERTNOM  = "N/A"
			  else
				 strCERTNOM  = RFecDocu.Fields.Item("FECH07").Value
			  end if
		 else
			if RFecDocu.Fields.Item("CLAV07").Value <>"" and ltrim(RFecDocu.Fields.Item("CLAV07").Value) = "CNS"  then
				 if RFecDocu.Fields.Item("DISP07").Value = "F"   then
					StrNUMSERIE = "N/A"
				 else
					StrNUMSERIE = RFecDocu.Fields.Item("FECH07").Value
				 end if
			end if
		 end if
		 RFecDocu.movenext
	 Wend
	 RFecDocu.close
	 set RFecDocu = Nothing
 end if

if isdate( StrNUMSERIE ) then
  retornaSerie = YEAR( StrNUMSERIE ) & Pd(Month( StrNUMSERIE ),2) & Pd(DAY( StrNUMSERIE ),2)  'SERIAL NUMBER
else
  retornaSerie = StrNUMSERIE  'SERIAL NUMBER
end if
  
end function

function AMOUNTOFDUTIES(referencia,oficina)
	if oficina="ALC" then oficina="LZR"
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	
 strImpuestos = ""
	if referencia <> "" then
		Set RImpuestos = Server.CreateObject("ADODB.Recordset")
		RImpuestos.ActiveConnection = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
		strSqlSel =  " SELECT SUM(import36) as Impuestos " & _
				  " FROM "& oficina &"_extranet.sscont36         " & _
				  " WHERE  REFCIA36 = '"&ltrim(referencia)&"'  AND " & _
				  "        FPAGOI36 = 0 " & _
				  " GROUP BY refcia36 "
		 'Response.Write(strSqlSel)
		 'Response.End
		RImpuestos.Source = strSqlSel
		RImpuestos.CursorType = 0
		RImpuestos.CursorLocation = 2
		RImpuestos.LockType = 1
		RImpuestos.Open()
		if not RImpuestos.eof then
			strImpuestos = RImpuestos.Fields.Item("Impuestos").Value
		else
			strImpuestos = ""
		end if
		RImpuestos.close
		set RImpuestos = Nothing
	end if
 
AMOUNTOFDUTIES = strImpuestos
  
end function

Function diasTrimFinSemana(DInicio, DFin)
	 x_Dias = 0
	 x_Dias = dateDiff("d", DInicio , DFin )

	 if x_Dias > 0 then
	   x_Con=1
	   x_finSemana=0
	   Do While (x_Con <= x_Dias)
		  x_diasemana=WeekDay( DateAdd("d",x_Con,  DInicio ) )
		  if x_diasemana=1 or x_diasemana=7 then
			 x_finSemana = x_finSemana +1
		  end if
		  x_Con = x_Con + 1
	   loop
	 x_Dias = x_Dias - x_finSemana ' Restamos los dias de fin de semana
	 end if
	 diasTrimFinSemana = x_Dias

End Function


Function SumarDiasSinFinSemana(DFecha,IntDayAdd)
	 x_Dias = 0
	 x_Dias = IntDayAdd
	 if x_Dias > 0 then
	   x_Con=1
	   x_finSemana=0
	   Do While (x_Con <= x_Dias)
		  x_diasemana=WeekDay( DateAdd("d",x_Con,  DFecha ) )
		  if x_diasemana=1 or x_diasemana=7 then
			 x_finSemana = x_finSemana +1
		  end if
		  x_Con = x_Con + 1
	   loop
	 x_Dias = x_Dias + x_finSemana ' sumamos los dias de fin de semana
	 end if
	 DNewFecha = DateAdd("d",x_Dias, DFecha  )

	 numDia= WeekDay( DNewFecha )
	 if numDia=1 then ' domingo
		DNewFecha = DateAdd("d",1, DNewFecha  )
	 else
		if numDia=7 then ' Sabado
			DNewFecha = DateAdd("d",2, DNewFecha  )
		end if
	 end if
	 SumarDiasSinFinSemana =  DNewFecha

End Function

Function SumarDias(DFecha,IntDayAdd,intType)
  if isdate(DFecha) then
	 if intType = 1 then ' dias Naturales
		SumarDias = DateAdd("d",IntDayAdd,  DFecha )
	 else ' dias habiles
		'if intType = 2 then
		  SumarDias = SumarDiasSinFinSemana(DFecha,IntDayAdd)
		'end if
	 end if
  else
	SumarDias = DFecha
  end if
End Function

function ETA_W_H(DFechEntAux,etalax01,DATE_CUSTOM,referencia,oficina)
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	
	if isdate(DFechEntAux) then
		if DFechEntAux > date() then
			DFechEntAux = ""
		end if
	end if
	
	
	strSQLPlSTD =" SELECT  " & _
				 "		D.n_orden as orden, " & _
				 "		E.d_abrev as inicio, " & _
				 "		B.d_abrev as fin, " & _
				 "	   I.transal as modalidad,  " & _
				 "		I.numdia01 as dias,  " & _
				 "		I.tipdia01 as tipod  " & _
				 "	FROM  " & _
				 "		"& oficina &"_status.ETXPL AS D " & _
				 "		INNER JOIN "& oficina &"_status.ETAPS AS E ON D.n_etapa = E.n_etapa " & _
				 "		INNER JOIN "& oficina &"_status.D01STD AS I ON  E.N_ETAPA= I.etpini01  " & _
				 "		INNER JOIN "& oficina &"_status.ETAPS AS B ON  B.n_etapa = I.etpfin01 " & _
				 "	WHERE " & _
				 "		D.n_plantilla = 1  " & _
				 "		AND I.tipoadu='AEREA' " & _
				 "		order by D.n_orden "
				 
	Set RsPlSTD = Server.CreateObject("ADODB.Recordset")
	RsPlSTD.ActiveConnection = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	RsPlSTD.Source = strSQLPlSTD
	RsPlSTD.CursorType = 0
	RsPlSTD.CursorLocation = 2
	RsPlSTD.LockType = 1
	RsPlSTD.Open()
   
	StdATAPORTDSP = 0 'std para ATAPORT A DESPACHO
	StdDSPWH      = 0 'std para DESPACHO A WAREHOUSE

	tipoStdEtdLoad    = 1 'tipo de dias de std ETDLOAD
	tipoStdATAPORTDSP = 1 'tipo de dias de std ATAPORT A DESPACHO
	tipoStdDSPWH      = 1 'tipo de dias de std DESPACHO A WAREHOUSE

	if not RsPlSTD.eof then
		While NOT RsPlSTD.EOF
			if RsPlSTD.Fields.Item("inicio").Value = "ATAPORT" and RsPlSTD.Fields.Item("fin").Value = "DSP" then
				StdATAPORTDSP     = RsPlSTD.Fields.Item("dias").Value
				tipoStdATAPORTDSP = RsPlSTD.Fields.Item("tipod").Value
			else
				if RsPlSTD.Fields.Item("inicio").Value = "DSP" and RsPlSTD.Fields.Item("fin").Value = "LLP" then
					StdDSPWH     = RsPlSTD.Fields.Item("dias").Value
					tipoStdDSPWH = RsPlSTD.Fields.Item("tipod").Value
				end if
			end if
			RsPlSTD.movenext
		wend
	end if
   RsPlSTD.close
   set RsPlSTD = Nothing
		   
	   
	if isdate(DFechEntAux) then
		StrETA_CUSTOM_CLEARANCE = SumarDias( DFechEntAux , StdATAPORTDSP,tipoStdATAPORTDSP)
	else
		StrETA_CUSTOM_CLEARANCE = SumarDias( etalax01 , StdATAPORTDSP,tipoStdATAPORTDSP)
	end if
	
	if isdate( StrETA_CUSTOM_CLEARANCE ) then
		if isdate( DATE_CUSTOM ) then
			IndFila = DateDiff("d",StrETA_CUSTOM_CLEARANCE , DATE_CUSTOM )
			if IndFila = 0 then
				StrColorfila = 1
				StrETA_W_H_AUX = SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH )
			else
				StrETA_W_H_AUX = SumarDias( DATE_CUSTOM , StdDSPWH, tipoStdDSPWH )
				if IndFila < 0 then
					StrColorfila = 2
				else
					StrColorfila = 3
				end if
			end if
		else
			StrETA_W_H_AUX = SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH )
			IndFila = DateDiff("d", StrETA_CUSTOM_CLEARANCE , DATE() )
			if IndFila > 0 then
				StrColorfila = 3
			end if
		end if
	else
		if isdate( DATE_CUSTOM ) then
			StrETA_W_H_AUX = SumarDias( DATE_CUSTOM , StdDSPWH, tipoStdDSPWH )
		else
			StrETA_W_H_AUX = SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH )
		end if
	end if
	
	ETA_W_H = formatofechaNum(StrETA_W_H_AUX)

end function

function STATUS(DFechEntAux,DFecOri,DATE_CUSTOM,referencia,oficina)
	strStatusTmp = ""
	if strFechaATAWH <> "" then
		strStatusTmp = "SEM"
	else
		if DATE_CUSTOM <> "" then
			strStatusTmp = "ADUANA"
			else
			if DFechEntAux <> "" then
				strStatusTmp = "AEROPUERTO"
			else
				if DFecOri <> "" then
					strStatusTmp = "TRANSITO AEREO"
				end if
			end if
		end if
	end if
	
	STATUS = strStatusTmp
end function

function KPISTATUS(DFechEntAux,etalax01,DATE_CUSTOM,referencia,oficina)
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	
	if isdate(DFechEntAux) then
		if DFechEntAux > date() then
			DFechEntAux = ""
		end if
	end if
	
	strSQLPlSTD =" SELECT  " & _
				 "		D.n_orden as orden, " & _
				 "		E.d_abrev as inicio, " & _
				 "		B.d_abrev as fin, " & _
				 "	   I.transal as modalidad,  " & _
				 "		I.numdia01 as dias,  " & _
				 "		I.tipdia01 as tipod  " & _
				 "	FROM  " & _
				 "		"& oficina &"_status.ETXPL AS D " & _
				 "		INNER JOIN "& oficina &"_status.ETAPS AS E ON D.n_etapa = E.n_etapa " & _
				 "		INNER JOIN "& oficina &"_status.D01STD AS I ON  E.N_ETAPA= I.etpini01  " & _
				 "		INNER JOIN "& oficina &"_status.ETAPS AS B ON  B.n_etapa = I.etpfin01 " & _
				 "	WHERE " & _
				 "		D.n_plantilla = 1  " & _
				 "		AND I.tipoadu='AEREA' " & _
				 "		order by D.n_orden "
				 
	Set RsPlSTD = Server.CreateObject("ADODB.Recordset")
	RsPlSTD.ActiveConnection = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	RsPlSTD.Source = strSQLPlSTD
	RsPlSTD.CursorType = 0
	RsPlSTD.CursorLocation = 2
	RsPlSTD.LockType = 1
	RsPlSTD.Open()
   
	StdATAPORTDSP = 0 'std para ATAPORT A DESPACHO
	StdDSPWH      = 0 'std para DESPACHO A WAREHOUSE

	tipoStdEtdLoad    = 1 'tipo de dias de std ETDLOAD
	tipoStdATAPORTDSP = 1 'tipo de dias de std ATAPORT A DESPACHO
	tipoStdDSPWH      = 1 'tipo de dias de std DESPACHO A WAREHOUSE

	if not RsPlSTD.eof then
		While NOT RsPlSTD.EOF
			if RsPlSTD.Fields.Item("inicio").Value = "ATAPORT" and RsPlSTD.Fields.Item("fin").Value = "DSP" then
				StdATAPORTDSP     = RsPlSTD.Fields.Item("dias").Value
				tipoStdATAPORTDSP = RsPlSTD.Fields.Item("tipod").Value
			else
				if RsPlSTD.Fields.Item("inicio").Value = "DSP" and RsPlSTD.Fields.Item("fin").Value = "LLP" then
					StdDSPWH     = RsPlSTD.Fields.Item("dias").Value
					tipoStdDSPWH = RsPlSTD.Fields.Item("tipod").Value
				end if
			end if
			RsPlSTD.movenext
		wend
	end if
   RsPlSTD.close
   set RsPlSTD = Nothing
   
	if isdate(DFechEntAux) then
		StrETA_CUSTOM_CLEARANCE = SumarDias( DFechEntAux , StdATAPORTDSP,tipoStdATAPORTDSP)
	else
		StrETA_CUSTOM_CLEARANCE = SumarDias( etalax01 , StdATAPORTDSP,tipoStdATAPORTDSP)
	end if
	
	if isdate( StrETA_CUSTOM_CLEARANCE ) then
		if isdate( DATE_CUSTOM ) then
			IndFila = DateDiff("d",StrETA_CUSTOM_CLEARANCE , DATE_CUSTOM )
			if IndFila = 0 then
				StrColorfila = 1
				StrETA_W_H_AUX = SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH )
			else
				StrETA_W_H_AUX = SumarDias( DATE_CUSTOM , StdDSPWH, tipoStdDSPWH )
				if IndFila < 0 then
					StrColorfila = 2
				else
					StrColorfila = 3
				end if
			end if
		else
			StrETA_W_H_AUX = SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH )
			IndFila = DateDiff("d", StrETA_CUSTOM_CLEARANCE , DATE() )
			if IndFila > 0 then
				StrColorfila = 3
			end if
		end if
	else
		if isdate( DATE_CUSTOM ) then
			StrETA_W_H_AUX = SumarDias( DATE_CUSTOM , StdDSPWH, tipoStdDSPWH )
		else
			StrETA_W_H_AUX = SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH )
		end if
	end if



	if isdate(strFechaATAWH) then
		if isdate(DFechEntAux) then
			intoTD = DiasTrimFinSemana( DFechEntAux ,strFechaATAWH )
		else
			if isdate( etalax01 ) then
				intoTD = DiasTrimFinSemana( etalax01 , strFechaATAWH )
			else
				intoTD = 0
			end if
		end if
	else
		if isdate(StrETA_W_H_AUX) then
			if isdate(DFechEntAux) then
				intoTD = DiasTrimFinSemana( DFechEntAux , StrETA_W_H_AUX )
			else
				if isdate( etalax01 ) then
					intoTD = DiasTrimFinSemana( etalax01 , StrETA_W_H_AUX )
				else
					intoTD = 0
				end if
			end if
		else
			intoTD = 0
		end if
	end if
									   
	strKPISTTmp  = ""
	
	if intoTD <= 2 then
		strKPISTTmp = "ON TIME"
	else
		strKPISTTmp = "DELAY"
	end if
	
	KPISTATUS = strKPISTTmp
end function

function HISTORIAL(referencia,oficina)
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	
	strObservaciones = ""
	if referencia <> "" then
		Set RObservEtapas = Server.CreateObject("ADODB.Recordset")
		RObservEtapas.ActiveConnection = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	 
		strSQL = " SELECT (n_secuenc), " & _
			  "        D.n_etapa,   " & _
			  "        f_fecha,     " & _
			  "        m_observ     " & _
			  " FROM "& oficina &"_status.ETXPD as D     " & _
			  " WHERE not(date_format(D.f_fecha,'%Y%m%d') = '00000000') and  " & _
			  "       D.c_referencia = '"&ltrim(referencia)&"' and n_etapa <>6 " & _
			  " ORDER BY N_ETAPA, N_SECUENC "

		RObservEtapas.Source = strSQL
		RObservEtapas.CursorType = 0
		RObservEtapas.CursorLocation = 2
		RObservEtapas.LockType = 1
		RObservEtapas.Open()
		intcontObs = 1
		While NOT RObservEtapas.EOF
			strObsTemp = RObservEtapas.Fields.Item("m_observ").Value
			if strObsTemp <>"" and ltrim(strObsTemp) <> "" and InStr( strObservaciones, strObsTemp) = 0 then
				if intcontObs = 1 then
					strObservaciones  =RObservEtapas.Fields.Item("m_observ").Value
				else
					strObservaciones  = strObservaciones & " ; "& RObservEtapas.Fields.Item("m_observ").Value
				end if
				intcontObs = intcontObs + 1
			end if
			RObservEtapas.movenext
		Wend
		RObservEtapas.close
		set RObservEtapas = Nothing
	end if

	'Esto es del otro reporte repLaySamsungAereo_excel.asp
	
	'if strComentarioATAWH <> "" AND ltrim(strComentarioATAWH) <> "" then
	'	 strObservaciones = strObservaciones&" ; "& strComentarioATAWH
	'  end if
	'  if strComentarioATAC_P <> "" and ltrim(strComentarioATAC_P) <> "" then
	'	 strObservaciones = strObservaciones&" ; "& strComentarioATAC_P
	'   end if
	'   if strComentarioETAW_H <> "" and ltrim(strComentarioETAW_H) <> "" then
	'	 strObservaciones = strObservaciones&" ; "& strComentarioETAW_H
	'   end if
	'   if strComentarioATASPLTMP <> "" and ltrim(strComentarioATASPLTMP) <> "" then
	'	 strObservaciones = strObservaciones&" ; "& strComentarioATASPLTMP
	'   end if
	HISTORIAL = strObservaciones
end function

function regresa_fecha_cuenta_gastos(referencia,oficina)
	dim c,valor
	c=chr(34)
	valor="PENDIENTE"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	 
	 
	sqlAct="select max(date_format(cta.fech31,'%d/%m/%Y')) as fech31 from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
	" where cta.cgas31 = r.cgas31 and not(date_format(cta.fech31,'%Y%m%d') = '00000000') and "&_
	" r.refe31 = '"&referencia&"' and cta.esta31 <> 'C' "

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	
	
	if not(act2.eof) then
		if isdate( act2.fields("fech31").value ) then
			  regresa_fecha_cuenta_gastos = YEAR( act2.Fields.Item("fech31").Value ) & Pd(Month( act2.Fields.Item("fech31").Value ),2) & Pd(DAY( act2.Fields.Item("fech31").Value ),2)
		else
			  regresa_fecha_cuenta_gastos = act2.Fields.Item("fech31").Value
		end if
	else
	  regresa_fecha_cuenta_gastos =valor
	end if

end function

function regresa_fecha_recepcion_cuenta_gastos(referencia,oficina)
	dim c,valor
	c=chr(34)
	valor="PENDIENTE"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	 
	 
	sqlAct="select max(date_format(cta.frec31,'%d/%m/%Y')) as frec31 from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
	" where cta.cgas31 = r.cgas31 and not(date_format(cta.frec31,'%Y%m%d') = '00000000') and "&_
	" r.refe31 = '"&referencia&"' and cta.esta31 <> 'C' "

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	
	
	if not(act2.eof) then
		if isdate( act2.fields("frec31").value ) then
			  regresa_fecha_recepcion_cuenta_gastos = YEAR( act2.Fields.Item("frec31").Value ) & Pd(Month( act2.Fields.Item("frec31").Value ),2) & Pd(DAY( act2.Fields.Item("frec31").Value ),2)
		else
			  regresa_fecha_recepcion_cuenta_gastos = act2.Fields.Item("frec31").Value
		end if
	else
	  regresa_fecha_recepcion_cuenta_gastos =valor
	end if

end function

function regresa_tasa_IGI_PROSEC(referencia,oficina) 'Tasa de Igi con prosec para el Sector IIa.
	valor="N/A"
	 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	 
	 
	sqlQry="select distinct fr.refcia02,fr.tasadv02 as campo "&_
			"from "&oficina&"_extranet.ssfrac02 fr "&_
			"inner join "&oficina&"_extranet.ssipar12 ip on fr.refcia02 = ip.refcia12 and fr.ordfra02 = ip.ordfra12 "&_
			"where ip.refcia12 = '"&referencia&"' and ip.cveide12 ='PS' and ip.comide12 ='IIa' "

	Set Retas= Server.CreateObject("ADODB.Recordset")
	conn= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	Retas.ActiveConnection = conn
	Retas.Source = sqlQry
	Retas.cursortype=0
	Retas.cursorlocation=2
	Retas.locktype=1
	Retas.open()
	
	if not(Retas.eof) then
		valor = Retas.Fields.Item("campo").Value
		Retas.movenext()
		while not Retas.eof
			valor = valor&", "&Retas.fields("campo").value
			Retas.movenext()
		wend
	end if

	regresa_tasa_IGI_PROSEC = valor
end function

function DESCRIPCION_RESULTADO_PREVIO(referencia,oficina)
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
	
	strObservaciones = ""
	if referencia <> "" then
		Set RObservEtapas = Server.CreateObject("ADODB.Recordset")
		RObservEtapas.ActiveConnection = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	 
		strSQL = " SELECT (n_secuenc), " & _
			  "        n_etapa,   " & _
			  "        f_fecha,     " & _
			  "        m_observ     " & _
			  " FROM "& oficina &"_status.ETXPD as D     " & _
			  " WHERE not(date_format(D.f_fecha,'%Y%m%d') = '00000000') and  " & _
			  "       D.c_referencia = '"&ltrim(referencia)&"' AND n_etapa = 6" & _
			  " ORDER BY N_SECUENC "

		RObservEtapas.Source = strSQL
		RObservEtapas.CursorType = 0
		RObservEtapas.CursorLocation = 2
		RObservEtapas.LockType = 1
		RObservEtapas.Open()
		intcontObs = 1
		While NOT RObservEtapas.EOF
			strObsTemp = RObservEtapas.Fields.Item("m_observ").Value
			if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
				if intcontObs = 1 then
					strObservaciones  =RObservEtapas.Fields.Item("m_observ").Value
				else
					strObservaciones  = strObservaciones & " ; "& RObservEtapas.Fields.Item("m_observ").Value
				end if
				intcontObs = intcontObs + 1
			end if
			RObservEtapas.movenext
		Wend
		RObservEtapas.close
		set RObservEtapas = Nothing
	end if

	DESCRIPCION_RESULTADO_PREVIO = strObservaciones
end function

function FECHA_ENVIO_FACTURA(referencia,oficina)
	dim valor
	valor=""
 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
 
 
	sqlAct =	"select distinct d35.fech35 as campo from "&oficina&"_extranet.d35entcg as d35 " &_
				" inner join "&oficina&"_extranet.e31cgast as cta on d35.foli35 = cta.foli31 " &_
				" inner join  "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
				" where  r.refe31 = '"& referencia &"' and cta.esta31 <> 'C' and not(date_format(d35.fech35,'%Y%m%d') = '00000000') "

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	if not(act2.eof) then
		valor = formatofechaNum(act2.fields("campo").value)
		act2.movenext()
		while not act2.eof
			valor = valor&", "& formatofechaNum(act2.fields("campo").value)
			act2.movenext()
		wend
	end if
	act2.close
	set act2 = Nothing
		
	FECHA_ENVIO_FACTURA = formatofechaNum(valor)
end function

function REGRESA_RECIBE(referencia,oficina)
 
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	end if
	if (ucase(oficina) = "PAN")then
		oficina = "DAI"
	end if
  
	sqlAct =	"select distinct e.refch01 ,e.recibe01,e.rehora01 from "&oficina&"_extranet.d01conte as d01 " &_
				" inner join "&oficina&"_extranet.e01oemb as e on e.peri01=d01.peri01 and d01.nemb01=e.nemb01 " &_
				" where  d01.refe01 = '"& referencia &"' and d01.nemb01<>0 "

	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()

	Dim array
	array = act2.getRows
	
	act2.close
	set act2 = Nothing
	 
	REGRESA_RECIBE = array
end function
%>
