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
<%
On Error Resume Next
'																										 		'
'																												'
' ---------------------------------------        IDOT       ------------------------------------------------	'
'																												'
' 																												'
'se ejecuta de la siguiente forma:
'http://10.66.1.9/portalmysql/extranet/ext-asp/reportes/samsung-idot-17082012.asp?finicio=2012-01-01&ffinal=2012-01-15&tipope=i&det=

Response.Buffer = TRUE
response.Charset = "utf-8"
Response.Addheader "Content-Disposition", "attachment; filename=BookletUNILEVER_IDOT_.xls"'
Response.ContentType = "application/vnd.ms-excel"

dim strTipoUsuario,fechaini,fechafin,oficina

strTipoUsuario = Session("GTipoUsuario")
tipope	 = Request.QueryString("tipope")
fechaini = Request.QueryString("finicio")
fechafin = Request.QueryString("ffinal")
'oficina = "RKU"

if not fechaini="" and not fechafin="" then
    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    finicio = tmpAnioIni & "-" &tmpMesIni & "-"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    ffinal = tmpAnioFin & "-" &tmpMesFin & "-"& tmpDiaFin

	dim orden(50)
	dim subrefaux,subref,bgcolor,strHTML
	subrefaux=""
	subref=""
	bgcolor="#FFFFFF"
	strHTML = ""
	
	Server.ScriptTimeOut=10000000
%>
<title> ReporteIDOT.. </title>
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
	background:#003366;
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
-->
</style>
</head>
<body>
<table x:str border=0 cellpadding=0 cellspacing=0 width=12637 style='border-collapse: collapse;table-layout:fixed;width:9479pt'>
	 <col width=125 style='mso-width-source:userset;mso-width-alt:4571;width:94pt'>
	 <col width=100 span=2 style='mso-width-source:userset;mso-width-alt:3657; width:75pt'>
	 <col class=xl26 width=370 style='mso-width-source:userset;mso-width-alt:8265; width:370pt'>
	 <col width=100 span=4 style='mso-width-source:userset;mso-width-alt:3657; width:75pt'>
	 <col width=259 style='mso-width-source:userset;mso-width-alt:7753;width:259pt'>
	 <col width=100 span=11 style='mso-width-source:userset;mso-width-alt:3657; width:75pt'>
	 <col class=xl26 width=214 style='mso-width-source:userset;mso-width-alt:7826; width:161pt'>
	 <col width=100 span=8 style='mso-width-source:userset;mso-width-alt:3657; width:75pt'>
	 <col class=xl26 width=100 style='mso-width-source:userset;mso-width-alt:3657; width:75pt'>
	 <col width=100 span=67 style='mso-width-source:userset;mso-width-alt:3657; width:75pt'>
	 <col width=80 span=32 style='width:60pt'>
	<tr class=xl27 height=52 style='height:39.0pt'>
<% genera_registros tipope %>
</table>
</body>
</html>
<%
else
	response.write("Algo esta mal en las fechas")

end if

sub genera_registros(tipope)
	
%>

<td height=52 class=xl28 id="_x0000_s1025" x:autofilter="all"
  x:autofilterrange="$A$1:$CS$800" width=125 style='height:39.0pt;width:94pt'
  x:str="DIVISION ">Division<span style='mso-spacerun:yes'> </span></td>
  <td class=xl33 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>No. De Trafico</td>
  <td class=xl29 id="_x0000_s1027" x:autofilter="all" width=100
  style='width:75pt'>Importancia(Planta de Entrega)</td>
  <td class=xl29 id="_x0000_s1028" x:autofilter="all" width=100
  style='width:75pt'>Categoria</td>
  <td class=xl29 id="_x0000_s1029" x:autofilter="all" width=475
  style='width:475pt'>Nombre del Material</td>
  <td class=xl28 id="_x0000_s1030" x:autofilter="all" width=100
  style='width:75pt'>Clase de Producto</td>
  <td class=xl31 id="_x0000_s1031" x:autofilter="all" width=212
  style='width:159pt'>Proveedor</td>
  <td class=xl31 id="_x0000_s1032" x:autofilter="all" width=212
  style='width:159pt'>Dimicilio Fiscal del Proveedor</td>
  <td class=xl33 id="_x0000_s1033" x:autofilter="all" width=100
  style='width:75pt'>TAX ID/ RFC</td>
  <td class=xl33 id="_x0000_s1034" x:autofilter="all" width=100
  style='width:75pt'>País de Procedencia</td> 
  <td class=xl33 id="_x0000_s1035" x:autofilter="all" width=100
  style='width:75pt' x:str="País de Origen ">País de Origen</td>
 <td class=xl33 id="_x0000_s1036" x:autofilter="all" width=100
  style='width:75pt'>Estado de procedencia</td> 
 <td class=xl33 id="_x0000_s1037" x:autofilter="all" width=100
  style='width:75pt'>Ciudad de procedencia</td> 
 <td class=xl33 id="_x0000_s1038" x:autofilter="all" width=100
  style='width:75pt'>Código Postal de procedencia</td>
 <td class=xl33 id="_x0000_s1039" x:autofilter="all" width=100
  style='width:75pt'>Region</td>
 <td class=xl33 id="_x0000_s1040" x:autofilter="all" width=100
  style='width:75pt'>PTO./CD DE ORIGEN</td>
 <td class=xl28 id="_x0000_s1041" x:autofilter="all" width=100
  style='width:75pt' x:str="Cuenta ">Cuenta</td>
  <td class=xl28 id="_x0000_s1042" x:autofilter="all" width=100
  style='width:75pt'>CECO</td>
  <td class=xl32 id="_x0000_s1043" x:autofilter="all" width=100
  style='width:75pt'>ODC</td>
  <td class=xl28 id="_x0000_s1044" x:autofilter="all" width=100
  style='width:75pt'>No IE</td>
   <td class=xl36 id="_x0000_s1045" x:autofilter="all" width=100
  style='width:75pt'>Factura</td>
  <td class=xl38 id="_x0000_s1046" x:autofilter="all" width=214
  style='width:161pt'>IMPORTADOR</td>
    <td class=xl39 id="_x0000_s1047" x:autofilter="all" width=100
  style='width:75pt' x:str="Cantidad ">Cantidad<span
  style='mso-spacerun:yes'> </span></td>
   <td class=xl33 id="_x0000_s1048" x:autofilter="all" width=100
  style='width:75pt'>Unidad de Medida</td>
  <td class=xl39 id="_x0000_s1049" x:autofilter="all" width=100 
  style='width:75pt' x:str="Peso Bruto ">Peso Bruto KG<span
  style='mso-spacerun:yes'> </span></td>
    <td class=xl33 id="_x0000_s1050" x:autofilter="all" width=100
  style='width:75pt'>Incoterms</td>
    <td class=xl33 id="_x0000_s1051" x:autofilter="all" width=100
  style='width:75pt'>Tipo de Transporte</td>
  <td class=xl33 id="_x0000_s1052" x:autofilter="all" width=100
  style='width:75pt'>Aduana</td>
  <td class=xl33 id="_x0000_s1053" x:autofilter="all" width=100
  style='width:75pt'>Agente Aduanal</td>
  <td class=xl40 id="_x0000_s1054" x:autofilter="all" width=100
  style='width:75pt'>Patente Agente Aduanal</td>
  <td class=xl33 id="_x0000_s1055" x:autofilter="all" width=100
  style='width:75pt'>Referencia del Agente Aduanal</td>
  <td class=xl38 id="_x0000_s1056" x:autofilter="all" width=100
  style='width:75pt' x:str="No de Contenedor ">No de Contenedor</td>
  <td class=xl33 id="_x0000_s1057" x:autofilter="all" width=100
  style='width:75pt'>Clave Pedimento</td>
  <td class=xl33 id="_x0000_s1058" x:autofilter="all" width=100
  style='width:75pt'>No. Pedimento</td>
  <td class=xl33 id="_x0000_s1059" x:autofilter="all" width=100
  style='width:75pt'>Fecha Pedimento</td>
   <td class=xl40 id="_x0000_s1060" x:autofilter="all" width=100
  style='width:75pt'>Mes</td>
  <td class=xl38 id="_x0000_s1061" x:autofilter="all" width=100
  style='width:75pt'>No.Semana de Operación</td>  
  <td class=xl41 id="_x0000_s1062" x:autofilter="all" width=100
  style='width:75pt' x:str="Cantidad de Operaciones ">Cantidad de Operaciones</td>
  <td class=xl41 id="_x0000_s1063" x:autofilter="all" width=100
  style='width:75pt'>Cantidad de Contenedores</td>
  <td class=xl41 id="_x0000_s1064" x:autofilter="all" width=100
  style='width:75pt' x:str="PALLETS/BULTOS ">PALLETS/BULTOS<span
  style='mso-spacerun:yes'> </span></td>
    <td class=xl42 id="_x0000_s1065" x:autofilter="all" width=100
  style='width:75pt'>MEDIDA DEL CONTENEDOR</td>
    <td class=xl42 id="_x0000_s1066" x:autofilter="all" width=100
  style='width:75pt'>TIPO DE CAJA(Seca/refrigerada)</td>
    <td class=xl37 id="_x0000_s1067" x:autofilter="all" width=100
  style='width:75pt'>Fecha de Factura</td>
  <td class=xl41 id="_x0000_s1068" x:autofilter="all" width=100
  style='width:75pt'>Fecha BL</td>
  <td class=xl43 id="_x0000_s1069" x:autofilter="all" width=100
  style='width:75pt'>Fecha de arribo a la aduana</td>
  <td class=xl43 id="_x0000_s1070" x:autofilter="all" width=100
  style='width:75pt'>Fecha Desaduanamiento</td>
  <td class=xl44 id="_x0000_s1071" x:autofilter="all" width=100
  style='width:75pt'>KPI Programación</td>
   <td class=xl44 id="_x0000_s1072" x:autofilter="all" width=100
  style='width:75pt'>KPI Tránsito</td>
  <td class=xl44 id="_x0000_s1073" x:autofilter="all" width=100
  style='width:75pt'>KPI Desaduanamiento</td>
  <td class=xl44 id="_x0000_s1077" x:autofilter="all" width=100
  style='width:75pt'>KPI lead TIME</td>
 <td class=xl45 id="_x0000_s1078" x:autofilter="all" width=100
  style='width:75pt'>TARGET TIME</td>
 <td class=xl53 id="_x0000_s1079" x:autofilter="all" width=100
  style='width:75pt'>NUMERO DE EMBARQUE</td>
  <td class=xl53 id="_x0000_s1080" x:autofilter="all" width=100
  style='width:75pt' x:str="IDOT ">IDOT<span style='mso-spacerun:yes'> </span></td>
 <td class=xl53 id="_x0000_s1081" x:autofilter="all" width=100
  style='width:75pt' x:str="IDOT LT ">IDOT LT<span style='mso-spacerun:yes'> </span></td>
  <td class=xl53 id="_x0000_s1082" x:autofilter="all" width=100
  style='width:75pt' x:str="CAUSA IDOT ">CAUSA IDOT<span style='mso-spacerun:yes'> </span></td>
  <td class=xl53 id="_x0000_s1083" x:autofilter="all" width=100
  style='width:75pt' x:str="CAUSA RAIZ ">CAUSA RAIZ<span style='mso-spacerun:yes'> </span></td>
  <td class=xl53 id="_x0000_s1084" x:autofilter="all" width=100
  style='width:75pt' x:str="PLAN DE ACCION ">PLAN DE ACCION<span style='mso-spacerun:yes'> </span></td>
  <td class=xl53 id="_x0000_s1085" x:autofilter="all" width=100
  style='width:75pt' x:str="RESPONSABLE">RESPONSABLE<span style='mso-spacerun:yes'> </span></td>
  <td class=xl53 id="_x0000_s1086" x:autofilter="all" width=100
  style='width:75pt' x:str="FECHA DE CUMPLIMIENTO ">FECHA DE CUMPLIMIENTO<span style='mso-spacerun:yes'> </span></td>
  <td class=xl53 id="_x0000_s1087" x:autofilter="all" width=100
  style='width:75pt' x:str="IMPACTO ">IMPACTO<span style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1088" x:autofilter="all" width=100
  style='width:75pt' x:str="No. CTA DE GASTOS"><span  style='mso-spacerun:yes'> </span>No. CTA DE GASTOS<span style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1089" x:autofilter="all" width=100
  style='width:75pt' x:str="FECHA CTA DE GASTOS"><span  style='mso-spacerun:yes'> </span>FECHA CTA DE GASTOS<span  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1090" x:autofilter="all" width=100
  style='width:75pt' x:str="TIPO DE CTA DE GASTOS"><span  style='mso-spacerun:yes'> </span>TIPO DE CTA DE GASTOS<span  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1091" x:autofilter="all" width=100
  style='width:75pt' x:str="Precio Pagado / valor comercial"><span  style='mso-spacerun:yes'> </span>Precio Pagado / valor comercial<span  style='mso-spacerun:yes'> </span></td>
    <td class=xl54 id="_x0000_s1092" x:autofilter="all" width=100
  style='width:75pt' x:str=" Valor comercial USD"><span
  style='mso-spacerun:yes'>  </span>Valor comercial USD<span
  style='mso-spacerun:yes'> </span></td>
    <td class=xl54 id="_x0000_s1093" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR FLETES INTERNACIONAL M.N."><span
  style='mso-spacerun:yes'> </span>VALOR FLETES INTERNACIONAL M.N.<span
  style='mso-spacerun:yes'> </span></td>
   <td class=xl53 id="_x0000_s1094" x:autofilter="all" width=100
  style='width:75pt'>VALOR FLETES INTERNACIONAL USD</td>
    <td class=xl54 id="_x0000_s1095" x:autofilter="all" width=100
  style='width:75pt' x:str="SEGUROS"><span
  style='mso-spacerun:yes'> </span>SEGUROS<span
  style='mso-spacerun:yes'> </span></td>
   <td class=xl54 id="_x0000_s1096" x:autofilter="all" width=100
  style='width:75pt' x:str="OTROS INCREMENTABLES"><span
  style='mso-spacerun:yes'> </span>OTROS INCREMENTABLES<span
  style='mso-spacerun:yes'> </span></td> 
    <td class=xl53 id="_x0000_s1097" x:autofilter="all" width=100
  style='width:75pt' x:str="OTROS INCREMENTABLES USD"><span
  style='mso-spacerun:yes'> </span>OTROS INCREMENTABLES USD<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1098" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR ADUANA M.N."><span
  style='mso-spacerun:yes'> </span>VALOR ADUANA M.N.<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1099" x:autofilter="all" width=100
  style='width:75pt' x:str="T.C."><span
  style='mso-spacerun:yes'> </span>T.C.<span style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1092" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR ADUANA  DLLS"><span
  style='mso-spacerun:yes'> </span>VALOR ADUANA DLLS<span style='mso-spacerun:yes'> 
  </span>DLLS<span style='mso-spacerun:yes'> </span></td>
    <td class=xl55 id="_x0000_s1100" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR FLETES AEREO DLLS"><span
  style='mso-spacerun:yes'> </span>VALOR FLETES AEREO DLLS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl55 id="_x0000_s1101" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR FLETES TERRESTRE DLLS."><span
  style='mso-spacerun:yes'> </span>VALOR FLETES TERRESTRE DLLS.<span
  style='mso-spacerun:yes'> </span></td>
    <td class=xl55 id="_x0000_s1102" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR FLETES MARITIMO DLLS."><span
  style='mso-spacerun:yes'> </span>VALOR FLETES MARITIMO DLLS.<span
  style='mso-spacerun:yes'> </span></td>
   <td class=xl56 id="_x0000_s1103" x:autofilter="all" width=100
  style='width:75pt' x:str="FRACC. ARANC."><span
  style='mso-spacerun:yes'> </span>FRACC. ARANC.<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl57 id="_x0000_s1104" x:autofilter="all" width=100
  style='width:75pt'>ARANCEL %</td>
  <td class=xl57 id="_x0000_s1105" x:autofilter="all" width=100
  style='width:75pt' x:str="ARANCEL PREFERENCIAL ">ARANCEL PREFERENCIAL<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl58 id="_x0000_s1106" x:autofilter="all" width=100
  style='width:75pt' x:str="MONTO DE RECUPERACION $ "><span
  style='mso-spacerun:yes'> </span>MONTO DE RECUPERACION $<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl58 id="_x0000_s1107" x:autofilter="all" width=100
  style='width:75pt' x:str="ADV. $ / IGI $"><span
  style='mso-spacerun:yes'> </span>ADV. $ / IGI $<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl58 id="_x0000_s1108" x:autofilter="all" width=100
  style='width:75pt' x:str="DTA $"><span style='mso-spacerun:yes'> </span>DTA
  $<span style='mso-spacerun:yes'> </span></td>
  <td class=xl57 id="_x0000_s1109" x:autofilter="all" width=100
  style='width:75pt'>IVA % </td>
  <td class=xl58 id="_x0000_s1110" x:autofilter="all" width=100
  style='width:75pt' x:str="IVA $"><span style='mso-spacerun:yes'> </span>IVA
  $<span style='mso-spacerun:yes'> </span></td>
  <td class=xl58 id="_x0000_s1111" x:autofilter="all" width=100
  style='width:75pt' x:str="PREVAL. "><span
  style='mso-spacerun:yes'> </span>PREVAL.<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl58 id="_x0000_s1112" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL IMPUESTOS"><span
  style='mso-spacerun:yes'> </span>TOTAL IMPUESTOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl59 id="_x0000_s1113" x:autofilter="all" width=100
  style='width:75pt' x:str="Total Impuestos USD "><span
  style='mso-spacerun:yes'> </span>Total Impuestos USD<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl60 id="_x0000_s1114" x:autofilter="all" width=100
  style='width:75pt' x:str="GTOS. ADUANA USD(SOLO FRONTERA)"><span
  style='mso-spacerun:yes'> </span>GTOS. ADUANA USD(SOLO FRONTERA)<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1115" x:autofilter="all" width=100
  style='width:75pt' x:str="DEMORAS"><span
  style='mso-spacerun:yes'> </span>DEMORAS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1116" x:autofilter="all" width=100
  style='width:75pt' x:str="ESTADIAS"><span
  style='mso-spacerun:yes'> </span>ESTADIAS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1117" x:autofilter="all" width=100
  style='width:75pt' x:str="MANIOBRAS "><span
  style='mso-spacerun:yes'> </span>MANIOBRAS<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl61 id="_x0000_s1118" x:autofilter="all" width=100
  style='width:75pt' x:str="ALMACENAJES"><span
  style='mso-spacerun:yes'> </span>ALMACENAJES<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1119" x:autofilter="all" width=100
  style='width:75pt' x:str="OTROS"><span
  style='mso-spacerun:yes'> </span>OTROS<span style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1120" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS"><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl59 id="_x0000_s1121" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS USD"><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS USD<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl53 id="_x0000_s1122" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS M.N."><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS M.N.<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl53 id="_x0000_s1123" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS USD"><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS USD<span
  style='mso-spacerun:yes'> </span></td>
    <td class=xl58 id="_x0000_s1124" x:autofilter="all" width=100
  style='width:75pt'>HONORARIOS AG AD. $</td>
  <td class=xl58 id="_x0000_s1125" x:autofilter="all" width=100
  style='width:75pt'>NAVIERA. $</td>
  <td class=xl58 id="_x0000_s1126" x:autofilter="all" width=100
  style='width:75pt'>TRANSPORTISTA NACIONAL</td>
    <td class=xl58 id="_x0000_s1127" x:autofilter="all" width=100
  style='width:75pt'>COSTO FLETE NACIONAL</td>>
    <td class=xl33 id="_x0000_s1128" x:autofilter="all" width=100
  style='width:75pt' x:str="TIPO TRANSPORTE"></td>
  <td class=xl54 id="_x0000_s1129" x:autofilter="all" width=100
  style='width:75pt' x:str="TIPO DE UNIDAD"></td>

<%
sqlAct=""
For i = 0 to 4
	
		Select Case i
				Case 0
					aduanaTmp = "rku"
		
				Case 1
					aduanaTmp = "dai"
			
				Case 2
					aduanaTmp = "sap"
	
				Case 3
					aduanaTmp = "tol"
				Case 4
					aduanaTmp="lzr"
			
						
		End Select
sqlAct= sqlAct & "select i.refcia01,fr2.fraarn02,fr2.ordfra02, " &_
  "i.cvecli01 as '1', "&_
  "'' as 'NameMat', "&_
  "prv.nompro22 as 'Proveedor', "&_
  "r.rcli01 as 'RefCliente', "&_
  "ar.pedi05 as 'ODC', "&_
  "i.cvepod01 as 'Pais Origen', "&_
  "i.cvepvc01 as 'Pais procedencia', "&_
  "r.ptoemb01 as 'Pto/Ciudad Origen', "&_
  "r.cveptoemb as 'cve embark',    "&_
  "prv.irspro22 as 'TaxID',     "&_
  "  concat_ws(' ',prv.dompro22 ,prv.ciupro22 ,prv.estpro22 ,'CP. ',prv.c_ppro22) as 'Domicilio Prove', "&_
  "f.numfac39 as 'Factura',     "&_
  "date_format(f.fecfac39,'%d/%m/%Y') as 'Fec.Fac',     "&_
  "r.impo01 as '21',     "&_
  "ar.caco05 as 'Cantidad',    "&_
  "um.descri31 as 'Unidad de medida',    "&_
  "i.pesobr01 as 'PesoB KG',    "&_
  "f.terfac39 as 'Imcoterms',    "
  if aduanaTmp="rku" or aduanaTmp="sap" or aduanaTmp="lzr" then
  sqlAct= sqlAct & "'MARITIMO' as '25',    "
  else
  sqlAct=sqlAct &"'AEREO' as '25',    "
  end if
  sqlAct=sqlAct &"i.adusec01 as 'AduanaSec',     "&_
  "i.patent01 as 'Patente',     "&_
"  i.refcia01 as 'Refe',    "&_
  "concat_ws(' ',DATE_FORMAT(i.fecpag01,'%y'),left(i.adusec01,2), i.patent01, i.numped01) as 'NumPedimento',    "&_
  "i.fecpag01 as 'FechaPagoPed',    "&_
  "Month(i.fecpag01) as 'Mes',     "&_
  "week(i.fecpag01) as 'Sem',    "&_
  "date_format(f.fecfac39,'%d/%m/%Y') as 'FechaFactura',    "&_
  "r.feorig01 as 'F.BL',    "&_
  "r.frev01 as 'F. ArriboAduana',     "&_
  "r.fdsp01 as 'F.Desaduanamiento',    "&_
  "fr.prepag02   as 'Presio Pag',     "&_
  "fr2.prepag02   as '590',     "&_
  "(fr.prepag02/i.tipcam01) as 'ValCom USD',     "&_
 "(fr2.prepag02/i.tipcam01) as '600',     "&_
 "'' as '67',      "&_
 "'' as '68', "&_
  "i.fletes01 as 'Valor Fletes MN',    "&_
  "i.segros01 as 'Seguros',    "&_
  "i.incble01 as 'Otros inc.',     "&_
  "fr.vaduan02 as 'Val Adua',     "&_
  "fr2.vaduan02 as '640',     "&_
  "CAST(i.tipcam01 AS CHAR) as 'TC',     "&_
  "(fr.vaduan02/i.tipcam01) as 'Val Adu DLS',    "&_
  "(fr2.vaduan02/i.tipcam01) as '660',    "&_
  "(i.fletes01/i.tipcam01)  as 'Fletes',     "&_
  "fr.fraarn02 as 'Frac.Ar', 	    "&_
  "ifnull(fr2.fraarn02,0) as '710', 	    "&_
  "fr.tasadv02 as 'Arancel',      "&_
  "fr2.tasadv02 as '720',      "&_
  "if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as 'Aran.Pref',     "&_
  "cf6.import36 as '75 ',    "&_
  "cf1.import36 as 'Total Imp',    "&_
  "(fr.i_adv102+fr.i_adv202) as '761',   "&_
  "(fr2.i_adv102+fr2.i_adv202) as '7610',   "&_
  "fr.tasiva02 as 'Iva',    "&_
  "fr2.tasiva02 as '770',    "&_
  "cf3.import36  as '78',    "&_
  "(fr.i_iva102+fr.i_iva202) as '781',   "&_
  "(fr2.i_iva102+fr2.i_iva202) as '7810',   "&_
  "cf15.import36  as 'Preval',     "&_
  "fr.ordfra02 as 'Ord FrcA',"&_
  "count(fr.ordfra02) as '82',    "&_
 "fr2.ordfra02 as '810',"&_
  "count(fr2.ordfra02) as '820',    "&_
   "ar2.item05 as 'Item05' ,"&_
	 "i.firmae01 as 'firmita',     "&_
   "transp.descri30  as 'sDescTransp',     "&_
   "ar2.fact05  as 'facmer',     "&_
  "i.cveped01 as 'cveped',    "&_
   "ar2.ffactp05  as 'ffac05',     "&_
   "ar2.pedi05 as 'odc05',    "&_
   "replace(replace(replace(replace(ar2.desc05,'\n',''),'\r',''),'\a',''),'\t','') as 'mer05x',    "&_
    "'' as 'mer05',    "&_
   "ar2.frac05 as 'far05',    "&_
   "ar2.caco05 as 'canco',"&_
	"GROUP_CONCAT(DISTINCT nav.nom01)  AS Naviera "&_
    "from "&aduanaTmp&"_extranet.ssdag"&tipope&"01 as i     "&_
    "left join "&aduanaTmp&"_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01     "&_
    "left join "&aduanaTmp&"_extranet.c01refer as r on r.refe01 = i.refcia01    "&_
  "LEFT join "&aduanaTmp&"_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01     "&_
  "LEFT join "&aduanaTmp&"_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C'     "&_
        "left join "&aduanaTmp&"_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01     "&_
        "left join "&aduanaTmp&"_extranet.c06barco as bar on bar.clav06 =r.cbuq01 "&_
        "left join "&aduanaTmp&"_extranet.c55navie as nav on nav.cve01 =bar.navi06 "&_
         " left join "&aduanaTmp&"_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01    "&_
          "left join "&aduanaTmp&"_extranet.d05artic as ar2 on ar2.refe05 = i.refcia01    "&_
            "left join "&aduanaTmp&"_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05      "&_
 			 "left join "&aduanaTmp&"_extranet.ssfrac02 as fr2 on i.refcia01 = fr2.refcia02  and ar2.frac05 = fr.fraarn02  and fr.ordfra02 = ar2.agru05      "&_
              "left join "&aduanaTmp&"_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01      "&_
                "left join "&aduanaTmp&"_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02      "&_
                     "left join "&aduanaTmp&"_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'      "&_
                     "left join "&aduanaTmp&"_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'      "&_
                     "left join "&aduanaTmp&"_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'      "&_
                     "left join "&aduanaTmp&"_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'      "&_
                     "left join "&aduanaTmp&"_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL')    "&_
                     "left join "&aduanaTmp&"_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01       "&_
     "where cc.rfccli18 in ('SEM950215S98') and i.firmae01 is not null and i.firmae01 <> ''  and  i.fecpag01 >='"&fechaini&"' and i.fecpag01 <= '"&fechafin&"'"&_
  "group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05  "
  
	if (i<>4) then
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
	
'response.write(sqlAct)
'response.end()
	dim refAnt, refActual, primerCiclo, pesoBruto
	dim ref,refAux
	dim cambio,rcli 
	cambio = 1
	rcli=""
	primerCiclo = 0
	
	while not act2.eof
		response.Write("<tr align="&c&"center"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")
		ref = act2.fields("Refe").value
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
		if (act2.fields("cveped").value<>"G1") then
			genera_html "d",retornaDivision(act2.fields("1").value,act2.fields("Frac.Ar").value),"center"  'DIVISION 
			genera_html "d",act2.fields("Refe").value,"center"  'No. De Trafico
			genera_html "d",retornaPlantaEntrega(act2.fields("1").value),"center"  'PLANTA DE ENTREGA
			genera_html "d","","center"  'CATEGORIA
			genera_html "d",act2.fields("NameMat").value,"center"  'Nombre del Material
			genera_html "d","","center"  'Clase de Producto
			genera_html "d",act2.fields("Proveedor").value,"center"  'Proveedor
			genera_html "d",act2.fields("Domicilio Prove").value,"center" 'Domicilio fiscal del proveedor 
			genera_html "d",act2.fields("TaxID").value,"center"  'TAX ID/ RFC
				if(act2.fields("Pto/Ciudad Origen").value<>"")then
					PuertEmb= retornaCampoPuertoEmb(act2.fields("Pto/Ciudad Origen").value,act2.fields("cve embark").value,"cvepai01",mid(act2.fields("Refe").value,1,3))	
				else
					PuertEmb= ""	
				end if 			
			genera_html "d",PuertEmb,"center"  'País de Procedencia
			genera_html "d",act2.fields("Pais Origen").value,"center"  'País de Origen 
			genera_html "d","","center" 'genera_html "d",act2.fields("EstaOri").value,"center" 'Estado de origen
			genera_html "d","","center" 'genera_html "d",act2.fields("Pto/Ciudad Origen").value,"center" 'Ciudad de origen
			genera_html "d","","center" 'genera_html "d",act2.fields("CP.Ori").value,"center" 'Codigo Postal de Origen
			genera_html "d",retornaRegion(act2.fields("Pais Origen").value),"center"  'Region
			genera_html "d",act2.fields("Pto/Ciudad Origen").value,"center" 'PTO./CD DE ORIGEN
			rcli =act2.fields("RefCliente").value
			genera_html "d",retornaCuenta(rcli),"center"  'Cuenta 
			genera_html "d",retornaCECO(rcli),"center"  'CECO
			genera_html "d",act2.fields("ODC").value,"center"  'ODC
			genera_html "d","","center"  'No IE
			genera_html "d",act2.fields("Factura").value,"center"  'Factura
			genera_html "d",retornaIMPORTADOR(act2.fields("21").value,mid(act2.fields("Refe").value,1,3)),"center"  'IMPORTADOR
			genera_html "d",act2.fields("Cantidad").value,"center"  'Cantidad 
		else
			genera_html "d",retornaDivision(act2.fields("1").value,act2.fields("710").value),"center"  'DIVISION 
			genera_html "d",act2.fields("Refe").value,"center"  'No. De Trafico
			genera_html "d",retornaPlantaEntrega(act2.fields("1").value),"center"  'PLANTA DE ENTREGA
			genera_html "d","","center"  'CATEGORIA
			genera_html "d",act2.fields("mer05").value,"center"  'Nombre del Material
			genera_html "d","","center"  'Clase de Producto
			genera_html "d",act2.fields("Proveedor").value,"center"  'Proveedor
			genera_html "d",act2.fields("Proveedor").value,"center"  'Domicilio fiscal del proveedor --- NUEVO
			genera_html "d",act2.fields("TaxID").value,"center"  'TAX ID/ RFC
			if(act2.fields("Pto/Ciudad Origen").value<>"")then
				PuertEmb= retornaCampoPuertoEmb(act2.fields("Pto/Ciudad Origen").value,act2.fields("cve embark").value,"cvepai01",mid(act2.fields("Refe").value,1,3))	
			else
				PuertEmb= ""	
			end if 
			genera_html "d",PuertEmb,"center"  'País de Procedencia
			genera_html "d",act2.fields("Pais Origen").value,"center"  'País de Origen 
			genera_html "d","","center" 	'genera_html "d",act2.fields("EstaOri").value,"center" 'Estado de origen
			genera_html "d","","center" 	'genera_html "d",act2.fields("Pto/Ciudad Origen").value,"center" 'Ciudad de origen
			genera_html "d","","center" 	'genera_html "d",act2.fields("CP.Ori").value,"center" 'Codigo Postal de Origen
			genera_html "d",retornaRegion(act2.fields("Pais Origen").value),"center"  'Region
			genera_html "d",act2.fields("Pto/Ciudad Origen").value,"center" 'PTO./CD DE ORIGEN
			rcli =act2.fields("RefCliente").value
			genera_html "d",retornaCuenta(rcli),"center"  'Cuenta 
			genera_html "d",retornaCECO(rcli),"center"  'CECO
			genera_html "d",act2.fields("odc05").value,"center"  'ODC G1
			genera_html "d","","center"  'No IE
			genera_html "d",act2.fields("facmer").value,"center" 'factura de G1
			genera_html "d",retornaIMPORTADOR(act2.fields("21").value,mid(act2.fields("Refe").value,1,3)),"center"  'IMPORTADOR
			genera_html "d",act2.fields("canco").value,"center"  'Cantidad 
							
		end if 'Termina G1 de Divicion a Cantidad
		
			' Para no repetir el peso bruto --
			refActual = act2.fields("Refe").value
			primerCiclo = primerCiclo + 1
			pesoBruto = act2.fields("PesoB KG").value
		
		if primerCiclo > 1 and refActual = refAnt then
			pesoBruto = " "
		end if
			
		refAnt = refActual
		genera_html "d",act2.fields("Unidad de medida").value,"center"  'Unidad de Medida
		genera_html "d", pesoBruto ,"center"  'Peso Bruto
		genera_html "d",act2.fields("Imcoterms").value,"center"  'Incoterms
		genera_html "d",act2.fields("25").value,"center"  'Tipo de Transporte
		genera_html "d",retornaAduana(act2.fields("AduanaSec").value),"center"  'Aduana
		genera_html "d",retornaAgenteAduanal(act2.fields("Patente").value),"center"  'Agente Aduanal
		genera_html "d",act2.fields("Patente").value,"center"  'Patente Agente Aduanal
		genera_html "d",act2.fields("Refe").value,"center"  'No. De Trafico

			ref = act2.fields("Refe").value
		if (ref <> refAux)then
			refAux=ref
			'Lote 2
			genera_html "d",retornaCampoContenedores(act2.fields("Refe").value,"numcon40",mid(act2.fields("Refe").value,1,3)),"center"  'No de Contenedor 
			genera_html "d",act2.fields("cveped").value,"center" 'Cve pedimento
			genera_html "d",act2.fields("NumPedimento").value,"center"  'No. Pedimento
			genera_html "d",act2.fields("FechaPagoPed").value,"center"  'Fecha Pedimento
			genera_html "d",act2.fields("Mes").value,"center"  'Mes
			genera_html "d",act2.fields("Sem").value,"center"  'No.Semana
			genera_html "d","1","center"  'Cantidad de Operaciones 
			genera_html "d",retornaCantContenedores40(act2.fields("Refe").value,"numcon40",mid(act2.fields("Refe").value,1,3)),"center"  'Cantidad de Contenedores
			genera_html "d",retornaCantContenedores(act2.fields("Refe").value,"'BUL','CAJ','BID','PAL'",mid(act2.fields("Refe").value,1,3)),"center"  'PALLETS/BULTOS 
				tipo=retornaTipoContenedores(act2.fields("Refe").value,mid(act2.fields("Refe").value,1,3))'TIPO DE CONTENEDOR/ CAJA
				n=len(tipo)
				dimension =left(tipo,2)
				tipo=mid(tipo,4,(n-3))
				caja=TipoCaja(act2.fields("Refe").value,mid(act2.fields("Refe").value,1,3))'Tipo Caja
			genera_html "d",dimension,"center"  'Medida del contenedor
			if tipo<>"" then 
				genera_html "d",caja,"center"  'Tipo de Caja (seca/Refrigerada)
			else
				genera_html "d","","center"  'Tipo de Caja (seca/Refrigerada)
			end if
			
			genera_html "d",act2.fields("FechaFactura").value,"center"  'Fecha Factura
			genera_html "d",act2.fields("F.BL").value,"center"  'Fecha BL
			genera_html "d",act2.fields("F. ArriboAduana").value,"center"  'Fecha de arribo a la aduana
			genera_html "d",act2.fields("F.Desaduanamiento").value,"center"  'Fecha Desaduanamiento
			genera_html "d","","center"  'KPI Programación
			genera_html "d","","center"  'KPI Tránsito
			genera_html "d","","center"  'KPI Desaduanamiento
			genera_html "d","","center"  'KPI lead TIME
			genera_html "d","","center"  'TARGET TIME
			genera_html "d","","center"  'NUMERO DE EMBARQUE
			genera_html "d","","center"  'IDOT
			genera_html "d","","center"  'IDOT LT
			genera_html "d","","center" 'CAUSA IDOT
			genera_html "d","","center"  'CAUSA RAIZ
			genera_html "d","","center"  'PLANTA DE ACCION
			genera_html "d","","center"  'RESPONSABLE
			genera_html "d","","center"  'FECHA DE CUMPLIMIENTO
			genera_html "d","","center"  'IMPACTO
			genera_html "d",retornaCampoCtaGastos(act2.fields("Refe").value,"cgas31",mid(act2.fields("Refe").value,1,3)),"center"  'No. CTA DE GASTOS
			genera_html "d",regresa_fecha_cuenta_gastos(act2.fields("Refe").value,mid(act2.fields("Refe").value,1,3)),"center"  'Fecha C.Gastos
			genera_html "d",regresa_tipo_Cgastos(act2.fields("Refe").value,mid(act2.fields("Refe").value,1,3)),"center"    'Tipo C.Gastos
			Subref = act2.fields("Ord FrcA").value
				if (ordenOcupado(Subref,act2.fields("Refe").value) = False)then
					ocuparOrd(Subref)
					'-----
					if (act2.fields("cveped").value<>"G1") then
						genera_html "d",act2.fields("Presio Pag").value,"center"  ' Precio Pagado / valor comercial 
						genera_html "d",act2.fields("ValCom USD").value,"center"  ' Valor comercial USD 
						genera_html "d",act2.fields("Valor Fletes MN").value,"center"  ' VALOR FLETES INTERNACIONAL M.N. 
						genera_html "d","","center"  'VALOR FLETES INTERNACIONAL USD
						genera_html "d",act2.fields("Seguros").value,"center"  ' SEGUROS 
						genera_html "d",act2.fields("Otros inc.").value,"center"  ' OTROS INCREMENTABLES 
						genera_html "d","","center"  'OTROS INCREMENTABLES USD
						genera_html "d",act2.fields("Val Adua").value,"center"  ' VALOR ADUANA M.N. 
						genera_html "d",act2.fields("TC").value,"center"  ' T.C. 
						genera_html "d",act2.fields("Val Adu DLS").value,"center"  ' VALOR ADUANA  DLLS 
	
							if( act2.fields("25").value = "MARITIMO") then
								genera_html "d",act2.fields("67").value,"center"  ' VALOR FLETES AEREO DLLS no estan en cosulta
								genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. no se sacan en consulta
								genera_html "d",act2.fields("Fletes").value,"center"  ' VALOR FLETES MARITIMO DLLS. 
							else
								genera_html "d",act2.fields("Fletes").value,"center"  ' VALOR FLETES AEREO DLLS 
								genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
								genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
							end if
				
						genera_html "d",act2.fields("Frac.Ar").value,"center"  ' FRACC. ARANC. 
						genera_html "d",act2.fields("Arancel").value,"center"  'ARANCEL %
						genera_html "d",act2.fields("Aran.Pref").value,"center"  'ARANCEL PREFERENCIAL 
						genera_html "d",retornaECI(act2.fields("Refe").value,tipope,mid(act2.fields("Refe").value,1,3)),"center"  ' MONTO DE RECUPERACION $  
						genera_html "d",act2.fields("761").value,"center"  ' ADV FRACC. $ 
						genera_html "d",act2.fields("Total Imp").value,"center"  ' DTA $ 
						genera_html "d",act2.fields("Iva").value,"center"  'IVA %
						genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
				
					else
				
						genera_html "d",act2.fields("590").value,"center"  ' Precio Pagado / valor comercial 
						genera_html "d",act2.fields("600").value,"center"  ' Valor comercial USD 
						genera_html "d",act2.fields("Valor Fletes MN").value,"center"  ' VALOR FLETES INTERNACIONAL M.N. 
						genera_html "d","","center"  'VALOR FLETES INTERNACIONAL USD
						genera_html "d",act2.fields("Seguros").value,"center"  ' SEGUROS 
						genera_html "d",act2.fields("Otros inc.").value,"center"  ' OTROS INCREMENTABLES 
						genera_html "d","","center"  'OTROS INCREMENTABLES USD
						genera_html "d",act2.fields("640").value,"center"  ' VALOR ADUANA M.N. 
						genera_html "d",act2.fields("TC").value,"center"  ' T.C. 
						genera_html "d",act2.fields("660").value,"center"  ' VALOR ADUANA  DLLS 

							if( act2.fields("25").value = "MARITIMO") then
								genera_html "d",act2.fields("67").value,"center"  ' VALOR FLETES AEREO DLLS 
								genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
								genera_html "d",act2.fields("Fletes").value,"center"  ' VALOR FLETES MARITIMO DLLS. 
							else
								genera_html "d",act2.fields("Fletes").value,"center"  ' VALOR FLETES AEREO DLLS 
								genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
								genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
							end if
				
						genera_html "d",act2.fields("710").value,"center"  ' FRACC. ARANC. 
						genera_html "d",act2.fields("720").value,"center"  'ARANCEL %
						genera_html "d",act2.fields("Aran.Pref").value,"center"  'ARANCEL PREFERENCIAL 
						genera_html "d",retornaECI(act2.fields("Refe").value,tipope,mid(act2.fields("Refe").value,1,3)),"center"  ' MONTO DE RECUPERACION $  
						genera_html "d",act2.fields("7610").value,"center"  ' ADV FRACC. $ 
						genera_html "d",act2.fields("Total Imp").value,"center"  ' DTA $ 
						genera_html "d",act2.fields("770").value,"center"  'IVA %
						genera_html "d",act2.fields("7810").value,"center"  ' IVA FRACC. $ 
				
				end if 'Termninan G1 para lote 2
			else 'No esta ocupado el orden
				genera_html "d","","center"  ' Precio Pagado / valor comercial 
				genera_html "d","","center"  '  Valor comercial USD 
				genera_html "d",act2.fields("Valor Fletes MN").value,"center"  ' VALOR FLETES INTERNACIONAL M.N. 
				genera_html "d","","center"  'VALOR FLETES INTERNACIONAL USD
				genera_html "d",act2.fields("Seguros").value,"center"  ' SEGUROS 
				genera_html "d",act2.fields("Otros inc.").value,"center"  ' OTROS INCREMENTABLES 
				genera_html "d","","center"  'OTROS INCREMENTABLES USD
				genera_html "d","","center"  ' VALOR ADUANA M.N. 
				genera_html "d",act2.fields("TC").value,"center"  ' T.C. 
				genera_html "d","","center"  ' VALOR ADUANA  DLLS 
				
				if( act2.fields("25").value = "MARITIMO") then
					genera_html "d",act2.fields("67").value,"center"  ' VALOR FLETES AEREO DLLS 
					genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
					genera_html "d",act2.fields("Fletes").value,"center"  ' VALOR FLETES MARITIMO DLLS. 
				else
					genera_html "d",act2.fields("Fletes").value,"center"  ' VALOR FLETES AEREO DLLS 
					genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
					genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
				end if
				
				if (act2.fields("cveped").value<>"G1") then
						genera_html "d",act2.fields("Frac.Ar").value,"center"  ' FRACC. ARANC. 
						genera_html "d",act2.fields("Arancel").value,"center"  'ARANCEL %
						genera_html "d",act2.fields("Aran.Pref").value,"center"  'ARANCEL PREFERENCIAL 
				else
						genera_html "d",act2.fields("710").value,"center"  ' FRACC. ARANC. 
						genera_html "d",act2.fields("720").value,"center"  'ARANCEL %
						genera_html "d",act2.fields("Aran.Pref").value,"center"  'ARANCEL PREFERENCIAL 
				end if
				genera_html "d",retornaECI(act2.fields("Refe").value,tipope,mid(act2.fields("Refe").value,1,3)),"center"  ' MONTO DE RECUPERACION $
				genera_html "d","","center"  ' ADV FRACC. $ 
				genera_html "d",act2.fields("Total Imp").value,"center"  ' DTA $ 
				genera_html "d","","center"  'IVA %
				genera_html "d","","center"  ' IVA FRACC. $ 
			end if 'Terminan Orden ocupado en lote 2
			'/Lote2
			'Lote 1
		genera_html "d",act2.fields("Preval").value,"center"  ' PREVAL. 
		genera_html "d",sumaTotalImpuestos(act2.fields("Refe").value,mid(act2.fields("Refe").value,1,3)),"center"  ' TOTAL IMPUESTOS 
		genera_html "d",cDbl(sumaTotalImpuestos(act2.fields("Refe").value,mid(act2.fields("Refe").value,1,3)))/  (act2.fields("TC").value),"center"  ' Total Impuestos USD 
		'-------------------EXTRA COSTOS-------------
		genera_html "d","N/A","center"  ' GASTOS ADUANA USD
		genera_html "d",retornaPagosHechos(act2.fields("Refe").value,retornaConceptosPH(mid(act2.fields("Refe").value,1,3),"DEMORAS"),tipope,mid(act2.fields("Refe").value,1,3)),"center"  ' DEMORAS 
		genera_html "d",retornaPagosHechos(act2.fields("Refe").value,retornaConceptosPH(mid(act2.fields("Refe").value,1,3),"ESTADIAS"),tipope,mid(act2.fields("Refe").value,1,3)),"center"  ' ESTADIAS 
		if ucase(mid(act2.fields("Refe").value,1,3)) ="RKU" or ucase(mid(act2.fields("Refe").value,1,3)) ="SAP" then
			genera_html "d",retornaPagosHechos(act2.fields("Refe").value,retornaConceptosPH(mid(act2.fields("Refe").value,1,3),"MANIOBRAS"),tipope,mid(act2.fields("Refe").value,1,3)),"center"  ' MANIOBRAS  
			genera_html "d",retornaPagosHechos(act2.fields("Refe").value,retornaConceptosPH(mid(act2.fields("Refe").value,1,3),"ALMACENAJES-MANIOBRAS"),tipope,mid(act2.fields("Refe").value,1,3)),"center"  ' ALMACENAJES 
		else
			genera_html "d",retornaPagosHechos(act2.fields("Refe").value,retornaConceptosPH(mid(act2.fields("Refe").value,1,3),"ALMACENAJES-MANIOBRAS"),tipope,mid(act2.fields("Refe").value,1,3)),"center"  ' MANIOBRAS
			genera_html "d","?","center"  ' ALMACENAJES 
		end if
				
		dim TPH,DEM,EST,ALMMAN,MAN
			TPH=0
			DEM=0
			EST=0
			ALMAN=0
			MAN=0
			STIMP=0
			STIVA=0
			TPH=retornaTOTALPagosHechos(act2.fields("Refe").value,tipope,mid(act2.fields("Refe").value,1,3))
			DEM= retornaPagosHechos(act2.fields("Refe").value,retornaConceptosPH(mid(act2.fields("Refe").value,1,3),"DEMORAS"),tipope,mid(act2.fields("Refe").value,1,3))
				EST = retornaPagosHechos(act2.fields("Refe").value,retornaConceptosPH(mid(act2.fields("Refe").value,1,3),"ESTADIAS"),tipope,mid(act2.fields("Refe").value,1,3))	
			ALMAN=retornaPagosHechos(act2.fields("Refe").value,retornaConceptosPH(mid(act2.fields("Refe").value,1,3),"ALMACENAJES-MANIOBRAS"),tipope,mid(act2.fields("Refe").value,1,3))
		if ucase(mid(act2.fields("Refe").value,1,3)) ="RKU" or ucase(mid(act2.fields("Refe").value,1,3)) ="SAP" then
				MAN=retornaPagosHechos(act2.fields("Refe").value,retornaConceptosPH(mid(act2.fields("Refe").value,1,3),"MANIOBRAS"),tipope,mid(act2.fields("Refe").value,1,3))
		end if
		if(revisaImpuestosFacturados( act2.fields("Refe").value,tipope,mid(act2.fields("Refe").value,1,3)) <> 0 )then
				STIMP=sumaTotalImpuestos(act2.fields("Refe").value,mid(act2.fields("Refe").value,1,3))
				STIVA=sumaTotalIVA(act2.fields("Refe").value,mid(act2.fields("Refe").value,1,3))
		end if
		on error resume next
		dim cueta
		cuenta=retornaCampoCtaGastos(act2.fields("Refe").value,"cgas31",mid(act2.fields("Refe").value,1,3))
			if cuenta<>"" then
				genera_html "d",TPH-DEM-EST-ALMAN-MAN-STIMP-STIVA,"center"  ' OTROS 
				genera_html "d",retornaTOTALPagosHechos(act2.fields("Refe").value,tipope,mid(act2.fields("Refe").value,1,3))-STIMP-STIVA,"center"  ' TOTAL GASTOS DIVERSOS 
			else
				genera_html "d","0","center"  ' OTROS 
				genera_html "d","0","center"  ' TOTAL GASTOS DIVERSOS 
			end if
			
			genera_html "d","","center" ' Pidieron campo que se dejara en blanco
				'genera_html "d",((retornaTOTALPagosHechos(act2.fields("Refe").value,"I",mid(act2.fields("Refe").value,1,3))-STIMP-STIVA)/act2.fields("TC").value),"center"  ' TOTAL GASTOS DIVERSOS USD 						
			if err.number <> 0 then
				genera_html "d","","center"  ' OTROS 
				genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS 
				genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS USD
			end if
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS M.N.
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS USD
			genera_html "d",retornaHonorarios(act2.fields("Refe").value,"chon31",mid(act2.fields("Refe").value,1,3)),"center"  ' HONORARIOS AG AD. $ 
			genera_html "d",act2.fields("Naviera").value,"center" 'NAVIERA
			genera_html "d",retornaTransportista(mid(act2.fields("Refe").value,1,3),act2.fields("Refe").value),"center"  ' TRANSPORTISTA NACIONAL
			genera_html "d",retornaFletesNacionales(mid(act2.fields("Refe").value,1,3),act2.fields("Refe").value),"center"  ' COSTO FLETE NACIONAL
			genera_html "d",act2.fields("sDescTransp").value,"center"  'TIPO TRANSPORTE
			genera_html "d",retornaTipoUnidad(act2.fields("Refe").value,tipope,mid(act2.fields("Refe").value,1,3)),"center" 'TIPO UNIDAD
			'/Lote 1
		else 
			'Lote 2
			genera_html "d","","center"  'No de Contenedor 
			genera_html "d",act2.fields("cveped").value,"center" 'Cve Pedimento
			genera_html "d",act2.fields("NumPedimento").value,"center"  'No. Pedimento
			genera_html "d",act2.fields("FechaPagoPed").value,"center"  'Fecha Pedimento
			genera_html "d",act2.fields("Mes").value,"center"  'Mes
			genera_html "d",act2.fields("Sem").value,"center"  'No.Semana
			genera_html "d","0","center"  'Cantidad de Operaciones 
			genera_html "d","","center"  'Cantidad de Contenedores
			genera_html "d","","center"  'PALLETS/BULTOS 
			genera_html "d","","center"  'Medida del contenedor
			genera_html "d","","center"  'TIPO DE CONTENEDOR/ CAJA
			if (act2.fields("cveped").value<>"G1") then
				genera_html "d",act2.fields("FechaFactura").value,"center"  'Fecha Factura
			else
				genera_html "d",act2.fields("ffac05").value,"center"  'Fecha Factura G1
			end if
			
			genera_html "d",act2.fields("F.BL").value,"center"  'Fecha BL
			genera_html "d",act2.fields("F. ArriboAduana").value,"center"  'Fecha de arribo a la aduana
			genera_html "d",act2.fields("F.Desaduanamiento").value,"center"  'Fecha Desaduanamiento
			genera_html "d","","center"  'KPI Programación
			genera_html "d","","center"  'KPI Tránsito
			genera_html "d","","center"  'KPI Desaduanamiento
			genera_html "d","","center"  'KPI lead TIME
			genera_html "d","","center"  'TARGET TIME
			genera_html "d","","center"  'NUMERO DE EMBARQUE
			genera_html "d","","center"  'IDOT
			genera_html "d","","center"  'IDOT LT
			genera_html "d","","center" 'CAUSA IDOT
			genera_html "d","","center"  'CAUSA RAIZ
			genera_html "d","","center"  'PLANTA DE ACCION
			genera_html "d","","center"  'RESPONSABLE
			genera_html "d","","center"  'FECHA DE CUMPLIMIENTO
			genera_html "d","","center"  'IMPACTO
			genera_html "d",retornaCampoCtaGastos(act2.fields("Refe").value,"cgas31",mid(act2.fields("Refe").value,1,3)),"center"  'No. CTA DE GASTOS
			genera_html "d",regresa_fecha_cuenta_gastos(act2.fields("Refe").value,mid(act2.fields("Refe").value,1,3)),"center"      'Fecha Cta de Gastos
			genera_html "d","","center"    'Tipo C.Gastos
	
			Subref = act2.fields("Ord FrcA").value
			if (ordenOcupado(Subref,act2.fields("Refe").value) = False)then
				ocuparOrd(Subref)
				if (act2.fields("cveped").value<>"G1") then
					genera_html "d",act2.fields("Presio Pag").value,"center"  ' Precio Pagado / valor comercial 
					genera_html "d",act2.fields("ValCom USD").value,"center"  '  Valor comercial USD 
					genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL M.N. 
					genera_html "d","","center"  'VALOR FLETES INTERNACIONAL USD
					genera_html "d","","center"  ' SEGUROS 
					genera_html "d","","center"  ' OTROS INCREMENTABLES 
					genera_html "d","","center"  'OTROS INCREMENTABLES USD
					genera_html "d",act2.fields("Val Adua").value,"center"  ' VALOR ADUANA M.N. 
					genera_html "d",act2.fields("TC").value,"center"  ' T.C. 
					genera_html "d",act2.fields("Val Adu DLS").value,"center"  ' VALOR ADUANA  DLLS 
					genera_html "d","","center"  ' VALOR FLETES AEREO DLLS 
					genera_html "d","","center"  ' VALOR FLETES TERRESTRE DLLS. 
					genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
					genera_html "d",act2.fields("Frac.Ar").value,"center"  ' FRACC. ARANC. 
					genera_html "d",act2.fields("Arancel").value,"center"  'ARANCEL %
					genera_html "d",act2.fields("Aran.Pref").value,"center"  'ARANCEL PREFERENCIAL 
					genera_html "d","","center"  ' MONTO DE RECUPERACION $  
					genera_html "d",act2.fields("761").value,"center"  ' ADV FRACC. $ 
					genera_html "d","","center"  ' DTA $ 
					genera_html "d",act2.fields("Iva").value,"center"  'IVA %
					genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
				
				else
				
					genera_html "d",act2.fields("590").value,"center"  ' Precio Pagado / valor comercial 
					genera_html "d",act2.fields("600").value,"center"  '  Valor comercial USD 
					genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL M.N. 
					genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL USD
					genera_html "d","","center"  ' SEGUROS 
					genera_html "d","","center"  ' OTROS INCREMENTABLES 
					genera_html "d","","center"  ' OTROS INCREMENTABLES USD
					genera_html "d",act2.fields("640").value,"center"  ' VALOR ADUANA M.N. 
					genera_html "d",act2.fields("TC").value,"center"  ' T.C. 
					genera_html "d",act2.fields("660").value,"center"  ' VALOR ADUANA  DLLS 
					genera_html "d","","center"  ' VALOR FLETES AEREO DLLS 
					genera_html "d","","center"  ' VALOR FLETES TERRESTRE DLLS. 
					genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
					genera_html "d",act2.fields("710").value,"center"  ' FRACC. ARANC. 
					genera_html "d",act2.fields("720").value,"center"  'ARANCEL %
					genera_html "d",act2.fields("Aran.Pref").value,"center"  'ARANCEL PREFERENCIAL 
					genera_html "d","","center"  ' MONTO DE RECUPERACION $  
					genera_html "d",act2.fields("7610").value,"center"  ' ADV FRACC. $ 
					genera_html "d","","center"  ' DTA $ 
					genera_html "d",act2.fields("770").value,"center"  'IVA %
					genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
				
				end if 'Terminan G1
			else  '--------- Si ordenOcupado(Subref,act2.fields("Refe").value) = False
			
			  	genera_html "d","","center"  ' Precio Pagado / valor comercial 
				genera_html "d","","center"  '  Valor comercial USD 
				genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL M.N. 
				genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL USD
				genera_html "d","","center"  ' SEGUROS 
				genera_html "d","","center"  ' OTROS INCREMENTABLES 
				genera_html "d","","center"  ' OTROS INCREMENTABLES USD
				genera_html "d","","center"  ' VALOR ADUANA M.N. 
				genera_html "d",act2.fields("TC").value,"center"  ' T.C. 
				genera_html "d","","center"  ' VALOR ADUANA  DLLS 
				genera_html "d","","center"  ' VALOR FLETES AEREO DLLS 
				genera_html "d","","center"  ' VALOR FLETES TERRESTRE DLLS. 
				genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
				if (act2.fields("cveped").value<>"G1") then
					genera_html "d",act2.fields("Frac.Ar").value,"center"  ' FRACC. ARANC. 
					genera_html "d",act2.fields("Arancel").value,"center"  'ARANCEL %
				else
					genera_html "d",act2.fields("710").value,"center"  ' FRACC. ARANC. 
					genera_html "d",act2.fields("720").value,"center"  'ARANCEL %
				end if
				
				genera_html "d",act2.fields("Aran.Pref").value,"center"  'ARANCEL PREFERENCIAL 
				genera_html "d","","center"  ' MONTO DE RECUPERACION $  
			   	genera_html "d","","center"  ' ADV FRACC. $ 
			    genera_html "d","","center"  ' DTA $ 
				genera_html "d","","center"  'IVA %
				genera_html "d","","center"  ' IVA FRACC. $ 
			end if 
			'/Lote2
		 'Lote 1
			genera_html "d","","center"  ' PREVAL. 
			genera_html "d","","center"  ' TOTAL IMPUESTOS 
			genera_html "d","","center"  ' Total Impuestos USD  
			genera_html "d","N/A","center"  ' GTOS. ADUANA USD(SOLO FRONTERA) 
			genera_html "d","","center"  ' DEMORAS 
			genera_html "d","","center"  ' ESTADIAS 
			genera_html "d","","center"  ' MANIOBRAS  
			genera_html "d","","center"  ' ALMACENAJES 
			genera_html "d","","center"  ' OTROS 
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS 
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS USD 
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS M.N
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS USD
			genera_html "d","","center"  ' HONORARIOS AG AD. $ 
			genera_html "d",act2.fields("Naviera").value,"center" 'NAVIERA
			genera_html "d",retornaTransportista(mid(act2.fields("Refe").value,1,3),act2.fields("Refe").value),"center"  ' TRANSPORTISTA NACIONAL
			genera_html "d","","center"
			genera_html "d",act2.fields("sDescTransp").value,"center"  'TIPO TRANSPORTE
			genera_html "d",retornaTipoUnidad(act2.fields("Refe").value,tipope,mid(act2.fields("Refe").value,1,3)),"center" 'TIPO UNIDAD
			'/Lote 1
		end if
		response.Write("</tr>")
		act2.movenext()
	'Termina la impresion de campos
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

function revisaImpuestosFacturados(referencia,tipoop,oficina)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif(ucase(oficina)="PAN")then
 oficina="DAI"
 end if
 
 
sqlAct="select count(i.refcia01) as Ref " & _
" from "& oficina &"_extranet.ssdag" & tipoop &"01 as i  " & _
"  inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  " & _
"     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 " & _
"          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 " & _
"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
"    where  i.firmae01 <> ''  and cta.esta31 <> 'C'  and i.refcia01 = '"& referencia &"' and ep.conc21 = 1"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()

if not(act2.eof) then
 revisaImpuestosFacturados =act2.fields("Ref").value
else
  revisaImpuestosFacturados = nothing
end if


end function


function regresa_fecha_cuenta_gastos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
 
sqlAct="select max(date_format(cta.fech31,'%d/%m/%Y')) as fech31 from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
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
 regresa_fecha_cuenta_gastos =act2.fields("fech31").value
else
  regresa_fecha_cuenta_gastos =valor
   end if
end function


function regresa_tipo_Cgastos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
 
sqlAct="select if(COUNT(cta.cgas31) > 1, 'COMPLEMENTARIA','NORMAL')  as tipo from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
" r.refe31 = '"&referencia&"'  "&_
"  and cta.esta31 <> 'C' "


Set act2= Server.CreateObject("ADODB.Recordset")
conn12= "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
if not(act2.eof) then
 regresa_tipo_Cgastos =act2.fields("tipo").value
else
  regresa_tipo_Cgastos =valor
   end if
end function


function codigoProveedor(desc)
dim res,desc2
res = "no"
Path=Server.MapPath("catprobd.xls")
desc2=replace(desc," ","%")
Set ConexionBD = Server.CreateObject("ADODB.Connection") 
ConexionBD.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & Path
Set rsVac = Server.CreateObject("ADODB.Recordset") 
rsVac.Open "Select * From A1:B50 where descpro like '" & desc2 & "'", ConexionBD,3,3 

if not(rsVac.eof)then
 res =rsVac.fields("cvepro") 
else
 res = desc
end if 

codigoProveedor = res
end function

function retornaCampoCtaGastos(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
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


function retornaMontoAnticipo(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
 
 sqlAct = " select sum(dm.mont11) as campo " & _
			" from "&oficina&"_extranet.d11movim as dm " & _
			" where dm.refe11='"& referencia&"' and dm.conc11 = '"&campo&"' "
Set act2= Server.CreateObject("ADODB.Recordset")
conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

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
  retornaMontoAnticipo = valor
 else
  retornaMontoAnticipo =valor
 end if
end function


function retornaIMPORTADOR(clave,oficina)
ON ERROR RESUME NEXT
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
  sqlAct2 = "select c.nomcli18 as campo from "&oficina&"_extranet.ssclie18 as c where c.cvecli18 = "&clave
Set act2= Server.CreateObject("ADODB.Recordset")
conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct2
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()

if err.number <> 0 then
	retornaIMPORTADOR = err.description
else

 if not(act2.eof) then
 
 
 valor = act2.fields("campo").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaIMPORTADOR = valor
 else
  retornaIMPORTADOR =valor
 end if
end if   
 
end function
function retornaCampoPuertoEmb(pto,val,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
 sqlAct="SELECT Distinct "& campo &" as campo FROM "&oficina&"_extranet.c01ptoemb where cvepto01 ="& val &" and nompto01 like '"& pto &"%'"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

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
  retornaCampoPuertoEmb = valor
 else
  retornaCampoPuertoEmb =valor
 end if
end function

function retornaAgenteAduanal(valor)
dim val
val = ""
if(valor = "3921")then
 val = "Luis E. de la Cruz Reyes"
else
if(valor = "3210")then
 val = "Rolando Reyes Kuri"
else
if(valor = "3945")then
 val = "Jesús Gómez Reyes"
else
if(valor = "3931")then
 val = "Sergio Alvarez Ramírez"
else
if(valor = "3044")then
 val = "Carlos Humberto Zesati Andrade"
else
val =""
end if
end if
end if
end if
end if

retornaAgenteAduanal = val
end function


function retornaAduana(valor)
dim val
val = ""
if(valor = "470")then
 val = "México"
else
if(valor = "430")then
 val = "Veracruz"
else
if(valor = "810")then
 val = "Altamira"
else
if(valor = "160")then
 val = "Manzanillo"
else
if(valor = "510")then
 val = "Lázaro Cardenas"
else
if(valor = "650")then
 val = "Toluca"
else
val =""
end if
end if
end if
end if
end if
end if

retornaAduana = val
end function




function codigoCliente(desc)
dim res,desc2
res = "no"
Path=Server.MapPath("catclibd.xls")
desc2=replace(desc," ","%")
Set ConexionBD = Server.CreateObject("ADODB.Connection") 
ConexionBD.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & Path
Set rsVac = Server.CreateObject("ADODB.Recordset") 
rsVac.Open "Select * From A1:B35 where desccli like '" & desc2 & "'", ConexionBD,3,3 
if not(rsVac.eof)then
 res =rsVac.fields("cvecli") 

else
 res = desc
end if 

codigoCliente = res
end function


function retornaHonorarios(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 

sqlAct=" select cta."&campo&" as campo from "&oficina&"_extranet.e31cgast as cta  " & _
       " inner join "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
       " where  r.refe31 = '"& referencia &"' and cta.esta31 = 'I' "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

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
  retornaHonorarios = valor
 else
  retornaHonorarios =valor
 end if
end function

function retornaConceptosPH(oficina,topico)
dim cad
cad = "NA"

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 

if oficina = "SAP" then

if topico = "ALMACENAJES-MANIOBRAS" then
  cad= "2,3,63,65,86,111,111,112,142,174,181,183,183,186,188,189,190,196,196,208,209,210,211,212,214,216,218,234,251,256,258,265,269,284,286,287,288,289,290,291,292,293,294,295,296,297,298,299,300,301,303,304,305,307,309,323,331,336,351"
end if
if topico = "DEMORAS" then
  cad= "6,14,46,63,129,156,352"
end if
if topico = "ESTADIAS" then
  cad="144"
end if
if topico = "OTROS" then
  cad="306,313,350,351,352"
end if


else 
  if oficina = "CEG" then
  
    if topico = "ALMACENAJES-MANIOBRAS" then
     cad= "2,4,59,77,100,223,223,235,235,241"
    end if
	if topico = "DEMORAS" then
	 cad= "11,48,99,150"
	end if
	if topico = "ESTADIAS" then
	 cad="NA"
	end if
	if topico = "OTROS" then
	 cad="239"
    end if
  
  else 
     if oficina = "TOL" then
	 
	    if topico = "ALMACENAJES-MANIOBRAS" then
		 cad= "2,2,10,127,128"
		end if
		if topico = "DEMORAS" then
		 cad= "79"
		end if
		if topico = "ESTADIAS" then
		 cad="123"
		end if
		if topico = "OTROS" then
		 cad="NA"
		end if
		
     else 
	   if oficina = "LZR" then
	         if topico = "ALMACENAJES-MANIOBRAS" then
			 cad= "4,78,115,116,119,125,160,167,167,203,230,230,244,297,312"
			end if
			if topico = "DEMORAS" then
			 cad= "11"
			end if
			if topico = "ESTADIAS" then
			 cad="77"
			end if
			if topico = "OTROS" then
			 cad="NA"
			end if
       else 
	       if oficina = "RKU" then
		          if topico = "ALMACENAJES-MANIOBRAS" then
					 'cad= "2,4,78,115,116,119,125,160,167,167,203,230,230,244,297,304,311,312,313,359,359"
					 cad="4"
					end if
					if topico = "DEMORAS" then
					 cad= "11,310,376"
					end if
					if topico = "ESTADIAS" then
					 cad="77"
					end if
					if topico = "MANIOBRAS" then
					 cad="2"
					end if
           else 
		       if oficina = "DAI" then
			           if topico = "ALMACENAJES-MANIOBRAS" then
						 cad= "2,2,10,93,127,128,155,163,163,166,170"
						end if
						if topico = "DEMORAS" then
						 cad= "79,171"
						end if
						if topico = "ESTADIAS" then
						 cad="123"
						end if
						if topico = "OTROS" then
						 cad="NA"
						end if
               else 
   			          cad = "NE"
               end if
           end if
       end if
     end if
  end if
end if
retornaConceptosPH = cad
end function


function retornaECI(referencia,tipope,oficina)
dim c,valor
 c=chr(34)
 valor=0

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 


sqlAct =" select c.import36 as Campo ,c.cveimp36,c.refcia36  from "& oficina &"_extranet.ssdag"& tipope &"01 as i " & _
		"  inner  join  "& oficina &"_extranet.sscont36 as c on i.refcia01 = c.refcia36 " & _
		"    where c.refcia36 = '"& referencia &"' and c.cveimp36 = '18' and i.rfccli01 in ('SEM950215S98')"


Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
	 if not(act2.eof) then
	 valor = act2.fields("Campo").value
	
	  retornaECI = valor
	 else
	  retornaECI = valor
	 end if

end function


function retornaPagosHechos(referencia,conceptos,tipope,oficina)
dim c,valor
 c=chr(34)
 valor=0

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then

sqlAct="select i.refcia01 as Ref, r.cgas31,ep.conc21,ep.piva21,sum(dp.mont21*if(ep.deha21 = 'C',-1,1)) as Importe, cp.desc21 " & _
" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
"  inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  " & _
"     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 " & _
"          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 " & _
"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
"    where  i.rfccli01 in ('SEM950215S98')  and i.firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in ("&conceptos&") and i.refcia01 = '"&referencia&"'  group by Ref,cgas31,conc21"

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("Importe").value
  retornaPagosHechos = valor
 else
  retornaPagosHechos = valor
 end if
 else
   retornaPagosHechos =valor
 end if

end function


function retornaTOTALPagosHechos(referencia,tipope,oficina)
dim c,valor
 c=chr(34)
 valor=0

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then

sqlAct =" select i.refcia01 as Ref,sum(dp.mont21*if(ep.deha21 = 'C',-1,1)) as Importe " & _
		" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
		" inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 " & _
		" inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
		" where i.rfccli01 in ('SEM950215S98')  and i.refcia01 = '"&referencia&"'  and i.firmae01 <> ''  group by i.refcia01 "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
 valor = act2.fields("Importe").value
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("Importe").value
   act2.movenext()
 wend
  retornaTOTALPagosHechos = valor
 else
  retornaTOTALPagosHechos = valor
 end if
 else
   retornaTOTALPagosHechos =0
 end if

end function

function retornaCampoContenedores(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
 
sqlAct="select Distinct "& campo &" as campo from "&oficina&"_extranet.sscont40 where refcia40 = '"&referencia&"'   "

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
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCampoContenedores = valor
 else
  retornaCampoContenedores =valor
 end if
end function


function retornaCantContenedores(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
sqlAct="select Distinct count(*) as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"' and clas01 in ("&campo&") "

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
 act2.movenext()
 while not act2.eof
   valor = valor&", "&act2.fields("campo").value
   act2.movenext()
 wend
  retornaCantContenedores = valor
 else
  retornaCantContenedores =valor
 end if
end function

function retornaCantContenedores40(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
sqlAct="select count(Distinct numcon40) as campo from  "&oficina&"_extranet.sscont40 where refcia40 = '"&referencia&"'  "

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
  retornaCantContenedores40 = valor
 else
  retornaCantContenedores40 =valor
 end if
end function

function DatosContenedor(Tipo)
dim val
val = Tipo
if(Tipo = "1")then
  'val="Contenedor Estandar 40 pulg (Standard Container 40 pulg)"
   val="20"
end if

if(Tipo = "2")then
'  val="Contenedor Estandar 40 pulg (Standard Container 40 pulg)"
    val="40"
end if

if(Tipo = "3")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="40 HighCube"
end if

if(Tipo = "4")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="20 Hardtop"
end if

if(Tipo = "5")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="40 Hardtop"
end if


if(Tipo = "6")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="20 OpenTop"
end if

if(Tipo = "7")then
'  val="Contenedor Estandar de cubo alto 40 pulg (High Cube Standard Container 40 pulg) "
  val="40 OpenTop"
end if


if(Tipo = "17")then
'val="Contenedor Refrigerante Cubo Alto 17 pulg (High Cube Refrigerated Container 40 pulg"
val="17 HighCube"
end if
DatosContenedor = val
end function

function retornaTipoContenedores(referencia,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
 
sqlAct="select distinct concat_ws(' ',cn7.dimen07,cn7.catego07) As campo from "&oficina&"_extranet.sscont40 as cn4 LEFT JOIN  "&oficina&"_extranet.sstcon07 cn7 on cn4.tipcon40 =cn7.cvecon07 where cn4.refcia40 = '"& referencia &"' "

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
	act2.movenext()
		while not act2.eof
			valor = valor &", "& act2.fields("campo").value
			act2.movenext()
		wend
	retornaTipoContenedores = valor
 else
	retornaTipoContenedores =valor
 end if
end function


function TipoCaja(referencia,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
 
sqlAct="select distinct cn4.tipcon40 As campo from "&oficina&"_extranet.sscont40 as cn4 LEFT JOIN  "&oficina&"_extranet.sstcon07 cn7 on cn4.tipcon40 =cn7.cvecon07 where cn4.refcia40 = '"& referencia &"' "

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
	act2.movenext()
		while not act2.eof
			valor = valor &", "& act2.fields("campo").value
			act2.movenext()
		wend
		select case valor
			case 1,2,3
			valor="SECA"
			case 17
			valor="REFRIGERADA"
		
		end select
		
	TipoCaja = valor
 else
 select case valor
			case 1,2,3
			valor="SECA"
			case 17
			valor="REFRIGERADA"
		
		end select	
	TipoCaja =valor
 end if
end function





function revisaFraccion(fraccion)
dim val
val ="(pendiente)"
fraccion = trim(fraccion)
'nueva fraccion: 33019001 añadida el 20/01/2011
if (fraccion = "90230001" or _
	fraccion = "49111099" or _
	fraccion = "48219099" or _
	fraccion = "39235001" or _
	fraccion = "34011101" or _
	fraccion = "33072001" or _
	fraccion = "33019001" or _
	fraccion = "33059099" or _
	fraccion = "33051001" or _
	fraccion = "33049999" or _
	fraccion = "76129099" or _
	fraccion = "33079099" or _
    fraccion = "39202099" or _
	fraccion = "85234099" or _
	fraccion = "49119999" or _
	fraccion = "09081001" or _
	fraccion = "13023902" or _
	fraccion = "21069003" or _
	fraccion = "39209999" or _
	fraccion = "39233099" or _
	fraccion = "84212199" or _
	fraccion = "33071001" or _
	fraccion = "39231001" or _
	fraccion = "49019906" or _	 
	fraccion = "33029099") then
	val ="HPC"
else
	if (fraccion = "21039099" or _
	fraccion = "12119001" or _
 	fraccion = "9023001" or _
	fraccion = "07103001" or _
	fraccion = "07108099" or _
	fraccion = "18069099" or _
 	fraccion = "21069099" or _
	fraccion = "21041001" or _
 	fraccion = "07129099" or _
 	fraccion = "11061001" or _
	fraccion = "15119099" or _
	fraccion = "17029099" or _
	fraccion = "39219099" or _
	fraccion = "09023001") then
	val ="FOODS"
	else
	val = "ERROR (CD) "&fraccion
	end if
end if

revisaFraccion = val
end function

function retornaDivision(clave,fraccion)
dim val,res
val= ""
res= "(pendiente)"
response.write("Estoy en division :D")

if clave = "11000" then
  val = revisaFraccion(fraccion) '"Centro de Distribución"
end if
if clave = "11001" then
val = "ICE CREAM" '"Planta Helados"
end if
if clave = "11002" then
val = "FOODS"
end if
if clave = "11003" then
val = "HPC" '"Planta HPC"
end if
if clave = "11004" then
val = "Todas"
end if
'OTRAS
if clave = "13000" then
val = "Todas"
end if
if clave = "14000" then
val = "Todas"
end if


retornaDivision = val
end function


function retornaPlantaEntrega(clave)
dim val
val= ""
if clave = "11000" then
val = "CDU"
end if
if clave = "11001" then
val = "TULTITLAN"
end if
if clave = "11002" then
val = "LERMA"
end if
if clave = "11003" then
val = "CIVAC"
end if
if clave = "11004" then
val = "ESPECIALES"
end if

'OTRAS
if clave = "13000" then
val = "Todas"
end if
if clave = "14000" then
val = "Todas"
end if


retornaPlantaEntrega = val
end function


function sumaTotalImpuestos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor="0"
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
 
sqlAct=" select ifnull(sum(import36),0) as campo from "& oficina &"_extranet.sscont36 as cf1 " & _
       " where cf1.cveimp36 in ('1', '6','15')   and refcia36 = '"&referencia&"' "
    

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
 act2.movenext()
 while not act2.eof
   valor = valor &", "& act2.fields("campo").value
   act2.movenext()
 wend
  sumaTotalImpuestos = valor
 else
  sumaTotalImpuestos =valor
 end if

'ADV/IGI+DTA+IVA+PREVAL

end function
function sumaTotalIVA(referencia,oficina)
dim c,valor
 c=chr(34)
 valor=0
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if
 
 
sqlAct=" select ifnull(sum(import36),0) as campo from "& oficina &"_extranet.sscont36 as cf1 " & _
       " where cf1.cveimp36 in ('3')   and refcia36 = '"&referencia&"' "
    
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
  sumaTotalIVA = valor
 else
  sumaTotalIVA =valor
 end if

'ADV/IGI+DTA+IVA+PREVAL

end function


function retornaRegion(clave)
dim val
val= "("& clave &")"

if clave = "MEX" then
val = "LATINOAMERICA"
end if

if clave = "TUR" then
val = "ASIA"
end if

if clave = "ARG" then
val = "SUDAMERICA"
end if
if clave = "BRA" then
val = "SUDAMERICA"
end if
if clave = "CHL" then
val = "SUDAMERICA"
end if
if clave = "COL" then
val = "SUDAMERICA"
end if
if clave = "DEU" then
val = "EUROPA"
end if


if clave = "DOM" then
val = "ANTILLAS"
end if
if clave = "ESP" then
val = "EUROPA"
end if
if clave = "FRA" then
val = "EUROPA"
end if
if clave = "GBR" then
val = "EUROPA"
end if
if clave = "SLV" then
val = "CENTROAMERICA"
end if

if clave = "THA" then
val = "ASIA"
end if
if clave = "USA" then
val = "NORTE AMERICA"
end if
if clave = "ZYA" then
val = "EUROPA"
end if

if clave = "CAN" then
val = "NORTE AMERICA"
end if

if clave = "CHN" then
val = "ASIA"
end if

if clave = "ITA" then
val = "EUROPA"
end if

if clave = "DNK" then
val = "EUROPA"
end if

if clave = "PHL" then
val = "ASIA"
end if

if clave = "BEL" then
val = "EUROPA"
end if

if clave = "BOL" then
val = "SUDAMERICA"
end if

if clave = "IND" then
val = "ASIA"
end if

if clave = "JPN" then
val = "ASIA"
end if



if clave = "PER" then
val = "SUDAMERICA"
end if

if clave = "VEN" then
val = "SUDAMERICA"
end if

if clave = "CRI" then
val = "LATINOAMERICA"
end if

if clave = "AUT" then
val = "ASIA"
end if


if clave = "SGP" then
val = "ASIA"
end if

if clave = "NZL" then
val = "EUROPA"
end if

if clave = "CHE" then
val = "EUROPA"
end if

if clave = "MYS" then
val = "ASIA"
end if



retornaRegion = val
end function


function ordenOcupado(Subref,referencia)
dim res
res = False


if(referencia <> subrefAux)then
subrefaux=referencia
  orden(1)=""
  orden(2)=""
  orden(3)=""
   orden(4)=""
	orden(5)=""
	 orden(6)=""
	  orden(7)=""
	   orden(8)=""
		orden(9)=""
		orden(10)=""
		orden(11)=""
	 orden(12)=""
	  orden(13)=""
	   orden(14)=""
	    orden(15)=""
		 orden(16)=""
		  orden(17)=""
end if

if(subref <> "" ) then
'for i=0 to 50 
if orden(Subref) = "1" then
 res = True
else
 res = False
end if
'next
end if


ordenOcupado = res
end function

function ocuparOrd(Subref)
dim res
res = False

if(subref <> "" ) then

orden(Subref) ="1"
'redim orden
end if

ocuparOrd = res
end function

function retornaCECO(rcli)
	dim val,aux,rcli_aux
	val =""
	rcli_aux = trim(rcli)
	
	if InStr(rcli_aux,"CECO")>0 then
		'primero eliminamos los caracteres extra y separamos el array de acuerdo a los espacios en blanco
		rcli_aux = Cstr(rcli_aux)
		rcli_aux = Replace(rcli_aux,vbCrLf," ")
		aux = split(rcli_aux)
		For Each item In aux
			if InStr(item,"CECO")>0 then 
				val = item
			end if
		next
		'una ultima validacion por si las dudas
		if InStr(val,"CECO") = 0 then
			val = "N/E"
		end if
	else
		val = "N/E"
	end if

	retornaCECO = val
end function

function retornaCuenta(rcli)
	dim val,aux,rcli_aux
	val = ""
	rcli_aux = trim(rcli)
	
	if InStr(rcli_aux,"CUENTA")>0 then
		'primero eliminamos los caracteres extra y separamos el array de acuerdo a los espacios en blanco
		rcli_aux = Cstr(rcli_aux)
		rcli_aux = Replace(rcli_aux,vbCrLf," ")
		aux = split(rcli_aux)
		For Each item In aux
			if InStr(item,"CUENTA")>0 then 
				val = item
			end if
		next
		if InStr(val,"CUENTA") = 0 then
			val = "N/E"
		end if
	else
		val = "N/E"
	end if

	retornaCuenta = val
end function


function retornaTipoUnidad(referencia,tipope,oficina)
dim valor,unidad
 valor=""
 unidad=""
 
if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
 end if

sqlAct=" SELECT tracto01 as campo from "& oficina &"_extranet.ssdag"&tipope&"01 AS i LEFT JOIN "& oficina &"_extranet.sscont40 " & _
		"AS con ON con.refcia40 =i.refcia01 AND con.patent40 = i.patent01 AND con.adusec40 = i.adusec01 LEFT JOIN" & _
		" "& oficina &"_extranet.d01conte AS d01 ON d01.refe01 =i.refcia01 AND con.numcon40 = REPLACE(REPLACE(d01.marc01, '/', ''), '-', '')" & _
		" where  i.refcia01 = '"& referencia &"'"
		
Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427" 
act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
  while not act2.eof
      valor = act2.fields("campo").value
   if valor= "S" then 
     unidad = "SENCILLO"
  end if
  if valor ="F" then
 	unidad = "FULL"
  end if 
  act2.movenext()
 wend
  retornaTipoUnidad = unidad
 else
  retornaTipoUnidad = unidad
 end if
end function
function retornaFletesNacionales(oficina,referencia)
if oficina="ALC" then oficina="LZR"
sqlAct="SELECT 	ifnull(SUM(IF(eph.deha21 = 'A', dph.mont21, (dph.mont21)*-1)),0) AS 'FleteNac'	" & _
	"FROM "&oficina&"_extranet.ssdag"&tipope&"01 AS imp LEFT JOIN  "&oficina&"_extranet.d31refer AS ref ON ref.refe31 = imp.refcia01 " & _
	"LEFT JOIN  "&oficina&"_extranet.e31cgast AS cgt ON cgt.cgas31 = ref.cgas31 AND cgt.esta31 <> 'C' "& _
	"INNER JOIN "&oficina&"_extranet.d21paghe AS dph ON dph.refe21 = imp.refcia01 AND dph.cgas21 = cgt.cgas31 " & _
	"INNER JOIN "&oficina&"_extranet.e21paghe AS eph ON eph.foli21 = dph.foli21 AND YEAR(eph.fech21) = YEAR(dph.fech21) AND eph.tmov21 = dph.tmov21 AND eph.esta21 <> 'S'" & _
	"INNER JOIN "&oficina&"_extranet.c21paghe AS cph ON cph.clav21 = eph.conc21 AND cph.desc21 LIKE 'flete%terres%' " & _
	"WHERE imp.rfccli01 IN ('SEM950215S98') AND " & _
	"imp.firmae01 <> '' AND imp.firmae01 IS NOT NULL AND imp.cveped01 <> 'R1' and imp.refcia01 ='"&referencia&"'"

	Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427" 
act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
 if not(act2.eof) then
  while not act2.eof
	monto =act2.fields("FleteNac").value
  
  act2.movenext()
 wend

  retornaFletesNacionales = monto
 else
  retornaFletesNacionales = monto
   end if
end function 

function retornaTransportista(oficina,referencia)
if oficina="ALC" then 
oficina="LZR"
elseif (ucase(oficina) ="PAN") then
 oficina="DAI"
end if

	sqlAct="select group_concat(distinct c.nom02) as Transportista from "&oficina&"_extranet.d01conte d" & _
	" left join "&oficina&"_extranet.e01oemb e on d.peri01 = e.peri01 and d.nemb01 = e.nemb01" & _
	" left join "&oficina&"_extranet.c56trans c on c.cve02 = e.ctra01 " & _
	"where d.refe01 ='"&referencia&"' and d.nemb01 != 0 group by d.refe01;"
	
	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427" 
	act2.ActiveConnection = conn12
	act2.Source = sqlAct
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()
	if not(act2.eof) then
		while not act2.eof
			Trans =act2.fields("Transportista").value
			act2.movenext()
		wend
	retornaTransportista = Trans
	else
	retornaTransportista =Trans
	end if 
end function 
%>

