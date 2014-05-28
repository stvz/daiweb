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

'																										 		'
'																												'
' ---------------------------------------        EXTRACOSTOS       --------------------------------------------	'
'																												'
' 
'se ejecuta de la siguiente forma:
'http://10.66.1.9/portalmysql/extranet/ext-asp/reportes/unilever-estados-112011-RSG.asp?finicio=01/01/2011&ffinal=15/11/2011&tipope=e&det=

Response.Buffer = TRUE
response.Charset = "utf-8"
Response.Addheader "Content-Disposition", "attachment; filename=BookletUNILEVER_EXTRACOSTOS_.xls"'
Response.ContentType = "application/vnd.ms-excel"

dim strTipoUsuario,fechaini,fechafin,oficina

'Estas lineas se comentaron para ejecutar el archivo directamente del explorador
'tipope = request.Form("rbnTipoDate")
'fechaini = trim(request.Form("txtDateIni")) 
'fechafin = trim(request.Form("txtDateFin"))
'oficina=Request.QueryString("ofi")

strTipoUsuario = Session("GTipoUsuario")
tipope	 = Request.QueryString("tipope")
fechaini = Request.QueryString("finicio")
fechafin = Request.QueryString("ffinal")
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
  
	dim orden(50)
	dim subrefaux,subref,bgcolor,strHTML
	subrefaux=""
	subref=""
	bgcolor="#FFFFFF"
	strHTML = ""
	
	Server.ScriptTimeOut=10000000

%>
<title> ReporteEXTRACOSTOS.. </title>
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
<table x:str border=0 cellpadding=0 cellspacing=0 width=12637 style='border-collapse:
 collapse;table-layout:fixed;width:9479pt'>
	 <col width=125 style='mso-width-source:userset;mso-width-alt:4571;width:94pt'>
	 <col width=100 span=2 style='mso-width-source:userset;mso-width-alt:3657;
	 width:75pt'>
	 <col class=xl26 width=370 style='mso-width-source:userset;mso-width-alt:8265;
	 width:370pt'>
	 <col width=100 span=4 style='mso-width-source:userset;mso-width-alt:3657;
	 width:75pt'>
	 <col width=259 style='mso-width-source:userset;mso-width-alt:7753;width:259pt'>
	 <col width=100 span=11 style='mso-width-source:userset;mso-width-alt:3657;
	 width:75pt'>
	 <col class=xl26 width=214 style='mso-width-source:userset;mso-width-alt:7826;
	 width:161pt'>
	 <col width=100 span=8 style='mso-width-source:userset;mso-width-alt:3657;
	 width:75pt'>
	 <col class=xl26 width=100 style='mso-width-source:userset;mso-width-alt:3657;
	 width:75pt'>
	 <col width=100 span=67 style='mso-width-source:userset;mso-width-alt:3657;
	 width:75pt'>
	 <col width=80 span=32 style='width:60pt'>
	<tr class=xl27 height=52 style='height:39.0pt'>
<% genera_registros tipope %>
</table>
</body>
</html>
<%

end if


sub genera_registros(tipope)
	dim c
	c=chr(34)

%>
<!-- Genera Encabezados -->


  <td class=xl31 id="_x0000_s1033" x:autofilter="all" width=212
  style='width:159pt'>Proveedor</td>
    
  <td class=xl36 id="_x0000_s1043" x:autofilter="all" width=100
  style='width:75pt'>Factura</td>
  <td class=xl37 id="_x0000_s1044" x:autofilter="all" width=100
  style='width:75pt'>Fecha de Factura</td>
  <td class=xl33 id="_x0000_s1053" x:autofilter="all" width=100
  style='width:75pt'>No. De Trafico</td>
      <td class=xl33 id="_x0000_s1053" x:autofilter="all" width=100
  style='width:75pt'>No. De Trafico Rectificado</td>
	<td class=xl54 id="_x0000_s1081" x:autofilter="all" width=100
  style='width:75pt' x:str="No. CTA DE GASTOS"><span
  style='mso-spacerun:yes'> </span>No. CTA DE GASTOS<span
  style='mso-spacerun:yes'> </span></td>
  
  <td class=xl33 id="_x0000_s1055" x:autofilter="all" width=100
  style='width:75pt'>No. Pedimento</td>
   <td class=xl33 id="_x0000_s1055" x:autofilter="all" width=100
  style='width:75pt'>No. Pedimento Rectificado</td>
  <td class=xl33 id="_x0000_s1056" x:autofilter="all" width=100
  style='width:75pt'>Fecha Pedimento</td>
  <td class=xl60 id="_x0000_s1108" x:autofilter="all" width=100
  style='width:75pt' x:str="FLETE TERRESTRE"><span
  style='mso-spacerun:yes'> </span>FLETE TERRESTRE<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1109" x:autofilter="all" width=100
  style='width:75pt' x:str="IMPUESTOS SEGUN PEDIMENTO"><span
  style='mso-spacerun:yes'> </span>IMPUESTOS SEGUN PEDIMENTO<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1110" x:autofilter="all" width=100
  style='width:75pt' x:str="ESTADIAS"><span
  style='mso-spacerun:yes'> </span>ESTADIAS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1111" x:autofilter="all" width=100
  style='width:75pt' x:str="MANIOBRAS "><span
  style='mso-spacerun:yes'> </span>MANIOBRAS<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl61 id="_x0000_s1112" x:autofilter="all" width=100
  style='width:75pt' x:str="ALMACENAJES"><span
  style='mso-spacerun:yes'> </span>ALMACENAJES<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1113" x:autofilter="all" width=100
  style='width:75pt' x:str="OTROS"><span
  style='mso-spacerun:yes'> </span>OTROS<span style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1114" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS"><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl59 id="_x0000_s1115" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS USD"><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS USD<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl33 id="_x0000_s1053" x:autofilter="all" width=100
  style='width:75pt'>TIPO DE OPERACIÓN</td>
<td class=xl33 id="_x0000_s1053" x:autofilter="all" width=100
  style='width:75pt'>Planta(origen)</td>
<td class=xl33 id="_x0000_s1053" x:autofilter="all" width=100
  style='width:75pt'>Puerto (origen)</td>
    <td class=xl33 id="_x0000_s1053" x:autofilter="all" width=100
  style='width:75pt'>País destino (destino)</td>
  <td class=xl33 id="_x0000_s1053" x:autofilter="all" width=100
  style='width:75pt'>OBSERVACIONES</td>




</tr>
 <%
sqlAct= "select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'MARITIMO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01)  as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" cf6.import36 as '75', " & _
" cf1.import36 as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" cf3.import36  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from rku_extranet.ssdage01 as i  " & _
"   left join rku_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join rku_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" LEFT join rku_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" LEFT join rku_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join rku_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join rku_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join rku_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join rku_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join rku_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join rku_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join rku_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join rku_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join rku_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join rku_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join rku_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null " & _
"    and i.firmae01 <> '' and  cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
"select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'MARITIMO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01)  as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" cf6.import36 as '75', " & _
" cf1.import36 as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" cf3.import36  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from sap_extranet.ssdage01 as i  " & _
"   left join sap_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join sap_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" LEFT join sap_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" LEFT join sap_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join sap_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join sap_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join sap_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join sap_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join sap_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join sap_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join sap_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join sap_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join sap_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join sap_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join sap_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null " & _
"    and i.firmae01 <> '' and  cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
"select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'MARITIMO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01)  as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" cf6.import36 as '75', " & _
" cf1.import36 as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" cf3.import36  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from lzr_extranet.ssdage01 as i  " & _
"   left join lzr_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join lzr_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" LEFT join lzr_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" LEFT join lzr_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join lzr_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join lzr_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join lzr_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join lzr_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join lzr_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join lzr_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join lzr_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join lzr_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join lzr_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join lzr_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join lzr_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null " & _
"    and i.firmae01 <> '' and  cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
"select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'MARITIMO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01)  as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" cf6.import36 as '75', " & _
" cf1.import36 as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" cf3.import36  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from ceg_extranet.ssdage01 as i  " & _
"   left join ceg_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join ceg_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" LEFT join ceg_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" LEFT join ceg_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join ceg_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join ceg_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join ceg_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join ceg_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join ceg_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join ceg_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join ceg_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join ceg_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join ceg_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join ceg_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join ceg_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null " & _
"    and i.firmae01 <> '' and  cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
" select  " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'AEREO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01)  as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" cf6.import36 as '75', " & _
" cf1.import36 as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" cf3.import36  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from dai_extranet.ssdage01 as i  " & _
"   left join dai_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join dai_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" LEFT join dai_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" LEFT join dai_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join dai_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join dai_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join dai_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join dai_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join dai_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join dai_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join dai_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join dai_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join dai_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join dai_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join dai_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> ''  and   cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " & _
" union all " & _
" select   " & _
" i.refcia01,fr.fraarn02,fr.ordfra02, " & _
" i.cvecli01 as '1', " & _
" 'unilever'  as '2',  " & _
" 'unilever'  as '3', " & _
" ar.desc05 as '4', " & _
" 'unilever' as '5', " & _
" 'unilever' as '6', " & _
" 'unilever' as '7', " & _
" 'unilever' as '8', " & _
" prv.nompro22 as '9', " & _
" '' as '10',  " & _
" r.rcli01 as '11', " & _
" ar.pedi05 as '12',  " & _
" 'unilever' as '13',  " & _
" i.cvepod01 as '14', " & _
" i.cvepvc01 as '15', " & _
" '?' as '16',  " & _
" r.ptoemb01 as '177', " & _
" r.cveptoemb as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" r.impo01 as '21',  " & _
" ar.caco05 as '22', " & _
" um.descri31 as '23', " & _
" f.terfac39 as '24', " & _
" 'AEREO' as '25', " & _
" i.adusec01 as '26',  " & _
" i.patent01 as '27',  " & _
" i.patent01 as '28',  " & _
" i.refcia01 as '29', " & _
" '?' as '30', " & _
" i.numped01 as '31', " & _
" i.fecpag01 as '32', " & _
" Month(i.fecpag01) as '33',  " & _
" week(i.fecpag01) as '34', " & _
" 'N/A' as '35',  " & _
" '?' as '36', " & _
" '?' as '37', " & _
" '?' as '38', " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '39', " & _
" r.feorig01 as '40', " & _
" r.frev01 as '41',  " & _
" r.fdsp01 as '42', " & _
" '?' as '57', " & _
" '?' as '58',  " & _
" fr.prepag02   as '59',  " & _
" (fr.prepag02/i.tipcam01) as '60',  " & _
" i.fletes01 as '61', " & _
" i.segros01 as '62', " & _
" i.incble01 as '63',  " & _
" fr.vaduan02 as '64',  " & _
" i.tipcam01 as '65',  " & _
" (fr.vaduan02/i.tipcam01) as '66', " & _
" '' as '67', " & _
" '' as '68', " & _
" (i.fletes01/i.tipcam01) as '69',  " & _
" 'unilever' as '70',  " & _
" fr.fraarn02 as '71', 	 " & _
" fr.tasadv02 as '72',   " & _
" if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73',  " & _
" 'unilever' as '74', " & _
" cf6.import36 as '75', " & _
" cf1.import36 as '76', " & _
" (fr.i_adv102+fr.i_adv202) as '761'," & _
" fr.tasiva02 as '77', " & _
" cf3.import36  as '78', " & _
" (fr.i_iva102+fr.i_iva202) as '781'," & _
" cf15.import36  as '79',  " & _
" '?'as '80', fr.ordfra02 as '81', count(fr.ordfra02) as '82', " & _
"  ar.item05 as 'Item05' , i.firmae01 as 'firmita',  " & _
"  transp.descri30  as 'sDescTransp'  " & _
"   from tol_extranet.ssdage01 as i  " & _
"   left join tol_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join tol_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" LEFT join tol_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _ 
" LEFT join tol_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _ 
"       left join tol_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join tol_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join tol_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join tol_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join tol_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join tol_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join tol_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join tol_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join tol_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join tol_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join tol_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.firmae01 is not null and i.firmae01 <> '' and  cta.fech31 >=  '"& finicio &"' and cta.fech31 <= '"& ffinal &"' " & _ 
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 "
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

	dim ref,refAux
	dim cambio,rcli 
	cambio = 1
	rcli=""
	contass=0
	while not act2.eof
		response.Write("<tr align="&c&"center"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")
		ref = act2.fields("29").value
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
		
		' Dim ref22 = 
		 
		 
		'comienza la impresion
		
		
		
		
		genera_html "d",act2.fields("9").value,"center"  'Proveedor
		
		rcli =act2.fields("11").value
						

		genera_html "d",act2.fields("19").value,"center"  'Factura
		genera_html "d",act2.fields("20").value,"center"  'Fecha de Factura
		genera_html "d",act2.fields("29").value,"center"  'No. De Trafico
		genera_html "d","","center"  'No. De Trafico Rectificado

		ref = act2.fields("29").value
		if (ref <> refAux)then
			refAux=ref
			
		 'Lote 2
			genera_html "d",retornaCampoCtaGastos(act2.fields("29").value,"cgas31",mid(act2.fields("29").value,1,3)),"center"  'No. CTA DE GASTOS
			genera_html "d",act2.fields("31").value,"center"  'No. Pedimento
			genera_html "d","","center"  'No. Pedimento Rectificado
			genera_html "d",act2.fields("32").value,"center"  'Fecha Pedimento

			Subref = act2.fields("81").value
			'if (ordenOcupado(Subref,act2.fields("29").value) = False)then
				ocuparOrd(Subref)
			'-----
				
 
			genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"FLETES"),"E",mid(act2.fields("29").value,1,3)),"center"  ' FLETES GTOS. ADUANA USD(SOLO FRONTERA) 
			genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"IMPUESTOS"),"E",mid(act2.fields("29").value,1,3)),"center"  ' IMPUESTOS
			genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"E",mid(act2.fields("29").value,1,3)),"center"  ' ESTADIAS 
			 
			if ucase(mid(act2.fields("29").value,1,3)) ="RKU" or ucase(mid(act2.fields("29").value,1,3)) ="SAP" then
				genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"E",mid(act2.fields("29").value,1,3)),"center"  ' MANIOBRAS  
				genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"E",mid(act2.fields("29").value,1,3)),"center"  ' ALMACENAJES 
			else
				genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"E",mid(act2.fields("29").value,1,3)),"center"  ' MANIOBRAS
				genera_html "d","?","center"  ' ALMACENAJES 
			end if
			
			dim TPH,DEM,EST,ALMMAN,MAN
			TPH=0
			DEM=0
			EST=0
			ALMAN=0
			MAN=0

			TPH=retornaTOTALPagosHechos(act2.fields("29").value,"E",mid(act2.fields("29").value,1,3))
			FLET= retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"FLETES"),"E",mid(act2.fields("29").value,1,3))  ' FLETES
			DEM= retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"IMPUESTOS"),"E",mid(act2.fields("29").value,1,3))
			EST = retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"E",mid(act2.fields("29").value,1,3))
			ALMAN=retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"E",mid(act2.fields("29").value,1,3))
			if ucase(mid(act2.fields("29").value,1,3)) ="RKU" or ucase(mid(act2.fields("29").value,1,3)) ="SAP" then
				MAN=retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"E",mid(act2.fields("29").value,1,3))
			end if
			if(revisaImpuestosFacturados( act2.fields("29").value,"E",mid(act2.fields("29").value,1,3)) <> 0 )then

				
			end if
			contass=contass+1

			on error resume next
				'genera_html "d",TPH-DEM-EST-ALMAN-MAN-STIMP-STIVA,"center"  ' OTROS 
				genera_html "d",TPH-DEM-EST-ALMAN-MAN,"center"  ' OTROS 
				genera_html "d",retornaTOTALPagosHechos(act2.fields("29").value,"E",mid(act2.fields("29").value,1,3)),"center"  ' TOTAL GASTOS DIVERSOS 
				genera_html "d",((retornaTOTALPagosHechos(act2.fields("29").value,"E",mid(act2.fields("29").value,1,3)))/act2.fields("65").value),"center"  ' TOTAL GASTOS DIVERSOS USD 						
			if err.number <> 0 then
				genera_html "d","Error","center"  ' OTROS 
				genera_html "d","Error","center"  ' TOTAL GASTOS DIVERSOS 
				genera_html "d","Error","center"  ' TOTAL GASTOS DIVERSOS USD
			end if
			 

		else
			'Lote 2
			genera_html "d",retornaCampoCtaGastos(act2.fields("29").value,"cgas31",mid(act2.fields("29").value,1,3)),"center"  'No. CTA DE GASTOS
			genera_html "d",act2.fields("31").value,"center"  'No. Pedimento
			genera_html "d",act2.fields("32").value,"center"  'Fecha Pedimento


			Subref = act2.fields("81").value
	

			genera_html "d","","center"  ' FLETES
			genera_html "d","","center"  ' DEMORAS 
			genera_html "d","","center"  ' ESTADIAS 
			genera_html "d","","center"  ' MANIOBRAS  
			genera_html "d","","center"  ' ALMACENAJES 
			genera_html "d","","center"  ' OTROS 
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS 
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS USD 
			
			'/Lote 1
		end if

		
		genera_html "d","","center"  'TIPO DE OPERACIÓN
		genera_html "d","","center"  'PLANTA ORIGEN
		genera_html "d",act2.fields("26").value,"center"  'PUERTO ORIGEN
		genera_html "d",act2.fields("14").value,"center"  'PAIS DESTINO
		genera_html "d","","center"  'OBSERVACIONES
		
		

		
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
 end if
 
 
sqlAct="select count(i.refcia01) as Ref " & _
" from "& oficina &"_extranet.ssdag" & tipoop &"01 as i  " & _
"  inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  " & _
"     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 " & _
"          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 " & _
"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'   and ep.tmov21 =dp.tmov21 " & _
"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
"    where  i.firmae01 <> ''  and cta.esta31 <> 'C'  and i.refcia01 = '"& referencia &"' and ep.conc21 = 1"

Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = cadena_de_conexion()
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="& oficina &"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

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
 end if
 
 
sqlAct="select max(date_format(cta.fech31,'%d/%m/%Y')) as fech31 from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
" r.refe31 = '"&referencia&"' and cta.esta31 <> 'C' "

Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = cadena_de_conexion()
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

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
 end if
 


sqlAct="select ifnull(COUNT(cta.cgas31),0)  as total from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
" r.refe31 = '"&referencia&"'  "

Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

act2.ActiveConnection = conn12
act2.Source = sqlAct
act2.cursortype=0
act2.cursorlocation=2
act2.locktype=1
act2.open()
if not(act2.eof) then
	if(cInt(act2.fields("total").value) = 1) then
	
		valor="NORMAL"
	end if
	if(cInt(act2.fields("total").value) = 2)then
	
		 
		sqlAct321="select ifnull(COUNT(cta.cgas31),0)  as totalcg from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
		" where cta.cgas31 = r.cgas31 and "&_
		" r.refe31 = '"&referencia&"'  "&_
		"  and cta.esta31 <> 'C' "
		
		Set act211= Server.CreateObject("ADODB.Recordset")
		conn1211 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

		act211.ActiveConnection = conn1211
		act211.Source = sqlAct321
		act211.cursortype=0
		act211.cursorlocation=2
		act211.locktype=1
		act211.open()
		if not(act211.eof) then
			if(cInt(act211.fields("totalcg").value) = 1)then
					
				valor="REFACTURADA"
			end if
			if(cInt(act211.fields("totalcg").value) = 2)then
			
				valor="COMPLEMENTARIA"
			end if
		end if
	end if
	if(cInt(act2.fields("total").value) > 2)then
	
		valor="REFACTURADA COMPLEMENTARIA"
	end if
	regresa_tipo_Cgastos =valor
else
  regresa_tipo_Cgastos =valor
   end if
end function




function retornaCampoCtaGastos(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
 sqlAct = "select  r."& campo &" as campo from "&oficina&"_extranet.e31cgast as cta " &_
 " inner join  "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
 " where  r.refe31 = '"& referencia &"' and cta.esta31 <> 'C' "

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
 end if
 
 
 sqlAct = " select sum(dm.mont11) as campo " & _
			" from "&oficina&"_extranet.d11movim as dm " & _
			" where dm.refe11='"& referencia&"' and dm.conc11 = '"&campo&"' "
	'and dm.cgas11 <> ''"

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
  retornaMontoAnticipo = valor
 else
  retornaMontoAnticipo =valor
 end if
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
 'response.write("Select * From A1:B15 where desccli like '%" & desc2 & "%'")
 'response.End()
if not(rsVac.eof)then
 res =rsVac.fields("cvecli")  '&","&rsVac.fields("desccli")

else
 res = desc
end if 

codigoCliente = res
end function


function retornaConceptosPH(oficina,topico)
dim cad
cad = "NA"

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if oficina = "SAP" then
if topico = "FLETES" then
  cad= "5"
end if
if topico = "ALMACENAJES-MANIOBRAS" then
  cad= "3"
end if
if topico = "IMPUESTOS" then
  cad= "6"
end if
if topico = "ESTADIAS" then
  cad="144"
end if
if topico = "OTROS" then
  cad="306,313,350,351,352"
end if
if topico = "MANIOBRAS" then
  cad="2"
end if

' if oficina = "SAP" then

' if topico = "ALMACENAJES-MANIOBRAS" then
  ' cad= "2"
' end if
' if topico = "DEMORAS" then
  ' cad= "6,14,46,63,129,156,352"
' end if
' if topico = "ESTADIAS" then
  ' cad="144"
' end if
' if topico = "OTROS" then
  ' cad="306,313,350,351,352"
' end if
' if topico = "MANIOBRAS" then
  ' cad="2"
' end if



else 
  if oficina = "CEG" then
  
    if topico = "ALMACENAJES-MANIOBRAS" then
     cad= "2,4,59,77,100,223,223,235,235,241"
    end if
	if topico = "IMPUESTOS" then
	 cad= "1"
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
		if topico = "IMPUESTOS" then
		 cad= "1"
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
			if topico = "IMPUESTOS" then
			 cad= "1"
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
					if topico = "IMPUESTOS" then
					 cad= "1"
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
						if topico = "IMPUESTOS" then
						 cad= "1"
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
 end if
 


sqlAct =" select  c.import36 as Campo ,c.cveimp36,c.refcia36  from "& oficina &"_extranet.ssdag"& tipope &"01 as i " & _
		"  inner  join  "& oficina &"_extranet.sscont36 as c on i.refcia01 = c.refcia36 " & _
		"    where c.refcia36 = '"& referencia &"' and c.cveimp36 = '18' and i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3')"


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
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then


'sqlAct =" select i.refcia01 as Ref,sum(dp.mont21*if(ep.deha21 = 'C',-1,1)) as Importe " & _
'		" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
'		" inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 " & _
'		" inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and ep.fech21 = dp.fech21 and ep.conc21 in ("&conceptos&") and ep.esta21 <> 'S' and ep.esta21 <> 'C' " & _
'		" where i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3') and i.refcia01 = '"&referencia&"'  and i.firmae01 <> ''   group by i.refcia01 "

sqlAct="select i.refcia01 as Ref, r.cgas31,ep.conc21,ep.piva21,ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as Importe, cp.desc21 " & _
" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
"  inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  " & _
"     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 " & _
"          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 " & _
"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  and ep.tmov21 =dp.tmov21 " & _
"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
"    where  i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3')  and i.firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in ("&conceptos&") and i.refcia01 = '"&referencia&"'  group by Ref,cgas31,conc21"

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

 if not(act2.eof) then
 valor = act2.fields("Importe").value
 'act2.movenext()
 'while not act2.eof
 '  valor = valor&", "&act2.fields("Importe").value
 '  act2.movenext()
 'wend
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
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then

sqlAct =" select i.refcia01 as Ref,ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as Importe " & _
		" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
		" inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 " & _
		" inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.tmov21 =dp.tmov21 " & _
		" where i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3')  and i.refcia01 = '"&referencia&"'  and i.firmae01 <> ''  group by i.refcia01 "

'response.Write(sqlAct)
'response.End()

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
   retornaTOTALPagosHechos = 0
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


function revisaFraccion(fraccion)
dim val
val ="(pendiente)"
fraccion = trim(fraccion)
if (fraccion = "90230001" or _
 fraccion = "49111099" or _
 fraccion = "48219099" or _
 fraccion = "39235001" or _
 fraccion = "34011101" or _
 fraccion = "33072001" or _
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



function retornaTipoUnidad(referencia,tipope,oficina)
dim valor,unidad
 valor=""
 unidad=""
 
if (ucase(oficina) = "ALC")then
 oficina = "LZR"
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
   'valor = valor &", "& act2.fields("campo").value
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
%>

