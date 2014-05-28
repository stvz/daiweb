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

Response.Buffer = TRUE
response.Charset = "utf-8"
Response.Addheader "Content-Disposition", "attachment; filename=Booklet_EXTRACOSTOS_.xls"
Response.ContentType = "application/vnd.ms-excel"

dim strTipoUsuario,fechaini,fechafin,oficina

strTipoUsuario = Session("GTipoUsuario")
tipope	 = Request.QueryString("tipope")
fechaini = Request.QueryString("finicio")
fechafin = Request.QueryString("ffinal")
oficina = Request.QueryString("OficinaG")


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
<td height=52 class=xl46 id="_x0000_s1071" x:autofilter="all"
  x:autofilterrange="$A$1:$CD$800" width=125 style='height:39.0pt;width:94pt'
  x:str="EJECUTIVO UL ">EJECUTIVO UL<span style='mso-spacerun:yes'> </span></td>
  <td class=xl29 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>DIVISION</td>
  <td class=xl33 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>No. De Trafico</td>
  <td class=xl46 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>RO</td>
  <td class=xl29 id="_x0000_s1121" x:autofilter="all" width=100
  style='width:75pt' x:str="Planta de entrega">Planta de entrega</td>
    <td class=xl46 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>Rectificacion</td>
  <td class=xl29 id="_x0000_s1027" x:autofilter="all" width=100
  style='width:75pt'>CATEGORIA</td>
  <td class=xl29 id="_x0000_s1028" x:autofilter="all" width=475
  style='width:475pt'>Nombre del material</td>
  <td class=xl29 id="_x0000_s1030" x:autofilter="all" width=100
  style='width:75pt'>Clase de Producto</td>
    <td class=xl31 id="_x0000_s1033" x:autofilter="all" width=212
  style='width:159pt'>Proveedor</td>
  <td class=xl28 id="_x0000_s1034" x:autofilter="all" width=100
  style='width:75pt' x:str="CUENTA ">CUENTA</td>
  <td class=xl28 id="_x0000_s1035" x:autofilter="all" width=100
  style='width:75pt'>CECO</td>
  <td class=xl32 id="_x0000_s1036" x:autofilter="all" width=100
  style='width:75pt'>ODC</td>
    <td class=xl33 id="_x0000_s1039" x:autofilter="all" width=100
  style='width:75pt'>País Origen</td>
  <td class=xl33 id="_x0000_s1039" x:autofilter="all" width=100
  style='width:75pt'>País de Procedencia</td>
  <td class=xl36 id="_x0000_s1043" x:autofilter="all" width=100
  style='width:75pt'>Factura</td>
  <td class=xl38 id="_x0000_s1045" x:autofilter="all" width=214
  style='width:161pt'>IMPORTADOR</td>
  <td class=xl39 id="_x0000_s1046" x:autofilter="all" width=100
  style='width:75pt' x:str="Cantidad ">Cantidad<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl33 id="_x0000_s1047" x:autofilter="all" width=100
  style='width:75pt'>Unidad de Medida</td>
  <td class=xl33 id="_x0000_s1044" x:autofilter="all" width=100
  style='width:75pt'>Peso Bruto KG</td>
    <td class=xl46 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>Tipo Producto</td>
  <td class=xl33 id="_x0000_s1048" x:autofilter="all" width=100
  style='width:75pt'>Incoterms</td>
  <td class=xl33 id="_x0000_s1049" x:autofilter="all" width=100
  style='width:75pt'>Tipo de Transporte</td>
  <td class=xl33 id="_x0000_s1050" x:autofilter="all" width=100
  style='width:75pt'>Aduana</td>
  <td class=xl33 id="_x0000_s1051" x:autofilter="all" width=100
  style='width:75pt'>Agente Aduanal</td>
  <td class=xl40 id="_x0000_s1052" x:autofilter="all" width=100
  style='width:75pt'>Patente Agente Aduanal</td>
  <td class=xl38 id="_x0000_s1054" x:autofilter="all" width=100
  style='width:75pt' x:str="No de Contenedor "></td>
  <td class=xl46 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>Numero BL</td>
  <td class=xl38 id="_x0000_s1054" x:autofilter="all" width=100
  style='width:75pt' x:str="Clave Pedimento "></td>
    <td class=xl33 id="_x0000_s1055" x:autofilter="all" width=100
  style='width:75pt'>No. Pedimento</td>
  <td class=xl33 id="_x0000_s1056" x:autofilter="all" width=100
  style='width:75pt'>Fecha Pedimento</td>
  <td class=xl40 id="_x0000_s1057" x:autofilter="all" width=100
  style='width:75pt'>Mes</td>
  <td class=xl38 id="_x0000_s1058" x:autofilter="all" width=100
  style='width:75pt'>No.Semana</td>
  <td class=xl41 id="_x0000_s1059" x:autofilter="all" width=100
  style='width:75pt' x:str="Cantidad de Operaciones "></td>
  <td class=xl41 id="_x0000_s1060" x:autofilter="all" width=100
  style='width:75pt'>Cantidad de Contenedores</td>
  <td class=xl41 id="_x0000_s1061" x:autofilter="all" width=100
  style='width:75pt' x:str="PALLETS/BULTOS ">PALLETS/BULTOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl42 id="_x0000_s1062" x:autofilter="all" width=100
  style='width:75pt'>TIPO DE CONTENEDOR/ CAJA</td>
  <td class=xl41 id="_x0000_s1063" x:autofilter="all" width=100
  style='width:75pt'>Fecha Factura</td>
  <td class=xl48 id="_x0000_s1064" x:autofilter="all" width=100
  style='width:75pt'>Fecha BL</td>
  <td class=xl43 id="_x0000_s1065" x:autofilter="all" width=100
  style='width:75pt'>Fecha de arribo a la aduana</td>
  <td class=xl43 id="_x0000_s1066" x:autofilter="all" width=100
  style='width:75pt'>Fecha Desaduanamiento</td>
  <td class=xl47 id="_x0000_s1067" x:autofilter="all" width=100
  style='width:75pt'>EXC</td>
  <td class=xl47 id="_x0000_s1068" x:autofilter="all" width=100
  style='width:75pt'>EXC $ M.N.</td>
  <td class=xl47 id="_x0000_s1069" x:autofilter="all" width=100
  style='width:75pt'>EXC S USD.</td>
  <td class=xl47 id="_x0000_s1070" x:autofilter="all" width=100
  style='width:75pt'>EXP</td>
  <td class=xl47 id="_x0000_s1071" x:autofilter="all" width=100
  style='width:75pt' x:str="EXP & USD ? ">EXP $ USD<span style='mso-spacerun:yes'> </span></td>
  <td class=xl47 id="_x0000_s1072" x:autofilter="all" width=100
  style='width:75pt' x:str="EXC & EXPT LT $ ">EXC & EXP LT $<span style='mso-spacerun:yes'> </span></td>
  <td class=xl47 id="_x0000_s1073" x:autofilter="all" width=100
  style='width:75pt'>CAUSAL LT</td>
  <td class=xl47 id="_x0000_s1074" x:autofilter="all" width=100
  style='width:75pt'>CAUSA RAIZ</td>
  <td class=xl47 id="_x0000_s1075" x:autofilter="all" width=100
  style='width:75pt'>PLAN DE ACCION</td>
  <td class=xl47 id="_x0000_s1076" x:autofilter="all" width=100
  style='width:75pt'>RESPONSABLE</td>
  <td class=xl47 id="_x0000_s1077" x:autofilter="all" width=100
  style='width:75pt'>FECHA DE CUMPLIMIENTO</td>
  <td class=xl47 id="_x0000_s1078" x:autofilter="all" width=100
  style='width:75pt'>IMPACTO</td>
    <td class=xl54 id="_x0000_s1081" x:autofilter="all" width=100
  style='width:75pt' x:str="No. CTA DE GASTOS"><span
  style='mso-spacerun:yes'> </span>No. CTA DE GASTOS<span
  style='mso-spacerun:yes'> </span></td>
    <td class=xl46 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>Fecha CG</td>
  <td class=xl56 id="_x0000_s1097" x:autofilter="all" width=100
  style='width:75pt' x:str="FRACC. ARANC."><span
  style='mso-spacerun:yes'> </span>FRACC. ARANC.<span
  style='mso-spacerun:yes'> </span></td>
    <td class=xl58 id="_x0000_s1101" x:autofilter="all" width=100
  style='width:75pt' x:str="ADV. $ / IGI $"><span
  style='mso-spacerun:yes'> </span>ADV. $ / IGI $<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl58 id="_x0000_s1102" x:autofilter="all" width=100
  style='width:75pt' x:str="DTA $"><span style='mso-spacerun:yes'> </span>DTA
  $<span style='mso-spacerun:yes'> </span></td>
  <td class=xl57 id="_x0000_s1103" x:autofilter="all" width=100
  style='width:75pt'>IVA %</td>
  <td class=xl58 id="_x0000_s1104" x:autofilter="all" width=100
  style='width:75pt' x:str="IVA $"><span style='mso-spacerun:yes'> </span>IVA
  $<span style='mso-spacerun:yes'> </span></td>
  <td class=xl58 id="_x0000_s1105" x:autofilter="all" width=100
  style='width:75pt' x:str="PREVAL. "><span
  style='mso-spacerun:yes'> </span>PREVAL.<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl46 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>ECI</td>
  <td class=xl58 id="_x0000_s1106" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL IMPUESTOS"><span
  style='mso-spacerun:yes'> </span>TOTAL IMPUESTOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl47 id="_x0000_s1107" x:autofilter="all" width=100
  style='width:75pt' x:str="Total Impuestos USD "><span
  style='mso-spacerun:yes'> </span>Total Impuestos USD<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl47 id="_x0000_s1091" x:autofilter="all" width=100
  style='width:75pt' x:str="Tipo de Cambio"><span
  style='mso-spacerun:yes'> </span>Tipo de Cambio<span style='mso-spacerun:yes'> </span></td>
    <td class=xl46 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>Valor de la Mercancia en aduana</td>
  <td class=xl60 id="_x0000_s1108" x:autofilter="all" width=100
  style='width:75pt' x:str="GTOS. ADUANA USD(SOLO FRONTERA)"><span
  style='mso-spacerun:yes'> </span>GTOS. ADUANA USD(SOLO FRONTERA)<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1109" x:autofilter="all" width=100
  style='width:75pt' x:str="DEMORAS (Costo Aduanal)"><span
  style='mso-spacerun:yes'> </span>DEMORAS (Costo Aduanal)<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1110" x:autofilter="all" width=100
  style='width:75pt' x:str="ESTADIAS (Costo Aduanal)"><span
  style='mso-spacerun:yes'> </span>ESTADIAS (Costo Aduanal)<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1111" x:autofilter="all" width=100
  style='width:75pt' x:str="MANIOBRAS (Costo Aduanal)"><span
  style='mso-spacerun:yes'> </span>MANIOBRAS (Costo Aduanal)<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl61 id="_x0000_s1112" x:autofilter="all" width=100
  style='width:75pt' x:str="ALMACENAJES (Costo Aduanal)"><span
  style='mso-spacerun:yes'> </span>ALMACENAJES (Costo Aduanal)<span
  style='mso-spacerun:yes'> </span></td>
    <td class=xl61 id="_x0000_s1112" x:autofilter="all" width=100
  style='width:75pt' x:str="CONEXION REFRIGERADO (Costo Aduanal)"><span
  style='mso-spacerun:yes'> </span>CONEXION REFRIGERADO (Costo Aduanal)<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1112" x:autofilter="all" width=100
  style='width:75pt' x:str="CONEXION REFRIGERADO (Costo Directo)"><span
  style='mso-spacerun:yes'> </span>CONEXION REFRIGERADO (Costo Directo)<span
  style='mso-spacerun:yes'> </span></td>
    <td class=xl61 id="_x0000_s1111" x:autofilter="all" width=100
  style='width:75pt' x:str="MANIOBRAS (Tarifa FLAT)"><span
  style='mso-spacerun:yes'> </span>MANIOBRAS (Tarifa FLAT)<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl61 id="_x0000_s1114" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS"><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1114" x:autofilter="all" width=100
  style='width:75pt' x:str="TC"><span
  style='mso-spacerun:yes'> </span>TC<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl62 id="_x0000_s1116" x:autofilter="all" width=100
  style='width:75pt' x:str="HONORARIOS AG AD. $"></td>
  <td class=xl59 id="_x0000_s1115" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS USD"><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS USD<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl62 id="_x0000_s1038" x:autofilter="all" width=100
  style='width:75pt' x:str="NAVIERA"></td>
  <td class=xl62 id="_x0000_s1044" x:autofilter="all" width=100
  style='width:75pt'>TRANSPORTISTA NACIONAL</td>
  <td class=xl33 id="_x0000_s1126" x:autofilter="all" width=100
  style='width:75pt' x:str="TIPO TRANSPORTE"></td>
  <td class=xl33 id="_x0000_s1127" x:autofilter="all" width=100
  style='width:75pt' x:str="TIPO DE UNIDAD"></td>
</tr>
 <%
	Dim sqlQuery
	sqlQuery=""
	if oficina<>"Todas"then
		sqlQuery=QueryMySql(oficina)
	else 
	sqlQuery=""
		For i = 0 to 5
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
				case 5
					aduanaTmp="ceg"
			End Select
			sqlQuery=sqlQuery& QueryMySql(aduanaTmp)
			
			if (i<>5) then
				sqlQuery= sqlQuery& " UNION ALL " & chr(13) & chr(10)
			end if
		Next 
		
	end if
	
	Set act2= Server.CreateObject("ADODB.Recordset")
	conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
	act2.ActiveConnection = conn12
	act2.Source = sqlQuery
	act2.cursortype=0
	act2.cursorlocation=2
	act2.locktype=1
	act2.open()

	dim ref,refAux
	dim cambio,rcli 
	cambio = 1
	rcli=""
	contass=0 
	Dim PrimerCiclo,refAnt, refActual
	PrimerCiclo=0
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
		
		'comienza la impresion
		genera_html "d",act2.fields("Ejecutivo").value,"center"  'Ejecutivo 
		genera_html "d",retornaDivision(act2.fields("1").value,act2.fields("71").value),"center"  'DIVISION 
		genera_html "d",act2.fields("29").value,"center"  'No. De Trafico
		genera_html "d",act2.fields("RO").value,"center" 'RO 
		genera_html "d",retornaPlantaEntrega(act2.fields("1").value),"center"  'PLANTA DE ENTREGA
		genera_html "d",act2.fields("Rectificacion").value,"center" 'Rectificacion
		genera_html "d","","center"  'CATEGORIA
		genera_html "d",act2.fields("4").value,"center"  'Nombre del Material
		genera_html "d","","center"  'Clase de Producto
		genera_html "d",act2.fields("9").value,"center"  'Proveedor
		rcli =act2.fields("11").value
		genera_html "d",retornaCuenta(rcli),"center"  'CUENTA
		genera_html "d",retornaCECO(rcli),"center"  'CECO
		if(act2.fields("177").value<>"")then
			PuertEmb= retornaCampoPuertoEmb(act2.fields("177").value,act2.fields("17").value,"cvepai01",mid(act2.fields("29").value,1,3))	
		else
			PuertEmb= ""	
		end if 
		genera_html "d",act2.fields("12").value,"center"  'ODC
		genera_html "d",act2.fields("14").value,"center"  'País de Origen 
		genera_html "d",PuertEmb,"center"  'País de Procedencia
		genera_html "d",act2.fields("19").value,"center"  'Factura
		genera_html "d",retornaIMPORTADOR(act2.fields("21").value,mid(act2.fields("29").value,1,3)),"center"  'IMPORTADOR
		genera_html "d",act2.fields("22").value,"center"  'Cantidad 
		genera_html "d",act2.fields("23").value,"center"  'Unidad de Medida
		
		refActual = act2.fields("refcia01").value
			primerCiclo = primerCiclo + 1
			pesoBruto = act2.fields("pesobr").value
		
		if primerCiclo > 1 and refActual = refAnt then
			pesoBruto = " "
		end if
			
		refAnt = refActual
		genera_html "d",pesoBruto,"center"  'Peso Bruto KG
		genera_html "d",act2.fields("TipMerc").value,"center" ' Tipo de mercancia
		genera_html "d",act2.fields("24").value,"center"  'Incoterms
		genera_html "d",act2.fields("25").value,"center"  'Tipo de Transporte
		genera_html "d",retornaAduana(act2.fields("26").value),"center"  'Aduana
		genera_html "d",retornaAgenteAduanal(act2.fields("27").value),"center"  'Agente Aduanal
		genera_html "d",act2.fields("28").value,"center"  'Patente Agente Aduanal
		ref = act2.fields("29").value
		if (ref <> refAux)then
			refAux=ref
			
		 'Lote 2
			genera_html "d",retornaCampoContenedores(act2.fields("29").value,"numcon40",mid(act2.fields("29").value,1,3)),"center"  'No de Contenedor 
			genera_html "d",act2.fields("NumeroBL").value,"center" 'Guia BL
			genera_html "d",act2.fields("cveped").value,"center"  'Clave Pedimento
			genera_html "d",act2.fields("31").value,"center"  'No. Pedimento
			genera_html "d",act2.fields("32").value,"center"  'Fecha Pedimento
			genera_html "d",act2.fields("33").value,"center"  'Mes
			genera_html "d",act2.fields("34").value,"center"  'No.Semana
			genera_html "d","1","center"  'Cantidad de Operaciones 
			genera_html "d",retornaCantContenedores40(act2.fields("29").value,"numcon40",mid(act2.fields("29").value,1,3)),"center"  'Cantidad de Contenedores
			genera_html "d",retornaCantContenedores(act2.fields("29").value,"'BUL','CAJ','BID','PAL'",mid(act2.fields("29").value,1,3)),"center"  'PALLETS/BULTOS 
			genera_html "d",retornaTipoContenedores(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  'TIPO DE CONTENEDOR/ CAJA
			genera_html "d",act2.fields("39").value,"center"  'Fecha Factura
			genera_html "d",act2.fields("40").value,"center"  'Fecha BL
			genera_html "d",act2.fields("41").value,"center"  'Fecha de arribo a la aduana
			genera_html "d",act2.fields("42").value,"center"  'Fecha Desaduanamiento
			genera_html "d","","center"  'EXC
			genera_html "d","","center"  'EXC $ M.N.
			genera_html "d","","center"  'EXC $ USD
			genera_html "d","","center"  'EXP
			genera_html "d","","center"  'EXP & USD ?
			genera_html "d","","center"  'EXC & EXPT LT $
			genera_html "d","","center"  'CAUSAL LT
			genera_html "d","","center"  'CAUSA RAIZ
			genera_html "d","","center"  'PLAN DE ACCION
			genera_html "d","","center"  'RESPONSABLE
			genera_html "d","","center"  'FECHA DE CUMPLIMIENTO
			genera_html "d","","center"  'IMPACTO
			genera_html "d",retornaCampoCtaGastos(act2.fields("29").value,"cgas31",mid(act2.fields("29").value,1,3)),"center"  'No. CTA DE GASTOS
			genera_html "d",act2.fields("fech31").value,"center" 'Fecha facturacion
			Subref = act2.fields("81").value
			if (ordenOcupado(Subref,act2.fields("29").value) = False)then
				ocuparOrd(Subref)
				genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
				genera_html "d",act2.fields("761").value,"center"  ' ADV FRACC. $ 
				genera_html "d",act2.fields("76").value,"center"  ' DTA $ 
				genera_html "d",act2.fields("77").value,"center"  'IVA %
				genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
			else
				genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
				genera_html "d","","center"  ' ADV FRACC. $ 
				genera_html "d",act2.fields("76").value,"center"  ' DTA $ 
				genera_html "d","","center"  'IVA %
				genera_html "d","","center"  ' IVA FRACC. $ 
			end if
			'/Lote2

			'Lote 1
			genera_html "d",act2.fields("79").value,"center"  ' PREVAL. 
			genera_html "d",act2.fields("eci").value,"center" 'ECI
			
			genera_html "d",sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  ' TOTAL IMPUESTOS 
			genera_html "d",(cdbl(sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3)))/act2.fields("65").value),"center"  ' Total Impuestos USD 
			genera_html "d",act2.fields("65").value,"center"  ' T.C. 
			genera_html "d",act2.fields("ValorAduana").value,"center" 'Valor aduana
			genera_html "d","N/A","center"  ' GTOS. ADUANA USD(SOLO FRONTERA) 
			IF (mid(act2.fields("29").value,1,3)="RKU" or mid(act2.fields("29").value,1,3)="SAP" or mid(act2.fields("29").value,1,3)="ALC") then
				genera_html "d",retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"DEMORAS"),"I",mid(act2.fields("29").value,1,3),"CA"),"center"  ' DEMORAS (Costo aduanal)
				genera_html "d",retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"I",mid(act2.fields("29").value,1,3),"CA"),"center"  ' ESTADIAS (Costo Aduanal)
				genera_html "d",retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3),"CA"),"center"  ' MANIOBRAS  (Costo Aduanal)
				genera_html "d",retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES"),"I",mid(act2.fields("29").value,1,3),"CA"),"center"  ' ALMACENAJES (Costo Aduanal)				
			else 
				genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"DEMORAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' DEMORAS (Costo aduanal) AEREO
				genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' ESTADIAS (Costo Aduanal) AEREO
				genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3)),"center"  ' MANIOBRAS  (Costo Aduanal) AEREO
					genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES"),"I",mid(act2.fields("29").value,1,3)),"center"  ' ALMACENAJES (Costo Aduanal) AEREO
			end if
			IF (mid(act2.fields("29").value,1,3)="RKU" or mid(act2.fields("29").value,1,3)="SAP" or mid(act2.fields("29").value,1,3)="ALC") then
				genera_html "d",retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"CONEXION REFRIGERADO"),"I",mid(act2.fields("29").value,1,3),"CA"),"center"  ' CONEXION DE REFRIGERADO  (Costo Aduanal)
			else
				if(mid(act2.fields("29").value,1,3)="TOL") THEN
					genera_html "d","0","center"
				else
					genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"CONEXION REFRIGERADO"),"I",mid(act2.fields("29").value,1,3)),"center"  ' CONEXION DE REFRIGERADO  (Costo Aduanal) Aereo
				end if
			end if
			IF (mid(act2.fields("29").value,1,3)="RKU" or mid(act2.fields("29").value,1,3)="SAP" or mid(act2.fields("29").value,1,3)="ALC") then
				genera_html "d",retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"CONEXION REFRIGERADO"),"I",mid(act2.fields("29").value,1,3),"CD"),"center"  ' CONEXION DE REFRIGERADO  (Costo Directo)
				genera_html "d",retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3),"CD"),"center"  ' MANIOBRA (Costo Directo)
			else
				genera_html "d","0","center"  ' CONEXION DE REFRIGERADO  (Costo Directo) Aereo
				genera_html "d","0","center"  ' MANIOBRA (Costo Directo) AEREO
			end if
			
			
			dim STIMP,STIVA,Demoras,Estadias,Almacenajes,Maniobras,CRefrigerado,Tarifa,TGD
			STIMP=0
			STIVA=0
			TGD=0
			Demoras=0
			Estadias=0
			Almacenajes=0
			Maniobras=0
			CRefrigerado=0
			Tarifa=0
				if(revisaImpuestosFacturados( act2.fields("29").value,"I",mid(act2.fields("29").value,1,3)) <> 0 )then
				STIMP=sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))
				STIVA=sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3))
			end if
			'Total de Gastos diversos
			IF (mid(act2.fields("29").value,1,3)="RKU" or mid(act2.fields("29").value,1,3)="SAP" or mid(act2.fields("29").value,1,3)="ALC") then
				Demoras=retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"DEMORAS"),"I",mid(act2.fields("29").value,1,3),"CA")' Retorna DEMORAS (Costo aduanal)
				Estadias= retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"I",mid(act2.fields("29").value,1,3),"CA")  ' ESTADIAS (Costo Aduanal)
				Maniobras=retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3),"CA")  ' MANIOBRAS  (Costo Aduanal)
				Almacenajes=retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES"),"I",mid(act2.fields("29").value,1,3),"CA")  ' ALMACENAJES (Costo Aduanal)				
				Tarifa=retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3),"CD")  ' MANIOBRA (Costo Directo)
			else	
				Demoras=retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"DEMORAS"),"I",mid(act2.fields("29").value,1,3))  ' DEMORAS (Costo aduanal) AEREO
				Estadias=retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"I",mid(act2.fields("29").value,1,3))  ' ESTADIAS (Costo Aduanal) AEREO
				Maniobras=retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3))  ' MANIOBRAS  (Costo Aduanal) AEREO
				Almacenajes=retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES"),"I",mid(act2.fields("29").value,1,3))  ' ALMACENAJES (Costo Aduanal) AEREO
			end if
			IF (mid(act2.fields("29").value,1,3)="RKU" or mid(act2.fields("29").value,1,3)="SAP" or mid(act2.fields("29").value,1,3)="ALC") then
				CRefrigerado=retornaPagosHechosMaritimo(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"CONEXION REFRIGERADO"),"I",mid(act2.fields("29").value,1,3),"CA")  ' CONEXION DE REFRIGERADO  (Costo Aduanal)
			else
				CRefrigerado=retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"CONEXION REFRIGERADO"),"I",mid(act2.fields("29").value,1,3))' CONEXION DE REFRIGERADO  (Costo Aduanal) Aereo
				IF CRefrigerado="" then 
					CRefrigerado=0
				end if
			end if
			
			contass=contass+1
			TGD=Demoras+Estadias+Maniobras+Almacenajes+Tarifa+CRefrigerado'-STIMP-STIVA
			on error resume next
				genera_html "d",TGD,"center"  ' TOTAL GASTOS DIVERSOS 
				genera_html "d",act2.fields("65").value,"center"  ' T.C. 
				genera_html "d",retornaHonorarios(act2.fields("29").value,"chon31",mid(act2.fields("29").value,1,3)),"center"  ' HONORARIOS AG AD. $ 
				genera_html "d",TGD/act2.fields("65").value,"center"  ' TOTAL GASTOS DIVERSOS USD 						
			if err.number <> 0 then
				genera_html "d","Error","center"  ' TOTAL GASTOS DIVERSOS 
				genera_html "d","Error","center"  ' TC
				genera_html "d","Error","center"  ' HONORARIOS AG AD. $
				genera_html "d","Error","center"  ' TOTAL GASTOS DIVERSOS USD
			end if
	
			
			'/Lote 1
		else
			'Lote 2
			genera_html "d","","center"  'No de Contenedor 
			genera_html "d",act2.fields("NumeroBL").value,"center" 'Guia BL
			genera_html "d",act2.fields("cveped").value,"center"  'Cve pedimento
			genera_html "d",act2.fields("31").value,"center"  'No. Pedimento
			genera_html "d",act2.fields("32").value,"center"  'Fecha Pedimento
			genera_html "d",act2.fields("33").value,"center"  'Mes
			genera_html "d",act2.fields("34").value,"center"  'No.Semana
			genera_html "d","0","center"  'Cantidad de Operaciones 
			genera_html "d","","center"  'Cantidad de Contenedores
			genera_html "d","","center"  'PALLETS/BULTOS 
			genera_html "d","","center"  'TIPO DE CONTENEDOR/ CAJA
			genera_html "d",act2.fields("39").value,"center"  'Fecha Factura
			genera_html "d",act2.fields("40").value,"center"  'Fecha BL
			genera_html "d",act2.fields("41").value,"center"  'Fecha de arribo a la aduana
			genera_html "d",act2.fields("42").value,"center"  'Fecha Desaduanamiento
			genera_html "d","","center"  'EXC
			genera_html "d","","center"  'EXC $ M.N.
			genera_html "d","","center"  'EXC $USD.
			genera_html "d","","center"  'EXP
			genera_html "d","","center"  'EXP & USD ?
			genera_html "d","","center"  'EXC & EXPT LT $
			genera_html "d","","center"  'CAUSAL LT
			genera_html "d","","center"  'CAUSA RAIZ
			genera_html "d","","center"  'PLAN DE ACCION
			genera_html "d","","center"  'REMPONSABLE
			genera_html "d","","center"  'FECHA DE CUMPLIMIENTO
			genera_html "d","","center"  'IMPACTO
			genera_html "d",retornaCampoCtaGastos(act2.fields("29").value,"cgas31",mid(act2.fields("29").value,1,3)),"center"  'No. CTA DE GASTOS
				genera_html "d",act2.fields("fech31").value,"center" 'Fecha facturacion
			
			Subref = act2.fields("81").value
			if (ordenOcupado(Subref,act2.fields("29").value) = False)then
				ocuparOrd(Subref)
				genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
				genera_html "d",act2.fields("761").value,"center"  ' ADV FRACC. $ 
				genera_html "d","","center"  ' DTA $ 
				genera_html "d",act2.fields("77").value,"center"  'IVA %
				genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
			else

				genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
			   genera_html "d","","center"  ' ADV FRACC. $ 
			   genera_html "d","","center"  ' DTA $ 
			   genera_html "d","","center"  'IVA %
			   genera_html "d","","center"  ' IVA FRACC. $ 
			end if
			'/Lote2
		 
			'Lote 1
			genera_html "d","","center"  ' PREVAL. 
			genera_html "d","","center"  'ECI
			genera_html "d","","center"  ' TOTAL IMPUESTOS 
			genera_html "d","","center"  ' Total Impuestos USD  
			genera_html "d",act2.fields("65").value,"center"  ' T.C. 
			genera_html "d","","center" 'Valor aduana
			genera_html "d","N/A","center"  ' GTOS. ADUANA USD(SOLO FRONTERA) 
			genera_html "d","","center"  ' DEMORAS (Costo Aduanal)
			genera_html "d","","center"  ' ESTADIAS (Costo Aduanal)
			genera_html "d","","center"  ' ALMACENAJES  (Costo Aduanal)
			genera_html "d","","center"  ' MANIOBRAS (Costo Aduanal) 
			genera_html "d","","center"  ' CONEXION DE REFRIGERADO (Costo aduanal)
			genera_html "d","","center"  ' CONEXION DE REFRIGERADO (Costo Directo)
			genera_html "d","","center"  ' MANIOBRAS (Costo Directo)
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS 
			genera_html "d","","center"  ' TC

			genera_html "d","","center"  ' HONORARIOS AG AD. $ 
			genera_html "d","","center"  ' TOTAL GASTOS DIVERSOS USD 
			
			'/Lote 1
		end if

		genera_html "d",act2.fields("Naviera").value,"center"  'Naviera 
		genera_html "d",retornaTransportista(mid(act2.fields("refcia01").value,1,3),act2.fields("refcia01").value),"center"  'Transportista Nacional
		genera_html "d",act2.fields("sDescTransp").value,"center"  'TIPO TRANSPORTE
		genera_html "d",retornaTipoUnidad(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3)),"center" 'TIPO UNIDAD
		
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

function retornaIMPORTADOR(clave,oficina)
ON ERROR RESUME NEXT
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

 sqlAct2 = "select  c.nomcli18 as campo from "&oficina&"_extranet.ssclie18 as c where c.cvecli18 = "&clave

 
Set act2= Server.CreateObject("ADODB.Recordset")
conn12 = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=10.66.1.9; DATABASE="&oficina&"_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"

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
 end if
 
'sqlAct="select "& campo &" as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"'  "
sqlAct="SELECT Distinct "& campo &" as campo FROM "&oficina&"_extranet.c01ptoemb where cvepto01 ="& val &" and nompto01 like '"& pto &"%'"

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

function retornaHonorarios(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
   if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
sqlAct=" select round(cta."&campo&",2) as campo from "&oficina&"_extranet.e31cgast as cta  " & _
       " inner join "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
       " where  r.refe31 = '"& referencia &"' and cta.esta31 = 'I' "

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
	elseif (ucase(oficina)="PAN")then
		oficina="DAI"
 end if
 
if oficina = "SAP" then

	if topico = "ALMACENAJES" then
		cad= "3,198,200,286,351"
	end if
	if topico = "DEMORAS" then
		cad= "6,14,46,63,129,156,352"
	end if
	if topico = "ESTADIAS" then
		cad="144,222,158"
	end if
	if topico = "MANIOBRAS" then
		cad="NA"
	end if
	if topico="CONEXION REFRIGERADO" THEN 
		cad="311"
	end if
else 
  if oficina = "CEG" then
  
    if topico = "ALMACENAJES" then
     cad= "2,4,59,77,100,223,223,235,235,241"
    end if
	if topico = "DEMORAS" then
	 cad= "11,48,99,150"
	end if
	if topico = "ESTADIAS" then
	 cad="NA"
	end if
	if topico = "MANIOBRAS" then
	 cad="NA" '"239"
    end if
  
  else 
     if oficina = "TOL" then
	 
	    if topico = "ALMACENAJES" then
		 cad= "10"
		end if
		if topico = "DEMORAS" then
		 cad= "NE"
		end if
		if topico = "ESTADIAS" then
		 cad="123"
		end if
		if topico = "MANIOBRAS" then
		 cad="141,82,102,14"
		end if
		if topico="CONEXION DE REFRIGERADO" then 
			cad="60"
		end if
		
     else 
	   if oficina = "LZR" then
	         if topico = "ALMACENAJES" then
			 cad= "4,119"
			end if
			if topico = "DEMORAS" then
			 cad= "11"
			end if
			if topico = "ESTADIAS" then
			 cad="77"
			end if
			if topico = "MANIOBRAS" then
			 cad="NA" '"2,78,115,125,203,244,297,312"
			end if
			if topico="CONEXION REFRIGERADO" then
				cad="135"
			end if
       else 
	       if oficina = "RKU" then
		          if topico = "ALMACENAJES" then
					  cad="4"
					end if
					if topico = "DEMORAS" then
					 cad= "11,376"
					end if
					if topico = "ESTADIAS" then
					 cad="77"
					end if
					if topico = "MANIOBRAS" then
					 cad="NA"'"2"
					end if
					if topico="CONEXION REFRIGERADO" then 	
						cad="135"
					end if
           else 
		       if oficina = "DAI" then
			           if topico = "ALMACENAJES" then
						 cad= "10"
						end if
						if topico = "DEMORAS" then
						 cad= "NE"
						end if
						if topico = "ESTADIAS" then
						 cad="NE"
						end if
						if topico = "MANIOBRAS" then
						 cad="1,127,82,102,35,6,11,3,141,63,315,219,257"
						end if
						if topico="CONEXION REFRIGERADO" then 	
						cad="60"
					end if
					else 
   			          cad = "NA"
               end if
           end if
       end if
     end if
  end if
end if
retornaConceptosPH = cad
end function

function retornaPagosHechosMaritimo(referencia,conceptos,tipope,oficina,tcosto)
dim c,valor,trafico,campo
	trafico=""
	c=chr(34)
	valor=0
	if (ucase(oficina) = "ALC")then
		oficina = "LZR"
	elseif (ucase(oficina)="PAN") then
		oficina="DAI"
	end if
	
	if (oficina="ALC" or oficina="RKU" or oficina="SAP") then
		trafico="MARITIMO"
	else 
		trafico="AEREO"
	end if
	campo="Importe"
if conceptos<>"NE" then 	
	if conceptos="NA" and tcosto="CA" then 'Esto indica que es Maniobra (Costo Aduanal)
		conceptos=retornaConceptosPH(oficina,"DEMORAS")&","&retornaConceptosPH(oficina,"ESTADIAS")&","&retornaConceptosPH(oficina,"ALMACENAJES")&","&retornaConceptosPH(oficina,"CONEXION REFRIGERADO")
		conceptos= " and ep.conc21 not in("&conceptos&") "
		tcosto=" and ep.tpag21 in(2,1) "
	
	elseif conceptos="NA" and tcosto="CD" then 'Esto indica que es Maniobra (Costo directo)
		conceptos=" "
		tcosto="  "
		campo="TFlat"
	else ' Es cualquier otro concepto como costo aduanal, costo directo o pago hecho 
		conceptos=" and ep.conc21 in ("&conceptos&") "
		if tcosto="CD" then 
			tcosto=" and ep.tpag21=3 "
		elseif tcosto="CA" then
			tcosto=" and ep.tpag21 in(2,1) "
		end if
	end if

	sqlAct="select i.refcia01 as Ref, r.cgas31,ep.conc21,ep.piva21,ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as Importe, cp.desc21,cta.csce31 as TFlat " 
	sqlAct=sqlAct& " from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " 
	sqlAct=sqlAct& "  inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  " 
	sqlAct=sqlAct& "     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 "
	sqlAct=sqlAct& "          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 " 
	sqlAct=sqlAct& "             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  and ep.tmov21 =dp.tmov21 " 
	sqlAct=sqlAct& "                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " 
	sqlAct=sqlAct& "    where  i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3','UMA011214255')  and i.firmae01 <> ''  and cta.esta31 <> 'C'  "
	sqlAct=sqlAct&  "and i.refcia01 = '"&referencia&"' "
	sqlAct=sqlAct& conceptos & tcosto 
	sqlAct=sqlAct&  " group by Ref,cgas31"

'response.write(sqlAct)
'response.end()
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
 valor = act2.fields(campo).value

  retornaPagosHechosMaritimo= valor
 else
  retornaPagosHechosMaritimo = valor
 end if
else
	retornaPagosHechosMaritimo=valor
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

sqlAct="select i.refcia01 as Ref, r.cgas31,ep.conc21,ep.piva21,ifnull(sum(dp.mont21*if(ep.deha21 = 'C',-1,1)),0) as Importe, cp.desc21 " & _
" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
"  inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  " & _
"     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 " & _
"          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 " & _
"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S'  and ep.tmov21 =dp.tmov21 " & _
"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
"    where  i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3','UMA011214255')  and i.firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in ("&conceptos&") and i.refcia01 = '"&referencia&"' "&_
"  group by Ref"

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
		" where i.rfccli01 in ('UME651115N48','BRM711115GI8','ISI011214HM3','UMA011214255')  and i.refcia01 = '"&referencia&"'  and i.firmae01 <> '' group by i.refcia01 "


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
   retornaTOTALPagosHechos =0
 end if

end function

function retornaCampoContenedores(referencia,campo,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select Distinct "& campo &" as campo from "&oficina&"_extranet.sscont40 where refcia40 = '"&referencia&"'  "

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
 end if
 
sqlAct="select count(Distinct numcon40) as campo from  "&oficina&"_extranet.sscont40 where refcia40 = '"&referencia&"'  "

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
val="40 HighCube Refrigerated"
end if
DatosContenedor = val
end function

function retornaTipoContenedores(referencia,oficina)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
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
	     fraccion = "33029099" or _
		 fraccion="33053001") then
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
	 fraccion = "09023001" or  _
	 fraccion="19019099") then
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

if clave = "11000" or clave="15001" then
  val = revisaFraccion(fraccion) '"Centro de Distribución"
end if
if clave = "11001" then
val = "ICE CREAM" '"Planta Helados"
end if
if clave = "11002" or clave="15002" or clave="15003" then
val = "FOODS"
end if
if clave = "11003" or clave="15004"  then
val = "HPC" '"Planta HPC"
end if
if clave = "11004" then
val = "Todas"
end if
'OTRAS
if clave = "13000" then
val = "Todas"
end if
if clave = "14000" or clave="14015" or clave="15005" then
val = "Todas"
end if

retornaDivision = val
end function


function retornaPlantaEntrega(clave)
dim val
val= ""
if clave = "11000" or clave="15001" then
val = "CDU"
end if
if clave = "11001" or clave="15002" then
val = "TULTITLAN"
end if
if clave = "11002" or clave="15003" then
val = "LERMA"
end if
if clave = "11003" or clave="15004"then
val = "CIVAC"
end if
if clave = "11004" then
val = "ESPECIALES"
end if

'OTRAS
if clave = "13000" or clave="14000" or clave="15005" then
val = "Todas"
end if


retornaPlantaEntrega = val
end function

function sumaTotalImpuestos(referencia,oficina)
dim c,valor
 c=chr(34)
 valor=0
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
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
   valor = act2.fields("campo").value
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
 end if
if oficina<>"TOL" or oficina<>"DAI" then
condicion="  AND (if(con.numcon40 is null,true,con.numcon40 = REPLACE(REPLACE(d01.marc01, '/', ''), '-', ''))) "
else 
condicion=""
end if
sqlAct=" SELECT tracto01 as campo from "& oficina &"_extranet.ssdag"&tipope&"01 AS i LEFT JOIN "& oficina &"_extranet.sscont40 " & _
		"AS con ON con.refcia40 =i.refcia01 AND con.patent40 = i.patent01 AND con.adusec40 = i.adusec01 LEFT JOIN" & _
		" "& oficina &"_extranet.d01conte AS d01 ON d01.refe01 =i.refcia01 "&condicion & _
		" where  i.refcia01 = '"& referencia &"' and d01.peri01<>0 "
		
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
				if valor="S" or valor="F" then
					if valor= "S" then 
						unidad = "SENCILLO"
					end if
					if valor ="F" then
						unidad = "FULL"
					end if 
			
				end if
	retornaTipoUnidad = unidad
 end if
 		if(oficina="TOL" or oficina="DAI") then
						unidad="SENCILLO"
		end if
 		retornaTipoUnidad = unidad
end function

function retornaTransportista(oficina,referencia)
if oficina="ALC" then oficina="LZR"
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
 function QueryMySql(Oficina)
 
  sqlAct= "select "&_
 "i.refcia01, "&_
 "i.cveped01 'cveped', "&_
 "i.pesobr01 'pesobr', "&_
 "fr.fraarn02, "&_
 "fr.ordfra02, "&_
 "i.cvecli01 as '1',  "&_
 "if(i.cvecli01 in (11000,15001),'Virginia Leon', if(i.cvecli01 in(11001,11004,15003,14015),'Montserrat Rodriguez',if(i.cvecli01 =11002,'Iray Hinojosa', if(i.cvecli01 in(11003,15004),'Francisco Bernal',if(i.cvecli01 in(15002,15005),'Jazmin Osornio',if(i.cvecli01='12002','Jorge Islas', if(i.cvecli01='12000','Georgina Perez', cast( i.cvecli01 as char)))) )))) Ejecutivo,  "&_
 "ar.desc05 as '4', "&_
 "prv.nompro22 as '9',  "&_ 
 "r.rcli01 as '11',  "&_
 "ar.pedi05 as '12',  "&_
 "i.cvepod01 as '14',  "&_
 "i.cvepvc01 as '15',   "&_
 "r.ptoemb01 as '177',  "&_
 "r.cveptoemb as '17',  "&_
 "prv.irspro22 as '18',  "&_
 "f.numfac39 as '19',  "&_
 "date_format(f.fecfac39,'%d/%m/%Y') as '20',  "&_
 "r.impo01 as '21',  "&_
 "ar.caco05 as '22',  "&_
 "um.descri31 as '23', "&_
 "f.terfac39 as '24', "&_
 "IF(ar.frac05 in('39239099','39173299'),'PACK', if(ar.tpmerc05='PM','ROH',if(ar.tpmerc05='PT','FERT',if(ar.tpmerc05='R','REF',if(ar.tpmerc05='MU','MUESTRA',IF(ar.tpmerc05='PA','PARTES',ar.tpmerc05)))))) TipMerc, "&_
 "max(cta.fech31) fech31, "&_
 "imp.eci,"&_
 " group_concat(distinct g.numgui04) NumeroBL, "&_
  "(select sum(va.vaduan02) as campo from "&Oficina&"_extranet.ssfrac02 as va where va.refcia02 =i.refcia01) as ValorAduana , "&_
 "(select distinct if(rec.refcia06 is null,'N','S') from "&Oficina&"_extranet.ssdagi01 as i2 left join "&Oficina&"_extranet.ssrecp06 as rec on rec.reforg06 =i2.refcia01 where i2.refcia01=i.refcia01 and i2.firmae01<>'' and i2.firmae01 is not null) as Rectificacion, "&_
 "(select if(ip.cveide11='RO','Si','')from "&Oficina&"_extranet.ssiped11 as ip where ip.refcia11=i.refcia01 and ip.patent11=i.patent01 and ip.cveide11='RO') RO,"
 if Oficina="rku" or Oficina="lzr" or Oficina="sap" then
 sqlAct=sqlAct& "'MARITIMO' as '25', "
 else
 sqlAct=sqlAct& "'AEREO' as '25', "
 end if 
  sqlAct=sqlAct&"i.adusec01 as '26',  "&_
 "i.patent01 as '27',  "&_
 "i.patent01 as '28',  "&_
 "i.refcia01 as '29',  "&_
 "i.numped01 as '31',  "&_
 "i.fecpag01 as '32',  "&_
 "Month(i.fecpag01) as '33', "&_
 "weekofyear(i.fecpag01) as '34',  "&_
 "'N/A' as '35',  "&_
 "date_format(f.fecfac39,'%d/%m/%Y') as '39',  "&_
 "r.feorig01 as '40',  "&_
 "r.frev01 as '41', "&_
 "r.fdsp01 as '42',  "&_
 "fr.prepag02 as '59',  "&_
 "(fr.prepag02/i.tipcam01) as '60',  "&_
 "i.fletes01 as '61',  "&_
 "i.segros01 as '62',  "&_
 "i.incble01 as '63',  "&_
 "fr.vaduan02 as '64',  "&_
 "i.tipcam01 as '65',  "&_
 "(fr.vaduan02/i.tipcam01) as '66',  "&_
 "(i.fletes01/i.tipcam01) as '69',  "&_
 "fr.fraarn02 as '71', fr.tasadv02 as '72', if(ipar2.cveide12 ='TL',concat(concat(ipar2.cveide12,'-'),ipar2.comide12) ,ifnull(ipar2.cveide12,'TG')) as '73', "&_
 "cf6.import36 as '75', cf1.import36 as '76', (fr.i_adv102+fr.i_adv202) as '761', fr.tasiva02 as '77', cf3.import36 as '78', (fr.i_iva102+fr.i_iva202) as '781', "&_ 
 "cf15.import36 as '79',  fr.ordfra02 as '81', count(fr.ordfra02) as '82', ar.item05 as 'Item05' , "&_
 "transp.descri30 as 'sDescTransp' , "&_
 " 	if((mid(i.refcia01 ,1,3)='DAI' or (mid(i.refcia01 ,1,3)='PAN')),if(lin.desc01<>'',lin.desc01,lin.dir01) ,nav.nom01)  AS Naviera  "&_
 "from "&Oficina&"_extranet.ssdag"&tipope&"01 as i  "&_
 " left join "&Oficina&"_extranet.ssguia04 as g on g.refcia04=i.refcia01 and g.patent04=i.patent01 and g.adusec04=i.adusec01 and idngui04 =1 "&_
 "left join "&Oficina&"_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  "&_
 "left join "&Oficina&"_extranet.c01refer as r on r.refe01 = i.refcia01  "&_
 "left join "&Oficina&"_extranet.c06barco as bar on bar.clav06 =r.cbuq01  "&_
 "left join "&Oficina&"_extranet.c55navie as nav on nav.cve01 =bar.navi06  and nav.Status55='T' "&_
 "left join "&Oficina&"_extranet.c01airln as lin on lin.cvela01 =r.cvela01 "&_
 "INNER join "&Oficina&"_extranet.d31refer as ctar on ctar.refe31 = i.refcia01 "&_ 
 "INNER join "&Oficina&"_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31 and cta.esta31 <> 'C'  "&_
 "left join "&Oficina&"_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and f.adusec39 = i.adusec01 and f.patent39 = i.patent01  "&_
 "left join "&Oficina&"_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05 and ar.refe05 = r.refe01 "&_
 "left join "&Oficina&"_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02 and ar.frac05 = fr.fraarn02 and fr.ordfra02 = ar.agru05  "&_
 "left join "&Oficina&"_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01  "&_
 "left join "&Oficina&"_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02  "&_
 "left join "&Oficina&"_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'  "&_
 "left join "&Oficina&"_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'  "&_
 "left join "&Oficina&"_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'  "&_
 "left join "&Oficina&"_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'  "&_
 "left join "&Oficina&"_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL')  "&_
 "left join "&Oficina&"_extranet.ssmtra30 as transp on transp.clavet30 = i.cvemts01 "&_
  "left join ( 	select ii.refcia01 "&_
		",sum(if( ii.cveped01<>'R1',if( c36.cveimp36 = 1, c36.import36, 0),if(c33.cveimp33=1,c33.import33 ,0))) 'dta' "&_
		", sum(if( ii.cveped01<>'R1',if( c36.cveimp36 = 3, c36.import36,0),if(c33.cveimp33=3,c33.import33 ,0)))  'iva' "&_
		", sum(if( ii.cveped01<>'R1',if( c36.cveimp36 = 15 , c36.import36,0),if(c33.cveimp33=15,c33.import33,0))) 'prv' "&_
		", sum(if( ii.cveped01<>'R1',if( c36.cveimp36 = 6 , c36.import36,0),if(c33.cveimp33=6,c33.import33 ,0))) 'igi' "&_
		", sum(if( ii.cveped01<>'R1',if( c36.cveimp36=18, c36.import36 ,0),if(c33.cveimp33=18,c33.import33 ,0))) 'eci' "&_
		", sum(if( ii.cveped01<>'R1',if( c36.cveimp36 in(1,6,15),c36.import36,0),if(c33.cveimp33 in (1,6,15),c33.import33,0))) 'TotalImp' "&_
		"from "&Oficina&"_extranet.ssdagi01 as ii "&_
		"left join "&Oficina&"_extranet.c01refer as rr  on rr.refe01=ii.refcia01 "&_
		"left join "&Oficina&"_extranet.sscont36 as c36 on c36.refcia36=ii.refcia01 and c36.patent36=ii.patent01 "&_
		"left join "&Oficina&"_extranet.sscont33 as c33 on c33.refcia33=ii.refcia01 and c33.patent33=ii.patent01 "&_
		"where ii.firmae01 is not null and ii.firmae01<>'' and ii.rfccli01 in('UME651115N48','BRM711115GI8','ISI011214HM3','UMA011214255') "&_
		"and rr.fdsp01 >='"&finicio&"' and rr.fdsp01<='"&ffinal&"' "&_
		"group by ii.refcia01 "&_
		") as imp on imp.refcia01=i.refcia01 "&_
 "where cc.rfccli18 in ('UME651115N48','BRM711115GI8','ISI011214HM3','UMA011214255') "&_
 "and i.cveped01 <> 'G1' and i.firmae01 is not null and i.firmae01 <> ''  "&_
 "and cta.fech31 >= '"&finicio&"' and cta.fech31 <= '"&ffinal&"' "&_
 "group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " 
'response.write(sqlAct)
'response.end()
	QueryMySql=sqlAct
end function	

%>

