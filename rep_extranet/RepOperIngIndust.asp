<!-- #include virtual="/PortalMySQL/Extranet/ext-Asp/Clases/cConexion.asp" -->

<META HTTP-EQUIV="Content-Type" CONTENT="text/html"; charset="utf-8">
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 11">
<%
Response.Buffer = TRUE
response.Charset = "utf-8"
'Response.Addheader "Content-Disposition", "attachment; filename=BookletIFF.xls"
'Response.ContentType = "application/vnd.ms-excel"

dim oficina,cvesoficina,validacion
oficina="RKU"
cvesoficina=""
validacion=""


 strTipoUsuario = Session("GTipoUsuario")
 fechaini = trim(request.Form("txtDateIni"))
 fechafin = trim(request.Form("txtDateFin"))
 strTipoOperaciones = request.Form("rbnTipoDate")

oficina=Request.QueryString("ofi")
tipope=Request.QueryString("tipope")
det=Request.QueryString("det")
mes=Request.QueryString("mes")
fechaini=Request.QueryString("finicio")
fechafin=Request.QueryString("ffinal")



 if not fechaini="" and not fechafin="" then


    tmpDiaIni = cstr(datepart("d",fechaini))
    tmpMesIni = cstr(datepart("m",fechaini))
    tmpAnioIni = cstr(datepart("yyyy",fechaini))
    finicio = tmpAnioIni & "-" &tmpMesIni & "-"& tmpDiaIni

    tmpDiaFin = cstr(datepart("d",fechafin))
    tmpMesFin = cstr(datepart("m",fechafin))
    tmpAnioFin = cstr(datepart("yyyy",fechafin))
    ffinal = tmpAnioFin & "-" &tmpMesFin & "-"& tmpDiaFin

	oficina_adu=GAduana
	jnxadu=Session("GAduana")
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
		case "TOL"
			strOficina="tol"
	end select
Response.write(strOficina)&"aki"

dim orden(300)
dim subrefaux,subref,bgcolor
subrefaux=""
subref=""
bgcolor="#FFFFFF"


dim strHTML 
strHTML = ""

Server.ScriptTimeOut=10000000
%>
<title> Reporte1.. </title>
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
	text-align:left;}
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
	text-align:left;
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
	text-align:left;
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
	text-align:left;
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
	text-align:left;
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
	text-align:left;
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
	text-align:left;
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
	text-align:left;
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
	text-align:left;
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
<!-- <table width="982" border="2" align="center" cellpadding="0" cellspacing="0" bordercolor="#C1C1C1"> 
<tr align="center" bordercolor="#999999" bgcolor="#CCFF99"> -->

<table x:str border=0 cellpadding=0 cellspacing=0 width=12637 style='border-collapse:
 collapse;table-layout:fixed;width:9479pt'>
 <col width=125 style='mso-width-source:userset;mso-width-alt:4571;width:94pt'>
 <col width=100 span=2 style='mso-width-source:userset;mso-width-alt:3657;
 width:75pt'>
 <col class=xl26 width=226 style='mso-width-source:userset;mso-width-alt:8265;
 width:170pt'>
 <col width=100 span=4 style='mso-width-source:userset;mso-width-alt:3657;
 width:75pt'>
 <col width=212 style='mso-width-source:userset;mso-width-alt:7753;width:159pt'>
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
<% genera_registros det,tipope %>
</table>
</body>
</html>
<%


end if


sub genera_registros(det,tipope)
dim c
 c=chr(34)




%>
 <!-- <td height=52 class=xl28 id="_x0000_s1025" x:autofilter="all"
  x:autofilterrange="$A$1:$CS$800" width=125 style='height:39.0pt;width:94pt'
  x:str="DIVISION ">DIVISION<span style='mso-spacerun:yes'> </span></td>
  <td class=xl29 id="_x0000_s1026" x:autofilter="all" width=100
  style='width:75pt'>IMPORTANCIA</td>
  <td class=xl29 id="_x0000_s1027" x:autofilter="all" width=100
  style='width:75pt'>CATEGORIA</td>
  <td class=xl29 id="_x0000_s1028" x:autofilter="all" width=226
  style='width:170pt'>Nombre del Material</td>
  <td class=xl30 id="_x0000_s1029" x:autofilter="all" width=100
  style='width:75pt'>Codigo SAP</td>
  <td class=xl28 id="_x0000_s1030" x:autofilter="all" width=100
  style='width:75pt'>Clase de Producto</td>
  <td class=xl29 id="_x0000_s1031" x:autofilter="all" width=100
  style='width:75pt'>Contrato Marco</td>
  <td class=xl29 id="_x0000_s1032" x:autofilter="all" width=100
  style='width:75pt'>Código SAP Proveedor</td>
  <td class=xl31 id="_x0000_s1033" x:autofilter="all" width=212
  style='width:159pt'>Proveedor</td>
  <td class=xl28 id="_x0000_s1034" x:autofilter="all" width=100
  style='width:75pt' x:str="Cuenta ">Cuenta<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl28 id="_x0000_s1035" x:autofilter="all" width=100
  style='width:75pt'>CECO</td>
  <td class=xl32 id="_x0000_s1036" x:autofilter="all" width=100
  style='width:75pt'>ODC</td>
  <td class=xl28 id="_x0000_s1037" x:autofilter="all" width=100
  style='width:75pt'>No IE</td>
  <td class=xl33 id="_x0000_s1038" x:autofilter="all" width=100
  style='width:75pt' x:str="País de Origen ">País de Origen<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl33 id="_x0000_s1039" x:autofilter="all" width=100
  style='width:75pt'>País de Procedencia</td>
  <td class=xl33 id="_x0000_s1040" x:autofilter="all" width=100
  style='width:75pt'>Region</td>
  <td class=xl34 id="_x0000_s1041" x:autofilter="all" width=100
  style='width:75pt'>PTO./CD DE ORIGEN</td>
  <td class=xl35 id="_x0000_s1042" x:autofilter="all" width=100
  style='width:75pt'>TAX ID/ RFC</td>
  <td class=xl36 id="_x0000_s1043" x:autofilter="all" width=100
  style='width:75pt'>Factura</td>
  <td class=xl37 id="_x0000_s1044" x:autofilter="all" width=100
  style='width:75pt'>Fecha de Factura</td>
  <td class=xl38 id="_x0000_s1045" x:autofilter="all" width=214
  style='width:161pt'>IMPORTADOR</td>
  <td class=xl39 id="_x0000_s1046" x:autofilter="all" width=100
  style='width:75pt' x:str="Cantidad ">Cantidad<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl33 id="_x0000_s1047" x:autofilter="all" width=100
  style='width:75pt'>Unidad de Medida</td>
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
  <td class=xl33 id="_x0000_s1053" x:autofilter="all" width=100
  style='width:75pt'>No. De Trafico</td>
  <td class=xl38 id="_x0000_s1054" x:autofilter="all" width=100
  style='width:75pt' x:str="No de Contenedor ">No de Contenedor<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl33 id="_x0000_s1055" x:autofilter="all" width=100
  style='width:75pt'>No. Pedimento</td>
  <td class=xl33 id="_x0000_s1056" x:autofilter="all" width=100
  style='width:75pt'>Fecha Pedimento</td>
  <td class=xl40 id="_x0000_s1057" x:autofilter="all" width=100
  style='width:75pt'>Mes</td>
  <td class=xl38 id="_x0000_s1058" x:autofilter="all" width=100
  style='width:75pt'>No.Semana</td>
  <td class=xl41 id="_x0000_s1059" x:autofilter="all" width=100
  style='width:75pt' x:str="Cantidad de Operaciones ">Cantidad de
  Operaciones<span style='mso-spacerun:yes'> </span></td>
  <td class=xl41 id="_x0000_s1060" x:autofilter="all" width=100
  style='width:75pt'>Cantidad de Contenedores</td>
  <td class=xl41 id="_x0000_s1061" x:autofilter="all" width=100
  style='width:75pt' x:str="PALLETS/BULTOS ">PALLETS/BULTOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl42 id="_x0000_s1062" x:autofilter="all" width=100
  style='width:75pt'>TIPO DE CONTENEDOR/ CAJA</td>
  <td class=xl41 id="_x0000_s1063" x:autofilter="all" width=100
  style='width:75pt'>Fecha Factura</td>
  <td class=xl41 id="_x0000_s1064" x:autofilter="all" width=100
  style='width:75pt'>Fecha BL</td>
  <td class=xl43 id="_x0000_s1065" x:autofilter="all" width=100
  style='width:75pt'>Fecha de arribo a la aduana</td>
  <td class=xl43 id="_x0000_s1066" x:autofilter="all" width=100
  style='width:75pt'>Fecha Desaduanamiento</td>
  <td class=xl44 id="_x0000_s1067" x:autofilter="all" width=100
  style='width:75pt'>KPI Desaduanamiento</td>
  <td class=xl44 id="_x0000_s1068" x:autofilter="all" width=100
  style='width:75pt'>KPI lead TIME</td>
  <td class=xl45 id="_x0000_s1069" x:autofilter="all" width=100
  style='width:75pt'>TARGET TIME</td>
  <td class=xl44 id="_x0000_s1070" x:autofilter="all" width=100
  style='width:75pt'>Fecha Arribo Planta</td>
  <td class=xl46 id="_x0000_s1071" x:autofilter="all" width=100
  style='width:75pt' x:str="BW ? ">BW ?<span style='mso-spacerun:yes'> </span></td>
  <td class=xl46 id="_x0000_s1072" x:autofilter="all" width=100
  style='width:75pt' x:str="BW $ ">BW $<span style='mso-spacerun:yes'> </span></td>
  <td class=xl47 id="_x0000_s1073" x:autofilter="all" width=100
  style='width:75pt'>PL</td>
  <td class=xl48 id="_x0000_s1074" x:autofilter="all" width=100
  style='width:75pt'>TT</td>
  <td class=xl49 id="_x0000_s1075" x:autofilter="all" width=100
  style='width:75pt'>PR</td>
  <td class=xl50 id="_x0000_s1076" x:autofilter="all" width=100
  style='width:75pt'>AA</td>
  <td class=xl51 id="_x0000_s1077" x:autofilter="all" width=100
  style='width:75pt'>CO</td>
  <td class=xl47 id="_x0000_s1078" x:autofilter="all" width=100
  style='width:75pt'>AL</td>
  <td class=xl52 id="_x0000_s1079" x:autofilter="all" width=100
  style='width:75pt'>NUMERO DE EMBARQUE</td>
  <td class=xl53 id="_x0000_s1080" x:autofilter="all" width=100
  style='width:75pt' x:str="IDOT ">IDOT<span style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1081" x:autofilter="all" width=100
  style='width:75pt' x:str="No. CTA DE GASTOS"><span
  style='mso-spacerun:yes'> </span>No. CTA DE GASTOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1082" x:autofilter="all" width=100
  style='width:75pt' x:str="Monto de Anticipo"><span
  style='mso-spacerun:yes'> </span>Monto de Anticipo<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1083" x:autofilter="all" width=100
  style='width:75pt' x:str="Precio Pagado / valor comercial"><span
  style='mso-spacerun:yes'> </span>Precio Pagado / valor comercial<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1084" x:autofilter="all" width=100
  style='width:75pt' x:str=" Valor comercial USD"><span
  style='mso-spacerun:yes'>  </span>Valor comercial USD<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1085" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR FLETES INTERNACIONAL M.N."><span
  style='mso-spacerun:yes'> </span>VALOR FLETES INTERNACIONAL M.N.<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1086" x:autofilter="all" width=100
  style='width:75pt' x:str="SEGUROS"><span
  style='mso-spacerun:yes'> </span>SEGUROS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1087" x:autofilter="all" width=100
  style='width:75pt' x:str="OTROS INCREMENTABLES"><span
  style='mso-spacerun:yes'> </span>OTROS INCREMENTABLES<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1088" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR ADUANA M.N."><span
  style='mso-spacerun:yes'> </span>VALOR ADUANA M.N.<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1089" x:autofilter="all" width=100
  style='width:75pt' x:str="T.C."><span
  style='mso-spacerun:yes'> </span>T.C.<span style='mso-spacerun:yes'> </span></td>
  <td class=xl54 id="_x0000_s1090" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR ADUANA  DLLS"><span
  style='mso-spacerun:yes'> </span>VALOR ADUANA<span style='mso-spacerun:yes'> 
  </span>DLLS<span style='mso-spacerun:yes'> </span></td>
  <td class=xl55 id="_x0000_s1091" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR FLETES AEREO DLLS"><span
  style='mso-spacerun:yes'> </span>VALOR FLETES AEREO DLLS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl55 id="_x0000_s1092" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR FLETES TERRESTRE DLLS."><span
  style='mso-spacerun:yes'> </span>VALOR FLETES TERRESTRE DLLS.<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl55 id="_x0000_s1093" x:autofilter="all" width=100
  style='width:75pt' x:str="VALOR FLETES MARITIMO DLLS."><span
  style='mso-spacerun:yes'> </span>VALOR FLETES MARITIMO DLLS.<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl55 id="_x0000_s1094" x:autofilter="all" width=100
  style='width:75pt' x:str="SAVING"><span
  style='mso-spacerun:yes'> </span>SAVING<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl56 id="_x0000_s1095" x:autofilter="all" width=100
  style='width:75pt' x:str="FRACC. ARANC."><span
  style='mso-spacerun:yes'> </span>FRACC. ARANC.<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl57 id="_x0000_s1096" x:autofilter="all" width=100
  style='width:75pt'>ARANCEL %</td>
  <td class=xl57 id="_x0000_s1097" x:autofilter="all" width=100
  style='width:75pt' x:str="ARANCEL PREFERENCIAL ">ARANCEL PREFERENCIAL<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl58 id="_x0000_s1098" x:autofilter="all" width=100
  style='width:75pt' x:str="MONTO DE RECUPERACION $ "><span
  style='mso-spacerun:yes'> </span>MONTO DE RECUPERACION $<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl58 id="_x0000_s1099" x:autofilter="all" width=100
  style='width:75pt' x:str="ADV. $ / IGI $"><span
  style='mso-spacerun:yes'> </span>ADV. $ / IGI $<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl58 id="_x0000_s1100" x:autofilter="all" width=100
  style='width:75pt' x:str="DTA $"><span style='mso-spacerun:yes'> </span>DTA
  $<span style='mso-spacerun:yes'> </span></td>
  <td class=xl57 id="_x0000_s1101" x:autofilter="all" width=100
  style='width:75pt'>IVA %</td>
  <td class=xl58 id="_x0000_s1102" x:autofilter="all" width=100
  style='width:75pt' x:str="IVA $"><span style='mso-spacerun:yes'> </span>IVA
  $<span style='mso-spacerun:yes'> </span></td>
  <td class=xl58 id="_x0000_s1103" x:autofilter="all" width=100
  style='width:75pt' x:str="PREVAL. "><span
  style='mso-spacerun:yes'> </span>PREVAL.<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl58 id="_x0000_s1104" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL IMPUESTOS"><span
  style='mso-spacerun:yes'> </span>TOTAL IMPUESTOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl59 id="_x0000_s1105" x:autofilter="all" width=100
  style='width:75pt' x:str="Total Impuestos USD "><span
  style='mso-spacerun:yes'> </span>Total Impuestos USD<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl60 id="_x0000_s1106" x:autofilter="all" width=100
  style='width:75pt' x:str="GTOS. ADUANA USD(SOLO FRONTERA)"><span
  style='mso-spacerun:yes'> </span>GTOS. ADUANA USD(SOLO FRONTERA)<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1107" x:autofilter="all" width=100
  style='width:75pt' x:str="DEMORAS"><span
  style='mso-spacerun:yes'> </span>DEMORAS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1108" x:autofilter="all" width=100
  style='width:75pt' x:str="ESTADIAS"><span
  style='mso-spacerun:yes'> </span>ESTADIAS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1109" x:autofilter="all" width=100
  style='width:75pt' x:str="MANIOBRAS "><span
  style='mso-spacerun:yes'> </span>MANIOBRAS<span
  style='mso-spacerun:yes'>  </span></td>
  <td class=xl61 id="_x0000_s1110" x:autofilter="all" width=100
  style='width:75pt' x:str="ALMACENAJES"><span
  style='mso-spacerun:yes'> </span>ALMACENAJES<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1111" x:autofilter="all" width=100
  style='width:75pt' x:str="OTROS"><span
  style='mso-spacerun:yes'> </span>OTROS<span style='mso-spacerun:yes'> </span></td>
  <td class=xl61 id="_x0000_s1112" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS"><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl59 id="_x0000_s1113" x:autofilter="all" width=100
  style='width:75pt' x:str="TOTAL GASTOS DIVERSOS USD"><span
  style='mso-spacerun:yes'> </span>TOTAL GASTOS DIVERSOS USD<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl62 id="_x0000_s1114" x:autofilter="all" width=100
  style='width:75pt' x:str="HONORARIOS AG AD. $"><span
  style='mso-spacerun:yes'> </span>HONORARIOS AG AD. $<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl59 id="_x0000_s1115" x:autofilter="all" width=100
  style='width:75pt' x:str="Total Gastos Indirectos  USD"><span
  style='mso-spacerun:yes'> </span>Total Gastos Indirectos<span
  style='mso-spacerun:yes'>  </span>USD<span style='mso-spacerun:yes'> </span></td>
  <td class=xl63 id="_x0000_s1116" x:autofilter="all" width=100
  style='width:75pt'>INLAND</td>
  <td class=xl64 id="_x0000_s1117" x:autofilter="all" width=100
  style='width:75pt' x:str="Impacto Valor Factura. ">Impacto Valor
  Factura.<span style='mso-spacerun:yes'> </span></td>
  <td class=xl65 id="_x0000_s1118" x:autofilter="all" width=100
  style='width:75pt' x:str="VAL FACT USD"><span
  style='mso-spacerun:yes'> </span>VAL FACT USD<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl66 id="_x0000_s1119" x:autofilter="all" width=100
  style='width:75pt'>PLANTA DE ENTREGA</td> -->
  
 <!--  <td class=xl66 id="_x0000_s1120" x:autofilter="all" width=100
  style='width:75pt'>ORDEN FRACCION</td>
  <td class=xl66 id="_x0000_s1121" x:autofilter="all" width=100
  style='width:75pt'>CUENTA ORDEN FRACCION</td> -->


<%
%></tr><%
 

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
" ifnull(r.cveptoemb,0) as '17', " & _
" prv.irspro22 as '18',  " & _
" f.numfac39 as '19',  " & _
" date_format(f.fecfac39,'%d/%m/%Y') as '20',  " & _
" ifnull(r.impo01,0) as '21',  " & _
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
"   from " & strOficina & "_extranet.ssdagi01 as i  " & _
"   left join " & strOficina & "_extranet.ssclie18 as cc on cc.cvecli18 = i.cvecli01  " & _
"   left join " & strOficina & "_extranet.c01refer as r on r.refe01 = i.refcia01 " & _
" left join " & strOficina & "_extranet.d31refer as ctar on  ctar.refe31 = i.refcia01 " & _
" left join " & strOficina & "_extranet.e31cgast as cta on cta.cgas31 = ctar.cgas31  and cta.esta31 <> 'C' " & _
"       left join " & strOficina & "_extranet.ssfact39 as f on f.refcia39 = i.refcia01 and  f.adusec39 = i.adusec01 and f.patent39 = i.patent01  " & _
"         left join " & strOficina & "_extranet.d05artic as ar on ar.refe05 = i.refcia01 and f.numfac39 = ar.fact05  and ar.refe05 = r.refe01 " & _
"           left join " & strOficina & "_extranet.ssfrac02 as fr on i.refcia01 = fr.refcia02  and ar.frac05 = fr.fraarn02  and fr.ordfra02 = ar.agru05   " & _
"             left join " & strOficina & "_extranet.ssprov22 as prv on prv.cvepro22 = i.cvepro01   " & _
"               left join " & strOficina & "_extranet.ssumed31 as um on um.clavem31 = fr.u_medc02   " & _
"                    left join " & strOficina & "_extranet.sscont36 as cf1 on cf1.refcia36 = i.refcia01 and cf1.cveimp36 = '1'   " & _
"                    left join " & strOficina & "_extranet.sscont36 as cf3 on cf3.refcia36 = i.refcia01 and cf3.cveimp36 = '3'   " & _
"                    left join " & strOficina & "_extranet.sscont36 as cf6 on cf6.refcia36 = i.refcia01 and cf6.cveimp36 = '6'   " & _
"                    left join " & strOficina & "_extranet.sscont36 as cf15 on cf15.refcia36 = i.refcia01 and cf15.cveimp36 = '15'   " & _
"                    left join " & strOficina & "_extranet.ssipar12 as ipar2 on ipar2.refcia12 = i.refcia01 and ipar2.ordfra12 = fr.ordfra02 and ipar2.cveide12 in ('PS','TG','TL','OC','AL') " & _
"                    left join " & strOficina & "_extranet.ssmtra30  as transp on transp.clavet30 = i.cvemts01    " & _
"    where cc.rfccli18 in ('IIN850215MIA') and i.firmae01 is not null and i.firmae01 <> ''  and  i.fecpag01 >=  '"& finicio &"' and i.fecpag01 <= '"& ffinal &"' " & _
" group by i.refcia01,f.numfac39,ar.item05,fr.fraarn02,fr.ordfra02,ar.pfac05 " 


Set act2= Nothing
Set oConex = New cConexion
oConex.Open_Conn	
oConex.Create_Rst act2
oConex.Ex_Sql sqlAct,act2


dim ref,refAux  ',Subref,SubrefAux
dim cambio,rcli 
cambio = 1
rcli=""
'refAux=""


while not act2.eof
	response.Write("<tr align="&c&"center"&c&" bordercolor="&c&"#999999"&c&" bgcolor="&c&"#FFFFFF"&c&">")
	ref = act2.fields("29").value
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
 

genera_html "d",retornaDivision(act2.fields("1").value,act2.fields("71").value),"center"  'DIVISION 
genera_html "d","","center"  'IMPORTANCIA
genera_html "d","","center"  'CATEGORIA
genera_html "d",act2.fields("4").value,"center"  'Nombre del Material
genera_html "d","","center"  'Codigo SAP
genera_html "d","","center"  'Clase de Producto
genera_html "d","","center"  'Contrato Marco
genera_html "d","","center"  'Código SAP Proveedor
genera_html "d",act2.fields("9").value,"center"  'Proveedor

rcli =act2.fields("11").value
genera_html "d",retornaCuenta(rcli),"center"  'Cuenta 
genera_html "d",retornaCECO(rcli),"center"  'CECO

genera_html "d",act2.fields("12").value,"center"  'ODC
genera_html "d","","center"  'No IE
genera_html "d",act2.fields("14").value &"--P-"&act2.fields("177").value&"--P--"&act2.fields("17").value,"center"  'País de Origen 
genera_html "d",retornaCampoPuertoEmb(act2.fields("177").value,act2.fields("17").value,"cvepai01",mid(act2.fields("29").value,1,3),oConex),"center"  'País de Procedencia
genera_html "d",retornaRegion(act2.fields("14").value),"center"  'Region
genera_html "d",act2.fields("177").value,"center" 'PTO./CD DE ORIGEN
genera_html "d",act2.fields("18").value,"center"  'TAX ID/ RFC
genera_html "d",act2.fields("19").value,"center"  'Factura
genera_html "d",act2.fields("20").value & "--P--" &act2.fields("21").value & "--P--" & mid(act2.fields("29").value,1,3) ,"center"  'Fecha de Factura
genera_html "d",retornaIMPORTADOR(act2.fields("21").value,mid(act2.fields("29").value,1,3),oConex),"center"  'IMPORTADOR
genera_html "d",act2.fields("22").value,"center"  'Cantidad 
genera_html "d",act2.fields("23").value,"center"  'Unidad de Medida
genera_html "d",act2.fields("24").value,"center"  'Incoterms
genera_html "d",act2.fields("25").value,"center"  'Tipo de Transporte
genera_html "d",retornaAduana(act2.fields("26").value),"center"  'Aduana
genera_html "d",retornaAgenteAduanal(act2.fields("27").value),"center"  'Agente Aduanal
genera_html "d",act2.fields("28").value,"center"  'Patente Agente Aduanal
genera_html "d",act2.fields("29").value,"center"  'No. De Trafico

   
ref = act2.fields("29").value
 if (ref <> refAux)then
  refAux=ref

  
  'Lote 2
genera_html "d",retornaCampoContenedores(act2.fields("29").value,"marc01",mid(act2.fields("29").value,1,3),oConex),"center"  'No de Contenedor 
genera_html "d",act2.fields("31").value,"center"  'No. Pedimento
genera_html "d",act2.fields("32").value,"center"  'Fecha Pedimento
genera_html "d",act2.fields("33").value,"center"  'Mes
genera_html "d",act2.fields("34").value,"center"  'No.Semana
genera_html "d",act2.fields("35").value,"center"  'Cantidad de Operaciones 
genera_html "d",retornaCantContenedores(act2.fields("29").value,"'ISO','CON'",mid(act2.fields("29").value,1,3),oConex),"center"  'Cantidad de Contenedores
genera_html "d",retornaCantContenedores(act2.fields("29").value,"'BUL','CAJ','BID','PAL'",mid(act2.fields("29").value,1,3),oConex),"center"  'PALLETS/BULTOS 
genera_html "d",retornaTipoContenedores(act2.fields("29").value,mid(act2.fields("29").value,1,3), oConex),"center"  'TIPO DE CONTENEDOR/ CAJA
genera_html "d",act2.fields("39").value,"center"  'Fecha Factura
genera_html "d",act2.fields("40").value,"center"  'Fecha BL
genera_html "d",act2.fields("41").value,"center"  'Fecha de arribo a la aduana
genera_html "d",act2.fields("42").value,"center"  'Fecha Desaduanamiento
 genera_html "d","","center"  'KPI Desaduanamiento
 genera_html "d","","center"  'KPI lead TIME
 genera_html "d","","center"  'TARGET TIME
 genera_html "d","","center"  'Fecha de arribo a la planta
 genera_html "d","","center"  'BW ? 
 genera_html "d","","center"  ' BW $  
 genera_html "d","","center"  'PL
 genera_html "d","","center"  'TT
 genera_html "d","","center"  'PR
 genera_html "d","","center"  'AA
 genera_html "d","","center"  'CO
 genera_html "d","","center"  'AL
 genera_html "d","","center"  'NUMERO DE EMBARQUE
 genera_html "d","","center"  'IDOT 
genera_html "d",retornaCampoCtaGastos(act2.fields("29").value,"cgas31",mid(act2.fields("29").value,1,3),oConex),"center"  'No. CTA DE GASTOS
genera_html "d",regresa_fecha_cuenta_gastos(act2.fields("29").value,mid(act2.fields("29").value,1,3),oConex),"center"  'Fecha C.Gastos
genera_html "d",regresa_tipo_Cgastos(act2.fields("29").value,mid(act2.fields("29").value,1,3),oConex),"center"    'Tipo C.Gastos
genera_html "d",retornaMontoAnticipo(act2.fields("29").value,"ANT",mid(act2.fields("29").value,1,3),oConex),"center"  ' Monto de Anticipo 



Subref = act2.fields("81").value
 if (ordenOcupado(Subref,act2.fields("29").value) = False)then
   ocuparOrd(Subref)
   
   '-----
	genera_html "d",act2.fields("59").value,"center"  ' Precio Pagado / valor comercial 
	genera_html "d",act2.fields("60").value,"center"  '  Valor comercial USD 
	genera_html "d",act2.fields("61").value,"center"  ' VALOR FLETES INTERNACIONAL M.N. 
	genera_html "d",act2.fields("62").value,"center"  ' SEGUROS 
	genera_html "d",act2.fields("63").value,"center"  ' OTROS INCREMENTABLES 
	genera_html "d",act2.fields("64").value,"center"  ' VALOR ADUANA M.N. 
	genera_html "d",act2.fields("65").value,"center"  ' T.C. 
	genera_html "d",act2.fields("66").value,"center"  ' VALOR ADUANA  DLLS 

	if( act2.fields("25").value = "MARITIMO") then
		genera_html "d",act2.fields("67").value,"center"  ' VALOR FLETES AEREO DLLS 
		genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
		genera_html "d",act2.fields("69").value,"center"  ' VALOR FLETES MARITIMO DLLS. 
	else
		genera_html "d",act2.fields("69").value,"center"  ' VALOR FLETES AEREO DLLS 
		genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
		genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
	end if
	
	
	genera_html "d","","center"  ' SAVING 
	genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
	genera_html "d",act2.fields("72").value,"center"  'ARANCEL %
	genera_html "d",act2.fields("73").value,"center"  'ARANCEL PREFERENCIAL 
	genera_html "d",retornaECI(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3), oConex),"center"  ' MONTO DE RECUPERACION $  
   '-----
      
   genera_html "d",act2.fields("761").value,"center"  ' ADV FRACC. $ 

   genera_html "d",act2.fields("76").value,"center"  ' DTA $ 
   genera_html "d",act2.fields("77").value,"center"  'IVA %
   genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
 else
 
 
    '-----
	genera_html "d","","center"  ' Precio Pagado / valor comercial 
	genera_html "d","","center"  '  Valor comercial USD 
	genera_html "d",act2.fields("61").value,"center"  ' VALOR FLETES INTERNACIONAL M.N. 
	genera_html "d",act2.fields("62").value,"center"  ' SEGUROS 
	genera_html "d",act2.fields("63").value,"center"  ' OTROS INCREMENTABLES 
	genera_html "d","","center"  ' VALOR ADUANA M.N. 
	genera_html "d",act2.fields("65").value,"center"  ' T.C. 
	genera_html "d","","center"  ' VALOR ADUANA  DLLS 

	if( act2.fields("25").value = "MARITIMO") then
		genera_html "d",act2.fields("67").value,"center"  ' VALOR FLETES AEREO DLLS 
		genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
		genera_html "d",act2.fields("69").value,"center"  ' VALOR FLETES MARITIMO DLLS. 0
	else
		genera_html "d",act2.fields("69").value,"center"  ' VALOR FLETES AEREO DLLS 
		genera_html "d",act2.fields("68").value,"center"  ' VALOR FLETES TERRESTRE DLLS. 
		genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
	end if

	genera_html "d","","center"  ' SAVING 
	genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
	genera_html "d",act2.fields("72").value,"center"  'ARANCEL %
	genera_html "d",act2.fields("73").value,"center"  'ARANCEL PREFERENCIAL 
	genera_html "d",retornaECI(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3),oConex),"center"  ' MONTO DE RECUPERACION $  
   '-----
   
   
   genera_html "d","","center"  ' ADV FRACC. $ 
   genera_html "d",act2.fields("76").value,"center"  ' DTA $ 
   genera_html "d","","center"  'IVA %
   genera_html "d","","center"  ' IVA FRACC. $ 
 end if




'/Lote2
  
  
  
  
  
'Lote 1
 '  genera_html "d",act2.fields("76").value,"center"  ' DTA $ 
 '  genera_html "d",act2.fields("77").value,"center"  'IVA %
 '  genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
   genera_html "d",act2.fields("79").value,"center"  ' PREVAL. 
   genera_html "d",sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3),oConex),"center"  ' TOTAL IMPUESTOS 
   genera_html "d",sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3),oConex)& "Falta dividirle T.C.","center"  ' Total Impuestos USD 
   
	genera_html "d","N/A","center"  ' GTOS. ADUANA USD(SOLO FRONTERA) 
	
	genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"DEMORAS"),"I",mid(act2.fields("29").value,1,3),oConex),"center"  ' DEMORAS 
	genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"I",mid(act2.fields("29").value,1,3),oConex),"center"  ' ESTADIAS 
	 
	 if ucase(mid(act2.fields("29").value,1,3)) ="RKU" then
	 genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3),oConex),"center"  ' MANIOBRAS  
 	 genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"I",mid(act2.fields("29").value,1,3),oConex),"center"  ' ALMACENAJES 
	else
 	 genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"I",mid(act2.fields("29").value,1,3),oConex),"center"  ' MANIOBRAS
	 genera_html "d","?","center"  ' ALMACENAJES 
	end if
	

	'genera_html "d",retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"OTROS"),"I",mid(act2.fields("29").value,1,3),oConex),"center"  ' OTROS 
	dim TPH,DEM,EST,ALMMAN,MAN
	TPH=0
	DEM=0
	EST=0
	ALMAN=0
	MAN=0
	TPH=CDbl(retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3),oConex))
	DEM= CDbl (retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"DEMORAS"),"I",mid(act2.fields("29").value,1,3),oConex))
	EST = CDbl (retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ESTADIAS"),"I",mid(act2.fields("29").value,1,3),oConex))
	ALMAN= CDbl (retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"ALMACENAJES-MANIOBRAS"),"I",mid(act2.fields("29").value,1,3),oConex))
	
	if ucase(mid(act2.fields("29").value,1,3)) ="RKU" then
	  MAN= CDbl (retornaPagosHechos(act2.fields("29").value,retornaConceptosPH(mid(act2.fields("29").value,1,3),"MANIOBRAS"),"I",mid(act2.fields("29").value,1,3),oConex))
	end if
	
	'if(TPH = "" )then
	'TPH="0"
	'end if
	'if(DEM = "" )then
	'DEM="0"
	'end if
	'dim okey 
	'okey = TPH &"-"& DEM-EST &"-"& ALMAN-MAN & "-" &sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))& "-" &sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3))
	'genera_html "d",TPH-DEM-EST-ALMAN-MAN-sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))-sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  ' OTROS 
	if(revisaImpuestosFacturados( act2.fields("29").value,"I",mid(act2.fields("29").value,1,3), oConex) <> 0 )then
			STIMP=	CDbl(sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3),oConex))
			STIVA= 	CDbl(sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3), oConex))		

end if
	genera_html "d",TPH-DEM-EST-ALMAN-MAN-STIMP-STIVA,"center"  ' OTROS 
'if(act2.fields("29").value="RKU10-00902")then
'	response.write(TPH&","&DEM&","&EST&","&ALMAN&","&MAN&","&sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))&","&sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3)))
'	response.end()
'end if

	'genera_html "d",retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3),oConex)-sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))-sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3)),"center"  ' TOTAL GASTOS DIVERSOS 
	genera_html "d",retornaTOTALPagosHechos(act2.fields("29").value,"I",mid(act2.fields("29").value,1,3),oConex)-STIMP-STIVA,"center"  ' TOTAL GASTOS DIVERSOS 
	
	'genera_html "d",((retornaTOTALPagosHechos(act2.fields("29").value,"E",mid(act2.fields("29").value,1,3),oConex)-sumaTotalImpuestos(act2.fields("29").value,mid(act2.fields("29").value,1,3))-sumaTotalIVA(act2.fields("29").value,mid(act2.fields("29").value,1,3))) /act2.fields("65").value),"center"  ' TOTAL GASTOS DIVERSOS USD 
	genera_html "d",retornaTOTALPagosHechos(act2.fields("29").value,"E",mid(act2.fields("29").value,1,3),oConex)-STIMP-STIVA,"center"    ' TOTAL GASTOS DIVERSOS USD 

	genera_html "d",retornaHonorarios(act2.fields("29").value,"chon31",mid(act2.fields("29").value,1,3), oConex),"center"  ' HONORARIOS AG AD. $ 
	
'/Lote 1
 else
'bgcolor="#D7ECF4"

 
 'Lote 2
genera_html "d","","center"  'No de Contenedor 
genera_html "d",act2.fields("31").value,"center"  'No. Pedimento
genera_html "d",act2.fields("32").value,"center"  'Fecha Pedimento
genera_html "d",act2.fields("33").value,"center"  'Mes
genera_html "d",act2.fields("34").value,"center"  'No.Semana
genera_html "d",act2.fields("35").value,"center"  'Cantidad de Operaciones 
genera_html "d","","center"  'Cantidad de Contenedores
genera_html "d","","center"  'PALLETS/BULTOS 
genera_html "d","","center"  'TIPO DE CONTENEDOR/ CAJA
genera_html "d",act2.fields("39").value,"center"  'Fecha Factura
genera_html "d",act2.fields("40").value,"center"  'Fecha BL
genera_html "d",act2.fields("41").value,"center"  'Fecha de arribo a la aduana
genera_html "d",act2.fields("42").value,"center"  'Fecha Desaduanamiento
 genera_html "d","","center"  'KPI Desaduanamiento
 genera_html "d","","center"  'KPI lead TIME
 genera_html "d","","center"  'TARGET TIME
 genera_html "d","","center"  'Fecha de arribo a la planta
 genera_html "d","","center"  'BW ? 
 genera_html "d","","center"  ' BW $  
 genera_html "d","","center"  'PL
 genera_html "d","","center"  'TT
 genera_html "d","","center"  'PR
 genera_html "d","","center"  'AA
 genera_html "d","","center"  'CO
 genera_html "d","","center"  'AL
 genera_html "d","","center"  'NUMERO DE EMBARQUE
 genera_html "d","","center"  'IDOT 
genera_html "d",retornaCampoCtaGastos(act2.fields("29").value,"cgas31",mid(act2.fields("29").value,1,3),oConex),"center"  'No. CTA DE GASTOS
genera_html "d",regresa_fecha_cuenta_gastos(act2.fields("29").value,mid(act2.fields("29").value,1,3),oConex),"center"      'Fecha Cta de Gastos
genera_html "d","","center"    'Tipo C.Gastos
genera_html "d","","center"  ' Monto de Anticipo 



Subref = act2.fields("81").value
 if (ordenOcupado(Subref,act2.fields("29").value) = False)then
   ocuparOrd(Subref)
   
   '---------
    genera_html "d",act2.fields("59").value,"center"  ' Precio Pagado / valor comercial 
	genera_html "d",act2.fields("60").value,"center"  '  Valor comercial USD 
	genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL M.N. 
	genera_html "d","","center"  ' SEGUROS 
	genera_html "d","","center"  ' OTROS INCREMENTABLES 
	genera_html "d",act2.fields("64").value,"center"  ' VALOR ADUANA M.N. 
	genera_html "d",act2.fields("65").value,"center"  ' T.C. 
	genera_html "d",act2.fields("66").value,"center"  ' VALOR ADUANA  DLLS 
	genera_html "d","","center"  ' VALOR FLETES AEREO DLLS 
	genera_html "d","","center"  ' VALOR FLETES TERRESTRE DLLS. 
	genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
	genera_html "d","","center"  ' SAVING 
	genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
	genera_html "d",act2.fields("72").value,"center"  'ARANCEL %
	genera_html "d",act2.fields("73").value,"center"  'ARANCEL PREFERENCIAL 
	genera_html "d","","center"  ' MONTO DE RECUPERACION $  
   '---------
   
   
   genera_html "d",act2.fields("761").value,"center"  ' ADV FRACC. $ 
  
   genera_html "d","","center"  ' DTA $ 
   genera_html "d",act2.fields("77").value,"center"  'IVA %
   genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
 else
 
   '---------
    genera_html "d","","center"  ' Precio Pagado / valor comercial 
	genera_html "d","","center"  '  Valor comercial USD 
	genera_html "d","","center"  ' VALOR FLETES INTERNACIONAL M.N. 
	genera_html "d","","center"  ' SEGUROS 
	genera_html "d","","center"  ' OTROS INCREMENTABLES 
	genera_html "d","","center"  ' VALOR ADUANA M.N. 
	genera_html "d",act2.fields("65").value,"center"  ' T.C. 
	genera_html "d","","center"  ' VALOR ADUANA  DLLS 
	genera_html "d","","center"  ' VALOR FLETES AEREO DLLS 
	genera_html "d","","center"  ' VALOR FLETES TERRESTRE DLLS. 
	genera_html "d","","center"  ' VALOR FLETES MARITIMO DLLS. 
	genera_html "d",act2.fields("70").value,"center"  ' SAVING 
	genera_html "d",act2.fields("71").value,"center"  ' FRACC. ARANC. 
	genera_html "d",act2.fields("72").value,"center"  'ARANCEL %
	genera_html "d",act2.fields("73").value,"center"  'ARANCEL PREFERENCIAL 
	genera_html "d","","center"  ' MONTO DE RECUPERACION $  
   '---------
 
   genera_html "d","","center"  ' ADV FRACC. $ 
   
   genera_html "d","","center"  ' DTA $ 
   genera_html "d","","center"  'IVA %
   genera_html "d","","center"  ' IVA FRACC. $ 
 end if





'/Lote2
 
 

'Lote 1
   'genera_html "d","","center"  ' DTA $ 
   'genera_html "d",act2.fields("77").value,"center"  'IVA %
   'genera_html "d",act2.fields("781").value,"center"  ' IVA FRACC. $ 
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
	genera_html "d","","center"  ' HONORARIOS AG AD. $ 
'/Lote 1
 end if
 

'genera_html "d","","center"  ' FLETE TARIFA NORMAL  
'genera_html "d","?","center"  'TRANSPORTISTA
'genera_html "d","?","center"  ' COSTO EXTRA EN FLETE 
'genera_html "d","?","center"  ' TOTAL FLETE NAL 
genera_html "d","","center"  ' Total Gastos Indirectos  USD 
genera_html "d","","center"  'INLAND
genera_html "d","","center"  'Impacto Valor Factura. 
genera_html "d","","center"  ' VAL FACT USD 
genera_html "d",retornaPlantaEntrega(act2.fields("1").value),"center"  'PLANTA DE ENTREGA

genera_html "d",act2.fields("81").value,"center"  'ORDEN FRACCION
genera_html "d",act2.fields("82").value,"center"  'CUENTA ORDEN FRACCION

genera_html "d",act2.fields("Item05").value,"center"  'item
genera_html "d",act2.fields("firmita").value,"center"  'firmita
genera_html "d",act2.fields("sDescTransp").value,"center"  'descripcion del transporte

 response.Write("</tr>")

 act2.movenext()
wend

end sub

sub genera_html(tipo,valor,alineacion)
 if(tipo = "e")then
  'response.Write("<td width="&c&"100"&c&" align="&c&alineacion&c&" nowrap bgcolor="&c&"#CCFF99"&c&"><div align="&c&alineacion&c&"><strong><em><font size="&c&"2"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></em></strong></div></td>")
   response.Write("<td width="&c&"100"&c&" align="&c&alineacion&c&" nowrap bgcolor="&c&"#CCFF99"&c&"><div align="&c&alineacion&c&"><strong><em><font size="&c&"2"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></em></strong></div></td>")
 else 
  'response.Write("<td align="&c&alineacion&c&" nowrap background="&c&bgcolor&c&"><div align="&c&alineacion&c&"><font color="&c&"#000000"&c&" size="&c&"1"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></div></td>")
  'response.Write("<td align="&c&alineacion&c&" nowrap background="&c&bgcolor&c&"><div align="&c&alineacion&c&"><font color="&c&"#000000"&c&" size="&c&"1"&c&" face="&c&"Verdana, Arial, Helvetica, sans-serif"&c&">"&valor&"</font></div></td>")
   if bgcolor ="#D7ECF4" then
     response.Write("<td align="&c&alineacion&c&" class=xl73>"&valor&"</td>")
   else
     response.Write("<td align="&c&alineacion&c&" class=xl78>"&valor&"</td>")
   end if
 '"#D7ECF4"
 end if
end sub

function revisaImpuestosFacturados(referencia,tipoop,oficina, ByRef oConex)
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
"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
"    where  i.firmae01 <> ''  and cta.esta31 <> 'C'  and i.refcia01 = '"& referencia &"' and ep.conc21 = 1"

Set Rst2= Nothing
oConex.Create_Rst Rst2
oConex.Ex_Sql sqlAct,Rst2	
		
' Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = cadena_de_conexion()
'conn12 ="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=dai_extranet; UID=jorgel; PWD=lorenzana86; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="& oficina &"_extranet; UID=pedrobm; PWD=123; OPTION=16427"
' act2.ActiveConnection = conn12
' act2.Source = sqlAct
' act2.cursortype=0
' act2.cursorlocation=2
' act2.locktype=1
' act2.open()


if not(Rst2.eof) then
 revisaImpuestosFacturados =Rst2.fields("Ref").value
else
  revisaImpuestosFacturados = nothing
end if


end function

function regresa_fecha_cuenta_gastos(referencia,oficina,ByRef oConex)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select max(date_format(cta.fech31,'%d/%m/%Y')) as fech31 from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
" r.refe31 = '"&referencia&"' and cta.esta31 <> 'C' "

Set Rst3= Nothing
oConex.Create_Rst Rst3
oConex.Ex_Sql sqlAct,Rst3


' Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = cadena_de_conexion()
' conn12 ="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=dai_extranet; UID=jorgel; PWD=lorenzana86; OPTION=16427"
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=pedrobm; PWD=123; OPTION=16427"
' act2.ActiveConnection = conn12
' act2.Source = sqlAct
' act2.cursortype=0
' act2.cursorlocation=2
' act2.locktype=1
' act2.open()

if not(Rst3.eof) then
 regresa_fecha_cuenta_gastos =Rst3.fields("fech31").value
else
  regresa_fecha_cuenta_gastos =valor
   end if
end function

function regresa_tipo_Cgastos(referencia,oficina,ByRef oConex)
dim c,valor
 c=chr(34)
 valor="PENDIENTE"
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select if(COUNT(cta.cgas31) > 1, 'COMPLEMENTARIA','NORMAL')  as tipo from "&oficina&"_extranet.e31cgast as cta, "&oficina&"_extranet.d31refer as r "&_
" where cta.cgas31 = r.cgas31 and "&_
" r.refe31 = '"&referencia&"'  "&_
"  and cta.esta31 <> 'C' "

Set Rst4= Nothing
oConex.Create_Rst Rst4
oConex.Ex_Sql sqlAct,Rst4

if not(Rst4.eof) then
 regresa_tipo_Cgastos =Rst4.fields("tipo").value
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
 res =rsVac.fields("cvepro") '&","&rsVac.fields("descpro")
else
 res = desc
end if 

codigoProveedor = res
end function

function retornaCampoCtaGastos(referencia,campo,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
 sqlAct = "select r."& campo &" as campo from "&oficina&"_extranet.e31cgast as cta " &_
 " inner join  "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
 " where  r.refe31 = '"& referencia &"' and cta.esta31 <> 'C' "

 Set Rst5= Nothing
oConex.Create_Rst Rst5
oConex.Ex_Sql sqlAct,Rst5


 if not(Rst5.eof) then
 valor = Rst5.fields("campo").value
 Rst5.movenext()
 while not Rst5.eof
   valor = valor&", "&Rst5.fields("campo").value
   Rst5.movenext()
 wend
  retornaCampoCtaGastos = valor
 else
  retornaCampoCtaGastos =valor
 end if
end function

function retornaMontoAnticipo(referencia,campo,oficina,ByRef oConex)
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

	Set Rst6= Nothing
oConex.Create_Rst Rst6
oConex.Ex_Sql sqlAct,Rst6


 if not(Rst6.eof) then
 valor = Rst6.fields("campo").value
 Rst6.movenext()
 while not Rst6.eof
   valor = valor&", "&Rst6.fields("campo").value
   Rst6.movenext()
 wend
  retornaMontoAnticipo = valor
 else
  retornaMontoAnticipo =valor
 end if
end function

function retornaIMPORTADOR(clave,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
 sqlAct = "select c.nomcli18 as campo from "&oficina&"_extranet.ssclie18 as c where c.cvecli18 = "&clave

Set Rst7= Nothing
oConex.Create_Rst Rst7
oConex.Ex_Sql sqlAct,Rst7

 if not(Rst7.eof) then
 valor = Rst7.fields("campo").value
 Rst7.movenext()
 while not Rst7.eof
   valor = valor&", "&Rst7.fields("campo").value
   Rst7.movenext()
 wend
  retornaIMPORTADOR = valor
 else
  retornaIMPORTADOR =valor
 end if
end function

function retornaCampoPuertoEmb(pto,val,campo,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor=""
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
'sqlAct="select "& campo &" as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"'  "
sqlAct="SELECT "& campo &" as campo FROM "&oficina&"_extranet.c01ptoemb where cvepto01 ="& val &" and nompto01 like '"& pto &"%'"

Set Rst19= Nothing
oConex.Create_Rst Rst19
oConex.Ex_Sql sqlAct,Rst19


 if not(Rst19.eof) then
 valor = Rst19.fields("campo").value
 Rst19.movenext()
 while not Rst19.eof
   valor = valor&", "&Rst19.fields("campo").value
   Rst19.movenext()
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
 'response.write("Select * From A1:B15 where desccli like '%" & desc2 & "%'")
 'response.End()
if not(rsVac.eof)then
 res =rsVac.fields("cvecli")  '&","&rsVac.fields("desccli")

else
 res = desc
end if 

codigoCliente = res
end function

function retornaHonorarios(referencia,campo,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor=""
 
  if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

sqlAct=" select cta."&campo&" as campo from "&oficina&"_extranet.e31cgast as cta  " & _
       " inner join "&oficina&"_extranet.d31refer as r on cta.cgas31 = r.cgas31 " & _
       " where  r.refe31 = '"& referencia &"' and cta.esta31 = 'I' "

Set Rst8= Nothing
oConex.Create_Rst Rst8
oConex.Ex_Sql sqlAct,Rst8

 if not(Rst8.eof) then
 valor = Rst8.fields("campo").value
 Rst8.movenext()
 while not Rst8.eof
   valor = valor&", "&Rst8.fields("campo").value
   Rst8.movenext()
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

function retornaECI(referencia,tipope,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor=0

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 


sqlAct =" select c.import36 as Campo ,c.cveimp36,c.refcia36  from "& oficina &"_extranet.ssdag"& tipope &"01 as i " & _
		"  inner  join  "& oficina &"_extranet.sscont36 as c on i.refcia01 = c.refcia36 " & _
		"    where c.refcia36 = '"& referencia &"' and c.cveimp36 = '18' and i.rfccli01 in ('IF&610526C95','IFF610526PQ6')"

Set Rst9= Nothing
oConex.Create_Rst Rst9
oConex.Ex_Sql sqlAct,Rst9


	 if not(Rst9.eof) then
	 valor = Rst9.fields("Campo").value
	
	  retornaECI = valor
	 else
	  retornaECI = valor
	 end if
end function

function retornaPagosHechos(referencia,conceptos,tipope,oficina,ByRef oConex)
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
'		" where i.rfccli01 in ('IF&610526C95','IFF610526PQ6','ISI011214HM3') and i.refcia01 = '"&referencia&"'  and i.firmae01 <> ''   group by i.refcia01 "

sqlAct="select i.refcia01 as Ref, r.cgas31,ep.conc21,ep.piva21,sum(dp.mont21*if(ep.deha21 = 'C',-1,1)) as Importe, cp.desc21 " & _
" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
"  inner join "& oficina &"_extranet.d31refer as r on r.refe31 = i.refcia01  " & _
"     inner join "& oficina &"_extranet.e31cgast as cta on cta.cgas31 = r.cgas31 " & _
"          inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 and dp.cgas21 = r.cgas31 " & _
"             inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
"                  inner join  "& oficina &"_extranet.c21paghe as cp on cp.clav21 = ep.conc21 " & _
"    where  i.rfccli01 in ('IF&610526C95','IFF610526PQ6')  and i.firmae01 <> ''  and cta.esta31 <> 'C'  and ep.conc21 in ("&conceptos&") and i.refcia01 = '"&referencia&"'  group by Ref,cgas31,conc21"


Set Rst10= Nothing
oConex.Create_Rst Rst10
oConex.Ex_Sql sqlAct,Rst10


 if not(Rst10.eof) then
 valor = Rst10.fields("Importe").value
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

function retornaTOTALPagosHechos(referencia,tipope,oficina,ByRef oConex)
dim c,valor
 c=chr(34)
 valor=0

 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 

if(conceptos <> "NA" and conceptos <> "NE")then

sqlAct =" select i.refcia01 as Ref,sum(dp.mont21*if(ep.deha21 = 'C',-1,1)) as Importe " & _
		" from "& oficina &"_extranet.ssdag"&tipope&"01 as i  " & _
		" inner join "& oficina &"_extranet.d21paghe as dp on dp.refe21 = i.refcia01 " & _
		" inner join "& oficina &"_extranet.e21paghe as ep on ep.foli21 = dp.foli21 and year(ep.fech21) = year(dp.fech21) and ep.esta21 <> 'S' and ep.esta21 <> 'C'  and ep.tmov21 =dp.tmov21 " & _
		" where i.rfccli01 in ('IF&610526C95','IFF610526PQ6')  and i.refcia01 = '"&referencia&"'  and i.firmae01 <> ''  group by i.refcia01 "

'response.Write(sqlAct)
'response.End()

Set Rst11= Nothing
oConex.Create_Rst Rst11
oConex.Ex_Sql sqlAct,Rst11


 if not(Rst11.eof) then
 valor = Rst11.fields("Importe").value
 Rst11.movenext()
 while not Rst11.eof
   valor = valor&", "&Rst11.fields("Importe").value
   Rst11.movenext()
 wend
  retornaTOTALPagosHechos = valor
 else
  retornaTOTALPagosHechos = valor
 end if
 else
   retornaTOTALPagosHechos =0
 end if
end function

function retornaCampoContenedores(referencia,campo,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select "& campo &" as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"'  "

Set Rst12= Nothing
oConex.Create_Rst Rst12
oConex.Ex_Sql sqlAct,Rst12


 if not(Rst12.eof) then
 valor = Rst12.fields("campo").value
 Rst12.movenext()
 while not Rst12.eof
   valor = valor&", "&Rst12.fields("campo").value
   Rst12.movenext()
 wend
  retornaCampoContenedores = valor
 else
  retornaCampoContenedores =valor
 end if
end function

function retornaCantContenedores(referencia,campo,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
sqlAct="select count(*) as campo from "&oficina&"_extranet.d01conte where refe01 = '"&referencia&"' and clas01 in ("&campo&") "

Set Rst13= Nothing
oConex.Create_Rst Rst13
oConex.Ex_Sql sqlAct,Rst13

 if not(Rst13.eof) then
 valor = Rst13.fields("campo").value
 Rst13.movenext()
 while not Rst13.eof
   valor = valor&", "&Rst13.fields("campo").value
   Rst13.movenext()
 wend
  retornaCantContenedores = valor
 else
  retornaCantContenedores =valor
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

function retornaTipoContenedores(referencia,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor=""
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct="select distinct cn4.tipcon40 as campo from "& oficina &"_extranet.sscont40 as cn4 where cn4.refcia40 = '"& referencia &"' "

Set Rst14= Nothing
oConex.Create_Rst Rst14
oConex.Ex_Sql sqlAct,Rst14

 if not(Rst14.eof) then
 valor = DatosContenedor(Rst14.fields("campo").value)
 Rst14.movenext()
 while not Rst14.eof
   valor = DatosContenedor(valor) &", "& DatosContenedor(Rst14.fields("campo").value)
   Rst14.movenext()
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

if clave = "11000" then
  res = revisaFraccion(fraccion)
  val = res '"Centro de Distribución"
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

function sumaTotalImpuestos(referencia,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor="0"
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct=" select ifnull(sum(import36),0) as campo from "& oficina &"_extranet.sscont36 as cf1 " & _
       " where cf1.cveimp36 in ('1', '6','15')   and refcia36 = '"&referencia&"' "
      ' " where cf1.cveimp36 in ('1','3','6','15')   and refcia36 = '"&referencia&"' "
	  
Set Rst15= Nothing
oConex.Create_Rst Rst15
oConex.Ex_Sql sqlAct,Rst15

' Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=pedrobm; PWD=123; OPTION=16427"
' conn12 ="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=dai_extranet; UID=jorgel; PWD=lorenzana86; OPTION=16427"

' act2.ActiveConnection = conn12
' act2.Source = sqlAct
' act2.cursortype=0
' act2.cursorlocation=2
' act2.locktype=1
' act2.open()
 if not(Rst15.eof) then
 valor = Rst15.fields("campo").value
 Rst15.movenext()
 while not Rst15.eof
   valor = valor &", "& Rst15.fields("campo").value
   Rst15.movenext()
 wend
  sumaTotalImpuestos = valor
 else
  sumaTotalImpuestos =valor
 end if

'ADV/IGI+DTA+IVA+PREVAL

end function

function sumaTotalIVA(referencia,oficina, ByRef oConex)
dim c,valor
 c=chr(34)
 valor=0
 
 if (ucase(oficina) = "ALC")then
 oficina = "LZR"
 end if
 
 
sqlAct=" select ifnull(sum(import36),0) as campo from "& oficina &"_extranet.sscont36 as cf1 " & _
       " where cf1.cveimp36 in ('3')   and refcia36 = '"&referencia&"' "
      ' " where cf1.cveimp36 in ('1','3','6','15')   and refcia36 = '"&referencia&"' "
Set Rst16= Nothing
oConex.Create_Rst Rst16
oConex.Ex_Sql sqlAct,Rst16

' Set act2= Server.CreateObject("ADODB.Recordset")
'conn12 = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=10.66.1.5; DATABASE="&oficina&"_extranet; UID=pedrobm; PWD=123; OPTION=16427"
' conn12 ="DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; DATABASE=dai_extranet; UID=jorgel; PWD=lorenzana86; OPTION=16427"

' act2.ActiveConnection = conn12
' act2.Source = sqlAct
' act2.cursortype=0
' act2.cursorlocation=2
' act2.locktype=1
' act2.open()
 if not(Rst16.eof) then
 valor = Rst16.fields("campo").value
 'act2.movenext()
 'while not act2.eof
   'valor = valor &", "& act2.fields("campo").value
 '  act2.movenext()
 'wend
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
dim val,aux
val =""



If InStr(rcli,"CECO")>0 then
'If inStr(ucase(rcli),"CUENTA")>0 then

 aux=split(rcli," ")
' Response.write(rcli & ", "& Ubound(aux) & ":" & aux(0) &"," & aux(1))
' response.End()
 if Ubound(aux) = 0  then
'  aux=ucase(mid(aux,"CECO:")
	  if InStr(aux(0),"CECO")>0then
	   val = aux(0)
	  else
	  
	  
	  
	    if InStr(aux(1),"CECO")>0then
		 val = aux(1)
	    else
	     val = "N/E"
	    end if
		
		
		
		
	  end if
 else
  if(Ubound(aux) = 1) then
  
      if InStr(aux(0),"CECO")>0then
	    val = aux(0)
	  else
	    if InStr(aux(1),"CECO")>0then
		 val = aux(1)
	    else
	     val = "N/E"
	    end if
	  end if
	  
   else
    val="ERROR:"&rcli& "," &Ubound(aux)
   end if
 end if
Else
val = "N/E"
End if

retornaCECO = val
end function

function retornaCuenta(rcli)
dim val,aux
val =""



If InStr(rcli,"CUENTA")>0 then
 aux=split(rcli," ")
 if Ubound(aux) = 0  then
'  aux=ucase(mid(aux,"CECO:")
	  if InStr(aux(0),"CUENTA")>0then
	   val = aux(0)
	  else
	    'if InStr(aux(1),"CUENTA")>0then
		' val = aux(1)
	    'else
	     val = "N/E"
	    'end if
	  end if
 else
  if(Ubound(aux) = 1) then
  
      if InStr(aux(0),"CUENTA")>0then
	    val = aux(0)
	  else
	    if InStr(aux(1),"CUENTA")>0then
		 val = aux(1)
	    else
	     val = "N/E"
	    end if
	  end if
	  
   else
    val="ERROR:"&rcli
   end if
 end if
Else
val = "N/E"
End if


retornaCuenta = val
end function
%>

