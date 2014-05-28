<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<%
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))

Response.Buffer = TRUE
Response.Addheader "Content-Disposition", "attachment;"
Response.ContentType = "application/vnd.ms-excel"
Server.ScriptTimeOut = 2000

strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")

if not permi = "" then
  permi = "  and (" & permi & ") "
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")
if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if
if blnAplicaFiltro then
   permi = " AND cvecli01 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi = ""
end if

dim strDateIni,strDateFin,strTipoPedimento,strCodError,strHTML
dim strCtaGas,dblSumPH,dblchon31,dblpiva31,dblcsce31,dblsuph31,dblpaho31,SumaIva
dim strCvePH,strConcepto,intCvePH,strFecPag,strNumPed,strpagpre,strEsta,strMoneda
dim strcveped,intCountRef,intCounter,dblSumProrPH,dblProrPH
'Variales de Acumulamiento para el registro de resumen por Referencia
dim dblAcumValFac,dblAcumADV,dblAcumIVA,dblAcumPRV,dblAcumDTA

dim arrCvePH()	'Array de Claves de PH
dim arrConcepto()	'Array de Conceptos
dim arrRefer()	'Array de Referencias
dim arrAcumProrPH() 'Array de Acumulacion de Prorrateos por Tipo de PH
dim arrAcumIVAPH() 'Array de Acumulacion de Prorrateos de IVA de PH por Tipo de PH

strHTML = ""
strDateIni = ""
strDateFin = ""
strTipoPedimento = ""
strCodError = "0"

strDateIni = trim(request.Form("txtDateIni"))
strDateFin = trim(request.Form("txtDateFin"))
strTipoPedimento = trim(request.Form("rbnTipoDate"))
strUsuario = request.Form("user")
strTipoUsuario = request.Form("TipoUser")

if not isdate(strDateIni) then
	strCodError = "5"
end if
if not isdate(strDateFin) then
	strCodError = "6"
end if
if strDateIni = "" or strDateFin = "" then
	strCodError = "1"
end if
if strCodError = "0" then
  strSQL = "SELECT refcia01 FROM " & TablaPed(strTipoPedimento) & " WHERE fecpag01 >='" & FormatoFechaInv(strDateIni) & "' AND fecpag01 <='" & FormatoFechaInv(strDateFin) & "' " & Permi & " and firmae01 != '' order by fecpag01"
  'Buscamos las referencias y sus pagos hechos para poder determinar que titulos de PH se pintarán dinamicamente
  Set RsRep = Server.CreateObject("ADODB.Recordset")
  RsRep.ActiveConnection = MM_EXTRANET_STRING
  RsRep.Source = strSQL
  RsRep.CursorType = 0
  RsRep.CursorLocation = 2
  RsRep.LockType = 1
  RsRep.Open()

  if not RsRep.eof then
	  intCountRef = 1 'Contador de Referencias
	  intCounter = 1 'Contador de Conceptos

	  ' Comienza el HTML, se pintan los titulos
	  strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DESGLOSADO DE MERCANCIAS DE " & TipoOperDescr(strTipoPedimento) & " </p></font></strong>"
	  strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p> </p></font></strong>"
	  strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>" & strDateIni & " al " & strDateFin & "</p></font></strong>"

		While NOT RsRep.EOF
			redim preserve arrRefer(intCountRef)

			'Se asigna el nombre de la referencia
			arrRefer(intCountRef) = RsRep.Fields.Item("refcia01").Value	'Guarda Referencias

      strSQL2 = "select E.conc21, C.desc21 from d21paghe AS D,E21paghe AS E , c21paghe as C where D.refe21='" & arrRefer(intCountRef) & "' and E.foli21 = D.foli21 and E.fech21 = D.fech21 and C.clav21 = E.conc21"
      Set RsRep2 = Server.CreateObject("ADODB.Recordset")
			RsRep2.ActiveConnection = MM_EXTRANET_STRING
			RsRep2.Source = strSQL2
			RsRep2.CursorType = 0
			RsRep2.CursorLocation = 2
			RsRep2.LockType = 1
			RsRep2.Open()

			if not RsRep2.eof then
			'lleva los registros de la d21paghe para una cuenta de gastos y una referencia determinada
				while not RsRep2.eof
					redim preserve arrCvePH(intCounter)
					redim preserve arrConcepto(intCounter)
					'por cada detalle busco el concepto
          strCvePH = cstr(RsRep2.Fields.Item("conc21").Value)
					'strConcepto = RsRep2.Fields.Item("desc21").Value
					'Mantiene que el concepto no existe en el array hasta que es encontrado
					strEsta = False
					'Busca el concepto entre los que ya tenemos
					For i = 1 To intCounter
						if (arrCvePH(i) = strCvePH) then
							strEsta = true
						end if
					Next

					'Guardo el concepto si no se encuentra ya en el arreglo de conceptos
					if strEsta = false then
						arrCvePH(intCounter) = strCvePH	'Guarda Claves de PH
						arrConcepto(intCounter) = cstr(RsRep2.Fields.Item("desc21").Value)'Guarda conceptos de PH
						intCounter = intCounter + 1
					end if
				RsRep2.movenext
				wend
				end if
			RsRep2.close
			Set RsRep2 = Nothing

  intCountRef = intCountRef + 1
	RsRep.movenext
	Wend
	RsRep.close
	Set RsRep = Nothing

  'Total de Referencias
  intTotRefer = intCountRef - 1
  'Total de Conceptos
  intTotConcep = intCounter - 1

  'Procedimiento de quitar los puntos que tiene el catalogo de Pagos Hechos
  if not arrConcepto(1) = "" then
    For i = 1 To intTotConcep
      arrConcepto(i) = QuitaPuntosPH(arrConcepto(i))
    Next
  end if

      'Pintado de las columnas estaticas y dinamicas
      strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
      strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Observaciones</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Status</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Indicador NAFTA</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Documento</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Numero de Proveedor de Material</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripción de Proveedor de Material</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Factura de Proveedor</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedido</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fracción</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Código de Producto</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Descripción</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cantidad</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Unidad de Medida</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Valor Factura</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Moneda</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IGI</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DTA</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PRV</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Moneda</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Factura C.G.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Honorarios</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Complementarios</td>" & chr(13) & chr(10)
      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA de la C.G.</td>" & chr(13) & chr(10)

      if not arrConcepto(1) = "" then
        For i = 1 To intTotConcep
          strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" & arrConcepto(i) & "</td>" & chr(13) & chr(10)
          strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA " & arrConcepto(i) & "</td>" & chr(13) & chr(10)
        Next
      end if

      strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Importe Total de la C.G.</td>" & chr(13) & chr(10)
      strHTML = strHTML & "</tr>"& chr(13) & chr(10)

      'Aqui se despliegan las referencias
      For b = 1 To intTotRefer
        intContReg = 1 'Cuenta Registros
        dbl_i_dta = 0
        dbl_i_prv = 0
        dblAcumValFac = 0
        dblAcumADV = 0
        dblAcumIVA = 0
        dblAcumPRV = 0
        dblAcumDTA = 0
        dblSumProrPH = 0  'Suma de Prorrateos de PH
        dblProrPH = 0
        dbladv = 0
        dbliva = 0
        dbldta = 0
        dblprv = 0
        SumaIva = 0

        strCtaGas = ""
        dblchon31 = 0
        dblpiva31 = 0
        dblpaho31 = 0
        dblcsce31 = 0
        dblsuph31 = 0
        dblImporteCG = 0

        strRefer = arrRefer(b)
        'Checa si la referencia es rectificada o rectificacion
        'Si es una de ellas lo almacena en ObservRect
        ObservRect = ""
        strRect = ""
        strRect = RegresaRect(strRefer,"Rectificado")
        if not strRect = ""  then
          ObservRect = strRect
        end if
        strRect = RegresaRect(strRefer,"Rectificacion")
        if not strRect = ""  then
          ObservRect = strRect
        end if

        'Datos generales del pedimento y los impuestos prv
        if strTipoPedimento ="1" then
          strSQL7 = "select A.numped01,A.fecpag01,A.pagpre01,A.cveped01,A.valmer01,B.cveimp36, B.import36, C.cgas31, D.chon31, D.piva31, D.paho31, D.suph31,D.csce31 from ssdagi01 as A, sscont36 as B, d31refer as C, e31cgast as D  where A.refcia01='" & strRefer & "' and A.refcia01 = B.refcia36 and A.refcia01 = C.refe31 and C.cgas31 = D.cgas31 and D.esta31 = 'I'"
        end if
        if strTipoPedimento = "2" then
          strSQL7 = "select A.numped01,A.fecpag01,A.pagpre01,A.cveped01,A.valfac01,B.cveimp36, B.import36, C.cgas31, D.chon31, D.piva31, D.paho31, D.suph31,D.csce31 from ssdage01 as A, sscont36 as B, d31refer as C, e31cgast as D  where A.refcia01='" & strRefer & "' and A.refcia01 = B.refcia36 and A.refcia01 = C.refe31 and C.cgas31 = D.cgas31 and D.esta31 = 'I'"
        end if

        Set RsRep7 = Server.CreateObject("ADODB.Recordset")
        RsRep7.ActiveConnection = MM_EXTRANET_STRING
        RsRep7.Source = strSQL7
        RsRep7.CursorType = 0
        RsRep7.CursorLocation = 2
        RsRep7.LockType = 1
        RsRep7.Open()

        if not RsRep7.eof then
		      strFecPag = RsRep7.Fields.Item("fecpag01").Value
          strNumPed = RsRep7.Fields.Item("numped01").Value
          strpagpre = RsRep7.Fields.Item("pagpre01").Value
          strcveped = RsRep7.Fields.Item("cveped01").Value

          if strTipoPedimento ="1" then
           dblvalmer = cdbl(RsRep7.Fields.Item("valmer01").Value)
          else
           dblvalmer = cdbl(RsRep7.Fields.Item("valfac01").Value)
          end if

          strCtaGas = RsRep7.Fields.Item("cgas31").Value
          dblchon31 = cdbl(RsRep7.Fields.Item("chon31").Value)
          dblpiva31 = cdbl(RsRep7.Fields.Item("piva31").Value)
          dblpaho31 = cdbl(RsRep7.Fields.Item("paho31").Value)
          dblcsce31 = cdbl(RsRep7.Fields.Item("csce31").Value)
          dblsuph31 = cdbl(RsRep7.Fields.Item("suph31").Value)
          'IVA de la Cuenta de Gastos
          SumaIva = dblchon31 + dblpaho31 + dblcsce31
          SumaIva = SumaIva * (dblpiva31/100)

           While not RsRep7.eof
            Select case trim(RsRep7.Fields.Item("cveimp36").Value)
                  case "1"    dbl_i_dta = trim(cdbl(RsRep7.Fields.Item("import36").Value))
                  case "15"   dbl_i_prv = trim(cdbl(RsRep7.Fields.Item("import36").Value))
                End Select
            RsRep7.movenext
           Wend
          end if
          RsRep7.close
          Set RsRep7 = Nothing
        strHTML = DespliegaRepDesgMerc(strRefer)
        response.Write(strHTML)
        strHTML = ""
        strHTML = DespliegaRefer(strRefer)
        response.Write(strHTML)
      Next

	   strHTML = strHTML & "</table>"& chr(13) & chr(10)
  else
	   strHTML = ""
  end if

  if strHTML = "" then
    strHTML = "NO EXISTEN REGISTROS"
  end if

  'Se pinta todo el HTML formado
  response.Write(strHTML)
else
	select case strCodError
    case "1"
	   strMenjError = "Campo en Blanco Especifique!.."
	case "5","6"
	   strMenjError = "Fechas Erroneas, Verifique!"
	case "7"
	   strMenjError = "Registros No Encontrados!"
	end select
	%>
	<table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
	  <tr>
		<td><%=strMenjError%></td>
	  </tr>
	</table>
<%end if%>
<%
'Funcion que va elaborando el reporte desglosado de referencias y devuelve el HTML
Function DespliegaRepDesgMerc(pRefer)
'Variables de resultado de impuestos
dim dbladv,dblprv,dbldta,dbliva,intCountFracc,strFact,strFracc
dim arrSumProrPH()  'Array de Suma de Prorrateos para cada registro
dim arrCodPH() 'array de codigos de PH
dim arrImpPH() 'Array de Importes de PH
dim Loencontro
dim arrCodPHCA() 'Array de Codigos de PH por Costo Aduanal
dim arrImpPHCA() 'Array de Importes de PH por Costo Aduanal
dim LoencontroCA
redim arrAcumProrPH(intTotConcep)
redim arrAcumIVAPH(intTotConcep)
'Variables de Acumulamiento para el registro x resumen por Factura y por Fraccion
dim dblResFraccVF, dblResFactVF, dblResADV, dblResIVA, dblResDTA, dblResPRV,dblmerme

strFact = ""
strFracc = ""
strEntra = ""
dblResFraccVF = 0
dblResFactVF = 0
dblResADV = 0
dblResDTA = 0
dblResIVA = 0
dblResPRV = 0
strMoneda = ""
dblSumFac = 0

'si la cuenta de gasto tiene mas de una referencia
'Cuenta cuantas referencias tiene una cuenta de gastos
'Regresa 1 si tiene una, y el llena un array de referencias si tiene mas
if not strCtaGas = "" then
  strSQL5 = "select count(refe31) as CountRef, refe31 from d31refer where cgas31='" & strCtaGas & "' group by refe31"

  Set RsRep5 = Server.CreateObject("ADODB.Recordset")
  RsRep5.ActiveConnection = MM_EXTRANET_STRING
  RsRep5.Source = strSQL5
  RsRep5.CursorType = 0
  RsRep5.CursorLocation = 2
  RsRep5.LockType = 1
  RsRep5.Open()

  if not RsRep5.eof then
  'Total de Referencias para esa cuenta de gastos
    intCountValfac = cint(RsRep5.Fields.Item("CountRef").Value)
    if not intCountValfac = 1 then
     'Tiene mas de una cuenta de gastos
     dblSumFac = 0
      While not RsRep5.eof
      'Obtengo la suma de valores factura para los productos de distintas referencias de una misma cuenta de gastos
        dblSumFac = SumaValFacxRef(RsRep5.Fields.Item("refe31").Value,dblSumFac)
      RsRep5.movenext
      Wend
    end if
  end if
  RsRep5.close
  Set RsRep5 = Nothing
end if

'Aqui se obtienen los datos del producto por referencia
strSQL6 = "select frac05,fact05,pedi05,prov05 as cveprov,item05 as cveprod, desc05 as descrip, caco05 as ccomer, umco05 as unicom, vafa05 as valfac from d05artic where refe05='" & pRefer & "' ORDER BY fact05,frac05"

Set RsRep6 = Server.CreateObject("ADODB.Recordset")
RsRep6.ActiveConnection = MM_EXTRANET_STRING
RsRep6.Source = strSQL6
RsRep6.CursorType = 3
RsRep6.CursorLocation = 2
RsRep6.LockType = 1
RsRep6.Open()

if not RsRep6.eof then
intCountFracc = 1 'Contador de Fracciones
strEntra = ""
'Guarda la ultima fraccion y factura
RsRep6.MoveLast 'Va al ultimo registro
strUltimaFracc = RsRep6.Fields.Item("frac05").Value
strUltimaFact = RsRep6.Fields.Item("fact05").Value
RsRep6.MoveFirst'Va al primer registro

	While not RsRep6.eof
  'Cuando hay varios registros
    if intCountFracc > 1 then
      'Si cambia el numero de fraccion,(abarca la ultima de cada factura)
        if not strFracc = RsRep6.Fields.Item("frac05").Value and (strFact = RsRep6.Fields.Item("fact05").Value) then
          'response.write("entra 1")
          strEntra = "Fraccion"
          strHTML = strHTML & DespliegaFracc(strFracc,dblResFraccVF,dblResADV,dblResIVA,dblResDTA,dblResPRV)
          strFracc = RsRep6.Fields.Item("frac05").Value 'Guarda Fraccion
          dblResFraccVF = 0
          dblResADV = 0
          dblResIVA = 0
          dblResDTA = 0
          dblResPRV = 0
        end if
      'Si cambia el numero de factura, (no abarca la ultima)
        if not strFact = RsRep6.Fields.Item("fact05").Value then
          strEntra ="Factura"
          strHTML = strHTML & DespliegaFracc(strFracc,dblResFraccVF,dblResADV,dblResIVA,dblResDTA,dblResPRV)
          strHTML = strHTML & DespliegaFact(strFact,dblResFactVF)
          strFact = RsRep6.Fields.Item("fact05").Value 'Guarda Factura
          strFracc = RsRep6.Fields.Item("frac05").Value 'Guarda Fraccion
          dblResFactVF = 0
          dblResFraccVF = 0
          dblResADV = 0
          dblResIVA = 0
          dblResDTA = 0
          dblResPRV = 0
        end if
    end if

    'Aqui se obtienen los impuestos por fraccion
    strSQL7 = "select A.vmerme02, A.i_adv102, A.i_adv202, A.i_adv302, A.i_iva102, A.i_iva202, A.i_iva302, B.monfac39, B.facmon39 from ssfrac02 as A, ssfact39 as B where A.refcia02='" & pRefer & "' AND A.fraarn02 = '" & RsRep6.Fields.Item("frac05").Value & "' and A.refcia02 = B.refcia39 order by A.fraarn02,A.ordfra02"
    Set RsRep7 = Server.CreateObject("ADODB.Recordset")
    RsRep7.ActiveConnection = MM_EXTRANET_STRING
    RsRep7.Source = strSQL7
    RsRep7.CursorType = 0
    RsRep7.CursorLocation = 2
    RsRep7.LockType = 1
    RsRep7.Open()

    if not RsRep7.eof then
      dbl_i_adv = cdbl(RsRep7.Fields.Item("i_adv102").Value) + cdbl(RsRep7.Fields.Item("i_adv202").Value) + cdbl(RsRep7.Fields.Item("i_adv302").Value)
		  dbl_i_iva = cdbl(RsRep7.Fields.Item("i_iva102").Value) + cdbl(RsRep7.Fields.Item("i_iva202").Value) + cdbl(RsRep7.Fields.Item("i_iva302").Value)
      dblmerme = cdbl(RsRep7.Fields.Item("vmerme02").Value)
      dbl_factmon = cdbl(RsRep7.Fields.Item("facmon39").Value)
      strMoneda = RsRep7.Fields.Item("monfac39").Value
    end if
    RsRep7.close
    Set RsRep7 = Nothing

  	'Datos de la referencia
		strHTML = strHTML&"<tr>" & chr(13) & chr(10)
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Numero de Referencia
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Observaciones
    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Status
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Numero de Pedimento
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Indica si el pedimento tiene algun Tratado de TLC America del Norte
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">//</font></td>" & chr(13) & chr(10)'Fecha de Pago

		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep6.Fields.Item("cveprov").Value & "&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Proveedor de producto
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & CampoProveedor("nompro22",RsRep6.Fields.Item("cveprov").Value) & "</font></td>" & chr(13) & chr(10) 'Descripcion de Proveedor de producto

    'Datos de la factura
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Factura
    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep6.Fields.Item("pedi05").Value & "&nbsp;</font></td>" & chr(13) & chr(10) 'Pedido

    'Datos de la fraccion
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Fraccion del producto

    'Datos del producto
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep6.Fields.Item("cveprod").Value & "</font></td>" & chr(13) & chr(10) 'Clave de Producto
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep6.Fields.Item("descrip").Value & "</font></td>" & chr(13) & chr(10) 'Descripcion del Producto
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep6.Fields.Item("ccomer").Value & "</font></td>" & chr(13) & chr(10) 'Cantidad comercial del Producto
		strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RegresaDescUM(RsRep6.Fields.Item("unicom").Value) & "</font></td>" & chr(13) & chr(10) 'Unidad Comercial del producto

    'Acumula el valor Factura
    dblAcumValFac = dblAcumValFac + cdbl(RsRep6.Fields.Item("valfac").Value)

    'Acumula el Valor Factura para el registro resumen de Factura y Fracciones
    dblResFactVF = dblResFactVF + cdbl(RsRep6.Fields.Item("valfac").Value)
    dblResFraccVF = dblResFraccVF + cdbl(RsRep6.Fields.Item("valfac").Value)
    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & RsRep6.Fields.Item("valfac").Value & "</font></td>" & chr(13) & chr(10)	'Valor del producto en la Factura

    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & strMoneda & "</font></td>" & chr(13) & chr(10)

    if intCountValfac = 1 then
      dbladv = (cdbl(RsRep6.Fields.Item("valfac").Value) * dbl_i_adv) / dblmerme
      dbldta = (cdbl(RsRep6.Fields.Item("valfac").Value) * dbl_i_dta) / dblvalmer
      dblprv = (cdbl(RsRep6.Fields.Item("valfac").Value) * dbl_i_prv) / dblvalmer
      dbliva = (cdbl(RsRep6.Fields.Item("valfac").Value) * dbl_i_iva) / dblmerme
    else
      dbladv = (dblSumFac * dbl_i_adv) / dblmerme
      dbldta = (dblSumFac * dbl_i_dta) / dblvalmer
      dblprv = (dblSumFac * dbl_i_prv) / dblvalmer
      dbliva = (dblSumFac * dbl_i_iva) / dblmerme
    end if

    'Variables de Acumulado de IGI e Iva
    dblAcumADV = dblAcumADV + dbladv
    dblAcumIVA = dblAcumIVA + dbliva
    dblAcumDTA = dblAcumDTA + dbldta
    dblAcumPRV = dblAcumPRV + dblprv
    dblResADV = dblResADV + dbladv
    dblResDTA = dblResDTA + dbldta
    dblResIVA = dblResIVA + dbliva
    dblResPRV = dblResPRV + dblprv

    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dbladv & "</font></td>" & chr(13) & chr(10) 'Prorrateo del impuesto ADV
    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dbldta & "</font></td>" & chr(13) & chr(10) 'Prorrateo del impuesto DTA
    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblprv & "</font></td>" & chr(13) & chr(10) 'Prorrateo del impuesto de PRV
    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dbliva & "</font></td>" & chr(13) & chr(10) 'Prorrateo de IVA
    strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">MXP</font></td>" & chr(13) & chr(10)

    'Datos de la Cuenta
		if not strCtaGas = "" then

		  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Cuenta de Gasto

      if dblchon31 > 0 and cdbl(RsRep6.Fields.Item("valfac").Value) > 0 then
        if intCountValfac = 1 then
        'Honorarios normales para una cuenta de gastos con solo una referencia
        dblHonorarios = (cdbl(RsRep6.Fields.Item("valfac").Value) * dblchon31)/ dblvalmer
        else
        'HONORARIOS *  VALDOLITEM / SUMAGLOBAL EN DLS
         dblHonorarios = dblchon31 * ((cdbl(RsRep6.Fields.Item("valfac").Value) * dbl_factmon) /dblSumFac)
         dblHonorarios = dblHonorarios / dblvalmer
        end if
      end if

      if dblcsce31 > 0 and cdbl(RsRep6.Fields.Item("valfac").Value) > 0 then
        if intCountValfac = 1 then
        'Complementarios normales para una cuenta de gastos con solo una referencia
        dblCompl = (cdbl(RsRep6.Fields.Item("valfac").Value) * dblcsce31) / dblvalmer
        else
        'Complementarios *  VALDOLITEM / SUMAGLOBAL EN DLS
         dblCompl = dblcsce31 * ((cdbl(RsRep6.Fields.Item("valfac").Value) * dbl_factmon) /dblSumFac)
         dblCompl = dblCompl / dblvalmer
        end if
      end if

      if SumaIva > 0 and cdbl(RsRep6.Fields.Item("valfac").Value) > 0 then
        if intCountValfac = 1 then
        'IVA normal para una cuenta de gastos con solo una referencia
        dblIVACG = (cdbl(RsRep6.Fields.Item("valfac").Value) * SumaIva) / dblvalmer
        else
        'IVA *  VALDOLITEM / SUMAGLOBAL EN DLS
         dblIVACG = SumaIva * ((cdbl(RsRep6.Fields.Item("valfac").Value) * dbl_factmon) /dblSumFac)
         dblIVACG = dblIVACG / dblvalmer
        end if
      end if

      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblHonorarios & "</font></td>" & chr(13) & chr(10) 'Prorrateo de Honorarios en base a la cantidad de producto
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblCompl & "</font></td>" & chr(13) & chr(10) 'Prorrateo de Complementarios
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblIVACG & "</font></td>" & chr(13) & chr(10) 'IVA de la Cuenta de Gastos

	  'Datos de Pagos Hechos
      strSQL7 = "SELECT E.conc21,E.deha21,sum(D.mont21) as Suma FROM D21PAGHE AS D, E21PAGHE AS E WHERE D.refe21='" & pRefer & "' and D.cgas21 ='"& strCtaGas &"' and D.fech21 = E.FECH21 AND D.foli21 = E.foli21 and E.TPAG21 = 1 and (E.tmov21='P' or E.esta21='A') group by E.conc21 ,E.deha21 ORDER BY E.conc21"
      Set RsRep7 = Server.CreateObject("ADODB.Recordset")
			RsRep7.ActiveConnection = MM_EXTRANET_STRING
			RsRep7.Source = strSQL7
			RsRep7.CursorType = 0
			RsRep7.CursorLocation = 2
			RsRep7.LockType = 1
			RsRep7.Open()

      intReg = 0
			if not RsRep7.eof then
			  intReg = 1
        redim arrCodPH(intReg)
			  redim arrImpPH(intReg)
				while not RsRep7.eof
         LoEncontro = false
			   redim preserve arrCodPH(intReg +1)
			   redim preserve arrImpPH(intReg +1)
			   for intI = 1 to intReg
			    if cstr(arrCodPH(intI)) = cstr(RsRep7.fields("conc21").value)  then
				    arrImpPH(intI) = cdbl(arrImpPH(intI)) - cdbl(RsRep7.fields("Suma").value)
            if arrImpPH(intI)  = 0 then
               arrCodPH(intI) = ""
               arrImpPH(intI) = ""
            end if
				  LoEncontro = true
			   end if
			  next
			  if not LoEncontro then
			     arrCodPH(intReg) = RsRep7.fields("conc21").value
			     arrImpPH(intReg) = RsRep7.fields("Suma").value
			  end if
        intReg = intReg + 1
				RsRep7.movenext
				wend
			end if
			RsRep7.close
			Set RsRep7 = nothing

      if  dblpiva31 > 0 then
        strSQL7 = "SELECT E.conc21,E.deha21,sum(D.mont21/" &  (cdbl(dblpiva31)/100) +1 & ") as Suma,sum(D.mont21) as SumaCA FROM D21PAGHE AS D, E21PAGHE AS E WHERE  D.refe21='" & pRefer & "' and D.cgas21 ='"& strCtaGas &"' and D.fech21 = E.FECH21 AND D.foli21 = E.foli21 and E.TPAG21 = 2 and (E.tmov21='P' or E.esta21='A') group by E.conc21 ,E.deha21 ORDER BY E.conc21"
      else
        strSQL7 = "SELECT E.conc21,E.deha21,sum(D.mont21/1.15) as Suma,sum(D.mont21) as SumaCA FROM D21PAGHE AS D, E21PAGHE AS E WHERE D.refe21='" & pRefer & "' and D.cgas21 ='"& strCtaGas &"' and D.fech21 = E.FECH21 AND D.foli21 = E.foli21 and E.TPAG21 = 2 and (E.tmov21='P' OR E.esta21='A') group by E.conc21 ,E.deha21 ORDER BY  E.conc21"
      end if

      Set RsRep7 = Server.CreateObject("ADODB.Recordset")
			RsRep7.ActiveConnection = MM_EXTRANET_STRING
			RsRep7.Source = strSQL7
			RsRep7.CursorType = 0
			RsRep7.CursorLocation = 2
			RsRep7.LockType = 1
			RsRep7.Open()
      intRegCA = 0
			if not RsRep7.eof then
        intRegCA = 1
        redim arrCodPHCA(intRegCA)
			  redim arrImpPHCA(intRegCA)
				while not RsRep7.eof
         LoEncontroCA = false
			   redim preserve arrCodPHCA(intRegCA +1)
			   redim preserve arrImpPHCA(intRegCA +1)
			   for intI = 1 to intRegCA
			    if cstr(arrCodPHCA(intI)) = cstr(RsRep7.fields("conc21").value)  then
				    arrImpPHCA(intI) = cdbl(arrImpPHCA(intI)) - cdbl(RsRep7.fields("Suma").value)
            if arrImpPHCA(intI)  = 0 then
               arrCodPHCA(intI) = ""
               arrImpPHCA(intI) = ""
            end if
				  LoEncontroCA = true
			   end if
			  next
			  if not LoEncontroCA then
			     arrCodPHCA(intReg) = RsRep7.fields("conc21").value
			     arrImpPHCA(intReg) = RsRep7.fields("Suma").value
			  end if
        intRegCA = intRegCA + 1
				RsRep7.movenext
				wend
			end if
			RsRep7.close
			Set RsRep7 = nothing

     dblSumProrPH  = 0
     if intTotConcep > 0 then
        SumProrPH = 0
        For x = 1 To intTotConcep
          dblSumPH = 0
          dblProrPH = 0
          dblIVAPH = 0

          ' PAGOS HECHOS
          for y=1 to IntReg-1
           intCvePH = ""
           intCvePH = arrCodPH(y)
            if trim(cstr(arrCvePH(x))) = trim(cstr(intCvePH)) then
            'Se suma el monto del PH y se va acumulando para ese tipo de PH
             if not trim(cstr(arrImpPH(y))) = "" then
                dblSumPH = dblSumPH +  cdbl(arrImpPH(y))
             end if
            end if
          next
          ' COSTOS ADUANALES
          if x <= (intRegCA-1) and intRegCA > 0 then
            for y=1 to intRegCA - 1
              intCvePH = ""
              intCvePHCA = arrCodPHCA(y)
            if trim(cstr(arrCvePH(x))) = trim(cstr(intCvePHCA)) then
            'Se suma el monto del PH y se va acumulando para ese tipo de PH
             if not trim(cstr(arrImpPHCA(y))) = "" then
                dblSumPH = dblSumPH +  cdbl(arrImpPHCA(y))
             end if
            end if
            next
          end if

          if intCountValfac = 1 and dblSumPH > 0 then
            dblProrPH = ((cdbl(RsRep6.Fields.Item("valfac").Value) * dblSumPH)/ dblvalmer)
            dblIVAPH = (dblProrPH / ((dblpiva31/100) + 1)) * (dblpiva31/100)
            dblProrPH = dblProrPH  - dblIVAPH
          else
            if dblSumPH > 0 then
              dblProrPH = (dblSumFac * dblSumPH)/ dblvalmer
              dblIVAPH = (dblProrPH / ((dblpiva31/100) + 1)) * (dblpiva31/100)
              dblProrPH = dblProrPH  - dblIVAPH
            end if
          end if

          'Acumulado de Prorrateos
          arrAcumProrPH(x) = arrAcumProrPH(x) + dblProrPH

          'Acumulado de Prorrateos de IVA de cada PH
          arrAcumIVAPH(x) = arrAcumIVAPH(x) + dblIVAPH

          'Suma de Prorrateos
          SumProrPH = SumProrPH + dblProrPH + dblIVAPH

          if dblProrPH > 0 then
            strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblProrPH & "</font></td>" & chr(13) & chr(10) 'Prorrateo de PH en base a la cantidad de producto
          else
            strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Prorrateo de PH en base a la cantidad de producto
          end if

          if dblProrPH > 0 and dblIVAPH > 0 then
            strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblIVAPH  & "</font></td>" & chr(13) & chr(10)'IVA de PH
          else
            strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA de PH
          end if

			Next

      dblImporteCG = 0
      dblImporteCG = SumProrPH + dblHonorarios + dblCompl + dblIVACG 'Importe Total de la Cuenta de gastos prorrateada

      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblImporteCG  & "</font></td>" & chr(13) & chr(10)'Suma de Prorrateos

     else
     	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Prorrateo de PH en base a la cantidad de producto
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA de PH
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Importe Total de la Cuenta de gastos prorrateada
     end if

		else
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Cuenta de Gasto
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Prorrateo de Honorarios en base a la cantidad de producto
			strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'Complemetarios prorrateados
      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10) 'IVA de la Cuenta de Gastos prorrateado

      if not arrConcepto(1) = "" then
        For i = 1 To intTotConcep
          strHTML = strHTML & "<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Prorrateo de PH
          strHTML = strHTML & "<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA de PH prorrateado
        Next
      end if

      strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Importe Total de la Cuenta de gastos prorrateada
		end if

    'Rutina que determina si la fraccion en la que estamos es la misma que la anterior
    'Sino es la misma entonces manda a llamar a la funcion de DespliegaFracc, que pinta el resumen por Fraccion
      'Si es el primer registro de la referencia
      if intCountFracc = 1 then
        strFracc = RsRep6.Fields.Item("frac05").Value 'Guarda Fraccion
        strFact = RsRep6.Fields.Item("fact05").Value 'Guarda Factura
        'Si la referencia solo tiene un registro
        if (RsRep6.RecordCount = 1 ) then
          strEntra ="entra"
          strHTML = strHTML & DespliegaFracc(strFracc,dblResFraccVF,dblResADV,dblResIVA,dblResDTA,dblResPRV)
          strHTML = strHTML & DespliegaFact(strFact,dblResFactVF)
          dblResFactVF = 0
          dblResFraccVF = 0
          dblResADV = 0
          dblResIVA = 0
          dblResDTA = 0
          dblResPRV = 0
        end if
      end if

    intCountFracc = intCountFracc + 1 'Incrementa el registro de Fraccion
    intContReg = intContReg + 1	'Incrementa el numero de registro
    RsRep6.movenext
	Wend

  'Cuando son varios registros y todos son iguales, no entra a ninguna de las condiciones (abarca la ultima)
  if strEntra = "" then
    strEntra = "Uno"
    strHTML = strHTML & DespliegaFracc(strUltimaFracc,dblResFraccVF,dblResADV,dblResIVA,dblResDTA,dblResPRV)
    strHTML = strHTML & DespliegaFact(strUltimaFact,dblResFactVF)
    dblResFactVF = 0
    dblResFraccVF = 0
    dblResADV = 0
    dblResIVA = 0
    dblResDTA = 0
    dblResPRV = 0
  else
    'Cuando son varios registros y ha entrado a DespliegaFact pero este es el ultimo registro (abarca la ultima)
    if intCountFracc > 1 then
      strEntra = "Ultima"
      strHTML = strHTML & DespliegaFracc(strUltimaFracc,dblResFraccVF,dblResADV,dblResIVA,dblResDTA,dblResPRV)
      strHTML = strHTML & DespliegaFact(strUltimaFact,dblResFactVF)
      dblResFactVF = 0
      dblResFraccVF = 0
      dblResADV = 0
      dblResIVA = 0
      dblResDTA = 0
      dblResPRV = 0
    end if
  end if
else
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Proveedor de producto
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Descripcion de Proveedor de producto
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Factura de proveedor
  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Fraccion
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Codigo de producto
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Descripcion de producto
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Cantidad
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Unidad de medida
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Valor factura
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Moneda
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'IGI
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'DTA
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'PRV
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Moneda
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Factura CG
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Honorarios
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Complemetarios
	strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA de la CG

  if not arrConcepto(1) = "" then
    For i = 1 To intTotConcep
      strHTML = strHTML & "<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Prorrateo de PH
      strHTML = strHTML & "<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA de PH
    Next
  end if

  strHTML = strHTML&"<td width=""90"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Importe Total de la CG

end if

RsRep6.close
Set RsRep6 = Nothing

'Regresa el HTML de la referencia
DespliegaRepDesgMerc = strHTML

End Function

Function DespliegaRefer(pRefer)
'Pinta el registro de resumen para cada referencia
	strHTML = strHTML&"<tr>" & chr(13) & chr(10)
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & pRefer & "&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Referencia
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & ObservRect & "&nbsp;</font></td>" & chr(13) & chr(10) 'Observaciones
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & TieneFIN(pRefer) & "</font></td>" & chr(13) & chr(10) 'Status
  strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & strNumPed & "</font></td>" & chr(13) & chr(10) 'Numero de Pedimento
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & TieneNAFTA(pRefer) & "&nbsp;</font></td>" & chr(13) & chr(10) 'Indica si el pedimento tiene algun Tratado de TLC America del Norte
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & strFecPag & "</font></td>" & chr(13) & chr(10) 'Fecha de Pago
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Proveedor de producto
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Descripcion de Proveedor de producto
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Factura de proveedor
  strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Pedido
  strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Fraccion
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Codigo de producto
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Descripcion de producto
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Cantidad
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Unidad de medida
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblAcumValFac & "</font></td>" & chr(13) & chr(10)'Valor factura
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & strMoneda & "</font></td>" & chr(13) & chr(10)'Moneda
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblAcumADV & "</font></td>" & chr(13) & chr(10)'IGI
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblAcumDTA & "</font></td>" & chr(13) & chr(10)'DTA
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblAcumPRV & "</font></td>" & chr(13) & chr(10)'PRV
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblAcumIVA & "</font></td>" & chr(13) & chr(10)'IVA
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">MXP</font></td>" & chr(13) & chr(10)'Moneda
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & strCtaGas & "</font></td>" & chr(13) & chr(10)'Factura CG
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblchon31 & "</font></td>" & chr(13) & chr(10)'Honorarios
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblcsce31 & "</font></td>" & chr(13) & chr(10)'Complementarios
	strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & SumaIva & "</font></td>" & chr(13) & chr(10)'IVA de la CG

  if not arrConcepto(1) = "" then
    For i = 1 To intTotConcep
      strHTML = strHTML & "<td width=""90"" bgcolor =""#99FFCC"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & arrAcumProrPH(i) & "</font></td>" & chr(13) & chr(10)'Prorrateo de PH
      strHTML = strHTML & "<td width=""90"" bgcolor =""#99FFCC"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & arrAcumIVAPH(i) & "</font></td>" & chr(13) & chr(10)'IVA de PH
    arrAcumProrPH(x) = 0
    arrAcumIVAPH(x) = 0
    Next
  end if

  strHTML = strHTML&"<td width=""90"" bgcolor =""#99FFCC"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & dblchon31 + dblcsce31 + SumaIva + dblsuph31 & "</font></td>" & chr(13) & chr(10)'Importe Total de la CG
  strHTML = strHTML&"</tr>" & chr(13) & chr(10)

  DespliegaRefer = strHTML

End Function


Function DespliegaFact(pFactura,pValFac)
'Pinta el registro de resumen para cada factura
dim strHTML2
strHTML2 = ""

	strHTML2 = strHTML2&"<tr>" & chr(13) & chr(10)
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Referencia
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Observaciones
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Status
  strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Pedimento
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Indica si el pedimento tiene algun Tratado de TLC America del Norte
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">//</font></td>" & chr(13) & chr(10) 'Fecha de Pago
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Proveedor de producto
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Descripcion de Proveedor de producto
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & pFactura & "</font></td>" & chr(13) & chr(10)'Factura de proveedor
  strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Pedido
  strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Fraccion
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Codigo de producto
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Descripcion de producto
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Cantidad
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Unidad de medida
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & pValFac & "</font></td>" & chr(13) & chr(10)'Valor factura
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & strMoneda & "</font></td>" & chr(13) & chr(10)'Moneda
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IGI
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'DTA
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'PRV
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">MXP</font></td>" & chr(13) & chr(10)'Moneda
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Factura CG
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Honorarios
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Complementarios
	strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA de la CG

  if not arrConcepto(1) = "" then
    For i = 1 To intTotConcep
      strHTML2 = strHTML2 & "<td width=""90"" bgcolor =""#99CCFF"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Prorrateo de PH
      strHTML2 = strHTML2 & "<td width=""90"" bgcolor =""#99CCFF"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA de PH
    Next
  end if

  strHTML2 = strHTML2&"<td width=""90"" bgcolor =""#99CCFF"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Importe Total de la CG
  strHTML2 = strHTML2&"</tr>" & chr(13) & chr(10)

  DespliegaFact = strHTML2

End Function

Function DespliegaFracc(pFraccion,pValFac,pADV,pIVA,pDTA,pPRV)
'Pinta el registro de resumen para cada fraccion

dim strHTML3
strHTML3 = ""

	strHTML3 = strHTML3&"<tr>" & chr(13) & chr(10)
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Referencia
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Observaciones
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Status
  strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Pedimento
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Indica si el pedimento tiene algun Tratado de TLC America del Norte
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">//</font></td>" & chr(13) & chr(10) 'Fecha de Pago
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Numero de Proveedor de producto
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10) 'Descripcion de Proveedor de producto
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Factura de proveedor
  strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Pedido
  strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pFraccion&"</font></td>" & chr(13) & chr(10)'Fraccion
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Codigo de producto
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Descripcion de producto
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Cantidad
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Unidad de medida
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & pValFac & "</font></td>" & chr(13) & chr(10)'Valor factura
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & strMoneda & "</font></td>" & chr(13) & chr(10)'Moneda
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & pADV & "</font></td>" & chr(13) & chr(10)'IGI
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & pDTA & "</font></td>" & chr(13) & chr(10)'DTA
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & pPRV & "</font></td>" & chr(13) & chr(10)'PRV
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & pIVA & "</font></td>" & chr(13) & chr(10)'IVA
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">MXP</font></td>" & chr(13) & chr(10)'Moneda
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">&nbsp;</font></td>" & chr(13) & chr(10)'Factura CG
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Honorarios
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Complementarios
	strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF""  nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA de la CG

  if not arrConcepto(1) = "" then
    For i = 1 To intTotConcep
      strHTML3 = strHTML3 & "<td width=""90"" bgcolor =""#CCCCFF"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Prorrateo de PH
      strHTML3 = strHTML3 & "<td width=""90"" bgcolor =""#CCCCFF"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'IVA de PH
    Next
  end if

  strHTML3 = strHTML3&"<td width=""90"" bgcolor =""#CCCCFF"" nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">0</font></td>" & chr(13) & chr(10)'Importe Total de la CG
  strHTML3 = strHTML3&"</tr>" & chr(13) & chr(10)

  DespliegaFracc = strHTML3

End Function


%>
