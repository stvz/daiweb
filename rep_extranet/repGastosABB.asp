
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->


<%
' ESTE ASP ES EL SEGUNDO Y ES PARA ADMINISTRADORES
MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
Response.Buffer = TRUE


Response.Addheader "Content-Disposition", "attachment;"
Response.ContentType = "application/vnd.ms-excel"


Server.ScriptTimeOut=100000




strPermisos = Request.Form("Permisos")
permi = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
permi2 = PermisoClientesTabla("B",Session("GAduana") ,strPermisos,"clie31")




if not permi2 = "" then
  permi2 = "  and (" & permi2 & ") "
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")
if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if
if blnAplicaFiltro then
   permi2 = " AND B.clie31 =" & strFiltroCliente
end if
if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
   permi2 = ""
end if







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

'response.Write("Permisos="&permi)



strDateIni=""
strDateFin=""
strTipoPedimento= ""
strCodError = "0"

strDateIni=trim(request.Form("txtDateIni"))
strDateFin=trim(request.Form("txtDateFin"))
'*******************************************************
' Si es Impo o Expo
strTipoPedimento=trim(request.Form("rbnTipoDate"))
'*******************************************************

'***************************************************************************************************************
strDescripcion=trim(request.Form("txtDescripcion"))
strDateIni2=trim(request.Form("txtDateIni2"))
strDateFin2=trim(request.Form("txtDateFin2"))
strTipoPedimento2=trim(request.Form("rbnTipoDate2"))

strTipoFiltro=trim(request.Form("TipoFiltro"))

if not isdate(strDateIni) then
	strCodError = "5"
end if
if not isdate(strDateFin) then
	strCodError = "6"
end if
if strDateIni="" or strDateFin="" then
	strCodError = "1"
end if


if strCodError = "0" then

strHTML = ""
strDateIni=trim(request.Form("txtDateIni"))
strDateFin=trim(request.Form("txtDateFin"))
strTipoPedimento=trim(request.Form("rbnTipoDate"))
strUsuario = request.Form("user")
strTipoUsuario = request.Form("TipoUser")

tmpTipo = ""
strSQL = ""


if strTipoPedimento  = "1" then
   tmpTipo = "IMPORTACION"
   strSQL = "SELECT tipopr01, " & _
            "       valmer01, " & _
            "       factmo01, " & _
            "       p_dta101, " & _
            "       t_reca01, " & _
            "       i_dta101, " & _
            "       cvecli01, " & _
            "       refcia01, " & _
            "       fecpag01, " & _
            "       valfac01, " & _
            "       fletes01, " & _
            "       segros01, " & _
            "       cvepvc01, " & _
            "       tipcam01, " & _
            "       patent01, " & _
            "       numped01, " & _
            "       totbul01, " & _
            "       cveped01, " & _
            "       adusec01, " & _
            "       desf0101, " & _
            "       nompro01, " & _
            "       cvepod01, " & _
            "       nombar01, " & _
            "       tipopr01, " & _
            "       fecent01, " & _
            "       if(alea01=1,'ROJO',if(alea01=2, 'VERDE',' ' ) ) as semaforo, " & _
            "       CVEMTS01 as destino " & _
            "FROM ssdagi01 ,c01refer " & _
            "WHERE refcia01=refe01 and " & _
            "      fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND " & _
            "      fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & _
                   Permi & " and " & _
            "      firmae01 !='' " & _
            "order by refcia01"

'            "       formfa01 as incoterm  " & _

            StrSQLPH=" SELECT distinct conc21,desc21 " & _
                     " FROM SSDAGI01,                "  & _
                     "      d21paghe,                " & _
                     "      e21paghe,                " & _
                     "      c21paghe                 " & _
                     " WHERE refcia01                  =  d21paghe.refe21 and " & _
                     "       YEAR(d21paghe.fech21)     =  YEAR(e21paghe.fech21) and " & _
                     "       d21paghe.foli21           =  e21paghe.foli21 and " & _
                     "       d21paghe.tmov21           =  e21paghe.tmov21 and " & _
                     "       e21paghe.fech21           =  d21paghe.fech21 AND " & _
                     "       e21paghe.esta21          !=  'S' and " & _
                     "       conc21                    =  clav21 and              " & _
                     "       fecpag01                 >=  '"&FormatoFechaInv(strDateIni)&"' AND " & _
                     "       fecpag01                 <='"&FormatoFechaInv(strDateFin)&"' " & _
                             Permi & " and " & _
                     "       firmae01 !='' " & _
                     " ORDER BY CONC21  "



end if
if strTipoPedimento  = "2" then
   tmpTipo = "EXPORTACION"
   strSQL = "SELECT tipopr01, " & _
            "       ' ' as  valmer01, " & _
            "       factmo01, " & _
            "       p_dta101, " & _
            "       t_reca01, " & _
            "       i_dta101, " & _
            "       cvecli01, " & _
            "       refcia01, " & _
            "       fecpag01, " & _
            "       valfac01, " & _
            "       fletes01, " & _
            "       segros01, " & _
            "       cvepvc01, " & _
            "       tipcam01, " & _
            "       patent01, " & _
            "       numped01, " & _
            "       totbul01, " & _
            "       cveped01, " & _
            "       adusec01, " & _
            "       desf0101, " & _
            "       nompro01, " & _
            "       cvepod01, " & _
            "       nombar01, " & _
            "       tipopr01, " & _
            "       fecpre01, " & _
            "       if(alea01=1,'ROJO',if(alea01=2, 'VERDE',' ' ) ) as semaforo, " & _
            "       CVEMTS01 as destino " & _
            "FROM ssdage01 ,c01refer " & _
            "WHERE refcia01=refe01 and " & _
            "      fecpag01 >='"&FormatoFechaInv(strDateIni)&"' AND " & _
            "      fecpag01 <='"&FormatoFechaInv(strDateFin)&"' " & _
                   Permi & " and " & _
            " firmae01 !='' " & _
            "order by refcia01"

            StrSQLPH=" SELECT distinct conc21,desc21 " & _
                     " FROM SSDAGE01,                "  & _
                     "      d21paghe,                " & _
                     "      e21paghe,                " & _
                     "      c21paghe                 " & _
                     " WHERE refcia01                  =  d21paghe.refe21 and " & _
                     "       YEAR(d21paghe.fech21)     =  YEAR(e21paghe.fech21) and " & _
                     "       d21paghe.foli21           =  e21paghe.foli21 and " & _
                     "       d21paghe.tmov21           =  e21paghe.tmov21 and " & _
                     "       e21paghe.fech21           =  d21paghe.fech21 AND " & _
                     "       e21paghe.esta21          !=  'S' and " & _
                     "       conc21                    =  clav21 and              " & _
                     "       fecpag01                 >=  '"&FormatoFechaInv(strDateIni)&"' AND " & _
                     "       fecpag01                 <='"&FormatoFechaInv(strDateFin)&"' " & _
                             Permi & " and " & _
                     "       firmae01 !='' " & _
                     " ORDER BY CONC21  "

end if

 'response.Write("Query="&strSQL)



if not trim(strSQL)="" then
		Set RsRep = Server.CreateObject("ADODB.Recordset")
		RsRep.ActiveConnection = MM_EXTRANET_STRING
		RsRep.Source = strSQL
		RsRep.CursorType = 0
		RsRep.CursorLocation = 2
		RsRep.LockType = 1
		RsRep.Open()

	if not RsRep.eof then
  ' Comienza el HTML, se pintan los titulos de las columnas
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>REPORTE DE GASTOS DE " & tmpTipo & " </p></font></strong>"
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>GRUPO REYES KURI, S.C. </p></font></strong>"
	   strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif""><p>" & strDateIni & " al " & strDateFin & "</p></font></strong>"
     strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
	   strHTML = strHTML & "<tr bgcolor=""#006699"" align=""center"">"& chr(13) & chr(10)

    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Aduana                     </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Patente                    </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Pedimento                  </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""70"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Referencia                 </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de Pago              </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""100"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Tipo de Pedimento         </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Clave de Documento         </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Cuenta de Gastos           </font></td>" & chr(13) & chr(10)
		strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Fecha de la C.G.           </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Total Pagos Hechos         </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Servicios Complementarios  </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Honorarios                 </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA de la C.G.             </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Total de la C.G.           </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Anticipo de la C.G.        </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Saldo de la C.G.           </font></td>" & chr(13) & chr(10)

    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IVA                        </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FP IVA                     </font></td>" & chr(13) & chr(10)

    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">IGI                        </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FP IGI                     </font></td>" & chr(13) & chr(10)

    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">PREV                       </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">DTA                        </font></td>" & chr(13) & chr(10)
    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">FP DTA                     </font></td>" & chr(13) & chr(10)

    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">OTROS IMPUESTOS            </font></td>" & chr(13) & chr(10)

    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TOTAL IMPUESTOS            </font></td>" & chr(13) & chr(10)

    strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">TOTAL IMPUESTOS EN EFECTIVO</font></td>" & chr(13) & chr(10)

    ' Desglozar los pagos hechos
    	Set RsRepPHConceptos = Server.CreateObject("ADODB.Recordset")
		  RsRepPHConceptos.ActiveConnection = MM_EXTRANET_STRING
		  RsRepPHConceptos.Source = StrSQLPH
		  RsRepPHConceptos.CursorType = 0
		  RsRepPHConceptos.CursorLocation = 2
		  RsRepPHConceptos.LockType = 1
		  RsRepPHConceptos.Open()
      IntLonPH=0

      While NOT RsRepPHConceptos.EOF
         IntLonPH=IntLonPH + 1
         RsRepPHConceptos.movenext
      wend
      'RsRepPHConceptos.close

      Dim ConceptosPagosHechos()
      redim ConceptosPagosHechos(IntLonPH)

      IntIndice=0
      'RsRepPHConceptos.Open()
      RsRepPHConceptos.MoveFirst
      While NOT RsRepPHConceptos.EOF
         strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">" & RsRepPHConceptos.Fields.Item("desc21").Value & "</td>" & chr(13) & chr(10)
         ConceptosPagosHechos(IntIndice) = RsRepPHConceptos.Fields.Item("conc21").Value
         RsRepPHConceptos.movenext
         IntIndice=IntIndice + 1
      wend

      RsRepPHConceptos.close

      Set RsRepPHConceptos = Nothing

      'response.Write(ConceptosPagosHechos(0) )
      'response.Write(ConceptosPagosHechos(1) )

		  strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">Contacto</td>" & chr(13) & chr(10)
		  strHTML = strHTML & "<td width=""80"" nowrap><strong><font color=""#FFFFFF"" size=""2"" face=""Arial, Helvetica, sans-serif"">R.F.C. del Cliente</td>" & chr(13) & chr(10)
      strHTML = strHTML & "</tr>"& chr(13) & chr(10)


'*************************************************************

	While NOT RsRep.EOF

    'Se asigna el nombre de la referencia
     strRefer = RsRep.Fields.Item("refcia01").Value

     'Checa si la referencia es rectificada o rectificacion
     'Si es una de ellas lo almacena en ObservRect

     ObservRect = ""
     strRect = ""
     strRect = RegresaRect(strRefer,"Rectificado")
     if not strRect = ""  then
        ObservRect=strRect
     end if
     strRect = RegresaRect(strRefer,"Rectificacion")
     if not strRect = ""  then
        ObservRect=strRect
     end if

     dim intContReg
     dim intContCtas
     dim strCtaGas
	   dim intContCG

	   intContCtas=1
     intContReg=1
     strCtaGas=""
	   strfech31=""
	   dblsuph31=0
	   dblcoad31=0
	   dblcsce31=0
	   dblchon31=0
	   dblpiva31=0
	   dblanti31=0
	   dblsald31=0
     dbltotal31=0

        ' Aqui se buscan los datos de la cuenta de gastos
        strSQL1="select A.cgas31,B.fech31,B.suph31, B.coad31, B.csce31, B.chon31, B.piva31, B.anti31,B.sald31,B.TOTA31 from d31refer as A LEFT JOIN e31cgast as B ON A.cgas31=B.cgas31 where B.esta31='I' and A.refe31='"&strRefer&"' " & permi2


'response.Write(strSQL1)
'response.end

        Set RsRep1 = Server.CreateObject("ADODB.Recordset")
        RsRep1.ActiveConnection = MM_EXTRANET_STRING
        RsRep1.Source = strSQL1
        RsRep1.CursorType = 0
        RsRep1.CursorLocation = 2
        RsRep1.LockType = 1
        RsRep1.Open()

        'Si no tiene una cuenta de gastos, igual se mandan a desplegar los datos de la referencia
       if RsRep1.eof then
          strHTML=DespliegaRepDesgRef(strRefer, strCtaGas)

       else
        'Si tiene varias cuentas de gastos se repiten los datos de la referencia y se despliegan los distintos datos de las diferentes cuentas de gastos
          while not RsRep1.eof
           intContCG=1
           strCtaGas=RsRep1.Fields.Item("cgas31").Value
           strfech31=RsRep1.Fields.Item("fech31").Value
           dblsuph31=cdbl(RsRep1.Fields.Item("suph31").Value)
           dblcoad31=cdbl(RsRep1.Fields.Item("coad31").Value)
           dblcsce31=cdbl(RsRep1.Fields.Item("csce31").Value)
           dblchon31=cdbl(RsRep1.Fields.Item("chon31").Value)
           dblpiva31=cdbl(RsRep1.Fields.Item("piva31").Value)
           dblanti31=cdbl(RsRep1.Fields.Item("anti31").Value)
           dblsald31=cdbl(RsRep1.Fields.Item("sald31").Value)
           dbltotal31=cdbl(RsRep1.Fields.Item("TOTA31").Value)
           strHTML=DespliegaRepDesgRef(strRefer, strCtaGas)
           intContCtas=intContCtas + 1
           RsRep1.movenext
          wend
       end if
		RsRep1.close
        Set RsRep1=Nothing
   RsRep.movenext
   Wend

'************************

   strHTML = strHTML & "</table>"& chr(13) & chr(10)
   'Se pinta todo el HTML formado
   response.Write(strHTML)
   end if
   RsRep.close
   Set RsRep = Nothing




   if strHTML = "" then
      strHTML = "NO EXISTEN REGISTROS"
      response.Write(strHTML)
   end if
 else
   strHTML = "NO EXISTEN REGISTROS"
   response.Write(strHTML)
end if
%>




<%
'Funcion que va elaborando el reporte desglosado de referencias y devuelve el HTML
function DespliegaRepDesgRef(pRefer, pCtaGas)
      'Sus parametros son
      'pRefer    Referencia
      'pCtaGas  Cuenta de Gastos


     '  dim dblEfectivo
     '  dim dblOtros
     '   strSQL3="select fpagoi36, " & _
     '           "        import36 " & _
     '           "from sscont36    " & _
     '           "where refcia36='"&pRefer&"'"

		 '	Set RsRep3 = Server.CreateObject("ADODB.Recordset")
		 '	RsRep3.ActiveConnection = MM_EXTRANET_STRING
		 '	RsRep3.Source = strSQL3
		 '	RsRep3.CursorType = 0
		 '	RsRep3.CursorLocation = 2
		 '	RsRep3.LockType = 1
		 '	RsRep3.Open()
     '  if not RsRep3.eof then
     '  ' Aqui se obtienen los campos de Suma de Efectivo y Suma de Otros
     '    dblEfectivo=0
     '    dblOtros=0
     '     while not RsRep3.eof
     '      If cdbl(RsRep3.Fields.Item("fpagoi36").Value)=0 then
     '          dblEfectivo=dblEfectivo+cdbl(RsRep3.Fields.Item("import36").Value)  'Sumamos el efectivo para una referencia
     '      else
     '          dblOtros=dblOtros+cdbl(RsRep3.Fields.Item("import36").Value)         'Sumamos los otros conceptos para una referencia
     '      end if
		 '	  RsRep3.movenext
     '     wend
     '     RsRep3.close
     '     Set RsRep3=Nothing
     '  end if

      '*************************************************************

                           StrImpContrCC    = 0
                           StrImpContrTotal = 0
                           padv_igi         = 0
                           pprv             = 0
                           pdta             = 0
                           piva             = 0

                           FPpadv_igi       = 0
                           FPpdta           = 0
                           OtrosImpuestos   = 0

                           ' facturas comerciales
                           sqlImpContr = " SELECT SUM(import36) as Impuestos,                  " & _
                                         "        SUM(IF(cveimp36=3,import36,0) )     AS IVA,  " & _
                                         "        SUM(IF(cveimp36=3,FPAGOI36,0) )     AS FPIVA,  " & _
                                         "        SUM(IF(cveimp36=6,import36,0) )     AS ADV,  " & _
                                         "        SUM(IF(cveimp36=6,FPAGOI36,0) )     AS FPADV," & _
                                         "        SUM(IF(cveimp36=1,import36,0) )     AS DTA,  " & _
                                         "        SUM(IF(cveimp36=1,FPAGOI36,0) )     AS FPDTA," & _
                                         "        SUM(IF(cveimp36=15,import36,0) )    AS PRV,  " & _
                                         "        SUM(IF(cveimp36<>3 AND cveimp36<>6 AND cveimp36<>15 AND cveimp36<>1,import36,0) )  AS otrosimp,   " & _
                                         "        SUM(IF(cveimp36=2,import36,0) )     AS CC,   " & _
                                         "        SUM(IF(cveimp36=4,import36,0) )     AS ISAN, " & _
                                         "        SUM(IF(cveimp36=5,import36,0) )     AS IEPS, " & _
                                         "        SUM(import36) AS TOTAL                    " & _
                                         " FROM sscont36                                    " & _
                                         " WHERE REFCIA36 = '"&ltrim(pRefer)&"' " & _
                                         " GROUP BY refcia36   "



                           'Response.Write(sqlImpContr)
                           'Response.End
                           set RSImpContr = server.CreateObject("ADODB.Recordset")
                           RSImpContr.ActiveConnection = MM_EXTRANET_STRING
                           RSImpContr.Source= sqlImpContr
                           RSImpContr.CursorType = 0
                           RSImpContr.CursorLocation = 2
                           RSImpContr.LockType = 1
                           RSImpContr.Open()
                           if not RSImpContr.eof then
                                    StrImpContrCC    = RSImpContr.fields.item("CC").value
                                    padv_igi         = RSImpContr.fields.item("ADV").value
                                    pdta             = RSImpContr.fields.item("DTA").value
                                    piva             = RSImpContr.fields.item("IVA").value
                                    pprv             = RSImpContr.fields.item("PRV").value
                                    StrImpContrTotal = RSImpContr.fields.item("TOTAL").value
                                    FPpadv_igi       = RSImpContr.fields.item("FPADV").value
                                    FPpdta           = RSImpContr.fields.item("FPDTA").value
                                    FPpiva           = RSImpContr.fields.item("FPIVA").value
                                    OtrosImpuestos   = RSImpContr.fields.item("otrosimp").value
                            end if
                           RSImpContr.close
                           set RSImpContr = nothing
            '**************************************************************************






      'strSQLIncoterm=" SELECT refcia39, " & _
      '               "      terfac39    " & _
      '               " FROM ssfact39    " & _
      '               " WHERE refcia39='"&pRefer&"'"

			'Set RsRepIncoterm = Server.CreateObject("ADODB.Recordset")
			'RsRepIncoterm.ActiveConnection = MM_EXTRANET_STRING
			'RsRepIncoterm.Source = strSQLIncoterm
			'RsRepIncoterm.CursorType = 0
			'RsRepIncoterm.CursorLocation = 2
			'RsRepIncoterm.LockType = 1
			'RsRepIncoterm.Open()
      'StrIncoterm=" "

      'if not RsRepIncoterm.eof then
      '    StrIncoterm=RsRepIncoterm.Fields.Item("terfac39").Value
      'end if
      'RsRepIncoterm.close
      'Set RsRepIncoterm=Nothing


      '***************************************************************************************************
             sqlImpContrEfec = " SELECT SUM(import36) AS TOTAL         " & _
                               " FROM sscont36                         " & _
                               " WHERE REFCIA36 = '"&ltrim(pRefer)&"'  " & _
                               " AND FPAGOI36 = '0' " & _
                               " GROUP BY refcia36   "
             'Response.Write(sqlImpContrEfec)
             'Response.End
             set RSImpContrEfec = server.CreateObject("ADODB.Recordset")
             RSImpContrEfec.ActiveConnection = MM_EXTRANET_STRING
             RSImpContrEfec.Source = sqlImpContrEfec
             RSImpContrEfec.CursorType = 0
             RSImpContrEfec.CursorLocation = 2
             RSImpContrEfec.LockType = 1
             RSImpContrEfec.Open()
             if not RSImpContrEfec.eof then
                  StrImpContrTotalEfec = RSImpContrEfec.fields.item("TOTAL").value
              end if
             RSImpContrEfec.close
             set RSImpContrEfec = nothing
      '********************************************************************************************************








            'Recorremos las Fracciones
            Strpatente=RsRep.Fields.Item("patent01").Value
            'Patente
            'Pedimento
            'Referencia
            'Fecha de Pago
            'Tipo de Pedimento
            'Clave de Documento
            'Cuenta de Gastos
		        'Fecha de la C.G.
            'Total Pagos Hechos
            'Servicios Complementarios
            'Honorarios
            'IVA de la C.G.
            'Subtotal de la C.G.
            'Anticipo de la C.G.
            'Total de la C.G.
            'IVA
            'IGI
            'PREV
            'DTA
            'TOTAL IMPUESTOS

            strHTML = strHTML&"<tr>" & chr(13) & chr(10)
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("adusec01").Value    &"</font></td>" & chr(13) & chr(10) 'Clave de aduana
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("patent01").Value    &"</font></td>" & chr(13) & chr(10) 'Patente
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("numped01").Value    &"</font></td>" & chr(13) & chr(10) 'Numero de Pedimento
				    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pRefer                                 &"</font></td>" & chr(13) & chr(10)  'Referencia
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("fecpag01").Value    &"</font></td>" & chr(13) & chr(10) 'Fecha de pago
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&TipoOper(RsRep.Fields.Item("tipopr01").Value)&"&nbsp;</font></td>" & chr(13) & chr(10) 'Tipo de Pedimento
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&RsRep.Fields.Item("cveped01").Value    &"</font></td>" & chr(13) & chr(10) 'Clave de pedimento
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pCtaGas                                &"</font></td>" & chr(13) & chr(10) 'Cuenta de Gastos
  		      strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&strfech31                              &"</font></td>" & chr(13) & chr(10) 'Fecha de la C.G.
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsuph31 + dblcoad31                  &"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
		        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblcsce31                              &"</font></td>" & chr(13) & chr(10) 'Servicios Complemetarios
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblchon31                              &"</font></td>" & chr(13) & chr(10) 'Honorarios
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&(dblcsce31 + dblchon31)*(dblpiva31/100)&"</font></td>" & chr(13) & chr(10) 'IVA de la C.G.
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dbltotal31                             &"</font></td>" & chr(13) & chr(10) 'Total de la C.G.
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblanti31                              &"</font></td>" & chr(13) & chr(10) 'Anticipos
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&dblsald31                              &"</font></td>" & chr(13) & chr(10) 'Saldo de la C.G.
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&piva                                   &"</font></td>" & chr(13) & chr(10) 'IVA

            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&FPpiva                                 &"</font></td>" & chr(13) & chr(10) 'FORMA DE PAGO IVA

            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&padv_igi                               &"</font></td>" & chr(13) & chr(10) 'IGI

            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&FPpadv_igi                             &"</font></td>" & chr(13) & chr(10) 'FORMA DE PAGO IGI

            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pprv                                   &"</font></td>" & chr(13) & chr(10) 'PREV
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&pdta                                   &"</font></td>" & chr(13) & chr(10) 'DTA
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&FPpdta                                 &"</font></td>" & chr(13) & chr(10) 'FORMA DE PAGO DTA
            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&OtrosImpuestos                         &"</font></td>" & chr(13) & chr(10) 'OTROS IMPUESTOS



            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&StrImpContrTotal                       &"</font></td>" & chr(13) & chr(10) 'Total Impuestos

            strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&StrImpContrTotalEfec                   &"</font></td>" & chr(13) & chr(10) 'Total Impuestos en Efectivo



             '--------------------------------------------
             ' Desglozar los pagos hechos
             '--------------------------------------------

              'response.Write(ConceptosPagosHechos(0))
              'response.Write(ConceptosPagosHechos(1))
              'response.Write("PagosHechos")
              'for each valor in ConceptosPagosHechos
             '    response.Write("PH"&valor)
              'next

            if intContReg=1 then
               StrQryConcppH= ""
               StrQryConcppH= " SELECT  "
               for inti=0 to (IntLonPH-1)
                 '       strHTML = strHTML & "<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">" & ConceptosPagosHechos(inti) & "</FONT></td>" & chr(13) & chr(10) 'Pagos hechos desglozados
                  IF inti=0 THEN
                     StrQryConcppH = StrQryConcppH & "sum(if(conc21="& ConceptosPagosHechos(inti) &",(if(trim(e21paghe.deha21)='A',d21paghe.mont21,(d21paghe.mont21)*-1)  ),0 ) ) as var"&CStr(inti)
                  ELSE
                     StrQryConcppH = StrQryConcppH & "," & "sum(if(conc21="& ConceptosPagosHechos(inti) &",(if(trim(e21paghe.deha21)='A',d21paghe.mont21,(d21paghe.mont21)*-1)  ),0 ) ) as var"&CStr(inti)
                  END IF
               next
               StrQryConcppH = StrQryConcppH & "  from  d21paghe, e21paghe " & _
                                               "  where  d21paghe.refe21  = '" &  pRefer & "' AND " & _
                                               "         YEAR(d21paghe.fech21) = YEAR(e21paghe.fech21)  and " & _
                                               "         e21paghe.foli21 = d21paghe.foli21   and " & _
                                               "     e21paghe.tmov21      = d21paghe.tmov21   and " & _
                                               "     e21paghe.fech21      = d21paghe.fech21   AND " & _
                                               "     TRIM(e21paghe.esta21) <> 'S'             AND " & _
                                               "     cgas21                <>''               and " & _
                                               "     cgas21='"& pCtaGas &"' "   & _
                                               "  group by cgas21                        "


                                         '      "         d21paghe.foli21 =  e21paghe.foli21  AND " & _
                                         '      "         d21paghe.tmov21 =  e21paghe.tmov21  AND " & _
                                         '      "         e21paghe.esta21 <> 'S'    " & _
                                         '      "  group by d21paghe.refe21  "



               'response.Write(StrQryConcppH& chr(13) & chr(10))
               Set RsConcppHdesg = Server.CreateObject("ADODB.Recordset")
			         RsConcppHdesg.ActiveConnection = MM_EXTRANET_STRING
			         RsConcppHdesg.Source = StrQryConcppH
			         RsConcppHdesg.CursorType = 0
			         RsConcppHdesg.CursorLocation = 2
			         RsConcppHdesg.LockType = 1
			         RsConcppHdesg.Open()
               inti=0
			         if not RsConcppHdesg.eof then
                  for inti=0 to (IntLonPH-1)
                    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& RsConcppHdesg.Fields.Item("var"&inti).Value &"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
                  next
               else
                  'Response.Write("No tiene Registros")
                  for inti=0 to (IntLonPH-1)
                    strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& "0" &"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
                  next
               end if
               RsConcppHdesg.close
               set RsConcppHdesg = Nothing
            else
               for inti=0 to (IntLonPH-1)
                 strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& "0" &"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
               next
            end if



              'While not RsConcppHdesg.eof
              '   strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"& RsConcppHdesg.Fields.Item("var"&inti).Value &"</font></td>" & chr(13) & chr(10) 'Pagos Hechos
              '   inti=inti+1
              '   RsConcppHdesg.movenext
              'wend
              '  select
              '             sum(if(conc21=1,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)  ),0 )      ) as impuesto,
              '              sum(if(conc21=2,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as Maniobras,
              '              sum(if(conc21=3,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as Almacenajes,
              '              sum(if(conc21=5,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as FleteTerrestre,
              '              sum(if(conc21=6,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as DemorasPorContenedor,
              '              sum(if(conc21=16,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as fumigaciones,
              '              sum(if(conc21=41,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as pagoLiberacion,
              '              sum(if(conc21=70,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as LimpiezaContenedores,
              '              sum(if(conc21=86,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ManiobrasylimpiezaCont,
              '              sum(if(conc21=100,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as Revalidacion,
              '              sum(if(conc21=107,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ReparacionCont,
              '              sum(if(conc21=111,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ManiobrasyAlmacenajes,
              '              sum(if(conc21=130,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ReparacionContenedor,
              '              sum(if(conc21=151,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as Prevalidacion,
              '              sum(if(conc21=153,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as ProcesamientoElectronicoDatos,
              '              sum(if(conc21=181,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as maniobrasyMuellajes,
              '              sum(if(conc21=183,(if(e21paghe.deha21='A',d21paghe.mont21,(d21paghe.mont21)*-1)),0 ) ) as almacenajemaniobrasymuellajes
              ' from  d21paghe, e21paghe
              ' where  d21paghe.refe21               =  'RKU06-03322' and
              '             YEAR(d21paghe.fech21)    =  YEAR(e21paghe.fech21) AND
              '             d21paghe.foli21                  =  e21paghe.foli21 AND
              '             d21paghe.tmov21               =  e21paghe.tmov21 and
              '             e21paghe.esta21                <>  'S'
              '      group by d21paghe.refe21


		        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"repcli18")&"</font></td>" & chr(13) & chr(10) 'Contacto
		        strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif"">"&CampoCliente(RsRep.Fields.Item("cvecli01").Value,"rfccli18")&"</font></td>" & chr(13) & chr(10) 'R.F.C. del Cliente
            strHTML = strHTML&"</tr>"& chr(13) & chr(10)




      'Regresa el HTML del reporte
      DespliegaRepDesgRef=strHTML
end function







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
<br>
<%end if%>


