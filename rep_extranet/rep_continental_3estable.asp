<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

 <style type="text/css">
.style20 {color: #FFFFFF}
 </style>

<% 
Response.Buffer = TRUE
Response.Addheader "Content-Disposition", "attachment;filename=RepContinental.xls"
Response.ContentType = "application/vnd.ms-excel"
Server.ScriptTimeOut=100000

STRFINI=request.form("txtDateIni")
STRFFIN=request.form("txtDateFin")
strTipo=request.Form("rbnTipoDate")

if isdate(STRFINI) and isdate(STRFFIN) AND DateDiff("d",STRFINI,STRFFIN)>=0 then
	DiaI = cstr(datepart("d",STRFINI))
	Mesi = cstr(datepart("m",STRFINI))
	AnioI = cstr(datepart("yyyy",STRFINI))
	ISTRFINI = Anioi&"/"&Mesi&"/"&Diai
	DiaF = cstr(datepart("d",STRFFIN))
	MesF = cstr(datepart("m",STRFFIN))
	AnioF = cstr(datepart("yyyy",STRFFIN))
	FSTRFFIN = AnioF&"/"&MesF&"/"&DiaF
else
    Response.Write("VERIFIQUE SUS FECHAS")
	' Response.End
end if

strTipoUsuario = request.Form("TipoUser")
strPermisos    = Request.Form("Permisos")
permi  = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01")
'Response.Write(permi)
'Response.End

if not permi = "" then
	permi = "  and (" & permi & ") "
end if
cvecli = request.Form("txtCliente")

if cvecli="Todos" then
	nomb=cvecli
end if

AplicaFiltro = false
strFiltroCliente = ""
strFiltroCliente = request.Form("txtCliente")

if not strFiltroCliente= "" and not strFiltroCliente  = "Todos" then
   blnAplicaFiltro = true
end if

if blnAplicaFiltro then
   permi = " AND cvecli01 =" & strFiltroCliente
end if

if  strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
  permi = ""
end if


'Response.Write(permi)
'Response.End


    if cvecli<>"Todos" then
        'MM_EXTRANET_STRING = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER="& IPHost &"; DATABASE=rku_extranet; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
        MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
		' Response.Write(MM_EXTRANET_STRING)
		' Response.End()
        Set RsRevisa = Server.CreateObject("ADODB.Recordset")
        RsRevisa.ActiveConnection = MM_EXTRANET_STRING
        strSQL=  "SELECT cvecli18, " & _
                 "NOMCLI18 AS NOMBRE " & _
                 "FROM SSCLIE18 " & _
                 "where cvecli18 = '" & cvecli &"' "
                 '"SELECT cvecli18,NOMCLI18 AS NOMBRE FROM SSCLIE18 where cvecli18='"&cvecli&"' " & _

        RsRevisa.Source = strSQL
        RsRevisa.CursorType = 0
        RsRevisa.CursorLocation = 2
        RsRevisa.LockType = 1
        RsRevisa.Open()
        if not RsRevisa.eof then
			' Response.Write("NO ESTA VACIO")
            nomb =  RsRevisa.Fields.Item("nombre").Value
          end if
        RsRevisa.close
        set RsRevisa = nothing
    end if

    '***********aqui empiezan los querys del reporte
     MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
     Set Conn = Server.CreateObject ("ADODB.Connection")
     Set REFE = Server.CreateObject ("ADODB.RecordSet")
     Conn.Open MM_EXTRANET_STRING

    If strTipo=1 then 'importacion --chekar si fecent01 aparece al = modo01='T' xeso no lo pues en expo

          STRSQL= " select refcia01,                   " &_
                  "        year(fecpag01) as Anio,     " &_
                  "        patent01,                   " &_
                  "        numped01,                   " &_
                  "        feta01,                     " &_
                  "        fecpag01,                   " &_
                  "        cveped01,                   " &_
                  "        '1' as Tipo,                " &_
                  "        embala01,                   " &_
                  "        tipcam01,                   " &_
                  "        cveadu01,                   " &_
                  "        valmer01,                   " &_
                  "        factmo01,                   " &_
                  "        fletes01,                   " &_
                  "        segros01,                   " &_
                  "        incble01,                   " &_
                  "        cvepod01,                   " &_
                  "        fecent01,                   " &_
                  "        regime01,                   " &_
				  "        pesobr01,                   " &_
				  "        otros01,                    " &_
                  "        nombar01                    " &_
                  " from ssdagi01 , c01refer           " &_
                  " where refcia01=refe01              " &_
                  "       and fecpag01>='"&ISTRFINI&"' " &_
                  "       and fecpag01<='"&FSTRFFIN&"' " &_
                  "       and firmae01<>''             " &_
                  "       and cveped01 <>'R1'          " &_
                  "       and modo01='T' "&permi&" "&_
				  " ORDER BY refcia01 "

    'Response.Write(STRSQL)
    'Response.End
    'STRSQL= " select refcia01,year(fecpag01) as Anio,patent01,numped01,feta01,fecpag01,cveped01,'1' as Tipo,embala01, " &_
    '        " tipcam01,cveadu01,valmer01,factmo01,fletes01,segros01,otros01,cvepod01,fecent01,regime01,nombar01 " &_
    '        " from ssdagi01 , c01refer  " &_
    '    " where refcia01=refe01 and fecpag01>='"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"' and firmae01<>'' and cveped01 <>'R1' " &_
    '    " and modo01='T' "&permi&" "

    else

          STRSQL= " select refcia01,                     " &_
                  "        year(fecpag01) as Anio,       " &_
                  "        patent01,                     " &_
                  "        numped01,                     " &_
                  "        feta01,                       " &_
                  "        fecpag01,                     " &_
                  "        cveped01,                     " &_
                  "        '2' as Tipo,                  " &_
                  "        embala01,                     " &_
                  "        tipcam01,                     " &_
                  "        valfac01,                     " &_
                  "        cveadu01,                     " &_
                  "        factmo01,                     " &_
                  "        fletes01,                     " &_
                  "        segros01,                     " &_
                  "        incble01,                      " &_
                  "        cvepod01,                     " &_
                  "        regime01,                     " &_
                  "        pesobr01,                     " &_
				  "        otros01,                   	 " &_
                  "        nombar01                      " &_
                  " from ssdage01 , c01refer             " &_
                  " where refcia01=refe01                " &_
                  "        and fecpag01>='"&ISTRFINI&"'  " &_
                  "        and fecpag01<='"&FSTRFFIN&"'  " &_
                  "        and firmae01<>''              " &_
                  "        and cveped01 <>'R1'           " &_
                  " "&permi&" "&_
				  " ORDER BY refcia01 "

    'STRSQL= " select refcia01,year(fecpag01) as Anio,patent01,numped01,feta01,fecpag01,cveped01,'2' as Tipo,embala01, " &_
    '        " tipcam01,valfac01,cveadu01,factmo01,fletes01,segros01,otros01,cvepod01,regime01,nombar01 " &_
    '        " tipcam01,pesobr01,regime01 from ssdage01 , c01refer  " &_
    '    " where refcia01=refe01 and fecpag01>='"&ISTRFINI&"' and fecpag01<='"&FSTRFFIN&"' and firmae01<>'' and cveped01 <>'R1' " &_
    '    " "&permi&" "

    '    Response.Write(STRSQL)
    '    Response.End

    end if

  Set REFE= Conn.Execute(strSQL)
 'response.Write(strsql)
 'response.end()
  If strTipo=1 then
    tipoper="Importacion"
 else
    tipoper="Exportacion"
 end if

  %>


 <strong><font color="#000000" size="3" face="Arial, Helvetica, sans-serif"><p>Reporte de Anexo 24 de <%=tipoper%> </p></font></  strong>
 <strong><font color="#000000" size="3" face="Arial, Helvetica, sans-serif"><p></p></font></strong>
 <strong><font color="#000000" size="3" face="Arial, Helvetica, sans-serif"><p> Del <%=STRFINI%> al <%=STRFFIN%> </p></font></  strong>

    <!--
        PEDIMENTO
        FECHA PAGO MM/DD/AA
        FECHA ENTRADA MM/DD/AA
        ADUANA
        CLAVE DE PEDIMENTO
        País de Origen
        Clave Pais de Origen
        País de Venta
        Clave País de Venta
        Referencia
        DESCONSOLIDADORA
        GUIA MASTER
        GUIA  HOUSE
        FECHA FACTURA   DD/MM/AA
        NUM. FACTURA
        PROVEEDOR
        NUM MATERIAL CONTINENTAL
        DESCRIPCION DE MATERIAL
        UNIDAD DE MEDIDA
        FRACCION ARANCELARIA
        TASA (IMMEX, PROSEC)
        PZAS FACTURA
        VALOR MONEDA EXTRANJERA
        MONEDA
        FACTOR MONEDA EXTRANJERA
        VALOR USD	ç
        TIPO DE CAMBIO
        FACTOR AJUSTE
        VALOR COMERCIAL
        VALOR ADUANA
        PRECIO UNITARIO USD
        PRECIO UNITARIO MNX
        TASA IGI (SIN IMMEX)
        DTA
        PROGRAMA APLICADO (IMMEX-PROSEC)
        SECTOR PROSEC
        TASA PROSEC
        FORMA DE PAGO
    -->

    <table align="left">
      <tr>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PEDIMENTO	                       </b></FONT></td>
		 <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PATENTE	                           </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> FECHA PAGO MM/DD/AA                 </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> FECHA ENTRADA MM/DD/AA	           </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> ADUANA                              </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> CLAVE DE PEDIMENTO                  </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> País de Origen                      </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> Clave Pais de Origen                </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> País de Venta                       </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> Clave País de Venta                 </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> Referencia                          </b></FONT></td>
		 <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PESO BRUTO					       </b></FONT></td>
		 <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> TOTAL DE INCREMENTABLES             </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> DESCONSOLIDADORA                    </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> GUIA MASTER                         </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> GUIA  HOUSE                         </b></FONT></td>
		 
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> CR                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> FI                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> FR                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> IC                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> IM                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> MS                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PC                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PD                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PP                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PS                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> RC                                   </b></FONT></td>
		<td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> RO                                   </b></FONT></td>
		
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> FECHA FACTURA   DD/MM/AA            </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> NUM. FACTURA                        </b></FONT></td>
		 <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> INCOTERM                            </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PROVEEDOR                           </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> NUM MATERIAL CONTINENTAL            </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> DESCRIPCION DE MATERIAL             </b></FONT></td>
		 <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> DESCRIPCION DE LA OBSERVACION       </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> UNIDAD DE MEDIDA                    </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> FRACCION ARANCELARIA                </b></FONT></td>
		 <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> NO. PARTIDA		                   </b></FONT></td>
		 <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> IDENTIFICADORES                     </b></FONT></td>
		 <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> COMP. IDENTIFICADORES               </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> TASA (IMMEX, PROSEC)                </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PZAS FACTURA                        </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> VALOR MONEDA EXTRANJERA             </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> MONEDA                              </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> FACTOR MONEDA EXTRANJERA            </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> VALOR USD                           </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> TIPO DE CAMBIO                      </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> FACTOR AJUSTE                       </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> VALOR COMERCIAL                     </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> VALOR ADUANA                        </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PRECIO UNITARIO USD                 </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PRECIO UNITARIO MNX                 </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> TASA IGI (SIN IMMEX)                </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> DTA                                 </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> PROGRAMA APLICADO (IMMEX-PROSEC)    </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> SECTOR PROSEC	                   </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> TASA PROSEC                         </b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b> FORMA DE PAGO                       </b></FONT></td>
      </tr>

      <!--
      <tr>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Referencia</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Año</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Patente</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Pedimento</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Fecha Entrada</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Fecha Pago</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Clave</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Tipo</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Regimen</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Tipo Cambio</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Aduana</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Factura</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Fecha Factura</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Moneda Fact</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Factor Mon</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Proveedor</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Pais Proveedor</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Cantidad Comercial</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>UMC</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Cantidad Tarifa</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>UMT</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Valor ME</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Valor USD</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Fletes</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Seguros</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Embalajes</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Otros</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Transporte</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Guia Master</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Guia House</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Material</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Fraccion</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Descripcion</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Identificador</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Complemento</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Tasa IGI</b></FONT></td>
         <%If strTipo=1 then %>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Pais Origen</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Pais Comprador</b></FONT></td>
         <%else%>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Pais Destino</b></FONT></td>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Pais Vendedor</b></FONT></td>
         <%end if%>
         <td bgcolor="#FF9900" ><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><b>Vinculacion</b></FONT></td>
      </tr>
      -->



	 <%

       index=1
	     tempo=index
       if not REFE.eof then
       While (NOT  REFE.EOF)
          identificador = " "
          referencia    = REFE("refcia01")
          ano           = REFE("anio")
          patente       = REFE("patent01")
          pedimento     = REFE("numped01")
          ETA           = REFE("Feta01")
          fecpago       = REFE("fecpag01")
          clave         = REFE("cveped01")
          tipo          = REFE("tipo")
          tcam          = REFE("tipcam01")
          adu           = REFE("cveadu01")
          transportista = REFE("nombar01")
		  pesobr = REFE("pesobr01")
		  totinc = REFE("otros01")

          if strTipo=1 then
             if REFE("fecent01")<>"" then
                fecent = REFE("fecent01")
             else
                fecent=""
             end if
          end if
          fletes    = REFE("fletes01")
		      seguros   = REFE("segros01")
	        embalaje  = REFE("embala01")
		      otros     = REFE("incble01")
		      regimen   = REFE("regime01")
		      Paiss     = REFE("cvepod01")
          if strTipo=2 then
             valfact   = cdbl(REFE("valfac01"))
             tipcambio = cdbl(REFE("tipcam01"))
             valcomer  = Round(tipcambio*valfact)
          end if

		  
		txtCR=""
		txtFI=""
		txtFR=""
		txtIC=""
		txtIM=""
		txtMS=""
		txtPC=""
		txtPD=""
		txtPP=""
		txtPS=""
		txtRC=""
		txtRO=""
		'******************************IDENTIFICADORES
		    ' ix=1
		    Set Conn  = Server.CreateObject ("ADODB.Connection")
        Set RsIDE = Server.CreateObject ("ADODB.RecordSet")
        Conn.Open MM_EXTRANET_STRING
     	  ' strSQL = " SELECT * FROM ssiped11 where refcia11='"&REFE("refcia01")&"' "
	      ' Set RsIDE = Conn.Execute(strSQL)
        ' Do while not RsIDE.Eof
		       ' iden   = RsIDE("deside11")
			     ' comple = RsIDE("comide11")
           ' if ix=1 then
               ' identificador = iden
               ' complemento   = comple
           ' else
               ' identificador = identificador+","+iden
               ' complemento   = complemento+","+comple
           ' end if
           ' ix = ix+1
           ' RsIDE.MoveNext  ' de ssiped11
        ' Loop ' de ssiped11
		strSQL = "select if(cveide11 = 'CR', 'SI','NO') as CR, " &_
				 "if(cveide11 = 'FI', 'SI','NO') as FI, " &_
				 "if(cveide11 = 'FR', 'SI','NO') as FR, " &_
				 "if(cveide11 = 'IC', 'SI','NO') as IC, " &_
				 "if(cveide11 = 'IM', 'SI','NO') as IM, " &_
				 "if(cveide11 = 'MS', 'SI','NO') as MS, " &_
				 "if(cveide11 = 'PC', 'SI','NO') as PC, " &_
				 "if(cveide11 = 'PD', 'SI','NO') as pede, " &_
				 "if(cveide11 = 'PP', 'SI','NO') as PP, " &_
				 "if(cveide11 = 'PS', 'SI','NO') as PS, " &_
				 "if(cveide11 = 'RC', 'SI','NO') as RC, " &_
				 "if(cveide11 = 'RO', 'SI','NO') as RO " &_
				 "from ssiped11 " &_
				 "where refcia11='"&REFE("refcia01")&"' " &_
				 "group by cveide11 "
		Set RsIDE = Conn.Execute(strSQL)
		Do while not RsIDE.Eof
			CR = RsIDE("CR")
			FI = RsIDE("FI")
			FR = RsIDE("FR")
			IC = RsIDE("IC")
			IM = RsIDE("IM")
			MS = RsIDE("MS")
			PC = RsIDE("PC")
			pede = RsIDE("pede")
			PP = RsIDE("PP")
			PS = RsIDE("PS")
			RC = RsIDE("RC")
			RO = RsIDE("RO")
			if CR = "SI" then
				txtCR = "SI"
			end if
			if FI = "SI" then
				txtFI = "SI"
			end if
			if FR = "SI" then
				txtFR = "SI"
			end if
			if IC = "SI" then
				txtIC = "SI"
			end if
			if IM = "SI" then
				txtIM = "SI"
			end if
			if MS = "SI" then
				txtMS = "SI"
			end if
			if PC = "SI" then
				txtPC = "SI"
			end if
			if pede = "SI" then
				txtPD = "SI"
			end if
			if PP = "SI" then
				txtPP = "SI"
			end if
			if PS = "SI" then
				txtPS = "SI"
			end if
			if RC = "SI" then
				txtRC = "SI"
			end if
			if RO = "SI" then
				txtRO = "SI"
			end if
		RsIDE.MoveNext
		Loop

			if txtCR = "" then
				txtCR = "NO"
			end if
			if txtFI = "" then
				txtFI = "NO"
			end if
			if txtFR = "" then
				txtFR = "NO"
			end if
			if txtIC = "" then
				txtIC = "NO"
			end if
			if txtIM = "" then
				txtIM = "NO"
			end if
			if txtMS = "" then
				txtMS = "NO"
			end if
			if txtPC = "" then
				txtPC = "NO"
			end if
			if txtPD = "" then
				txtPD = "NO"
			end if
			if txtPP = "" then
				txtPP = "NO"
			end if
			if txtPS = "" then
				txtPS = "NO"
			end if
			if txtRC = "" then
				txtRC = "NO"
			end if
			if txtRO = "" then
				txtRO = "NO"
			end if		

    '****************************************GUIAS*************************************
          Set Conx = Server.CreateObject ("ADODB.Connection")
          Set RSGUIA = Server.CreateObject ("ADODB.RecordSet")
          Conx.Open MM_EXTRANET_STRING
	        SQLx="SELECT * FROM Ssguia04 WHERE REFCIA04='"&REFE("refcia01")&"' and numgui04<>'' "
	        Set RSGUIA= Conx.Execute(SQLx)
          'response.Write(sqlx)
          'response.end
          Do while not RSGUIA.Eof
             idngui04 = RSGUIA("idngui04")
			       numgui04= RSGUIA("numgui04")
             if idngui04=1 then
	              gmaster= numgui04
		         else
               if idngui04=2 then
                  gmaster2= numgui04
               end if
             end if
 	           RSGUIA.MoveNext
          Loop
          RSGUIA.close
          set RSGUIA = nothing

    '****************************************CONTRIBUCIONES*************************************

                           padv_igi  = ""
                           strdta    = ""
                           piva      = ""
                           sqlImpContr = " SELECT SUM(IF(cveimp36=1,import36,0) )  AS DTA,  " & _
                                         "        SUM(IF(cveimp36=3,import36,0) )  AS IVA,  " & _
                                         "        SUM(IF(cveimp36=6,import36,0) )  AS ADV  " & _
                                         " FROM sscont36                                    " & _
                                         " WHERE REFCIA36 = '"&ltrim(REFE("refcia01"))&"' " & _
                                         " GROUP BY refcia36 "

                           'Response.Write(sqlImpContr)
                           'Response.End

                           set RSImpContr = server.CreateObject("ADODB.Recordset")
                           RSImpContr.ActiveConnection = MM_EXTRANET_STRING
                           RSImpContr.Source= sqlImpContr
                           RSImpContr.CursorType = 0
                           RSImpContr.CursorLocation = 2
                           RSImpContr.LockType = 1
                           RSImpContr.Open()
                           'pnpf=0
                           'pNumPartFact=""
                           if not RSImpContr.eof then
                                    padv_igi         = RSImpContr.fields.item("ADV").value
                                    strdta           = RSImpContr.fields.item("DTA").value
                                    piva             = RSImpContr.fields.item("IVA").value
                            end if
                           RSImpContr.close
                           set RSImpContr = nothing
            '**************************************************************************



    '************************************* Desconsolidadora ****************************************
          strDesCon = ""
          'Set ConDescon = Server.CreateObject ("ADODB.Connection")
          'Set RSDescon = Server.CreateObject ("ADODB.RecordSet")
          'ConDescon.Open MM_EXTRANET_STRING
	        ''SQLx="SELECT * FROM Ssguia04 WHERE REFCIA04='"&REFE("refcia01")&"' and numgui04<>'' "
          'SQLx= " SELECT  refe01,zona01 " & _
          '      " FROM D01CONTE         " & _
          '      " WHERE REFE01 =  '"&REFE("refcia01")&"' " & _
          '      "      and zona01 <> '' "
	        'Set RSDescon= ConDescon.Execute(SQLx)
          ''response.Write(sqlx)
          ''response.end
          'if not RSDescon.Eof then
          '   strDesCon = RSDescon("zona01")
 	        '   RSDescon.MoveNext
          'end if
          'RSDescon.close
          'set RSDescon = nothing

		'******************************d05artic
		    Set Connx = Server.CreateObject ("ADODB.Connection")
        Set RsD05art = Server.CreateObject ("ADODB.RecordSet")
        Connx.Open MM_EXTRANET_STRING
     	  strSQL= " select  item05, frac05, fact05, agru05, desc05, obse05, caco05,umco05, cata05,umta05, vafa05 " & _
                " from d05artic " & _
                " where refe05 = '"&REFE("refcia01")&"' "&_
				" ORDER BY agru05 "
        'strSQL= " select  item05, frac05, fact05, agru05, desc05 from d05artic  where refe05='"&REFE("refcia01")&"' group by agru05,fact05 "
		'response.write(strSQL)
		'response.end()
	      Set RsD05art= Connx.Execute(strSQL)
        Do while not RsD05art.Eof
		    strfrac         = RsD05art("frac05")
		    strfact         = RsD05art("fact05")
	        strAgru         = RsD05art("agru05")
		    strmate         = RsD05art("desc05")
			strobse			= RsD05art("obse05")
            strcodPro       = RsD05art("item05")
            strumcMer       = RsD05art("umco05")
            strCancomMer    = RsD05art("caco05")
            strvalfacMer    = RsD05art("vafa05")
            strvalAduMer    = 0
            strvalPreUniMer = 0

            '************************************************************** fracciones
                 Set RSfracc = Server.CreateObject("ADODB.Recordset")
                 RSfracc.ActiveConnection =  MM_EXTRANET_STRING
                 strSQL= "SELECT fraarn02, " & _
                         "       d_mer102, " & _
                         "       ifnull(tasadv02,0) as tasadv02, " & _
                         "       ifnull(I_iva102,0) as I_iva102, " &_
                         "       u_medc02, " & _
                         "       cantar02, " & _
                         "       u_medt02, " & _
                         "       cancom02, " & _
                         "       paiori02, " & _
                         "       paiscv02, " &_
                         "       preuni02, " &_
                         "       vaduan02, " &_
                         "       vmerme02,  " &_
						 "       ordfra02  " &_
                         " FROM SSFRAC02 WHERE REFCIA02='"&REFE("refcia01")&"' and ordfra02='"&RsD05art("agru05")&"' and  " &_
                         " fraarn02='"&RsD05art("frac05")&"' "


                 'strSQL= "SELECT fraarn02,d_mer102,ifnull(tasadv02,0) as tasadv02, ifnull(I_iva102,0) as I_iva102, " &_
                 '        " u_medc02,cantar02,u_medt02,cancom02,paiori02,paiscv02 " &_
                 '        " FROM SSFRAC02 WHERE REFCIA02='"&REFE("refcia01")&"' and ordfra02='"&RsD05art("agru05")&"' and  " &_
                 '        " fraarn02='"&RsD05art("frac05")&"'"

                 RSfracc.Source = strSQL
                 RSfracc.CursorType = 0
                 RSfracc.CursorLocation = 2
                 RSfracc.LockType = 1
                 RSfracc.Open()
                 if not RSfracc.eof then
                    nfraccion    = RSfracc.Fields.Item("fraarn02").Value
                    mercancia    = RSfracc.Fields.Item("d_mer102").Value
                    igi          = RSfracc.Fields.Item("tasadv02").Value
                    ivaimpuesto  = RSfracc.Fields.Item("I_iva102").Value
                    umc          = RSfracc.Fields.Item("u_medc02").Value
                    cantarifa    = RSfracc.Fields.Item("cantar02").Value
                    umt          = RSfracc.Fields.Item("u_medt02").Value
                    cancom       = RSfracc.Fields.Item("cancom02").Value
                    paisOD       = RSfracc.Fields.Item("paiori02").Value
                    paisCV       = RSfracc.Fields.Item("paiscv02").Value
                    strvalfacfra = RSfracc.Fields.Item("vmerme02").Value
                    strvalAduFra = RSfracc.Fields.Item("vaduan02").Value
					ordfra = RSfracc.Fields.Item("ordfra02").Value
                 end if
                 RSfracc.close
                 set RSfracc = nothing
            '***********************************facturas


                    'strvalfacfra = RSfracc.Fields.Item("vmerme02").Value
                    'strvalAduFra = RSfracc.Fields.Item("vaduan02").Value
                    'strvalfacMer
                    if strvalfacfra > 0 then
						if strvalAduFra <> 0 and strvalfacMer <> 0 and strvalfacfra <> 0 then
							strvalAduMer    = ( strvalAduFra * strvalfacMer) / strvalfacfra
						Else
							strvalAduMer = 0
						End If
						if strvalAduMer <> 0 and strvalfacMer <> 0 then
							strvalPreUniMer = strvalAduMer / strvalfacMer
						Else
							strvalPreUniMer = 0
						End If
                    end if



                                 permisProImm = ""
                                 sectorprosec = ""
                                 sqlpermisos= " select cveide12 as TipoTasa, " &_
                                              " concat_ws(',',cveide12,comide12) as cvepermis,  "&_
                                              " comide12,  "&_
											  " group_concat(distinct ifnull(cveide12,'') separator ' | ') as 'Identificadores' " &_
                                              " from ssipar12 " &_
                                              " where refcia12='"&REFE("refcia01")&"' and ordfra12= '"&ordfra&"' "&_
                                              " group by cveide12"
                                 'Response.Write(sqlpermisos)
                                 'Response.End

                                 set RSPermiso = server.CreateObject("ADODB.Recordset")
                                 RSPermiso.ActiveConnection = MM_EXTRANET_STRING
                                 RSPermiso.Source= sqlpermisos
                                 RSPermiso.CursorType = 0
                                 RSPermiso.CursorLocation = 2
                                 RSPermiso.LockType = 1
                                 RSPermiso.Open()
                                 pp=0
                                 if not RSPermiso.eof then
									identif = RSPermiso.fields.item("Identificadores").value
									comide = RSPermiso.fields.item("comide12").value
                                     while not RSPermiso.eof
                                       if RSPermiso.fields.item("TipoTasa").value = "IM" or RSPermiso.fields.item("TipoTasa").value = "PS" then

                                          if pp=0 then
                                              permisProImm = RSPermiso.fields.item("TipoTasa").value
                                              pp=1
                                          else
                                              permisProImm = permisProImm &" , "& RSPermiso.fields.item("TipoTasa").value
                                          end if

                                          if RSPermiso.fields.item("TipoTasa").value = "PS"  then
                                             sectorprosec = RSPermiso.fields.item("comide12").value
                                          end if
                                       end if
                                      RSPermiso.movenext
                                     wend
                                 end if
                                 RSPermiso.close
                                 set RSPermiso= nothing


                                 'response.write(permisProImm)
                                 'Response.End

                             'StrTipoTasaProImm = ""
                             ''if isNull(permis) or permis="" or isEmpty(permis) then
                             'if Len(Trim(permis)) = 0 then
                             '  StrTipoTasaProImm = ""
                             'else
                             '    if InStr(permis,"IM") > 0 then
                             '      StrTipoTasaProImm = "IM"
                             '    else
                             '        if InStr(permis,"PS") > 0 then
                             '          StrTipoTasaProImm = "PS"
                             '        else
                             '          StrTipoTasaProImm = ""
                             '        end if
                             '    end if
                             'end if



            '***********************************Pais
                 ' Nombre Pais Origen Destino
                 strNompaisOD = ""
                 if paisOD <> "" then
                     Set RsPaisOD = Server.CreateObject("ADODB.Recordset")
                     RsPaisOD.ActiveConnection = MM_EXTRANET_STRING
                     strSqlPaisOD =  " SELECT nompai19 " & _
                                     " FROM sspais19   " & _
                                     " WHERE cvepai19 = '"&ltrim(paisOD)&"'"

                     'Response.Write(strSqlPaisOD)
                     'Response.End
                     RsPaisOD.Source = strSqlPaisOD
                     RsPaisOD.CursorType = 0
                     RsPaisOD.CursorLocation = 2
                     RsPaisOD.LockType = 1
                     RsPaisOD.Open()
                     if not RsPaisOD.eof then
                            strNompaisOD    = RsPaisOD.Fields.Item("nompai19").Value
                     end if
                     RsPaisOD.close
                     set RsPaisOD = Nothing
                 end if


                 ' Nombre Pais Comprador-Vendedor
                 strNompaisCV = ""
                 'strNompaisOD = ""
                 if paisCV  <> "" then
                     Set RsPaisCV = Server.CreateObject("ADODB.Recordset")
                     RsPaisCV.ActiveConnection = MM_EXTRANET_STRING
                     strSqlPaisCV =  " SELECT nompai19 " & _
                                     " FROM sspais19   " & _
                                     " WHERE cvepai19 = '"&ltrim(paisCV)&"'"

                     'Response.Write(strSqlPaisCV)
                     'Response.End
                     RsPaisCV.Source = strSqlPaisCV
                     RsPaisCV.CursorType = 0
                     RsPaisCV.CursorLocation = 2
                     RsPaisCV.LockType = 1
                     RsPaisCV.Open()
                     if not RsPaisCV.eof then
                            strNompaisCV    = RsPaisCV.Fields.Item("nompai19").Value
                     end if
                     RsPaisCV.close
                     set RsPaisCV = Nothing
                 end if
                 '**************************************************************************************************************

            '***********************************


                  Set Rfac = Server.CreateObject("ADODB.Recordset")
                  Rfac.ActiveConnection =  MM_EXTRANET_STRING
                  sqL= " select refcia39," & _
                       "        numfac39," & _
                       "        idfisc39," & _
                       "        nompro39," & _
                       "        dompro39," & _
                       "        terfac39," & _
                       "        monfac39," & _
                       "        IF( fecfac39 >'1900-00-00' AND fecfac39 is not null,DATE_FORMAT(fecfac39,'%d/%m/%Y'),'00/00/0000') as fecfac, " & _
                       "        valdls39," & _
                       "        valmex39," & _
                       "        cvepro39," &_
                       "        vincul39," & _
                       "        facmon39 " & _
                       " from ssfact39   " & _
                       " where refcia39='"&REFE("refcia01")&"' "  & _
                       "       and numfac39='"&RsD05art("fact05")&"' "


                  Rfac.Source = sqL
                  Rfac.CursorType = 0
                  Rfac.CursorLocation = 2
                  Rfac.LockType = 1
                  Rfac.Open()
                  if not Rfac.eof then
                      nompro = Rfac.Fields.Item("nompro39").Value
                      nfactura= Rfac.Fields.Item("numfac39").Value
                      nomprov= Rfac.Fields.Item("nompro39").Value
					  incoterm= Rfac.Fields.Item("terfac39").Value
                      monfact= Rfac.Fields.Item("monfac39").Value
                      valor_usd=Rfac.Fields.Item("valdls39").value
                      valor_me=Rfac.Fields.Item("valmex39").value
                      fechafac=Rfac.Fields.Item("fecfac").Value
                      cvepro=Rfac.Fields.Item("cvepro39").value
                      vinculacion=Rfac.Fields.Item("vincul39").value
                      factomon=Rfac.Fields.Item("facmon39").value
                      if vinculacion=1 then
                         vinc="si"
                      else
                         vinc="no"
                      end if
                  end if
                  Rfac.close
                  set Rfac = nothing
            '************************************************************** ssprov22

                  Set RSprov = Server.CreateObject("ADODB.Recordset")
                  RSprov.ActiveConnection =  MM_EXTRANET_STRING
                  strSQL="select paipro22 from ssprov22 where cvepro22='"&cvepro&"' "
                  RSprov.Source = strSQL
                  RSprov.CursorType = 0
                  RSprov.CursorLocation = 2
                  RSprov.LockType = 1
                  RSprov.Open()
                  if not RSprov.eof then
                     paiprov = RSprov.Fields.Item("paipro22").Value
                  end if
                  RSprov.close
                  set RSprov = nothing
            '*****************************DATOS FILAS*****************************

                  x=0
                 %>
                <!--
                    PEDIMENTO
                    FECHA PAGO MM/DD/AA
                    FECHA ENTRADA MM/DD/AA
                    ADUANA
                    CLAVE DE PEDIMENTO
                    País de Origen
                    Clave Pais de Origen
                    País de Venta
                    Clave País de Venta
                    Referencia
                    DESCONSOLIDADORA
                    GUIA MASTER
                    GUIA  HOUSE
                    FECHA FACTURA   DD/MM/AA
                    NUM. FACTURA
                    PROVEEDOR
                    NUM MATERIAL CONTINENTAL
                    DESCRIPCION DE MATERIAL
                    UNIDAD DE MEDIDA
                    FRACCION ARANCELARIA
                    TASA (IMMEX, PROSEC)
                    PZAS FACTURA
                    VALOR MONEDA EXTRANJERA
                    MONEDA
                    FACTOR MONEDA EXTRANJERA
                    VALOR USD	ç
                    TIPO DE CAMBIO
                    FACTOR AJUSTE
                    VALOR COMERCIAL
                    VALOR ADUANA
                    PRECIO UNITARIO USD
                    PRECIO UNITARIO MNX
                    TASA IGI (SIN IMMEX)
                    DTA
                    PROGRAMA APLICADO (IMMEX-PROSEC)
                    SECTOR PROSEC
                    TASA PROSEC
                    FORMA DE PAGO
                -->
                 <tr>
                     <td align="center"><font size="-1"><%RESPONSE.Write(pedimento)     %>   </font></td> <!-- Pedimento -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(patente)       %>   </font></td> <!-- Patente -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(formatofechaNum(fecpago))%>   </font></td> <!-- Fecha Pago -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(formatofechaNum(fecent)) %>   </font></td> <!-- Fecha Entrada -->
                     <TD align="center"><font size="-1"><%RESPONSE.Write(adu)           %>   </font></TD> <!-- Aduana -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(clave)         %>   </font></td> <!-- Clave -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strNompaisOD)  %>   </font></td> <!-- Pais Origen/Destino -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(paisOD)        %>   </font></td> <!-- Clave Pais Origen/Destino -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strNompaisCV)  %>   </font></td> <!-- Pais Venta Vendedor/Comprador -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(paisCV)        %>   </font></td> <!-- Clave Pais  Venta Vendedor/Comprador -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(referencia)    %>   </font></td> <!-- Referencia -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(pesobr)    %>   </font></td> <!-- Peso Bruto -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(totinc)    %>   </font></td> <!-- Total de Incrementables -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strDesCon)     %>   </font></td> <!-- Desconsolidadora -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(gmaster)       %>   </font></td> <!-- Guia Master -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(gmaster2)      %>   </font></td> <!-- Guia House -->
					 
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtCR)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtFI)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtFR)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtIC)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtIM)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtMS)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtPC)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtPD)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtPP)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtPS)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtRC)      %>   </font></td> <!-- IDENTIFICADOR -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(txtRO)      %>   </font></td> <!-- IDENTIFICADOR -->
					 
                     <td align="center"><font size="-1"><%RESPONSE.Write(fechafac)      %>   </font></td> <!-- Fecha Factura -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(nfactura)      %>   </font></td> <!-- Factura -->
					 <td align="center"><font size="-1"><%RESPONSE.Write(incoterm)      %>   </font></td> <!-- Incoterm -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(nomprov)       %>   </font></td> <!-- Proveedor -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strcodPro)     %>   </font></td> <!-- Material -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strmate)       %>   </font></td> <!-- Descripcion del material-->
					 <td align="center"><font size="-1"><%RESPONSE.Write(strobse)       %>   </font></td> <!-- Descripcion de la Observacion (INGLES)-->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strumcMer)     %>   </font></td> <!-- UMC - Unidad de medida-->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strfrac)       %>   </font></td> <!-- Fraccion Arancelaria-->
					 <td align="center"><font size="-1"><%RESPONSE.Write(ordfra)       %>   </font></td> <!-- No. Partida-->
					 <td align="center"><font size="-1"><%RESPONSE.Write(identif)       %>   </font></td> <!-- Identificadores Partida-->
					 <td align="center"><font size="-1"><%RESPONSE.Write(comide)       %>   </font></td> <!-- Complemento Identificadores Partida-->
                     <td align="center"><font size="-1">                                     </font></td> <!-- Tasa (IMMEX,PROSEC) -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strCancomMer)  %>   </font></td> <!-- Piezas factura --><!-- Cantidad Comercial -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strvalfacMer)  %>   </font></td> <!-- Valor moneda extranjera--><!--Valor USD--> <!-- valor_usd -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(monfact)       %>   </font></td> <!-- Moneda  --> <!-- Moneda Fact -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(factomon)      %>   </font></td> <!-- Factor moneda extranjera  --> <!-- Factor Mon -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strvalfacMer*factomon)     %>   </font></td> <!-- Valor USD --> <!-- valor_usd -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(tcam)          %>   </font></td> <!-- Tipo Cambio -->
                     <td align="center"><font size="-1">                                     </font></td> <!-- Factor ajuste -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strvalfacMer*factomon*tcam)      %> </font></td> <!-- Valor Comercial--><!-- Valor ME --> <!-- valor_me -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strvalAduMer)  %>                 </font></td> <!-- Valor aduana -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strvalPreUniMer*factomon) %>      </font></td> <!-- Precio unitario USD -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strvalPreUniMer*factomon*tcam) %> </font></td> <!-- Precio Unitario MNX -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(igi)           %>   </font></td> <!-- TASA IGI (SIN IMMEX) -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(strdta)%>           </font></td> <!-- DTA -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(permisProImm)  %>   </font></td> <!-- PROGRAMA APLICADO (IMMEX-PROSEC) -->
                     <td align="center"><font size="-1"><%RESPONSE.Write(sectorprosec)  %>                                     </font></td> <!-- SECTOR PROSEC -->
                     <td align="center"><font size="-1">                                     </font></td> <!-- TASA PROSEC -->
                     <td align="center"><font size="-1">                                     </font></td> <!-- FORMA DE PAGO -->
                 </tr>
         <%

         '********d05artic

	        RsD05art.MoveNext  ' de d05artic
	        tempo=index+1
       Loop ' de d05artic
     %>

<%

'**********************************************************************

         Refe.MoveNext 'avanza referencia  ---->
		     index=index+1
   wend 'REFErencia
 else
%>
<tr>
  <th colspan=12>
    <font size="1" face="Arial">No se Encontro ningun registro con esos parametros
  </th>
</tr>
<table>
<%end if
'end if'del movimiento %>
</form>







<%
    Function pd(n, totalDigits)
        if totalDigits > len(n) then
            pd = String(totalDigits-len(n),"0") & n
        else
            pd = n
        end if
    End Function

    Function formatofechaNum(DFecha)
       if isdate( DFecha ) then
          formatofechaNum = Pd(Month( DFecha ),2)& "/"&Pd(DAY( DFecha ),2)&"/"&YEAR(DFecha)
       else
          formatofechaNum	= DFecha
       end if
    End Function
%>