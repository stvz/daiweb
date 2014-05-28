<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% 
	Server.ScriptTimeout=150000
	Dim strOficina,aduSec 
%>
<HTML>
	<HEAD>
		<TITLE>:: .... REPORTE DE FRACCIONES EN EL Unilever vs DOF .... ::</TITLE>
	</HEAD>
	<BODY>
<% 
	if  Session("GAduana") = "" then %>
		<table border="0" align="center" cellpadding="0" cellspacing="7" class="titulosconsultas">
			<tr>
				<td><%=strMenjError%></td>
			</tr>
		</table>
<% 
	else 
	    mov=request.form("mov") ' Tipo de movimiento IMPO o EXPO "i"
        fi=	trim(request.form("fi")) ' Fecha inicio reporte "10/10/2010"
        ff=	trim(request.form("ff")) ' Fecha final reporte "14/10/2010"
		fid= trim(request.form("fid")) ' Fecha inicio del DOF "07/10/2010"	
		ffd= trim(request.form("ffd")) ' Fecha final del DOF "14/10/2010"
		Vrfc= Request.Form("rfcCliente") ' RFC cliente "UME651115N48"
        Vckcve=	Request.Form("ckcve") ' la seleccion si usares RFC o CVECLIE si es 0 es por RFC y si es 1 es por CVE CLI  
    	txtcli= Request.Form("txtCliente") ' clave de cliente "Todos"	
		multiofi =	Request.Form("multi") ' Multioficina "t"
						  
        if isdate(fi) and isdate(ff) then
			DiaI = cstr(datepart("d",fi))
            Mesi = cstr(datepart("m",fi))
            AnioI = cstr(datepart("yyyy",fi))
            DateI = Anioi & "/" & Mesi & "/" & Diai            
			DiaF = cstr(datepart("d",ff))
            MesF = cstr(datepart("m",ff))
            AnioF = cstr(datepart("yyyy",ff))
            DateF = AnioF & "/" & MesF & "/" & DiaF
			
			DiaId = cstr(datepart("d",fid))
            MesId = cstr(datepart("m",fid))
            AnioId = cstr(datepart("yyyy",fid))
			DateIDof = AnioId & "/" & MesId & "/" & DiaId
			
			DiaFd = cstr(datepart("d",ffd))
            MesFd = cstr(datepart("m",ffd))
            AnioFd = cstr(datepart("yyyy",ffd))
            DateFDof = AnioFd & "/" & MesFd & "/" & DiaFd
            		
			
			 if request.form("tipRep") = "2" then
				 Response.Addheader "Content-Disposition", "attachment; filename=ReporteDOF"
				 Response.ContentType = "application/vnd.ms-excel"
				 
			 end if
			if multiofi = "t" and Vckcve = "1" Then
			Response.Write("<table border='0' align='center' cellpadding='0' cellspacing='7' class='titulosconsultas'>" &_
								"<tr>" &_
									"<td>No es posible elegir por clave de cliente y MultiOficina elijalo por RFC</td>" &_
								"</tr>" &_
							"</table>")
			Else
				Dim htlm
										
					set miCon=Server.CreateObject("ADODB.Connection")
					Set miRS = Server.CreateObject("ADODB.Recordset")
					ConnectionString="DRIVER={SQL Server};SERVER=10.66.1.19;UID=sa;PWD=S0l1umF0rW;DATABASE=SIR"
	
					if multiofi<>"t" then 
						jnxadu=Session("GAduana")
						select case jnxadu
							case "VER"
								strOficina="rku"
								aduSec="'430'"
							case "MEX"
								strOficina="dai"
								aduSec="'470'"
							case "MAN"
								strOficina="sap"
								aduSec="'160'"
							case "TAM"
								strOficina="ceg"
								aduSec="'810'"
							case "LAR"
								strOficina="LAR"
								aduSec="'800','240'"
							case "LZR"
								strOficina="lzr"
								aduSec="'510'"
							case "TOL"
								strOficina="tol"
								aduSec="'650'"
						end select
						if strOficina="LAR" then 
							set miCon=Server.CreateObject("ADODB.Connection")
							Set miRS = Server.CreateObject("ADODB.Recordset")
							ConnectionString="DRIVER={SQL Server};SERVER=10.66.1.19;UID=sa;PWD=S0l1umF0rW;DATABASE=SIR"
							query = GeneraSQL("SIR",DateI,Datef,DateIDof,DateFDof,Vrfc,txtcli,mov,aduSec)
							miRS.Open query, ConnectionString
							IF miRS.BOF = True And miRS.EOF = True Then
							else
								htlm = generahtml(miRS)
							end if							
						elseif strOficina="lzr" or  strOficina="ceg" then 
							set miCon=Server.CreateObject("ADODB.Connection")
							Set miRS = Server.CreateObject("ADODB.Recordset")
							ConnectionString="DRIVER={SQL Server};SERVER=10.66.1.19;UID=sa;PWD=S0l1umF0rW;DATABASE=SIR"
							query = GeneraSQL("SIR",DateI,Datef,DateIDof,DateFDof,Vrfc,txtcli,mov,aduSec)
							miRS.Open query, ConnectionString
							IF miRS.BOF = True And miRS.EOF = True Then
							else
								htlm = generahtml(miRS)
							end if
							Set ConnStr = Server.CreateObject ("ADODB.Connection")
							ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
							Set RSops = CreateObject("ADODB.RecordSet")
							query = GeneraSQL("SAAI",DateI,Datef,DateIDof,DateFDof,Vrfc,txtcli,mov,strOficina)
							Set RSops = ConnStr.Execute(query)
							IF RSops.BOF = True And RSops.EOF = True Then
							else
								htlm=htlm &generahtml(RSops)
							end if
						else 
							Set ConnStr = Server.CreateObject ("ADODB.Connection")
							ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
							Set RSops = CreateObject("ADODB.RecordSet")
							query = GeneraSQL("SAAI",DateI,Datef,DateIDof,DateFDof,Vrfc,txtcli,mov,strOficina)
							
							Set RSops = ConnStr.Execute(query)
							IF RSops.BOF = True And RSops.EOF = True Then
							else
								htlm= generahtml(RSops)
							end if
						end if 	
					elseif multiofi="t" then 
						
						For indi = 1 To 7
						
							Select Case indi
								Case 1
									strOficina = "rku"
									aduSec="'430'"
								Case 2
									strOficina = "dai"
									aduSec="'470'"
								Case 3
									strOficina = "sap"
									aduSec="'160'"
								Case 4
									strOficina = "lzr"
									aduSec="'510'"
								Case 5
									strOficina = "ceg"
									aduSec="'810'"
								Case 6
									strOficina = "tol"
									aduSec="'650'"
								case 7 
									strOficina="LAR"
									aduSec="'800','240'"
							End Select
							if strOficina="LAR" then 
								set miCon=Server.CreateObject("ADODB.Connection")
								Set miRS = Server.CreateObject("ADODB.Recordset")
								ConnectionString="DRIVER={SQL Server};SERVER=10.66.1.19;UID=sa;PWD=S0l1umF0rW;DATABASE=SIR"
								query = GeneraSQL("SIR",DateI,Datef,DateIDof,DateFDof,Vrfc,txtcli,mov,aduSec)
								miRS.Open query, ConnectionString
								IF miRS.BOF = True And miRS.EOF = True Then
								else
									htlm =htlm & generahtml(miRS)
								end if							
							elseif strOficina="lzr" or  strOficina="ceg" then 
								set miCon=Server.CreateObject("ADODB.Connection")
								Set miRS = Server.CreateObject("ADODB.Recordset")
								ConnectionString="DRIVER={SQL Server};SERVER=10.66.1.19;UID=sa;PWD=S0l1umF0rW;DATABASE=SIR"
								query = GeneraSQL("SIR",DateI,Datef,DateIDof,DateFDof,Vrfc,txtcli,mov,aduSec)
								miRS.Open query, ConnectionString
								IF miRS.BOF = True And miRS.EOF = True Then
								else
									htlm =htlm &  generahtml(miRS)
								end if
								
								Set ConnStr = Server.CreateObject ("ADODB.Connection")
								ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
								Set RSops = CreateObject("ADODB.RecordSet")
								query = GeneraSQL("SAAI",DateI,Datef,DateIDof,DateFDof,Vrfc,txtcli,mov,strOficina)
								Set RSops = ConnStr.Execute(query)
								IF RSops.BOF = True And RSops.EOF = True Then
								else
									htlm=htlm &generahtml(RSops)
								end if
								
							else 
						
							Set ConnStr = Server.CreateObject ("ADODB.Connection")
							ConnStr.Open "DRIVER={MySQL ODBC 5.1 Driver}; SERVER=localhost; UID=EXTRANET; PWD=rku_admin; OPTION=16427"
							Set RSops = CreateObject("ADODB.RecordSet")
							query = GeneraSQL("SAAI",DateI,Datef,DateIDof,DateFDof,Vrfc,txtcli,mov,strOficina)
							
							Set RSops = ConnStr.Execute(query)
							IF RSops.BOF = True And RSops.EOF = True Then
							else
								htlm =htlm &  generahtml(RSops)
							end if
						end if 	
						Next
					end if 
						
				if htlm="" Then
					Response.Write("<table border='1' align='center' cellpadding='0' cellspacing='7' class='titulosconsultas'>" &_
										"<tr>" &_
											"<td>No existen datos que mostrar</td>" &_
										"</tr>" &_
									"</table>")
				Else
					
					encabezado = ""
					encabezado = "<table align='center' Width='1000' bordercolor='#C1C1C1' border='2' align='center' cellpadding='0' cellspacing='0'>"
					encabezado = encabezado	&		"<tr>" &_
														"<td colspan='7'>Fracciones Usadas Mencionadas en el Diario de la Federacion (DOF)</td>" &_
													"</tr>" &_
													"<tr>" &_
														"<td colspan='7'>Del " & fi & " al " & ff & "</td>" &_
													"</tr>"
					encabezado = encabezado & 		"<tr bgcolor='#006699' class='boton'>" &_
														CeldaHead("Aduana / Sección") &_
														CeldaHead("Pais Origen") &_
														CeldaHead("Codigo Producto") &_
														CeldaHead("Desc. Producto") &_
														CeldaHead("Fracción") &_
														CeldaHead("Fecha Publicacion")&_
														CeldaHead("Fecha Entrada Vigor")&_
														CeldaHead("Acuerdo")&_
													"</tr>"
				htlm = encabezado & htlm & "</table>"
				Response.Write(htlm)
				End If
			End if
		End If
	End If



Function GeneraSQL(Sistema,FIoperacion,FFoperacion,FIdof,FFdof,RFC_,CVE_,mov2,aduSec2) 
	SQL = ""
		if Sistema="SAAI"  then ' Si el reporte solicitado es del sistema SAAI
			
			movim = mov2
			if mov2 = "a" Then
				movim = "i"
				SQL = SQL & OfiSQL(movim, aduSec2,FIoperacion,FFoperacion,FIdof,FFdof,RFC_,CVE_) & " UNION ALL "
				movim = "e"
				SQL = SQL & OfiSQL(movim, aduSec2,FIoperacion,FFoperacion,FIdof,FFdof,RFC_,CVE_) & "  "
			Else
				SQL = SQL & OfiSQL(movim, aduSec2,FIoperacion,FFoperacion,FIdof,FFdof,RFC_,CVE_) & " "
			End If
			
		elseif Sistema="SIR" then 'Obtener informacion de SIR
			dim mov3 
			if mov2="i" then 
				mov3="'impo'"
			elseif mov2="e" then 
				mov3="'expo'"
			
			end if 
			SQL="select distinct dbo.GSI_F_RetornaSucursal(Partes.R_ID_SucPatAdu71,'ADUANASEC') aduana,Partes.PaisOrigen PaisO,Partes.PF_FraccionA fraccion, cast(Partes.sDescripcionAA as varchar(550)) descripcion,Partes.sParte codigoprod ,dof.fechPub fechapub , " & _
					"dof.FechVigor fechvigor, dof.acuerdo " & _
				"from (select distinct m.r_sRFC, " & _
						"m.R_ID_SucPatAdu71, " & _
						"F_PaisFac AS PaisOrigen, " & _
						"m.PF_FraccionA , " & _
						"p.sDescripcionAA,p.nIdParte99,p.sParte " & _
						"from sir.SIR.SIR_99_PARTES as p " & _
						"inner join sir.dbo.GSI_VT_MER_SVG_OperacionesMercancias as m on p.nIdParte99=m.PF_IDparte99 and p.sFraccion=m.PF_FraccionA " & _
						"inner join sir.dbo.VT_Pedimentos as pe on pe.nIdPedimento149=m.R_IDPedimento149 " & _
						"where convert(date,pe.dFechaPago) between convert(date,'"& FIoperacion &"') and CONVERT(date,'"& FFoperacion &"') " 
					if Vckcve="0"  then
						SQL=SQL &	" and  m.r_sRFC='"&RFC_&"'  " 
					elseif CVE_<>"Todos" and Vckcve="1" then 
						SQL=SQL &" and  m.R_sClaveCliente='"& CVE_ &"'  " 
					end if
					if mov2<>"a" then 
						SQL=SQL& " and pe.[Tipo Operacion]="&mov3
					end if 
					if Ofi_<>"t" then 
						SQL=SQL& " and pe.adusec in("&aduSec2&") " 
					end if
					SQL=SQL & "group by m.r_sRFC,p.nIdParte99,p.sParte,m.R_ID_SucPatAdu71,m.F_PaisFac,m.PF_FraccionA,p.sDescripcionAA ) as Partes " & _
				"left join  sir.dbo.dof_unilever aS dOF on replace(dOF.fraccion,'.','')=Partes.PF_FraccionA " & _
				"where dof.fechPub between convert(date,'"&FIdof&"')and CONVERT(date,'" & FFdof & "') " & _
				"order by fechPub,codigoprod " 
		
	
	End if 
	
	GeneraSQL=SQL
End Function

Function OfiSQL(movi, ofi,Fio,Ffo,Fip,Ffp,rfc,cve) 'OfiSQL(mov2, aduSec2,FIoperacion,FFoperacion,FIdof,FFdof,RFC_,CVE_)
	SQL2 = ""
	if movi = "i" then
		movto = "'IMPO' as Mov, "
		fecentpre = "fecent01"
	Else
		movto = "'EXPO' as Mov, "
		fecentpre = "fecpre01"
	End If
	SQL2 = 	"SELECT " & movto &_
			"i.nomcli01 AS 'nomcliente', " &_
			"i.adusec01 AS 'aduana', " &_
			"i.cvepod01 AS 'PaisO', " &_
			"fr.fraarn02 AS 'fraccion', " &_
			"fr.d_mer102 AS 'descripcion', " &_
			"d.cpro05 AS 'codigoprod', " &_
			"dof.fraccion as 'fracciondof', " &_
			"dof.descripcion as 'descdof', " &_
			"dof.FechPub as 'fechapub', " &_
			"dof.FechVigor as 'fechvigor', " &_
			"dof.Acuerdo as 'acuerdo' " &_
			"FROM " & Ofi & "_extranet.ssdag" & movi & "01 AS i " &_
			"LEFT JOIN " & Ofi & "_extranet.ssfrac02 AS fr ON i.refcia01 = fr.refcia02 " &_
			"LEFT JOIN " & Ofi & "_extranet.d05artic AS d ON d.refe05 = i.refcia01 AND d.agru05 = fr.ordfra02 AND d.frac05 = fr.fraarn02 " &_
			"LEFT JOIN sistemas.dof_unilever as dof ON fr.fraarn02 = REPLACE(dof.fraccion,'.','') " 
			
			if Vckcve="0"  then
				SQL2=SQL2 &	"WHERE i.rfccli01 in('" & rfc & "') AND "	
			elseif txtcli<>"Todos" and Vckcve="1" then 
				SQL2=SQL2 &"WHERE i.cvecli01 in('" & cve & "') AND "	
			elseif txtcli="Todos" then 
				SQL2=SQL2 &"WHERE "
			end if 
			
			SQL2=SQL2 &" i.firmae01 <> '' AND i.firmae01 IS NOT NULL AND i.fecpag01 >= '" & Fio & "' AND i.fecpag01 <='" & Ffo & "' "&_
			"AND dof.fechpub >= '" & Fip & "' AND dof.fechpub <='" & Ffp & "' " &_
			ExcluirOp_SIR(Fio,Ffo,rfc,cve,movi,Ofi) &_
			" GROUP BY fracciondof , codigoprod, acuerdo " &_
			"HAVING fracciondof IS NOT NULL "
		
	OfiSQL = SQL2
End Function

Function ExcluirOp_SIR(DI,DF,Rfc,cvec,T_op,ofici)
	dim refes,i
	refes=""
	i=0
	select case ofici
		case "rku" 
			strOficina="rku"
			aduSec="'430'"
		case "dai" 
			strOficina="dai"
			aduSec="'470'"
		case "sap" 
			strOficina="sap"
			aduSec="'160'"
		case "ceg" 
			strOficina="ceg"
			aduSec="'810'"
		case "lzr"
			strOficina="lzr"
			aduSec="'510'"
		case "TOL" 
			strOficina="tol"
			aduSec="'650'"
	end select
	Set miConec = Server.CreateObject("ADODB.Recordset")
		ConnectionString="DRIVER={SQL Server};SERVER=10.66.1.19;UID=sa;PWD=S0l1umF0rW;DATABASE=SIR"
		 
		dim mov4 
			if T_op="i" then 
				mov4="'impo'"
				else
				mov4="'expo'"
			end if 
					strSQL="select distinct pe.sReferencia  Referencia " & _
						"from sir.SIR.SIR_99_PARTES as p " & _
						"inner join sir.dbo.GSI_VT_MER_SVG_OperacionesMercancias as m on p.nIdParte99=m.PF_IDparte99 and p.sFraccion=m.PF_FraccionA " & _
						"inner join sir.dbo.VT_Pedimentos as pe on pe.nIdPedimento149=m.R_IDPedimento149 " & _
						"where convert(date,pe.dFechaPago) between convert(date,'"& DI &"') and CONVERT(date,'"& DF &"') " 
						if Vckcve="0"  then
							strSQL=strSQL &	" and  m.r_sRFC='"&Rfc&"'  " 
						elseif cvec<>"Todos" and Vckcve="1" then 
							strSQL=strSQL &" and  m.R_sClaveCliente='" & cvec & "'  " 
						end if
						if T_op<>"a" then 
							strSQL=strSQL& " and pe.[Tipo Operacion]="&mov4
						end if 
						if Ofi_<>"t" then 
							strSQL=strSQL& " and pe.adusec in("&aduSec&" )" 
						end if

		miConec.Open strSQL, ConnectionString
	If err.number =0 then  
			IF miConec.BOF = True And miConec.EOF = True Then
				refes=""
			else
				While Not  miConec.eof
					
					if i>0 then 
						refes=refes &","
					end if
					refes=refes &"'"& miConec("Referencia") &"'"
					i=i+1
					miConec.movenext
				Wend
				if refes<>"" then 
					refes= " and i.refcia01 not in("& refes &") "
				end if
			End if
	else 
		response.write err.description
	end if 
	miConec.close		
	
	ExcluirOp_SIR=refes
End Function

Function generahtml(RecSet)
	codigo = ""
	Do Until RecSet.EOF
	
			codigo = codigo & 	"<tr>" &_
								CeldaCuerpo(RecSet("aduana")) &_
								CeldaCuerpo(RecSet("PaisO")) &_
								CeldaCuerpoN(Recset("codigoprod")) &_
								CeldaCuerpo(Recset("descripcion")) &_
								CeldaCuerpo(Recset("fraccion")) &_
								CeldaCuerpo(Recset("fechapub")) &_
								CeldaCuerpo(Recset("fechvigor")) &_
								CeldaCuerpo(Recset("acuerdo")) &_
							"</tr>"
	RecSet.MoveNext
	
	Loop
	
	generahtml = codigo
End Function

Function CeldaCuerpo(txtcelda)
	tags = ""
	tags = "<td align='center'><font size='1' face='Arial'>" & txtcelda & "</font></td>"
	CeldaCuerpo = tags
End Function
Function CeldaCuerpoN(txtcelda)
	If IsNull(txtcelda) = True Or txtcelda = "" or txtcelda=" " Then
		txtcelda = "&nbsp;"
	End If
	tags = ""
	tags = "<td align='center' style=""mso-number-format:\@""><font size='1' face='Arial'>" & txtcelda & "</font></td>"
	CeldaCuerpoN = tags
End Function
Function CeldaHead(txtcelda)
	tags = ""
	tags = "<td align='center' nowrap><strong><font color='#FFFFFF' size='2' face='Arial, Helvetica, sans-serif'>" & txtcelda & "</font></strong></td>"
	CeldaHead = tags
End Function


%>
	</BODY>
</HTML>