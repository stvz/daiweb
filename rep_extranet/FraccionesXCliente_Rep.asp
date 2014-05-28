<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp" -->
<!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->
<% Server.ScriptTimeout=1500 %>

<%if  Session("GAduana") <> "" then %>
      <%
      DIM vOPC
      'vOPC=Request.Form("OPC")
      if request.form("OPC") = "2" then
        Response.Addheader "Content-Disposition", "attachment;"
        Response.ContentType = "application/vnd.ms-excel"
      end if
      %>
      <%  %>
      <link href="Includes/Styles/Consult.css" rel="stylesheet" type="text/css">

      <head>
      <%
      dim serv, usu,pass,Query,Vrfc,Vfracc,Vdescrip,Vcod,strDesplegado,Vckcve,Vclave,Vini,Vfin,tempstrOficina
         serv="localhost"
         usu="EXTRANET"
         pass="rku_admin"
         Query= ""
         tempstrOficina= ""
         tempstrOficina=""
        Vbf=Request.Form("bf")
        Vbd=Request.Form("bd")
        Vrfc=Request.Form("txtCliente")
        Vckcve=Request.Form("ckcve")
        Vclave=Request.Form("cveCliente")
        Vfracc=Request.Form("fracc")
        Vdescrip=Request.Form("descrip")
        Vcod=Request.Form("codigo")
        VAdu1=Request.Form("aduan1")
        VAdu2=Request.Form("aduan2")
        VAdu3=Request.Form("aduan3")
        VAdu4=Request.Form("aduan4")
        VAdu5=Request.Form("aduan5")
        tempstrOficina=Request.Form("NomAdu")

'Response.Write(" RFC ")
'  Response.Write(Vrfc)
'Response.Write(" CVE ")
'  Response.Write(Vclave)
'Response.Write(" check ")
'  Response.Write(Vckcve)
'Response.Write(" fracc ")
'  Response.Write(Vfracc)
'Response.Write(" DEsc ")
'  Response.Write(Vdescrip)
'Response.Write(" cod ")
'  Response.Write(Vcod)
'Response.Write(" aduanas:  ")
'  Response.Write(VAdu1)
'Response.Write(" -- ")
'  Response.Write(VAdu2)
'Response.Write(" -- ")
'  Response.Write(VAdu3)
'Response.Write(" -- ")
'  Response.Write(VAdu4)
'Response.Write(" -- ")
'  Response.Write(VAdu5)
'Response.Write(" -- ")
'  Response.Write(tempstrOficina)

        'Response.End
      FILA=0
      ' para ver si ya esta el WHERE
      p=0

      ' variables para el FOR
      Vini=1
      Vfin=5
      %>
      <%
        Query="select nomcli18,rfccli18,cvecli18,fraant05,desc05,cpro05 "&_
               "from c05artic inner join ssclie18 on c05artic.clie05=ssclie18.cvecli18 "&_
               "where " 'rfccli18=variable and cpro05=variable and fraant05 like ' %' and desc05 like ' %"
      'concatena el RFC para

      if Vckcve="0" then
          if Vrfc <> "t" and p=0 then
            Query= Query & " rfccli18='"&Vrfc&"'"
            p=1
            else
            if Vrfc = "t" and p=0 then
            Query= "select nomcli18,rfccli18,cvecli18,fraant05,desc05,cpro05 "&_
                   "from c05artic inner join ssclie18 on c05artic.clie05=ssclie18.cvecli18 where "
            p=0
            end if
          end if
      else
          if Vclave<>"" and p=0 then
              Query="select nomcli18,rfccli18,cvecli18,fraant05,desc05,cpro05 "&_
                   "from c05artic inner join ssclie18 on c05artic.clie05=ssclie18.cvecli18 "&_
                   "where cvecli18='"&Vclave&"'"
              p=1
              ' este case es para ver por cual OFICINA se hara la consulta del CLIENTE
              Select Case (tempstrOficina)
               Case "LAR":
                   Vini=1
                   Vfin=1
               Case "MAN":
                  Vini=2
                  Vfin=2
               Case "MEX":
                   Vini=3
                   Vfin=3
               Case "VER":
                   Vini=4
                   Vfin=4
               Case "TAM":
                   Vini=5
                   Vfin=5
               End Select
          end if
      end if
      'concatena la parte de fracciones
      if Vfracc <> "" and p=0 then

          select case(vbf)
            case 1:
                 Query= Query & "fraant05 like '"&Vfracc&"%'"
                 p=1
            case 2:
                  Query= Query & "fraant05 like '%"&Vfracc&"%'"
                  p=1
            case 3:
                 Query= Query & "fraant05 like '%"&Vfracc&"'"
                 p=1
          end select

        else
        if Vfracc <> "" and p=1 then

          select case(vbf)
            case 1:
                 Query= Query & " and fraant05 like '"&Vfracc&"%'"
                 p=1
            case 2:
                  Query= Query & " and fraant05 like '%"&Vfracc&"%'"
                  p=1
            case 3:
                 Query= Query & " and fraant05 like '%"&Vfracc&"'"
                 p=1
          end select

        end if
      end if

      'concatena la descripcion
      if Vdescrip <> "" and p=0 then

          select case(Vbd)
          case 1:
             Query= Query & "desc05 like '"&Vdescrip&"%'"
             p=1
          case 2:
              Query= Query & "desc05 like '%"&Vdescrip&"%'"
              p=1
          end select

        else
        if Vdescrip <> "" and p=1 then
          select case(Vbd)
            case 1:
               Query= Query & " and desc05 like '"&Vdescrip&"%'"
               p=1
            case 2:
                Query= Query & " and desc05 like '%"&Vdescrip&"%'"
                p=1
          end select

        end if
      end if

      'concatena el codigo
      if Vcod <> "" and p=0 then
        Query= Query & "cpro05= '"&Vcod&"'"
        p=1
        else
        if Vcod <> "" and p=1 then
        Query= Query & " and cpro05= '"&Vcod&"'"
        p=1
        end if
      end if
      'si es TODOS y los demas campos estan vacios
      if Vrfc = "t" and p=0 and Vckcve="0" then
        Query= "select nomcli18,rfccli18,fraant05,cvecli18,desc05,cpro05 "&_
               "from c05artic inner join ssclie18 on c05artic.clie05=ssclie18.cvecli18 "
        p=1
      end if

      'Query=Query&" LIMIT  "&FILA&",40"
 '     Response.Write(Query)
'      Response.end

          'strDesplegado=strDesplegado & "<TR colspan=""7"" ><TH  bgcolor=""#336699"" align=""center""><FONT size=""3"" COLOR=""#ffffFF""><STRONG>..:º.º::. FRACCIONES POR CLIENTES .::º.º:..</STRONG></FONT></TD></TR>"& chr(13) & chr(10)
          strDesplegado= strDesplegado &   "<table align=""center""  border=""1"" Width=""1000"">" & chr(13) & chr(10)
          strDesplegado=strDesplegado &      "<TR><TH  colspan=""8""  bgcolor=""#336699"" align=""center""  ><FONT size=""3"" COLOR=""#ffffFF""><STRONG>FRACCIONES POR CLIENTE</STRONG></FONT></TD></TR>"& chr(13) & chr(10)
          strDesplegado= strDesplegado &     "<tr bgcolor=""#336699"">" & chr(13) & chr(10)
          strDesplegado= strDesplegado &          "<th><font color=""#FFFFFF"" size=""2"">i</th>" & chr(13) & chr(10)
          strDesplegado= strDesplegado &          "<th><font color=""#FFFFFF"" size=""2"">NOMBRE DEL CLIENTE</th>" & chr(13) & chr(10)
          strDesplegado= strDesplegado &          "<th><font color=""#FFFFFF"" size=""2"">RFC</th>" & chr(13) & chr(10)
          strDesplegado= strDesplegado &          "<th><font color=""#FFFFFF"" size=""2"">CLAVE</th>" & chr(13) & chr(10)
          strDesplegado= strDesplegado &          "<th><font color=""#FFFFFF"" size=""2"">FRACCION</th>" & chr(13) & chr(10)
          'strDesplegado= strDesplegado &          "<th bgcolor=""#3366CC""><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">OFICINA</font></strong></th>" & chr(13) & chr(10)
          strDesplegado= strDesplegado &          "<th><font color=""#FFFFFF"" size=""2"">DESCRIPCION</th>" & chr(13) & chr(10)
          strDesplegado= strDesplegado &          "<th><font color=""#FFFFFF"" size=""2"">CODIGO</th>" & chr(13) & chr(10)
          strDesplegado= strDesplegado &          "<th><font color=""#FFFFFF"" size=""2"">OFICINA</th>" & chr(13) & chr(10)
          strDesplegado= strDesplegado &     "</font></tr> " & chr(13) & chr(10)


      i=0

      for index= Vini to Vfin
           strOficina= ""
           MM_STRING  = ""
           ok="NOok"
           ii=0
           IF index=1 and Vadu1 = "lzr" THEN
              MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=lzr_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
              Vofi="LZR"
              ok="ok"
              'Response.Write("entro a: ")
              'Response.Write(index)
           END IF
           IF index=2 and Vadu2 = "sap" THEN
              MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=sap_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
              Vofi="SAP"
              ok="ok"
              'Response.Write("entro a: ")
              'Response.Write(index)
           END IF
           IF index=3 and Vadu3 = "dai" THEN
              MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=dai_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
              Vofi="DAI"
              ok="ok"
              'Response.Write("entro a: ")
              'Response.Write(index)
           END IF
           IF index= 4 and Vadu4 = "rku" THEN
              MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=rku_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
              Vofi="RKU"
              ok="ok"
              'Response.Write("entro a: ")
              'Response.Write(index)
           END IF
           IF index=5  and Vadu5 = "ceg" THEN
              MM_STRING = "DRIVER={MySQL ODBC 5.1 Driver}; SERVER="&serv&"; DATABASE=ceg_extranet; UID="&usu&"; PWD="&pass&"; OPTION=16427"
              Vofi="CEG"
              ok="ok"
              'Response.Write("entro a: ")
              'Response.Write(index)
           END IF
           '''''''

          if ok="ok" then
              set Rsio = server.CreateObject("ADODB.Recordset")
               Rsio.ActiveConnection = MM_STRING
               Rsio.Source= Query
               Rsio.CursorType = 0
               Rsio.CursorLocation = 2
               Rsio.LockType = 1
               Rsio.Open()

              do while not Rsio.EOF
              ii=ii+1
                 strDesplegado=strDesplegado & "<tr>" & chr(13) & chr(10)
                     strDesplegado=strDesplegado & "<td><font color=""#000000"" size=""1"">"&ii&"</font></td>" & chr(13) & chr(10)
                     'if Rsio.Fields.Item("nomcli18").Value<>"" then
                      strDesplegado=strDesplegado & "<td><font color=""#000000"" size=""1"">"&Rsio.Fields.Item("nomcli18").Value&"</td>" & chr(13) & chr(10)
                     'else
                      'strDesplegado=strDesplegado & "<td >".....</td>" & chr(13) & chr(10)
                     'end if
                     if Rsio.Fields.Item("rfccli18").Value<>"" then
                      strDesplegado=strDesplegado & "<td><font color=""#000000"" size=""1"">" &Rsio.Fields.Item("rfccli18").Value& "</td>" & chr(13) & chr(10)
                     else
                      strDesplegado=strDesplegado & "<td >.....</td>" & chr(13) & chr(10)
                     end if
                     if Rsio.Fields.Item("cvecli18").Value<>"" then
                      strDesplegado=strDesplegado & "<td><font color=""#000000"" size=""1"">" &Rsio.Fields.Item("cvecli18").Value& "</td>" & chr(13) & chr(10)
                     else
                      strDesplegado=strDesplegado & "<td >.....</td>" & chr(13) & chr(10)
                     end if
                     if Rsio.Fields.Item("fraant05").Value<>"" then
                      strDesplegado=strDesplegado & "<td><font color=""#000000"" size=""1"">"&Rsio.Fields.Item("fraant05").Value&"</td>" & chr(13) & chr(10)
                     else
                      strDesplegado=strDesplegado & "<td>.....</td>" & chr(13) & chr(10)
                     end if
                     descripcion=Rsio.Fields.Item("desc05").Value
                     if descripcion<>"" then
                      strDesplegado=strDesplegado & "<td><font color=""#000000"" size=""1"">"&descripcion&"</td>" & chr(13) & chr(10)
                     else
                      strDesplegado=strDesplegado & "<td >.....</td>" & chr(13) & chr(10)
                     end if
                     if Rsio.Fields.Item("cpro05").Value<>"" then
                      strDesplegado=strDesplegado & "<td><font color=""#000000"" size=""1"">"&Rsio.Fields.Item("cpro05").Value&"</td>" & chr(13) & chr(10)
                     else
                      strDesplegado=strDesplegado & "<td>.....</td>" & chr(13) & chr(10)
                     end if
                     'if Vofi<>"" then
                      strDesplegado=strDesplegado & "<td><font color=""#000000"" size=""1"">" &Vofi& "</td>" & chr(13) & chr(10)
                     'else
                      'strDesplegado=strDesplegado & "<td >.....</td>" & chr(13) & chr(10)
                     'end if
                     'strDesplegado=strDesplegado & "<td>"&Rsio.Fields.Item("fraant05").Value&"</td>" & chr(13) & chr(10)
                  strDesplegado=strDesplegado & "</font></tr>" & chr(13) & chr(10)
              Rsio.MoveNext
               i=i+1
              loop

          end if
      'Rsio.Close()
      'Set Rsio = Nothing
      next

          strDesplegado=strDesplegado & "</tr>" & chr(13) & chr(10)
          strDesplegado=strDesplegado & "</table><br>" & chr(13) & chr(10)


      Response.Write(strDesplegado)
      Response.Write("TOTAL DE REGISTROS ENCONTRADOS: ")
      Response.Write(i)

      %>
        <!--tr>
           <td>
              <br></br><br></br><div align="center"><a href="..Extranet\ext-Asp\FraccionesXCliente.asp" class="Boton"> .:..: VOLVER :..:. </a><br></div>
           </td>
        </tr-->
      </head>
      <%  %>
<%else
  response.write("<br></br><div align=""center""><p  class=""Titulo1"">:: USUARIO NO HABILITADO PARA ACCESAR A LOS REPORTES ::</div></p></div>")
end if%>