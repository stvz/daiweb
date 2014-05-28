
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/BDsystem.asp"   -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Connections/Configura.asp"  -->
 <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/Includes/ext_funciones.asp" -->

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
          formatofechaNum = YEAR(DFecha) & Pd(Month( DFecha ),2) & Pd(DAY( DFecha ),2)
       else
          formatofechaNum	= DFecha
       end if
    End Function


    Function diasTrimFinSemana(DInicio, DFin)
         x_Dias = 0
         x_Dias = dateDiff("d", DInicio , DFin )

         'Response.Write("entrada")
         'Response.Write(DInicio)
         'Response.Write("Salida")
         'Response.Write(DFin)
         'Response.Write("Diferencia")
         'Response.Write(x_Dias)



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

         'Response.Write("Diferencia")
         'Response.Write(x_Dias)
         'Response.End

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

         ' if isdate(DFecha) then
         '    DNewFecha = DateAdd("d",IntDayAdd, DFecha )
         '    numDia= WeekDay( DNewFecha )
         '    if numDia=1 then ' domingo
         '      DNewFecha = DateAdd("d",1, DNewFecha  )
         '    else
         '      if numDia=7 then ' Sabado
         '         DNewFecha = DateAdd("d",2, DNewFecha  )
         '      end if
         '    end if
         ' else
         '    Response.Write("no es fecha")
         ' end if
         ' SumarDiasSinFinSemana =  DNewFecha
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

    '--------------------------------------------------------------------------------------------------------------------------------
    'Funcion para escribir el encabezado del reporte en la cadena HTML
    function DespliegaEncabezado()
       strHTML = strHTML & " <br> "
       strHTML = strHTML & "<strong><font color=""#000066"" size=""3"" face=""Arial, Helvetica, sans-serif"">GRUPO REYES KURI, S.C. </font></strong> <br> "
       strHTML = strHTML & "<strong><font color=""#969696"" size=""3"" face=""Arial, Helvetica, sans-serif""> TRACKING AEREO " & " </font></strong>"
       'strHTML = strHTML & "<table bordercolor=""#C1C1C1"" border=""2"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)
       strHTML = strHTML & "<table bordercolor=""#7D997D"" border=""1"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & chr(13) & chr(10)


       contCamposplantilla = UBound(arrcampos,2) - 1
       strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)
       For intRow = 0 To contCamposplantilla
           strHTML = strHTML & "<td width=""120"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & arrcampos(2,intRow) & " </font></strong></td>" & chr(13) & chr(10)
       next
       strHTML = strHTML & "</tr>"& chr(13) & chr(10)




       ''*****************************************************************
       'strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""80""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REFERENCIA                           </font></strong></td>" & chr(13) & chr(10) '1 REFERENCIA
       'strHTML = strHTML & "<td width=""45""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> OTD2                                 </font></strong></td>" & chr(13) & chr(10) '2 OTD2
       'strHTML = strHTML & "<td width=""110"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ITTS/NOTIF DATE                      </font></strong></td>" & chr(13) & chr(10) '3 ITTS/NOTIF DATE
       'strHTML = strHTML & "<td width=""120"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> B. OF L. / AW. B. M.                 </font></strong></td>" & chr(13) & chr(10) '4 B. OF L. / AW. B. M.
       'strHTML = strHTML & "<td width=""120"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CONTAINER/ AW. B. H.                 </font></strong></td>" & chr(13) & chr(10) '5 CONTAINER/ AW. B. H.
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> P/O                                  </font></strong></td>" & chr(13) & chr(10) '6 P/O
       'strHTML = strHTML & "<td width=""120"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PORT/AIRPORT OF DEPARTURE            </font></strong></td>" & chr(13) & chr(10) '7 PORT/AIRPORT OF DEPARTURE --AEROPUERTO DE SALIDA
       'strHTML = strHTML & "<td width=""120"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ARRIVAL PORT/AIRPORT                 </font></strong></td>" & chr(13) & chr(10) '8 ARRIVAL PORT/AIRPORT      --PUERTO DESTINO
       'strHTML = strHTML & "<td width=""100"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CUSTOM OF DISPATCH                   </font></strong></td>" & chr(13) & chr(10) '8 9 CUSTOM OF DISPATCH      --PUERTO DESPACHO
       'strHTML = strHTML & "<td width=""100""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SHIPPING LINE /FORWARDER             </font></strong></td>" & chr(13) & chr(10) '10 SHIPPING LINE /FORWARDER  --anteriormente FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""100""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VESSEL                               </font></strong></td>" & chr(13) & chr(10) '11 VESSEL
       'strHTML = strHTML & "<td width=""100""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IMPORT DOCUMENT	                    </font></strong></td>" & chr(13) & chr(10) '12 IMPORT DOCUMENT
       'strHTML = strHTML & "<td width=""90""   bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SHIPPER    	                        </font></strong></td>" & chr(13) & chr(10) '13 SHIPPER  --anteriormente PROVEEDOR
       'strHTML = strHTML & "<td width=""90""   bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> INVOICE	                            </font></strong></td>" & chr(13) & chr(10) '14 INVOICE
       'strHTML = strHTML & "<td width=""120""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESCRIPTION CODE	                    </font></strong></td>" & chr(13) & chr(10) '15 DESCRIPTION CODE
       'strHTML = strHTML & "<td width=""120""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> MODEL 	                              </font></strong></td>" & chr(13) & chr(10) '16 MODEL
       'strHTML = strHTML & "<td width=""90""   bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESCRIPTION          	              </font></strong></td>" & chr(13) & chr(10) '17 DESCRIPTION  --  DESCRIPCION COMERCIAL
       'strHTML = strHTML & "<td width=""80""   bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> QTY                                  </font></strong></td>" & chr(13) & chr(10) '18 QTY
       'strHTML = strHTML & "<td width=""90""   bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETD LOAD /ATD ORIGIN                 </font></strong></td>" & chr(13) & chr(10) '19 ETD LOAD /ATD ORIGIN
       'strHTML = strHTML & "<td width=""90""   bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REMARKS BEFORE ARRIVAL PORT/LAX      </font></strong></td>" & chr(13) & chr(10) '20 REMARKS BEFORE ARRIVAL PORT/LAX
       'strHTML = strHTML & "<td width=""90""   bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETA PORT /ETA LAX                    </font></strong></td>" & chr(13) & chr(10) '21 ETA PORT/LAX
       'strHTML = strHTML & "<td width=""90""   bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATA PORT/CUSTOM (LAX)                </font></strong></td>" & chr(13) & chr(10) '22 ATA PORT/CUSTOM (LAX)
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REMARKS PORT/LAX                     </font></strong></td>" & chr(13) & chr(10) '23 REMARKS  PORT/LAX
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SERIAL NUMBER                        </font></strong></td>" & chr(13) & chr(10) '24 SERIAL NUMBER
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CERT. NOM                            </font></strong></td>" & chr(13) & chr(10) '25 CERT. NOM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DATE OF RELEASE                      </font></strong></td>" & chr(13) & chr(10) '26 DATE OF RELEASE -- REVALIDACION
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> RESQUEST DUTIES                      </font></strong></td>" & chr(13) & chr(10) '27 RESQUEST DUTIES
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> AMOUNT OF DUTIES                     </font></strong></td>" & chr(13) & chr(10) '28 AMOUNT OF DUTIES
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PREVIO                               </font></strong></td>" & chr(13) & chr(10) '29 PREVIO
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETA DATE OF CLEARANCE                </font></strong></td>" & chr(13) & chr(10) '30 ETA DATE OF CLEARANCE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DATE OF CLEARANCE                    </font></strong></td>" & chr(13) & chr(10) '31 DATE OF CLEARANCE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REMARKS CLEARANCE                    </font></strong></td>" & chr(13) & chr(10) '32 REMARKS CLEARANCE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATD  RAIL                            </font></strong></td>" & chr(13) & chr(10) '33 ATD RAIL
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REMARKS ATD RAIL                     </font></strong></td>" & chr(13) & chr(10) '34 REMARKS ATD RAIL
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETA C./P.                            </font></strong></td>" & chr(13) & chr(10) '35 ETA C./P.
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATA C./P.                            </font></strong></td>" & chr(13) & chr(10) '36 ATA C./P.
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REMARKS  C./P.                       </font></strong></td>" & chr(13) & chr(10) '37 REMARKS  C./P.
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETA W/H                              </font></strong></td>" & chr(13) & chr(10) '38 ETA W/H
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATA W-H                              </font></strong></td>" & chr(13) & chr(10) '39 ATA W-H
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TRUCK ARRIVE TIME                    </font></strong></td>" & chr(13) & chr(10) '40 TRUCK ARRIVE TIME -- TIME OF DELIVERY IN SEM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TRUCK  DEPARTURE FROM W/H            </font></strong></td>" & chr(13) & chr(10) '41 TRUCK  DEPARTURE FROM W/H  --SALIDA DE ALMACEN
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> STATUS                               </font></strong></td>" & chr(13) & chr(10) '42 STATUS
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REMARKS (ULTIMO)                     </font></strong></td>" & chr(13) & chr(10) '43 REMARKS (ULTIMO)
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> KPI STATUS                           </font></strong></td>" & chr(13) & chr(10) '44 KPI STATUS
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> MODALIDAD                            </font></strong></td>" & chr(13) & chr(10) '45 MODALIDAD
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> WEEK                                 </font></strong></td>" & chr(13) & chr(10) '46 WEEK
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> NUM. INVOICE CUSTOM                  </font></strong></td>" & chr(13) & chr(10) '47 NUM. INVOICE CUSTOM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DATE OF INVOICE CUSTOM               </font></strong></td>" & chr(13) & chr(10) '48 DATE OF INVOICE CUSTOM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> HISTORIAL                            </font></strong></td>" & chr(13) & chr(10) '49 HISTORIAL
       'strHTML = strHTML & "</tr>"& chr(13) & chr(10)
       ''*****************************************************************


       'strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)
       'strHTML = strHTML & "<td width=""80""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REFERENCIA                           </font></strong></td>" & chr(13) & chr(10) 'OTD 2
       'strHTML = strHTML & "<td width=""50""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> OTD2                                 </font></strong></td>" & chr(13) & chr(10) 'OTD2
       'strHTML = strHTML & "<td width=""100"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ASIGNADO ITTS/DATE OF NOTIFICATION   </font></strong></td>" & chr(13) & chr(10) 'ASIGNADO ITTS
       'strHTML = strHTML & "<td width=""100"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> BILL OF LADING / AIRWAY BILL MASTER  </font></strong></td>" & chr(13) & chr(10) 'GUIA MASTER
       'strHTML = strHTML & "<td width=""100"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CONTAINER/ AIRWAY BILL HOUSE         </font></strong></td>" & chr(13) & chr(10) 'GUIA HOUSE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> P/O                                  </font></strong></td>" & chr(13) & chr(10) 'P/O
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PORT/AIRPORT OF DEPARTURE            </font></strong></td>" & chr(13) & chr(10) 'PORT / AIRPORT OF DEPARTURE --AEROPUERTO DE SALIDA
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CUSTOM OF DISPATCH                   </font></strong></td>" & chr(13) & chr(10) 'PORT OF DISCHARGE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SHIPPING LINE /FORWARDER             </font></strong></td>" & chr(13) & chr(10) 'SHIPPING LINE /FORWARDER  --anteriormente FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> VESSEL                               </font></strong></td>" & chr(13) & chr(10) 'VESSEL
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IMPORT DOCUMENT	                    </font></strong></td>" & chr(13) & chr(10) 'IMPORT DOCUMENT
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SHIPPER    	                        </font></strong></td>" & chr(13) & chr(10) 'SHIPPER  --anteriormente PROVEEDOR
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> INVOICE	                            </font></strong></td>" & chr(13) & chr(10) 'INVOICE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESCRIPTION CODE	                    </font></strong></td>" & chr(13) & chr(10) 'DESCRIPTION CODE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> MODEL 	                              </font></strong></td>" & chr(13) & chr(10) 'MODEL
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESCRIPTION          	              </font></strong></td>" & chr(13) & chr(10) 'DESCRIPCION COMERCIAL
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> QTY                                  </font></strong></td>" & chr(13) & chr(10) 'QTY
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETD LOAD /ATD  ORIGIN                </font></strong></td>" & chr(13) & chr(10) 'ETD LOAD
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETA PORT /ETA LAX                    </font></strong></td>" & chr(13) & chr(10) 'ETA PORT
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATA PORT/ATA CUSTOM                  </font></strong></td>" & chr(13) & chr(10) 'ATA PORT
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SERIAL NUMBER                        </font></strong></td>" & chr(13) & chr(10) 'NUMS. SERIE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> CERT. NOM                            </font></strong></td>" & chr(13) & chr(10) 'CERT. NOM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DATE OF RELEASE                      </font></strong></td>" & chr(13) & chr(10) 'REVALIDACION
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> RESQUEST DUTIES                      </font></strong></td>" & chr(13) & chr(10) 'RESQUEST DUTIES
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> AMOUNT OF DUTIES                     </font></strong></td>" & chr(13) & chr(10) 'AMOUNT OF DUTIES
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PREVIO                               </font></strong></td>" & chr(13) & chr(10) 'PREVIO
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETA DATE OF CLEARANCE                </font></strong></td>" & chr(13) & chr(10) 'DATE OF CUSTOM CLEARANCE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DATE OF CLEARANCE                    </font></strong></td>" & chr(13) & chr(10) 'DATE OF CUSTOM CLEARANCE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATD  RAIL                            </font></strong></td>" & chr(13) & chr(10) 'ATD  RAIL
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETA C./P.                            </font></strong></td>" & chr(13) & chr(10) 'ETA C./P.
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATA C./P.                            </font></strong></td>" & chr(13) & chr(10) 'ATA C./P.
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETA W/H                              </font></strong></td>" & chr(13) & chr(10) 'ETA W/H
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATA W-H                              </font></strong></td>" & chr(13) & chr(10) 'ATA W-H
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TIME OF DELIVERY IN SEM              </font></strong></td>" & chr(13) & chr(10) 'TIME OF DELIVERY IN SEM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REMARKS                              </font></strong></td>" & chr(13) & chr(10) 'REMARKS
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> MODALIDAD                            </font></strong></td>" & chr(13) & chr(10) 'MODALIDAD
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> WEEK                                 </font></strong></td>" & chr(13) & chr(10) 'WEEK
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> NUM. INVOICE CUSTOM                  </font></strong></td>" & chr(13) & chr(10) 'NUM. INVOICE CUSTOM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DATE OF INVOICE CUSTOM               </font></strong></td>" & chr(13) & chr(10) 'DATE OF INVOICE CUSTOM
       'strHTML = strHTML & "</tr>"& chr(13) & chr(10)




       'strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)
       ''--------------------------------------------------
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REFERENCIA              </font></strong></td>" & chr(13) & chr(10) 'OTD 2
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> OTD2                    </font></strong></td>" & chr(13) & chr(10) 'OTD2
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ASIGNADO ITTS           </font></strong></td>" & chr(13) & chr(10) 'OTD2
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> GUIA MASTER	           </font></strong></td>" & chr(13) & chr(10) 'GUIA MASTER
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> GUIA HOUSE	             </font></strong></td>" & chr(13) & chr(10) 'GUIA HOUSE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> P/O                     </font></strong></td>" & chr(13) & chr(10) 'P/O
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> AEROPUERTO DE SALIDA	   </font></strong></td>" & chr(13) & chr(10) 'AEROPUERTO DE SALIDA
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">                          </font></strong></td>" & chr(13) & chr(10)'FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FORWARDER Y/O AIR  LINE </font></strong></td>" & chr(13) & chr(10)'FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">                          </font></strong></td>" & chr(13) & chr(10)'FORWARDER Y/O AIR  LINE
       ''strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> NOTIFICACION DE GUIA	   </font></strong></td>" & chr(13) & chr(10) 'NOTIFICACION DE GUIA
       ''strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA DE NOTIFICACION 	 </font></strong></td>" & chr(13) & chr(10) 'FECHA DE NOTIFICACION
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> IMPORT DOCUMENT	       </font></strong></td>" & chr(13) & chr(10) 'IMPORT DOCUMENT
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> PROVEEDOR	             </font></strong></td>" & chr(13) & chr(10) 'PROVEEDOR
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> INVOICE	               </font></strong></td>" & chr(13) & chr(10) 'INVOICE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESCRIPTION CODE	       </font></strong></td>" & chr(13) & chr(10) 'DESCRIPTION CODE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> MODEL 	                 </font></strong></td>" & chr(13) & chr(10) 'MODEL
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DESCRIPCION COMERCIAL	 </font></strong></td>" & chr(13) & chr(10) 'DESCRIPCION COMERCIAL
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> QTY	                   </font></strong></td>" & chr(13) & chr(10) 'QTY
       ''strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> INCOTERMS	             </font></strong></td>" & chr(13) & chr(10) 'INCOTERMS
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ORIGIN   ETD	           </font></strong></td>" & chr(13) & chr(10) 'ORIGIN   ETD
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ETA  LAX./ 	           </font></strong></td>" & chr(13) & chr(10) 'ETA  LAX./
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATA CUSTOM	             </font></strong></td>" & chr(13) & chr(10) 'ATA CUSTOM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> SERIAL NUMBER	         </font></strong></td>" & chr(13) & chr(10) 'SERIAL NUMBER
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">                          </font></strong></td>" & chr(13) & chr(10)'FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA DE  REVALIDACION	 </font></strong></td>" & chr(13) & chr(10) 'FECHA DE  REVALIDACION
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> RESQUEST DUTIES	       </font></strong></td>" & chr(13) & chr(10) 'RESQUEST DUTIES
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> AMOUNT OF DUTIES	       </font></strong></td>" & chr(13) & chr(10) 'AMOUNT OF DUTIES
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> FECHA DE PREVIO	       </font></strong></td>" & chr(13) & chr(10) 'FECHA DE PREVIO
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DATE OF CLEARANCE	     </font></strong></td>" & chr(13) & chr(10) 'DATE OF CLEARANCE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">                          </font></strong></td>" & chr(13) & chr(10)'FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">                          </font></strong></td>" & chr(13) & chr(10)'FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">                          </font></strong></td>" & chr(13) & chr(10)'FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">                          </font></strong></td>" & chr(13) & chr(10)'FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> ATA W-H	               </font></strong></td>" & chr(13) & chr(10) 'ATA W-H
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> TIME OF DELIVERY IN SEM </font></strong></td>" & chr(13) & chr(10)'TIME OF DELIVERY IN SEM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> REMARKS	               </font></strong></td>" & chr(13) & chr(10) 'REMARKS
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif"">                          </font></strong></td>" & chr(13) & chr(10)'FORWARDER Y/O AIR  LINE
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> WEEK	                   </font></strong></td>" & chr(13) & chr(10) 'WEEK
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> NUM. INVOICE CUSTOM	   </font></strong></td>" & chr(13) & chr(10) 'NUM. INVOICE CUSTOM
       'strHTML = strHTML & "<td width=""90""  bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> DATE OF INVOICE CUSTOM  </font></strong></td>" & chr(13) & chr(10) 'DATE OF INVOICE CUSTOM
       'strHTML = strHTML & "</tr>"& chr(13) & chr(10)

    end function




     '-------------------------------------------------------------------------------------------------------------------------------
    'Funcion para escribir el encabezado del reporte en la cadena HTML
    function agregarfilaHTML(COLOR,      REFERENCIA,      OTD2,     GUIA_MASTER,     GUIA_HOUSE   ,  P_OA ,   FORWARDER_AIR_LINE,     AEROPUERTO_SALIDA,     CUSTOM_OF_DISPATCH,    NOTIFICACION_GUIA,     FECHA_NOTIFICACION,     VESSEL,     IMPORT_DOCUMENT,     PROVEEDOR,     INVOICE,     MODEL,     DESCRIPCION_COMERCIAL,     DESCRIPTION_CODE,     QTY,     INCOTERMS,     SERIAL_NUMBER,     CERT_NOM,	    ORIGIN_ETD,     ETA_LAX,     ATA_CUSTOM,     RESQUEST_DUTIES,     FECHA_DE_REVALIDACION,     FECHA_DE_PREVIO,     ETA_CUSTOM_CLEARANCE ,    DATE_OF_CLEARANCE,     ETA_C_P,      ATA_CP ,    ETA_W_H,     ATA_WH,     TIMEOFDELIVERY,      REMARKS,     MODALIDAD ,     WEEK,     AMOUNTOFDUTIES,     NUM_INVOICECUSTOM,     DATEINVOICECUSTOM,     ADUDESPACHO, RMKATDORIGIN,    RMKATAPORT,    RMKDEPACHO,    RMKATDRAIL,    RMKCP,    ATASPL,    STATUS,    LASTRMK,    KPISTATUS )
              'agregarfilaHTML StrColorfila, StrREFERENCIA, StrPOTD2, StrPGUIA_MASTER, StrPGUIA_HOUSE, StrPP_O, StrPFORWARDER_AIR_LINE, StrPAEROPUERTO_SALIDA, StrCUSTOM_OF_DISPATCH, StrPNOTIFICACION_GUIA, StrPFECHA_NOTIFICACION, StrPVESSEL, StrPIMPORT_DOCUMENT, StrPPROVEEDOR, StrPINVOICE, StrPMODEL, StrPDESCRIPCION_COMERCIAL, StrPDESCRIPTION_CODE, StrPQTY, StrPINCOTERMS, StrPSERIAL_NUMBER, StrPCERT_NOM,	StrPORIGIN_ETD, StrPETA_LAX, StrPATA_CUSTOM, StrPRESQUEST_DUTIES, StrPFECHA_DE_REVALIDACION, StrPFECHA_DE_PREVIO, StrPETA_CUSTOM_CLEARANCE, StrPDATE_OF_CLEARANCE, StrPETA_C_P,  StrPATA_CP ,StrPETA_W_H, StrPATA_WH, StrPTIMEOFDELIVERY , StrPREMARKS, StrPMODALIDAD , StrPWEEK, StrPAMOUNTOFDUTIES, StrPNUM_INVOICECUSTOM, StrPDATEINVOICECUSTOM,              strRMKATDORIGIN, strRMKATAPORT, strRMKDEPACHO, strRMKATDRAIL, strRMKCP, strATASPL, strSTATUS, strLASTRMK, strKPISTATUS

       'COLOR
       'REFERENCIA
       'OTD2
       'GUIA_MASTER
       'GUIA_HOUSE
       'P_OA
       'FORWARDER_AIR_LINE
       'AEROPUERTO_SALIDA
       'CUSTOM_OF_DISPATCH
       'NOTIFICACION_GUIA
       'FECHA_NOTIFICACION
       'VESSEL
       'IMPORT_DOCUMENT
       'PROVEEDOR
       'INVOICE
       'MODEL
       'DESCRIPCION_COMERCIAL
       'DESCRIPTION_CODE
       'QTY
       'INCOTERMS
       'SERIAL_NUMBER
       'CERT_NOM
       'ORIGIN_ETD
       'ETA_LAX
       'ATA_CUSTOM
       'RESQUEST_DUTIES
       'FECHA_DE_REVALIDACION
       'FECHA_DE_PREVIO
       'ETA_CUSTOM_CLEARANCE
       'DATE_OF_CLEARANCE
       'ETA_C_P
       'ATA_CP
       'ETA_W_H
       'ATA_WH
       'TIMEOFDELIVERY
       'REMARKS
       'MODALIDAD
       'WEEK
       'AMOUNTOFDUTIES
       'NUM_INVOICECUSTOM
       'DATEINVOICECUSTOM
       'ADUDESPACHO
       'RMKATDORIGIN
       'RMKATAPORT
       'RMKDEPACHO
       'RMKATDRAIL
       'RMKCP
       'ATASPL
       'STATUS
       'LASTRMK
       'KPISTATUS

       contCamposplantilla = UBound(arrcampos,2) - 1
       'strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)
       For intRow = 0 To contCamposplantilla

           if arrcampos(1,intRow) = "REFERENCIA"  then' Nombre del campo
             arrcampos(4,intRow) = REFERENCIA  ' titulo
           end if
           if arrcampos(1,intRow) = "OTD2"  then' Nombre del campo
             arrcampos(4,intRow) = OTD2  ' titulo
           end if
           if arrcampos(1,intRow) = "ITTSNOTIFDATE"  then' Nombre del campo
             arrcampos(4,intRow) =  FECHA_NOTIFICACION ' titulo
           end if
           if arrcampos(1,intRow) = "BOFL_AWBM"  then' Nombre del campo
             arrcampos(4,intRow) = GUIA_MASTER  ' titulo
           end if
           if arrcampos(1,intRow) = "CONTAINER_AWBH"  then' Nombre del campo
             arrcampos(4,intRow) = GUIA_HOUSE  ' titulo
           end if
           if arrcampos(1,intRow) = "PO"  then' Nombre del campo
             arrcampos(4,intRow) = P_OA  ' titulo
           end if
           if arrcampos(1,intRow) = "PORT_AIRPORTOFDEPARTURE"  then' Nombre del campo
             arrcampos(4,intRow) = AEROPUERTO_SALIDA  ' titulo
           end if
           if arrcampos(1,intRow) = "ARRIVALPORT_AIRPORT"  then' Nombre del campo
             arrcampos(4,intRow) = CUSTOM_OF_DISPATCH  ' titulo
           end if

           if arrcampos(1,intRow) = "SHIPPINGLINE_FORWARDER"  then' Nombre del campo
             arrcampos(4,intRow) = FORWARDER_AIR_LINE  ' titulo
           end if
           if arrcampos(1,intRow) = "VESSEL"  then' Nombre del campo
             arrcampos(4,intRow) = VESSEL  ' titulo
           end if
           if arrcampos(1,intRow) = "IMPORTDOCUMENT"  then' Nombre del campo
             arrcampos(4,intRow) = IMPORT_DOCUMENT  ' titulo
           end if
           if arrcampos(1,intRow) = "SHIPPER"  then' Nombre del campo
             arrcampos(4,intRow) = PROVEEDOR  ' titulo
           end if
           if arrcampos(1,intRow) = "INVOICE"  then' Nombre del campo
             arrcampos(4,intRow) = INVOICE  ' titulo
           end if
           if arrcampos(1,intRow) = "MODEL"  then' Nombre del campo
             arrcampos(4,intRow) = MODEL  ' titulo
           end if
           if arrcampos(1,intRow) = "DESCRIPTION"  then' Nombre del campo
             arrcampos(4,intRow) = DESCRIPCION_COMERCIAL  ' titulo
           end if
           if arrcampos(1,intRow) = "DESCRIPTIONCODE"  then' Nombre del campo
             arrcampos(4,intRow) = DESCRIPTION_CODE  ' titulo
           end if
           if arrcampos(1,intRow) = "QTY"  then' Nombre del campo
             arrcampos(4,intRow) = QTY ' titulo
           end if
           if arrcampos(1,intRow) = "ETDLOAD_ATDORIGIN"  then' Nombre del campo
             arrcampos(4,intRow) = ORIGIN_ETD  ' titulo
           end if
           if arrcampos(1,intRow) = "ETAPORT_LAX"  then' Nombre del campo
             arrcampos(4,intRow) = ETA_LAX  ' titulo
           end if
           if arrcampos(1,intRow) = "ATAPORT_CUSTOMLAX"  then' Nombre del campo
             arrcampos(4,intRow) = ATA_CUSTOM  ' titulo
           end if
           if arrcampos(1,intRow) = "SERIALNUMBER"  then' Nombre del campo
             arrcampos(4,intRow) = SERIAL_NUMBER  ' titulo
           end if
           if arrcampos(1,intRow) = "CERTNOM"  then' Nombre del campo
             arrcampos(4,intRow) = CERT_NOM  ' titulo
           end if
           if arrcampos(1,intRow) = "DATEOFRELEASE"  then' Nombre del campo
             arrcampos(4,intRow) = FECHA_DE_REVALIDACION  ' titulo
           end if
           if arrcampos(1,intRow) = "RESQUESTDUTIES"  then' Nombre del campo
             arrcampos(4,intRow) = RESQUEST_DUTIES  ' titulo
           end if
           if arrcampos(1,intRow) = "AMOUNTOFDUTIES"  then' Nombre del campo
             arrcampos(4,intRow) = AMOUNTOFDUTIES  ' titulo
           end if
           if arrcampos(1,intRow) = "PREVIO"  then' Nombre del campo
             arrcampos(4,intRow) = FECHA_DE_PREVIO  ' titulo
           end if
           if arrcampos(1,intRow) = "ETADATEOFCLEARANCE"  then' Nombre del campo
             arrcampos(4,intRow) = ETA_CUSTOM_CLEARANCE  ' titulo
           end if
           if arrcampos(1,intRow) = "DATEOFCLEARANCE"  then' Nombre del campo
             arrcampos(4,intRow) = DATE_OF_CLEARANCE  ' titulo
           end if
           if arrcampos(1,intRow) = "ATDRAIL"  then' Nombre del campo
             arrcampos(4,intRow) = "N/A"  ' titulo
           end if
           if arrcampos(1,intRow) = "ETACP"  then' Nombre del campo
             arrcampos(4,intRow) = "N/A"  ' titulo
           end if
           if arrcampos(1,intRow) = "ATACP"  then' Nombre del campo
             arrcampos(4,intRow) = "N/A"  ' titulo
           end if
           if arrcampos(1,intRow) = "ETAWH"  then' Nombre del campo
             arrcampos(4,intRow) = ETA_W_H  ' titulo
           end if
           if arrcampos(1,intRow) = "ATAWH"  then' Nombre del campo
             arrcampos(4,intRow) = ATA_WH  ' titulo
           end if
           if arrcampos(1,intRow) = "TRUCKARRIVETIMEWH"  then' Nombre del campo
             arrcampos(4,intRow) = TIMEOFDELIVERY  ' titulo
           end if
           if arrcampos(1,intRow) = "HISTORIAL"  then' Nombre del campo
             arrcampos(4,intRow) = REMARKS  ' titulo
           end if
           if arrcampos(1,intRow) = "MODALIDAD"  then' Nombre del campo
             arrcampos(4,intRow) = MODALIDAD  ' titulo
           end if
           if arrcampos(1,intRow) = "WEEK"  then' Nombre del campo
             arrcampos(4,intRow) = WEEK  ' titulo
           end if
           if arrcampos(1,intRow) = "NUMINVOICECUSTOM"  then' Nombre del campo
             arrcampos(4,intRow) = NUM_INVOICECUSTOM  ' titulo
           end if
           if arrcampos(1,intRow) = "DATEOFINVOICECUSTOM"  then' Nombre del campo
             arrcampos(4,intRow) = DATEINVOICECUSTOM  ' titulo
           end if
           if arrcampos(1,intRow) = "CUSTOMOFDISPATCH"  then' Nombre del campo
             arrcampos(4,intRow) = ADUDESPACHO  ' titulo
           end if
           if arrcampos(1,intRow) = "RMKBFARRIVALPORT_LAX"  then' Nombre del campo
             arrcampos(4,intRow) = RMKATDORIGIN  ' titulo
           end if
           if arrcampos(1,intRow) = "REMARKS_PORTLAX"  then' Nombre del campo
             arrcampos(4,intRow) = RMKATAPORT  ' titulo
           end if
           if arrcampos(1,intRow) = "REMARKSCLEARANCE"  then' Nombre del campo - Despacho
             arrcampos(4,intRow) = "N/A"  ' titulo
           end if
           if arrcampos(1,intRow) = "REMARKSATDRAIL"  then' Nombre del campo
             arrcampos(4,intRow) = "N/A"  ' titulo
           end if
           if arrcampos(1,intRow) = "REMARKCP"  then' Nombre del campo
             arrcampos(4,intRow) = "N/A"  ' titulo
           end if
           if arrcampos(1,intRow) = "TRUCKDEPARTUREWH"  then' Nombre del campo
             arrcampos(4,intRow) = ATASPL  ' titulo
           end if
           if arrcampos(1,intRow) = "STATUS"  then' Nombre del campo
             arrcampos(4,intRow) = STATUS  ' titulo
           end if
           if arrcampos(1,intRow) = "REMARKSULTIMO"  then' Nombre del campo
             arrcampos(4,intRow) = LASTRMK  ' titulo
           end if
           if arrcampos(1,intRow) = "KPISTATUS"  then' Nombre del campo
             arrcampos(4,intRow) = KPISTATUS  ' titulo
           end if
       next

       '*******************************************************************************************************


       if COLOR=1 then
          str_color = "#FFFFFF"
          str_fcolor = "#000000"
       else
         if COLOR=2 then ' AZUL DIFERENCIA A FAVOR AGENCIA
            'str_color = "#0099FF"
            'str_fcolor = "#000000"
            str_color = "#FFFFFF"
            str_fcolor = "#0099FF"
         else
            if COLOR=3 then ' ROJO RETRASO
               'str_color = "#FF0000"
               'str_fcolor = "#000000"
               str_color = "#FFFFFF"
               str_fcolor = "#FF0000"
            end if
         end if
       end if
       'strColorNA = "#7D997D"
       'strColorNA = "#C1C1C1"
       strColorNA = "#DCDCDC"

       if strTipoFiltro  = "BotonOtrosOpVivas" and ATA_WH <> "" and not isnull(ATA_WH) then
           str_color = "#FFFFCC"
          'str_color = "#99CCFF"
       end if




       'ADUDESPACHO, RMKATDORIGIN, RMKATAPORT, RMKDEPACHO, RMKATDRAIL, RMKCP, ATASPL, STATUS, LASTRMK, KPISTATUS

       if COLOR <> 2 and COLOR <> 3 then

           strHTML = strHTML& "<tr bgcolor= '"&str_color&"' align=""center"" >"& chr(13) & chr(10)
           For intRow = 0 To contCamposplantilla
               'strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & arrcampos(4,intRow) &" </font></td>" & chr(13) & chr(10) '
               if arrcampos(4,intRow) = "N/A" then
                  strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " " & " </font></td>" & chr(13) & chr(10) '
               else
                  strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & arrcampos(4,intRow) & " </font></td>" & chr(13) & chr(10) '
               end if
           next
           strHTML = strHTML & "</tr>"& chr(13) & chr(10)
           '****************************************************************************************


            '           strHTML = strHTML& "<tr bgcolor= '"&str_color&"' align=""center"" >"& chr(13) & chr(10)
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & REFERENCIA            & " </font></td>" & chr(13) & chr(10) 'REFERENCIA
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & OTD2                  & " </font></td>" & chr(13) & chr(10) 'OTD2
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_NOTIFICACION 	  & " </font></td>" & chr(13) & chr(10) 'ASIGNADO ITTS /FECHA DE NOTIFICACION
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & GUIA_MASTER	          & " </font></td>" & chr(13) & chr(10) 'GUIA MASTER
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & GUIA_HOUSE	          & " </font></td>" & chr(13) & chr(10) 'GUIA HOUSE
            '           if P_OA = "N/A" then
            '           'if P_O <> "" and InStr(P_O,"N/A") > 0 and Len(P_O) = 3 then
            '              strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "              & " </font></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
            '           else
            '              strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & P_OA               & " </font></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
            '           end if
            '           'if trim(P_OA) = "N/A" then
            '           '   strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "              & " </font></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
            '           '   'response.write(P_OA)
            '           '   'Response.End
            '           'else
            '           '   strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & P_OA               & " </font></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
            '           'end if
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & AEROPUERTO_SALIDA	    & " </font></td>" & chr(13) & chr(10) 'AEROPUERTO DE SALIDA
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & CUSTOM_OF_DISPATCH    &" </font></td>" & chr(13) & chr(10) '7 PORT/AIRPORT OF DEPARTURE --AEROPUERTO DE SALIDA --PORT OF LOADING
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ADUDESPACHO           & " </font></td>" & chr(13) & chr(10) 'CUSTOM OF DISPATCH -- ADUANA SEGUN APENDICE 1 DEL ANEXO 22
            '
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FORWARDER_AIR_LINE    & " </font></td>" & chr(13) & chr(10) 'FORWARDER Y/O AIR  LINE
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></td>" & chr(13) & chr(10) 'VESSELN_GUIA	        & " </font></strong></td>" & chr(13) & chr(10) 'NOTIFICACION DE GUIA
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & IMPORT_DOCUMENT	      & " </font></td>" & chr(13) & chr(10) 'IMPORT DOCUMENT
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & PROVEEDOR	            & " </font></td>" & chr(13) & chr(10) 'PROVEEDOR
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & INVOICE	              & " </font></td>" & chr(13) & chr(10) 'INVOICE
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DESCRIPTION_CODE	    & " </font></td>" & chr(13) & chr(10) 'DESCRIPTION CODE
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & MODEL 	              & " </font></td>" & chr(13) & chr(10) 'MODEL
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DESCRIPCION_COMERCIAL & " </font></td>" & chr(13) & chr(10) 'DESCRIPCION COMERCIAL
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & QTY	                  & " </font></td>" & chr(13) & chr(10) 'QTY
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ORIGIN_ETD	          & " </font></td>" & chr(13) & chr(10) 'ORIGIN   ETD
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & RMKATDORIGIN      &" </font></td>" & chr(13) & chr(10) '20 REMARKS BEFORE ARRIVAL PORT/LAX
            '
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_LAX   	          & " </font></td>" & chr(13) & chr(10) 'ETA  LAX./
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATA_CUSTOM	          & " </font></td>" & chr(13) & chr(10) 'ATA CUSTOM
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & RMKATAPORT        &" </font></td>" & chr(13) & chr(10) '23 REMARKS  PORT/LAX
            '
            '           if SERIAL_NUMBER = "N/A" then
            '              strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ""  & " </font></td>" & chr(13) & chr(10) 'SERIAL NUMBER
            '           else
            '              strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & SERIAL_NUMBER	        & " </font></td>" & chr(13) & chr(10) 'SERIAL NUMBER
            '           end if
            '
            '           if CERT_NOM = "N/A" then
            '              strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ""  & " </font></td>" & chr(13) & chr(10) 'CERT. NOM
            '           else
            '              strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & CERT_NOM              & " </font></td>" & chr(13) & chr(10) 'CERT. NOM
            '           end if

            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_DE_REVALIDACION & " </font></td>" & chr(13) & chr(10) 'FECHA DE  REVALIDACION
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & RESQUEST_DUTIES	      & " </font></td>" & chr(13) & chr(10) 'RESQUEST DUTIES
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & AMOUNTOFDUTIES	      & " </font></td>" & chr(13) & chr(10) 'AMOUNT OF DUTIES
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_DE_PREVIO	      & " </font></td>" & chr(13) & chr(10) 'FECHA DE PREVIO
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_CUSTOM_CLEARANCE  & " </font></td>" & chr(13) & chr(10) 'ETA DATE OF CLEARANCE
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DATE_OF_CLEARANCE	    & " </font></td>" & chr(13) & chr(10) 'DATE OF CLEARANCE
            '
            '           'strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & RMKDEPACHO        &" </font></td>" & chr(13) & chr(10) '32 REMARKS CLEARANCE
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></td>" & chr(13) & chr(10) 'REMARKS CLEARANCE
            '
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></td>" & chr(13) & chr(10) 'ATD RAIL
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></td>" & chr(13) & chr(10) 'REMARKS ATD RAIL
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></td>" & chr(13) & chr(10) 'ETA C/P
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></td>" & chr(13) & chr(10) 'ATA C/P
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></td>" & chr(13) & chr(10) 'REMARKS  C./P.
            '
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_W_H               & " </font></td>" & chr(13) & chr(10) 'ETA W/H
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATA_WH	              & " </font></td>" & chr(13) & chr(10) 'ATA W-H
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & TIMEOFDELIVERY        & " </font></td>" & chr(13) & chr(10) 'TIME OF DELIVERY IN SEM
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATASPL                & " </font></td>" & chr(13) & chr(10) '41 TRUCK  DEPARTURE FROM W/H  --SALIDA DE ALMACEN
            '
            '           '**********************************************************************
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & STATUS            &" </font></td>" & chr(13) & chr(10) '42 STATUS
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & LASTRMK           &" </font></td>" & chr(13) & chr(10) '43 REMARKS (ULTIMO)
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & KPISTATUS         &" </font></td>" & chr(13) & chr(10) '44 KPI STATUS
            '
            '           'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & MODALIDAD             & " </font></strong></td>" & chr(13) & chr(10) 'MODALIDAD
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "               &" </font></td>" & chr(13) & chr(10) 'MODALIDAD
            '
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & WEEK              &" </font></td>" & chr(13) & chr(10) 'WEEK
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & NUM_INVOICECUSTOM &" </font></td>" & chr(13) & chr(10) 'NUM. INVOICE CUSTOM
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & DATEINVOICECUSTOM &" </font></td>" & chr(13) & chr(10) 'DATE OF INVOICE CUSTOM
            '           strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & REMARKS   &" </font></td>" & chr(13) & chr(10) 'HISTORIAL
            '           '**********************************************************************
            '           'strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & REMARKS	              & " </font></td>" & chr(13) & chr(10) 'REMARKS
            '           'strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " " & " </font></td>" & chr(13) & chr(10) 'MODALIDAD
            '           'strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & WEEK	                & " </font></td>" & chr(13) & chr(10) 'WEEK
            '           'strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & NUM_INVOICECUSTOM	    & " </font></td>" & chr(13) & chr(10) 'NUM. INVOICE CUSTOM
            '           'strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DATEINVOICECUSTOM     & " </font></td>" & chr(13) & chr(10) 'DATE OF INVOICE CUSTOM
            '           strHTML = strHTML & "</tr>"& chr(13) & chr(10)


       else

             strHTML = strHTML& "<tr bgcolor= '"&str_color&"' align=""center"" >"& chr(13) & chr(10)
             For intRow = 0 To contCamposplantilla
                 'strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & arrcampos(4,intRow) &" </font></td>" & chr(13) & chr(10) '
                 if arrcampos(4,intRow) = "N/A" then
                    'strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " " & " </font></td>" & chr(13) & chr(10) '
                    strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "  &" </font></strong></td>" & chr(13) & chr(10) '
                 else
                    'strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & arrcampos(4,intRow) & " </font></td>" & chr(13) & chr(10) '
                    strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & arrcampos(4,intRow)        &" </font></strong></td>" & chr(13) & chr(10) '
                 end if
             next
             strHTML = strHTML & "</tr>"& chr(13) & chr(10)



            '           strHTML = strHTML& "<tr bgcolor= '"&str_color&"' align=""center"" >"& chr(13) & chr(10)
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & REFERENCIA            & " </font></strong></td>" & chr(13) & chr(10) 'REFERENCIA
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & OTD2                  & " </font></strong></td>" & chr(13) & chr(10) 'OTD2
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_NOTIFICACION 	  & " </font></strong></td>" & chr(13) & chr(10) 'ASIGNADO ITTS /FECHA DE NOTIFICACION
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & GUIA_MASTER	          & " </font></strong></td>" & chr(13) & chr(10) 'GUIA MASTER
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & GUIA_HOUSE	          & " </font></strong></td>" & chr(13) & chr(10) 'GUIA HOUSE
            '           if P_OA = "N/A" then
            '           'if P_O <> "" and InStr(P_O,"N/A") > 0 and Len(P_O) = 3 then
            '              strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "              & " </font></strong></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
            '           else
            '              strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & P_OA               & " </font></strong></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
            '           end if
            '           'if trim(P_OA) = "N/A" then
            '           '   strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "              & " </font></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
            '           '   'response.write(P_OA)
            '           '   'Response.End
            '           'else
            '           '   strHTML = strHTML&"<td nowrap><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & P_OA               & " </font></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
            '           'end if
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & AEROPUERTO_SALIDA	    & " </font></strong></td>" & chr(13) & chr(10) 'AEROPUERTO DE SALIDA
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & CUSTOM_OF_DISPATCH    & " </font></strong></td>" & chr(13) & chr(10) '7 PORT/AIRPORT OF DEPARTURE --AEROPUERTO DE SALIDA --PORT OF LOADING
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ADUDESPACHO           & " </font></strong></td>" & chr(13) & chr(10) 'CUSTOM OF DISPATCH -- ADUANA SEGUN APENDICE 1 DEL ANEXO 22
            '
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FORWARDER_AIR_LINE    & " </font></strong></td>" & chr(13) & chr(10) 'FORWARDER Y/O AIR  LINE
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'VESSELN_GUIA	        & " </font></strong></td>" & chr(13) & chr(10) 'NOTIFICACION DE GUIA
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & IMPORT_DOCUMENT	      & " </font></strong></td>" & chr(13) & chr(10) 'IMPORT DOCUMENT
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & PROVEEDOR	            & " </font></strong></td>" & chr(13) & chr(10) 'PROVEEDOR
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & INVOICE	              & " </font></strong></td>" & chr(13) & chr(10) 'INVOICE
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DESCRIPTION_CODE	    & " </font></strong></td>" & chr(13) & chr(10) 'DESCRIPTION CODE
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & MODEL 	              & " </font></strong></td>" & chr(13) & chr(10) 'MODEL
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DESCRIPCION_COMERCIAL & " </font></strong></td>" & chr(13) & chr(10) 'DESCRIPCION COMERCIAL
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & QTY	                  & " </font></strong></td>" & chr(13) & chr(10) 'QTY
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ORIGIN_ETD	          & " </font></strong></td>" & chr(13) & chr(10) 'ORIGIN   ETD
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & RMKATDORIGIN      &" </font></strong></td>" & chr(13) & chr(10) '20 REMARKS BEFORE ARRIVAL PORT/LAX
            '
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_LAX   	          & " </font></strong></td>" & chr(13) & chr(10) 'ETA  LAX./
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATA_CUSTOM	          & " </font></strong></td>" & chr(13) & chr(10) 'ATA CUSTOM
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & RMKATAPORT        &" </font></strong></td>" & chr(13) & chr(10) '23 REMARKS  PORT/LAX
            '
            '           if SERIAL_NUMBER = "N/A" then
            '              strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ""  & " </font></strong></td>" & chr(13) & chr(10) 'SERIAL NUMBER
            '           else
            '              strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & SERIAL_NUMBER	        & " </font></strong></td>" & chr(13) & chr(10) 'SERIAL NUMBER
            '           end if
            '
            '           if CERT_NOM = "N/A" then
            '              strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ""  & " </font></strong></td>" & chr(13) & chr(10) 'CERT. NOM
            '           else
            '              strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & CERT_NOM              & " </font></strong></td>" & chr(13) & chr(10) 'CERT. NOM
            '           end if
            '
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_DE_REVALIDACION & " </font></strong></td>" & chr(13) & chr(10) 'FECHA DE  REVALIDACION
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & RESQUEST_DUTIES	      & " </font></strong></td>" & chr(13) & chr(10) 'RESQUEST DUTIES
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & AMOUNTOFDUTIES	      & " </font></strong></td>" & chr(13) & chr(10) 'AMOUNT OF DUTIES
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_DE_PREVIO	      & " </font></strong></td>" & chr(13) & chr(10) 'FECHA DE PREVIO
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_CUSTOM_CLEARANCE  & " </font></strong></td>" & chr(13) & chr(10) 'ETA DATE OF CLEARANCE
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DATE_OF_CLEARANCE	    & " </font></strong></td>" & chr(13) & chr(10) 'DATE OF CLEARANCE
            '
            '           'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & RMKDEPACHO        &" </font></strong></td>" & chr(13) & chr(10) '32 REMARKS CLEARANCE
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></td>" & chr(13) & chr(10) 'REMARKS CLEARANCE
            '
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'ATD RAIL
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'REMARKS ATD RAIL
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'ETA C/P
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'ATA C/P
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'REMARKS  C./P.
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_W_H               & " </font></strong></td>" & chr(13) & chr(10) 'ETA W/H
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATA_WH	              & " </font></strong></td>" & chr(13) & chr(10) 'ATA W-H
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & TIMEOFDELIVERY        & " </font></strong></td>" & chr(13) & chr(10) 'TIME OF DELIVERY IN SEM
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATASPL                & " </font></strong></td>" & chr(13) & chr(10) '41 TRUCK  DEPARTURE FROM W/H  --SALIDA DE ALMACEN
            '
            '           '**********************************************************************
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & STATUS            &" </font></strong></td>" & chr(13) & chr(10) '42 STATUS
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & LASTRMK           &" </font></strong></td>" & chr(13) & chr(10) '43 REMARKS (ULTIMO)
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & KPISTATUS         &" </font></strong></td>" & chr(13) & chr(10) '44 KPI STATUS
            '           strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"' ><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "               &" </font></strong></td>" & chr(13) & chr(10) 'MODALIDAD
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & WEEK              &" </font></strong></td>" & chr(13) & chr(10) 'WEEK
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & NUM_INVOICECUSTOM &" </font></strong></td>" & chr(13) & chr(10) 'NUM. INVOICE CUSTOM
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & DATEINVOICECUSTOM &" </font></strong></td>" & chr(13) & chr(10) 'DATE OF INVOICE CUSTOM
            '           strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"'  size=""1"" face=""Arial, Helvetica, sans-serif""> " & REMARKS   &" </font></strong></td>" & chr(13) & chr(10) 'HISTORIAL
            '           '**********************************************************************

       end if



             'strHTML = strHTML& "<tr bgcolor= '"&str_color&"' align=""center"" >"& chr(13) & chr(10)
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & REFERENCIA            & " </font></strong></td>" & chr(13) & chr(10) 'REFERENCIA
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & OTD2                  & " </font></strong></td>" & chr(13) & chr(10) 'OTD2
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_NOTIFICACION 	  & " </font></strong></td>" & chr(13) & chr(10) 'ASIGNADO ITTS /FECHA DE NOTIFICACION
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & GUIA_MASTER	          & " </font></strong></td>" & chr(13) & chr(10) 'GUIA MASTER
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & GUIA_HOUSE	          & " </font></strong></td>" & chr(13) & chr(10) 'GUIA HOUSE
             'if P_OA = "N/A" then
             ''if P_O <> "" and InStr(P_O,"N/A") > 0 and Len(P_O) = 3 then
             '   strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "              & " </font></strong></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
             'else
             '   strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & P_OA               & " </font></strong></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
             'end if
             ''if trim(P_OA) = "N/A" then
             ''   strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " "              & " </font></strong></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
             ''   'response.write(P_OA)
             ''   'Response.End
             ''else
             ''   strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & P_OA               & " </font></strong></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
             ''end if
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & AEROPUERTO_SALIDA	    & " </font></strong></td>" & chr(13) & chr(10) 'AEROPUERTO DE SALIDA
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & CUSTOM_OF_DISPATCH    & " </font></strong></td>" & chr(13) & chr(10) 'CUSTOM OF DISPATCH -- ADUANA SEGUN APENDICE 1 DEL ANEXO 22
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FORWARDER_AIR_LINE    & " </font></strong></td>" & chr(13) & chr(10) 'FORWARDER Y/O AIR  LINE
             'strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'VESSELN_GUIA	        & " </font></strong></td>" & chr(13) & chr(10) 'NOTIFICACION DE GUIA
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & IMPORT_DOCUMENT	      & " </font></strong></td>" & chr(13) & chr(10) 'IMPORT DOCUMENT
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & PROVEEDOR	            & " </font></strong></td>" & chr(13) & chr(10) 'PROVEEDOR
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & INVOICE	              & " </font></strong></td>" & chr(13) & chr(10) 'INVOICE
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DESCRIPTION_CODE	    & " </font></strong></td>" & chr(13) & chr(10) 'DESCRIPTION CODE
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & MODEL 	              & " </font></strong></td>" & chr(13) & chr(10) 'MODEL
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DESCRIPCION_COMERCIAL & " </font></strong></td>" & chr(13) & chr(10) 'DESCRIPCION COMERCIAL
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & QTY	                  & " </font></strong></td>" & chr(13) & chr(10) 'QTY
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ORIGIN_ETD	          & " </font></strong></td>" & chr(13) & chr(10) 'ORIGIN   ETD
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_LAX   	          & " </font></strong></td>" & chr(13) & chr(10) 'ETA  LAX./
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATA_CUSTOM	          & " </font></strong></td>" & chr(13) & chr(10) 'ATA CUSTOM
             'if SERIAL_NUMBER = "N/A" then
             '  strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ""  & " </font></strong></td>" & chr(13) & chr(10) 'SERIAL NUMBER
             'else
             '  strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & SERIAL_NUMBER	        & " </font></strong></td>" & chr(13) & chr(10) 'SERIAL NUMBER
             'end if
             'if CERT_NOM = "N/A" then
             '  strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ""   & " </font></strong></td>" & chr(13) & chr(10) 'CERT. NOM
             'else
             '  strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & CERT_NOM              & " </font></strong></td>" & chr(13) & chr(10) 'CERT. NOM
             'end if
             ''strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & SERIAL_NUMBER	        & " </font></strong></td>" & chr(13) & chr(10) 'SERIAL NUMBER
             ''strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & CERT_NOM              & " </font></strong></td>" & chr(13) & chr(10) 'CERT. NOM
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_DE_REVALIDACION & " </font></strong></td>" & chr(13) & chr(10) 'FECHA DE  REVALIDACION
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & RESQUEST_DUTIES	      & " </font></strong></td>" & chr(13) & chr(10) 'RESQUEST DUTIES
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & AMOUNTOFDUTIES	      & " </font></strong></td>" & chr(13) & chr(10) 'AMOUNT OF DUTIES
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_DE_PREVIO	      & " </font></strong></td>" & chr(13) & chr(10) 'FECHA DE PREVIO
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_CUSTOM_CLEARANCE  & " </font></strong></td>" & chr(13) & chr(10) 'ETA DATE OF CLEARANCE
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DATE_OF_CLEARANCE	    & " </font></strong></td>" & chr(13) & chr(10) 'DATE OF CLEARANCE
             'strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'ATD RAIL
             'strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'ETA C/P
             'strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & "             " & " </font></strong></td>" & chr(13) & chr(10) 'ATA C/P
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_W_H               & " </font></strong></td>" & chr(13) & chr(10) 'ETA W/H
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATA_WH	              & " </font></strong></td>" & chr(13) & chr(10) 'ATA W-H
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & TIMEOFDELIVERY        & " </font></strong></td>" & chr(13) & chr(10) 'TIME OF DELIVERY IN SEM
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & REMARKS	              & " </font></strong></td>" & chr(13) & chr(10) 'REMARKS
             'strHTML = strHTML&"<td nowrap bgcolor='"&strColorNA&"'><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & " " & " </font></strong></td>" & chr(13) & chr(10) 'MODALIDAD
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & WEEK	                & " </font></strong></td>" & chr(13) & chr(10) 'WEEK
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & NUM_INVOICECUSTOM	    & " </font></strong></td>" & chr(13) & chr(10) 'NUM. INVOICE CUSTOM
             'strHTML = strHTML&"<td nowrap><strong><font color='"&str_fcolor&"' size=""1"" face=""Arial, Helvetica, sans-serif""> " & DATEINVOICECUSTOM     & " </font></strong></td>" & chr(13) & chr(10) 'DATE OF INVOICE CUSTOM
             'strHTML = strHTML & "</tr>"& chr(13) & chr(10)

             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & REFERENCIA            & " </font></strong></td>" & chr(13) & chr(10) 'REFERENCIA
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & OTD2                  & " </font></strong></td>" & chr(13) & chr(10) 'OTD2
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_NOTIFICACION 	 & " </font></strong></td>" & chr(13) & chr(10) 'ASIGNADO ITTS /FECHA DE NOTIFICACION
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & GUIA_MASTER	         & " </font></strong></td>" & chr(13) & chr(10) 'GUIA MASTER
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & GUIA_HOUSE	           & " </font></strong></td>" & chr(13) & chr(10) 'GUIA HOUSE
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & P_OA                  & " </font></strong></td>" & chr(13) & chr(10) 'P/O --ORDEN DE COMPRA
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & AEROPUERTO_SALIDA	   & " </font></strong></td>" & chr(13) & chr(10) 'AEROPUERTO DE SALIDA
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & CUSTOM_OF_DISPATCH    & " </font></strong></td>" & chr(13) & chr(10) 'CUSTOM OF DISPATCH -- ADUANA SEGUN APENDICE 1 DEL ANEXO 22
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & FORWARDER_AIR_LINE    & " </font></strong></td>" & chr(13) & chr(10) 'FORWARDER Y/O AIR  LINE
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & VESSEL                & " </font></strong></td>" & chr(13) & chr(10) 'VESSEL
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & IMPORT_DOCUMENT	     & " </font></strong></td>" & chr(13) & chr(10) 'IMPORT DOCUMENT
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & PROVEEDOR	           & " </font></strong></td>" & chr(13) & chr(10) 'PROVEEDOR
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & INVOICE	             & " </font></strong></td>" & chr(13) & chr(10) 'INVOICE
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & DESCRIPTION_CODE	     & " </font></strong></td>" & chr(13) & chr(10) 'DESCRIPTION CODE
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & MODEL 	               & " </font></strong></td>" & chr(13) & chr(10) 'MODEL
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & DESCRIPCION_COMERCIAL & " </font></strong></td>" & chr(13) & chr(10) 'DESCRIPCION COMERCIAL
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & QTY	                 & " </font></strong></td>" & chr(13) & chr(10) 'QTY
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & ORIGIN_ETD	           & " </font></strong></td>" & chr(13) & chr(10) 'ORIGIN   ETD
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_LAX   	           & " </font></strong></td>" & chr(13) & chr(10) 'ETA  LAX./
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATA_CUSTOM	           & " </font></strong></td>" & chr(13) & chr(10) 'ATA CUSTOM
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & SERIAL_NUMBER	       & " </font></strong></td>" & chr(13) & chr(10) 'SERIAL NUMBER
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & CERT_NOM              & " </font></strong></td>" & chr(13) & chr(10) 'CERT. NOM
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_DE_REVALIDACION & " </font></strong></td>" & chr(13) & chr(10) 'FECHA DE  REVALIDACION
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & RESQUEST_DUTIES	     & " </font></strong></td>" & chr(13) & chr(10) 'RESQUEST DUTIES
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & AMOUNTOFDUTIES	       & " </font></strong></td>" & chr(13) & chr(10) 'AMOUNT OF DUTIES
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & FECHA_DE_PREVIO	     & " </font></strong></td>" & chr(13) & chr(10) 'FECHA DE PREVIO
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_CUSTOM_CLEARANCE  & " </font></strong></td>" & chr(13) & chr(10) 'ETA DATE OF CLEARANCE
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & DATE_OF_CLEARANCE	   & " </font></strong></td>" & chr(13) & chr(10) 'DATE OF CLEARANCE
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & "                "    & " </font></strong></td>" & chr(13) & chr(10) 'ATD RAIL
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_C_P               & " </font></strong></td>" & chr(13) & chr(10) 'ETA C/P
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATA_CP                & " </font></strong></td>" & chr(13) & chr(10) 'ATA C/P
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & ETA_W_H               & " </font></strong></td>" & chr(13) & chr(10) 'ETA W/H
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & ATA_WH	               & " </font></strong></td>" & chr(13) & chr(10) 'ATA W-H
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & TIMEOFDELIVERY        & " </font></strong></td>" & chr(13) & chr(10) 'TIME OF DELIVERY IN SEM
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & REMARKS	             & " </font></strong></td>" & chr(13) & chr(10) 'REMARKS
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & MODALIDAD             & " </font></strong></td>" & chr(13) & chr(10) 'MODALIDAD
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & WEEK	                 & " </font></strong></td>" & chr(13) & chr(10) 'WEEK
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & NUM_INVOICECUSTOM	   & " </font></strong></td>" & chr(13) & chr(10) 'NUM. INVOICE CUSTOM
             'strHTML = strHTML&"<td nowrap><font color=""#000000"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & DATEINVOICECUSTOM     & " </font></strong></td>" & chr(13) & chr(10) 'DATE OF INVOICE CUSTOM

    end function

     '-------------------------------------------------------------------------------------------------------------------------------
%>

<%
    'TipoFiltro

     tempstrOficina = adu_ofi( Session("GAduana") )
     IF NOT enproceso(tempstrOficina) THEN



    MM_EXTRANET_STRING = ODBC_POR_ADUANA(Session("GAduana"))
    MM_EXTRANET_STRING_STATUS = ODBC_POR_ADUANA(Session("GAduana")&"_STATUS")

    'Dim arrRefEtapas()



    Response.Buffer = TRUE
    Response.Addheader "Content-Disposition", "attachment;filename=TRACKING_AEREO.xls"
    Response.ContentType = "application/vnd.ms-excel"
    Server.ScriptTimeOut=100000

    strUsuario     = request.Form("user")
    strTipoUsuario = request.Form("TipoUser")
    strPermisos    = Request.Form("Permisos")
    permi          = PermisoClientes(Session("GAduana"),strPermisos,"cvecli01 ")
    permi2         = PermisoClientesTabla("B",Session("GAduana") ,strPermisos,"clie31")

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
    'strFiltroCliente = ""
    'strFiltroCliente = request.Form("txtCliente")
    if not strFiltroCliente  = "" and not strFiltroCliente  = "Todos" then
       blnAplicaFiltro = true
    end if
    if blnAplicaFiltro then
       permi = " AND cvecli01 =" & strFiltroCliente
    end if
    if strTipoUsuario = MM_Cod_Admon and not blnAplicaFiltro then
       permi = ""
    end if

    strDateIni = ""
    strDateFin = ""
    strTipoPedimento = ""
    strCodError      = "0"


    '*******************************************************
    strTipoFiltro         = trim(request.Form("TipoFiltro"))
    strDateIni            = trim(request.Form("txtDateIni"))
    strDateFin            = trim(request.Form("txtDateFin"))
    strTipoFecha          = trim(request.Form("txtTipoFecha"))
    strLinNav             = trim(request.Form("txtLinNav"))
    strModalidad          = trim(request.Form("txtMod"))
    strProv               = trim(request.Form("txtProv"))
    strfiltrosrestantes   = trim(request.Form("txtfiltrosrestantes"))
    strTipoOtrosFiltros   = trim(request.Form("txttipoOtrosFiltros"))

    'strTipoConte = trim(request.Form("txttipoConte"))


    '*******************************************************

    if not isdate(strDateIni) then
      strCodError = "5"
    end if
    if not isdate(strDateFin) then
      strCodError = "6"
    end if
    if strDateIni="" or strDateFin="" then
      strCodError = "1"
    end if

    'Response.Write(strTipoFiltro)
	

    strHTML = ""


    if strCodError = "0" then

    tmpTipo = ""
    strSQL  = ""

    'BotonxLinNaviera  - Filtro de linea naviera
    'BotonxModalidad   - Filtro de modalidad (Tipo de transporte)
    'BotonxProveedor   - Filtro de Proveedor
    'BotonOtrosFiltros - Otros Flitros de captura libre

         if strTipoFiltro  = "BotonxLinNaviera" then 'Filtro de linea naviera
              'filtro para la linea aerea
              strCadFiltroLinNav = ""
              if ltrim(strLinNav) <> "Todos" then ' Selecciono una linea  naviera
                  strCadFiltroLinNav = " and C01REFER.cvela01 = " & strLinNav
              end if

              'filtro para el proveedor
              strCadFiltroProv = ""
              if ltrim(strProv) <> "Todos" then ' Selecciono un proveedor
                  strCodProvt = ""
                  Set RsProv = Server.CreateObject("ADODB.Recordset")
                  RsProv.ActiveConnection = MM_EXTRANET_STRING
                  RsProv.Source = " SELECT CVEPRO22 FROM SSPROV22 WHERE NPSCLI22='"&strProv&"'"

                  'response.write(RsProv.Source)
                  'Response.End

                  RsProv.CursorType = 0
                  RsProv.CursorLocation = 2
                  RsProv.LockType = 1
                  RsProv.Open()
                  if not RsProv.eof then
                        strCodProvt = RsProv.Fields.Item("CVEPRO22").Value
                  end if
                  RsProv.close
                  set RsProv = Nothing
                  if strCodProvt <> "" then
                     strCadFiltroProv = " and SSDAGI01.cvepro01 = " & strCodProvt
                  end if
              end if


              if ltrim(strTipoFecha)  = "DAITTS" OR ltrim(strTipoFecha) = "DFPAG" then ' Selecciona la fecha de ITTS
                    ' ESTE QUERY TRAE LOS REGISTROS A NIVEL PEDIMENTO
                    IF ltrim(strTipoFecha) = "DAITTS" then


                       strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                                "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                                "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                                "        SSDAGI01.PATENT01,                     " & _
                                "        CONCAT(SSDAGI01.PATENT01, CONCAT( '',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                                "        C01REFER.feta01    AS ETA_PORT,        " & _
                                "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                                "        SSDAGI01.fecent01,                     " & _
                                "        C01REFER.frev01    AS REVALIDACION,    " & _
                                "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                "        C01REFER.fpre01    AS PREVIO,          " & _
                                "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                                "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                                "        firmae01,                              " & _
                                "        frec01 as FecITTS,                     " & _
                                "        cvrexp01 as FORWARDER,                 " & _
                                "        cvela01  as  AIRLINE,                  " & _
                                "        feorig01,                              " & _
                                "        etalax01,                              " & _
                                "        cbuq01,                                " & _
                                "        CVEPED01,                              " & _
                                "        cveptoemb,                             " & _
                                "        ADUDES01                               " & _
                                " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01        " & _
                                " WHERE adusec01=470 and modo01 = 'T' AND                                    " & _
                                "        firmae01   <> ''        AND                                    " & _
                                "        frec01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                                "        frec01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                                "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv & _
                                " UNION ALL                                     " & _
                                " SELECT SSDAGE01.REFCIA01  AS REFERENCIA,      " & _
                                "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                "        SSDAGE01.adusec01  AS PORT_DISCHARGE,  " & _
                                "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                "        SSDAGE01.REGBAR01  AS VESSEL,          " & _
                                "        SSDAGE01.PATENT01,                     " & _
                                "        CONCAT(SSDAGE01.PATENT01, CONCAT( '',SSDAGE01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                "        SSDAGE01.CVEPRO01  AS PROVEEDOR,       " & _
                                "        C01REFER.feta01    AS ETA_PORT,        " & _
                                "        SSDAGE01.fecpre01  AS ETA_PORT2,       " & _
                                "        SSDAGE01.fecpre01  AS fecent01,                     " & _
                                "        C01REFER.frev01    AS REVALIDACION,    " & _
                                "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                "        C01REFER.fpre01    AS PREVIO,          " & _
                                "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                "        SSDAGE01.cvemts01  AS MODALIDAD,       " & _
                                "        SSDAGE01.desf0101  AS FACTURAS,        " & _
                                "        firmae01,                              " & _
                                "        frec01 as FecITTS,                     " & _
                                "        cvrexp01 as FORWARDER,                 " & _
                                "        cvela01  as  AIRLINE,                  " & _
                                "        feorig01,                              " & _
                                "        etalax01,                              " & _
                                "        cbuq01,                                " & _
                                "        CVEPED01,                              " & _
                                "        cveptoemb,                             " & _
                                "        ADUDES01                               " & _
                                " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01        " & _
                                " WHERE adusec01=470 and modo01 = 'T' AND                                    " & _
                                "        firmae01   <> ''        AND                                    " & _
                                "        frec01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                                "        frec01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                                "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv & _
                                " ORDER BY ETA_PORT2, ETA_PORT  "


                     'strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                     '         "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                     '         "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                     '         "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                     '         "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                     '         "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                     '         "        SSDAGI01.PATENT01,                     " & _
                     '         "        CONCAT(SSDAGI01.PATENT01, CONCAT( '-',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                     '         "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                     '         "        C01REFER.feta01    AS ETA_PORT,        " & _
                     '         "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                     '         "        SSDAGI01.fecent01,                     " & _
                     '         "        C01REFER.frev01    AS REVALIDACION,    " & _
                     '         "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                     '         "        C01REFER.fpre01    AS PREVIO,          " & _
                     '         "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                     '         "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                     '         "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                     '         "        firmae01,                              " & _
                     '         "        frec01 as FecITTS,                     " & _
                     '         "        cvrexp01 as FORWARDER,                 " & _
                     '         "        cvela01  as  AIRLINE,                  " & _
                     '         "        feorig01,                              " & _
                     '         "        etalax01,                              " & _
                     '         "        cbuq01                                 " & _
                     '         " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01        " & _
                     '         "        WHERE ( cvep01 <> 'R1') AND                                    " & _
                     '         "        firmae01   <> ''        AND                                    " & _
                     '         "        frec01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                     '         "        frec01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                     '         "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv & _
                     '         " UNION ALL  " & _
                     '         " SELECT SSDAGE01.REFCIA01  AS REFERENCIA,      " & _
                     '         "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                     '         "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                     '         "        SSDAGE01.adusec01  AS PORT_DISCHARGE,  " & _
                     '         "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                     '         "        SSDAGE01.REGBAR01  AS VESSEL,          " & _
                     '         "        SSDAGE01.PATENT01,                     " & _
                     '         "        CONCAT(SSDAGE01.PATENT01, CONCAT( '-',SSDAGE01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                     '         "        SSDAGE01.CVEPRO01  AS PROVEEDOR,       " & _
                     '         "        C01REFER.feta01    AS ETA_PORT,        " & _
                     '         "        SSDAGE01.fecpre01  AS ETA_PORT2,       " & _
                     '         "        SSDAGE01.fecpre01 as fecent01,         " & _
                     '         "        C01REFER.frev01    AS REVALIDACION,    " & _
                     '         "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                     '         "        C01REFER.fpre01    AS PREVIO,          " & _
                     '         "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                     '         "        SSDAGE01.cvemts01  AS MODALIDAD,       " & _
                     '         "        SSDAGE01.desf0101  AS FACTURAS,        " & _
                     '         "        firmae01,                              " & _
                     '         "        frec01 as FecITTS,                     " & _
                     '         "        cvrexp01 as FORWARDER,                 " & _
                     '         "        cvela01  as  AIRLINE,                  " & _
                     '         "        feorig01,                              " & _
                     '         "        feorig01,                              " & _
                     '         "        etalax01,                              " & _
                     '         "        cbuq01                                 " & _
                     '         " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01        " & _
                     '         "        WHERE ( cvep01 <> 'R1') AND                                    " & _
                     '         "        firmae01   <> ''        AND                                    " & _
                     '         "        frec01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                     '         "        frec01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                     '         "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv

                    ELSE

                       strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                                "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                                "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                                "        SSDAGI01.PATENT01,                     " & _
                                "        CONCAT(SSDAGI01.PATENT01, CONCAT( '',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                                "        C01REFER.feta01    AS ETA_PORT,        " & _
                                "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                                "        SSDAGI01.fecent01,                     " & _
                                "        C01REFER.frev01    AS REVALIDACION,    " & _
                                "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                "        C01REFER.fpre01    AS PREVIO,          " & _
                                "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                                "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                                "        firmae01,                              " & _
                                "        frec01 as FecITTS,                     " & _
                                "        cvrexp01 as FORWARDER,                 " & _
                                "        cvela01  as  AIRLINE,                  " & _
                                "        feorig01,                              " & _
                                "        etalax01,                              " & _
                                "        cbuq01,                                " & _
                                "        CVEPED01,                              " & _
                                "        cveptoemb,                             " & _
                                "        ADUDES01                               " & _
                                " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01        " & _
                                " WHERE adusec01=470 and modo01 = 'T' AND                                    " & _
                                "        firmae01   <> ''        AND                                    " & _
                                "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                                "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                                "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv & _
                                " UNION ALL                                     " & _
                                " SELECT SSDAGE01.REFCIA01  AS REFERENCIA,      " & _
                                "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                "        SSDAGE01.adusec01  AS PORT_DISCHARGE,  " & _
                                "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                "        SSDAGE01.REGBAR01  AS VESSEL,          " & _
                                "        SSDAGE01.PATENT01,                     " & _
                                "        CONCAT(SSDAGE01.PATENT01, CONCAT( '',SSDAGE01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                "        SSDAGE01.CVEPRO01  AS PROVEEDOR,       " & _
                                "        C01REFER.feta01    AS ETA_PORT,        " & _
                                "        SSDAGE01.fecpre01  AS ETA_PORT2,       " & _
                                "        SSDAGE01.fecpre01  AS fecent01,                     " & _
                                "        C01REFER.frev01    AS REVALIDACION,    " & _
                                "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                "        C01REFER.fpre01    AS PREVIO,          " & _
                                "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                "        SSDAGE01.cvemts01  AS MODALIDAD,       " & _
                                "        SSDAGE01.desf0101  AS FACTURAS,        " & _
                                "        firmae01,                              " & _
                                "        frec01 as FecITTS,                     " & _
                                "        cvrexp01 as FORWARDER,                 " & _
                                "        cvela01  as  AIRLINE,                  " & _
                                "        feorig01,                              " & _
                                "        etalax01,                              " & _
                                "        cbuq01,                                " & _
                                "        CVEPED01,                              " & _
                                "        cveptoemb,                             " & _
                                "        ADUDES01                               " & _
                                " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01        " & _
                                " WHERE adusec01=470 and modo01 = 'T' AND                                    " & _
                                "        firmae01   <> ''        AND                                    " & _
                                "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                                "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                                "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv & _
                                " ORDER BY ETA_PORT2, ETA_PORT  "

                    END IF
                     'strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                     '         "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                     '         "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                     '         "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                     '         "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                     '         "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                     '         "        SSDAGI01.PATENT01,                     " & _
                     '         "        CONCAT(SSDAGI01.PATENT01, CONCAT( '-',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                     '         "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                     '         "        C01REFER.feta01    AS ETA_PORT,        " & _
                     '         "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                     '         "        SSDAGI01.fecent01,                     " & _
                     '         "        C01REFER.frev01    AS REVALIDACION,    " & _
                     '         "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                     '         "        C01REFER.fpre01    AS PREVIO,          " & _
                     '         "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                     '         "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                     '         "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                     '         "        firmae01,                              " & _
                     '         "        frec01 as FecITTS,                     " & _
                     '         "        cvrexp01 as FORWARDER,                 " & _
                     '         "        cvela01  as  AIRLINE,                  " & _
                     '         "        feorig01,                              " & _
                     '         "        etalax01                               " & _
                     '         " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01        " & _
                     '         "        WHERE ( cvep01 <> 'R1') AND                                    " & _
                     '         "        firmae01   <> ''        AND                                    " & _
                     '         "        frec01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                     '         "        frec01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                     '         "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv


              else
                    if ltrim(strTipoFecha)  = "DCusCl" then ' selecciono Date of custom clearance (fecha de despacho)
                       ' ESTE QUERY TRAE LOS REGISTROS A NIVEL PEDIMENTO
                       strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                                "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                                "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                                "        SSDAGI01.PATENT01,                     " & _
                                "        CONCAT(SSDAGI01.PATENT01, CONCAT( '',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                                "        C01REFER.feta01    AS ETA_PORT,        " & _
                                "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                                "        SSDAGI01.fecent01,                     " & _
                                "        C01REFER.frev01    AS REVALIDACION,    " & _
                                "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                "        C01REFER.fpre01    AS PREVIO,          " & _
                                "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                                "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                                "        firmae01,                              " & _
                                "        frec01 as FecITTS,                     " & _
                                "        cvrexp01 as FORWARDER,                 " & _
                                "        cvela01  as  AIRLINE,                  " & _
                                "        feorig01,                              " & _
                                "        etalax01,                              " & _
                                "        cbuq01,                                " & _
                                "        CVEPED01,                              " & _
                                "        cveptoemb,                             " & _
                                "        ADUDES01                               " & _
                                " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01        " & _
                                " WHERE adusec01=470 and modo01 = 'T' AND                                    " & _
                                "        firmae01   <> ''        AND                                    " & _
                                "        fdsp01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                                "        fdsp01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                                "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv & _
                                " UNION ALL                                     " & _
                                " SELECT SSDAGE01.REFCIA01  AS REFERENCIA,      " & _
                                "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                "        SSDAGE01.adusec01  AS PORT_DISCHARGE,  " & _
                                "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                "        SSDAGE01.REGBAR01  AS VESSEL,          " & _
                                "        SSDAGE01.PATENT01,                     " & _
                                "        CONCAT(SSDAGE01.PATENT01, CONCAT( '',SSDAGE01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                "        SSDAGE01.CVEPRO01  AS PROVEEDOR,       " & _
                                "        C01REFER.feta01    AS ETA_PORT,        " & _
                                "        SSDAGE01.fecpre01  AS ETA_PORT2,       " & _
                                "        SSDAGE01.fecpre01  AS fecent01,                     " & _
                                "        C01REFER.frev01    AS REVALIDACION,    " & _
                                "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                "        C01REFER.fpre01    AS PREVIO,          " & _
                                "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                "        SSDAGE01.cvemts01  AS MODALIDAD,       " & _
                                "        SSDAGE01.desf0101  AS FACTURAS,        " & _
                                "        firmae01,                              " & _
                                "        frec01 as FecITTS,                     " & _
                                "        cvrexp01 as FORWARDER,                 " & _
                                "        cvela01  as  AIRLINE,                  " & _
                                "        feorig01,                              " & _
                                "        etalax01,                              " & _
                                "        cbuq01,                                " & _
                                "        CVEPED01,                              " & _
                                "        cveptoemb,                             " & _
                                "        ADUDES01                               " & _
                                " FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01        " & _
                                " WHERE adusec01=470 and modo01 = 'T' AND                                    " & _
                                "        firmae01   <> ''        AND                                    " & _
                                "        fdsp01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                                "        fdsp01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                                "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv & _
                                " ORDER BY ETA_PORT2, ETA_PORT  "



                                'strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                                '"        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                '"        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                '"        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                                '"        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                '"        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                                '"        SSDAGI01.PATENT01,                     " & _
                                '"        CONCAT(SSDAGI01.PATENT01, CONCAT( '-',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                '"        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                                '"        C01REFER.feta01    AS ETA_PORT,        " & _
                                '"        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                                '"        SSDAGI01.fecent01,                     " & _
                                '"        C01REFER.frev01    AS REVALIDACION,    " & _
                                '"        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                '"        C01REFER.fpre01    AS PREVIO,          " & _
                                '"        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                '"        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                                '"        SSDAGI01.desf0101  AS FACTURAS,        " & _
                                '"        firmae01,                              " & _
                                '"        frec01 as FecITTS,                     " & _
                                '"        cvrexp01 as FORWARDER,                 " & _
                                '"        cvela01  as  AIRLINE,                  " & _
                                '"        feorig01,                              " & _
                                '"        etalax01                               " & _
                                '" FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01        " & _
                                '"        WHERE ( cvep01 <> 'R1') AND                                    " & _
                                '"        firmae01   <> ''        AND                                    " & _
                                '"        fdsp01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                                '"        fdsp01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                                '"        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv


                       'strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                       '         "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                       '         "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                       '         "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                       '         "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                       '         "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                       '         "        SSDAGI01.PATENT01,                     " & _
                       '         "        CONCAT(SSDAGI01.PATENT01, CONCAT( '-',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                       '         "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                       '         "        C01REFER.feta01    AS ETA_PORT,        " & _
                       '         "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                       '         "        SSDAGI01.fecent01,                     " & _
                       '         "        C01REFER.frev01    AS REVALIDACION,    " & _
                       '         "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                       '         "        C01REFER.fpre01    AS PREVIO,          " & _
                       '         "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                       '         "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                       '         "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                       '         "        firmae01,                              " & _
                       '         "        frec01 as FecITTS,                     " & _
                       '         "        cvrexp01 as FORWARDER,                 " & _
                       '         "        cvela01  as  AIRLINE,                  " & _
                       '         "        feorig01,                              " & _
                       '         "        etalax01                               " & _
                       '         " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01        " & _
                       '         "        WHERE ( cvep01 <> 'R1') AND                                    " & _
                       '         "        firmae01   <> ''        AND                                    " & _
                       '         "        fecpag01 >= '"&FormatoFechaInv(strDateIni)&"' AND              " & _
                       '         "        fecpag01 <= '"&FormatoFechaInv(strDateFin)&"' AND              " & _
                       '         "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv
                    else
                        if ltrim(strTipoFecha)  = "ATAAlm" then ' selecciono ATA W/H (fecha real de llegada a planta)
                            strAduana =Session("GAduana")
                            StrPreAdu = ""
                            if strAduana="VER" then
                               StrPreAdu = "RKU"
                            else
                               if strAduana="LZR" then
                                  StrPreAdu = "LZR"
                               else
                                  if strAduana="TAM" then
                                     StrPreAdu = "CEG"
                                  else
                                     if strAduana="MEX" then
                                        StrPreAdu = "DAI"
                                     else
                                        if strAduana="MAN" then
                                           StrPreAdu = "SAP"
                                        end if
                                     end if
                                  end if
                               end if
                            end if
                            strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                                     "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                     "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                     "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                                     "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                     "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                                     "        SSDAGI01.PATENT01,                     " & _
                                     "        CONCAT(SSDAGI01.PATENT01, CONCAT( '',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                     "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                                     "        C01REFER.feta01    AS ETA_PORT,        " & _
                                     "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                                     "        SSDAGI01.fecent01,                     " & _
                                     "        C01REFER.frev01    AS REVALIDACION,    " & _
                                     "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                     "        C01REFER.fpre01    AS PREVIO,          " & _
                                     "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                     "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                                     "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                                     "        firmae01,                              " & _
                                     "        frec01 as FecITTS,                     " & _
                                     "        cvrexp01 as FORWARDER,                 " & _
                                     "        cvela01  as  AIRLINE,                  " & _
                                     "        feorig01,                              " & _
                                     "        etalax01,                              " & _
                                     "        cbuq01,                                " & _
                                     "        CVEPED01,                              " & _
                                     "        cveptoemb,                             " & _
                                     "        ADUDES01                               " & _
                                     "  FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01   " & _
                                     "      INNER JOIN D01CONTE ON D01CONTE.REFE01=REFCIA01              " & _
                                     "      INNER JOIN "&StrPreAdu&"_STATUS.ETXCOI                                 " & _
                                     "            ON REFCIA01="&StrPreAdu&"_STATUS.ETXCOI.C_REFERENCIA and         " & _
                                     "      marc01="&StrPreAdu&"_STATUS.ETXCOI.C_Conte                             " & _
                                     "      INNER JOIN "&StrPreAdu&"_STATUS.ETAPS                                    " & _
                                     "            ON  "&StrPreAdu&"_STATUS.ETXCOI.N_ETAPA = "&StrPreAdu&"_STATUS.ETAPS.N_ETAPA " & _
                                     "  WHERE adusec01=470 and modo01 = 'T' AND                                      " & _
                                     "        firmae01   <> ''        AND                                " & _
                                     StrPreAdu & "_STATUS.ETXCOI.f_fecha  >= '"&FormatoFechaInv(strDateIni)&"' AND        " & _
                                     StrPreAdu & "_STATUS.ETXCOI.f_fecha <=  '"&FormatoFechaInv(strDateFin)&"' AND        " & _
                                     StrPreAdu&"_STATUS.ETAPS.D_ABREV ='LLP' AND                         " & _
                                     "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv     & _
                                     "  GROUP BY REFCIA01 " & _
                                     " ORDER BY ETA_PORT2, ETA_PORT  "

                                     '" UNION ALL                                       " & _
                                     '"   SELECT SSDAGE01.REFCIA01  AS REFERENCIA,      " & _
                                     '"  C01REFER.PTOEMB01  AS PORT_LOADING,            " & _
                                     '"  C01REFER.PAISEM01  AS VESSEL_LOADING,          " & _
                                     '"  SSDAGE01.adusec01  AS PORT_DISCHARGE,          " & _
                                     '"  C01REFER.Naim01    AS SHIPPING_LINE,           " & _
                                     '"  SSDAGE01.REGBAR01  AS VESSEL,                  " & _
                                     '"  SSDAGE01.PATENT01,                             " & _
                                     '"  CONCAT(SSDAGE01.PATENT01, CONCAT( '-',SSDAGE01.NUMPED01 ) ) AS IMPORT_DOCUMENT,  " & _
                                     '"  SSDAGE01.CVEPRO01  AS PROVEEDOR,               " & _
                                     '"  C01REFER.feta01    AS ETA_PORT,                " & _
                                     '"  SSDAGE01.fecpre01  AS ETA_PORT2,               " & _
                                     '"  SSDAGE01.fecpre01  AS fecent01,                " & _
                                     '"  C01REFER.frev01    AS REVALIDACION,            " & _
                                     '"  C01REFER.fcot01    AS RESQUEST_DUTIES,         " & _
                                     '"  C01REFER.fpre01    AS PREVIO,                  " & _
                                     '"  C01REFER.fdsp01    AS DATE_CUSTOM,             " & _
                                     '"  SSDAGE01.cvemts01  AS MODALIDAD,               " & _
                                     '"  SSDAGE01.desf0101  AS FACTURAS,                " & _
                                     '"  firmae01,                                      " & _
                                     '"  frec01 as FecITTS,                             " & _
                                     '"  cvrexp01 as FORWARDER,                         " & _
                                     '"  cvela01  as  AIRLINE,                          " & _
                                     '"  feorig01,                                      " & _
                                     '"  etalax01                                       " & _
                                     '"  FROM C01REFER INNER JOIN SSDAGE01 ON REFCIA01= C01REFER.REFE01  " & _
                                     '"  INNER JOIN D01CONTE ON D01CONTE.REFE01=REFCIA01                 " & _
                                     '"  INNER JOIN DAI_STATUS.ETXCOI                                    " & _
                                     '"  ON REFCIA01=DAI_STATUS.ETXCOI.C_REFERENCIA and
                                     '"  marc01=DAI_STATUS.ETXCOI.C_Conte
                                     '"  INNER JOIN DAI_STATUS.ETAPS
                                     '"  ON  DAI_STATUS.ETXCOI.N_ETAPA = DAI_STATUS.ETAPS.N_ETAPA
                                     '"  WHERE ( cvep01 <> 'R1') AND
                                     '"  firmae01   <> ''        AND
                                     '"  DAI_STATUS.ETXCOI.f_fecha  >= '2007/1/1' AND
                                     '"  DAI_STATUS.ETXCOI.f_fecha <=  '2007/6/6' AND
                                     '"  DAI_STATUS.ETAPS.D_ABREV ='LLP' AND
                                     '"  C01REFER.REFE01 <> ''   AND
                                     '"  cvecli01 =1903
                                     '"  GROUP BY REFCIA01



                        end if
                    end if
              end if
         else
          if strTipoFiltro  = "BotonOtrosFiltros" then  'Otros Flitros de captura libre
                if strTipoOtrosFiltros = "Modelo" then ' filtro por modelo
                    'strfiltrosrestantes
                    strSQL =  " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                              "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                              "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                              "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                              "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                              "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                              "        SSDAGI01.PATENT01,                     " & _
                              "        CONCAT(SSDAGI01.PATENT01, CONCAT( '',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                              "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                              "        C01REFER.feta01    AS ETA_PORT,        " & _
                              "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                              "        SSDAGI01.fecent01,                     " & _
                              "        C01REFER.frev01    AS REVALIDACION,    " & _
                              "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                              "        C01REFER.fpre01    AS PREVIO,          " & _
                              "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                              "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                              "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                              "        firmae01,                              " & _
                              "        frec01 as FecITTS,                     " & _
                              "        cvrexp01 as FORWARDER,                 " & _
                              "        cvela01  as  AIRLINE,                  " & _
                              "        feorig01,                              " & _
                              "        etalax01,                              " & _
                              "        cbuq01,                                " & _
                              "        CVEPED01                               " & _
                              " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01      " & _
                              "               INNER JOIN D05ARTIC ON REFE01=REFE05                  " & _
                              " WHERE adusec01=470 and modo01 = 'T' and cpro05 = '"&LTRIM(strfiltrosrestantes)   & "' " & Permi & _
                              " ORDER BY ETA_PORT2, ETA_PORT  "
                else
                   if strTipoOtrosFiltros = "Descripcion" then ' filtro po Descripcion de mercancia
                       strSQL =  " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                              "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                              "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                              "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                              "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                              "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                              "        SSDAGI01.PATENT01,                     " & _
                              "        CONCAT(SSDAGI01.PATENT01, CONCAT( '',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                              "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                              "        C01REFER.feta01    AS ETA_PORT,        " & _
                              "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                              "        SSDAGI01.fecent01,                     " & _
                              "        C01REFER.frev01    AS REVALIDACION,    " & _
                              "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                              "        C01REFER.fpre01    AS PREVIO,          " & _
                              "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                              "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                              "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                              "        firmae01,                              " & _
                              "        frec01 as FecITTS,                     " & _
                              "        cvrexp01 as FORWARDER,                 " & _
                              "        cvela01  as  AIRLINE,                  " & _
                              "        feorig01,                              " & _
                              "        etalax01,                              " & _
                              "        cbuq01,                                " & _
                              "        CVEPED01,                              " & _
                              "        cveptoemb,                             " & _
                              "        ADUDES01                               " & _
                              " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01      " & _
                              "               INNER JOIN D05ARTIC ON REFE01=REFE05                  " & _
                              " WHERE adusec01=470 and modo01 = 'T' and desc05 LIKE'%" & LTRIM(strfiltrosrestantes)& "%' " & Permi & _
                              " ORDER BY ETA_PORT2, ETA_PORT  "

                   end if
                end if
                  'strfiltrosrestantes   = trim(request.Form("txtfiltrosrestantes"))
                  'strTipoOtrosFiltros   = trim(request.Form("txttipoOtrosFiltros"))
                  ' txttipoOtrosFiltros
                  ' Modelo
                  ' Descripcion
          else
               if strTipoFiltro  = "BotonOtrosOpVivas" then  'Otros Flitros de captura libre
               '-------------------------------------------------
                              strAduana =Session("GAduana")
                              StrPreAdu = ""
                              if strAduana="VER" then
                                 StrPreAdu = "RKU"
                              else
                                 if strAduana="LZR" then
                                    StrPreAdu = "LZR"
                                 else
                                    if strAduana="TAM" then
                                       StrPreAdu = "CEG"
                                    else
                                       if strAduana="MEX" then
                                          StrPreAdu = "DAI"
                                       else
                                          if strAduana="MAN" then
                                             StrPreAdu = "SAP"
                                          end if
                                       end if
                                    end if
                                 end if
                              end if

                              strCodEtapa =  ""
                              Set RBusEtapa = Server.CreateObject("ADODB.Recordset")
                              RBusEtapa.ActiveConnection = MM_EXTRANET_STRING_STATUS
                              strSqlSel = " SELECT N_ETAPA " & _
                                          " FROM  ETAPS "  & _
                                          " WHERE d_abrev = 'LLP'"
                              'Response.Write(strSqlSel)
                              'Response.End
                              RBusEtapa.Source = strSqlSel
                              RBusEtapa.CursorType = 0
                              RBusEtapa.CursorLocation = 2
                              RBusEtapa.LockType = 1
                              RBusEtapa.Open()
                              if not RBusEtapa.eof then
                                 strCodEtapa = RBusEtapa.Fields.Item("N_ETAPA").Value
                              end if
                                 RBusEtapa.close
                              set RBusEtapa = Nothing

                              strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                                     "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                     "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                     "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                                     "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                     "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                                     "        SSDAGI01.PATENT01,                     " & _
                                     "        CONCAT(SSDAGI01.PATENT01, CONCAT( '',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                     "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                                     "        C01REFER.feta01    AS ETA_PORT,        " & _
                                     "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                                     "        SSDAGI01.fecent01,                     " & _
                                     "        C01REFER.frev01    AS REVALIDACION,    " & _
                                     "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                     "        C01REFER.fpre01    AS PREVIO,          " & _
                                     "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                     "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                                     "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                                     "        firmae01,                              " & _
                                     "        frec01 as FecITTS,                     " & _
                                     "        cvrexp01 as FORWARDER,                 " & _
                                     "        cvela01  as  AIRLINE,                  " & _
                                     "        feorig01,                              " & _
                                     "        etalax01,                              " & _
                                     "        cbuq01,                                " & _
                                     "        CVEPED01,                              " & _
                                     "        cveptoemb,                             " & _
                                     "        ADUDES01                               " & _
                                     "  FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01   " & _
                                     "      LEFT JOIN "&StrPreAdu&"_STATUS.ETXCOI                       " & _
                                     "            ON REFCIA01="&StrPreAdu&"_STATUS.ETXCOI.C_REFERENCIA  AND " & _
                                            StrPreAdu&"_STATUS.ETXCOI.N_ETAPA = "& strCodEtapa & _
                                     "  WHERE adusec01=470 and modo01 = 'T' AND                                     " & _
                                     "        ( isnull( f_fecha )  or f_fecha >= '" & FormatoFechaInv( DateAdd("d",-7, date() ) )  & "' )   AND " & _
                                     "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv     & _
                                     "  GROUP BY REFCIA01 " & _
                                     " ORDER BY ETA_PORT2, ETA_PORT  "

                              'strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                              '       "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                              '       "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                              '       "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                              '       "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                              '       "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                              '       "        SSDAGI01.PATENT01,                     " & _
                              '       "        CONCAT(SSDAGI01.PATENT01, CONCAT( '-',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                              '       "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                              '       "        C01REFER.feta01    AS ETA_PORT,        " & _
                              '       "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                              '       "        SSDAGI01.fecent01,                     " & _
                              '       "        C01REFER.frev01    AS REVALIDACION,    " & _
                              '       "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                              '       "        C01REFER.fpre01    AS PREVIO,          " & _
                              '       "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                              '       "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                              '       "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                              '       "        firmae01,                              " & _
                              '       "        frec01 as FecITTS,                     " & _
                              '       "        cvrexp01 as FORWARDER,                 " & _
                              '       "        cvela01  as  AIRLINE,                  " & _
                              '       "        feorig01,                              " & _
                              '       "        etalax01                               " & _
                              '       "  FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01   " & _
                              '       "      LEFT JOIN "&StrPreAdu&"_STATUS.ETXCOI                       " & _
                              '       "            ON REFCIA01="&StrPreAdu&"_STATUS.ETXCOI.C_REFERENCIA   " & _
                              '       "      LEFT JOIN "&StrPreAdu&"_STATUS.ETAPS                        " & _
                              '       "            ON  "&StrPreAdu&"_STATUS.ETXCOI.N_ETAPA = "&StrPreAdu&"_STATUS.ETAPS.N_ETAPA  AND " & _
                              '                    StrPreAdu&"_STATUS.ETAPS.D_ABREV ='LLP'               " & _
                              '       "  WHERE ( cvep01 <> 'R1') AND                                     " & _
                              '       "        ISNULL(F_FECHA)   AND " & _
                              '       "        C01REFER.REFE01 <> ''  " & Permi & strCadFiltroLinNav & strCadFiltroProv     & _
                              '       "  GROUP BY REFCIA01 "

                                    'FROM C01REFER INNER JOIN SSDAGI01
                                    '                                  ON REFCIA01= C01REFER.REFE01
                                    '              LEFT JOIN DAI_STATUS.ETXCOI
                                    '                                  ON REFCIA01=DAI_STATUS.ETXCOI.C_REFERENCIA
                                    '              LEFT JOIN DAI_STATUS.ETAPS
                                    '                                  ON  DAI_STATUS.ETXCOI.N_ETAPA = DAI_STATUS.ETAPS.N_ETAPA
                                    '                                  AND DAI_STATUS.ETAPS.D_ABREV ='LLP'
                                    'WHERE ( cvep01 <> 'R1') AND
                                    '                C01REFER.REFE01 <> ''   AND
                                    '                ISNULL(F_FECHA)    AND
                                    '                cvecli01 =1903
                                    'GROUP BY REFCIA01

                                    'strSQL = " SELECT SSDAGI01.REFCIA01  AS REFERENCIA,      " & _
                                    '         "        C01REFER.PTOEMB01  AS PORT_LOADING,    " & _
                                    '         "        C01REFER.PAISEM01  AS VESSEL_LOADING,  " & _
                                    '         "        SSDAGI01.adusec01  AS PORT_DISCHARGE,  " & _
                                    '         "        C01REFER.Naim01    AS SHIPPING_LINE,   " & _
                                    '         "        SSDAGI01.REGBAR01  AS VESSEL,          " & _
                                    '         "        SSDAGI01.PATENT01,                     " & _
                                    '         "        CONCAT(SSDAGI01.PATENT01, CONCAT( '-',SSDAGI01.NUMPED01 ) ) AS IMPORT_DOCUMENT, " & _
                                    '         "        SSDAGI01.CVEPRO01  AS PROVEEDOR,       " & _
                                    '         "        C01REFER.feta01    AS ETA_PORT,        " & _
                                    '         "        SSDAGI01.fecent01  AS ETA_PORT2,       " & _
                                    '         "        SSDAGI01.fecent01,                     " & _
                                    '         "        C01REFER.frev01    AS REVALIDACION,    " & _
                                    '         "        C01REFER.fcot01    AS RESQUEST_DUTIES, " & _
                                    '         "        C01REFER.fpre01    AS PREVIO,          " & _
                                    '         "        C01REFER.fdsp01    AS DATE_CUSTOM,     " & _
                                    '         "        SSDAGI01.cvemts01  AS MODALIDAD,       " & _
                                    '         "        SSDAGI01.desf0101  AS FACTURAS,        " & _
                                    '         "        firmae01,                              " & _
                                    '         "        frec01 as FecITTS,                     " & _
                                    '         "        feorig01 as FECETDLOAD                 " & _
                                    '         " FROM C01REFER INNER JOIN SSDAGI01 ON REFCIA01= C01REFER.REFE01      " & _
                                    '         "      LEFT JOIN "&StrPreAdu&"_STATUS.ETXCOI                         " & _
                                    '         "            ON REFCIA01="&StrPreAdu&"_STATUS.ETXCOI.C_REFERENCIA   " & _
                                    '         "      LEFT JOIN "&StrPreAdu&"_STATUS.ETAPS                                    " & _
                                    '         "            ON  "&StrPreAdu&"_STATUS.ETXCOI.N_ETAPA = "&StrPreAdu&"_STATUS.ETAPS.N_ETAPA AND " & _
                                    '                           StrPreAdu&"_STATUS.ETAPS.D_ABREV ='LLP'                         " & _
                                    '         "  WHERE ( cvep01 <> 'R1') AND                                        " & _
                                    '         "        C01REFER.REFE01 <> '' " & Permi & strCadFiltroLinNav & strCadFiltroModal & strCadFiltroProv & _
                                    '         "  GROUP BY REFCIA01 "
               '-------------------------------------------------
               end if
          end if
         end if


       'response.Write(strSQL)
       'response.end


         '*********************************************************************************************************************************************************************************************************************************************
         ' Traemos plantilla del tracking y STD de cada una de las etapas
         '*********************************************************************************************************************************************************************************************************************************************
           tmpEnviaOper = "I"
           strAduana =Session("GAduana")
           if strAduana="VER" then
              strPrefijo = "430"
           else
              if strAduana="LZR" then
                 strPrefijo = "510"
              else
                 if strAduana="TAM" then
                    strPrefijo = "810"
                 else
                    if strAduana="MEX" then
                       strPrefijo = "470"
                    else
                       if strAduana="MAN" then
                          strPrefijo = "160"
                       end if
                    end if
                 end if
              end if
           end if
           NumPlantilla = BuscaPlantillaConte(strUsuario,tmpEnviaOper,strPrefijo,"TRACKING")
           'Response.End
           ' strSQLPlSTD = " SELECT D.n_orden as orden,     " & _
                         ' "        E.d_abrev as inicio,    " & _
                         ' "        B.d_abrev as fin,       " & _
                         ' "        transal   as modalidad, " & _
                         ' "        numdia01  as dias,      " & _
                         ' "        tipdia01  as tipod      " & _
                         ' " FROM ETXPL AS D,   " & _
                         ' "      ETAPS AS E ,  " & _
                         ' "      ETAPS AS B    " & _
                         ' " INNER JOIN D01STD ON E.N_ETAPA= ETPINI01 and tipoadu='AEREA' " & _
                         ' "      and B.N_ETAPA= etpfin01  " & _
                         ' " WHERE D.n_plantilla = " & NumPlantilla & " and " & _
                         ' "       D.n_etapa = E.n_etapa " & _
                         ' " order by D.n_orden "
						 
		     strSQLPlSTD =" SELECT  " & _
						 "		D.n_orden as orden, " & _
						 "		E.d_abrev as inicio, " & _
						 "		B.d_abrev as fin, " & _
						 "	   I.transal as modalidad,  " & _
						 "		I.numdia01 as dias,  " & _
						 "		I.tipdia01 as tipod  " & _
						 "	FROM  " & _
						 "		ETXPL AS D " & _
						 "		INNER JOIN ETAPS AS E ON D.n_etapa = E.n_etapa " & _
						 "		INNER JOIN D01STD AS I ON  E.N_ETAPA= I.etpini01  " & _
						 "		INNER JOIN ETAPS AS B ON  B.n_etapa = I.etpfin01 " & _
						 "	WHERE " & _
						 "		D.n_plantilla = 1  " & _
						 "		AND I.tipoadu='AEREA' " & _
						 "		order by D.n_orden "
						 
           'Response.Write(strSQLPlSTD)
		  ' Response.Write(MM_EXTRANET_STRING_STATUS)
           'Response.End

           Set RsPlSTD = Server.CreateObject("ADODB.Recordset")
           RsPlSTD.ActiveConnection = MM_EXTRANET_STRING_STATUS
           RsPlSTD.Source = strSQLPlSTD
           RsPlSTD.CursorType = 0
           RsPlSTD.CursorLocation = 2
           RsPlSTD.LockType = 1
           RsPlSTD.Open()

           'variables para los std
           StdEtdLoad = 0 'std para ETDLOAD
           StdATAPORTDSP = 0 'std para ATAPORT A DESPACHO
           StdDSPWH      = 0 'std para DESPACHO A WAREHOUSE

           tipoStdEtdLoad    = 1 'tipo de dias de std ETDLOAD
           tipoStdATAPORTDSP = 1 'tipo de dias de std ATAPORT A DESPACHO
           tipoStdDSPWH      = 1 'tipo de dias de std DESPACHO A WAREHOUSE

           'response.end
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
         '*********************************************************************************************************************************************************************************************************************************************






               '*********************************************************************************************************************************************************************************************************************************************
               Dim arrcampos()      ' Vector utilizado para almacenar los campos
               ReDim arrcampos(4,1) ' Necesitamos todos los campos
               'ReDim PRESERVE arrcampos(2,intconcmp + 1)

               Set rslistacampos = Server.CreateObject("ADODB.Recordset")
               rslistacampos.ActiveConnection = MM_EXTRANET_STRING_STATUS
               strSqlSel = " SELECT ORDEN ,   " & _
                           "        NOMCAM01,   " & _
                           "        TITULO01, " & _
                           "        DESC01    " & _
                           " FROM etpxcmp     " & _
                           " INNER JOIN campostrk " & _
                           "       ON CVECAM   =  CVECAM01 " & _
                           " WHERE n_plantilla = " & NumPlantilla & _
                           " order by orden "
               'Response.Write(strSqlSel)
               'Response.End
               rslistacampos.Source = strSqlSel
               rslistacampos.CursorType = 0
               rslistacampos.CursorLocation = 2
               rslistacampos.LockType = 1
               rslistacampos.Open()
               if not rslistacampos.eof then
                   'strHTML = strHTML & "<tr  align=""center"" >"& chr(13) & chr(10)
                   intconcmp = 0
                   While NOT rslistacampos.EOF
                       'ORDEN
                       strcmpaux = rslistacampos.Fields.Item("TITULO01").Value
                       'NOMCAM01
                       if strcmpaux <>"" and ltrim(strcmpaux) <> "" then
                          'strHTML = strHTML & "<td width=""120"" bgcolor=""#0066FF"" nowrap><strong><font color=""#FFFFFF"" size=""1"" face=""Arial, Helvetica, sans-serif""> " & strcmpaux & " </font></strong></td>" & chr(13) & chr(10)
                          arrcampos(0,intconcmp) = rslistacampos.Fields.Item("ORDEN").Value    ' orden
                          arrcampos(1,intconcmp) = rslistacampos.Fields.Item("NOMCAM01").Value ' cvecam
                          arrcampos(2,intconcmp) = strcmpaux  ' titulo
                          ReDim PRESERVE arrcampos(4,intconcmp + 1)
                          intconcmp = intconcmp + 1
                       end if
                   rslistacampos.movenext
                   Wend
                   'strHTML = strHTML & "</tr>"& chr(13) & chr(10)
               end if
               rslistacampos.close
               set rslistacampos = Nothing
               '*******************************************************************









         if not trim(strSQL)="" then
            Set RsRep = Server.CreateObject("ADODB.Recordset")
            RsRep.ActiveConnection = MM_EXTRANET_STRING
            RsRep.Source = strSQL
            RsRep.CursorType = 0
            RsRep.CursorLocation = 2
            RsRep.LockType = 1
            RsRep.Open()

            banCargaRun = false

            'response.write(strSQL)
            'response.end

            if not RsRep.eof then
               DespliegaEncabezado()
               'response.end
               intColumna = 1
               While NOT RsRep.EOF AND not banCargaRun
                         ' ya tenemos los registros a nivel pedimento, ahora vamos por todos los campos a nivel pedimento restantes
                         StrRefer = RsRep.Fields.Item("REFERENCIA").Value
                         '**************************************************************************************************************

                          Bolbanrecti = True
                         ' verificamos que la referenccia no sea una R1
                         ' si es una R1 entonces vericamos si la original no tiene

                         'StrCvePed = RsRep.Fields.Item("cveped01").Value
                         'if StrCvePed <> "" and ltrim(StrCvePed) = "R1" and strTipoFiltro  = "BotonOtrosOpVivas"  then
                         '    strRefRecti = ""
                         '    if StrRefer <> "" then
                         '        Set Rsrecti = Server.CreateObject("ADODB.Recordset")
                         '        Rsrecti.ActiveConnection = MM_EXTRANET_STRING
                         '        strSqlSel =  " SELECT REFORG06 " & _
                         '                     " FROM SSRECP06 " & _
                         '                     " WHERE REFCIA06 ='" & ltrim(StrRefer)&"'"
                         '        'Response.Write(strSqlSel)
                         '        'Response.End
                         '        Rsrecti.Source = strSqlSel
                         '        Rsrecti.CursorType = 0
                         '        Rsrecti.CursorLocation = 2
                         '        Rsrecti.LockType = 1
                         '        Rsrecti.Open()
                         '        if not Rsrecti.eof then
                         '            While NOT Rsrecti.EOF
                         '               strRefRecti      = Rsrecti.Fields.Item("REFORG06").Value
                         '               Rsrecti.movenext
                         '            Wend
                         '            Bolbanrecti = False
                         '              Dim oConn
                         '              Set oConn = Server.CreateObject("ADODB.Connection")
                         '              oConn.Open(MM_EXTRANET_STRING_STATUS)
                         '              strSQL = " INSERT INTO etxcoi    " & _
                         '                       " SELECT n_secuenc,     " & _
                         '                       "        n_etapa,       " & _
                         '                       "        f_fecha,       " & _
                         '                       "        t_hora,        " & _
                         '                       "'" & StrRefer & "',    "& _
                         '                       "        c_conte,       " & _
                         '                       "        I_completo,    " & _
                         '                       "        I_visible,     " & _
                         '                       "        m_observ       " & _
                         '                       " FROM etxcoi           " & _
                         '                       " where C_referencia='" &  ltrim(strRefRecti)&"'"
                         '              'Response.Write(strSQL)
                         '              'Response.End
                         '              oConn.Execute(strSQL)
                         '              oConn.Close
                         '              set oConn = nothing
                         '        end if
                         '        Rsrecti.close
                         '        set Rsrecti = Nothing
                         '    end if
                         'end if

                    if Bolbanrecti = True then
                         ' GUIA
                         strGuiaMaster      = ""
                         strGuiaMasterHouse = ""
                         if StrRefer <> "" then
                             Set Recguia = Server.CreateObject("ADODB.Recordset")
                             Recguia.ActiveConnection = MM_EXTRANET_STRING
                             'strSqlSel =  "select numgui04 from ssguia04 where refcia04='" & ltrim(StrRefer)&"'"
                             strSqlSel =  " SELECT  IF( IDNGUI04=1,numgui04,'') AS guiaMaster,  " & _
                                          "         IF( IDNGUI04=2,numgui04,'') AS guiaHouse    " & _
                                          " from ssguia04  " & _
                                          " where refcia04='" & ltrim(StrRefer)&"'"
                             'Response.Write(strSqlSel)
                             'Response.End
                             Recguia.Source = strSqlSel
                             Recguia.CursorType = 0
                             Recguia.CursorLocation = 2
                             Recguia.LockType = 1
                             Recguia.Open()
                             if not Recguia.eof then
                                 strGuiaMaster      = Recguia.Fields.Item("guiaMaster").Value
                                 strGuiaMasterHouse = Recguia.Fields.Item("guiaHouse").Value
                                 intcountguia1=1
                                 intcountguia2=1
                                 While NOT Recguia.EOF
                                    if Recguia.Fields.Item("guiaMaster").Value <> "" then
                                       if intcountguia1 = 1 then
                                           strGuiaMaster      = Recguia.Fields.Item("guiaMaster").Value
                                       else
                                           strGuiaMaster      = strGuiaMaster & ", "& Recguia.Fields.Item("guiaMaster").Value
                                       end if
                                       intcountguia1= intcountguia1 + 1
                                    end if

                                    if Recguia.Fields.Item("guiaHouse").Value <> "" then
                                       if intcountguia2 = 1 then
                                           strGuiaMasterHouse = Recguia.Fields.Item("guiaHouse").Value
                                       else
                                           strGuiaMasterHouse = strGuiaMasterHouse & ", "& Recguia.Fields.Item("guiaHouse").Value
                                       end if
                                       intcountguia2= intcountguia2 + 1
                                    end if

                                 Recguia.movenext

                                 Wend
                             end if
                             Recguia.close
                             set Recguia = Nothing
                         end if

                         '**************************************************************************************************************
                         ' fORWARDWERS Y LINEAS AEREAS
                          strForwarder = RsRep.Fields.Item("FORWARDER").Value
                           if strForwarder <> "" then
                              Set RForwLA = Server.CreateObject("ADODB.Recordset")
                              RForwLA.ActiveConnection = MM_EXTRANET_STRING
                              strSqlSel =  " select CVEFOR01  " & _
                                           " from c01reexp        " & _
                                           " where cvrexP01='"&ltrim(strForwarder)&"' "
                              'Response.Write(strSqlSel)
                              'Response.End
                              RForwLA.Source = strSqlSel
                              RForwLA.CursorType = 0
                              RForwLA.CursorLocation = 2
                              RForwLA.LockType = 1
                              RForwLA.Open()
                              if not RForwLA.eof then
                                  strForwarder = RForwLA.Fields.Item("CVEFOR01").Value
                              else
                                  strForwarder = ""
                              end if
                              RForwLA.close
                              set RForwLA = Nothing
                          end if

                         ' strAirLine = RsRep.Fields.Item("AIRLINE").Value
                         ' if strAirLine <> "" then
                         '     Set RForwLA = Server.CreateObject("ADODB.Recordset")
                         '     RForwLA.ActiveConnection = MM_EXTRANET_STRING
                         '     strSqlSel =  " select DESC01,cvefor01  " & _
                         '                  " from c01airln  " & _
                         '                  " WHERE CVELA01='"&ltrim(strAirLine)&"' "
                         '     'Response.Write(strSqlSel)
                         '     'Response.End
                         '     RForwLA.Source = strSqlSel
                         '     RForwLA.CursorType = 0
                         '     RForwLA.CursorLocation = 2
                         '     RForwLA.LockType = 1
                         '     RForwLA.Open()
                         '     if not RForwLA.eof then
                         '         'strAirLine = RForwLA.Fields.Item("DESC01").Value
                         '         strAirLine = RForwLA.Fields.Item("cvefor01").Value
                         '     else
                         '         strAirLine = ""
                         '     end if
                         '     RForwLA.close
                         '     set RForwLA = Nothing
                         ' end if
                         '   if strAirLine <> "" then
                         '      'strForwarder = strForwarder&"/"&strAirLine
                         '       strForwarder = strAirLine
                         '   end if
                         '**************************************************************************************************************


                         '**************************************************************************************************************
                         ' TRAEMOS EL STD DEL ETD LOAD DE ACUERDO A LA NAVIERA Y AL PUERTO DE ORIGEN
                         '**************************************************************************************************************

                         'strCvepto01 = RsRep.Fields.Item("cveptoemb").Value
                         'StrAdutmp   = RsRep.Fields.Item("PORT_DISCHARGE").Value
                         'strNaimtmp  = RsRep.Fields.Item("SHIPPING_LINE").Value

                         strForwardertmp = RsRep.Fields.Item("FORWARDER").Value

                         if strForwardertmp <> "" then
                             Set Rshipping_line = Server.CreateObject("ADODB.Recordset")
                             Rshipping_line.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT clifor01, " & _
                                          "        numdia01  " & _
                                          " FROM d01reexp    " & _
                                          " where cverex01 = '" & ltrim(strForwardertmp) & "' " & permi
                             'if StrRefer = "SAP07-6550" then
                             'Response.Write(strSqlSel)
                             'Response.End
                             'end if
                             Rshipping_line.Source = strSqlSel
                             Rshipping_line.CursorType = 0
                             Rshipping_line.CursorLocation = 2
                             Rshipping_line.LockType = 1
                             Rshipping_line.Open()
                             if not Rshipping_line.eof then
                                 strForwardertmp = Rshipping_line.Fields.Item("clifor01").Value
                                 StdEtdLoad = Rshipping_line.Fields.Item("numdia01").Value
                             else
                                 StdEtdLoad = 0
                                 strForwardertmp = ""
                             end if
                             Rshipping_line.close
                             set Rshipping_line = Nothing
                         end if
                         'Response.End
                         if strForwardertmp <> "" then
                           strForwarder = strForwardertmp
                         end if
                         '**************************************************************************************************************


                         '**************************************************************************************************************
                         ' c01refer.Naim01 AS SHIPPING_LINE, ' catalogo c01navi  c01refer.Naim01=c01naviE.CVE01
                         strNaim01 = RsRep.Fields.Item("SHIPPING_LINE").Value
                         if strNaim01 <> "" then
                             Set Rshipping_line = Server.CreateObject("ADODB.Recordset")
                             Rshipping_line.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  "select nom01 from c01navie where cve01='" & ltrim(strNaim01)&"'"
                             'Response.Write(strSqlSel)
                             'Response.End
                             Rshipping_line.Source = strSqlSel
                             Rshipping_line.CursorType = 0
                             Rshipping_line.CursorLocation = 2
                             Rshipping_line.LockType = 1
                             Rshipping_line.Open()
                             if not Rshipping_line.eof then
                                 strNaim01 = Rshipping_line.Fields.Item("nom01").Value
                             else
                                 strNaim01 = ""
                             end if
                             Rshipping_line.close
                             set Rshipping_line = Nothing
                         end if
                         'Response.End
                         '**************************************************************************************************************

                         ' catalogo ssreba17  SSDAGI01.REGBAR01=ssreba17.regbar17
                         ' strfirmae01 = RsRep.Fields.Item("firmae01").Value
                         strvessel = ""
                         'if strfirmae01 <> "" then
                         ' if not isnull(strfirmae01) and not isempty(strfirmae01) then
                         '     strvessel = RsRep.Fields.Item("VESSEL").Value
                         '     if strvessel <> "" then
                         '         Set Rvessel = Server.CreateObject("ADODB.Recordset")
                         '         Rvessel.ActiveConnection = MM_EXTRANET_STRING
                         '         strSqlSel =  "select nombar17 from ssreba17 where regbar17='" & ltrim(strvessel)&"'"
                         '         'Response.Write(strSqlSel)
                         '         'Response.End
                         '         Rvessel.Source = strSqlSel
                         '         Rvessel.CursorType = 0
                         '         Rvessel.CursorLocation = 2
                         '         Rvessel.LockType = 1
                         '         Rvessel.Open()
                         '         if not Rvessel.eof then
                         '             strvessel = Rvessel.Fields.Item("nombar17").Value
                         '         else
                         '             strvessel = ""
                         '         end if
                         '         Rvessel.close
                         '         set Rvessel = Nothing
                         '     end if
                         ' else
                         '       strvessel = RsRep.Fields.Item("cbuq01").Value
                         '       if strvessel <> "" then
                         '           Set Rvessel = Server.CreateObject("ADODB.Recordset")
                         '           Rvessel.ActiveConnection = MM_EXTRANET_STRING
                         '           strSqlSel =  " select nomb06 " & _
                         '                        " from c06barco " & _
                         '                        " where clav06='" & Cstr(strvessel)&"'"
                         '           'Response.Write(strSqlSel)
                         '           'Response.End
                         '           Rvessel.Source = strSqlSel
                         '           Rvessel.CursorType = 0
                         '           Rvessel.CursorLocation = 2
                         '           Rvessel.LockType = 1
                         '           Rvessel.Open()
                         '           if not Rvessel.eof then
                         '               strvessel = Rvessel.Fields.Item("nomb06").Value
                         '           else
                         '               strvessel = ""
                         '           end if
                         '           Rvessel.close
                         '           set Rvessel = Nothing
                         '       end if
                         ' end if
                         'Response.End

                         strvessel = "N/A"

                         '**************************************************************************************************************

                         'catalogo ssprov22  SSDAGI01.CVEPRO01=ssprov22.cvepro22

                         strProveedor = RsRep.Fields.Item("PROVEEDOR").Value
                         if strProveedor <> "" then
                             Set RProv = Server.CreateObject("ADODB.Recordset")
                             RProv.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  "select nompro22,npscli22 from ssprov22 where cvepro22=" & ltrim(strProveedor)
                             'Response.Write(strSqlSel)
                             'Response.End
                             RProv.Source = strSqlSel
                             RProv.CursorType = 0
                             RProv.CursorLocation = 2
                             RProv.LockType = 1
                             RProv.Open()
                             if not RProv.eof then
                                 strProveedor = RProv.Fields.Item("npscli22").Value
                             else
                                 strProveedor = ""
                             end if
                             RProv.close
                             set RProv = Nothing
                         end if
                         'Response.End
                         '**************************************************************************************************************

                         ' 24-09-07 - Samsung pidio que en el campo modalidad
                         ' para el tracking aereo se pusiera "N/A"

                         ''catalogo ssmtra30 ssdagi01=ssmtra30.clavet30
                         ' strModalidad = RsRep.Fields.Item("MODALIDAD").Value
                         ' if strModalidad <> "" then
                         '     Set RModalidad = Server.CreateObject("ADODB.Recordset")
                         '     RModalidad.ActiveConnection = MM_EXTRANET_STRING
                         '     strSqlSel =  "select descri30 from ssmtra30 where clavet30=" & ltrim(strModalidad)
                         '     'Response.Write(strSqlSel)
                         '     'Response.End
                         '     RModalidad.Source = strSqlSel
                         '     RModalidad.CursorType = 0
                         '     RModalidad.CursorLocation = 2
                         '     RModalidad.LockType = 1
                         '     RModalidad.Open()
                         '     if not RModalidad.eof then
                         '         strModalidad = RModalidad.Fields.Item("descri30").Value
                         '     else
                         '         strModalidad = ""
                         '     end if
                         '     RModalidad.close
                         '     set RModalidad = Nothing
                         ' end if
                         ''Response.End
                         '**************************************************************************************************************
                         strModalidad = "N/A"

                         ' Impuestos
                         strImpuestos = ""
                         if StrRefer <> "" then
                             Set RImpuestos = Server.CreateObject("ADODB.Recordset")
                             RImpuestos.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT SUM(import36) as Impuestos " & _
                                          " FROM sscont36         " & _
                                          " WHERE  REFCIA36 = '"&ltrim(StrRefer)&"'  AND " & _
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
                         'Response.End
                         '**************************************************************************************************************

                         ' Cuentas de Gastos
                         strCuentaGastos = ""
                         strFecCuentaGastos = ""
                         if StrRefer <> "" then
                             Set RCuentaGastos = Server.CreateObject("ADODB.Recordset")
                             RCuentaGastos.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT e31cgast.cgas31,  " & _
                                          "        e31cgast.fech31   " & _
                                          " FROM d31refer,e31cgast   " & _
                                          " WHERE refe31  = '"&ltrim(StrRefer)&"' AND " & _
                                          "       d31refer.cgas31=e31cgast.cgas31 AND " & _
                                          "       esta31='I' "
                             'Response.Write(strSqlSel)
                             'Response.End
                             RCuentaGastos.Source = strSqlSel
                             RCuentaGastos.CursorType = 0
                             RCuentaGastos.CursorLocation = 2
                             RCuentaGastos.LockType = 1
                             RCuentaGastos.Open()
                             if not RCuentaGastos.eof then
                               intcontemp = 1
                               While NOT RCuentaGastos.EOF
                                 if intcontemp = 1 then
                                    strCuentaGastos    = RCuentaGastos.Fields.Item("cgas31").Value
                                    strFecCuentaGastos = RCuentaGastos.Fields.Item("fech31").Value
                                 else
                                    strCuentaGastos    = strCuentaGastos &", "& RCuentaGastos.Fields.Item("cgas31").Value
                                    strFecCuentaGastos = strFecCuentaGastos&", "& RCuentaGastos.Fields.Item("fech31").Value
                                 end if
                                 intcontemp = intcontemp + 1
                                 RCuentaGastos.movenext
                               Wend
                             end if
                             RCuentaGastos.close
                             set RCuentaGastos = Nothing
                         end if
                         'Response.End
                         '**************************************************************************************************************

                         ' Incoterms
                         strIncoterms = ""
                         ' if StrRefer <> "" then
                         '     Set RCuentaGastos = Server.CreateObject("ADODB.Recordset")
                         '     RCuentaGastos.ActiveConnection = MM_EXTRANET_STRING
                         '     strSqlSel =  " SELECT TERFAC39 " & _
                         '                  " FROM SSFACT39   " & _
                         '                  " WHERE REFCIA39='"&ltrim(StrRefer)&"' "
                         '     'Response.Write(strSqlSel)
                         '      'Response.End
                         '     RCuentaGastos.Source = strSqlSel
                         '     RCuentaGastos.CursorType = 0
                         '     RCuentaGastos.CursorLocation = 2
                         '     RCuentaGastos.LockType = 1
                         '     RCuentaGastos.Open()
                         '     if not RCuentaGastos.eof then
                         '       intcontemp = 1
                         '       While NOT RCuentaGastos.EOF
                         '         if intcontemp = 1 then
                         '            strIncoterms    = RCuentaGastos.Fields.Item("TERFAC39").Value
                         '         else
                         '            strIncoterms    = strIncoterms &", "& RCuentaGastos.Fields.Item("TERFAC39").Value
                         '         end if
                         '         intcontemp = intcontemp + 1
                         '         RCuentaGastos.movenext
                         '       Wend
                         '     end if
                         '     RCuentaGastos.close
                         '     set RCuentaGastos = Nothing
                         ' end if
                           'Response.End

                         '**************************************************************************************************************
                         ' Vamos por las mercancias
                           strPO_Pedido = ""
                           strPO_PedidoNA = ""
                           strDescMerc  = ""
                           strModelo    = ""
                           strDescCode  = ""
                           Set RMercancias = Server.CreateObject("ADODB.Recordset")
                           RMercancias.ActiveConnection = MM_EXTRANET_STRING
                           'strSqlSel =  "Select  pedi05, desc05, cpro05 from d05artic where refe05='" & ltrim(StrRefer) & "' "
                           'strSqlSel = " Select  refe05,pedi05, desc05, cpro05, descri05,tpmerc05 " & _
                           '            " from d05artic left join c05tmerc on tpmerc05 =clave05 " & _
                           '            " where refe05='" & ltrim(StrRefer) & "' "
                           strSqlSel = " Select  refe05,pedi05, desc05, cpro05,descod05 " & _
                                       " from d05artic  " & _
                                       " where refe05='" & ltrim(StrRefer) & "' "
                           'Response.Write(strSqlSel)
                           'Response.End
                           RMercancias.Source = strSqlSel
                           RMercancias.CursorType = 0
                           RMercancias.CursorLocation = 2
                           RMercancias.LockType = 1
                           RMercancias.Open()
                           if not RMercancias.eof then
                           intcontemp = 1
                           intcontped = 1
                             While NOT RMercancias.EOF
                                 if RMercancias.Fields.Item("pedi05").Value <> ""  then
                                    'AND UCase(ltrim(RMercancias.Fields.Item("pedi05").Value)) <> "S/N" AND UCase(ltrim(RMercancias.Fields.Item("pedi05").Value)) <> "N/A" AND UCase(ltrim(RMercancias.Fields.Item("pedi05").Value)) <> "SN" AND UCase(ltrim(RMercancias.Fields.Item("pedi05").Value)) <> "NA"
                                    if UCase(ltrim(RMercancias.Fields.Item("pedi05").Value)) <> "N/A" then
                                        if intcontped = 1 then
                                           strPO_Pedido  = RMercancias.Fields.Item("pedi05").Value
                                        else
                                           strPO_Pedido  = strPO_Pedido& ", "&RMercancias.Fields.Item("pedi05").Value
                                        end if
                                        intcontped = intcontped + 1
                                    else
                                       strPO_PedidoNA = "N/A"
                                    end if
                                 end if

                                 if intcontemp = 1 then
                                    strDescMerc   = RMercancias.Fields.Item("desc05").Value
                                    strModelo     = RMercancias.Fields.Item("cpro05").Value
                                    strDescCode   = RMercancias.Fields.Item("descod05").Value
                                 else
                                    strDescMerc   = strDescMerc & ", " & RMercancias.Fields.Item("desc05").Value
                                    strModelo     = strModelo & ", " & RMercancias.Fields.Item("cpro05").Value
                                    strDescCode   = strDescCode & ", " & RMercancias.Fields.Item("descod05").Value
                                 end if
                                 intcontemp = intcontemp + 1
                                 RMercancias.movenext
                             Wend

                             if strPO_Pedido = "" then
                              strPO_Pedido = strPO_PedidoNA

                             end if
                           end if


                         '**************************************************************************************************************

                         ' Cantidad de fracciones
                         strQTY = ""
                         if StrRefer <> "" then
                             Set Rfracciones = Server.CreateObject("ADODB.Recordset")
                             Rfracciones.ActiveConnection = MM_EXTRANET_STRING
                             strSqlSel =  " SELECT SUM(CANCOM02) as QTY " & _
                                          " FROM ssfrac02         " & _
                                          " WHERE  REFCIA02 = '"&ltrim(StrRefer)&"' " & _
                                          " GROUP BY refcia02 "
                             'Response.Write(strSqlSel)
                             'Response.End
                             Rfracciones.Source = strSqlSel
                             Rfracciones.CursorType = 0
                             Rfracciones.CursorLocation = 2
                             Rfracciones.LockType = 1
                             Rfracciones.Open()
                             if not Rfracciones.eof then
                                 strQTY = Rfracciones.Fields.Item("QTY").Value
                             else
                                 strQTY = ""
                             end if
                             Rfracciones.close
                             set Rfracciones = Nothing
                         end if
                         'Response.End
                         '**************************************************************************************************************

                         ' Fechas de documentos
                         'CNO -> CERTIFICADO NOM
                         'CNS -> CARTA CON NUMERO DE INSTRUCCIONES
                         'GUA -> GUIA AEREA
                         strCERTNOM  = ""
                         StrNUMSERIE = ""
                         strGUANotificacion = ""
                         strGUA      = ""
                         if StrRefer <> "" then
                             Set RFecDocu = Server.CreateObject("ADODB.Recordset")
                             RFecDocu.ActiveConnection = MM_EXTRANET_STRING
                             'strSqlSel =  " SELECT CLAV07, FECH07,ORIG07 " & _
                             '             " FROM C07DOCRE " & _
                             '             " WHERE REFE07 ='"&ltrim(StrRefer)&"' AND " & _
                             '             " CLAV07='CNO' or clav07='CNS' or clav07='GUA'  "
                             strSqlSel =  " SELECT C07DOCRE.CLAV07,  " & _
                                          "         C07DOCRE.FECH07, " & _
                                          "         C07DOCRE.ORIG07, " & _
                                          "         C07DOCOR.DESC07, " & _
                                          "         DISP07            " & _
                                          " FROM C07DOCRE LEFT JOIN C07DOCOR ON C07DOCRE.ORIG07=C07DOCOR.CLAV07 " & _
                                          " WHERE  C07DOCRE.REFE07 ='"&ltrim(StrRefer)&"' AND " & _
                                          "       (C07DOCRE.CLAV07='CNO' or " & _
                                          "        C07DOCRE.clav07='CNS' or " & _
                                          "        C07DOCRE.clav07='GUA'   )"
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
                                      'strCERTNOM  = RFecDocu.Fields.Item("FECH07").Value
                                 else
                                    if RFecDocu.Fields.Item("CLAV07").Value <>"" and ltrim(RFecDocu.Fields.Item("CLAV07").Value) = "CNS"  then
                                         if RFecDocu.Fields.Item("DISP07").Value = "F"   then
                                            StrNUMSERIE = "N/A"
                                         else
                                            StrNUMSERIE = RFecDocu.Fields.Item("FECH07").Value
                                         end if
                                         'StrNUMSERIE = RFecDocu.Fields.Item("FECH07").Value
                                    else
                                       if RFecDocu.Fields.Item("CLAV07").Value <>"" and ltrim(RFecDocu.Fields.Item("CLAV07").Value) = "GUA"  then
                                          strGUA = RFecDocu.Fields.Item("FECH07").Value
                                          strGUANotificacion = RFecDocu.Fields.Item("DESC07").Value
                                       end if
                                    end if
                                 end if
                                 RFecDocu.movenext
                             Wend
                             RFecDocu.close
                             set RFecDocu = Nothing
                         end if
                         'Response.End
                         '**************************************************************************************************************



                         ' OBSERVACIONES
                         strObservaciones = ""
                         if StrRefer <> "" then
                             Set RObservEtapas = Server.CreateObject("ADODB.Recordset")
                             RObservEtapas.ActiveConnection = MM_EXTRANET_STRING_STATUS

                             'strSQL = "   SELECT max(n_secuenc), D.n_etapa, f_fecha, m_observ " & _
                             '         "   FROM ETXPD as D                                     " & _
                             '         "   WHERE not(date_format(D.f_fecha,'%Y%m%d') = '00000000') and  " & _
                             '         "         D.c_referencia = '"&ltrim(StrRefer)&"'" & _
                             '         "   group by c_referencia,D.n_etapa     "

                             strSQL = " SELECT (n_secuenc), " & _
                                      "        D.n_etapa,   " & _
                                      "        f_fecha,     " & _
                                      "        m_observ     " & _
                                      " FROM ETXPD as D     " & _
                                      " WHERE not(date_format(D.f_fecha,'%Y%m%d') = '00000000') and  " & _
                                      "       D.c_referencia = '"&ltrim(StrRefer)&"'" & _
                                      " ORDER BY N_ETAPA, N_SECUENC "

                             'strSQL = " SELECT max(n_secuenc), " & _
                             '         "        D.n_etapa,      " & _
                             '         "        f_fecha,        " & _
                             '         "        m_observ,       " & _
                             '         "        d_abrev ,       " & _
                             '         "        d_nombre        " & _
                             '         " FROM ETXPD as D,       " & _
                             '         "      etaps             " & _
                             '         " WHERE D.n_etapa= etaps.n_etapa and                " & _
                             '         "       not(date_format(D.f_fecha,'%Y%m%d') = '00000000') " & _
                             '         " GROUP BY c_referencia,D.n_etapa   "
                             'Response.Write(strSQL)
                             'Response.End
                             RObservEtapas.Source = strSQL
                             RObservEtapas.CursorType = 0
                             RObservEtapas.CursorLocation = 2
                             RObservEtapas.LockType = 1
                             RObservEtapas.Open()
                             intcontObs = 1
                             While NOT RObservEtapas.EOF
                                 strObsTemp = RObservEtapas.Fields.Item("m_observ").Value
                                 'strObservaciones  = strObservaciones & RObservEtapas.Fields.Item("d_nombre").Value & "(" & RObservEtapas.Fields.Item("d_abrev").Value& ") "& RObservEtapas.Fields.Item("f_fecha").Value & " .-" & RObservEtapas.Fields.Item("m_observ").Value & "<br>"
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
                         'Response.End
                         '**************************************************************************************************************

                         ' FACTURAS
                         'if strfirmae01 <> "" then
                         '   StrINVOICE = RsRep.Fields.Item("FACTURAS").Value 'INVOICE
                         'else
                             StrINVOICE = ""
                             if StrRefer <> "" then
                                 Set RFactuRef = Server.CreateObject("ADODB.Recordset")
                                 RFactuRef.ActiveConnection = MM_EXTRANET_STRING
                                 strSQL = " SELECT NUMFAC39, FECFAC39 " & _
                                          " FROM SSFACT39 " & _
                                          " WHERE REFCIA39='" & ltrim(StrRefer) & "'"
                                 RFactuRef.Source = strSQL
                                 RFactuRef.CursorType = 0
                                 RFactuRef.CursorLocation = 2
                                 RFactuRef.LockType = 1
                                 RFactuRef.Open()
                                 intcontObs = 1
                                 While NOT RFactuRef.EOF
                                     StrINVOICETemp    = RFactuRef.Fields.Item("NUMFAC39").Value
                                     StrfECINVOICETemp = RFactuRef.Fields.Item("FECFAC39").Value
                                     'strObservaciones  = strObservaciones & RObservEtapas.Fields.Item("d_nombre").Value & "(" & RObservEtapas.Fields.Item("d_abrev").Value& ") "& RObservEtapas.Fields.Item("f_fecha").Value & " .-" & RObservEtapas.Fields.Item("m_observ").Value & "<br>"
                                     if StrINVOICETemp <> "" and StrfECINVOICETemp <> "" then
                                        if intcontObs = 1 then
                                           StrINVOICE  = StrINVOICETemp
                                           'StrINVOICE  = StrINVOICETemp&" de "&StrfECINVOICETemp
                                        else
                                           'StrINVOICE  = StrINVOICE & "; "& StrINVOICETemp&" de "&StrfECINVOICETemp
                                           StrINVOICE  = StrINVOICE & "; "& StrINVOICETemp
                                        end if
                                        intcontObs = intcontObs + 1
                                     end if
                                 RFactuRef.movenext
                                 Wend
                                 'PONER EL RESUMEN DE CUANTAS FACTURAS SON EJ (3)
                                 'if intcontObs > 1 then
                                 '   StrINVOICE = "("&CStr(intcontObs - 1)&"), "&StrINVOICE
                                 'end if
                                 RFactuRef.close
                                 set RFactuRef = Nothing
                             end if
                         'end if



                         ' Contenedores
                         strNumConte = ""
                         strATDRAIL  = ""
                         strETA_CP   = ""
                         strATAC_P   = ""
                         strETAW_H   = ""
                         '----------------
                         strFechaATAWH      = ""
                         strComentarioATAWH = ""
                         strHoraATAWH       = ""

                         if StrRefer <> "" then
                             Set RContenedores = Server.CreateObject("ADODB.Recordset")
                             RContenedores.ActiveConnection = MM_EXTRANET_STRING
                             'strSqlSel =  "select marc01 from d01conte where refe01 = '" & ltrim(StrRefer) & "' "
                             strSqlSel =  " select marc01, " & _
                                          "       fcTren01 as ATDRAIL, " & _
                                          "       feCont01 as ETA_CP,  " & _
                                          "       frCont01 as ATAC_P,  " & _
                                          "       feAlma01 as ETAW_H   " & _
                                          " from d01conte where refe01 = '" & ltrim(StrRefer) & "' "

                             'ATD RAIL (Fecha de Carga en Tren) d01Conte.fcTren01
                             'ETA C./P. (Estimada de Arribo Contrimodal)  d01Conte.feCont01
                             'ATA C./P. (Real de Arribo Contrimodal) d01Conte.frCont01
                             'ETA W/H (Fecha de llegada a Almacen de SEM) d01Conte.feAlma01

                             'Response.Write(strSqlSel)
                             'Response.End
                             RContenedores.Source = strSqlSel
                             RContenedores.CursorType = 0
                             RContenedores.CursorLocation = 2
                             RContenedores.LockType = 1
                             RContenedores.Open()
                             if not RContenedores.eof then
                               While NOT RContenedores.EOF
                                       strNumConte = RContenedores.Fields.Item("marc01").Value
                                       strATDRAIL  = RContenedores.Fields.Item("ATDRAIL").Value
                                       'strETA_CP   = RContenedores.Fields.Item("ETA_CP").Value
                                       'strATAC_P   = RContenedores.Fields.Item("ATAC_P").Value
                                       'strETAW_H   = RContenedores.Fields.Item("ETAW_H").Value
                                       '*********************************************
                                         strFechaATAWH      = ""
                                         strComentarioATAWH = ""
                                         strHoraATAWH       = ""
                                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                         strSqlSel = " SELECT f_fecha,   " & _
                                                     "        t_hora,   " & _
                                                     "        m_observ  " & _
                                                     " FROM etxcoi, etaps " & _
                                                     " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                     "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and    " & _
                                                     "       ltrim(c_conte)      = '" & ltrim(strNumConte) & "' and " & _
                                                     "       d_abrev      = 'LLP'             " & _
                                                     " order by n_secuenc desc                  "
                                         'Response.Write(strSqlSel)
                                         'Response.End
                                         RConteDetalle.Source = strSqlSel
                                         RConteDetalle.CursorType = 0
                                         RConteDetalle.CursorLocation = 2
                                         RConteDetalle.LockType = 1
                                         RConteDetalle.Open()
                                         if not RConteDetalle.eof then
                                             strFechaATAWH       = RConteDetalle.Fields.Item("f_fecha").Value
                                             strHoraATAWH        = RConteDetalle.Fields.Item("t_hora").Value
                                             'strComentarioATAWH  = RConteDetalle.Fields.Item("m_observ").Value
                                             strObsTemp = ""
                                             intcontObs = 1
                                             While NOT RConteDetalle.EOF
                                                 strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                 if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                    if intcontObs = 1 then
                                                       strComentarioATAWH  = RConteDetalle.Fields.Item("m_observ").Value
                                                    else
                                                       strComentarioATAWH  = strComentarioATAWH & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                    end if
                                                    intcontObs = intcontObs + 1
                                                 end if
                                             RConteDetalle.movenext
                                             Wend
                                         end if
                                         RConteDetalle.close
                                         set RConteDetalle = Nothing

                                         'Response.Write(strHoraATAWH)
                                         'Response.End
                                       '*********************************************


                                         '*********************************************
                                         strATAC_P           = ""
                                         strComentarioATAC_P = ""
                                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                         strSqlSel = " SELECT f_fecha,  " & _
                                                     "        m_observ  " & _
                                                     " FROM etxcoi, etaps " & _
                                                     " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                     "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and    " & _
                                                     "       ltrim(c_conte)      = '" & ltrim(strNumConte) & "' and " & _
                                                     "       d_abrev      = 'CP'             " & _
                                                     " order by n_secuenc desc                  "
                                         'Response.Write(strSqlSel)
                                         'Response.End
                                         RConteDetalle.Source = strSqlSel
                                         RConteDetalle.CursorType = 0
                                         RConteDetalle.CursorLocation = 2
                                         RConteDetalle.LockType = 1
                                         RConteDetalle.Open()
                                         if not RConteDetalle.eof then
                                             strATAC_P            = RConteDetalle.Fields.Item("f_fecha").Value
                                             'strComentarioATAC_P  = RConteDetalle.Fields.Item("m_observ").Value
                                             strObsTemp = ""
                                             intcontObs = 1
                                             While NOT RConteDetalle.EOF
                                                 strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                 if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                    if intcontObs = 1 then
                                                       strComentarioATAC_P  = RConteDetalle.Fields.Item("m_observ").Value
                                                    else
                                                       strComentarioATAC_P  = strComentarioATAC_P & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                    end if
                                                    intcontObs = intcontObs + 1
                                                 end if
                                             RConteDetalle.movenext
                                             Wend
                                             'strFechaConteSPL       = RConteDetalle.Fields.Item("f_fecha").Value
                                             'strHoraATAWH           = RConteDetalle.Fields.Item("t_hora").Value
                                             'strComentarioConteSPL  = RConteDetalle.Fields.Item("m_observ").Value
                                         end if
                                         RConteDetalle.close
                                         set RConteDetalle = Nothing
                                       '*********************************************

                                       '*********************************************
                                         strATDRAIL          = ""
                                         strComentarioATDRAIL = ""
                                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                         strSqlSel = " SELECT f_fecha,  " & _
                                                     "        m_observ  " & _
                                                     " FROM etxcoi, etaps " & _
                                                     " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                     "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and    " & _
                                                     "       ltrim(c_conte)      = '" & ltrim(strNumConte) & "' and " & _
                                                     "       d_abrev      = 'RAIL'            " & _
                                                     " order by n_secuenc desc                  "
                                         'Response.Write(strSqlSel)
                                         'Response.End
                                         RConteDetalle.Source = strSqlSel
                                         RConteDetalle.CursorType = 0
                                         RConteDetalle.CursorLocation = 2
                                         RConteDetalle.LockType = 1
                                         RConteDetalle.Open()
                                         if not RConteDetalle.eof then
                                             strATDRAIL            = RConteDetalle.Fields.Item("f_fecha").Value
                                             'strComentarioETAW_H  = RConteDetalle.Fields.Item("m_observ").Value
                                             strObsTemp = ""
                                             intcontObs = 1
                                             While NOT RConteDetalle.EOF
                                                 strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                 if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                    if intcontObs = 1 then
                                                       strComentarioATDRAIL  = RConteDetalle.Fields.Item("m_observ").Value
                                                    else
                                                       strComentarioATDRAIL  = strComentarioATDRAIL & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                    end if
                                                    intcontObs = intcontObs + 1
                                                 end if
                                             RConteDetalle.movenext
                                             Wend
                                             'strFechaConteSPL       = RConteDetalle.Fields.Item("f_fecha").Value
                                             'strHoraATAWH           = RConteDetalle.Fields.Item("t_hora").Value
                                             'strComentarioConteSPL  = RConteDetalle.Fields.Item("m_observ").Value
                                         end if
                                         RConteDetalle.close
                                         set RConteDetalle = Nothing
                                       '*********************************************

                                       '*********************************************
                                         strATASPLTMP           = ""
                                         strTimeSLP             = ""
                                         strComentarioATASPLTMP = ""
                                         Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                         RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                         strSqlSel = " SELECT f_fecha,  " & _
                                                     "        t_hora,   " & _
                                                     "        m_observ  " & _
                                                     " FROM etxcoi, etaps " & _
                                                     " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                     "       c_referencia = '" & ltrim(StrRefer)    & "' and    " & _
                                                     "       c_conte      = '" & ltrim(strNumConte) & "' and " & _
                                                     "       d_abrev      = 'SPL'             " & _
                                                     " order by n_secuenc desc                  "
                                         'Response.Write(strSqlSel)
                                         'Response.End
                                         RConteDetalle.Source = strSqlSel
                                         RConteDetalle.CursorType = 0
                                         RConteDetalle.CursorLocation = 2
                                         RConteDetalle.LockType = 1
                                         RConteDetalle.Open()
                                         if not RConteDetalle.eof then
                                             strATASPLTMP = RConteDetalle.Fields.Item("f_fecha").Value
                                             strTimeSLP   = RConteDetalle.Fields.Item("t_hora").Value
                                             'strComentarioATAC_P  = RConteDetalle.Fields.Item("m_observ").Value
                                             strObsTemp = ""
                                             intcontObs = 1
                                             While NOT RConteDetalle.EOF
                                                 strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                 if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                    if intcontObs = 1 then
                                                       strComentarioATASPLTMP  = RConteDetalle.Fields.Item("m_observ").Value
                                                    else
                                                       strComentarioATASPLTMP  = strComentarioATASPLTMP & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                    end if
                                                    intcontObs = intcontObs + 1
                                                 end if
                                             RConteDetalle.movenext
                                             Wend

                                         end if
                                         RConteDetalle.close
                                         set RConteDetalle = Nothing
                                       '*********************************************
                                       'strATAC_P    'fecha de arribo cuatrimodal
                                       'strETAW_H    'fecha estimada de arribo a planta samsung

                                 RContenedores.movenext
                               Wend
                             else
                             'Response.Write("No hay ningun contenedor")
                             'Response.End

                                          'strNumConte = RContenedores.Fields.Item("marc01").Value
                                          'strATDRAIL  = RContenedores.Fields.Item("ATDRAIL").Value
                                          strATDRAIL  = ""
                                          '*********************************************
                                           strFechaATAWH      = ""
                                           strComentarioATAWH = ""
                                           strHoraATAWH       = ""
                                           Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                           RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                           strSqlSel = " SELECT f_fecha,   " & _
                                                       "        t_hora,   " & _
                                                       "        m_observ  " & _
                                                       " FROM etxcoi, etaps " & _
                                                       " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                       "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and " & _
                                                       "       d_abrev      = 'LLP'             " & _
                                                       " order by n_secuenc desc                  "
                                           'Response.Write(strSqlSel)
                                           'Response.End
                                           RConteDetalle.Source = strSqlSel
                                           RConteDetalle.CursorType = 0
                                           RConteDetalle.CursorLocation = 2
                                           RConteDetalle.LockType = 1
                                           RConteDetalle.Open()
                                           if not RConteDetalle.eof then
                                               strFechaATAWH       = RConteDetalle.Fields.Item("f_fecha").Value
                                               strHoraATAWH        = RConteDetalle.Fields.Item("t_hora").Value
                                               'strComentarioATAWH  = RConteDetalle.Fields.Item("m_observ").Value
                                               strObsTemp = ""
                                               intcontObs = 1
                                               While NOT RConteDetalle.EOF
                                                   strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                   if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                      if intcontObs = 1 then
                                                         strComentarioATAWH  = RConteDetalle.Fields.Item("m_observ").Value
                                                      else
                                                         strComentarioATAWH  = strComentarioATAWH & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                      end if
                                                      intcontObs = intcontObs + 1
                                                   end if
                                               RConteDetalle.movenext
                                               Wend
                                           end if
                                           RConteDetalle.close
                                           set RConteDetalle = Nothing

                                           'Response.Write(strHoraATAWH)
                                           'Response.End
                                           '*********************************************

                                           '*********************************************
                                           strATAC_P           = ""
                                           strComentarioATAC_P = ""
                                           Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                           RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                           strSqlSel = " SELECT f_fecha,  " & _
                                                       "        m_observ  " & _
                                                       " FROM etxcoi, etaps " & _
                                                       " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                       "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and    " & _
                                                       "       d_abrev      = 'CP'             " & _
                                                       " order by n_secuenc desc                  "
                                           'Response.Write(strSqlSel)
                                           'Response.End
                                           RConteDetalle.Source = strSqlSel
                                           RConteDetalle.CursorType = 0
                                           RConteDetalle.CursorLocation = 2
                                           RConteDetalle.LockType = 1
                                           RConteDetalle.Open()
                                           if not RConteDetalle.eof then
                                               strATAC_P            = RConteDetalle.Fields.Item("f_fecha").Value
                                               'strComentarioATAC_P  = RConteDetalle.Fields.Item("m_observ").Value
                                               strObsTemp = ""
                                               intcontObs = 1
                                               While NOT RConteDetalle.EOF
                                                   strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                   if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                      if intcontObs = 1 then
                                                         strComentarioATAC_P  = RConteDetalle.Fields.Item("m_observ").Value
                                                      else
                                                         strComentarioATAC_P  = strComentarioATAC_P & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                      end if
                                                      intcontObs = intcontObs + 1
                                                   end if
                                               RConteDetalle.movenext
                                               Wend
                                               'strFechaConteSPL       = RConteDetalle.Fields.Item("f_fecha").Value
                                               'strHoraATAWH           = RConteDetalle.Fields.Item("t_hora").Value
                                               'strComentarioConteSPL  = RConteDetalle.Fields.Item("m_observ").Value
                                           end if
                                           RConteDetalle.close
                                           set RConteDetalle = Nothing
                                         '*********************************************

                                         '*********************************************
                                           strATDRAIL          = ""
                                           strComentarioATDRAIL = ""
                                           Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                           RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                           strSqlSel = " SELECT f_fecha,  " & _
                                                       "        m_observ  " & _
                                                       " FROM etxcoi, etaps " & _
                                                       " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                       "       ltrim(c_referencia) = '" & ltrim(StrRefer)    & "' and " & _
                                                       "       d_abrev      = 'RAIL'            " & _
                                                       " order by n_secuenc desc                  "
                                           'Response.Write(strSqlSel)
                                           'Response.End
                                           RConteDetalle.Source = strSqlSel
                                           RConteDetalle.CursorType = 0
                                           RConteDetalle.CursorLocation = 2
                                           RConteDetalle.LockType = 1
                                           RConteDetalle.Open()
                                           if not RConteDetalle.eof then
                                               strATDRAIL            = RConteDetalle.Fields.Item("f_fecha").Value
                                               'strComentarioETAW_H  = RConteDetalle.Fields.Item("m_observ").Value
                                               strObsTemp = ""
                                               intcontObs = 1
                                               While NOT RConteDetalle.EOF
                                                   strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                   if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                      if intcontObs = 1 then
                                                         strComentarioATDRAIL  = RConteDetalle.Fields.Item("m_observ").Value
                                                      else
                                                         strComentarioATDRAIL  = strComentarioATDRAIL & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                      end if
                                                      intcontObs = intcontObs + 1
                                                   end if
                                               RConteDetalle.movenext
                                               Wend
                                               'strFechaConteSPL       = RConteDetalle.Fields.Item("f_fecha").Value
                                               'strHoraATAWH           = RConteDetalle.Fields.Item("t_hora").Value
                                               'strComentarioConteSPL  = RConteDetalle.Fields.Item("m_observ").Value
                                           end if
                                           RConteDetalle.close
                                           set RConteDetalle = Nothing
                                         '*********************************************

                                         '*********************************************
                                           strATASPLTMP           = ""
                                           strTimeSLP             = ""
                                           strComentarioATASPLTMP = ""
                                           Set RConteDetalle = Server.CreateObject("ADODB.Recordset")
                                           RConteDetalle.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                           strSqlSel = " SELECT f_fecha,  " & _
                                                       "        t_hora,   " & _
                                                       "        m_observ  " & _
                                                       " FROM etxcoi, etaps " & _
                                                       " where etxcoi.n_etapa = etaps.n_etapa and " & _
                                                       "       c_referencia = '" & ltrim(StrRefer)    & "' and " & _
                                                       "       d_abrev      = 'SPL'             " & _
                                                       " order by n_secuenc desc                  "
                                           'Response.Write(strSqlSel)
                                           'Response.End
                                           RConteDetalle.Source = strSqlSel
                                           RConteDetalle.CursorType = 0
                                           RConteDetalle.CursorLocation = 2
                                           RConteDetalle.LockType = 1
                                           RConteDetalle.Open()
                                           if not RConteDetalle.eof then
                                               strATASPLTMP = RConteDetalle.Fields.Item("f_fecha").Value
                                               strTimeSLP   = RConteDetalle.Fields.Item("t_hora").Value
                                               'strComentarioATAC_P  = RConteDetalle.Fields.Item("m_observ").Value
                                               strObsTemp = ""
                                               intcontObs = 1
                                               While NOT RConteDetalle.EOF
                                                   strObsTemp = RConteDetalle.Fields.Item("m_observ").Value
                                                   if strObsTemp <>"" and ltrim(strObsTemp) <> "" then
                                                      if intcontObs = 1 then
                                                         strComentarioATASPLTMP  = RConteDetalle.Fields.Item("m_observ").Value
                                                      else
                                                         strComentarioATASPLTMP  = strComentarioATASPLTMP & " ; "& RConteDetalle.Fields.Item("m_observ").Value
                                                      end if
                                                      intcontObs = intcontObs + 1
                                                   end if
                                               RConteDetalle.movenext
                                               Wend

                                           end if
                                           RConteDetalle.close
                                           set RConteDetalle = Nothing
                                         '*********************************************

                             end if
                             RContenedores.close
                             set RContenedores = Nothing


                             '*************************************************************

                                       'StrPOTD               = "" 'OTD 2
                                       'StrPITTS              = RsRep.Fields.Item("FecITTS").Value'ASIGNADO ITTS
                                       'StrPBL                = strGuia 'BILL OF LADING
                                       'StrPCONTAINER         = strNumConte  'CONTAINER
                                       'StrPP_O               = strPO_Pedido  'P/O
                                       'StrPPORT_OF_LOADING   = RsRep.Fields.Item("PORT_LOADING").Value&","&RsRep.Fields.Item("VESSEL_LOADING").Value 'PORT OF LOADING

                                       'StrAdu = RsRep.Fields.Item("PORT_DISCHARGE").Value
                                       'if ltrim(StrAdu)="430" then
                                       '   StrPPORT_OF_DISCHARGE =  "VERACRUZ"
                                       'else
                                       '  if ltrim(StrAdu)="160" then
                                       '     StrPPORT_OF_DISCHARGE =  "MANZANILLO"
                                       '  else
                                       '     if ltrim(StrAdu)="200" or ltrim(StrAdu)="202" or ltrim(StrAdu)="470" then
                                       '        StrPPORT_OF_DISCHARGE =  "MEXICO"
                                       '     else
                                       '        if ltrim(StrAdu)="380" or ltrim(StrAdu)="810" then
                                       '           StrPPORT_OF_DISCHARGE =  "TAMPICO"
                                       '        else
                                       '           if ltrim(StrAdu)="510" then
                                       '              StrPPORT_OF_DISCHARGE =  "LAZARO CARDENAS"
                                       '           end if
                                       '        end if
                                       '     end if
                                       '  end if
                                       'end if

                                       'StrPSHIPPING_LINE     = strNaim01    'SHIPPING LINE
                                       'StrPVESSEL            = strvessel    'VESSEL
                                       'StrPIMPORT_DOCUMENT   = RsRep.Fields.Item("IMPORT_DOCUMENT").Value'IMPORT DOCUMENT
                                       'StrPPROVEEDOR         = strProveedor 'PROVEEDOR
                                       'StrPINVOICE           = RsRep.Fields.Item("FACTURAS").Value 'INVOICE
                                       'StrPMODEL             = strModelo    'MODEL
                                       'StrPDESCRIPTION       = strDescMerc  'DESCRIPTION
                                       'StrPDESCRIPTION_CODE  = ""           'DESCRIPTION CODE
                                       'StrPQTY               = strQTY       'QTY
                                       ''StrPETD_LOAD          = RsRep.Fields.Item("FECETDLOAD").Value'ETD LOAD
                                       'StrPETD_LOAD          = ""  'ETD LOAD
                                       'strfirmae01 = RsRep.Fields.Item("firmae01").Value
                                       'if strfirmae01 = "" then
                                       '   StrPETA_PORT       = RsRep.Fields.Item("ETA_PORT").Value  'ETA PORT
                                       'else
                                       '   StrPETA_PORT       = RsRep.Fields.Item("ETA_PORT2").Value 'ETA PORT
                                       'end if
                                       'StrPATA_PORT          = RsRep.Fields.Item("fecent01").Value  'ATA PORT
                                       'StrPNUMS_SERIE        = StrNUMSERIE        'NUMS. SERIE
                                       'StrPCERT_NOM          = strCERTNOM         'CERT. NOM
                                       'StrPREVALIDACION      = RsRep.Fields.Item("REVALIDACION").Value    'REVALIDACION
                                       'StrPRESQUEST_DUTIES   = RsRep.Fields.Item("RESQUEST_DUTIES").Value 'RESQUEST DUTIES
                                       'StrPAMOUNT_OF_DUTIES  = strImpuestos                               'AMOUNT OF DUTIES
                                       'StrPPREVIO            = RsRep.Fields.Item("PREVIO").Value          'PREVIO
                                       'StrPDATE_OF_CUSTOM    = RsRep.Fields.Item("DATE_CUSTOM").Value     'DATE OF CUSTOM CLEARANCE
                                       'StrPATD_RAIL          = strATDRAIL     'ATD  RAIL
                                       'StrPETA_CP            = strETA_CP      'ETA C./P.
                                       'StrPATA_CP            = strATAC_P      'ATA C./P.
                                       'StrPETA_WH            = strETAW_H      'ETA W/H
                                       'StrPATA_WH            = strFechaATAWH  'ATA W-H
                                       'StrPTIME_OF_DELIVERY  = strHoraATAWH   'TIME OF DELIVERY IN SEM
                                       'Concatenado de todos los comentarios
                                       'strComentarioATAWH
                                       'StrPREMARKS           = ""             'REMARKS
                                       'StrPMODALIDAD         = StrModalidad   'MODALIDAD
                                       'StrPWEEK              = ""             'WEEK
                                       'StrPNUM_INVOICE       = strCuentaGastos    'NUM. INVOICE CUSTOM
                                       'StrPDATE_OF_INVOICE   = strFecCuentaGastos 'DATE OF INVOICE CUSTOM
                                       'agregarfilaHTML  StrPOTD,StrPITTS,StrPBL,StrPCONTAINER,StrPP_O,StrPPORT_OF_LOADING,StrPPORT_OF_DISCHARGE,StrPSHIPPING_LINE,StrPVESSEL,StrPIMPORT_DOCUMENT,StrPPROVEEDOR,StrPINVOICE,StrPMODEL,StrPDESCRIPTION,StrPDESCRIPTION_CODE,StrPQTY,StrPETD_LOAD,StrPETA_PORT,StrPATA_PORT,StrPNUMS_SERIE,StrPCERT_NOM,REVALIDACION,StrPRESQUEST_DUTIES,StrPAMOUNT_OF_DUTIES,StrPPREVIO,StrPDATE_OF_CUSTOM,StrPATD_RAIL,StrPETA_CP,StrPATA_CP,StrPETA_WH,StrPATA_WH,StrPTIME_OF_DELIVERY,StrPREMARKS,StrPMODALIDAD,StrPWEEK,StrPNUM_INVOICE,StrPDATE_OF_INVOICE
                                       'agregarfilaHTML (StrPOTD,StrPITTS,StrPBL,StrPCONTAINER,StrPP_O,StrPPORT_OF_LOADING,StrPPORT_OF_DISCHARGE,StrPSHIPPING_LINE,StrPVESSEL,StrPIMPORT_DOCUMENT,StrPINVOICE,StrPMODEL,StrPDESCRIPTION,StrPDESCRIPTION_CODE,StrPQTY,StrPETD_LOAD,StrPETA_PORT,StrPATA_PORT,StrPNUMS_SERIE,StrPCERT_NOM,StrPREVALIDACION,StrPRESQUEST_DUTIES,StrPAMOUNT_OF_DUTIES,StrPPREVIO,StrPDATE_OF_CUSTOM,StrPATD_RAIL,StrPETA_CP,StrPATA_CP,StrPETA_WH,StrPATA_WH,StrPTIME_OF_DELIVERY,StrPREMARKS,StrPMODALIDAD,StrPWEEK,StrPNUM_INVOICE,StrPDATE_OF_INVOICE)

                                       'strFechaATAWH
                                       '


                                       StrREFERENCIA = StrRefer
                                       StrPGUIA_MASTER	         = strGuiaMaster      'GUIA MASTER


                                       StrPFORWARDER_AIR_LINE    = strForwarder       'FORWARDER Y/O AIR  LINE
                                       'StrAdu = RsRep.Fields.Item("PORT_DISCHARGE").Value  ' aduana

                                       if RsRep.Fields.Item("PORT_DISCHARGE").Value <> "" and RsRep.Fields.Item("PORT_DISCHARGE").Value = "200" then
                                         StrPGUIA_HOUSE	           = strNumConte 'GUIA HOUSE - para pantaco tomar el contenedor
                                       else
                                         StrPGUIA_HOUSE	           = strGuiaMasterHouse 'GUIA HOUSE
                                       end if


                                       strADUDESPACHO = ""
                                         StrAdutmp = RsRep.Fields.Item("PORT_DISCHARGE").Value
                                         if ltrim(StrAdutmp)="430" then
                                            strADUDESPACHO = StrAdutmp&"-VERACRUZ" 'aduana aduana de despacho
                                         else
                                           if ltrim(StrAdutmp)="160" then
                                              strADUDESPACHO = StrAdutmp&"-MANZANILLO" 'aduana aduana de despacho
                                           else
                                              if ltrim(StrAdutmp)="200" or ltrim(StrAdu)="202" or ltrim(StrAdu)="470" then
                                                 strADUDESPACHO = StrAdutmp&"-PANTACO" 'aduana aduana de despacho
                                              else
                                                 if ltrim(StrAdutmp)="380" or ltrim(StrAdu)="810" then
                                                    strADUDESPACHO = StrAdutmp&"-TAMPICO" 'aduana aduana de despacho
                                                 else
                                                    if ltrim(StrAdutmp)="510" then
                                                       strADUDESPACHO = StrAdutmp&"-LAZARO CARDENAS" 'aduana aduana de despacho
                                                    else
                                                       if ltrim(StrAdutmp)="470" then
                                                          strADUDESPACHO = StrAdutmp&"-AEROPUERTO" 'aduana aduana de despacho
                                                       end if
                                                    end if
                                                 end if
                                              end if
                                           end if
                                         end if

                                         StrCUSTOM_OF_DISPATCH = ""
                                         StrAdutmp = RsRep.Fields.Item("ADUDES01").Value
                                         if ltrim(StrAdutmp)="430" then
                                            StrCUSTOM_OF_DISPATCH = StrAdutmp&"-VERACRUZ" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                         else
                                           if ltrim(StrAdutmp)="160" then
                                              StrCUSTOM_OF_DISPATCH = StrAdutmp&"-MANZANILLO" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                           else
                                              if ltrim(StrAdutmp)="200" or ltrim(StrAdu)="202" or ltrim(StrAdu)="470" then
                                                 StrCUSTOM_OF_DISPATCH = StrAdutmp&"-PANTACO" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                              else
                                                 if ltrim(StrAdutmp)="380" or ltrim(StrAdu)="810" then
                                                    StrCUSTOM_OF_DISPATCH = StrAdutmp&"-TAMPICO" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                                 else
                                                    if ltrim(StrAdutmp)="510" then
                                                       StrCUSTOM_OF_DISPATCH = StrAdutmp&"-LAZARO CARDENAS" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                                    else
                                                       if ltrim(StrAdutmp)="470" then
                                                          StrCUSTOM_OF_DISPATCH = StrAdutmp&"-AEROPUERTO" 'aduana de destino (en la que llega la mercancia directo de Origen)
                                                       end if
                                                    end if
                                                 end if
                                              end if
                                           end if
                                         end if
                                       'StrCUSTOM_OF_DISPATCH = RsRep.Fields.Item("PORT_DISCHARGE").Value  ' aduana


                                       if RsRep.Fields.Item("PORT_LOADING").Value <> "" then 'Puerto
                                          if RsRep.Fields.Item("VESSEL_LOADING").Value <> "" then 'Pais
                                             StrPAEROPUERTO_SALIDA   = RsRep.Fields.Item("PORT_LOADING").Value&","&RsRep.Fields.Item("VESSEL_LOADING").Value 'PORT OF LOADING
                                          else
                                             StrPAEROPUERTO_SALIDA   = RsRep.Fields.Item("PORT_LOADING").Value 'PORT OF LOADING
                                          end if
                                       else
                                          if RsRep.Fields.Item("VESSEL_LOADING").Value <> "" then
                                             ' StrPAEROPUERTO_SALIDA   = RsRep.Fields.Item("VESSEL_LOADING").Value 'PORT OF LOADING
                                             ' no dejar solo el pais
                                             StrPAEROPUERTO_SALIDA   = "" 'PORT OF LOADING
                                          else
                                             StrPAEROPUERTO_SALIDA   = "" 'PORT OF LOADING
                                          end if
                                       end if

                                       'StrPAEROPUERTO_SALIDA	   = RsRep.Fields.Item("PORT_LOADING").Value&","&RsRep.Fields.Item("VESSEL_LOADING").Value  'AEROPUERTO DE SALIDA

                                       StrPNOTIFICACION_GUIA	   = strGUANotificacion 'NOTIFICACION DE GUIA
                                       ' SE CAMBIO LA FECHA DE NOTIFICACION DE GUIA
                                       ' POR LA FECHA DE ALTA DE LA REFERENCIA EN EL ZEGO
                                       ' SOLICITADO POR JOSE LUIS CHANG 04-10-2007
                                       ' if isdate(strGUA) then
                                       '    StrPFECHA_NOTIFICACION 	 = YEAR(strGUA ) & Pd(Month(strGUA ),2) & Pd(DAY(strGUA ),2)  'FECHA DE NOTIFICACION
                                       ' else
                                       '    StrPFECHA_NOTIFICACION = strGUA 'FECHA DE NOTIFICACION
                                       ' end if
                                       StrPFECHA_NOTIFICACION = formatofechaNum(RsRep.Fields.Item("FecITTS").Value) 'ASIGNADO ITTS


                                       'StrPFECHA_NOTIFICACION 	 = YEAR(strGUA ) & Pd(Month(strGUA ),2) & Pd(DAY(strGUA ),2)  'FECHA DE NOTIFICACION
                                       'StrPFECHA_NOTIFICACION 	 = cStr(year(strGUA)) & cStr(month(strGUA)) & cStr(day(strGUA)) strGUA 'FECHA DE NOTIFICACION
                                       'StrPFECHA_NOTIFICACION 	 = strGUA 'FECHA DE NOTIFICACION

                                       StrPVESSEL                = strvessel    'VESSEL
                                       StrPIMPORT_DOCUMENT	     = RsRep.Fields.Item("IMPORT_DOCUMENT").Value 'IMPORT DOCUMENT
                                       StrPPROVEEDOR	           = strProveedor 'PROVEEDOR
                                       'StrPINVOICE	             = RsRep.Fields.Item("FACTURAS").Value 'INVOICE
                                       StrPINVOICE	             = StrINVOICE
                                       StrPP_O                   = strPO_Pedido  'P/O
                                       StrPMODEL 	               = strModelo    'MODEL
                                       StrPDESCRIPCION_COMERCIAL = strDescMerc  'DESCRIPCION COMERCIAL
                                       StrPDESCRIPTION_CODE	     = strDescCode  'DESCRIPTION CODE
                                       StrPQTY	                 = strQTY       'QTY
                                       StrPINCOTERMS	           = strIncoterms 'INCOTERMS

                                       '*************************************************************
                                       '***                Vamos por los remarks                  ***
                                       '*************************************************************
                                       'variables para los Remarks
                                       rmkEtdLoad    = "" 'rmk para ETDLOAD
                                       rmkATAPORT    = "" 'rmk para ATAPORT
                                       rmkDSP        = "" 'rmk para DESPACHO

                                       diaRmkEtdLoad  = 0 'rmk para ETDLOAD
                                       diaRmkATAPORT  = 0 'rmk para ATAPORT
                                       diaRmkDSP      = 0 'rmk para DESPACHO

                                       tipoRmkEtdLoad  = 1 'rmk para ETDLOAD
                                       tipoRmkATAPORT  = 1 'rmk para ATAPORT
                                       tipoRmkDSP      = 1 'rmk para DESPACHO

                                       descRmkEtdLoad  = "" 'Descripcion del rmk para ETDLOAD
                                       descRmkATAPORT  = "" 'Descripcion del rmk para ATAPORT
                                       descRmkDSP      = "" 'Descripcion del rmk para DESPACHO

                                       strLastRMKtmp = "" ' El ultimo Remark en el que se encuentre

                                       Set RsRmk = Server.CreateObject("ADODB.Recordset")
                                       RsRmk.ActiveConnection = MM_EXTRANET_STRING_STATUS
                                       strSqlrmk = " SELECT c_refer           as referencia, " & _
                                                   "        c_conte           as contenedor, " & _
                                                   "        c_desc            as remark,     " & _
                                                   "        c01rmrks.n_cvermk as claveint,   " & _
                                                   "        d_abrev           as etapa,      " & _
                                                   "        c_cvefor          as clavefor,   " & _
                                                   "        n_dias            as dias,       " & _
                                                   "        n_tipodia         as tipodia     " & _
                                                   " FROM d01rmrks, c01rmrks, etaps          " & _
                                                   " where d01rmrks.n_cvermk = c01rmrks.n_cvermk  " & _
                                                   "       and c01rmrks.n_etapa = etaps.n_etapa " & _
                                                   "       and status = 'A'  " & _
                                                   "       and c_refer = '" & ltrim(StrRefer)    & "' " & _
                                                   "       and c_conte = '" & ltrim(strNumConte) & "' "
                                       'Response.Write(strSqlrmk)
                                       'Response.End
                                       RsRmk.Source = strSqlrmk
                                       RsRmk.CursorType = 0
                                       RsRmk.CursorLocation = 2
                                       RsRmk.LockType = 1
                                       RsRmk.Open()
                                       if not RsRmk.eof then
                                           While NOT RsRmk.EOF
                                              if RsRmk.Fields.Item("etapa").Value="ETDLOAD" then ' RMK de salida de origen
                                                 if RsRmk.Fields.Item("dias").Value > diaRmkEtdLoad then
                                                    rmkEtdLoad     = RsRmk.Fields.Item("clavefor").Value
                                                    diaRmkEtdLoad  = RsRmk.Fields.Item("dias").Value
                                                    tipoRmkEtdLoad = RsRmk.Fields.Item("tipodia").Value
                                                    descRmkEtdLoad = RsRmk.Fields.Item("remark").Value
                                                 end if
                                              else
                                                  if RsRmk.Fields.Item("etapa").Value="ATAPORT" then ' RMK de llegada a puerto
                                                      if RsRmk.Fields.Item("dias").Value > diaRmkATAPORT then
                                                         rmkATAPORT     = RsRmk.Fields.Item("clavefor").Value
                                                         diaRmkATAPORT  = RsRmk.Fields.Item("dias").Value
                                                         tipoRmkATAPORT = RsRmk.Fields.Item("tipodia").Value
                                                         descRmkATAPORT = RsRmk.Fields.Item("remark").Value
                                                      end if
                                                  else
                                                      if RsRmk.Fields.Item("etapa").Value="DSP" then ' RMK de despacho
                                                         if RsRmk.Fields.Item("dias").Value > diaRmkDSP then
                                                            rmkDSP     = RsRmk.Fields.Item("clavefor").Value
                                                            diaRmkDSP  = RsRmk.Fields.Item("dias").Value
                                                            tipoRmkDSP = RsRmk.Fields.Item("tipodia").Value
                                                            descRmkDSP = RsRmk.Fields.Item("remark").Value
                                                         end if
                                                      end if
                                                  end if
                                              end if
                                           RsRmk.movenext
                                           Wend
                                       end if
                                       RsRmk.close
                                       set RsRmk = Nothing


                                       if rmkDSP <> "" then
                                         strLastRMKtmp = descRmkDSP
                                       else
                                          if rmkATAPORT <> "" then
                                             strLastRMKtmp = descRmkATAPORT
                                          else
                                             if rmkEtdLoad <> "" then
                                                strLastRMKtmp = descRmkEtdLoad
                                             end if
                                          end if
                                       end if

                                       '**************************************************************************************
                                       '**************************************************************************************

                                       if isdate( StrNUMSERIE ) then
                                          StrPSERIAL_NUMBER 	   = YEAR( StrNUMSERIE ) & Pd(Month( StrNUMSERIE ),2) & Pd(DAY( StrNUMSERIE ),2)  'SERIAL NUMBER
                                       else
                                          StrPSERIAL_NUMBER	     = StrNUMSERIE  'SERIAL NUMBER
                                       end if
                                       'StrPSERIAL_NUMBER	       = StrNUMSERIE  'SERIAL NUMBER
                                       StrPCERT_NOM          = strCERTNOM         'CERT. NOM

                                       if isdate(RsRep.Fields.Item("feorig01").Value) then
                                          StrPORIGIN_ETD 	 = YEAR( RsRep.Fields.Item("feorig01").Value ) & Pd(Month(RsRep.Fields.Item("feorig01").Value ),2) & Pd(DAY(RsRep.Fields.Item("feorig01").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPORIGIN_ETD = RsRep.Fields.Item("feorig01").Value 'FECHA DE NOTIFICACION
                                       end if
                                       'StrPORIGIN_ETD	           = RsRep.Fields.Item("feorig01").Value  'ORIGIN   ETD

                                       'Response.End

                                       if StdEtdLoad > 0 then
                                          if RsRep.Fields.Item("feorig01").Value  <> ""  then
                                             StrPETA_LAX = formatofechaNum(DateAdd("d",diaRmkEtdLoad,  DateAdd("d",StdEtdLoad, RsRep.Fields.Item("feorig01").Value )) ) ' Calculamos ETA PORT apartir de la fecha de salida de origen
                                          end if
                                       else
                                          if isdate(RsRep.Fields.Item("etalax01").Value) then
                                             StrPETA_LAX 	 = YEAR( RsRep.Fields.Item("etalax01").Value ) & Pd(Month(RsRep.Fields.Item("etalax01").Value ),2) & Pd(DAY(RsRep.Fields.Item("etalax01").Value ),2)  'FECHA DE NOTIFICACION
                                          else
                                             StrPETA_LAX = formatofechaNum(RsRep.Fields.Item("etalax01").Value) 'FECHA DE NOTIFICACION
                                          end if
                                       end if

                                       'if isdate(RsRep.Fields.Item("etalax01").Value) then
                                       '   StrPETA_LAX 	 = YEAR( RsRep.Fields.Item("etalax01").Value ) & Pd(Month(RsRep.Fields.Item("etalax01").Value ),2) & Pd(DAY(RsRep.Fields.Item("etalax01").Value ),2)  'FECHA DE NOTIFICACION
                                       'else
                                       '   StrPETA_LAX = RsRep.Fields.Item("etalax01").Value 'FECHA DE NOTIFICACION
                                       'end if
                                       'StrPETA_LAX    	         = RsRep.Fields.Item("etalax01").Value  'ETA  LAX./

                                       ' ETA LAX va va ser mi pivote para DAI
                                       'StrPETA_CUSTOM_CLEARANCE = StrPETA_LAX + 2
                                       'StrPETA_C_P              = StrPETA_CUSTOM_CLEARANCE + 3
                                       'StrPETA_W_H              = StrPETA_C_P + 1
                                       'Response.Write("eta lax")
                                       'Response.Write(StrPETA_LAX)

                                       'hay veces que capturan la fecha de entrada antes de que haya atracado el buque
                                       'para delantar trabajo, por lo tanto hay que validar que la fecha de entrada
                                       'sea mayor o igual al día de hoy, en caso contrario no desplegarla.
                                       '************************************************************************
                                       DFechEntAux = RsRep.Fields.Item("fecent01").Value
                                       if isdate(DFechEntAux) then
                                          if DFechEntAux > date() then
                                             DFechEntAux = ""
                                          end if
                                       end if
                                       '************************************************************************

                                       'Response.End
                                       'StrETA_CUSTOM_CLEARANCE = SumarDiasSinFinSemana( RsRep.Fields.Item("etalax01").Value , 1)
                                       if isdate(DFechEntAux) then
                                         'StrETA_CUSTOM_CLEARANCE = SumarDiasSinFinSemana( DFechEntAux , 1)
                                         StrETA_CUSTOM_CLEARANCE = SumarDias( SumarDias( DFechEntAux , StdATAPORTDSP,tipoStdATAPORTDSP) , diaRmkATAPORT,tipoRmkATAPORT)
                                       else
                                         'StrETA_CUSTOM_CLEARANCE = SumarDiasSinFinSemana( RsRep.Fields.Item("etalax01").Value , 1)
                                         StrETA_CUSTOM_CLEARANCE = SumarDias( SumarDias( RsRep.Fields.Item("etalax01").Value , StdATAPORTDSP,tipoStdATAPORTDSP) , diaRmkATAPORT,tipoRmkATAPORT)
                                       end if

                                       StrColorfila = 1
                                       StrETA_C_P = "N/A"

                                       if isdate( StrETA_CUSTOM_CLEARANCE ) then
                                         if isdate( RsRep.Fields.Item("DATE_CUSTOM").Value ) then
                                           IndFila = DateDiff("d",StrETA_CUSTOM_CLEARANCE , RsRep.Fields.Item("DATE_CUSTOM").Value )
                                           if IndFila = 0 then
                                              StrColorfila = 1
                                              'StrETA_W_H_AUX = SumarDiasSinFinSemana(StrETA_CUSTOM_CLEARANCE , 1)
                                              StrETA_W_H_AUX = SumarDias( SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                           else
                                              'StrETA_W_H_AUX = SumarDiasSinFinSemana(RsRep.Fields.Item("DATE_CUSTOM").Value , 1) 'si las fechas estimada y la real no coinciden ponemos como eta la ata y de ahi recalculamos las siguientes
                                              StrETA_W_H_AUX = SumarDias( SumarDias( RsRep.Fields.Item("DATE_CUSTOM").Value , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                              if IndFila < 0 then
                                                  StrColorfila = 2
                                              else
                                                  StrColorfila = 3
                                              end if
                                           end if
                                         else
                                            'StrETA_W_H_AUX = SumarDiasSinFinSemana(StrETA_CUSTOM_CLEARANCE , 1)
                                            StrETA_W_H_AUX = SumarDias( SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                            IndFila = DateDiff("d", StrETA_CUSTOM_CLEARANCE , DATE() )
                                            if IndFila > 0 then
                                                 StrColorfila = 3
                                            end if
                                         end if
                                       else
                                          if isdate( RsRep.Fields.Item("DATE_CUSTOM").Value ) then
                                             'StrETA_W_H_AUX = SumarDiasSinFinSemana(RsRep.Fields.Item("DATE_CUSTOM").Value , 1)
                                             StrETA_W_H_AUX = SumarDias( SumarDias( RsRep.Fields.Item("DATE_CUSTOM").Value , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                             'StrColorfila = 1
                                          else
                                             'StrETA_W_H_AUX = SumarDiasSinFinSemana(StrETA_CUSTOM_CLEARANCE , 1)
                                             StrETA_W_H_AUX = SumarDias( SumarDias( StrETA_CUSTOM_CLEARANCE , StdDSPWH, tipoStdDSPWH ) , diaRmkDSP, tipoRmkDSP )
                                          end if
                                       end if

                                       if StrETA_W_H_AUX then
                                         if isdate(strFechaATAWH ) then
                                             'IndFila = DateDiff("d",StrETA_W_H , strFechaATAWH )
                                             IndFila = DateDiff("d",StrETA_W_H_AUX , strFechaATAWH )
                                             if IndFila <> 0 then
                                                if StrColorfila = 1 then
                                                   if IndFila < 0 then
                                                      StrColorfila = 2
                                                   else
                                                      StrColorfila = 3
                                                   end if
                                                end if
                                             end if
                                         else
                                            IndFila = DateDiff("d", StrETA_W_H_AUX , DATE() )
                                            if IndFila > 0 then
                                               StrColorfila = 3
                                            end if
                                         end if
                                       end if

                                       'Response.Write("StrPETA_CUSTOM_CLEARANCE")
                                       'Response.Write(StrPETA_CUSTOM_CLEARANCE)
                                       'Response.End

                                       StrPETA_CUSTOM_CLEARANCE = formatofechaNum(StrETA_CUSTOM_CLEARANCE)
                                       StrPETA_C_P              = formatofechaNum(StrETA_C_P)
                                       if isdate( strETAW_H ) then
                                          StrPETA_W_H 	 = YEAR( strETAW_H ) & Pd(Month( strETAW_H ),2) & Pd(DAY( strETAW_H ),2)  'FECHA DE NOTIFICACION
                                          StrETA_W_H_AUX = StrPETA_WH
                                       else
                                          StrPETA_W_H              = formatofechaNum(StrETA_W_H_AUX)
                                       end if
                                       'StrPATA_CP               = formatofechaNum(strATAC_P)       'ATA C./P.
                                       StrPATA_CP  = "N/A"       'ATA C./P.

                                       '********************************************************************************
                                       '********************************************************************************

                                       if isdate( DFechEntAux ) then
                                          StrPATA_CUSTOM = YEAR( DFechEntAux ) & Pd(Month( DFechEntAux ),2) & Pd(DAY( DFechEntAux ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPATA_CUSTOM = DFechEntAux 'FECHA DE NOTIFICACION
                                       end if

                                       'if isdate(RsRep.Fields.Item("fecent01").Value) then
                                       '   StrPATA_CUSTOM 	 = YEAR( RsRep.Fields.Item("fecent01").Value ) & Pd(Month( RsRep.Fields.Item("fecent01").Value ),2) & Pd(DAY( RsRep.Fields.Item("fecent01").Value ),2)  'FECHA DE NOTIFICACION
                                       'else
                                       '   StrPATA_CUSTOM = RsRep.Fields.Item("fecent01").Value 'FECHA DE NOTIFICACION
                                       'end if
                                       ''StrPATA_CUSTOM	           = RsRep.Fields.Item("fecent01").Value  'ATA CUSTOM

                                       if isdate( RsRep.Fields.Item("RESQUEST_DUTIES").Value ) then
                                          StrPRESQUEST_DUTIES 	 = YEAR( RsRep.Fields.Item("RESQUEST_DUTIES").Value ) & Pd(Month( RsRep.Fields.Item("RESQUEST_DUTIES").Value ),2) & Pd(DAY( RsRep.Fields.Item("RESQUEST_DUTIES").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPRESQUEST_DUTIES = RsRep.Fields.Item("RESQUEST_DUTIES").Value 'FECHA DE NOTIFICACION
                                       end if
                                       'StrPRESQUEST_DUTIES	     = RsRep.Fields.Item("RESQUEST_DUTIES").Value 'RESQUEST DUTIES

                                       if isdate( RsRep.Fields.Item("REVALIDACION").Value ) then
                                          StrPFECHA_DE_REVALIDACION 	 = YEAR( RsRep.Fields.Item("REVALIDACION").Value ) & Pd(Month( RsRep.Fields.Item("REVALIDACION").Value ),2) & Pd(DAY( RsRep.Fields.Item("REVALIDACION").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPFECHA_DE_REVALIDACION = RsRep.Fields.Item("REVALIDACION").Value 'FECHA DE NOTIFICACION
                                       end if
                                       'StrPFECHA_DE_REVALIDACION = RsRep.Fields.Item("REVALIDACION").Value    'FECHA DE  REVALIDACION

                                       if isdate( RsRep.Fields.Item("PREVIO").Value ) then
                                          StrPFECHA_DE_PREVIO 	 = YEAR( RsRep.Fields.Item("PREVIO").Value ) & Pd(Month( RsRep.Fields.Item("PREVIO").Value ),2) & Pd(DAY( RsRep.Fields.Item("PREVIO").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPFECHA_DE_PREVIO = RsRep.Fields.Item("PREVIO").Value 'FECHA DE NOTIFICACION
                                       end if
                                       'StrPFECHA_DE_PREVIO	     = RsRep.Fields.Item("PREVIO").Value          'FECHA DE PREVIO
                                       if isdate( RsRep.Fields.Item("DATE_CUSTOM").Value ) then
                                          StrPDATE_OF_CLEARANCE 	 = YEAR( RsRep.Fields.Item("DATE_CUSTOM").Value ) & Pd(Month( RsRep.Fields.Item("DATE_CUSTOM").Value ),2) & Pd(DAY( RsRep.Fields.Item("DATE_CUSTOM").Value ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPDATE_OF_CLEARANCE = RsRep.Fields.Item("DATE_CUSTOM").Value 'FECHA DE NOTIFICACION
                                       end if
                                       'StrPDATE_OF_CLEARANCE	   = RsRep.Fields.Item("DATE_CUSTOM").Value     'DATE OF CLEARANCE
                                       if isdate( strFechaATAWH ) then
                                          StrPATA_WH 	 = YEAR( strFechaATAWH ) & Pd(Month( strFechaATAWH ),2) & Pd(DAY( strFechaATAWH ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPATA_WH = strFechaATAWH 'FECHA DE NOTIFICACION
                                       end if
                                       'StrPATA_WH	               = strFechaATAWH 'ATA W-H

                                       'if isdate( strHoraATAWH ) then
                                       '   StrPTIMEOFDELIVERY 	 = YEAR( strHoraATAWH ) & Pd(Month( strHoraATAWH ),2) & Pd(DAY( strHoraATAWH ),2)  'FECHA DE NOTIFICACION
                                       'else
                                       '   StrPTIMEOFDELIVERY = strHoraATAWH 'FECHA DE NOTIFICACION
                                       'end if

                                       strATASPL            = strTimeSLP
                                       StrPTIMEOFDELIVERY   = strHoraATAWH  'TIME OF DELIVERY IN SEM

                                       'if strComentarioATAWH <> "" then
                                       '  strObservaciones = strObservaciones&"<BR>"& strComentarioATAWH
                                       'end if
                                       'if strComentarioATAC_P <> "" then
                                       '  strObservaciones = strObservaciones&"<BR>"& strComentarioATAC_P
                                       'end if
                                       'if strComentarioETAW_H <> "" then
                                       '  strObservaciones = strObservaciones&"<BR>"& strComentarioETAW_H
                                       'end if

                                       if strComentarioATAWH <> "" AND ltrim(strComentarioATAWH) <> "" then
                                         strObservaciones = strObservaciones&" ; "& strComentarioATAWH
                                       end if
                                       if strComentarioATAC_P <> "" and ltrim(strComentarioATAC_P) <> "" then
                                         strObservaciones = strObservaciones&" ; "& strComentarioATAC_P
                                       end if
                                       if strComentarioETAW_H <> "" and ltrim(strComentarioETAW_H) <> "" then
                                         strObservaciones = strObservaciones&" ; "& strComentarioETAW_H
                                       end if
                                       if strComentarioATASPLTMP <> "" and ltrim(strComentarioATASPLTMP) <> "" then
                                         strObservaciones = strObservaciones&" ; "& strComentarioATASPLTMP
                                       end if

                                       StrPREMARKS = strObservaciones  'REMARKS


                                       ' SEMANA DEL AÑO DE LA FECHA DE GENEREACIONS DEL REPORTE (NOW)
                                       'DCustomClear = ( RsRep.Fields.Item("DATE_CUSTOM").Value )
                                       'if isdate(strFechaATAWH) then
                                       '   if not isempty(strFechaATAWH) then
                                       '      numeroDiasAnio = dateDiff("d",CDate("01/01/"&Datepart("yyyy", strFechaATAWH )), strFechaATAWH )
                                       '      numeroDiasAnio =    int(numeroDiasAnio/7)+1
                                       '    else
                                       '      numeroDiasAnio = 0
                                       '    end if
                                       'else
                                       '   numeroDiasAnio = 0
                                       'end if



                                       if isdate(strFechaATAWH) then
                                          if not isempty(strFechaATAWH) then
                                             numeroDiasAnio = dateDiff("d",CDate("01/01/"&Datepart("yyyy", strFechaATAWH )), strFechaATAWH )
                                             numeroDiasAnio =    int(numeroDiasAnio/7)+1
                                           else
                                             numeroDiasAnio = 0
                                           end if
                                       else
                                          if isdate(StrETA_W_H_AUX) then
                                             numeroDiasAnio = dateDiff("d",CDate("01/01/"&Datepart("yyyy", StrETA_W_H_AUX )), StrETA_W_H_AUX )
                                             numeroDiasAnio =    int(numeroDiasAnio/7)+1
                                          else
                                             numeroDiasAnio = 0
                                          end if
                                       end if




                                       'DCustomClear = ( RsRep.Fields.Item("DATE_CUSTOM").Value )
                                       'if isdate(DCustomClear) then
                                       '   if not isempty(DCustomClear) then
                                       '      numeroDiasAnio = dateDiff("d",CDate("01/01/"&Datepart("yyyy",  DCustomClear  )), DCustomClear )
                                       '      numeroDiasAnio =    int(numeroDiasAnio/7)+1
                                       '    else
                                       '      numeroDiasAnio = 0
                                       '    end if
                                       'else
                                       '   numeroDiasAnio = 0
                                       'end if
                                       'numeroDiasAnio = dateDiff("d",CDate("01/01/"&Datepart("yyyy",  Date() )), Date() )

                                       StrPMODALIDAD         = StrModalidad   'MODALIDAD
                                       StrPWEEK	             = numeroDiasAnio   'WEEK

                                       ' SEMANA DEL AÑO DE LA FECHA DE GENEREACIONS DEL REPORTE (NOW)
                                       'numeroDiasAnio = dateDiff("d",CDate("01/01/"&Datepart("yyyy", Date)),Date())
                                       'numeroDiasAnio & ", " Cstr(numeroDiasAnio/7)
                                       'StrPWEEK	                 = Cstr(int(numeroDiasAnio/7)+1)   'WEEK

                                       StrPAMOUNTOFDUTIES	       = strImpuestos 'AMOUNT OF DUTIES
                                       StrPNUM_INVOICECUSTOM	   = strCuentaGastos 'NUM. INVOICE CUSTOM
                                       if isdate( strFecCuentaGastos ) then
                                          StrPDATEINVOICECUSTOM 	 = YEAR( strFecCuentaGastos ) & Pd(Month( strFecCuentaGastos ),2) & Pd(DAY( strFecCuentaGastos ),2)  'FECHA DE NOTIFICACION
                                       else
                                          StrPDATEINVOICECUSTOM = strFecCuentaGastos 'FECHA DE NOTIFICACION
                                       end if
                                       'StrPDATEINVOICECUSTOM     = strFecCuentaGastos 'DATE OF INVOICE CUSTOM

                                       'Response.Write("Ata WH")
                                       'Response.Write(strFechaATAWH)
                                       'Response.End



                                       'if (strFechaATAWH) <> "" and isdate(strFechaATAWH) then
                                       '    intoTD = DiasTrimFinSemana(RsRep.Fields.Item("fecent01").Value,strFechaATAWH)
                                       'else
                                       '    if isdate(StrETA_W_H_AUX) then
                                       '       intoTD = DiasTrimFinSemana(RsRep.Fields.Item("fecent01").Value, StrETA_W_H_AUX )
                                       '    else
                                       '       intoTD = 0
                                       '    end if
                                       'end if

                                       if isdate(strFechaATAWH) then
                                           if isdate(DFechEntAux) then
                                              intoTD = DiasTrimFinSemana( DFechEntAux ,strFechaATAWH )
                                           else
                                              'intoTD = 0
                                              if isdate( RsRep.Fields.Item("etalax01").Value ) then
                                                intoTD = DiasTrimFinSemana( RsRep.Fields.Item("etalax01").Value , strFechaATAWH )
                                              else
                                                intoTD = 0
                                              end if
                                           end if
                                       else
                                           if isdate(StrETA_W_H_AUX) then
                                              if isdate(DFechEntAux) then
                                                 intoTD = DiasTrimFinSemana( DFechEntAux , StrETA_W_H_AUX )
                                              else
                                                 'intoTD = 0
                                                 if isdate( RsRep.Fields.Item("etalax01").Value ) then
                                                   intoTD = DiasTrimFinSemana( RsRep.Fields.Item("etalax01").Value , StrETA_W_H_AUX )
                                                 else
                                                   intoTD = 0
                                                 end if
                                              end if
                                           else
                                              intoTD = 0
                                           end if
                                       end if

                                       StrPOTD2                  = intoTD 'OTD2

                                       'Response.Write("Diferencia")
                                       'Response.Write(intoTD)
                                       'Response.End

                                       strStatusTmp  = "" ' Exactamnete en donde se encuentra la mercancia
                                       strKPISTTmp  = "" ' Para saber si viene en tiempo o retrasado
                                       '*SI MODALIDAD ES “FERROVIARIO” O “CARRETERO” Y SI EXISTE ATA W/H
                                       '   ATA/W/H- ATA PORT <= 8 ES “ON TIME”     SINO ES “DELAY”
                                       ' *SI NO EXISTE ATA W/H PERO EXISTE ATA PORT/CUSTOM
                                       '   ENTONCES ETA W/H – ATA PORT/CUSTOM <=8  ES “ON TIME” SINO ES “DELAY”
                                       ' * SI NO EXISTE ATA W/H Y  ATAPORT/CUSTOM ES “ON TIME”

                                       if intoTD <= 2 then
                                         strKPISTTmp = "ON TIME"
                                       else
                                         strKPISTTmp = "DELAY"
                                       end if

                                       if strFechaATAWH <> "" then
                                          strStatusTmp = "SEM"
                                       else
                                          if StrPDATE_OF_CLEARANCE <> "" then
                                             strStatusTmp = "ADUANA"
                                          else
                                             if DFechEntAux <> "" then
                                                strStatusTmp = "AEROPUERTO"
                                             else
                                                if RsRep.Fields.Item("feorig01").Value <> "" then
                                                  strStatusTmp = "TRANSITO AEREO"
                                                end if
                                             end if
                                          end if
                                       end if
                                        'SI EXISTE ATA W/H ESCRIBE “SEM”
                                        'SI NO EXISTE ATA W/H PERO EXISTE ATA C./P. ESCRIBE “COUNTRY/ PANTACO”.
                                        'SI NO EXISTE ATD C./P. PERO EXISTE ATA RAIL ESCRIBE “ TRANSITO FERROVIARIO “
                                        'SI NO EXISTE ATD RAIL  PERO EXISTE DATE OF CLEARENCE ESCRIBE “ ADUANA”
                                        'SI NO EXISTE DATE OF CLEARENCE PERO EXISTE ATA PORT/CUSTOM ESCRIBE  “ PUERTO”
                                        'SI NO EXISTE ATA PORT/CUSTOM  PERO EXISTE ETD LOAD ESCRIBE “TRANSITO MARITIMO “

                                       strRMKATDORIGIN = rmkEtdLoad
                                       strRMKATAPORT   = rmkATAPORT
                                       strRMKDEPACHO   = rmkDSP
                                       strRMKATDRAIL   = rmkRAIL
                                       strRMKCP        = rmkCP
                                       'strATASPL       = strATASPL
                                       strSTATUS       = strStatusTmp
                                       strLASTRMK      = strLastRMKtmp
                                       strKPISTATUS    = strKPISTTmp


                                       'strRMKATDORIGIN = ""
                                       'strRMKATAPORT   = ""
                                       'strRMKDEPACHO   = ""
                                       'strRMKATDRAIL   = ""
                                       'strRMKCP        = ""
                                       ''strATASPL       = ""
                                       'strSTATUS       = ""
                                       'strLASTRMK      = ""
                                       'strKPISTATUS    = ""

                                       'strRMKATDORIGIN, strRMKATAPORT, strRMKDEPACHO, strRMKATDRAIL, strRMKCP, strATASPL, strSTATUS, strLASTRMK, strKPISTATUS
                                       'agregarfilaHTML  StrColorfila, StrReferencia,StrPOTD,StrPITTS,StrPBL,StrPCONTAINER,StrPP_O,StrPPORT_OF_LOADING,StrPPORT_OF_DISCHARGE,StrPSHIPPING_LINE,StrPVESSEL,StrPIMPORT_DOCUMENT,StrPPROVEEDOR,StrPINVOICE,StrPMODEL,StrPDESCRIPTION,StrPDESCRIPTION_CODE,StrPQTY,StrPETD_LOAD,StrPETA_PORT,StrPATA_PORT,StrPNUMS_SERIE,StrPCERT_NOM,StrPREVALIDACION ,StrPRESQUEST_DUTIES,StrPAMOUNT_OF_DUTIES,StrPPREVIO,StrPETA_CUSTOM_CLEARANCE ,StrPDATE_OF_CUSTOM,StrPATD_RAIL,StrPETA_CP,StrPATA_CP,StrPETA_WH,StrPATA_WH,StrPTIME_OF_DELIVERY,StrPREMARKS,StrPMODALIDAD,StrPWEEK,StrPNUM_INVOICE,StrPDATE_OF_INVOICE, strADUDESPACHO, strRMKATDORIGIN, strRMKATAPORT, strRMKDEPACHO, strRMKATDRAIL, strRMKCP, strATASPL, strSTATUS, strLASTRMK, strKPISTATUS
                                       agregarfilaHTML StrColorfila, StrREFERENCIA, StrPOTD2, StrPGUIA_MASTER, StrPGUIA_HOUSE, StrPP_O, StrPFORWARDER_AIR_LINE, StrPAEROPUERTO_SALIDA, StrCUSTOM_OF_DISPATCH, StrPNOTIFICACION_GUIA, StrPFECHA_NOTIFICACION, StrPVESSEL, StrPIMPORT_DOCUMENT, StrPPROVEEDOR, StrPINVOICE, StrPMODEL, StrPDESCRIPCION_COMERCIAL, StrPDESCRIPTION_CODE, StrPQTY, StrPINCOTERMS, StrPSERIAL_NUMBER, StrPCERT_NOM,	StrPORIGIN_ETD, StrPETA_LAX, StrPATA_CUSTOM, StrPRESQUEST_DUTIES, StrPFECHA_DE_REVALIDACION, StrPFECHA_DE_PREVIO, StrPETA_CUSTOM_CLEARANCE, StrPDATE_OF_CLEARANCE, StrPETA_C_P,  StrPATA_CP ,StrPETA_W_H, StrPATA_WH, StrPTIMEOFDELIVERY , StrPREMARKS, StrPMODALIDAD , StrPWEEK, StrPAMOUNTOFDUTIES, StrPNUM_INVOICECUSTOM, StrPDATEINVOICECUSTOM, strADUDESPACHO, strRMKATDORIGIN, strRMKATAPORT, strRMKDEPACHO, strRMKATDRAIL, strRMKCP, strATASPL, strSTATUS, strLASTRMK, strKPISTATUS
                             '*************************************************************




                         end if

                  end if 'if Bolbanrecti = True then

                     RsRep.movenext
                    'Response.Write(strHTML)
                    'Response.End

                    if enproceso( adu_ofi( Session("GAduana") ) ) then
                      banCargaRun=true
                    end if

               Wend

            strHTML = strHTML & "</table>"& chr(13) & chr(10)

            end if


            'DespliegaEncabezado()
            'response.Write(strHTML)
            'Response.End

            if banCargaRun = false then
                RsRep.close
                Set RsRep = Nothing
                'Se pinta todo el HTML formado
                response.Write(strHTML)
                if strHTML = "" then
                   strHTML = "NO EXISTEN REGISTROS"
                   response.Write(strHTML)
                end if
            else

                strHTML = "<table>"
                strHTML = strHTML &  "<tr bgcolor='#1B5296'>"
                strHTML = strHTML &  "      <td colspan='4' class='textForm2'><div align='right'></div></td> "
                strHTML = strHTML &  " </tr> "
                strHTML = strHTML &  " <tr>  "
                strHTML = strHTML &  "    <td colspan='4'><div align='center'></div></td> "
                strHTML = strHTML &  " </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td width='250' rowspan='4' align='center'><img src='http://rkzego.no-ip.org/PortalMySQL/Extranet/ext-Images/computadora_animo.jpg' width='150' height='157'></td> "
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=4 COLOR=red>Espere un momento...</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=5 COLOR=red>La Base de Datos se esta Actualizando</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=5 COLOR=red>Genere este Reporte unos minutos mas tarde</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='3' align='center'><FONT FACE='arial' SIZE=3 COLOR=red>estamos trabajando para brindarle un mejor servicio</FONT></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr>"
                strHTML = strHTML &  "    <td colspan='4'><div align='center'></div></td>"
                strHTML = strHTML &  "  </tr>"
                strHTML = strHTML &  "  <tr bgcolor='#1B5296'><td colspan='4'></td></tr>"
                strHTML = strHTML &  "  </table>"
                response.Write(strHTML)
            end if

          else
             strHTML = "NO EXISTEN REGISTROS"
             response.Write(strHTML)
          end if

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
<%end if

ELSE
%>
     <!--#include virtual="/PortalMySQL/Extranet/ext-Asp/carga_activa.asp" -->
<%

END IF

%>




