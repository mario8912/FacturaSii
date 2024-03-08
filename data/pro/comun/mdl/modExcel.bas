Attribute VB_Name = "modExcel"

Global VaValor As Variant
'211227 StrExtensionExcel
Global StrExtensionExcel As String

Public Function QVersionExcel() As String
          '211227 QVersionExcel
10        On Error GoTo Fallo
          Dim oExcel As excel.Application
20        Set oExcel = New excel.Application
          
          Dim ExcelVersion As Long
30        ExcelVersion = Val(oExcel.Version)
          
40        Select Case ExcelVersion
                 Case 11 'excel 2003
50                QVersionExcel = ".xls"
60            Case 12 'excel 2007
70                QVersionExcel = ".xlsx"
80            Case 14 'excel 2010
90                QVersionExcel = ".xlsx"
100           Case 15 'excel 2013
110               QVersionExcel = ".xlsx"
120       End Select
130       Set oExcel = Nothing
140       On Error GoTo 0
150       Exit Function
Fallo:
160       Set oExcel = Nothing
170       Msg mError, "", Err, Erl
End Function

'===
Public Sub CrearHojaExcelCompras2019(Rs As ADODB.Recordset, Fichero As String)
10        On Error GoTo Fallo
          Dim erroNoCampo As Boolean
          Dim oExcel As excel.Application
          Dim oBook As excel.Workbook
          Dim oSheet As excel.Worksheet

20        Set oExcel = New excel.Application
30        Set oBook = oExcel.Workbooks.Add
40        Set oSheet = oBook.Worksheets(1)
          '15/02/2014 modificada CrearExcelDesatendido para que guarde el fichero segun la version instalada
          'el fichero hay que pasarlo sin la extension
          
          Dim ExcelVersion As Long
50        ExcelVersion = Val(oExcel.Version)
          
60        Select Case ExcelVersion
              Case 11 'excel 2003
70                Fichero = Fichero & ".xls"
80            Case 12 'excel 2007
90                Fichero = Fichero & ".xlsx"
100           Case 14 'excel 2010
110               Fichero = Fichero & ".xlsx"
120       End Select
          Dim N As Integer
130       For N = 0 To Rs.Fields.Count - 1
140           oSheet.Cells(1, N + 1).Value = Rs.Fields(N).Name
150       Next
          
          Dim vFila As Long
          Dim vCol As Long
          Dim dblImporte As Double
          Dim dblSumaLinea As Double
          Dim blnSaltar As Boolean
          
          '170813 blnEsFactura en CrearHojaExcelCompras
          Dim blnEsFactura As Boolean
          Dim dblSumaIVA As Double
          Dim intNumFacturas As Integer
          
160       vFila = 2
170       While Not Rs.EOF
              '170813 blnEsFactura en CrearHojaExcelCompras
180           blnEsFactura = False
              'dblSumaIVA = rs.Fields(10) + rs.Fields(15) + rs.Fields(20)
              

190           dblSumaIVA = Rs.Fields(9) + Rs.Fields(14) + Rs.Fields(19) + Rs.Fields(24)
200           On Error GoTo NoCampo
210           dblSumaIVA = dblSumaIVA + Rs.Fields(29)
220           On errot GoTo Fallo
              'dblSumaIVA = rs.Fields(10) + rs.Fields(15) + rs.Fields(21)
230           If dblSumaIVA <> 0 Then blnEsFactura = True
              'esto es el dni
240           If Rs.Fields(6) = "" Then blnEsFactura = False
              
250           If blnEsFactura Then
260           For N = 0 To Rs.Fields.Count - 1
                  'esto es la suma del iva
                  
                  '230206 manolo
270               dblSumaLinea = Rs.Fields(10) + Rs.Fields(12) + Rs.Fields(15) + Rs.Fields(17) + Rs.Fields(20) + Rs.Fields(22) + Rs.Fields(25) + Rs.Fields(27) + Rs.Fields(30) + Rs.Fields(32)
                  'dblSumaLinea = rs.Fields(10) + rs.Fields(12) + rs.Fields(15) + rs.Fields(17) + rs.Fields(20) + rs.Fields(22)
                  'dblSumaLinea = rs.Fields(10) + rs.Fields(12) + rs.Fields(15) + rs.Fields(17) + rs.Fields(20) + rs.Fields(22)
280               Select Case N
      '===
                      Case 0
                          'idexterno
290                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
300                   Case 1
                          'numfactura
310                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
320                   Case 2
                          'ejercicio
330                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
340                   Case 3
                          'fecha expedicion
350                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
360                   Case 4
                          'fecha recepcion
370                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
380                   Case 5
                          'fecha fiscal
390                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
400                   Case 6
                          'nif
410                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
420                   Case 7
                          'razon social
430                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
440                   Case 8
                          'la base
                          'si el importe es 0 comprobar si  hay valor en alguna cuota
450                       If Rs.Fields(10) = 0 Then
460                           If dblSumaLinea = 0 Then
470                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
480                           End If
490                       Else
500                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
510                       End If
520                   Case 9
                          'el tipo de iva
530                       If Rs.Fields(10) = 0 Then
540                           If dblSumaLinea = 0 Then
550                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
560                           End If
570                       Else
580                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
590                       End If
600                   Case 10
                          'la cuota 1
610                       If Rs.Fields(10) <> 0 Then
620                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
630                       End If
640                   Case 11
                          'tipo RE 1
650                       If Rs.Fields(11) <> 0 Then
660                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
670                       End If
680                   Case 12
                          'cuota re 1
690                       If Rs.Fields(12) <> 0 Then
700                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
710                       End If
720                   Case 13
                          'base 2
730                       If Rs.Fields(13) <> 0 Then
740                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
750                       End If
760                   Case 14
                          'tipo2
770                       If Rs.Fields(N) <> 0 Then
780                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
790                       End If
800                   Case 15
                          'cuota2
810                       If Rs.Fields(N) <> 0 Then
820                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
830                       End If
840                   Case 16
                          'tipo RE2
850                       If Rs.Fields(N) <> 0 Then
860                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
870                       End If
880                   Case 17
                          'cuota RE2
890                       If Rs.Fields(N) <> 0 Then
900                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
910                       End If
920                   Case 18
                          'base 3
930                       If Rs.Fields(N) <> 0 Then
940                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
950                       End If
960                   Case 19
                          'tipo iva 3
970                       If Rs.Fields(N) <> 0 Then
980                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
990                       End If
1000                  Case 20
                          'cuota 3
1010                      If Rs.Fields(N) <> 0 Then
1020                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1030                      End If
1040                  Case 21
                          'tipo re3
1050                      If Rs.Fields(N) <> 0 Then
1060                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1070                      End If
1080                  Case 22
                          'cuota 3
1090                      If Rs.Fields(N) <> 0 Then
1100                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1110                      End If
1120                  Case 23
                          'Base4
1130                      If Rs.Fields(N) <> 0 Then
1140                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1150                      End If
1160                  Case 24
                          'Tipo4
1170                      If Rs.Fields(N) <> 0 Then
1180                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1190                      End If
1200                  Case 25
                          'Cuota4
1210                      If Rs.Fields(N) <> 0 Then
1220                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1230                      End If
1240                  Case 26
                          'TipoRE4
1250                      If Rs.Fields(N) <> 0 Then
1260                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1270                      End If
1280                  Case 27
                          'CuotaRE4
1290                      If Rs.Fields(N) <> 0 Then
1300                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1310                      End If
1320                  Case 28
                          'Base5
1330                      If Rs.Fields(N) <> 0 Then
1340                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1350                      End If
      '====
1360                  Case 29
                          'Tipo5
1370                      If Rs.Fields(N) <> 0 Then
1380                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1390                      End If
1400                  Case 30
                          'Cuota5
1410                      If Rs.Fields(N) <> 0 Then
1420                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1430                      End If
1440                  Case 31
                          'TipoRE5
1450                      If Rs.Fields(N) <> 0 Then
1460                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1470                      End If
1480                  Case 32
                          'CuotaRE5
1490                      If Rs.Fields(N) <> 0 Then
1500                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1510                      End If
1520                  Case 33
                          'idsujeto
1530                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1540                  Case 34
                          'base99
1550                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1560                  Case 35
                          'codcontable
1570                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1580                  Case 36
                          'contrapartida
1590                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1600                  Case 37
                          'irpf
1610                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1620                  Case 38
                          'cuentairpf
1630                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)


      '===

1640              End Select
1650          Next
1660          vFila = vFila + 1
1670          End If
1680          Rs.MoveNext
              'vFila = vFila + 1
1690      Wend
          
1700      intNumFacturas = vFila
1710      If Rs.RecordCount > 0 Then
1720          Rs.MoveFirst
1730      End If
      '===la segunda pasada
1740      While Not Rs.EOF
              '170813 blnEsFactura en CrearHojaExcelCompras
1750          blnEsFactura = False
              
1760          dblSumaIVA = Rs.Fields(9) + Rs.Fields(14) + Rs.Fields(20)
              'dblSumaIVA = rs.Fields(10) + rs.Fields(15) + rs.Fields(20)
              'dblSumaIVA = rs.Fields(11) + rs.Fields(16) + rs.Fields(21)
1770          If dblSumaIVA <> 0 Then blnEsFactura = True
              'esto es el dni
1780          If Rs.Fields(6) = "" Then blnEsFactura = False
              
1790          If Not blnEsFactura Then
1800          For N = 0 To Rs.Fields.Count - 1
                  'esto es la suma del iva
                  '230206 manolo
1810              dblSumaLinea = Rs.Fields(10) + Rs.Fields(12) + Rs.Fields(15) + Rs.Fields(17) + Rs.Fields(20) + Rs.Fields(22)
                  'dblSumaLinea = rs.Fields(10) + rs.Fields(12) + rs.Fields(15) + rs.Fields(17) + rs.Fields(20) + rs.Fields(22)
1820              dblSumaLinea = Rs.Fields(10) + Rs.Fields(12) + Rs.Fields(15) + Rs.Fields(17) + Rs.Fields(20) + Rs.Fields(22) + Rs.Fields(25) + Rs.Fields(27) + Rs.Fields(30) + Rs.Fields(32)
                  
                  
1830              Select Case N
      '===
                      Case 0
                          'idexterno
1840                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1850                  Case 1
                          'numfactura
1860                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1870                  Case 2
                          'ejercicio
1880                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1890                  Case 3
                          'fecha expedicion
1900                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1910                  Case 4
                          'fecha recepcion
1920                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1930                  Case 5
                          'fecha fiscal
1940                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1950                  Case 6
                          'nif
1960                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1970                  Case 7
                          'razon social
1980                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1990                  Case 8
                          'la base
                          'si el importe es 0 comprobar si  hay valor en alguna cuota
2000                      If Rs.Fields(10) = 0 Then
2010                          If dblSumaLinea = 0 Then
2020                              oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2030                          End If
2040                      Else
2050                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2060                      End If
2070                  Case 9
                          'el tipo de iva
2080                      If Rs.Fields(10) = 0 Then
2090                          If dblSumaLinea = 0 Then
2100                              oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2110                          End If
2120                      Else
2130                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2140                      End If
2150                  Case 10
                          'la cuota 1
2160                      If Rs.Fields(10) <> 0 Then
2170                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2180                      End If
2190                  Case 11
                          'tipo RE 1
2200                      If Rs.Fields(11) <> 0 Then
2210                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2220                      End If
2230                  Case 12
                          'cuota re 1
2240                      If Rs.Fields(12) <> 0 Then
2250                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2260                      End If
2270                  Case 13
                          'base 2
2280                      If Rs.Fields(13) <> 0 Then
2290                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2300                      End If
2310                  Case 14
                          'tipo2
2320                      If Rs.Fields(N) <> 0 Then
2330                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2340                      End If
2350                  Case 15
                          'cuota2
2360                      If Rs.Fields(N) <> 0 Then
2370                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2380                      End If
2390                  Case 16
                          'tipo RE2
2400                      If Rs.Fields(N) <> 0 Then
2410                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2420                      End If
2430                  Case 17
                          'cuota RE2
2440                      If Rs.Fields(N) <> 0 Then
2450                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2460                      End If
2470                  Case 18
                          'base 3
2480                      If Rs.Fields(N) <> 0 Then
2490                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2500                      End If
2510                  Case 19
                          'tipo iva 3
2520                      If Rs.Fields(N) <> 0 Then
2530                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2540                      End If
2550                  Case 20
                          'cuota 3
2560                      If Rs.Fields(N) <> 0 Then
2570                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2580                      End If
2590                  Case 21
                          'tipo re3
2600                      If Rs.Fields(N) <> 0 Then
2610                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2620                      End If
2630                  Case 22
                          'cuota 3
2640                      If Rs.Fields(N) <> 0 Then
2650                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2660                      End If
2670                  Case 23
                          'Base4
2680                      If Rs.Fields(N) <> 0 Then
2690                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2700                      End If
2710                  Case 24
                          'Tipo4
2720                      If Rs.Fields(N) <> 0 Then
2730                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2740                      End If
2750                  Case 25
                          'Cuota4
2760                      If Rs.Fields(N) <> 0 Then
2770                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2780                      End If
2790                  Case 26
                          'TipoRE4
2800                      If Rs.Fields(N) <> 0 Then
2810                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2820                      End If
2830                  Case 27
                          'CuotaRE4
2840                      If Rs.Fields(N) <> 0 Then
2850                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2860                      End If
2870                  Case 28
                          'Base5
2880                      If Rs.Fields(N) <> 0 Then
2890                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2900                      End If
      '====
2910                  Case 29
                          'Tipo5
2920                      If Rs.Fields(N) <> 0 Then
2930                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2940                      End If
2950                  Case 30
                          'Cuota5
2960                      If Rs.Fields(N) <> 0 Then
2970                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2980                      End If
2990                  Case 31
                          'TipoRE5
3000                      If Rs.Fields(N) <> 0 Then
3010                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
3020                      End If
3030                  Case 32
                          'CuotaRE5
3040                      If Rs.Fields(N) <> 0 Then
3050                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
3060                      End If
3070                  Case 33
                          'idsujeto
3080                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
3090                  Case 34
                          'base99
3100                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
3110                  Case 35
                          'codcontable
3120                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
3130                  Case 36
                          'contrapartida
3140                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
3150                  Case 37
                          'irpf
3160                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
3170                  Case 38
                          'cuentairpf
3180                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)

      '===
3190              End Select
3200          Next
3210          vFila = vFila + 1
3220          End If
3230          Rs.MoveNext
3240      Wend

      '===la segunda pasada fin
3250      With oSheet
3260          .Cells.Select
3270          oBook.Application.Selection.Columns.AutoFit
3280      End With
          
3290      If Not erroNoCampo Then
3300          oSheet.Columns("I:AG").Select
3310          oExcel.Selection.NumberFormat = "#,##0.00"
3320          oExcel.Selection.ColumnWidth = 10
              
3330          oSheet.Columns("AH:AH").Select
              'oExcel.Selection.NumberFormat = "#,##0"
3340          oExcel.Selection.NumberFormat = "0"
              
              'intNumFacturas = vFila
              'MiSql = "A1:S" & vFila
              'oSheet.Range(MiSql).Select
3350          intNumFacturas = intNumFacturas
          
              '230125 seleccion de casillas
3360          oSheet.Range("A1:AH" & vFila - 1).Select
3370          oSheet.Range("A1:AH" & intNumFacturas - 1).Select
              'oSheet.Range("A1:X" & vFila - 1).Select
              'oSheet.Range("A1:X" & intNumFacturas - 1).Select
3380      Else
3390          oSheet.Columns("I:W").Select
3400          oExcel.Selection.NumberFormat = "#,##0.00"
3410          oExcel.Selection.ColumnWidth = 10
              
3420          oSheet.Columns("X:X").Select
              'oExcel.Selection.NumberFormat = "#,##0"
3430          oExcel.Selection.NumberFormat = "0"
              
              'intNumFacturas = vFila
              'MiSql = "A1:S" & vFila
              'oSheet.Range(MiSql).Select
3440          intNumFacturas = intNumFacturas
3450          oSheet.Range("A1:X" & vFila - 1).Select
3460          oSheet.Range("A1:X" & intNumFacturas - 1).Select
3470      End If
          
          'Guardar el libro y cerrar Excel.
3480      oBook.SaveAs (Fichero)
          'oExcel.Visible = True
3490      Set oSheet = Nothing
3500      Set oBook = Nothing
3510      oExcel.Quit
3520      Set oExcel = Nothing
3530      On Error GoTo 0
3540      Exit Sub
Fallo:
3550      Msg mError, "", Err, Erl
3560      Resume
3570      Exit Sub
NoCampo:
3580      erroNoCampo = True
3590      Resume Next
Salir:
3600      oBook.SaveAs (Fichero)
3610      Set oSheet = Nothing
3620      Set oBook = Nothing
3630      oExcel.Quit
3640      Set oExcel = Nothing
3650      Err.Raise Err.Number, "CrearHojaExcelCompras2019 " & Erl, Err.Description
End Sub

'===
Public Sub CrearHojaExcelCompras(Rs As ADODB.Recordset, Fichero As String)
10        On Error GoTo Fallo
          Dim oExcel As excel.Application
          Dim oBook As excel.Workbook
          Dim oSheet As excel.Worksheet

20        Set oExcel = New excel.Application
30        Set oBook = oExcel.Workbooks.Add
40        Set oSheet = oBook.Worksheets(1)
          '15/02/2014 modificada CrearExcelDesatendido para que guarde el fichero segun la version instalada
          'el fichero hay que pasarlo sin la extension
          
          Dim ExcelVersion As Long
50        ExcelVersion = Val(oExcel.Version)
          
60        Select Case ExcelVersion
              Case 11 'excel 2003
70                Fichero = Fichero & ".xls"
80            Case 12 'excel 2007
90                Fichero = Fichero & ".xlsx"
100           Case 14 'excel 2010
110               Fichero = Fichero & ".xlsx"
120       End Select
          Dim N As Integer
130       For N = 0 To Rs.Fields.Count - 1
140           oSheet.Cells(1, N + 1).Value = Rs.Fields(N).Name
150       Next
          'GoTo Salir
          Dim vFila As Long
          Dim vCol As Long
          Dim dblImporte As Double
          Dim dblSumaLinea As Double
          Dim blnSaltar As Boolean
          
          '170813 blnEsFactura en CrearHojaExcelCompras
          Dim blnEsFactura As Boolean
          Dim dblSumaIVA As Double
          Dim intNumFacturas As Integer
          
160       vFila = 2
170       While Not Rs.EOF
              '170813 blnEsFactura en CrearHojaExcelCompras
180           blnEsFactura = False
190           dblSumaIVA = Rs.Fields(10) + Rs.Fields(15) + Rs.Fields(21)
200           If dblSumaIVA <> 0 Then blnEsFactura = True
              'esto es el dni
210           If Rs.Fields(6) = "" Then blnEsFactura = False
              
220           If blnEsFactura Then
230           For N = 0 To Rs.Fields.Count - 1
                  'esto es la suma del iva
240               dblSumaLinea = Rs.Fields(10) + Rs.Fields(12) + Rs.Fields(15) + Rs.Fields(17) + Rs.Fields(20) + Rs.Fields(22)
                  
250               Select Case N
      '===
                      Case 0
260                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
270                   Case 1
280                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
290                   Case 2
300                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
310                   Case 3
320                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
330                   Case 4
                          'si el importe es 0 comprobar si  hay valor en alguna cuota
340                       If Rs.Fields(6) = 0 Then
350                           If dblSumaLinea = 0 Then
360                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
370                           End If
380                       Else
390                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
400                       End If
410                   Case 5
420                       If Rs.Fields(6) = 0 Then
430                           If dblSumaLinea = 0 Then
440                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
450                           End If
460                       Else
470                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
480                       End If
490                   Case 6
500                       If Rs.Fields(6) = 0 Then
510                           If dblSumaLinea = 0 Then
520                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
530                           End If
540                       Else
550                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
560                       End If
570                   Case 7
580                       If Rs.Fields(8) <> 0 Then
590                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
600                       End If
610                   Case 8
620                       If Rs.Fields(8) <> 0 Then
630                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
640                       End If
650                   Case 9
660                       If Rs.Fields(11) <> 0 Then
670                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
680                       End If
690                   Case 10
700                       If Rs.Fields(11) <> 0 Then
710                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
720                       End If
730                   Case 11
740                       If Rs.Fields(11) <> 0 Then
750                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
760                       End If
770                   Case 12
780                       If Rs.Fields(13) <> 0 Then
790                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
800                       End If
810                   Case 13
820                       If Rs.Fields(13) <> 0 Then
830                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
840                       End If
850                   Case 14
860                       If Rs.Fields(16) <> 0 Then
870                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
880                       End If
890                   Case 15
900                       If Rs.Fields(16) <> 0 Then
910                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
920                       End If
930                   Case 16
940                       If Rs.Fields(16) <> 0 Then
950                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
960                       End If
970                   Case 16
980                       If Rs.Fields(18) <> 0 Then
990                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1000                      End If
1010                  Case 18
1020                      If Rs.Fields(18) <> 0 Then
1030                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1040                      End If
1050                  Case 19
1060                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1070                  Case 20
1080                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1090                  Case 21
1100                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1110                  Case 22
1120                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1130                  Case 23
1140                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1150                  Case 24
1160                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
      '===
1170              End Select
1180          Next
1190          vFila = vFila + 1
1200          End If
1210          Rs.MoveNext
              'vFila = vFila + 1
1220      Wend
          
1230      intNumFacturas = vFila
1240      If Rs.RecordCount > 0 Then
1250          Rs.MoveFirst
1260      End If
      '===la segunda pasada
1270      While Not Rs.EOF
              '170813 blnEsFactura en CrearHojaExcelCompras
1280          blnEsFactura = False
1290          dblSumaIVA = Rs.Fields(6) + Rs.Fields(11) + Rs.Fields(16)
1300          If dblSumaIVA <> 0 Then blnEsFactura = True
1310          If Rs.Fields(2) = "" Then blnEsFactura = False
              
1320          If Not blnEsFactura Then
1330          For N = 0 To Rs.Fields.Count - 1
1340              dblSumaLinea = Rs.Fields(8) + Rs.Fields(11) + Rs.Fields(13) + Rs.Fields(16) + Rs.Fields(18)
                  
1350              Select Case N
      '===
                      Case 0
1360                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1370                  Case 1
1380                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1390                  Case 2
1400                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1410                  Case 3
1420                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1430                  Case 4
                          'si el importe es 0 comprobar si  hay valor en alguna cuota
1440                      If Rs.Fields(6) = 0 Then
1450                          If dblSumaLinea = 0 Then
1460                              oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1470                          End If
1480                      Else
1490                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1500                      End If
1510                  Case 5
1520                      If Rs.Fields(6) = 0 Then
1530                          If dblSumaLinea = 0 Then
1540                              oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1550                          End If
1560                      Else
1570                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1580                      End If
1590                  Case 6
1600                      If Rs.Fields(6) = 0 Then
1610                          If dblSumaLinea = 0 Then
1620                              oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1630                          End If
1640                      Else
1650                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1660                      End If
1670                  Case 7
1680                      If Rs.Fields(8) <> 0 Then
1690                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1700                      End If
1710                  Case 8
1720                      If Rs.Fields(8) <> 0 Then
1730                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1740                      End If
1750                  Case 9
1760                      If Rs.Fields(11) <> 0 Then
1770                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1780                      End If
1790                  Case 10
1800                      If Rs.Fields(11) <> 0 Then
1810                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1820                      End If
1830                  Case 11
1840                      If Rs.Fields(11) <> 0 Then
1850                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1860                      End If
1870                  Case 12
1880                      If Rs.Fields(13) <> 0 Then
1890                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1900                      End If
1910                  Case 13
1920                      If Rs.Fields(13) <> 0 Then
1930                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1940                      End If
1950                  Case 14
1960                      If Rs.Fields(16) <> 0 Then
1970                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1980                      End If
1990                  Case 15
2000                      If Rs.Fields(16) <> 0 Then
2010                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2020                      End If
2030                  Case 16
2040                      If Rs.Fields(16) <> 0 Then
2050                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2060                      End If
2070                  Case 16
2080                      If Rs.Fields(18) <> 0 Then
2090                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2100                      End If
2110                  Case 18
2120                      If Rs.Fields(18) <> 0 Then
2130                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2140                      End If
2150                  Case 19
2160                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2170                  Case 20
2180                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2190                  Case 21
2200                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2210                  Case 22
2220                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2230                  Case 23
2240                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
2250                  Case 24
2260                      oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
      '===
2270              End Select
2280          Next
2290          vFila = vFila + 1
2300          End If
2310          Rs.MoveNext
              'vFila = vFila + 1
2320      Wend

      '===la segunda pasada fin
2330      With oSheet
2340          .Cells.Select
2350          oBook.Application.Selection.Columns.AutoFit
2360      End With
2370      oSheet.Columns("E:S").Select
2380      oExcel.Selection.NumberFormat = "#,##0.00"
2390      oExcel.Selection.ColumnWidth = 10
          
2400      oSheet.Columns("T:T").Select
          'oExcel.Selection.NumberFormat = "#,##0"
2410      oExcel.Selection.NumberFormat = "0"
          
          'MiSql = "A1:S" & vFila
          'oSheet.Range(MiSql).Select
2420      intNumFacturas = intNumFacturas
2430      oSheet.Range("A1:T" & vFila - 1).Select
2440      oSheet.Range("A1:T" & intNumFacturas - 1).Select
          'Guardar el libro y cerrar Excel.
2450      oBook.SaveAs (Fichero)
          'oExcel.Visible = True
2460      Set oSheet = Nothing
2470      Set oBook = Nothing
2480      oExcel.Quit
2490      Set oExcel = Nothing
2500      Exit Sub
          
Fallo:
2510      Msg mError, "", Err, Erl
2520      Set oSheet = Nothing
2530      Set oBook = Nothing
2540      oExcel.Quit
2550      Set oExcel = Nothing
End Sub
Public Sub CrearHojaExcel(Rs As ADODB.Recordset, Fichero As String)
10        On Error GoTo Fallo
          Dim errorNoCampo As Boolean
          Dim oExcel As excel.Application
          Dim oBook As excel.Workbook
          Dim oSheet As excel.Worksheet

20        Set oExcel = New excel.Application
30        Set oBook = oExcel.Workbooks.Add
40        Set oSheet = oBook.Worksheets(1)
          '15/02/2014 modificada CrearExcelDesatendido para que guarde el fichero segun la version instalada
          'el fichero hay que pasarlo sin la extension
          
          Dim ExcelVersion As Long
50        ExcelVersion = Val(oExcel.Version)
          
60        Select Case ExcelVersion
              Case 11 'excel 2003
70                Fichero = Fichero & ".xls"
80            Case 12 'excel 2007
90                Fichero = Fichero & ".xlsx"
100           Case 14 'excel 2010
110               Fichero = Fichero & ".xlsx"
120       End Select
          Dim N As Integer
130       For N = 0 To Rs.Fields.Count - 1
140           oSheet.Cells(1, N + 1).Value = Rs.Fields(N).Name
150       Next
          
          Dim vFila As Long
          Dim vCol As Long
          Dim dblImporte As Double
          Dim dblSumaLinea As Double
          Dim blnSaltar As Boolean
          
160       vFila = 2
170       While Not Rs.EOF
              
180                       If Rs.Fields(0) = "A9723" Then
190                           MiSql = MiSql
200                       End If
210           For N = 0 To Rs.Fields.Count - 1
                  'cuota1 = 6
                  'cuotare1 = 8
                  'cuota2 = 11
                  'cuotare2 = 13
                  'cuota3 = 16
                  'cuotare3 = 18
                  
                  'para saber si hay cuotas
                  
220               If vFila = 34 Then
230                   vFila = vFila
240               End If
                  
250               dblSumaLinea = Rs.Fields(6) + Rs.Fields(8) + Rs.Fields(11) + Rs.Fields(13) + Rs.Fields(16) + Rs.Fields(18)
260               On Error GoTo NoCampo
270               dblSumaLinea = dblSumaLinea + Rs.Fields(21) + Rs.Fields(23) + Rs.Fields(26) + Rs.Fields(28)
280               On Error GoTo Fallo
              'If dblSumaLinea <> 0 Then
290               Select Case N
                      Case 0
300                       If Rs.Fields(N) = "A9723" Then
310                           MiSql = MiSql
320                       End If
                          
330                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
                          'A83445

340                   Case 1
350                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
360                   Case 2
370                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
380                   Case 3
390                       oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
400                   Case 4
                          'si el importe es 0 comprobar si  hay valor en alguna cuota
                          'If rs.Fields(6) = 0 Then
410                       If Rs.Fields(6) <> 0 Then
                              'If dblSumaLinea <> 0 Then
                                  
420                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
                              'End If
                          'Else
                          '    oSheet.Cells(vFila, N + 1).Value = rs.Fields(N)
430                       Else
440                           If dblSumaLinea = 0 Then
450                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
460                           End If
470                       End If
480                   Case 5
      '                    If rs.Fields(6) = 0 Then
      '                        If dblSumaLinea = 0 Then
      '                            oSheet.Cells(vFila, N + 1).Value = rs.Fields(N)
      '                        End If
      '                    Else
      '                        oSheet.Cells(vFila, N + 1).Value = rs.Fields(N)
      '                    End If
490                       If Rs.Fields(6) <> 0 Then
                              'If dblSumaLinea = 0 Then
500                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
                              'End If
                          'Else
                              'oSheet.Cells(vFila, N + 1).Value = rs.Fields(N)
510                       Else
520                           If dblSumaLinea = 0 Then
530                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
540                           End If
550                       End If
                      
560                   Case 6
570                       If Rs.Fields(6) <> 0 Then
580                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
      '                        If dblSumaLinea = 0 Then
      '                            oSheet.Cells(vFila, N + 1).Value = rs.Fields(N)
      '                        End If
      '                    Else
      '                        oSheet.Cells(vFila, N + 1).Value = rs.Fields(N)
590                       Else
600                           If dblSumaLinea = 0 Then
610                               oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
620                           End If
630                       End If
640                   Case 7
650                       If Rs.Fields(8) <> 0 Then
660                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
670                       End If
680                   Case 8
690                       If Rs.Fields(8) <> 0 Then
700                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
710                       End If
720                   Case 9
730                       If Rs.Fields(11) <> 0 Then
740                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
750                       End If
760                   Case 10
770                       If Rs.Fields(11) <> 0 Then
780                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
790                       End If
800                   Case 11
810                       If Rs.Fields(11) <> 0 Then
820                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
830                       End If
840                   Case 12
850                       If Rs.Fields(13) <> 0 Then
860                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
870                       End If
880                   Case 13
890                       If Rs.Fields(13) <> 0 Then
900                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
910                       End If
920                   Case 14
930                       If Rs.Fields(14) <> 0 Then
940                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
950                       End If
960                   Case 15
970                       If Rs.Fields(15) <> 0 Then
980                           oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
990                       End If
1000                  Case 16
1010                      If Rs.Fields(16) <> 0 Then
1020                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1030                      End If
1040                  Case 17
1050                      If Rs.Fields(17) <> 0 Then
1060                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1070                      End If
1080                 Case 18
1090                      If Rs.Fields(18) <> 0 Then
1100                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1110                      End If
1120                  Case 19
1130                      If Rs.Fields(N) <> 0 Then
1140                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1150                      End If
1160                  Case 20
1170                      If Rs.Fields(N - 1) <> 0 Then
1180                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1190                      End If
1200                  Case 21
1210                      If Rs.Fields(N - 2) <> 0 Then
1220                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1230                      End If
1240                  Case 22
1250                      If Rs.Fields(N - 3) <> 0 Then
1260                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1270                      End If
1280                  Case 23
1290                      If Rs.Fields(N - 4) <> 0 Then
1300                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1310                      End If
      '===
1320                  Case 24
1330                      If Rs.Fields(N) <> 0 Then
1340                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1350                      End If
1360                  Case 25
1370                      If Rs.Fields(N) <> 0 Then
1380                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1390                      End If
1400                  Case 26
1410                      If Rs.Fields(N) <> 0 Then
1420                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1430                      End If
1440                  Case 27
1450                      If Rs.Fields(N) <> 0 Then
1460                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1470                      End If
1480                  Case 28
1490                      If Rs.Fields(N) <> 0 Then
1500                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1510                      End If
1520                  Case 29
1530                      If Rs.Fields(N) <> 0 Then
1540                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1550                      End If
1560                  Case 30
1570                      If Rs.Fields(N) <> 0 Then
1580                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1590                      End If
1600                  Case 31
1610                      If Rs.Fields(N) <> 0 Then
1620                          oSheet.Cells(vFila, N + 1).Value = Rs.Fields(N)
1630                      End If

      '===
1640              End Select
                  'End If
                  
1650          Next
1660          Rs.MoveNext
1670          vFila = vFila + 1
1680      Wend
1690      With oSheet
1700          .Cells.Select
1710          oBook.Application.Selection.Columns.AutoFit
1720      End With
      '    For N = 0 To rs.Fields.Count - 1
      '        If rs.Fields(N).Type = 131 Then
      '            oSheet.Columns(N + 1).Select
      '            oExcel.Selection.NumberFormat = "#,##0.00"
      '            oExcel.Selection.ColumnWidth = 15
      '
      '        End If
      '    Next
          
1730      If Not errorNoCampo Then
1740          oSheet.Columns("E:AC").Select
1750          oExcel.Selection.NumberFormat = "#,##0.00"
1760          oExcel.Selection.ColumnWidth = 10
              
1770          oSheet.Columns("AD:AD").Select
              'oExcel.Selection.NumberFormat = "#,##0"
1780          oExcel.Selection.NumberFormat = "0"
              
              'MiSql = "A1:S" & vFila
              'oSheet.Range(MiSql).Select
1790          oSheet.Range("A1:AD" & vFila - 1).Select
1800      Else
1810          oSheet.Columns("E:AC").Select
1820          oExcel.Selection.NumberFormat = "#,##0.00"
1830          oExcel.Selection.ColumnWidth = 10
              
1840          oSheet.Columns("V:V").Select
              'oExcel.Selection.NumberFormat = "#,##0"
1850          oExcel.Selection.NumberFormat = "0"
              
              'MiSql = "A1:S" & vFila
              'oSheet.Range(MiSql).Select
1860          oSheet.Range("A1:T" & vFila - 1).Select
1870      End If
          
          
          
          'Guardar el libro y cerrar Excel.
1880      oBook.SaveAs (Fichero)
          'oExcel.Visible = True
1890      Set oSheet = Nothing
1900      Set oBook = Nothing
1910      oExcel.Quit
1920      Set oExcel = Nothing
1930      On Error GoTo 0
1940      Exit Sub
Fallo:
1950      Msg mError, "", Err, Erl
1960      Exit Sub
          
NoCampo:
1970      errorNoCampo = True
1980      Resume Next
End Sub
Public Sub CrearExcelDesatendido(Rs As ADODB.Recordset, Fichero As String)
10        On Error GoTo Fallo
          
          '13/03/11 Añadia la funcion CrearExcelDesatendido
          '======exportar a excel un recordset desconectado
          'Fichero es la ruta completa del archivo
          'para la carpeta donde se guarda el fichero
          Dim Carpeta As String
          'Crear un libro en Excel.
          Dim oExcel As excel.Application
          Dim oBook As excel.Workbook
          Dim oSheet As excel.Worksheet
          
20        Set oExcel = New excel.Application
30        Set oBook = oExcel.Workbooks.Add
40        Set oSheet = oBook.Worksheets(1)
          'oExcel.Visible = True
          
          '15/02/2014 modificada CrearExcelDesatendido para que guarde el fichero segun la version instalada
          'el fichero hay que pasarlo sin la extension
          
          Dim ExcelVersion As Long
50        ExcelVersion = Val(oExcel.Version)
          
60        Select Case ExcelVersion
              Case 11 'excel 2003
70                Fichero = Fichero & ".xls"
80            Case 12 'excel 2007
90                Fichero = Fichero & ".xlsx"
100           Case 14 'excel 2010
110               Fichero = Fichero & ".xlsx"
120       End Select
          
          'Transferir los nombres de campo a la fila 1 de la hoja de cálculo:
          'Nota: CopyFromRecordset sólo copia los datos y no los nombres de campo,
          'por lo que puede transferir los nombres de campo recorriendo la colección de campos.
          Dim N As Long
130       For N = 0 To Rs.Fields.Count - 1
140           oSheet.Cells(1, N + 1).Value = Rs.Fields(N).Name
              'Debug.Print RS.Fields(n).Name
150       Next

          'Transferir los datos a Excel.
160       oSheet.Range("A2").CopyFromRecordset Rs
          
          'insertar las lineas iniciales para las cabeceras
          'oSheet.Rows("1:6").Insert Shift:=xlDown
          'dar ancho a las columas
          'oSheet.Columns("B:B").ColumnWidth = 34.14
          
          '28/07/12 Modificado CrearExcelDesatendido para que formatee las columnas numericas con dos decimales, y les asigne a todas el mismo ancho
170       With oSheet
180           .Cells.Select
190           oBook.Application.Selection.Columns.AutoFit
200       End With
          
210       For N = 0 To Rs.Fields.Count - 1
220           If Rs.Fields(N).Type = 5 Or Rs.Fields(N).Type = 6 Then
230               oSheet.Columns(N + 1).Select
240               oExcel.Selection.NumberFormat = "#,##0.00"
250               oExcel.Selection.ColumnWidth = 15

260           End If
              'If RS.Fields(n).Type =
270       Next
          '28/07/12 Modificado CrearExcelDesatendido para que formatee las columnas numericas con dos decimales, y les asigne a todas el mismo ancho
          
          Dim ZZ As Long
          
          'consumos
280       ZZ = InStr(1, Fichero, "consumos", vbTextCompare)
290       If ZZ <> 0 Then
300           ZZ = Rs.RecordCount
              
              '=== sumar columnas
              
              
              
310           MiSql = "=sum(r[-" & ZZ + 1 & "]C:R[-1]C)"
320           oExcel.ActiveSheet.Cells(ZZ + 2, 12).FormulaR1C1 = MiSql
330           oExcel.ActiveSheet.Cells(ZZ + 2, 13).FormulaR1C1 = MiSql
340           oExcel.ActiveSheet.Cells(ZZ + 2, 22).FormulaR1C1 = MiSql
350           oExcel.ActiveSheet.Cells(ZZ + 2, 23).FormulaR1C1 = MiSql
              'oExcel.ActiveSheet.Cells(ZZ + 1, 4).FormulaR1C1 = MiSql
360       End If
          
          'diferencias
370       ZZ = InStr(1, Fichero, "diferencias", vbTextCompare)
          
380       If ZZ <> 0 Then
390           ZZ = Rs.RecordCount
              
              '=== sumar columnas
              
              
              
400           MiSql = "=sum(r[-" & ZZ + 1 & "]C:R[-1]C)"
410           oExcel.ActiveSheet.Cells(ZZ + 2, 2).FormulaR1C1 = MiSql
420           oExcel.ActiveSheet.Cells(ZZ + 2, 3).FormulaR1C1 = MiSql
430           oExcel.ActiveSheet.Cells(ZZ + 2, 4).FormulaR1C1 = MiSql
440       End If
          
450       oExcel.Range("A1").Select
          'borrar el richero si existe
460       If ArchivoExistente(Fichero) Then
470           Kill Fichero
480       End If
          'Guardar el libro y cerrar Excel.
490       oBook.SaveAs (Fichero)
          'oExcel.Visible = True
500       Set oSheet = Nothing
510       Set oBook = Nothing
520       oExcel.Quit
          
530       Set oExcel = Nothing
          'esto obtiene la carpeta donde se encuentra el fichero
540       Carpeta = Mid(Fichero, 1, InStrRev(Fichero, "\", , vbTextCompare))
          'abre la carpeta con el explorador
          'Externa = Shell("explorer " & Carpeta, vbNormalFocus)
550       On Error GoTo 0
560       Exit Sub
Fallo:
570       Msg mError, "", Err, Erl
End Sub

Public Function Importar_Excel( _
    Libro As String, _
    Optional hoja As String = "Hoja1", _
    Optional Rango As String = "") As ADODB.Recordset
        
          Dim conexion As ADODB.Connection
          Dim Rs As ADODB.Recordset
          
10        On Error GoTo Fallo
20        Set conexion = New ADODB.Connection
          
          '200302 importar office 2007
          Dim CadConexion As String
          
          'Crear un libro en Excel.
          Dim oExcel As excel.Application
30        Set oExcel = New excel.Application
          Dim ExcelVersion As Long
40        ExcelVersion = Val(oExcel.Version)
          
50        Select Case ExcelVersion
              Case 11 'excel 2003
60                MiSql = MiSql
70                CadConexion = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                "Data Source=" & Libro & _
                                ";Extended Properties=""Excel 8.0;HDR=Yes;"""
80            Case 12 'excel 2007
90                MiSql = MiSql
100               CadConexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                "Data Source=" & Libro & _
                                ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;"""
110           Case 14 'excel 2010
120               MiSql = MiSql
130               CadConexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                "Data Source=" & Libro & _
                                ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;"""
140       End Select
          
150       Set oExcel = Nothing
160       conexion.Open CadConexion
            
          'conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source=" & Libro & _
                        ";Extended Properties=""Excel 8.0;HDR=Yes;"""
                        
                        'MaxScanRows=0;
          '200302 importar office 2007 fin
          
          
          ' Nuevo recordset
170       Set Rs = New ADODB.Recordset
            
180       With Rs
190           .CursorLocation = adUseClient
200           .CursorType = adOpenStatic
210           .LockType = adLockOptimistic
220       End With
        
230       If Rango <> ":" Then
240          hoja = hoja & "$" & Rango
250       End If
            
260       Rs.Open "SELECT * FROM [" & hoja & "]", conexion, , , adCmdText
            
          ' Mostramos los datos en el datagrid
          'Set DataGrid1.DataSource = rs
270       Set Importar_Excel = Rs
          
          'rs.Close
          'Set rs = Nothing
280       On Error GoTo 0
290       Exit Function
Fallo:
300       Err.Raise Err.Number, "Importar_Excel", Err.Description
End Function





