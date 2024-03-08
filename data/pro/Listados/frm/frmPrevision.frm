VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrevision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros del listado de previsiones"
   ClientHeight    =   4365
   ClientLeft      =   12540
   ClientTop       =   5460
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Selección de informe"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   150
      TabIndex        =   26
      Top             =   3090
      Width           =   4290
      Begin VB.OptionButton optInforme 
         Caption         =   "Oculta Artículos"
         Height          =   195
         Index           =   1
         Left            =   2415
         TabIndex        =   13
         Top             =   300
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton optInforme 
         Caption         =   "Muestra Artículos"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   12
         Top             =   300
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3225
      TabIndex        =   0
      Top             =   3990
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   315
      Left            =   1800
      TabIndex        =   14
      Top             =   3990
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Height          =   30
      Left            =   150
      TabIndex        =   25
      Top             =   3915
      Width           =   4290
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mes"
      ForeColor       =   &H00FF0000&
      Height          =   915
      Left            =   1320
      TabIndex        =   16
      Top             =   255
      Width           =   3120
      Begin MSComCtl2.DTPicker dtpMesIni 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   435
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM"
         Format          =   93847555
         UpDown          =   -1  'True
         CurrentDate     =   36977
      End
      Begin MSComCtl2.DTPicker dtpMesFin 
         Height          =   330
         Left            =   1620
         TabIndex        =   3
         Top             =   435
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMMM"
         Format          =   93847555
         UpDown          =   -1  'True
         CurrentDate     =   36977
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Año"
      ForeColor       =   &H00FF0000&
      Height          =   915
      Left            =   150
      TabIndex        =   15
      Top             =   255
      Width           =   1020
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Año anterior"
         Top             =   435
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   93847555
         UpDown          =   -1  'True
         CurrentDate     =   36977
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selección Parámetros"
      ForeColor       =   &H00FF0000&
      Height          =   1290
      Left            =   150
      TabIndex        =   19
      Top             =   1245
      Width           =   4290
      Begin VB.TextBox txtRuta 
         Height          =   285
         Index           =   1
         Left            =   780
         TabIndex        =   5
         Text            =   "32767"
         Top             =   840
         Width           =   700
      End
      Begin VB.TextBox txtRuta 
         Height          =   285
         Index           =   0
         Left            =   780
         TabIndex        =   4
         Text            =   "0"
         Top             =   435
         Width           =   700
      End
      Begin VB.TextBox txtFamilia 
         Height          =   285
         Index           =   0
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   6
         Text            =   " "
         Top             =   435
         Width           =   700
      End
      Begin VB.TextBox txtFamilia 
         Height          =   285
         Index           =   1
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "ZZZ"
         Top             =   840
         Width           =   700
      End
      Begin VB.TextBox txtsubFamilia 
         Height          =   285
         Index           =   0
         Left            =   2610
         MaxLength       =   3
         TabIndex        =   8
         Text            =   " "
         Top             =   435
         Width           =   700
      End
      Begin VB.TextBox txtsubFamilia 
         Height          =   285
         Index           =   1
         Left            =   2610
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "ZZZ"
         Top             =   840
         Width           =   700
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   0
         Left            =   3450
         MaxLength       =   4
         TabIndex        =   10
         Text            =   " "
         Top             =   435
         Width           =   700
      End
      Begin VB.TextBox txtArticulo 
         Height          =   285
         Index           =   1
         Left            =   3450
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "ZZZZ"
         Top             =   840
         Width           =   700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta:"
         Height          =   195
         Index           =   7
         Left            =   780
         TabIndex        =   27
         Top             =   225
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "INICIO :"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   24
         Top             =   525
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Família:"
         Height          =   195
         Index           =   3
         Left            =   1755
         TabIndex        =   23
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Subfamília:"
         Height          =   195
         Index           =   4
         Left            =   2610
         TabIndex        =   22
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   5
         Left            =   3450
         TabIndex        =   21
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FINAL :"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   20
         Top             =   930
         Width           =   540
      End
   End
   Begin VB.Frame Frame6 
      Height          =   525
      Left            =   150
      TabIndex        =   28
      Top             =   2415
      Width           =   4290
      Begin VB.OptionButton optBulto 
         Caption         =   "Sólo Bultos"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   30
         Top             =   210
         Width           =   1665
      End
      Begin VB.OptionButton optBulto 
         Caption         =   "Todos"
         Height          =   195
         Index           =   1
         Left            =   2415
         TabIndex        =   29
         Top             =   210
         Value           =   -1  'True
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmPrevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsPrincipal As ADODB.Recordset

Dim MesIni, MesFin As Integer
Dim Anyo As Integer
Dim RutaIni As Integer
Dim RutaFin As Integer
Dim pBulto As String
Private Function Consulta() As Boolean
   On Error GoTo ControlError

   Dim oCadena As String
   Dim oCadena2 As String
   Dim oCadena3 As String
   Dim I As Integer
   Dim mMeses() As String
   
   'Cadena con los campos correspondientes a prevision que deben sumarse
   Dim CPrevision As String
   Dim SelBulto As String
   
   'matriz con los nombres de los campos correspondientes a la prevision de meses
   'de la tabla RUTA para construir la cadena CPrevision que suma dichos campos
   ReDim mMeses(12)
   mMeses(0) = "PrevENE"
   mMeses(1) = "PrevFEB"
   mMeses(2) = "PrevMAR"
   mMeses(3) = "PrevABR"
   mMeses(4) = "PrevMAY"
   mMeses(5) = "PrevJUN"
   mMeses(6) = "PrevJUL"
   mMeses(7) = "PrevAGO"
   mMeses(8) = "PrevSEP"
   mMeses(9) = "PrevOCT"
   mMeses(10) = "PrevNOV"
   mMeses(11) = "PrevDIC"
   
   Anyo = DTPicker1.Year
   
   MesIni = dtpMesIni.Month
   MesFin = dtpMesFin.Month
   
   CPrevision = mMeses(MesIni - 1)
   For I = MesIni To MesFin - 1
      CPrevision = CPrevision & " + " & mMeses(I)
   Next I
   
   RutaIni = CInt(txtRuta(0).Text)
   RutaFin = CInt(txtRuta(1).Text)
   
   If optBulto(0).Value Then
      SelBulto = " AND A.EstadisticaBulto = 0 "
      pBulto = "SOLO BULTOS"
   Else
      SelBulto = " "
      pBulto = "TODOS"
   End If
   ' *****************************************************************************
   '  SELECT Campos
   '  FROM (SELECT   ... FROM Cabecera, lineas articulo de la BD ROSELL)
   '  FULL OUTER JOIN
   '       (SELECT ... FROM Rosell & ANYO.Cabecera,Rosell & ANYO.Lineas, ROSELL.Articulo) "AÑO ANTERIOR"
   '  FULL OUTER JOIN
   '       (SELECT ... FROM Ruta)
   ' *****************************************************************************
   
   oCadena = " SELECT ISNULL(ISNULL( R.IdRuta, ACT.IdRuta), ANT.IdRuta) AS IdRuta, ISNULL(R.Descripcion, 'NO EXISTE'), ISNULL( ACT.IdFamilia, ANT.IdFamilia) AS IdFamilia, " & _
                     " ISNULL(ACT.SubFamilia, ANT.SubFamilia) AS Subfamilia, ISNULL(ACT.IdArticulo, ANT.IdArticulo) AS IdArticulo, ISNULL(ACT.DescripcionA, ANT.DescripcionA) AS DescripcionA," & _
                     " ISNULL(ANT.Unidades, 0) As UnidadesANT, ISNULL(R.UnidadesPrev, 0) AS UnidadesPrev, ISNULL(ACT.Unidades, 0) AS UnidadesACT" & _
            " FROM (SELECT  C.IdRuta AS IdRuta,  A.IdFamilia AS IdFamilia, A.Subfamilia AS Subfamilia,  L.IdArticulo AS IdArticulo, MIN(A.Descripcion) AS DescripcionA, SUM(L.Unidades) AS Unidades " & _
                     " FROM Cabecera C INNER JOIN Lineas L ON C.IdCabecera = L.IdCabecera " & _
                                    " INNER JOIN Articulo A ON A.IdArticulo = L.IdArticulo " & _
                     " Where C.Servido <> 0 " & _
                     " AND   L.TipoLinea = ''" & _
                     " AND   A.TipoArticulo <> 'ENV'" & _
                     SelBulto & _
                     " AND   MONTH(C.FechaServido) BETWEEN " & MesIni & " AND " & MesFin & _
                     " AND   YEAR(C.FechaServido) = YEAR (GETDATE())" & _
                     " AND   C.IdRuta BETWEEN " & RutaIni & " AND " & RutaFin & _
                     " AND   A.IdFamilia BETWEEN '" & txtFamilia(0).Text & "' AND '" & txtFamilia(1).Text & "' " & _
                     " AND   A.Subfamilia BETWEEN '" & txtsubFamilia(0).Text & "' AND '" & txtsubFamilia(1).Text & "' " & _
                     " AND   A.IdArticulo BETWEEN '" & txtArticulo(0).Text & "' AND '" & txtArticulo(1).Text & "' " & _
                     " GROUP BY C.IdRuta, A.IdFamilia, A.Subfamilia, L.IdArticulo ) AS ACT "
   oCadena2 = " FULL OUTER JOIN " & _
                  " (SELECT  C.IdRuta AS IdRuta,  A.IdFamilia AS IdFamilia, A.Subfamilia AS Subfamilia,  L.IdArticulo AS IdArticulo, MIN(A.Descripcion) AS DescripcionA, SUM(L.Unidades) As Unidades " & _
                     " FROM Rosell" & Anyo & ".dbo.Cabecera C INNER JOIN Rosell" & Anyo & ".dbo.Lineas L ON C.IdCabecera = L.IdCabecera " & _
                                                      " INNER JOIN Articulo A ON A.IdArticulo = L.IdArticulo " & _
                     " Where C.Servido <> 0 " & _
                     " AND   LTRIM(L.TipoLinea) = '' " & _
                     " AND   A.TipoArticulo <> 'ENV' " & _
                     SelBulto & _
                     " AND   MONTH(C.FechaServido) BETWEEN " & MesIni & " AND " & MesFin & _
                     " AND   YEAR(C.FechaServido) = " & Anyo & _
                     " AND   C.IdRuta BETWEEN " & RutaIni & " AND " & RutaFin & _
                     " AND   A.IdFamilia BETWEEN '" & txtFamilia(0).Text & "' AND '" & txtFamilia(1).Text & "' " & _
                     " AND   A.Subfamilia BETWEEN '" & txtsubFamilia(0).Text & "' AND '" & txtsubFamilia(1).Text & "' " & _
                     " AND   A.IdArticulo BETWEEN '" & txtArticulo(0).Text & "' AND '" & txtArticulo(1).Text & "' " & _
                     " GROUP BY C.IdRuta, A.IdFamilia, A.Subfamilia, L.IdArticulo ) AS ANT " & _
                     " ON ACT.IdRuta = ANT.IdRuta AND ACT.IdFamilia = ANT.IdFamilia AND ACT.SubFamilia = ANT.SubFamilia AND ACT.IdArticulo = ANT.IdArticulo " & _
               " FULL OUTER JOIN  (SELECT IdRuta, Descripcion, " & CPrevision & " AS UnidadesPrev FROM Ruta WHERE IdRuta BETWEEN " & RutaIni & " AND " & RutaFin & " ) AS R ON R.IdRuta = ISNULL(ACT.IdRuta, ANT.IdRuta) " & _
               " ORDER BY R.IdRuta "
      
      oCadena3 = oCadena & oCadena2
      
      Set rsPrincipal = New ADODB.Recordset
      
      DB.CommandTimeout = 120
      rsPrincipal.Open oCadena3, DB, adOpenStatic, adLockReadOnly
      DB.CommandTimeout = 30
      Consulta = True
Exit Function
ControlError:
   If Err.Number = -2147217865 Then
      MsgBox "No existe la base de datos del año seleccionado"
   Else
      MsgBox Err.Description
   End If
   Consulta = False
End Function


Private Sub cmdAceptar_Click()
Screen.MousePointer = vbHourglass
   Dim strRuta As String
   Dim pos As Long
   
   Dim Objeto As Form

   If Consulta Then
'''        FrmImpresoras.Show 1
'''
'''        If gNombreImpresora = "" Then Exit Sub
'''        Screen.MousePointer = vbHourglass
      'v PARA VER EL INFORME ANTES DE IMPRIMIR
   
      Set Objeto = New frmVisorListados
      Load Objeto
            strRuta = App.Path
            pos = InStrRev(strRuta, "\")
            strRuta = Mid(strRuta, 1, pos)
            strRuta = strRuta & "INFORMES\PREVISION\PrevRutaArtAdo.Rpt"
      Objeto.AbreInforme strRuta
      'Objeto.AbreInforme "..\INFORMES\PREVISION\PrevRutaArtAdo.Rpt"
      If optInforme(1).Value Then   'Oculta articulo
         Objeto.OcultaSeccion "GH2"
         Objeto.OcultaSeccion "GH3"
         Objeto.OcultaSeccion "D"
         Objeto.OcultaSeccion "GF2"
         Objeto.OcultaSeccion "GF3"
      End If
      Objeto.AsignaRs rsPrincipal, 0, ""
      
      Objeto.AsignaValorParam 1, Anyo
      Objeto.AsignaValorParam 2, Format(dtpMesIni.Value, "mmmm")
      Objeto.AsignaValorParam 3, Format(dtpMesFin.Value, "mmmm")
      Objeto.AsignaValorParam 4, RutaIni
      Objeto.AsignaValorParam 5, RutaFin
      Objeto.AsignaValorParam 6, txtFamilia(0).Text
      Objeto.AsignaValorParam 7, txtFamilia(1).Text
      Objeto.AsignaValorParam 8, txtsubFamilia(0).Text
      Objeto.AsignaValorParam 9, txtsubFamilia(1).Text
      Objeto.AsignaValorParam 10, txtArticulo(0).Text
      Objeto.AsignaValorParam 11, txtArticulo(1).Text
      Objeto.AsignaValorParam 12, pBulto
      
      Objeto.MuestraInforme
   
      Objeto.Show
   
      Set Objeto = Nothing
      
      Set rsPrincipal = Nothing
      '^
      Unload Me
   End If
Screen.MousePointer = vbDefault
End Sub


Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub txtArticulo_GotFocus(Index As Integer)
   txtArticulo(Index).Appearance = 0
   txtArticulo(Index).SelStart = 0
   txtArticulo(Index).SelLength = Len(txtArticulo(Index).Text)
End Sub


Private Sub txtArticulo_LostFocus(Index As Integer)
   txtArticulo(Index).Appearance = 1
End Sub


Private Sub txtFamilia_GotFocus(Index As Integer)
   txtFamilia(Index).Appearance = 0
   txtFamilia(Index).SelStart = 0
   txtFamilia(Index).SelLength = Len(txtFamilia(Index).Text)
End Sub


Private Sub txtFamilia_LostFocus(Index As Integer)
   txtFamilia(Index).Appearance = 1
End Sub

Private Sub txtRuta_GotFocus(Index As Integer)
   txtRuta(Index).Appearance = 0
   txtRuta(Index).SelStart = 0
   txtRuta(Index).SelLength = Len(txtRuta(Index).Text)
End Sub


Private Sub txtRuta_LostFocus(Index As Integer)
   txtRuta(Index).Appearance = 1
End Sub


Private Sub txtRuta_Validate(Index As Integer, Cancel As Boolean)
On Error GoTo ControlError
   Dim kk As Integer
   kk = CInt(txtRuta(Index).Text)
      
Exit Sub
ControlError:
   Cancel = True
   MsgBox "Introduzca valores correctos"
   txtRuta(Index).Text = ""
End Sub

Private Sub txtsubFamilia_GotFocus(Index As Integer)
   txtsubFamilia(Index).Appearance = 0
   txtsubFamilia(Index).SelStart = 0
   txtsubFamilia(Index).SelLength = Len(txtsubFamilia(Index).Text)
End Sub

Private Sub txtsubFamilia_LostFocus(Index As Integer)
   txtsubFamilia(Index).Appearance = 1
End Sub


