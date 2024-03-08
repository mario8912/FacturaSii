VERSION 5.00
Begin VB.UserControl Campo 
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HitBehavior     =   0  'None
   KeyPreview      =   -1  'True
   ScaleHeight     =   1485
   ScaleWidth      =   4005
   ToolboxBitmap   =   "Campo.ctx":0000
   Begin VB.CommandButton Boton 
      Height          =   195
      Left            =   2040
      Picture         =   "Campo.ctx":00FA
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Marquita 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2100
      ScaleHeight     =   315
      ScaleWidth      =   75
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox LinTxtAbj 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   840
      ScaleHeight     =   15
      ScaleWidth      =   1335
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line LinDesDer 
      BorderColor     =   &H80000014&
      Visible         =   0   'False
      X1              =   3600
      X2              =   3600
      Y1              =   420
      Y2              =   720
   End
   Begin VB.Line LinTxtIzq 
      BorderColor     =   &H80000015&
      Visible         =   0   'False
      X1              =   1020
      X2              =   1020
      Y1              =   360
      Y2              =   660
   End
   Begin VB.Image OptionMark 
      Height          =   195
      Left            =   1680
      Picture         =   "Campo.ctx":0684
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image CheckMark 
      Height          =   135
      Left            =   2400
      Picture         =   "Campo.ctx":07CE
      Stretch         =   -1  'True
      Top             =   1140
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line LinDesAbj 
      BorderColor     =   &H80000014&
      Visible         =   0   'False
      X1              =   2640
      X2              =   3600
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Line LinDesArr 
      BorderColor     =   &H80000015&
      Visible         =   0   'False
      X1              =   2640
      X2              =   3600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line LinDesIzq 
      BorderColor     =   &H80000015&
      Visible         =   0   'False
      X1              =   2580
      X2              =   2580
      Y1              =   480
      Y2              =   780
   End
   Begin VB.Line LinTxtDer 
      BorderColor     =   &H80000014&
      Visible         =   0   'False
      X1              =   1980
      X2              =   1980
      Y1              =   360
      Y2              =   660
   End
   Begin VB.Line LinTxtArr 
      BorderColor     =   &H80000015&
      Visible         =   0   'False
      X1              =   1020
      X2              =   1980
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   120
      Picture         =   "Campo.ctx":0AD8
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label LblEtiqueta 
      Caption         =   "Etiqueta:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label LblEtiquetaFnd 
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   300
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label LblPor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1860
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label LblDescripcion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label LblCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   3360
      TabIndex        =   7
      Top             =   1080
      Width           =   150
   End
End
Attribute VB_Name = "Campo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'--> Control ocx para la introduccion de datos en los formularios
'--> @author Manuel Peris

'Uso General:
Dim Editando As Boolean
Dim nEditando As Integer
Dim SeHaEditado As Boolean
Dim OldValue As String
Dim Validado As Boolean
Dim m_AlgunCambio As Boolean
Dim Acccion As String
Dim SignoDecimal As String
Dim VieneDeV As Boolean
'Eventos:
Event CheckClick(ByRef Cancel As Boolean)
Event TxtChange()
Event TxtGotFocus()
Event TxtKeyDown(KeyCode As Integer, Shift As Integer)
Event TxtKeyPress(KeyAscii As Integer)
Event TxtLostFocus()
Event BotonClick(Boton As Integer)

Public Enum TiposLetra 'No se usa, aunque funciona
    TLPequeña = 0
    TLMediana = 1
    TLGrande = 2
End Enum
Dim TFC As FormatosCadenas
Dim qTFC As FormatosCadenas

'Valores por defecto:
Const m_def_TipoLetra = TiposLetra.TLPequeña
Const m_def_Formato = Formatos.FCadena
Const m_def_Longitud = 50
Const m_def_Texto = ""
Const m_def_Etiqueta = "Etiqueta:"
Const m_def_AlineacionEtiqueta = TiposAlineacion.TAIzquierda
Const m_def_PosicionEtiqueta = TiposPosicion.TPIzquierda
Const m_def_Tip = ""
Const m_def_TablaBD = ""
Const m_def_CampoBD = ""
Const m_def_CampoBDDescripcion = ""
Const m_def_Descripcion = ""
Const m_def_EtiquetaVisible = True
Const m_def_TextoVisible = True
Const m_def_BotonVisible = False
Const m_def_DescripcionVisible = False
Const m_def_Activado = True
Const m_def_Visible = True
Const m_def_EtiquetaNegrita = False
Const m_def_Activar = True
Const m_def_Ocultar = False
Const m_def_Obligar = False

Const m_def_ColorFondoDescripcionConFoco = &H8000000F
Const m_def_ColorFondoDescripcionSinFoco = &H8000000F
Const m_def_ColorFondoEtiquetaConFoco = &H80000014
Const m_def_ColorFondoEtiquetaSinFoco = &H8000000F
Const m_def_ColorFondoTextoConFoco = &H80000005
Const m_def_ColorFondoTextoSinFoco = &H80000014
Const m_def_ColorLetraDescripcionConFoco = &H80000012
Const m_def_ColorLetraDescripcionSinFoco = &H80000011
Const m_def_ColorLetraEtiquetaConFoco = &H80000012
Const m_def_ColorLetraEtiquetaSinFoco = &H80000012
Const m_def_ColorLetraTextoConFoco = &H80000012
Const m_def_ColorLetraTextoSinFoco = &H80000012
Const m_def_LongitudEtiqueta = 1035
Const m_def_LongitudTexto = 1035
Const m_def_LongitudDescripcion = 1035
Const m_def_Movible = True
Const m_def_Mascara = ""

'Propiedades:
Dim m_TipoLetra As TiposLetra
Dim m_Formato As Formatos
Dim m_Longitud As Integer
Dim m_Texto As String
Dim m_VALOR As Variant
Dim m_Etiqueta As String
Dim m_AlineacionEtiqueta As TiposAlineacion
Dim m_PosicionEtiqueta As TiposPosicion
Dim m_Tip As String
Dim m_TablaBD As String
Dim m_CampoBD As String
Dim m_CampoBDDescripcion As String
Dim m_Descripcion As String
Dim m_EtiquetaVisible As Boolean
Dim m_TextoVisible As Boolean
Dim m_BotonVisible As Boolean
Dim m_DescripcionVisible As Boolean
Dim m_Activado As Boolean
Dim m_Visible As Boolean
Dim m_Activar As Boolean
Dim m_Ocultar As Boolean
Dim m_Obligar As Boolean

Dim m_EtiquetaNegrita As Boolean
Dim m_ColorLetraEtiquetaConFoco As OLE_COLOR
Dim m_ColorFondoEtiquetaConFoco As OLE_COLOR
Dim m_ColorLetraEtiquetaSinFoco As OLE_COLOR
Dim m_ColorFondoEtiquetaSinFoco As OLE_COLOR
Dim m_ColorLetraTextoConFoco As OLE_COLOR
Dim m_ColorFondoTextoConFoco As OLE_COLOR
Dim m_ColorLetraTextoSinFoco As OLE_COLOR
Dim m_ColorFondoTextoSinFoco As OLE_COLOR
Dim m_ColorLetraDescripcionConFoco As OLE_COLOR
Dim m_ColorFondoDescripcionConFoco As OLE_COLOR
Dim m_ColorLetraDescripcionSinFoco As OLE_COLOR
Dim m_ColorFondoDescripcionSinFoco As OLE_COLOR
Dim m_LongitudEtiqueta As Long
Dim m_LongitudTexto As Long
Dim m_LongitudDescripcion As Long
Dim m_Movible As Boolean
Dim m_Mascara As String
'Default Property Values:
Const m_def_ValorN = 0
Const m_def_hWnd = 0
Const m_def_IdEtiqueta = ""
Const m_def_IdTip = ""
Const m_def_IdToolTipText = ""
Const m_def_Mayusculas = 0
Const m_def_FechaConta = False
Const m_def_FechaIva = False
'Property Variables:
Dim m_ValorN As Variant
Dim m_hWnd As Long
Dim m_IdEtiqueta As String
Dim m_IdTip As String
Dim m_IdToolTipText As String
Dim m_Mayusculas As FormatosCadenas
Dim m_FechaConta As Boolean
Dim m_FechaIva As Boolean


'PROCEDIMIENTOS GENERALES:
Public Sub PrepararFormulario(Frm) 'Interno
    Dim DataArray()
    Dim ctrl As Control
    Dim I As Integer
    Dim j As Integer
    Dim TEMP1 As String
    Dim Temp2 As String
    Dim idx As Integer
    'Variables iniciales:
    idx = 0
    'Establecer nuevo TabIndex para CAMPOS:
    'Llenar Array con controles "Campo".
    I = 0
    For Each ctrl In Frm
        If TypeOf ctrl Is Campo Then
            ReDim Preserve DataArray(2, I)
            DataArray(0, I) = ctrl.Container.Name + "," + Format(ctrl.Top, "00000000") + "," + Format(ctrl.Left, "00000000")
            DataArray(1, I) = ctrl.Name
            I = I + 1
        End If
    Next
    'Ordenación.
    If I > 0 Then
        For I = 1 To UBound(DataArray, 2)
            For j = UBound(DataArray, 2) To I Step -1
                If DataArray(0, j) < DataArray(0, j - 1) Then
                    TEMP1 = DataArray(0, j - 1)
                    Temp2 = DataArray(1, j - 1)
                    DataArray(0, j - 1) = DataArray(0, j)
                    DataArray(1, j - 1) = DataArray(1, j)
                    DataArray(0, j) = TEMP1
                    DataArray(1, j) = Temp2
                End If
            Next
        Next
        'Establecer Tabindex.
        For I = 0 To UBound(DataArray, 2)
            For Each ctrl In Frm
                If TypeOf ctrl Is Campo Then
                    If ctrl.Name = DataArray(1, I) Then ctrl.TabIndex = idx: idx = idx + 1
                End If
            Next
        Next
    End If
End Sub

Private Function Ficha(Contenedor As Control) As Control
    If InStr(UCase(Contenedor.Name), "SSTAB") Or TypeOf Contenedor Is Form Then
        Set Ficha = Contenedor
        Exit Function
    Else
        Set Ficha = Ficha(Contenedor.Container)
    End If
End Function

Private Function NumeroDeFicha(Cadena As String) As Integer
    Dim I As Integer
    Dim HayNumero As Boolean
 
    HayNumero = False
    For I = 1 To Len(Cadena)
        If IsNumeric(Mid(Cadena, I, 1)) Then HayNumero = True: Exit For
    Next

    If HayNumero Then
        NumeroDeFicha = Val(Mid(Cadena, I, 1))
    Else
        NumeroDeFicha = 0
    End If
End Function

Private Sub PosicionarEtiqueta(X, Y) 'Interno del control
    If m_PosicionEtiqueta <> TPArriba Then
        LblEtiquetaFnd.Left = X: LblEtiquetaFnd.Top = Y: LblEtiquetaFnd.Width = LblEtiqueta.Width
    Else
        LblEtiquetaFnd.Left = -888: LblEtiquetaFnd.Top = -888
    End If
    LblEtiqueta.Left = X: LblEtiqueta.Top = Y + 32
End Sub

Private Sub PosicionarTexto(X, Y)
    Dim longitud As Long
        
    If m_Formato <> FChequeo And m_Formato <> FOpcion Then
        longitud = Txt.Width + IIf(m_Formato = FPorcentaje, LblPor.Width, 0) + (16 * 2) + 16
        
        LinTxtArr.X1 = X: LinTxtArr.X2 = X + longitud: LinTxtArr.Y1 = Y: LinTxtArr.Y2 = Y
        LinTxtAbj.Left = X: LinTxtAbj.Width = longitud: LinTxtAbj.Top = UserControl.Height - 16
        LinTxtIzq.X1 = X: LinTxtIzq.X2 = X: LinTxtIzq.Y1 = Y: LinTxtIzq.Y2 = Y + UserControl.Height
        LinTxtDer.X1 = X + longitud - 16: LinTxtDer.X2 = X + longitud - 16: LinTxtDer.Y1 = Y: LinTxtDer.Y2 = Y + UserControl.Height
        
        CheckMark.Visible = False
        OptionMark.Visible = False
        Txt.Left = X + 16 + 16: Txt.Top = Y + 16 + 16
        If m_Formato = FPorcentaje Then LblPor.Left = X + Txt.Width + 16 + 16: LblPor.Top = Y + 16 + 16
    Else
        X = X - 16
        Txt.Left = -888: Txt.Top = -888: Txt.Width = LblCheck.Width
        longitud = LblCheck.Width + (16 * 2)
        Y = Y + (UserControl.Height / 6)
        
        LinTxtArr.X1 = X: LinTxtArr.X2 = X + longitud: LinTxtArr.Y1 = Y: LinTxtArr.Y2 = Y
        LinTxtAbj.Left = X: LinTxtAbj.Width = longitud: LinTxtAbj.Top = Y + LblCheck.Height + 16
        LinTxtIzq.X1 = X: LinTxtIzq.X2 = X: LinTxtIzq.Y1 = Y: LinTxtIzq.Y2 = Y + LblCheck.Height + 16
        LinTxtDer.X1 = X + longitud - 16: LinTxtDer.X2 = X + longitud - 16: LinTxtDer.Y1 = Y: LinTxtDer.Y2 = Y + LblCheck.Height + 16
        
        LblCheck.Left = X + 16: LblCheck.Top = Y + 16
        If m_Formato = FChequeo Then CheckMark.Left = X + 16: CheckMark.Top = Y + 16
        If m_Formato = FOpcion Then OptionMark.Left = X + 16: OptionMark.Top = Y + 16
    End If
End Sub

Private Sub PosicionarDescripcion(X, Y)
    LinDesArr.X1 = X: LinDesArr.X2 = X + LblDescripcion.Width: LinDesArr.Y1 = Y: LinDesArr.Y2 = Y
    If m_PosicionEtiqueta = TPArriba Then
        LinDesAbj.X1 = X: LinDesAbj.X2 = X + LblDescripcion.Width: LinDesAbj.Y1 = Y + LblDescripcion.Height - 16: LinDesAbj.Y2 = Y + LblDescripcion.Height - 16
    Else
        LinDesAbj.X1 = X: LinDesAbj.X2 = X + LblDescripcion.Width: LinDesAbj.Y1 = Y + UserControl.Height - 16: LinDesAbj.Y2 = Y + UserControl.Height - 16
    End If
    LinDesIzq.X1 = X: LinDesIzq.X2 = X: LinDesIzq.Y1 = Y: LinDesIzq.Y2 = Y + UserControl.Height
    'LinDesDer.x1 = X + lblDescripcion.Width - 32: LinDesDer.X2 = X + lblDescripcion.Width - 32: LinDesDer.y1 = Y: LinDesDer.y2 = Y + UserControl.Height '- 16
    LinDesDer.X1 = X + LblDescripcion.Width - 16: LinDesDer.X2 = X + LblDescripcion.Width - 16: LinDesDer.Y1 = Y: LinDesDer.Y2 = Y + UserControl.Height '- 16
    LblDescripcion.Left = X + 16: LblDescripcion.Top = Y + 16: LblDescripcion.Width = LblDescripcion.Width - (16 * 2)
End Sub


Private Sub Boton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent BotonClick(Button)
End Sub

Private Sub CheckMark_Click()
    AlgunCambio = True
    Call Txt_KeyPress(32)
End Sub

Private Sub LblCheck_Click()
    Dim ctrl As Control
    Dim Contenedor As String
    Dim Cancel As Boolean
    
    Cancel = False

    If m_Formato = FChequeo Then
        AlgunCambio = True
        Call Txt_KeyPress(32)
    End If
    If m_Formato = FOpcion Then
        If OptionMark.Visible Then Exit Sub
        AlgunCambio = True
        Call Txt_KeyPress(32)
    End If
End Sub

Private Sub LblEtiqueta_Click()
    Dim ctrl As Control
    Dim Contenedor As String
    Dim Cancel As Boolean
    
    Cancel = False
    If m_Formato = FChequeo Then
         Call Txt_KeyPress(32)
    End If
    If m_Formato = FOpcion Then
        Call Txt_KeyPress(32)
    End If
End Sub

Private Sub OptionMark_Click()
        Call Txt_KeyPress(32)
End Sub

'EVENTOS CONTROL TEXTO:
Private Sub Txt_Change()
    'Ejecutar el Change "externo".
    Dim PosCursor As Integer 'Posicion del cursor dentro del texto
     PosCursor = Txt.SelStart
    If TFC = ForzarMinusculas Then
        Txt = StrConv(Txt, vbLowerCase)
    ElseIf TFC = ForzarMayusculas Then
        Txt = StrConv(Txt, vbUpperCase)
    ElseIf TFC = PrimeraMayuscula Then
        Txt = StrConv(Txt, vbProperCase)
    End If
    RaiseEvent TxtChange
    Txt.SelStart = PosCursor
End Sub

Private Sub Txt_GotFocus()
    Dim A As Integer
    Dim NumFicha As Integer
    Dim tmp As Control
    Dim nWidth, nHeight
    'Mostrar el Tip.
    On Error Resume Next
    UserControl.Parent.Status = " " + m_Tip
    On Error GoTo 0
    
    OldValue = Txt.Text
    Validado = False
    SeHaEditado = False
    'Ejecutar el GotFocus "externo".
    RaiseEvent TxtGotFocus
    'Foco.
    Txt.SelStart = 0
    Txt.SelLength = Len(Txt)
    LblEtiqueta.ForeColor = m_ColorLetraEtiquetaConFoco
    Txt.ForeColor = m_ColorLetraTextoConFoco: LblPor.ForeColor = m_ColorLetraTextoConFoco
    LblDescripcion.ForeColor = m_ColorLetraDescripcionConFoco
    LblEtiqueta.BackColor = m_ColorFondoEtiquetaConFoco
    LblEtiquetaFnd.BackColor = m_ColorFondoEtiquetaConFoco
    Txt.BackColor = m_ColorFondoTextoConFoco: LblPor.BackColor = m_ColorFondoTextoConFoco
    LblCheck.BackColor = m_ColorFondoTextoConFoco
    LblDescripcion.BackColor = m_ColorFondoDescripcionConFoco
    'Si está en un SSTab, cambiar Tab si procede a ello.
    If Not TypeOf UserControl.Extender.Container Is Form Then
        If UCase(Left(UserControl.Extender.Container.Name, 8)) <> "FRAMETAB" Then Exit Sub
        Set tmp = Ficha(UserControl.Extender.Container)
        If InStr(UCase(tmp.Name), "SSTAB") Then
            NumFicha = NumeroDeFicha(UserControl.Extender.Container.Name)
'            For A = NumFicha To tmp.Tabs - 1
'                If tmp.TabVisible(NumFicha) Then Exit For
'                NumFicha = A + 1
'            Next
            If NumFicha > tmp.Tabs - 1 Then NumFicha = 0
            If tmp.Tab <> NumFicha Then
               If tmp.TabVisible(NumFicha) Then
               tmp.Tab = NumFicha
               End If
            End If
        End If
    End If
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim AltDown As Boolean
    Dim CtrlDown As Boolean
    Dim ShiftDown As Boolean
    Dim I As Integer
    Dim NumTwips As Long
    
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    ShiftDown = (Shift And vbShiftMask) > 0
    'Tecla "Shift"+"Arriba": Convertir a Mayúsculas.
    'Tecla "Shift"+"Abajo" : Convertir a Minúsculas.
    If ShiftDown And (KeyCode = 38 Or KeyCode = 40) Then
        ''151202 vbProperCase en el campo
        If TFC = Normal1 Then
            I = Txt.SelStart
                If qTFC = Normal1 Then
                    qTFC = ForzarMinusculas
                End If
                
                If qTFC = ForzarMinusculas Then
                    'se pone en minusculas
                    Txt = StrConv(Txt, vbLowerCase)
                    If KeyCode = 38 Then
                        'se pasa qTFC a ForzarMayusculas
                        qTFC = ForzarMayusculas
                        
                    ElseIf KeyCode = 40 Then
                        'se pasa qTFC a PrimeraMayuscula
                        qTFC = PrimeraMayuscula
                    End If
                ElseIf qTFC = ForzarMayusculas Then
                    'se pone a mayusculas
                    Txt = StrConv(Txt, vbUpperCase)
                    If KeyCode = 38 Then
                        'se pasa qTFC a PrimeraMayuscula
                        qTFC = PrimeraMayuscula
                    ElseIf KeyCode = 40 Then
                        'se pasa qTFC a ForzarMinusculas
                        qTFC = ForzarMinusculas
                    End If
                ElseIf qTFC = PrimeraMayuscula Then
                    'se pone a PrimeraMayuscula
                    Txt = StrConv(Txt, vbProperCase)
                    If KeyCode = 38 Then
                        'se pasa qTFC a ForzarMinusculas
                        qTFC = ForzarMinusculas
                    ElseIf KeyCode = 40 Then
                        'se pasa qTFC a ForzarMayusculas
                        qTFC = ForzarMayusculas
                    End If
                End If
                
                'If KeyCode = 38 Then Txt = UCase(Txt)
                'If KeyCode = 40 Then Txt = LCase(Txt)
            Txt.SelStart = I
        End If
        '151202 vbProperCase en el campo fin
        KeyCode = 0: Exit Sub
    End If
    'Ejecutar el KeyDown "externo".
    RaiseEvent TxtKeyDown(KeyCode, Shift)
    'Tecla "Izquierda":
    If KeyCode = 37 And Txt.SelStart = 0 And Txt.SelLength = 0 Then KeyCode = 0: MySendKeys "+{TAB}", True: Exit Sub
    'Tecla "Arriba":
    If KeyCode = 38 Then
        If UserControl.Extender.TabIndex <> 0 Then
            KeyCode = 0: MySendKeys "+{TAB}", True: Exit Sub
        Else
            KeyCode = 0
            Txt.SelStart = 0
            Txt.SelLength = Len(Txt)
            'Txt.SelStart = Len(Txt)
        End If
    End If
    'Tecla "Derecha":
    If KeyCode = 39 And (Txt.SelStart = Len(Txt) Or Txt.SelLength = Len(Txt)) Then KeyCode = 0: MySendKeys "{TAB}", True: Exit Sub
    'Tecla  "Abajo":
    If KeyCode = 40 Then KeyCode = 0: MySendKeys "{TAB}", True: Exit Sub
End Sub


Private Sub Txt_KeyPress(KeyAscii As Integer)
    '161019 Campo activado
'    If Not m_Activado Then
'        If KeyAscii = 13 Then MySendKeys "{TAB}", True
'        KeyAscii = 0
'        Exit Sub
'    End If
    
    Dim Contenedor As String
    Dim ctrl As Control
    Dim Cancel As Boolean

    If Acccion = "CONSULTANDO" Then
        If KeyAscii = 13 Then
            KeyAscii = 0
            MySendKeys "{TAB}", True
        Else
            KeyAscii = 0
        End If
        Exit Sub
    End If
    'Tecla "*" para consultas:
    If m_BotonVisible And KeyAscii = Asc("*") Then
        KeyAscii = 0
        Accion = "consultando"
        RaiseEvent BotonClick(1)
        Exit Sub
    End If
    If m_BotonVisible And KeyAscii = 10 Then
        RaiseEvent TxtKeyPress(KeyAscii)
        'KeyAscii = 0
        Exit Sub
    End If
    Validado = (KeyAscii = 13) 'Or (KeyAscii = 9)
    'Marcado de Option y Check.
    If KeyAscii = 32 And (m_Formato = FChequeo Or m_Formato = FOpcion) Then
        If m_Formato = FChequeo Then AlgunCambio = True
        Cancel = False
        If m_Formato = FOpcion And OptionMark.Visible Then
            'esto genera el pulsar enter
            KeyAscii = 13
            RaiseEvent TxtKeyPress(KeyAscii)
            'MySendKeys "{TAB}", True
            Exit Sub
        End If
        RaiseEvent CheckClick(Cancel)
        AlgunCambio = True
        If Not Cancel Then
            If m_Formato = FChequeo Then
                CheckMark.Visible = Not CheckMark.Visible
                'esto genera el pulsar enter
                KeyAscii = 13
                RaiseEvent TxtKeyPress(KeyAscii)
                MySendKeys "{TAB}", True
                KeyAscii = 0
                Exit Sub
            End If
            If m_Formato = FOpcion Then
                OptionMark.Visible = Not OptionMark.Visible
                If OptionMark.Visible Then
                    Contenedor = UserControl.Extender.Container.Name
                    For Each ctrl In UserControl.Parent
                        If TypeOf ctrl Is Campo Then
                            If ctrl.Name <> UserControl.Extender.Name And ctrl.Container.Name = Contenedor And ctrl.Formato = FOpcion Then ctrl.valor = False
                        End If
                    Next
                'esto genera el pulsar enter
                RaiseEvent CheckClick(Cancel)
                KeyAscii = 13
                RaiseEvent TxtKeyPress(KeyAscii)
                End If
            End If
        End If
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 13 Or KeyAscii = 9 Then
        Call Txt_Validate(True)
    End If
    'If KeyAscii = 13 Then Call FormatearTxt
    
    '141213 VARIABLE gstrCadenaAComprobar
    If UCase(CampoBD) = gstrCadenaAComprobar Then
        MiSql = MiSql
    End If
    
    If Validado And Txt.Text <> OldValue Then m_AlgunCambio = True
    
    'Ejecutar el Keypress "externo".
    RaiseEvent TxtKeyPress(KeyAscii)
    'Filtrado de caracteres comunes.
    If KeyAscii = 8 Then Exit Sub
    'correcion validado
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Validado Then
        MySendKeys "{TAB}", True
    End If
    Exit Sub
    End If
    'Filtrado de caracteres según Formato.
    Select Case m_Formato
        Case FCadena
            If m_Mascara <> "" Then
                If Mid(m_Mascara, 1, 1) = ">" Then KeyAscii = Asc(LCase(Chr(KeyAscii)))
                If Mid(m_Mascara, 1, 1) = "<" Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Len(m_Mascara) > 1 Then
                    If InStr(m_Mascara, Chr(KeyAscii)) = 0 Then KeyAscii = 0
                End If
            End If
            Exit Sub
        Case FCodigo, FCodCli
            '02/05/14 No pegaba en los campos de codigo OJO
            If KeyAscii = 22 Then Exit Sub
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        Case FFecha, FHora
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
                If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
                If Len(Txt) = 2 Then Txt.SelStart = Len(Txt): Txt = Txt + IIf(m_Formato = FFecha, "/", ":"): Txt.SelStart = Len(Txt)
                If Len(Txt) = 5 Then Txt.SelStart = Len(Txt): Txt = Txt + IIf(m_Formato = FFecha, "/", ":"): Txt.SelStart = Len(Txt)
                Exit Sub
            End If
        Case FSubcuenta
            If Txt <> "" And KeyAscii = Asc(".") And InStr(Txt, ".") = 0 Then Exit Sub
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        Case FCantidad, FMoneda, FPorcentaje, FNumerico, FTotales
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
                If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
                If Len(Txt) > 0 Then
                    If Mid(Txt, Txt.SelStart + 1, 1) <> "-" Then Exit Sub
                Else
                    Exit Sub
                End If
            End If
            If SignoDecimal = "," Then
                If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
            End If
            If KeyAscii = Asc("-") Then
                If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
                If Txt.SelStart = 0 And InStr(1, Txt, "-") = 0 Then Exit Sub
            End If
            
            If SignoDecimal = "," Then

                If KeyAscii = Asc(",") And InStr(m_Mascara, ".") = 0 And m_Mascara <> "" Then KeyAscii = 0
                If KeyAscii = Asc(",") Then
                    If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
                    If InStr(1, Txt, ",") = 0 Then
                        If Len(Txt) > 0 Then
                            If Mid(Txt, Txt.SelStart + 1, 1) <> "-" Then Exit Sub
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            Else
                If KeyAscii = Asc(".") And InStr(m_Mascara, ".") = 0 And m_Mascara <> "" Then KeyAscii = 0
                If KeyAscii = Asc(".") Then
                    If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
                    If InStr(1, Txt, ",") = 0 Then
                        If Len(Txt) > 0 Then
                            If Mid(Txt, Txt.SelStart + 1, 1) <> "-" Then Exit Sub
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Case FCodObra
            If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) And KeyAscii = Asc(".") Then KeyAscii = 0: Exit Sub
            If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
            Dim HayPunto As Integer
            Dim PosPunto As Integer
            Dim ZZ As Integer
            HayPunto = 0
            For ZZ = 1 To Len(Txt)
                If Asc(Mid(Txt, ZZ, 1)) = Asc(".") Then
                    HayPunto = HayPunto + 1
                    PosPunto = ZZ
                End If
            Next
            If KeyAscii = Asc(".") Then
                If PosPunto = Len(Txt) Then KeyAscii = 0: Exit Sub
            End If
            If Len(Txt) = PosPunto + 3 Then
                If HayPunto = 2 Then KeyAscii = 0: Exit Sub
                Txt = Txt & ".": Txt.SelStart = Len(Txt)
            End If
            Exit Sub
        Case FTelefono
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
            
    End Select
    KeyAscii = 0
End Sub

Private Sub Txt_LostFocus()
    If UCase(CampoBD) = gstrCadenaAComprobar Then
        MiSql = MiSql
    End If
    
    If Txt.Text <> OldValue Then m_AlgunCambio = True
    'Limpiar el Tip.
    On Error Resume Next
    UserControl.Parent.Status = ""
    On Error GoTo 0
    'If m_BotonVisible And Not Validado Then Txt.Text = OldValue
    If Txt.Text = "" Then LblDescripcion = "": m_Descripcion = ""
    'Ejecutar el LostFocus "externo".
    RaiseEvent TxtLostFocus
    'Call FormatearTxt
    'Fuera Foco.
    LblEtiqueta.ForeColor = m_ColorLetraEtiquetaSinFoco
    Select Case m_Formato
        Case FFecha, FHora
'            If Txt <> "" And Not IsDate(Txt) Then
'                MensajeError "Introduzca una fecha valida"
'                Txt = ""
'                Txt.SetFocus
'                Exit Sub
'                Txt.ForeColor = QBColor(12): LblPor.ForeColor = QBColor(12)
'            Else
'                Txt.ForeColor = m_ColorLetraTextoSinFoco: LblPor.ForeColor = m_ColorLetraTextoSinFoco
'            End If
        Case FCodigo, FSubcuenta, FCantidad, FMoneda, FPorcentaje, FNumerico, FCodCli
            If Txt <> "" And Not IsNumeric(Txt) Then
                Txt.ForeColor = QBColor(12): LblPor.ForeColor = QBColor(12)
            Else
                Txt.ForeColor = m_ColorLetraTextoSinFoco: LblPor.ForeColor = m_ColorLetraTextoSinFoco
            End If
    End Select
    LblDescripcion.ForeColor = m_ColorLetraDescripcionSinFoco
    LblEtiqueta.BackColor = m_ColorFondoEtiquetaSinFoco
    LblEtiquetaFnd.BackColor = m_ColorFondoEtiquetaSinFoco
    Txt.BackColor = m_ColorFondoTextoSinFoco: LblPor.BackColor = m_ColorFondoTextoSinFoco
    LblCheck.BackColor = m_ColorFondoTextoSinFoco
    LblDescripcion.BackColor = m_ColorFondoDescripcionSinFoco
End Sub

Private Sub RefrescarControl()
    On Error Resume Next
    If m_Formato = FCantidad Or m_Formato = FMoneda Or m_Formato = FPorcentaje Or m_Formato = FNumerico Or m_Formato = FTotales Then
        If m_TextoVisible Then
            Txt.Alignment() = 1
        Else
            LblDescripcion.Alignment() = 1
        End If
    Else
        If m_TextoVisible Then
            Txt.Alignment() = 0
        Else
            LblDescripcion.Alignment() = 0
        End If
    End If
    If Mascara = "*" Then Txt.PasswordChar = "*"
    Txt.MaxLength() = m_Longitud
    Txt.Text() = m_Texto
    LblEtiqueta() = m_Etiqueta
    LblEtiqueta.Alignment() = m_AlineacionEtiqueta
    LblDescripcion() = m_Descripcion
    LblEtiqueta.Visible() = (m_Visible And m_EtiquetaVisible)
    LblEtiquetaFnd.Visible() = (m_Visible And m_EtiquetaVisible)
    LinTxtArr.Visible() = (m_Visible And m_TextoVisible)
    LinTxtAbj.Visible() = (m_Visible And m_TextoVisible)
    LinTxtIzq.Visible() = (m_Visible And m_TextoVisible)
    LinTxtDer.Visible() = (m_Visible And m_TextoVisible)
    Txt.Visible() = (m_Visible And m_TextoVisible)
    LblCheck.Visible() = (m_Visible And m_TextoVisible And (m_Formato = FChequeo Or m_Formato = FOpcion))
    LblPor.Visible() = (m_Visible And m_TextoVisible And (m_Formato = FPorcentaje))
    Boton.Visible() = (m_Visible And m_BotonVisible)
    LinDesArr.Visible() = (m_Visible And m_DescripcionVisible)
    LinDesAbj.Visible() = (m_Visible And m_DescripcionVisible)
    LinDesIzq.Visible() = (m_Visible And m_DescripcionVisible)
    LinDesDer.Visible() = (m_Visible And m_DescripcionVisible)
    LblDescripcion.Visible() = (m_Visible And m_DescripcionVisible)
    LblEtiqueta.Enabled() = m_Activado
    'Txt.Enabled() = m_Activado:
    LblPor.Enabled() = m_Activado: Boton.Enabled = m_Activado: LblDescripcion.Enabled() = m_Activado
    '161019 Campo activado
    UserControl.Enabled = m_Activado
    
    LblEtiqueta.Font.Bold() = m_EtiquetaNegrita
    LblEtiqueta.ForeColor() = m_ColorLetraEtiquetaSinFoco
    Txt.ForeColor() = m_ColorLetraTextoSinFoco
    LblPor.ForeColor() = m_ColorLetraTextoSinFoco
    LblDescripcion.ForeColor() = m_ColorLetraDescripcionSinFoco
    LblEtiqueta.BackColor() = m_ColorFondoEtiquetaSinFoco
    LblEtiquetaFnd.BackColor() = m_ColorFondoEtiquetaSinFoco
    Txt.BackColor() = m_ColorFondoTextoSinFoco
    LblCheck.BackColor() = m_ColorFondoTextoSinFoco
    LblPor.BackColor() = m_ColorFondoTextoSinFoco
    LblDescripcion.BackColor() = m_ColorFondoDescripcionSinFoco
    LblEtiqueta.Width() = m_LongitudEtiqueta
    Txt.Width() = m_LongitudTexto
    LblDescripcion.Width() = m_LongitudDescripcion
    On Error GoTo 0
    Call UserControl_Resize
End Sub

Private Sub Txt_Validate(Cancel As Boolean)
    Dim cad1 As String, cad2 As String, cad3 As String
    Select Case m_Formato
        Case FCadena, FTelefono
            Txt = Trim(Txt)
        Case FCodigo, FCodCli
            Txt = Format(Txt, String(Txt.MaxLength, "0"))
            Txt.Refresh
            'DoEvents
        Case FFecha ', FHora
            If Len(Txt) = 8 Then Txt = Mid(Txt, 1, 6) & Mid(Format(Date, "yyyy"), 1, 2) & Mid(Txt, 7)
            If Len(Txt) = 9 Then Txt = Mid(Txt, 1, 6) & Mid(Format(Date, "yyyy"), 1, 1) & Mid(Txt, 7)
            If Len(Txt) = 6 Then Txt = Txt & Format(Date, "yyyy")
            '141217 modificado el campo, para controlar la fecha, petaba con la fecha 31/11/
            If Len(Txt) = 5 Then Txt = Txt & "/" & Format(Date, "yyyy")
            If Txt <> "" And IsDate(Txt) Then Txt = CDate(Txt)
            If Txt <> "" And Not IsDate(Txt) Then
                'info0127=Introduzca una fecha valida
                't = IdActual.Recuperar("info0127")
                T = "Introduzca una fecha valida"
                'MensajeError T
                Msg mError, T
                If OldValue <> "" Then
                    Txt = OldValue
                Else
                    Txt = ""
                End If
                Txt.SetFocus
                Validado = False
                Cancel = True
                Exit Sub
            End If
        Case FHora
            If Len(Txt) = 1 Or Len(Txt) = 2 Then Txt = Txt & ":00:00"
            
            If Txt <> "" And Not IsDate(Txt) Then
                'info0128=Introduzca un valor valido
                't = IdActual.Recuperar("info0128")
                T = "Introduzca un valor valido"
                
                
                'MensajeError T
                Msg mError, T
                If OldValue <> "" Then
                    Txt = OldValue
                Else
                    Txt = ""
                End If
                Txt.SetFocus
                Validado = False
                Cancel = True
                Exit Sub
            End If
           
        Case FSubcuenta
            If Txt <> "" Then
                If Len(Txt) < Txt.MaxLength And InStr(Txt, ".") = 0 Then
                    Txt = Txt + String(Txt.MaxLength - Len(Txt), "0")
                ElseIf InStr(Txt, ".") Then
                    cad1 = Mid(Txt, 1, InStr(Txt, ".") - 1)
                    cad2 = Mid(Txt, InStr(Txt, ".") + 1)
                    Txt = cad1 + String(Txt.MaxLength - (Len(cad1) + Len(cad2)), "0") + cad2
                End If
            End If
        Case FCantidad, FMoneda, FPorcentaje, FNumerico, FTotales
            If IsNumeric(Txt) Then
                LblPor.ForeColor = m_ColorLetraTextoSinFoco
                Txt.ForeColor = m_ColorLetraTextoSinFoco
                If m_Mascara = "" Then
                    If CDbl(Txt) - Fix(CDbl(Txt)) <> 0 Then
                        Txt = Format(Txt, "#,##0.############")
                    Else
                        Txt = Format(Txt, "#,##0")
                    End If
                Else
                    Txt = Format(Txt, m_Mascara)
                End If
            End If
        Case FCodObra
            Dim Pu1 As Integer
            Dim Pu2 As Integer
            Dim Po1 As Integer
            Dim Po2 As Integer
            Dim ZZ As Integer
            If Txt = "" Then Exit Sub
            
            For ZZ = 1 To Len(Txt)
                If Asc(Mid(Txt, ZZ, 1)) = Asc(".") Then
                    If Po1 = 0 Then
                        Po1 = ZZ
                    Else
                        Po2 = ZZ
                    End If
                End If
            Next
            If Po1 = 0 Then
                Txt = String(9 - Len(Txt), "0") & Txt
                Po1 = 4
                Po2 = 8
            End If
            If Po1 <> 0 Then
                cad1 = Mid(Txt, 1, Po1 - 1)
            Else
                cad1 = Txt
            End If
            If Len(cad1) = 1 Then
                cad1 = "00" & cad1
            ElseIf Len(cad1) = 2 Then
                cad1 = "0" & cad1
            ElseIf Len(cad1) = 0 Then
                cad1 = "000"
            End If
            If Po2 <> 0 Then
            cad2 = Mid(Txt, Po1 + 1, (Po2 - Po1) - 1)
            Else
                If Po1 <> 0 Then
                cad2 = Mid(Txt, Po1 + 1)
                Else
                cad2 = "000"
                End If
            End If
            If Len(cad2) = 1 Then
                cad2 = "00" & cad2
            ElseIf Len(cad2) = 2 Then
                cad2 = "0" & cad2
            ElseIf Len(cad2) = 0 Then
                cad2 = "000"
            End If
            
            If Po2 <> 0 Then
                cad3 = Mid(Txt, Po2 + 1)
            Else
                cad3 = "000"
            End If
            If Len(cad3) = 1 Then
                cad3 = "00" & cad3
            ElseIf Len(cad3) = 2 Then
                cad3 = "0" & cad3
            ElseIf Len(cad3) = 0 Then
                cad3 = "000"
            End If
            
            Txt = cad1 & "." & cad2 & "." & cad3
    End Select
'Validado = True
End Sub

Private Sub UserControl_ExitFocus()
    If SeHaEditado Then
        PrepararFormulario UserControl.Parent
    End If
    SeHaEditado = False
    Editando = False
    Marquita.Visible = False
End Sub
'I N I C I A L I Z A R   C O N T R O L:
Private Sub UserControl_Initialize()
    Call QueSigno
    m_TipoLetra = m_def_TipoLetra
    m_Formato = m_def_Formato
    m_Longitud = m_def_Longitud
    m_Texto = m_def_Texto
    m_Etiqueta = m_def_Etiqueta
    m_AlineacionEtiqueta = m_def_AlineacionEtiqueta
    m_PosicionEtiqueta = m_def_PosicionEtiqueta
    m_Tip = m_def_Tip
    m_TablaBD = m_def_TablaBD
    m_CampoBD = m_def_CampoBD
    m_CampoBDDescripcion = m_def_CampoBDDescripcion
    m_Descripcion = m_def_Descripcion
    m_EtiquetaVisible = m_def_EtiquetaVisible
    m_TextoVisible = m_def_TextoVisible
    m_BotonVisible = m_def_BotonVisible
    m_DescripcionVisible = m_def_DescripcionVisible
    m_Activado = m_def_Activado
    m_Visible = m_def_Visible
    m_Activar = m_def_Activar
    m_Ocultar = m_def_Ocultar
    m_Obligar = m_def_Obligar
    m_EtiquetaNegrita = m_def_EtiquetaNegrita
    m_ColorLetraEtiquetaConFoco = m_def_ColorLetraEtiquetaConFoco
    m_ColorFondoEtiquetaConFoco = m_def_ColorFondoEtiquetaConFoco
    m_ColorLetraEtiquetaSinFoco = m_def_ColorLetraEtiquetaSinFoco
    m_ColorFondoEtiquetaSinFoco = m_def_ColorFondoEtiquetaSinFoco
    m_ColorLetraTextoConFoco = m_def_ColorLetraTextoConFoco
    m_ColorFondoTextoConFoco = m_def_ColorFondoTextoConFoco
    m_ColorLetraTextoSinFoco = m_def_ColorLetraTextoSinFoco
    m_ColorFondoTextoSinFoco = m_def_ColorFondoTextoSinFoco
    m_ColorLetraDescripcionConFoco = m_def_ColorLetraDescripcionConFoco
    m_ColorFondoDescripcionConFoco = m_def_ColorFondoDescripcionConFoco
    m_ColorLetraDescripcionSinFoco = m_def_ColorLetraDescripcionSinFoco
    m_ColorFondoDescripcionSinFoco = m_def_ColorFondoDescripcionSinFoco
    m_LongitudEtiqueta = m_def_LongitudEtiqueta
    m_LongitudTexto = m_def_LongitudTexto
    m_LongitudDescripcion = m_def_LongitudDescripcion
    m_Movible = m_def_Movible
    m_Mascara = m_def_Mascara
    SeHaEditado = False
    Editando = False
    nEditando = 1
    m_AlgunCambio = False
End Sub

Private Sub UserControl_InitProperties()
    Call RefrescarControl
    m_FechaConta = m_def_FechaConta
    m_FechaIva = m_def_FechaIva
    m_Mayusculas = m_def_Mayusculas
    m_hWnd = m_def_hWnd
    m_IdEtiqueta = m_def_IdEtiqueta
    m_IdTip = m_def_IdTip
    m_IdToolTipText = m_def_IdToolTipText
    m_ValorN = m_def_ValorN
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim AltDown As Boolean
    Dim CtrlDown As Boolean
    Dim ShiftDown As Boolean
    Dim I As Integer
    Dim NumTwips As Long
    
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    ShiftDown = (Shift And vbShiftMask) > 0
    'M O V I L I D A D
    If CtrlDown And AltDown And Movible Then
        If Editando Then
            nEditando = nEditando + 1
            If m_DescripcionVisible Then
                If nEditando = 6 Then Marquita.Visible = False: Editando = False: Exit Sub
                Select Case nEditando
                    Case 1: Marquita.Left = LblEtiqueta.Left
                    Case 2: Marquita.Left = Txt.Left
                    Case 3: Marquita.Left = Txt.Left + Txt.Width + Marquita.Width - 32
                    Case 4: Marquita.Left = LblDescripcion.Left + LblDescripcion.Width - Marquita.Width - 40
                    Case 5: Marquita.Left = LblEtiqueta.Left: Marquita.Width = UserControl.Width
                End Select
            Else
                If nEditando = 5 Then Marquita.Visible = False: Editando = False: Exit Sub
                Select Case nEditando
                    Case 1: Marquita.Width = 75: Marquita.Left = LblEtiqueta.Left
                    Case 2: Marquita.Left = Txt.Left
                    Case 3: Marquita.Left = Txt.Left + Txt.Width - Marquita.Width
                    Case 4: Marquita.Left = LblEtiqueta.Left: Marquita.Width = UserControl.Width
                End Select
            End If
        Else
            Marquita.Width = 75
            Select Case m_TipoLetra
                Case TLPequeña: Marquita.Top = 160
                Case TLMediana: Marquita.Top = 210
                Case TLGrande: Marquita.Top = 270
            End Select
            Marquita.Visible = True
            Marquita.Left = LblEtiqueta.Left
            Editando = True
            nEditando = 1
            SeHaEditado = True
        End If
    End If
    
    If ShiftDown Then
        NumTwips = 15
    Else
        NumTwips = 60
    End If
    
    If Editando Then
        'Tecla "Izquierda": Mover
        If KeyCode = 37 And CtrlDown And Marquita.Width = UserControl.Width Then
            If UserControl.Extender.Left > NumTwips Then
                UserControl.Extender.Left = UserControl.Extender.Left - NumTwips
            End If
            GoSub MostrarTip: KeyCode = 0: Exit Sub
        End If
        'Tecla "Derecha": Mover
        If KeyCode = 39 And CtrlDown And Marquita.Width = UserControl.Width Then
            If UserControl.Extender.Left + UserControl.Extender.Width <= UserControl.Extender.Container.Width - 120 Then
                UserControl.Extender.Left = UserControl.Extender.Left + NumTwips
            End If
            GoSub MostrarTip: KeyCode = 0: Exit Sub
        End If
        'Tecla "Arriba": Mover
        If KeyCode = 38 And CtrlDown And Marquita.Width = UserControl.Width Then
            If UserControl.Extender.Top > 120 Then
                UserControl.Extender.Top = UserControl.Extender.Top - NumTwips
            End If
            GoSub MostrarTip: KeyCode = 0: Exit Sub
        End If
        'Tecla "Abajo": Mover
        If KeyCode = 40 And CtrlDown And Marquita.Width = UserControl.Width Then
            If UserControl.Extender.Top + UserControl.Extender.Height <= UserControl.Extender.Container.Height - 120 Then
                UserControl.Extender.Top = UserControl.Extender.Top + NumTwips
            End If
            GoSub MostrarTip: KeyCode = 0: Exit Sub
        End If
        'Tecla "Izquierda": Redimensionar
        If KeyCode = 37 And CtrlDown Then
            If m_DescripcionVisible Then
                Select Case nEditando
                    Case 1
                        If UserControl.Extender.Left > NumTwips Then
                            LblEtiqueta.Width = LblEtiqueta.Width + NumTwips
                            UserControl.Width = UserControl.Width + NumTwips
                            UserControl.Extender.Left = UserControl.Extender.Left - NumTwips
                        End If
                    Case 2
                        If LblEtiqueta.Width > 180 Then
                            LblEtiqueta.Width = LblEtiqueta.Width - NumTwips
                            Txt.Width = Txt.Width + NumTwips
                            Txt.Left = Txt.Left - NumTwips
                            Marquita.Left = Marquita.Left - NumTwips
                        End If
                    Case 3
                        If Txt.Width > 180 Then
                            Txt.Width = Txt.Width - NumTwips
                            LblDescripcion.Width = LblDescripcion.Width + NumTwips
                            LblDescripcion.Left = LblDescripcion.Left - NumTwips
                            Marquita.Left = Marquita.Left - NumTwips
                        End If
                    Case 4
                        If LblDescripcion.Width > 180 Then
                            LblDescripcion.Width = LblDescripcion.Width - NumTwips
                            UserControl.Width = UserControl.Width - NumTwips
                            Marquita.Left = Marquita.Left - NumTwips
                        End If
                End Select
            Else
                Select Case nEditando
                    Case 1
                        If UserControl.Extender.Left > NumTwips Then
                            LblEtiqueta.Width = LblEtiqueta.Width + NumTwips
                            UserControl.Width = UserControl.Width + NumTwips
                            UserControl.Extender.Left = UserControl.Extender.Left - NumTwips
                        End If
                    Case 2
                        If LblEtiqueta.Width > 180 Then
                            LblEtiqueta.Width = LblEtiqueta.Width - NumTwips
                            Txt.Width = Txt.Width + NumTwips
                            Txt.Left = Txt.Left - NumTwips
                            Marquita.Left = Marquita.Left - NumTwips
                        End If
                    Case 3
                        If Txt.Width > 180 Then
                            Txt.Width = Txt.Width - NumTwips
                            UserControl.Width = UserControl.Width - NumTwips
                            Marquita.Left = Marquita.Left - NumTwips
                        End If
                End Select
            End If
            GoSub MostrarTip: KeyCode = 0: Call UserControl_Resize: Exit Sub
        End If
        'Tecla "Derecha": Redimensionar
        If KeyCode = 39 And CtrlDown Then
            If m_DescripcionVisible Then
                Select Case nEditando
                    Case 1
                        If LblEtiqueta.Width > 180 Then
                            LblEtiqueta.Width = LblEtiqueta.Width - NumTwips
                            UserControl.Width = UserControl.Width - NumTwips
                            UserControl.Extender.Left = UserControl.Extender.Left + NumTwips
                        End If
                    Case 2
                        If Txt.Width > 180 Then
                            LblEtiqueta.Width = LblEtiqueta.Width + NumTwips
                            Txt.Width = Txt.Width - NumTwips
                            Txt.Left = Txt.Left + NumTwips
                            Marquita.Left = Marquita.Left + NumTwips
                        End If
                    Case 3
                        If LblDescripcion.Width > 180 Then
                            Txt.Width = Txt.Width + NumTwips
                            LblDescripcion.Width = LblDescripcion.Width - NumTwips
                            LblDescripcion.Left = LblDescripcion.Left + NumTwips
                            Marquita.Left = Marquita.Left + NumTwips
                        End If
                    Case 4
                        If UserControl.Extender.Left + UserControl.Extender.Width <= UserControl.Extender.Container.Width - 120 Then
                            LblDescripcion.Width = LblDescripcion.Width + NumTwips
                            UserControl.Width = UserControl.Width + NumTwips
                            Marquita.Left = Marquita.Left + NumTwips
                        End If
                End Select
            Else
                Select Case nEditando
                    Case 1
                        If LblEtiqueta.Width > 180 Then
                            LblEtiqueta.Width = LblEtiqueta.Width - NumTwips
                            UserControl.Width = UserControl.Width - NumTwips
                            UserControl.Extender.Left = UserControl.Extender.Left + NumTwips
                        End If
                    Case 2
                        If Txt.Width > 180 Then
                            LblEtiqueta.Width = LblEtiqueta.Width + NumTwips
                            Txt.Width = Txt.Width - NumTwips
                            Txt.Left = Txt.Left + NumTwips
                            Marquita.Left = Marquita.Left + NumTwips
                        End If
                    Case 3
                        If UserControl.Extender.Left + UserControl.Extender.Width <= UserControl.Extender.Container.Width - 120 Then
                            Txt.Width = Txt.Width + NumTwips
                            UserControl.Width = UserControl.Width + NumTwips
                            Marquita.Left = Marquita.Left + NumTwips
                        End If
                End Select
            End If
            'Mostrar el Tip.
            On Error Resume Next
            UserControl.Parent.Status = " Control: [X: " + Trim(str(UserControl.Extender.Left)) + ", Y: " + Trim(str(UserControl.Extender.Top)) + ", ANCHO: " + Trim(str(UserControl.Width)) + ", ALTO: " + Trim(str(UserControl.Height)) + "]   " + _
            "Etiqueta: [X: " + Trim(str(LblEtiqueta.Left)) + ", Y: " + Trim(str(LblEtiqueta.Top)) + ", ANCHO: " + Trim(str(LblEtiqueta.Width)) + ", ALTO: " + Trim(str(LblEtiqueta.Height)) + "]   " + _
            "Texto: [X: " + Trim(str(Txt.Left)) + ", Y: " + Trim(str(Txt.Top)) + ", ANCHO: " + Trim(str(Txt.Width)) + ", ALTO: " + Trim(str(Txt.Height)) + "]   " + _
            "Descripción: [X: " + Trim(str(LblDescripcion.Left)) + ", Y: " + Trim(str(LblDescripcion.Top)) + ", ANCHO: " + Trim(str(LblDescripcion.Width)) + ", ALTO: " + Trim(str(LblDescripcion.Height)) + "]"
            On Error GoTo 0
            
            GoSub MostrarTip: KeyCode = 0: Call UserControl_Resize: Exit Sub
        End If
    End If
    'F I N   M O V I L I D A D
    Exit Sub
MostrarTip:
    'Mostrar el Tip.
    On Error Resume Next
    UserControl.Parent.Status = " Ctrl: [X: " + Trim(str(UserControl.Extender.Left)) + ", Y: " + Trim(str(UserControl.Extender.Top)) + ", AN: " + Trim(str(UserControl.Width)) + ", AL: " + Trim(str(UserControl.Height)) + "], " + _
    "Etq: [X: " + Trim(str(LblEtiqueta.Left)) + ", Y: " + Trim(str(LblEtiqueta.Top)) + ", AN: " + Trim(str(LblEtiqueta.Width)) + ", AL: " + Trim(str(LblEtiqueta.Height)) + "], " + _
    "Txt: [X: " + Trim(str(Txt.Left)) + ", Y: " + Trim(str(Txt.Top)) + ", AN: " + Trim(str(Txt.Width)) + ", AL: " + Trim(str(Txt.Height)) + "], " + _
    "Dsc: [X: " + Trim(str(LblDescripcion.Left)) + ", Y: " + Trim(str(LblDescripcion.Top)) + ", AN: " + Trim(str(LblDescripcion.Width)) + ", AL: " + Trim(str(LblDescripcion.Height)) + "]"
    On Error GoTo 0
    Return
End Sub
'P R O P I E D A D E S:
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    m_TipoLetra = PropBag.ReadProperty("TipoLetra", m_def_TipoLetra)
    m_Formato = PropBag.ReadProperty("Formato", m_def_Formato)
    m_Longitud = PropBag.ReadProperty("Longitud", m_def_Longitud)
    m_Texto = PropBag.ReadProperty("Texto", m_def_Texto)
    m_Etiqueta = PropBag.ReadProperty("Etiqueta", m_def_Etiqueta)
    m_AlineacionEtiqueta = PropBag.ReadProperty("AlineacionEtiqueta", m_def_AlineacionEtiqueta)
    m_PosicionEtiqueta = PropBag.ReadProperty("PosicionEtiqueta", m_def_PosicionEtiqueta)
    m_Tip = PropBag.ReadProperty("Tip", m_def_Tip)
    m_TablaBD = PropBag.ReadProperty("TablaBD", m_def_TablaBD)
    m_CampoBD = PropBag.ReadProperty("CampoBD", m_def_CampoBD)
    m_CampoBDDescripcion = PropBag.ReadProperty("CampoBDDescripcion", m_def_CampoBDDescripcion)
    m_Descripcion = PropBag.ReadProperty("Descripcion", m_def_Descripcion)
    m_EtiquetaVisible = PropBag.ReadProperty("EtiquetaVisible", m_def_EtiquetaVisible)
    m_TextoVisible = PropBag.ReadProperty("TextoVisible", m_def_TextoVisible)
    m_BotonVisible = PropBag.ReadProperty("BotonVisible", m_def_BotonVisible)
    m_DescripcionVisible = PropBag.ReadProperty("DescripcionVisible", m_def_DescripcionVisible)
    m_Activado = PropBag.ReadProperty("Activado", m_def_Activado)
    m_Visible = PropBag.ReadProperty("Visible", m_def_Visible)
    m_Activar = PropBag.ReadProperty("Activar", m_def_Activar)
    m_Ocultar = PropBag.ReadProperty("Ocultar", m_def_Ocultar)
    m_Obligar = PropBag.ReadProperty("Obligar", m_def_Obligar)
    m_EtiquetaNegrita = PropBag.ReadProperty("EtiquetaNegrita", m_def_EtiquetaNegrita)
    m_ColorLetraEtiquetaConFoco = PropBag.ReadProperty("ColorLetraEtiquetaConFoco", m_def_ColorLetraEtiquetaConFoco)
    m_ColorFondoEtiquetaConFoco = PropBag.ReadProperty("ColorFondoEtiquetaConFoco", m_def_ColorFondoEtiquetaConFoco)
    m_ColorLetraEtiquetaSinFoco = PropBag.ReadProperty("ColorLetraEtiquetaSinFoco", m_def_ColorLetraEtiquetaSinFoco)
    m_ColorFondoEtiquetaSinFoco = PropBag.ReadProperty("ColorFondoEtiquetaSinFoco", m_def_ColorFondoEtiquetaSinFoco)
    m_ColorLetraTextoConFoco = PropBag.ReadProperty("ColorLetraTextoConFoco", m_def_ColorLetraTextoConFoco)
    m_ColorFondoTextoConFoco = PropBag.ReadProperty("ColorFondoTextoConFoco", m_def_ColorFondoTextoConFoco)
    m_ColorLetraTextoSinFoco = PropBag.ReadProperty("ColorLetraTextoSinFoco", m_def_ColorLetraTextoSinFoco)
    m_ColorFondoTextoSinFoco = PropBag.ReadProperty("ColorFondoTextoSinFoco", m_def_ColorFondoTextoSinFoco)
    m_ColorLetraDescripcionConFoco = PropBag.ReadProperty("ColorLetraDescripcionConFoco", m_def_ColorLetraDescripcionConFoco)
    m_ColorFondoDescripcionConFoco = PropBag.ReadProperty("ColorFondoDescripcionConFoco", m_def_ColorFondoDescripcionConFoco)
    m_ColorLetraDescripcionSinFoco = PropBag.ReadProperty("ColorLetraDescripcionSinFoco", m_def_ColorLetraDescripcionSinFoco)
    m_ColorFondoDescripcionSinFoco = PropBag.ReadProperty("ColorFondoDescripcionSinFoco", m_def_ColorFondoDescripcionSinFoco)
    m_LongitudEtiqueta = PropBag.ReadProperty("LongitudEtiqueta", m_def_LongitudEtiqueta)
    m_LongitudTexto = PropBag.ReadProperty("LongitudTexto", m_def_LongitudTexto)
    m_LongitudDescripcion = PropBag.ReadProperty("LongitudDescripcion", m_def_LongitudDescripcion)
    m_Movible = PropBag.ReadProperty("Movible", m_def_Movible)
    m_Mascara = PropBag.ReadProperty("Mascara", m_def_Mascara)
    m_FechaConta = PropBag.ReadProperty("FechaConta", m_def_FechaConta)
    m_FechaIva = PropBag.ReadProperty("FechaIva", m_def_FechaIva)
    m_Mayusculas = PropBag.ReadProperty("Mayusculas", m_def_Mayusculas)
    TFC = m_Mayusculas
    qTFC = m_Mayusculas
    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
    m_IdEtiqueta = PropBag.ReadProperty("IdEtiqueta", m_def_IdEtiqueta)
    m_IdTip = PropBag.ReadProperty("IdTip", m_def_IdTip)
    m_IdToolTipText = PropBag.ReadProperty("IdToolTipText", m_def_IdToolTipText)
    m_ValorN = PropBag.ReadProperty("ValorN", m_def_ValorN)
    On Error GoTo 0
    Call RefrescarControl
    
End Sub

Private Sub UserControl_Show()
    Dim tControl As Control
    Dim sTip$
    sTip = Extender.ToolTipText
    On Local Error Resume Next
    For Each tControl In Controls
        tControl.ToolTipText = sTip
        If Err Then Err = 0
    Next
    On Local Error GoTo 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TipoLetra", m_TipoLetra, m_def_TipoLetra)
    Call PropBag.WriteProperty("Formato", m_Formato, m_def_Formato)
    Call PropBag.WriteProperty("Longitud", m_Longitud, m_def_Longitud)
    Call PropBag.WriteProperty("Texto", m_Texto, m_def_Texto)
    Call PropBag.WriteProperty("Etiqueta", m_Etiqueta, m_def_Etiqueta)
    Call PropBag.WriteProperty("AlineacionEtiqueta", m_AlineacionEtiqueta, m_def_AlineacionEtiqueta)
    Call PropBag.WriteProperty("PosicionEtiqueta", m_PosicionEtiqueta, m_def_PosicionEtiqueta)
    Call PropBag.WriteProperty("Tip", m_Tip, m_def_Tip)
    Call PropBag.WriteProperty("TablaBD", m_TablaBD, m_def_TablaBD)
    Call PropBag.WriteProperty("CampoBD", m_CampoBD, m_def_CampoBD)
    Call PropBag.WriteProperty("CampoBDDescripcion", m_CampoBDDescripcion, m_def_CampoBDDescripcion)
    Call PropBag.WriteProperty("Descripcion", m_Descripcion, m_def_Descripcion)
    Call PropBag.WriteProperty("EtiquetaVisible", m_EtiquetaVisible, m_def_EtiquetaVisible)
    Call PropBag.WriteProperty("TextoVisible", m_TextoVisible, m_def_TextoVisible)
    Call PropBag.WriteProperty("BotonVisible", m_BotonVisible, m_def_BotonVisible)
    Call PropBag.WriteProperty("DescripcionVisible", m_DescripcionVisible, m_def_DescripcionVisible)
    Call PropBag.WriteProperty("Activado", m_Activado, m_def_Activado)
    Call PropBag.WriteProperty("Visible", m_Visible, m_def_Visible)
    Call PropBag.WriteProperty("Activar", m_Activar, m_def_Activar)
    Call PropBag.WriteProperty("Ocultar", m_Ocultar, m_def_Ocultar)
    Call PropBag.WriteProperty("Obligar", m_Obligar, m_def_Obligar)
    Call PropBag.WriteProperty("EtiquetaNegrita", m_EtiquetaNegrita, m_def_EtiquetaNegrita)
    Call PropBag.WriteProperty("ColorLetraEtiquetaConFoco", m_ColorLetraEtiquetaConFoco, m_def_ColorLetraEtiquetaConFoco)
    Call PropBag.WriteProperty("ColorFondoEtiquetaConFoco", m_ColorFondoEtiquetaConFoco, m_def_ColorFondoEtiquetaConFoco)
    Call PropBag.WriteProperty("ColorLetraEtiquetaSinFoco", m_ColorLetraEtiquetaSinFoco, m_def_ColorLetraEtiquetaSinFoco)
    Call PropBag.WriteProperty("ColorFondoEtiquetaSinFoco", m_ColorFondoEtiquetaSinFoco, m_def_ColorFondoEtiquetaSinFoco)
    Call PropBag.WriteProperty("ColorLetraTextoConFoco", m_ColorLetraTextoConFoco, m_def_ColorLetraTextoConFoco)
    Call PropBag.WriteProperty("ColorFondoTextoConFoco", m_ColorFondoTextoConFoco, m_def_ColorFondoTextoConFoco)
    Call PropBag.WriteProperty("ColorLetraTextoSinFoco", m_ColorLetraTextoSinFoco, m_def_ColorLetraTextoSinFoco)
    Call PropBag.WriteProperty("ColorFondoTextoSinFoco", m_ColorFondoTextoSinFoco, m_def_ColorFondoTextoSinFoco)
    Call PropBag.WriteProperty("ColorLetraDescripcionConFoco", m_ColorLetraDescripcionConFoco, m_def_ColorLetraDescripcionConFoco)
    Call PropBag.WriteProperty("ColorFondoDescripcionConFoco", m_ColorFondoDescripcionConFoco, m_def_ColorFondoDescripcionConFoco)
    Call PropBag.WriteProperty("ColorLetraDescripcionSinFoco", m_ColorLetraDescripcionSinFoco, m_def_ColorLetraDescripcionSinFoco)
    Call PropBag.WriteProperty("ColorFondoDescripcionSinFoco", m_ColorFondoDescripcionSinFoco, m_def_ColorFondoDescripcionSinFoco)
    Call PropBag.WriteProperty("LongitudEtiqueta", m_LongitudEtiqueta, m_def_LongitudEtiqueta)
    Call PropBag.WriteProperty("LongitudTexto", m_LongitudTexto, m_def_LongitudTexto)
    Call PropBag.WriteProperty("LongitudDescripcion", m_LongitudDescripcion, m_def_LongitudDescripcion)
    Call PropBag.WriteProperty("Movible", m_Movible, m_def_Movible)
    Call PropBag.WriteProperty("Mascara", m_Mascara, m_def_Mascara)
    Call PropBag.WriteProperty("FechaConta", m_FechaConta, m_def_FechaConta)
    Call PropBag.WriteProperty("FechaIva", m_FechaIva, m_def_FechaIva)
    Call PropBag.WriteProperty("Mayusculas", m_Mayusculas, m_def_Mayusculas)
    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
    Call PropBag.WriteProperty("IdEtiqueta", m_IdEtiqueta, m_def_IdEtiqueta)
    Call PropBag.WriteProperty("IdTip", m_IdTip, m_def_IdTip)
    Call PropBag.WriteProperty("IdToolTipText", m_IdToolTipText, m_def_IdToolTipText)
    Call PropBag.WriteProperty("ValorN", m_ValorN, m_def_ValorN)
End Sub

Private Sub UserControl_Resize()
    Dim xx As Long
    Dim Alt1, Alt2, Alt3, Alt4, Tam1, Tam2
    
    UserControl.Extender.TabStop = m_TextoVisible
    
    On Error Resume Next
    
    Select Case m_TipoLetra
        Case TLPequeña
            If m_PosicionEtiqueta = TPArriba Then
                Alt1 = 240 * 2
            Else
                Alt1 = 240
            End If
            Alt2 = 240: Alt3 = 210: Alt4 = 140: Tam1 = 8: Tam2 = 8 '10
        Case TLMediana
            If m_PosicionEtiqueta = TPArriba Then
                Alt1 = 295 * 2
            Else
                Alt1 = 285
            End If
            Alt2 = 285: Alt3 = 260: Alt4 = 180: Tam1 = 10: Tam2 = 10 '12
        Case TLGrande
            If m_PosicionEtiqueta = TPArriba Then
                Alt1 = 355 * 2
            Else
                Alt1 = 345
            End If
            Alt2 = 345: Alt3 = 310: Alt4 = 220: Tam1 = 12: Tam2 = 12 '16
    End Select
    
    UserControl.Height = Alt1
    LblEtiqueta.Height = Alt2
    LblEtiquetaFnd.Height = Alt2
    Txt.Height = Alt3
    LblCheck.Height = Alt4
    CheckMark.Width = Alt4: CheckMark.Height = Alt4
    OptionMark.Width = Alt4: OptionMark.Height = Alt4
    LblPor.Height = Alt2
    LblDescripcion.Height = Alt2
    Boton.Height = Alt2
    LblEtiqueta.Font.Size() = Tam1
    Txt.Font.Size() = Tam2
    LblPor.Font.Size() = Tam2
    LblDescripcion.Font.Size() = Tam2
    Select Case m_PosicionEtiqueta
        Case TPIzquierda
            'Etiqueta
            If m_EtiquetaVisible And Not m_TextoVisible And Not m_BotonVisible And Not m_DescripcionVisible Then
                LblEtiqueta.Width = UserControl.Width
                Call PosicionarEtiqueta(0, 0)
            End If
            'Texto
            If Not m_EtiquetaVisible And m_TextoVisible And Not m_BotonVisible And Not m_DescripcionVisible Then
                Txt.Width = UserControl.Width - IIf(m_Formato = FPorcentaje, LblPor.Width, 0) - (16 * 2) - 16
                Call PosicionarTexto(0, 0)
            End If
            'Etiqueta Texto
            If m_EtiquetaVisible And m_TextoVisible And Not m_BotonVisible And Not m_DescripcionVisible Then
                Call PosicionarEtiqueta(0, 0)
                If m_Formato = FChequeo Or m_Formato = FOpcion Then
                    Select Case m_TipoLetra
                        Case TLPequeña: LblCheck.Width = 140
                        Case TLMediana: LblCheck.Width = 180
                        Case TLGrande: LblCheck.Width = 220
                    End Select
                    LblEtiqueta.Width = UserControl.Width - LblCheck.Width - (16 * 3) + 8
                    Call PosicionarTexto(LblEtiqueta.Width + 32, 0)
                ElseIf m_Formato = FFecha Then
                    LblEtiqueta.Width = UserControl.Width - 1100
                    Txt.Width = 1100
                    Call PosicionarTexto(LblEtiqueta.Width + 16, 0)
                Else
                    Txt.Width = UserControl.Width - (LblEtiqueta.Width + 16) - IIf(m_Formato = FPorcentaje, LblPor.Width, 0) - (16 * 2) - 16
                    Call PosicionarTexto(LblEtiqueta.Width + 16, 0)
                End If
            End If
            'Etiqueta Texto Botón
            If m_EtiquetaVisible And m_TextoVisible And m_BotonVisible And Not m_DescripcionVisible Then
                Txt.Width = UserControl.Width - (LblEtiqueta.Width + 16) - IIf(m_Formato = FPorcentaje, LblPor.Width, 0) - (16 * 2) - (Boton.Width + 16) - 16
                Call PosicionarEtiqueta(0, 0)
                Call PosicionarTexto(LblEtiqueta.Width + 16, 0)
                xx = Txt.Left + Txt.Width + IIf(m_Formato = FPorcentaje, LblPor.Width, 0) + (16 * 2)
                Boton.Left = xx: Boton.Top = 0
            End If
            'Etiqueta Texto Descripción
            If m_EtiquetaVisible And m_TextoVisible And Not m_BotonVisible And m_DescripcionVisible Then
                LblDescripcion.Width = UserControl.Width - (LblEtiqueta.Width + 16) - IIf(m_Formato = FPorcentaje, LblPor.Width, 0) - (Txt.Width) - (16 * 3)
                Call PosicionarEtiqueta(0, 0)
                Call PosicionarTexto(LblEtiqueta.Width + 16, 0)
                xx = Txt.Left + Txt.Width + IIf(m_Formato = FPorcentaje, LblPor.Width, 0) + (16 * 2)
                Call PosicionarDescripcion(xx, 0)
            End If
            'Etiqueta Texto Botón Descripción
            If m_EtiquetaVisible And m_TextoVisible And m_BotonVisible And m_DescripcionVisible Then
                LblDescripcion.Width = UserControl.Width - (LblEtiqueta.Width + 16) - IIf(m_Formato = FPorcentaje, LblPor.Width, 0) - (Txt.Width) - (Boton.Width + 16) - (16 * 3) - 16
                Call PosicionarEtiqueta(0, 0)
                Call PosicionarTexto(LblEtiqueta.Width + 16, 0)
                xx = Txt.Left + Txt.Width + IIf(m_Formato = FPorcentaje, LblPor.Width, 0) + (16 * 2)
                Boton.Left = xx: Boton.Top = 0
                Call PosicionarDescripcion(Boton.Left + Boton.Width + 16, 0)
                'Call PosicionarDescripcion(xx, 0)
            End If
            'Etiqueta Descripcion
            If m_EtiquetaVisible And Not m_TextoVisible And Not m_BotonVisible And m_DescripcionVisible Then
                LblDescripcion.Width = UserControl.Width - (LblEtiqueta.Width + 16)
                Call PosicionarEtiqueta(0, 0)
                Call PosicionarDescripcion(LblEtiqueta.Width + 16, 0)
            End If
        Case TPDerecha
        Case TPArriba
            'Etiqueta Texto
            If m_EtiquetaVisible And m_TextoVisible And Not m_BotonVisible And Not m_DescripcionVisible Then
                LblEtiqueta.Width = UserControl.Width + 16
                Select Case m_TipoLetra
                    Case TLPequeña: LblEtiqueta.Height = 195: xx = -16
                    Case TLMediana: LblEtiqueta.Height = 195 + 60: xx = -16
                    Case TLGrande: LblEtiqueta.Height = 195 + 60 + 60: xx = -32
                End Select
                If m_AlineacionEtiqueta = TADerecha Then xx = 16
                Call PosicionarEtiqueta(xx, 0)
                Txt.Width = UserControl.Width - IIf(m_Formato = FPorcentaje, LblPor.Width, 0) - (16 * 2) - 16
                Call PosicionarTexto(0, LblEtiqueta.Height + 48)
            End If
            'Etiqueta Texto Botón
            If m_EtiquetaVisible And m_TextoVisible And m_BotonVisible And Not m_DescripcionVisible Then
                LblEtiqueta.Width = UserControl.Width + 16
                Select Case m_TipoLetra
                    Case TLPequeña: LblEtiqueta.Height = 195: xx = -16
                    Case TLMediana: LblEtiqueta.Height = 195 + 60: xx = -16
                    Case TLGrande: LblEtiqueta.Height = 195 + 60 + 60: xx = -32
                End Select
                If m_AlineacionEtiqueta = TADerecha Then xx = 16
                Call PosicionarEtiqueta(xx, 0)
                Txt.Width = UserControl.Width - IIf(m_Formato = FPorcentaje, LblPor.Width, 0) - (16 * 2) - (Boton.Width + 16) - 16
                Call PosicionarTexto(0, LblEtiqueta.Height + 48)
                xx = Txt.Left + Txt.Width + IIf(m_Formato = FPorcentaje, LblPor.Width, 0) + (16 * 2)
                Boton.Left = xx: Boton.Top = LblEtiqueta.Height + 48
            End If
            'Etiqueta Descripcion
            If m_EtiquetaVisible And Not m_TextoVisible And Not m_BotonVisible And m_DescripcionVisible Then
                LblEtiqueta.Width = UserControl.Width
                LblEtiqueta.Height = LblEtiqueta.Height - 16
                LblDescripcion.Width = UserControl.Width
                Call PosicionarEtiqueta(0, 0)
                Call PosicionarDescripcion(0, LblEtiqueta.Height + 16)
            End If
    End Select
    On Error GoTo 0
End Sub
'TipoLetra:
Public Property Get TipoLetra() As TiposLetra
    TipoLetra = m_TipoLetra
End Property

Public Property Let TipoLetra(ByVal New_TipoLetra As TiposLetra)
    m_TipoLetra = New_TipoLetra
    PropertyChanged "TipoLetra"
    Call RefrescarControl
End Property
'Formato:
Public Property Get Formato() As Formatos
    Formato = m_Formato
End Property

Public Property Let Formato(ByVal New_Formato As Formatos)
    m_Formato = New_Formato
    PropertyChanged "Formato"
    If m_Formato = FCantidad Or m_Formato = FMoneda Or m_Formato = FPorcentaje Or m_Formato = FNumerico Or m_Formato = FTotales Then
        If m_TextoVisible Then
            Txt.Alignment() = 1
        Else
            LblDescripcion.Alignment() = 1
        End If
    Else
        If m_TextoVisible Then
            Txt.Alignment() = 0
        Else
            LblDescripcion.Alignment() = 0
        End If
    End If
    
    Select Case m_Formato
        Case FCadena: longitud = 50
        Case FCodigo: longitud = 4
        Case FFecha: longitud = 10
        Case FHora: longitud = 8
        Case FSubcuenta: longitud = 8
        Case FCantidad, FMoneda, FPorcentaje, FNumerico, FTotales: longitud = 15
        Case FChequeo, FOpcion: longitud = 1
        Case FCodCli: longitud = 4
        Case FCodObra: longitud = 11
        Case FTelefono: longitud = 9
    End Select

    If m_Formato = FChequeo Or m_Formato = FOpcion Then
        Select Case m_TipoLetra
            Case TLPequeña: LblCheck.Width = 140
            Case TLMediana: LblCheck.Width = 180
            Case TLGrande: LblCheck.Width = 220
        End Select
    End If
    
    Call RefrescarControl
End Property
'Longitud:
Public Property Get longitud() As String
Attribute longitud.VB_ProcData.VB_Invoke_Property = "General"
    longitud = m_Longitud
End Property

Public Property Let longitud(ByVal New_Longitud As String)
    m_Longitud = New_Longitud
    Txt.MaxLength() = m_Longitud
    PropertyChanged "Longitud"
End Property
'Texto:
Public Property Get Texto() As Variant 'el valor del control
'--> Obtiene el valor actual del control
'--> si el formato es tipo chequeo u opcion, devuelve true o false
'--> El resto de formato, una cadena de texto
    If m_Formato = FChequeo Then
        Texto = CheckMark.Visible
    ElseIf m_Formato = FOpcion Then
        Texto = OptionMark.Visible
    Else
        If m_TextoVisible Then
            Texto = Trim(Txt)
        Else
            Texto = Trim(LblDescripcion)
        End If
    End If
End Property

Public Property Let Texto(ByVal New_Texto As Variant) 'Asigna un nuevo valor
    Select Case m_Formato
        Case FCadena, FSubcuenta, FTelefono
            m_Texto = Trim("" & New_Texto)
        Case FChequeo
            CheckMark.Visible = New_Texto

        Case FOpcion
            OptionMark.Visible = New_Texto
        Case FCodigo, FCodCli
            m_Texto = Format(New_Texto, String(Txt.MaxLength, "0"))
        Case FFecha
            '05/05/13 modificada la propiedad Public Property Let en el campo
            'If IsNull(New_Texto) Then
            If IsNull(New_Texto) Or New_Texto = "" Then
                m_Texto = ""
            Else
                m_Texto = Format(CDate(New_Texto), "DD/MM/YYYY")
            End If
        Case FHora
            If IsNull(New_Texto) Then
                m_Texto = ""
            Else
                m_Texto = Format(CDate(New_Texto), "Hh:Nn:Ss")
            End If
        Case FCantidad, FMoneda, FPorcentaje, FNumerico, FTotales
            If IsNumeric(New_Texto) Then
                If Mascara <> "" Then
                    m_Texto = Format(New_Texto, Mascara)
                Else
                    If Fix(CDbl(New_Texto)) - CDbl(New_Texto) <> 0 Then
                        m_Texto = Format(New_Texto, "#,##0.########")
                    Else
                        m_Texto = Format(New_Texto, "#,##0")
                    End If
                End If
            Else
                m_Texto = ""
            End If
        Case FCodObra
            If Len(New_Texto) = 9 Then
                m_Texto = Trim(str(Mid(New_Texto, 1, 3) & "." & Mid(New_Texto, 4, 3) & "." & Mid(New_Texto, 7)))
            Else
                m_Texto = ""
            End If
            
    End Select
    If m_TextoVisible Then
        Txt.Text() = m_Texto
    Else
        LblDescripcion() = m_Texto
    End If
    PropertyChanged "Texto"
    AlgunCambio = False
End Property
'Limpiar:
Public Sub Limpiar() 'Limpia el control - Borra los datos
    Txt = "": m_Texto = ""
    LblDescripcion = "": m_Descripcion = ""
    CheckMark.Visible = False
    OptionMark.Visible = False
    AlgunCambio = False
End Sub
'Valor:
Public Property Let valor(ByVal New_VALOR As Variant) 'Asigna valor
'--> Adigna un nuevo valor, si el formato es chequeo u opcion, se puede asignar valor true/false
    If m_Formato = FChequeo Then CheckMark.Visible = New_VALOR
    If m_Formato = FOpcion Then OptionMark.Visible = New_VALOR
End Property

Public Property Get valor() As Variant 'Obtiene el valor
'--> Obtiene el valor del campo, devuelve NULL si el campo esta en blanco
    Select Case m_Formato
        Case FChequeo: valor = CheckMark.Visible
        Case FOpcion: valor = OptionMark.Visible
        Case FCodObra
            If Trim(Txt.Text) = "" Then valor = Null Else valor = Mid(Txt, 1, 3) & Mid(Txt, 5, 3) & Mid(Txt, 9)
        Case Else
            If m_TextoVisible Then
                If Trim(Txt.Text) = "" Then valor = Null Else valor = Trim(Txt.Text)
            Else
                If Trim(LblDescripcion) = "" Then valor = Null Else valor = Trim(LblDescripcion)
            End If
    End Select
End Property

Public Property Get ValorZ() As Double 'Obtiene el valor
'--> Devuelve un tipo de datos Double, si el campo no es numérico, devuelve 0
    Select Case m_Formato
        Case FCadena, FCodigo, FFecha, FHora, FSubcuenta, FChequeo, FOpcion, FCodCli, FCodObra, FTelefono: ValorZ = 0#
        Case Else
            If m_TextoVisible Then
                If IsNumeric(Txt.Text) Then
                    ValorZ = CDbl(Txt.Text)
                Else
                    ValorZ = 0#
                End If
            Else
                If IsNumeric(LblDescripcion) Then
                    ValorZ = CDbl(LblDescripcion)
                Else
                    ValorZ = 0#
                End If
            End If
    End Select
End Property
'AlgunCambio:
Public Property Get AlgunCambio() As Boolean 'Si ha variado el valor del campo
'--> Aunque se puede asignar por codigo, el control devolvera true si se ha introducido un nuevo valor
    AlgunCambio = m_AlgunCambio
End Property

Public Property Let AlgunCambio(ByVal New_AlgunCambio As Boolean)
    
    If UCase(CampoBD) = gstrCadenaAComprobar Then
        MiSql = MiSql
    End If
    
    m_AlgunCambio = New_AlgunCambio
    PropertyChanged "AlgunCambio"
End Property
'Valido:
Public Property Get Valido() As Boolean
    Valido = (Txt.ForeColor <> QBColor(12))
End Property
'Etiqueta:
Public Property Get etiqueta() As String
Attribute etiqueta.VB_ProcData.VB_Invoke_Property = "General"
    etiqueta = LblEtiqueta
End Property

Public Property Let etiqueta(ByVal New_Etiqueta As String)
    m_Etiqueta = "" & New_Etiqueta
    LblEtiqueta() = m_Etiqueta
    PropertyChanged "Etiqueta"
End Property
'AlineacionEtiqueta:
Public Property Get AlineacionEtiqueta() As TiposAlineacion
    AlineacionEtiqueta = m_AlineacionEtiqueta
End Property

Public Property Let AlineacionEtiqueta(ByVal New_AlineacionEtiqueta As TiposAlineacion)
    m_AlineacionEtiqueta = New_AlineacionEtiqueta
    LblEtiqueta.Alignment() = m_AlineacionEtiqueta
    PropertyChanged "AlineacionEtiqueta"
    Call UserControl_Resize
End Property
'PosicionEtiqueta:
Public Property Get PosicionEtiqueta() As TiposPosicion
    PosicionEtiqueta = m_PosicionEtiqueta
End Property

Public Property Let PosicionEtiqueta(ByVal New_PosicionEtiqueta As TiposPosicion)
    If m_Formato = FChequeo Or m_Formato = FOpcion And New_PosicionEtiqueta = TPArriba Then Exit Property
    If (m_DescripcionVisible Or m_BotonVisible) And New_PosicionEtiqueta = TPDerecha Then Exit Property
    If m_TextoVisible And m_DescripcionVisible And New_PosicionEtiqueta = TPArriba Then Exit Property
    If m_PosicionEtiqueta = TPArriba And New_PosicionEtiqueta <> TPArriba Then
        LblEtiqueta.Width() = LblEtiqueta.Width / 2
    End If
    m_PosicionEtiqueta = New_PosicionEtiqueta
    PropertyChanged "PosicionEtiqueta"
    Call UserControl_Resize
End Property
'Tip:
Public Property Get Tip() As String
Attribute Tip.VB_ProcData.VB_Invoke_Property = "General"
    Tip = m_Tip
End Property

Public Property Let Tip(ByVal New_Tip As String)
    m_Tip = "" & New_Tip
    PropertyChanged "Tip"
End Property
'TablaBD:
Public Property Get TablaBD() As String
    TablaBD = m_TablaBD
End Property

Public Property Let TablaBD(ByVal New_TablaBD As String)
    m_TablaBD = Trim(New_TablaBD)
    PropertyChanged "TablaBD"
End Property
'CampoBD:
Public Property Get CampoBD() As String 'El campo de la BD
Attribute CampoBD.VB_ProcData.VB_Invoke_Property = "General"
    CampoBD = m_CampoBD
End Property

Public Property Let CampoBD(ByVal New_CampoBD As String) 'El campo de la BD
    m_CampoBD = Trim(New_CampoBD)
    PropertyChanged "CampoBD"
End Property
'CampoBDDescripcion:
Public Property Get CampoBDDescripcion() As String 'El campo de la BD
Attribute CampoBDDescripcion.VB_ProcData.VB_Invoke_Property = "General"
    CampoBDDescripcion = m_CampoBDDescripcion
End Property

Public Property Let CampoBDDescripcion(ByVal New_CampoBDDescripcion As String)
    m_CampoBDDescripcion = Trim(New_CampoBDDescripcion)
    PropertyChanged "CampoBDDescripcion"
End Property
'Descripcion:
Public Property Get Descripcion() As Variant 'el valor de la descripcion
Attribute Descripcion.VB_ProcData.VB_Invoke_Property = "General"
    Descripcion = m_Descripcion
End Property

Public Property Let Descripcion(ByVal New_Descripcion As Variant)
    m_Descripcion = "" & New_Descripcion
    LblDescripcion() = m_Descripcion
    PropertyChanged "Descripcion"
End Property
'EtiquetaVisible:
Public Property Get EtiquetaVisible() As Boolean
Attribute EtiquetaVisible.VB_ProcData.VB_Invoke_Property = "General"
    EtiquetaVisible = m_EtiquetaVisible
End Property

Public Property Let EtiquetaVisible(ByVal New_EtiquetaVisible As Boolean)
    m_EtiquetaVisible = New_EtiquetaVisible
    LblEtiqueta.Visible() = (m_Visible And m_EtiquetaVisible)
    PropertyChanged "EtiquetaVisible"
    Call RefrescarControl
End Property
'TextoVisible:
Public Property Get TextoVisible() As Boolean
Attribute TextoVisible.VB_ProcData.VB_Invoke_Property = "General"
    TextoVisible = m_TextoVisible
End Property

Public Property Let TextoVisible(ByVal New_TextoVisible As Boolean)
    m_TextoVisible = New_TextoVisible
    Txt.Visible() = (m_Visible And m_TextoVisible)
    LblCheck.Visible() = (m_Visible And m_TextoVisible And (m_Formato = FChequeo Or m_Formato = FOpcion))
    If (m_Visible And m_TextoVisible And m_Formato <> FChequeo And m_Formato <> FOpcion) Then
        Txt.Width() = m_LongitudTexto
        UserControl.Width = UserControl.Width + m_LongitudTexto
    End If
    PropertyChanged "TextoVisible"
    Call RefrescarControl
End Property
'BotonVisible:
Public Property Get BotonVisible() As Boolean
Attribute BotonVisible.VB_ProcData.VB_Invoke_Property = "General"
    BotonVisible = m_BotonVisible
End Property

Public Property Let BotonVisible(ByVal New_BotonVisible As Boolean)
    If m_PosicionEtiqueta = TPDerecha Then Exit Property
    If m_Formato = FChequeo Or m_Formato = FOpcion Then Exit Property
    m_BotonVisible = New_BotonVisible
    Boton.Visible() = (m_Visible And m_BotonVisible)
    PropertyChanged "BotonVisible"
    Call RefrescarControl
End Property
'DescripcionVisible:
Public Property Get DescripcionVisible() As Boolean
Attribute DescripcionVisible.VB_ProcData.VB_Invoke_Property = "General"
    DescripcionVisible = m_DescripcionVisible
End Property

Public Property Let DescripcionVisible(ByVal New_DescripcionVisible As Boolean)
    If m_PosicionEtiqueta = TPDerecha Or m_PosicionEtiqueta = TPArriba Then Exit Property
    If m_Formato = FChequeo Or m_Formato = FOpcion Then Exit Property
    m_DescripcionVisible = New_DescripcionVisible
    LblDescripcion.Visible() = (m_Visible And m_DescripcionVisible)
    If (m_Visible And m_DescripcionVisible) Then
        LblDescripcion.Width() = m_LongitudDescripcion
        UserControl.Width = UserControl.Width + m_LongitudDescripcion
    End If
    PropertyChanged "DescripcionVisible"
    Call RefrescarControl
End Property
'Activado:
Public Property Get Activado() As Boolean 'Si el campo esta activado o no ENABLED/DISABLED
Attribute Activado.VB_ProcData.VB_Invoke_Property = "General"
    Activado = m_Activado
End Property

Public Property Let Activado(ByVal New_Activado As Boolean)
    m_Activado = New_Activado
    LblEtiqueta.Enabled() = m_Activado
    '161019 Campo activado
    Txt.Enabled() = m_Activado
    
    LblPor.Enabled() = m_Activado
    Boton.Enabled() = m_Activado
    LblDescripcion.Enabled() = m_Activado
    '161019 Campo activado
    UserControl.Enabled() = m_Activado
    
    PropertyChanged "Activado"
End Property
'Visible:
Public Property Get Visible() As Boolean
Attribute Visible.VB_ProcData.VB_Invoke_Property = "General"
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal New_Visible As Boolean)
    m_Visible = New_Visible
    LblEtiqueta.Visible() = (m_Visible And m_EtiquetaVisible)
    Txt.Visible() = (m_Visible And m_TextoVisible)
    LblCheck.Visible() = (m_Visible And m_TextoVisible And (m_Formato = FChequeo Or m_Formato = FOpcion))
    LblPor.Visible() = (m_Visible And m_TextoVisible)
    Boton.Visible() = (m_Visible And m_BotonVisible)
    LblDescripcion.Visible() = (m_Visible And m_DescripcionVisible)
    PropertyChanged "Visible"
End Property
'EtiquetaNegrita:
Public Property Get EtiquetaNegrita() As Boolean
Attribute EtiquetaNegrita.VB_ProcData.VB_Invoke_Property = "General"
    EtiquetaNegrita = m_EtiquetaNegrita
End Property

Public Property Let EtiquetaNegrita(ByVal New_EtiquetaNegrita As Boolean)
    m_EtiquetaNegrita = New_EtiquetaNegrita
    LblEtiqueta.Font.Bold() = m_EtiquetaNegrita
    PropertyChanged "EtiquetaNegrita"
End Property
'ColorLetraEtiquetaConFoco:
Public Property Get ColorLetraEtiquetaConFoco() As OLE_COLOR
    ColorLetraEtiquetaConFoco = m_ColorLetraEtiquetaConFoco
End Property

Public Property Let ColorLetraEtiquetaConFoco(ByVal New_ColorLetraEtiquetaConFoco As OLE_COLOR)
    m_ColorLetraEtiquetaConFoco = New_ColorLetraEtiquetaConFoco
    PropertyChanged "ColorLetraEtiquetaConFoco"
End Property
'ColorFondoEtiquetaConFoco:
Public Property Get ColorFondoEtiquetaConFoco() As OLE_COLOR
    ColorFondoEtiquetaConFoco = m_ColorFondoEtiquetaConFoco
End Property

Public Property Let ColorFondoEtiquetaConFoco(ByVal New_ColorFondoEtiquetaConFoco As OLE_COLOR)
    m_ColorFondoEtiquetaConFoco = New_ColorFondoEtiquetaConFoco
    PropertyChanged "ColorFondoEtiquetaConFoco"
End Property
'ColorLetraEtiquetaSinFoco:
Public Property Get ColorLetraEtiquetaSinFoco() As OLE_COLOR
    ColorLetraEtiquetaSinFoco = m_ColorLetraEtiquetaSinFoco
End Property

Public Property Let ColorLetraEtiquetaSinFoco(ByVal New_ColorLetraEtiquetaSinFoco As OLE_COLOR)
    m_ColorLetraEtiquetaSinFoco = New_ColorLetraEtiquetaSinFoco
    LblEtiqueta.ForeColor() = m_ColorLetraEtiquetaSinFoco
    PropertyChanged "ColorLetraEtiquetaSinFoco"
End Property
'ColorFondoEtiquetaSinFoco:
Public Property Get ColorFondoEtiquetaSinFoco() As OLE_COLOR
    ColorFondoEtiquetaSinFoco = m_ColorFondoEtiquetaSinFoco
End Property

Public Property Let ColorFondoEtiquetaSinFoco(ByVal New_ColorFondoEtiquetaSinFoco As OLE_COLOR)
    m_ColorFondoEtiquetaSinFoco = New_ColorFondoEtiquetaSinFoco
    LblEtiqueta.BackColor() = m_ColorFondoEtiquetaSinFoco
    LblEtiquetaFnd.BackColor() = m_ColorFondoEtiquetaSinFoco
    PropertyChanged "ColorFondoEtiquetaSinFoco"
End Property
'ColorLetraTextoConFoco:
Public Property Get ColorLetraTextoConFoco() As OLE_COLOR
    ColorLetraTextoConFoco = m_ColorLetraTextoConFoco
End Property

Public Property Let ColorLetraTextoConFoco(ByVal New_ColorLetraTextoConFoco As OLE_COLOR)
    m_ColorLetraTextoConFoco = New_ColorLetraTextoConFoco
    PropertyChanged "ColorLetraTextoConFoco"
End Property
'ColorFondoTextoConFoco:
Public Property Get ColorFondoTextoConFoco() As OLE_COLOR
    ColorFondoTextoConFoco = m_ColorFondoTextoConFoco
End Property

Public Property Let ColorFondoTextoConFoco(ByVal New_ColorFondoTextoConFoco As OLE_COLOR)
    m_ColorFondoTextoConFoco = New_ColorFondoTextoConFoco
    PropertyChanged "ColorFondoTextoConFoco"
End Property
'ColorLetraTextoSinFoco:
Public Property Get ColorLetraTextoSinFoco() As OLE_COLOR
    ColorLetraTextoSinFoco = m_ColorLetraTextoSinFoco
End Property

Public Property Let ColorLetraTextoSinFoco(ByVal New_ColorLetraTextoSinFoco As OLE_COLOR)
    m_ColorLetraTextoSinFoco = New_ColorLetraTextoSinFoco
    Txt.ForeColor() = m_ColorLetraTextoSinFoco
    LblPor.ForeColor() = m_ColorLetraTextoSinFoco
    PropertyChanged "ColorLetraTextoSinFoco"
End Property
'ColorFondoTextoSinFoco:
Public Property Get ColorFondoTextoSinFoco() As OLE_COLOR
    ColorFondoTextoSinFoco = m_ColorFondoTextoSinFoco
End Property

Public Property Let ColorFondoTextoSinFoco(ByVal New_ColorFondoTextoSinFoco As OLE_COLOR)
    m_ColorFondoTextoSinFoco = New_ColorFondoTextoSinFoco
    Txt.BackColor() = m_ColorFondoTextoSinFoco
    LblPor.BackColor() = m_ColorFondoTextoSinFoco
    PropertyChanged "ColorFondoTextoSinFoco"
End Property
'ColorLetraDescripcionConFoco:
Public Property Get ColorLetraDescripcionConFoco() As OLE_COLOR
    ColorLetraDescripcionConFoco = m_ColorLetraDescripcionConFoco
End Property

Public Property Let ColorLetraDescripcionConFoco(ByVal New_ColorLetraDescripcionConFoco As OLE_COLOR)
    m_ColorLetraDescripcionConFoco = New_ColorLetraDescripcionConFoco
    PropertyChanged "ColorLetraDescripcionConFoco"
End Property
'ColorFondoDescripcionConFoco:
Public Property Get ColorFondoDescripcionConFoco() As OLE_COLOR
    ColorFondoDescripcionConFoco = m_ColorFondoDescripcionConFoco
End Property

Public Property Let ColorFondoDescripcionConFoco(ByVal New_ColorFondoDescripcionConFoco As OLE_COLOR)
    m_ColorFondoDescripcionConFoco = New_ColorFondoDescripcionConFoco
    PropertyChanged "ColorFondoDescripcionConFoco"
End Property
'ColorLetraDescripcionSinFoco:
Public Property Get ColorLetraDescripcionSinFoco() As OLE_COLOR
    ColorLetraDescripcionSinFoco = m_ColorLetraDescripcionSinFoco
End Property

Public Property Let ColorLetraDescripcionSinFoco(ByVal New_ColorLetraDescripcionSinFoco As OLE_COLOR)
    m_ColorLetraDescripcionSinFoco = New_ColorLetraDescripcionSinFoco
    LblDescripcion.ForeColor() = m_ColorLetraDescripcionSinFoco
    PropertyChanged "ColorLetraDescripcionSinFoco"
End Property
'ColorFondoDescripcionSinFoco:
Public Property Get ColorFondoDescripcionSinFoco() As OLE_COLOR
    ColorFondoDescripcionSinFoco = m_ColorFondoDescripcionSinFoco
End Property

Public Property Let ColorFondoDescripcionSinFoco(ByVal New_ColorFondoDescripcionSinFoco As OLE_COLOR)
    m_ColorFondoDescripcionSinFoco = New_ColorFondoDescripcionSinFoco
    LblDescripcion.BackColor() = m_ColorFondoDescripcionSinFoco
    PropertyChanged "ColorFondoDescripcionSinFoco"
End Property
'Longitud Etiqueta:
Public Property Get LongitudEtiqueta() As Long
    LongitudEtiqueta = LblEtiqueta.Width
End Property

Public Property Let LongitudEtiqueta(ByVal New_Longitud As Long)
    m_LongitudEtiqueta = New_Longitud
    LblEtiqueta.Width() = m_LongitudEtiqueta
    PropertyChanged "LongitudEtiqueta"
    Call UserControl_Resize
End Property
'Longitud Texto:
Public Property Get LongitudTexto() As Long
    LongitudTexto = Txt.Width
End Property

Public Property Let LongitudTexto(ByVal New_Longitud As Long)
    m_LongitudTexto = New_Longitud
    Txt.Width() = m_LongitudTexto
    PropertyChanged "LongitudTexto"
    Call UserControl_Resize
End Property
'Longitud Descripcion:
Public Property Get LongitudDescripcion() As Long
    LongitudDescripcion = LblDescripcion.Width
End Property

Public Property Let LongitudDescripcion(ByVal New_Longitud As Long)
    m_LongitudDescripcion = New_Longitud
    LblDescripcion.Width() = m_LongitudDescripcion
    PropertyChanged "LongitudDescripcion"
    Call UserControl_Resize
End Property
'Movible:
Public Property Get Movible() As Boolean 'Si el usuario puede mover el campo
Attribute Movible.VB_ProcData.VB_Invoke_Property = "General"
    Movible = m_Movible
End Property

Public Property Let Movible(ByVal New_Movible As Boolean)
    m_Movible = New_Movible
    PropertyChanged "Movible"
End Property
'Mascara:
Public Property Get Mascara() As String 'Mascara para los formatos numericos
Attribute Mascara.VB_ProcData.VB_Invoke_Property = "General"
    Mascara = m_Mascara
End Property

Public Property Let Mascara(ByVal New_Mascara As String)
    m_Mascara = New_Mascara
    PropertyChanged "Mascara"
End Property
'Activar:
Public Property Get Activar() As Boolean
    Activar = m_Activar
End Property

Public Property Let Activar(ByVal New_Activar As Boolean)
    m_Activar = New_Activar
    PropertyChanged "Activar"
End Property
'Ocultar:
Public Property Get Ocultar() As Boolean
    Ocultar = m_Ocultar
End Property

Public Property Let Ocultar(ByVal New_Ocultar As Boolean)
    m_Ocultar = New_Ocultar
    PropertyChanged "Ocultar"
End Property
'Obligar:
Public Property Get Obligar() As Boolean 'Si es obligado poner un valor en el campo
    Obligar = m_Obligar
End Property

Public Property Let Obligar(ByVal New_Obligar As Boolean)
    m_Obligar = New_Obligar
    PropertyChanged "Obligar"
End Property
Public Property Get Accion() As String
    Accion = Acccion
End Property

Public Property Let Accion(ByVal New_Accion As String)
    Acccion = New_Accion
    PropertyChanged "Accion"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,false
Public Property Get FechaConta() As Boolean 'si el tipo de datos es fecha, controla la fecha minima para contabilizar
    FechaConta = m_FechaConta
End Property

Public Property Let FechaConta(ByVal New_FechaConta As Boolean)
    m_FechaConta = New_FechaConta
    PropertyChanged "FechaConta"
End Property
Public Property Get FechaIva() As Boolean 'si el tipo de datos es fecha, controla la fecha minima para contabilizar apuntes de iva
    FechaIva = m_FechaIva
End Property
Public Property Let FechaIva(ByVal New_FechaIva As Boolean)
    m_FechaIva = New_FechaIva
    PropertyChanged "FechaIva"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get Mayusculas() As FormatosCadenas
    Mayusculas = m_Mayusculas
End Property

Public Property Let Mayusculas(ByVal New_Mayusculas As FormatosCadenas)
    m_Mayusculas = New_Mayusculas
    TFC = m_Mayusculas
    PropertyChanged "Mayusculas"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Devuelve un controlador (de Microsoft Windows) a la ventana de un objeto."
    hwnd = m_hWnd
End Property

Public Property Let hwnd(ByVal New_hWnd As Long)
    m_hWnd = New_hWnd
    PropertyChanged "hWnd"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get IdEtiqueta() As String
    IdEtiqueta = m_IdEtiqueta
End Property

Public Property Let IdEtiqueta(ByVal New_IdEtiqueta As String)
    m_IdEtiqueta = New_IdEtiqueta
    PropertyChanged "IdEtiqueta"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get IdTip() As String
    IdTip = m_IdTip
End Property

Public Property Let IdTip(ByVal New_IdTip As String)
    m_IdTip = New_IdTip
    PropertyChanged "IdTip"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get IdToolTipText() As String
    IdToolTipText = m_IdToolTipText
End Property

Public Property Let IdToolTipText(ByVal New_IdToolTipText As String)
    m_IdToolTipText = New_IdToolTipText
    PropertyChanged "IdToolTipText"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get ValorN() As Variant
        If Accion = "consultando" Then
            ValorN = Txt
            Exit Sub
        End If

    Select Case m_Formato
        Case FChequeo: ValorN = CheckMark.Visible
        Case FOpcion: ValorN = OptionMark.Visible
        Case FCodObra
            If Trim(Txt.Text) = "" Then ValorN = Null Else ValorN = Mid(Txt, 1, 3) & Mid(Txt, 5, 3) & Mid(Txt, 9)
        Case FCadena, FCodigo, FFecha, FHora, FSubcuenta, FChequeo, FOpcion, FCodCli, FCodObra, FTelefono
            If m_TextoVisible Then
                If Trim(Txt.Text) = "" Then ValorN = Null Else ValorN = Trim(Txt.Text)
            Else
                If Trim(LblDescripcion) = "" Then ValorN = Null Else ValorN = Trim(LblDescripcion)
            End If
        Case Else
            If m_TextoVisible Then
                If IsNumeric(Txt.Text) Then
                    ValorN = CDbl(Txt.Text)
                Else
                    ValorN = 0#
                End If
            Else
                If IsNumeric(LblDescripcion) Then
                    ValorN = CDbl(LblDescripcion)
                Else
                    ValorN = 0#
                End If
            End If
    End Select
End Property

Public Property Let ValorN(ByVal New_ValorN As Variant)
    Select Case m_Formato
        Case FCadena, FSubcuenta, FTelefono
            m_Texto = Trim("" & New_ValorN)
        Case FChequeo
            CheckMark.Visible = New_ValorN
        Case FOpcion
            OptionMark.Visible = New_ValorN
        Case FCodigo, FCodCli
            m_Texto = Format(New_ValorN, String(Txt.MaxLength, "0"))
        Case FFecha
            If IsNull(New_ValorN) Then
                m_Texto = ""
            Else
                m_Texto = Format(CDate(New_ValorN), "DD/MM/YYYY")
            End If
        Case FHora
            If IsNull(New_ValorN) Then
                m_Texto = ""
            Else
                m_Texto = Format(CDate(New_ValorN), "Hh:Nn:Ss")
            End If
        Case FCantidad, FMoneda, FPorcentaje, FNumerico, FTotales
            If IsNumeric(New_ValorN) Then
                If Mascara <> "" Then
                    m_Texto = Format(New_ValorN, Mascara)
                Else
                    If Fix(CDbl(New_ValorN)) - CDbl(New_ValorN) <> 0 Then
                        m_Texto = Format(New_ValorN, "#,##0.########")
                    Else
                        m_Texto = Format(New_ValorN, "#,##0")
                    End If
                End If
            Else
                m_Texto = ""
            End If
        Case FCodObra
            If Len(New_ValorN) = 9 Then
                m_Texto = Trim(str(Mid(New_ValorN, 1, 3) & "." & Mid(New_ValorN, 4, 3) & "." & Mid(New_ValorN, 7)))
            Else
                m_Texto = ""
            End If
            
    End Select
    If m_TextoVisible Then
        Txt.Text() = m_Texto
    Else
        LblDescripcion() = m_Texto
    End If
    PropertyChanged "Texto"
    AlgunCambio = False
    
End Property
Public Sub QueSigno() 'Obtiene el separador decimal si es punto o coma
    Dim A As Double
    A = 1.1
    SignoDecimal = Mid(CStr(A), 2, 1)
End Sub
Public Property Get FechaSQL() As Variant
'--> Obtiene el valor en formato YYYY/MM/DD
    FechaSQL = Null
    Select Case m_Formato
        Case FFecha
            If IsDate(Txt.Text) Then
                FechaSQL = "#" & Format(CDate(Txt.Text), "YYYY/MM/DD") & "#"
            End If
    End Select
End Property

