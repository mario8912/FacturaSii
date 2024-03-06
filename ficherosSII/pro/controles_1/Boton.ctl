VERSION 5.00
Begin VB.UserControl Boton 
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   KeyPreview      =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   4305
   ToolboxBitmap   =   "Boton.ctx":0000
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H80000008&
      Height          =   615
      Left            =   840
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Picture1 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Etiqueta"
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
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   2520
      X2              =   4020
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   2400
      X2              =   2400
      Y1              =   360
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000015&
      X1              =   2580
      X2              =   4095
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000015&
      X1              =   4140
      X2              =   4140
      Y1              =   300
      Y2              =   660
   End
   Begin VB.Label Picture0 
      Height          =   1275
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Boton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Valores por defecto:
Const m_def_Activado = True
Const m_def_Cancelar = False
Const m_def_ColorEtiquetaConFoco = &H80000012
Const m_def_ColorEtiquetaSinFoco = &H80000012

'Propiedades:
Dim m_UsarIcono As Boolean
Dim m_Activado As Boolean
Dim m_Cancelar As Boolean
Dim m_ColorEtiquetaConFoco As OLE_COLOR
Dim m_ColorEtiquetaSinFoco As OLE_COLOR

'Dim ColorMascara As Long

Dim PulsandoTecla As Boolean
Dim PulsandoRaton As Boolean
Dim PulsandoEscape As Boolean

Dim TeclaRapida As String

Dim OCXRegistrado As Boolean

'Event Declarations:
Event BotonClick(Boton As Integer)
'Public Sub Tip(Valor As String)
'    Dim tControl As Control
'    'Dim sTip$
'    'sTip = Extender.ToolTipText
'    On Local Error Resume Next
'    For Each tControl In Controls
'        tControl.ToolTipText = Valor
'        If Err Then Err = 0
'    Next
'    On Local Error GoTo 0
'
'End Sub
Private Sub Pulsar()
    Line1.BorderColor = &H80000015
    Line2.BorderColor = &H80000015
    Line3.BorderColor = &H80000014
    Line4.BorderColor = &H80000014
    
    If m_UsarIcono Then
        If Label1 <> "" Then
            'Icono:
            Picture1.Left = ((Picture0.Width - Picture1.Width - Label1.Width) / 2) + 8 '+ 8
            Picture1.Top = ((Picture0.Height - Picture1.Height) / 2) + 16 + 8 + 8
            'Etiqueta:
            Label1.Left = Picture1.Width + ((Picture0.Width - Picture1.Width - Label1.Width) / 2) + 36
            
            Label1.Top = ((Picture0.Height - Label1.Height) / 2) + 8 + 8
        Else
            'Icono:
            Picture1.Left = ((Picture0.Width - Picture1.Width) / 2) + 16 + 8
            Picture1.Top = ((Picture0.Height - Picture1.Height) / 2) + 16 + 8
        End If
    Else
        'Etiqueta:
        Label1.Left = ((Picture0.Width - Label1.Width) / 2) + 24
        Label1.Top = ((Picture0.Height - Label1.Height) / 2) + 8 + 8
    End If
End Sub

Private Sub Soltar()
    Line1.BorderColor = &H80000014
    Line2.BorderColor = &H80000014
    Line3.BorderColor = &H80000015
    Line4.BorderColor = &H80000015
    
    If m_UsarIcono Then
        If Label1 <> "" Then
            'Icono:
            Picture1.Left = ((Picture0.Width - Picture1.Width - Label1.Width) / 2) '+ 8
            Picture1.Top = ((Picture0.Height - Picture1.Height) / 2) + 8 + 8
            'Escribir Etiqueta:
            Label1.Left = Picture1.Width + ((Picture0.Width - Picture1.Width - Label1.Width) / 2) + 30
            Label1.Top = ((Picture0.Height - Label1.Height) / 2) '- 8
        Else
            'Icono:
            Picture1.Left = ((Picture0.Width - Picture1.Width) / 2) + 8 + 8
            Picture1.Top = ((Picture0.Height - Picture1.Height) / 2) + 8 + 8
        End If
        If Not m_Activado Then
            Picture1.Enabled = False
        End If
    Else
        'Escribir Etiqueta:
        Label1.Left = ((Picture0.Width - Label1.Width) / 2) + 8
        Label1.Top = ((Picture0.Height - Label1.Height) / 2) '- 8
    End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PulsandoRaton = True: Call Pulsar
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PulsandoRaton = False: Call Soltar
    RaiseEvent BotonClick(Button)
End Sub

Private Sub Picture0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PulsandoRaton = True: Call Pulsar
End Sub

Private Sub Picture0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PulsandoRaton = False: Call Soltar
    RaiseEvent BotonClick(Button)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PulsandoRaton = True: Call Pulsar
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PulsandoRaton = False: Call Soltar
    RaiseEvent BotonClick(Button)
End Sub

Private Sub UserControl_ExitFocus()
    Label1.ForeColor = m_ColorEtiquetaSinFoco
End Sub

Private Sub UserControl_Initialize()
    Label1 = UserControl.Name
    m_UsarIcono = False
    m_Activado = m_def_Activado
    m_Cancelar = m_def_Cancelar
    m_ColorEtiquetaConFoco = m_def_ColorEtiquetaConFoco
    m_ColorEtiquetaSinFoco = m_def_ColorEtiquetaSinFoco
    PulsandoTecla = False: PulsandoRaton = False: PulsandoEscape = False
    Picture1.Left = -888
    'Picture3.Left = -888
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If PulsandoTecla Then Exit Sub
        
    If KeyCode = 13 Or KeyCode = 32 Then PulsandoTecla = True: Call Pulsar
    
    If TeclaRapida <> "" Then
        'la sig linea esto estaba comentada
        If GetAsyncKeyState(18) <= 0 And GetAsyncKeyState(Asc(TeclaRapida)) < 0 Then PulsandoTecla = True: Call Pulsar: Exit Sub
        If KeyCode = Asc(TeclaRapida) Then PulsandoTecla = True: Call Pulsar
    End If
    
    'Tecla "Izquierda":
    If KeyCode = 37 Then KeyCode = 0: MySendKeys "+{TAB}", True: Exit Sub
    
    'Tecla "Arriba":
    If KeyCode = 38 Then KeyCode = 0: MySendKeys "+{TAB}", True: Exit Sub
    
    'Tecla "Derecha":
    If KeyCode = 39 Then KeyCode = 0: MySendKeys "{TAB}", True: Exit Sub
    
    'Tecla  "Abajo":
    If KeyCode = 40 Then KeyCode = 0: MySendKeys "{TAB}", True: Exit Sub
End Sub

Private Sub UserControl_EnterFocus()
    Label1.ForeColor = m_ColorEtiquetaConFoco
    
    PulsandoTecla = False: PulsandoRaton = False: PulsandoEscape = False
    
    If TeclaRapida = "" Then Exit Sub
    If GetAsyncKeyState(18) < 0 And GetAsyncKeyState(Asc(TeclaRapida)) < 0 Then PulsandoTecla = True: Call Pulsar
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not PulsandoTecla Then Exit Sub
    
'    If TeclaRapida <> "" Then
'        If KeyCode = Asc(TeclaRapida) Then PulsandoTecla = False: Call Soltar: RaiseEvent BotonClick(1)
'    End If
    
'    If KeyCode = 13 Or KeyCode = 32 Or KeyCode = 27 Then
        PulsandoTecla = False: Call Soltar
        RaiseEvent BotonClick(1)
'    End If
    
End Sub

Private Sub UserControl_Resize()
    Picture0.Width = UserControl.Width - 28
    Picture0.Height = UserControl.Height - 28
    Picture0.Left = 8
    Picture0.Top = 8
    Line2.X1 = 0: Line2.X2 = 0: Line2.Y1 = 0: Line2.Y2 = UserControl.Height
    Line4.X1 = UserControl.Width - 20: Line4.X2 = UserControl.Width - 20: Line4.Y1 = 0: Line4.Y2 = UserControl.Height
    Line1.X1 = 0: Line1.X2 = UserControl.Width: Line1.Y1 = 0: Line1.Y2 = 0
    Line3.X1 = 0: Line3.X2 = UserControl.Width: Line3.Y1 = UserControl.Height - 20: Line3.Y2 = UserControl.Height - 20
    Shape1.Left = Picture0.Left + 16
    Shape1.Top = Picture0.Top + 16
    Shape1.Width = Picture0.Width - 32
    Shape1.Height = Picture0.Height - 32
    Call Soltar
End Sub

Public Property Get etiqueta() As String
    etiqueta = Label1.Caption
End Property

Public Property Let etiqueta(ByVal New_Etiqueta As String)
    Label1.Visible = (New_Etiqueta <> "")
    Label1.Caption() = New_Etiqueta
    PropertyChanged "Etiqueta"
    If PulsandoTecla Or PulsandoRaton Or PulsandoEscape Then
        Call Pulsar
    Else
        Call Soltar
    End If
End Property

Public Property Get UsarIcono() As Boolean
    UsarIcono = m_UsarIcono
End Property

Public Property Let UsarIcono(ByVal New_UsarIcono As Boolean)
    m_UsarIcono = New_UsarIcono
    If Not m_UsarIcono Then Picture1.Left = -888
    PropertyChanged "UsarIcono"
    Call UserControl_Resize
End Property

Public Property Get Icono() As Picture
    Set Icono = Picture1.Picture
End Property

Public Property Set Icono(ByVal New_Icono As Picture)
    Set Picture1.Picture = New_Icono
    PropertyChanged "Icono"
    Call UserControl_Resize
End Property

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim i As Integer
    
    OCXRegistrado = True '(Trim(GetRegValue(HKEY_CURRENT_USER, "Software\TraySystem", "FTSYS32")) = "WIK92-ROEI9-C2839-EQDS1-HBYVB-ZSD22")
    m_ColorEtiquetaConFoco = PropBag.ReadProperty("ColorEtiquetaConFoco", m_def_ColorEtiquetaConFoco)
    m_ColorEtiquetaSinFoco = PropBag.ReadProperty("ColorEtiquetaSinFoco", m_def_ColorEtiquetaSinFoco)
    m_UsarIcono = PropBag.ReadProperty("UsarIcono", False)
    If Not m_UsarIcono Then Picture1.Left = -888
    Label1.Caption = PropBag.ReadProperty("Etiqueta", "")
    Label1.Visible = (Label1.Caption <> "")
    Set Icono = PropBag.ReadProperty("Icono", Nothing)
    m_Activado = PropBag.ReadProperty("Activado", m_def_Activado)
    m_Cancelar = PropBag.ReadProperty("Cancelar", m_def_Cancelar)
    Label1.Enabled = m_Activado: Picture0.Enabled = m_Activado: Picture1.Enabled = m_Activado: UserControl.Enabled = m_Activado
    Label1.ForeColor() = m_ColorEtiquetaSinFoco
    If Label1.Caption <> "" Then
        For i = 1 To Len(Label1.Caption)
            If Mid(Label1.Caption, i, 1) = "&" And i < Len(Label1.Caption) Then
                TeclaRapida = Mid(Label1.Caption, i + 1, 1)
                Exit For
            End If
        Next
    End If
    Picture1.Enabled = m_Activado
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

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ColorEtiquetaConFoco", m_ColorEtiquetaConFoco, m_def_ColorEtiquetaConFoco)
    Call PropBag.WriteProperty("ColorEtiquetaSinFoco", m_ColorEtiquetaSinFoco, m_def_ColorEtiquetaSinFoco)
    Call PropBag.WriteProperty("Etiqueta", Label1.Caption, "")
    Call PropBag.WriteProperty("UsarIcono", m_UsarIcono, False)
    Call PropBag.WriteProperty("Icono", Icono, Nothing)
    Call PropBag.WriteProperty("Activado", m_Activado, m_def_Activado)
    Call PropBag.WriteProperty("Cancelar", m_Cancelar, m_def_Cancelar)
End Sub

'---------------------
'ColorEtiquetaConFoco:
'---------------------
Public Property Get ColorEtiquetaConFoco() As OLE_COLOR
    ColorEtiquetaConFoco = m_ColorEtiquetaConFoco
End Property

Public Property Let ColorEtiquetaConFoco(ByVal New_ColorEtiquetaConFoco As OLE_COLOR)
    m_ColorEtiquetaConFoco = New_ColorEtiquetaConFoco
    PropertyChanged "ColorEtiquetaConFoco"
End Property

'---------------------
'ColorEtiquetaSinFoco:
'---------------------
Public Property Get ColorEtiquetaSinFoco() As OLE_COLOR
    ColorEtiquetaSinFoco = m_ColorEtiquetaSinFoco
End Property

Public Property Let ColorEtiquetaSinFoco(ByVal New_ColorEtiquetaSinFoco As OLE_COLOR)
    m_ColorEtiquetaSinFoco = New_ColorEtiquetaSinFoco
    Label1.ForeColor() = m_ColorEtiquetaSinFoco
    PropertyChanged "ColorEtiquetaSinFoco"
End Property

'---------
'Activado:
'---------
Public Property Get Activado() As Boolean
    Activado = m_Activado
End Property

Public Property Let Activado(ByVal New_Activado As Boolean)
    Dim i As Long, j As Long
    Dim A As Boolean
    
    m_Activado = New_Activado
    Label1.Enabled = m_Activado: Picture0.Enabled = m_Activado: UserControl.Enabled = m_Activado
    Picture1.Enabled = m_Activado
    Shape1.Visible = Not m_Activado
    PropertyChanged "Activado"
End Property

'---------
'Cancelar:
'---------
Public Property Get Cancelar() As Boolean
    Cancelar = m_Cancelar
End Property

Public Property Let Cancelar(ByVal New_Cancelar As Boolean)
    m_Cancelar = New_Cancelar
    PropertyChanged "Cancelar"
End Property

