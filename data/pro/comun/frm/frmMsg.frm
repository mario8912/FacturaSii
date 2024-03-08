VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   11550
   ClientTop       =   3030
   ClientWidth     =   6600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMsg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMsg.frx":0442
   ScaleHeight     =   1950
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   2580
      Picture         =   "frmMsg.frx":2A3BA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1380
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   3360
      Picture         =   "frmMsg.frx":2B9DC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1380
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSi 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   1800
      Picture         =   "frmMsg.frx":2D000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1380
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   0
      Left            =   5400
      Picture         =   "frmMsg.frx":2E624
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   1
      Left            =   4200
      Picture         =   "frmMsg.frx":2EEEE
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   3
      Left            =   4800
      Picture         =   "frmMsg.frx":2F330
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Mensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   180
      TabIndex        =   3
      Top             =   900
      Width           =   6300
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008000&
      X1              =   60
      X2              =   6540
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      X1              =   60
      X2              =   6540
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   2
      Left            =   180
      Picture         =   "frmMsg.frx":2FBFA
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PREGUNTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   780
      TabIndex        =   2
      Top             =   240
      Width           =   3000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      Height          =   1815
      Left            =   60
      Top             =   60
      Width           =   6495
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CONTESTACION As Boolean
Public Tipo As Integer

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdNo_Click()
    CONTESTACION = False
    Unload Me
End Sub

Private Sub cmdSi_Click()
    CONTESTACION = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 83 Then KeyCode = 0: Call cmdSi_Click '83 = S
    If KeyCode = 78 Then KeyCode = 0: Call cmdNo_Click '78 = N
    If KeyCode = 65 Then KeyCode = 0: Call cmdAceptar_Click '65 = A
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    '180
    img(0).Left = 180
    img(1).Left = 180
    img(2).Left = 180
    img(3).Left = 180
    img(Tipo).Visible = True
    Select Case Tipo
        Case 0 'merror
            Line1.BorderColor = &HFF&
            Line2.BorderColor = &HFF&
            Shape1.BorderColor = &HFF&
            cmdAceptar.Visible = True
            Label1 = "ERROR"
        Case 1 'madvertencia
            Line1.BorderColor = &HFFFF&
            Line2.BorderColor = &HFFFF&
            Shape1.BorderColor = &HFFFF&
            cmdSi.Visible = True
            cmdNo.Visible = True
            Label1 = "ADVERTENCIA"
        Case 2 'mpregunta
            Line1.BorderColor = &H8000&
            Line2.BorderColor = &H8000&
            Shape1.BorderColor = &H8000&
            cmdSi.Visible = True
            cmdNo.Visible = True
            Label1 = "PREGUNTA"
        Case 3 'minformacion
            Line1.BorderColor = &HFF0000
            Line2.BorderColor = &HFF0000
            Shape1.BorderColor = &HFF0000
            cmdAceptar.Visible = True
            Label1 = "INFORMACION"
    End Select
    
End Sub
