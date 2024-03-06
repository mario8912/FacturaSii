VERSION 5.00
Begin VB.Form FrmImpresoras 
   Caption         =   "Seleccione la impresora"
   ClientHeight    =   2115
   ClientLeft      =   14610
   ClientTop       =   2550
   ClientWidth     =   4485
   Icon            =   "FrmImpresoras.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin Listados.Boton Boton2 
      Height          =   495
      Left            =   2820
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Etiqueta        =   "&Cancelar"
      UsarIcono       =   -1  'True
      Icono           =   "FrmImpresoras.frx":06EA
   End
   Begin Listados.Boton Boton1 
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Etiqueta        =   "&Aceptar"
      UsarIcono       =   -1  'True
      Icono           =   "FrmImpresoras.frx":0DE4
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4335
      Begin Listados.Campo CampoNumFacturas 
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   423
         Formato         =   8
         Longitud        =   15
         Etiqueta        =   "Numero de copias:"
         LongitudEtiqueta=   3000
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   450
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmImpresoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Boton1_BotonClick(Boton As Integer)
    gNombreImpresora = Combo1.Text
    Unload Me
End Sub

Private Sub Boton2_BotonClick(Boton As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    gNombreImpresora = ""
    Dim A As Integer, Marcador As Integer
    
    A = 0
    For Each G_Prn In Printers
        Combo1.AddItem G_Prn.DeviceName
        'if G_Prn.DeviceName = ImpresoraLis Then
        If Printer.DeviceName = G_Prn.DeviceName Then
            Marcador = A
        End If
        A = A + 1
    Next
    Combo1.ListIndex = Marcador
    'Combo1.ListIndex = 0

End Sub
