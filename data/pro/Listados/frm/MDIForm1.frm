VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Aplicación de Listados"
   ClientHeight    =   6615
   ClientLeft      =   7710
   ClientTop       =   750
   ClientWidth     =   7815
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuListados 
      Caption         =   "Listados"
   End
   Begin VB.Menu mnuAgregar 
      Caption         =   "Agregar Listados"
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "Ventana"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NivelAcceso As Integer = 9

Private Sub MDIForm_Load()
    Me.AutoShowChildren = False
    Me.Caption = "Listados-" & NombreBD
   If NivelUsuario >= NivelAcceso Then
      mnuAgregar.Visible = True
   Else
      mnuAgregar.Visible = False
   End If
   
   qimpresora
End Sub
Private Sub mnuAgregar_Click()
   Screen.MousePointer = vbHourglass
   frmAgregarListado.Show
   Screen.MousePointer = vbDefault

End Sub

Private Sub mnuListados_Click()
   Screen.MousePointer = vbHourglass
   frmListados.Show
   Screen.MousePointer = vbDefault
End Sub

Function qimpresora() As String
  
    Dim buffer As String
    Dim Ret As Integer
  
    buffer = Space(255)
  
    Ret = GetProfileString("Windows", ByVal "device", "", _
                                 buffer, Len(buffer))
  
    If Ret Then
        gImpresoraDefecto = UCase(Left(buffer, _
                                   InStr(buffer, ",") - 1))
    End If
  
End Function

