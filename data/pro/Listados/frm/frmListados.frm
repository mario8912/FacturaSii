VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListados 
   Caption         =   "Listados"
   ClientHeight    =   4980
   ClientLeft      =   5850
   ClientTop       =   3390
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   5730
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4920
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   8678
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsListados As ADODB.Recordset
Public Function AbreRs() As Boolean
On Error GoTo ControlError
Dim fs
Dim existeRs As Boolean

   Set fs = CreateObject("Scripting.FileSystemObject")
   existeRs = fs.FileExists(App.Path & "\rslistados")
   
   Set RsListados = New ADODB.Recordset
   
   If existeRs Then
      RsListados.Open App.Path & "\rsListados", , adOpenStatic, adLockBatchOptimistic
      RsListados.Sort = "IdOrden"
      AbreRs = True
   Else
      MsgBox "NO EXISTE LA INFORMACION DE LOS LISTADOS"
      AbreRs = False
   End If
   
Exit Function
ControlError:
   AbreRs = False
End Function




Private Sub LlenaArbol()
   
   ' Establece propiedades del control ImageList.
   TreeView1.LineStyle = tvwRootLines  ' Linestyle = 1

   ' Agrega los objetos Node.
   Dim nodX As Node    ' Declara una variable Node.
   
   RsListados.MoveFirst
   ' Primer nodo con el texto 'Raíz'.
   Set nodX = TreeView1.Nodes.Add(, , RsListados.Fields("IdListado").Value, RsListados.Fields("Titulo").Value)
   RsListados.MoveNext
   
   While Not RsListados.EOF
      If NivelUsuario > CByte(RsListados.Fields("NAcceso").Value) Then
         Set nodX = TreeView1.Nodes.Add(RsListados.Fields("IdPadre").Value, tvwChild, RsListados.Fields("IdListado").Value, RsListados.Fields("Titulo").Value)
      End If
      RsListados.MoveNext
   Wend
End Sub

Private Sub Form_Load()
   If AbreRs Then
      LlenaArbol
   End If

End Sub

Private Sub Form_Resize()
   On Error Resume Next
   
   TreeView1.Width = Me.ScaleWidth - (2 * TreeView1.Left)
   TreeView1.Height = Me.ScaleHeight - TreeView1.Top - TreeView1.Left
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RsListados = Nothing
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
   'Busco el registro del nodo seleccionado
   Dim strRuta As String
   Dim pos As Long
   Dim intOrientacion As Integer
   
   RsListados.MoveFirst
   RsListados.Find "IdListado='" & Node.Key & "'", 0, adSearchForward, 0
   If LTrim(RsListados.Fields("camino").Value) <> "" Then
      If Not RsListados.EOF Then
         'Obtengo los valores del registro
      'v PARA VER EL INFORME ANTES DE IMPRIMIR
            
            FrmImpresoras.Show 1
            
            If gNombreImpresora = "" Then Exit Sub
            
            
            Dim Objeto As Form
            Set Objeto = New frmVisorListados
            Load Objeto
            
 
            strRuta = App.Path
            pos = InStrRev(strRuta, "\")
            strRuta = Mid(strRuta, 1, pos)
            strRuta = strRuta & Mid(RsListados.Fields("camino").Value, 4)
            
            'strRuta = App.Path & Mid(RsListados.Fields("camino").Value, 3)

            Objeto.AbreInforme strRuta
            'Objeto.AbreInforme RsListados.Fields("camino").Value
            Objeto.AsignaBD
            If RsListados.Fields("moneda").Value Then
               Objeto.AsignaValorParam 1, CDbl(ValorMoneda)
            End If
            'Objeto.AsignaValorParam 1, oCarga.IdCarga
            
            'intOrientacion = Objeto.orientacion
            intOrientacion = Objeto.qOrientacion
            
            
            Objeto.AsignaImpresora gNombreImpresora
            Objeto.qOrientacion = intOrientacion
            'Objeto.orientacion = 2
            
            '221020 blnCancelado
            blnCancelado = False
            '221020 me.mousepointer en frmListados
            Me.MousePointer = vbHourglass
            Objeto.MuestraInforme
            Objeto.ZOrder
            
            '221020 blnCancelado
            If Not blnCancelado Then
            Objeto.Show
            End If
            '221020 me.mousepointer en frmListados
            Me.MousePointer = 0
            
      
            'Unload Objeto
            'Set Objeto = Nothing
      '^
      End If
   Else
      If LCase(RsListados.Fields("IdListado").Value) = "previs" Then
         frmPrevision.Show
      End If
      'vnuevos
      If LCase(RsListados.Fields("IdListado").Value) = "vnuevos" Then
         frmVentasNuevas.Show
      End If
   End If
End Sub
