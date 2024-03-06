VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAgregarListado 
   Caption         =   "Agregar LISTADOS"
   ClientHeight    =   5895
   ClientLeft      =   13620
   ClientTop       =   3210
   ClientWidth     =   7785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   7785
   Begin VB.PictureBox picButtons 
      Align           =   1  'Align Top
      Height          =   850
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7725
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   7785
      Begin VB.CommandButton Command2 
         Caption         =   "importar"
         Height          =   800
         Left            =   2340
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton Command1 
         Caption         =   "exportar"
         Height          =   800
         Left            =   1560
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdAnterior 
         Caption         =   "&Anterior"
         Height          =   800
         Left            =   780
         Picture         =   "frmAgregarListado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Edita la cuadrícula"
         Top             =   0
         Width           =   780
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Salir"
         Height          =   800
         Left            =   4905
         Picture         =   "frmAgregarListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cierra la ventana"
         Top             =   0
         Width           =   780
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Modificar"
         Height          =   800
         Left            =   0
         Picture         =   "frmAgregarListado.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Edita la cuadrícula"
         Top             =   0
         Width           =   780
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   800
         Left            =   780
         Picture         =   "frmAgregarListado.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela las modificaciones realizadas"
         Top             =   0
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Grabar"
         Height          =   800
         Left            =   0
         Picture         =   "frmAgregarListado.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Graba las modificaciones realizadas"
         Top             =   0
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin MSDataGridLib.DataGrid grdListados 
      Height          =   4245
      Left            =   90
      TabIndex        =   0
      Top             =   990
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   7488
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAgregarListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim WithEvents RsListados As ADODB.Recordset
Attribute RsListados.VB_VarHelpID = -1

Private Const COLIdOrden As Integer = 0
Private Const COLIdListado As Integer = 1
Private Const COLTitulo As Integer = 2
Private Const COLIdPadre As Integer = 3
Private Const COLCamino As Integer = 4
Private Const COLMoneda As Integer = 5
Private Const COLNAcceso As Integer = 6
Private AltDown As Boolean


Public Sub AbreRs()
Dim fs
Dim existeRs As Boolean

   Set fs = CreateObject("Scripting.FileSystemObject")
   existeRs = fs.FileExists(App.Path & "\rslistados")
   
   Set RsListados = New ADODB.Recordset
   
   If existeRs Then
      RsListados.Open App.Path & "\rsListados", , adOpenStatic, adLockBatchOptimistic
      RsListados.Sort = "IdOrden"
   Else
      RsListados.Fields.Append "IdOrden", adSmallInt
      RsListados.Fields.Append "IdListado", adVarChar, 6
      RsListados.Fields.Append "Titulo", adVarChar, 50
      RsListados.Fields.Append "IdPadre", adVarChar, 6
      RsListados.Fields.Append "Camino", adVarChar, 255
      RsListados.Fields.Append "Moneda", adBoolean
      RsListados.Fields.Append "NAcceso", adTinyInt
      
      RsListados.Open
   End If
   
   Set grdListados.DataSource = RsListados
   PropiedadesGrid
   
   Set fs = Nothing
   
End Sub



Private Sub PasoAEditar(bValor As Boolean)
   grdListados.AllowAddNew = bValor
   grdListados.AllowDelete = bValor
   grdListados.AllowUpdate = bValor
   
   cmdEdit.Visible = Not bValor
   cmdClose.Visible = Not bValor
   
   cmdUpdate.Visible = bValor
   cmdCancel.Visible = bValor
End Sub
Private Sub PropiedadesGrid()
   'Mostrar las barras de desplazamiento para el control.
   grdListados.ScrollBars = dbgBoth

   'Oculta las columnas no utilizadas en el control.

   'Encabezados de columna para el control DataGrid
   grdListados.Columns(COLIdOrden).Caption = "ORDEN"
   grdListados.Columns(COLIdListado).Caption = "IdListado"
   grdListados.Columns(COLTitulo).Caption = "Título"
   grdListados.Columns(COLIdPadre).Caption = "Id Padre"
   grdListados.Columns(COLCamino).Caption = "Camino"
   grdListados.Columns(COLMoneda).Caption = "Moneda"
   grdListados.Columns(COLNAcceso).Caption = "Nivel Usuario"
   
   grdListados.Columns(COLIdOrden).Width = 800
   grdListados.Columns(COLIdListado).Width = 800
   grdListados.Columns(COLTitulo).Width = 3000
   grdListados.Columns(COLIdPadre).Width = 800
   grdListados.Columns(COLCamino).Width = 3000
   grdListados.Columns(COLMoneda).Width = 800
   grdListados.Columns(COLNAcceso).Width = 800
End Sub

Private Sub RecuperaAnterior()
Dim fs
Dim existeRs As Boolean
   Dim rsListadosAnt
   Set fs = CreateObject("Scripting.FileSystemObject")
   existeRs = fs.FileExists(App.Path & "\rslistadosAnt")
   
   Set rsListadosAnt = New ADODB.Recordset
   
   If existeRs Then
      rsListadosAnt.Open App.Path & "\rslistadosAnt", , adOpenStatic, adLockBatchOptimistic
      rsListadosAnt.Sort = "IdOrden"
            
      Set RsListados = New ADODB.Recordset
      RsListados.Fields.Append "IdOrden", adSmallInt
      RsListados.Fields.Append "IdListado", adVarChar, 6
      RsListados.Fields.Append "Titulo", adVarChar, 50
      RsListados.Fields.Append "IdPadre", adVarChar, 6
      RsListados.Fields.Append "Camino", adVarChar, 255
      RsListados.Fields.Append "Moneda", adBoolean
      RsListados.Fields.Append "NAcceso", adTinyInt
      
      RsListados.Open
      
      rsListadosAnt.MoveFirst
      While Not rsListadosAnt.EOF
         RsListados.AddNew
         RsListados.Fields("IdOrden") = rsListadosAnt.Fields("IdOrden")
         RsListados.Fields("IdListado") = rsListadosAnt.Fields("IdListado")
         RsListados.Fields("Titulo") = rsListadosAnt.Fields("Titulo")
         RsListados.Fields("IdPadre") = rsListadosAnt.Fields("IdPadre")
         RsListados.Fields("Camino") = rsListadosAnt.Fields("Camino")
         RsListados.Fields("Moneda") = rsListadosAnt.Fields("Moneda")
         RsListados.Fields("NAcceso") = rsListadosAnt.Fields("NAcceso")
         rsListadosAnt.MoveNext
      Wend
      Set grdListados.DataSource = RsListados
      PropiedadesGrid

   End If
   
   Set fs = Nothing

End Sub

Private Sub cmdAnterior_Click()
   RecuperaAnterior
End Sub


Private Sub cmdCancel_Click()
   PasoAEditar False
   RsListados.CancelBatch adAffectAll
   If RsListados.RecordCount = 0 Then
      AbreRs
   End If
End Sub
Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdEdit_Click()
   PasoAEditar True
End Sub
Private Sub cmdUpdate_Click()
   PasoAEditar False
   RsListados.Save App.Path & "\RsListados"
End Sub

Private Sub Form_Load()
   AbreRs
End Sub

Private Sub Form_Resize()
On Error Resume Next
   grdListados.Width = Me.ScaleWidth - 180
   grdListados.Height = Me.ScaleHeight - grdListados.Top
   
   cmdClose.Left = Me.ScaleWidth - cmdClose.Width - 75
      
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RsListados = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AltDown = (Shift And vbAltMask) > 0
    If AltDown Then
        If KeyCode = vbKeyF12 Then
            Command1.Visible = Not Command1.Visible
            Command2.Visible = Not Command2.Visible
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    AltDown = (Shift And vbAltMask) > 0
End Sub


Private Sub Command1_Click()
    CrearExcelDesatendido RsListados, App.Path & "\rsexcellistados"
End Sub

Private Sub Command2_Click()
'Set rs = Importar_Excel(strFicheroExcel, Cadt)
    Dim strFicheroExcel As String
    Dim strFicheroRSact As String
    Dim strFicheroRSGuardar As String
    
    strFicheroExcel = App.Path & "\rsexcellistados.xls"
    Set RsListados = Nothing
    Set RsListados = Importar_Excel(strFicheroExcel)
    
    'guardar la copa anterior
    strFicheroRSact = App.Path & "\RsListados"
    Dim i As Integer
    Do
        strFicheroRSGuardar = strFicheroRSact & Format(i, "00")
        If Not ArchivoExistente(strFicheroRSGuardar) Then Exit Do
        i = i + 1
    Loop Until i = 100
    Name strFicheroRSact As strFicheroRSGuardar
    grdListados.ClearFields
    Set grdListados.DataSource = RsListados
    grdListados.Refresh
    RsListados.Save App.Path & "\RsListados"
    PropiedadesGrid
    Unload Me
End Sub


