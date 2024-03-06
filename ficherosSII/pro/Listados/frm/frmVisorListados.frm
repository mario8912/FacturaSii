VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmVisorListados 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   10320
   ClientTop       =   2850
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11940
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmVisorListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As New CrystalReport1

Option Explicit
Public ImpresoraElegida As Boolean

Dim Apli As CRAXDRT.Application
Dim Informe As CRAXDRT.Report
Dim SubInforme As CRAXDRT.SubreportObject
Dim CRXSubReport As CRAXDRT.Report

Dim blnHorizontal As Boolean

'220706 orientacion del report
Dim intHorizontal As enumOrientacion

Public Function orientacion() As Integer
    orientacion = Informe.PaperOrientation
End Function

Public Property Get qOrientacion() As enumOrientacion
    intHorizontal = Informe.PaperOrientation
    qOrientacion = intHorizontal
End Property

Public Property Let qOrientacion(ByVal New_qOrientacion As enumOrientacion)
    intHorizontal = New_qOrientacion
    Informe.PaperOrientation = intHorizontal
End Property


Public Sub AsignaRs(rs As ADODB.Recordset, Tipo As Integer, Nombre As String)
'************************************************************************
'                 ADO
'        ASIGNA EL RECORDSET CORRESPONDIENTE AL INFORME PRINCIPAL O
'        A LOS SUBINFORMES
'************************************************************************
Dim Seccion As CRAXDRT.Section
Dim Objeto As Object
   
   If Tipo = 0 Then
      'Es el recordset principal
      Informe.Database.SetDataSource rs, 3, 1
   Else
      'es un subinforme
      For Each Seccion In Informe.Sections
         For Each Objeto In Seccion.ReportObjects
            If Objeto.Kind = crSubreportObject Then
               'Select Case LCase(Objeto.Name)
               Select Case LCase(Objeto.SubreportName)
                  Case LCase(Nombre)
                     Set SubInforme = Objeto
                     Set CRXSubReport = SubInforme.OpenSubreport
                     CRXSubReport.Database.SetDataSource rs, 3, 1
               End Select
            End If
         Next
      Next
   End If
   
   Set Seccion = Nothing
   Set Objeto = Nothing
End Sub
Public Sub AbreInforme(vInforme As String)
    On Error GoTo Fallo
    Dim cad As String
    Dim pos As Integer
    Screen.MousePointer = vbHourglass
    'Set Apli = CreateObject("CrystalRuntime.Application")
    Set Apli = New CRAXDRT.Application
    DoEvents
    Set Informe = Apli.OpenReport(vInforme)
    DoEvents
    Me.Caption = Informe.ReportTitle
    DoEvents
    Screen.MousePointer = vbDefault

    
    pos = InStrRev(vInforme, ".")
    cad = Right(vInforme, 5)
    
    blnHorizontal = False
    If Mid(cad, 1, 1) = "_" Then
        'Informe.PaperOrientation = crLandscape
        blnHorizontal = True
    'Else
        'Informe.PaperOrientation = crDefaultPaperOrientation
    End If
    
    On Error GoTo 0
    Exit Sub
Fallo:
    If Err.Number = -2147206458 Then
        MsgBox "Fichero no encontrado"
        bolContinuar = False
        Screen.MousePointer = 0
        Exit Sub
    End If
    MsgBox Err.Number & " " & Err.Description

End Sub

Public Sub AsignaBD()
'************************************************************************
'                 OLE DB
'        ASIGNA LA BD CORRESPONDIENTE A LAS TABLAS DEL INFORME PRINCIPAL O
'        A LOS SUBINFORMES
'************************************************************************
Screen.MousePointer = vbHourglass
   Dim Tabla As CRAXDRT.DatabaseTable
   Dim Seccion As CRAXDRT.Section
   Dim Objeto As Object
   
   For Each Tabla In Informe.Database.Tables
      'Tabla.SetLogOnInfo "POLLO", "ROSELL"
      'Tabla.SetLogOnInfo "ROSELL", "ROSELL"
      Tabla.SetLogOnInfo NombreServidor, NombreBD
      DoEvents
   Next
   'si hay subinformes
   For Each Seccion In Informe.Sections
      For Each Objeto In Seccion.ReportObjects
         If Objeto.Kind = crSubreportObject Then
            Set SubInforme = Objeto
            DoEvents
            Set CRXSubReport = SubInforme.OpenSubreport
            DoEvents
            For Each Tabla In CRXSubReport.Database.Tables
               Tabla.SetLogOnInfo NombreServidor, NombreBD
               DoEvents
               'Tabla.SetLogOnInfo "ROSELL", "ROSELL"
            Next
         End If
      Next
   Next
   
   Set Tabla = Nothing
   Set Seccion = Nothing
   Set Objeto = Nothing

Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaImpresora(impresora As String)
    Informe.SelectPrinter ObtenerDriverImpresora(impresora), impresora, ObtenerPuertoImpresora(impresora)
    ImpresoraElegida = True
End Sub

Public Sub MuestraInforme()
   'Muestra el setup de la impresora
    On Error GoTo Fallo
'   If gNombreImpresora <> "" Then
'    Informe.SelectPrinter ObtenerDriverImpresora(gNombreImpresora), gNombreImpresora, ObtenerPuertoImpresora(gNombreImpresora)
'   Else
'    Informe.PrinterSetup Me.hWnd
'   End If
   
   If Not ImpresoraElegida Then
   Informe.PrinterSetup Me.hwnd
   End If
   
   
   CRViewer1.ReportSource = Informe
   DoEvents
   CRViewer1.ViewReport
   DoEvents
'    If blnHorizontal Then
'        Informe.PaperOrientation = crLandscape
'    Else
'        Informe.PaperOrientation = crPortrait
'    End If
    
    On Error GoTo 0
    Exit Sub
Fallo:
    MsgBox Err.Number & " " & Err.Description
   
End Sub

Public Sub AsignarFormula(Nombreformula As String, vValor As Variant)
   Dim CrxFFD As CRAXDRT.FormulaFieldDefinitions
   Dim CrxFormula As CRAXDRT.FormulaFieldDefinition
   Dim I As Integer
   
   Set CrxFFD = Informe.FormulaFields
   For I = 1 To CrxFFD.Count
        Set CrxFormula = CrxFFD.Item(I)
        If CrxFormula.FormulaFieldName = Nombreformula Then
        CrxFormula.Text = "'" & vValor & "'"
        End If
   Next
End Sub
Public Sub AsignaValorParam(vLugar As Integer, vValor As Variant)
   Informe.ParameterFields(vLugar).SetCurrentValue vValor
End Sub
Public Sub OcultaSeccion(nSec As Variant)
   'las secciones pueden representarse por numero o por su nombre
   '  RH-->Report Header, PH-->Page Header, GHn-->Group Header n, D-->Details
   '  GFn-->Group Foot n, PF-->Page Foot, RF-->Report Foot
   ' Si hay multiples secciones, se representan con una letra minuscula p.e. "Da", "Db"...
   
   Informe.Sections(nSec).Suppress = True
End Sub



Private Sub CRViewer1_OnReportSourceError(ByVal errorMsg As String, ByVal errorCode As Long, UseDefault As Boolean)
    If errorCode = -2147206395 Then
        '221020 blnCancelado
        blnCancelado = True
        Unload Me
    End If
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
'220706 Cancelar el boton dialogo de impresora cuando se pincha en imprimir
   
   'Cancela el dialog box, pero tambien la impresion
'   UseDefault = False
   'Muestra el setup de la impresora
   'Informe.PrinterSetup Me.hwnd
   
   'imprime el informe ya que lo he cancelado con el usedefault
'   Informe.PrintOut False
End Sub




Private Sub Form_Load()
Me.Move 50, 1000, Screen.Width - 1500, Screen.Height - 3000
'Screen.MousePointer = vbHourglass
'
'CRViewer1.ReportSource = Report
'CRViewer1.ViewReport
'Screen.MousePointer = vbDefault

'220706 orientacion del report
intHorizontal = vertical
End Sub
Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
Private Sub Form_Unload(Cancel As Integer)
    ImpresoraElegida = False
    Set Informe = Nothing
   Set Apli = Nothing
   Set SubInforme = Nothing
   Set CRXSubReport = Nothing
End Sub
Function ObtenerDriverImpresora(impresora As String) As String
   Dim pr As Printer
   For Each pr In Printers
      If pr.DeviceName = impresora Then
         ObtenerDriverImpresora = pr.DriverName
         Exit Function
      End If
   Next
   ObtenerDriverImpresora = ""
End Function

Function ObtenerPuertoImpresora(impresora As String) As String
   Dim pr As Printer
   For Each pr In Printers
      If pr.DeviceName = impresora Then
         ObtenerPuertoImpresora = pr.Port
         Exit Function
      End If
   Next
   ObtenerPuertoImpresora = ""
End Function
