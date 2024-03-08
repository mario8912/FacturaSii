VERSION 5.00
Begin VB.Form frmVentasNuevas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes Nuevos Articulos Venta"
   ClientHeight    =   5805
   ClientLeft      =   16965
   ClientTop       =   4350
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   4860
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Top             =   4440
         Width           =   2535
      End
      Begin Listados.Campo CampoRutaD 
         Height          =   480
         Left            =   480
         TabIndex        =   5
         Top             =   1279
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Formato         =   8
         Longitud        =   15
         Etiqueta        =   "Ruta:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
      End
      Begin Listados.Campo CampoFechaD 
         Height          =   480
         Left            =   480
         TabIndex        =   3
         Top             =   767
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Formato         =   2
         Longitud        =   10
         Etiqueta        =   "Fecha:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
      End
      Begin Listados.Campo CampoFechaH 
         Height          =   480
         Left            =   2880
         TabIndex        =   4
         Top             =   767
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Formato         =   2
         Longitud        =   10
         Etiqueta        =   "Fecha:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
      End
      Begin Listados.Campo CampoRutaH 
         Height          =   480
         Left            =   2880
         TabIndex        =   6
         Top             =   1279
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Formato         =   8
         Longitud        =   15
         Etiqueta        =   "Ruta:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
      End
      Begin Listados.Campo CampoPreventistaD 
         Height          =   480
         Left            =   480
         TabIndex        =   7
         Top             =   1791
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Formato         =   8
         Longitud        =   15
         Etiqueta        =   "Preventista:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
      End
      Begin Listados.Campo CampoPreventistaH 
         Height          =   480
         Left            =   2880
         TabIndex        =   8
         Top             =   1791
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Formato         =   8
         Longitud        =   15
         Etiqueta        =   "Preventista:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
      End
      Begin Listados.Campo CampoFamiliaD 
         Height          =   480
         Left            =   480
         TabIndex        =   9
         Top             =   2303
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Etiqueta        =   "Familia:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
         Mascara         =   "<"
      End
      Begin Listados.Campo CampoFamiliaH 
         Height          =   480
         Left            =   2880
         TabIndex        =   10
         Top             =   2303
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Etiqueta        =   "Familia:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
         Mascara         =   "<"
      End
      Begin Listados.Campo CampoSubFamiliaD 
         Height          =   480
         Left            =   480
         TabIndex        =   11
         Top             =   2815
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Etiqueta        =   "Sub Familia:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
         Mascara         =   "<"
      End
      Begin Listados.Campo CampoSubFamiliaH 
         Height          =   480
         Left            =   2880
         TabIndex        =   12
         Top             =   2815
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Etiqueta        =   "Sub Familia:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
         Mascara         =   "<"
      End
      Begin Listados.Campo CampoArticuloD 
         Height          =   480
         Left            =   480
         TabIndex        =   13
         Top             =   3327
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Etiqueta        =   "Articulo:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
         Mascara         =   "<"
      End
      Begin Listados.Campo CampoArticuloH 
         Height          =   480
         Left            =   2880
         TabIndex        =   14
         Top             =   3327
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   847
         Etiqueta        =   "Articulo:"
         AlineacionEtiqueta=   2
         PosicionEtiqueta=   2
         Mascara         =   "<"
      End
      Begin Listados.Boton Boton2 
         Height          =   495
         Left            =   2880
         TabIndex        =   16
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Etiqueta        =   "&Cancelar"
         UsarIcono       =   -1  'True
         Icono           =   "frmVentasNuevas.frx":0000
      End
      Begin Listados.Boton Boton1 
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Etiqueta        =   "&Aceptar"
         UsarIcono       =   -1  'True
         Icono           =   "frmVentasNuevas.frx":06FA
      End
      Begin VB.Label Label3 
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
         Left            =   480
         TabIndex        =   18
         Top             =   4500
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmVentasNuevas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPrincipal As ADODB.Recordset
Private v_DtmFechaD As Date
Private v_DtmFechaH As Date
Private v_RutaD As Integer
Private v_RutaH As Integer

Private v_PreventistaD As Integer
Private v_PreventistaH As Integer
Private v_FamiliaD As String
Private v_FamiliaH As String
Private v_SubFamiliaD As String
Private v_SubFamiliaH As String
Private v_ArticuloD As String
Private v_ArticuloH As String

Private Function Consulta() As Boolean
    On Error GoTo ControlError
    Dim strProcedimiento As String
    Dim ocommand As ADODB.Command
    strProcedimiento = "dbo.ListadoCliNuevosXArticulo"
    'strProcedimiento = "dbo.CliNuevosXArticulo.rpt"
    
    
    Set ocommand = New ADODB.Command
    Set ocommand.ActiveConnection = DB
    ocommand.CommandText = strProcedimiento
    ocommand.CommandType = adCmdStoredProc
    
    Set rsPrincipal = New ADODB.Recordset
    'ocommand.Parameters("@FechaIni").Value = v_DtmFechaD
    'ocommand.Parameters("@FechaFin").Value = v_DtmFechaH
    
    ocommand.Parameters("@FechaIni").Value = v_DtmFechaD
    ocommand.Parameters("@FechaFin").Value = v_DtmFechaH
    ocommand.Parameters("@RutaIni").Value = v_RutaD
    ocommand.Parameters("@RutaFin").Value = v_RutaH
    ocommand.Parameters("@PreventaIni").Value = v_PreventistaD
    ocommand.Parameters("@PreventaFin").Value = v_PreventistaH
    ocommand.Parameters("@FamiliaIni").Value = v_FamiliaD
    ocommand.Parameters("@FamiliaFin").Value = v_FamiliaH
    ocommand.Parameters("@SubfamiliaIni").Value = v_SubFamiliaD
    ocommand.Parameters("@SubfamiliaFin").Value = v_SubFamiliaH
    ocommand.Parameters("@ArticuloIni").Value = v_ArticuloD
    ocommand.Parameters("@ArticuloFin").Value = v_ArticuloH
    
    
    DB.CommandTimeout = 120
    'rsPrincipal.Open oCadena3, DB, adOpenStatic, adLockReadOnly
    Set rsPrincipal = ocommand.Execute
    DB.CommandTimeout = 30
    Consulta = True
    Exit Function
ControlError:
    Msg mError, "", Err, Erl
'''       strProcedimiento = "dbo.ActualizarInventarioX"
'''       Set ocommand = New ADODB.Command
'''       Set ocommand.ActiveConnection = BDrs
'''       ocommand.CommandText = strProcedimiento
'''       ocommand.CommandType = adCmdStoredProc
'''
'''          'ocommand.Parameters("@DesdeFecha").Value = dtmFechaIni
'''       ocommand.Parameters("@HastaFecha").Value = dtmFechaFin
'''       ocommand.Execute
'''       retorno = ocommand.Parameters(0).Value
    
End Function


Private Sub Boton1_BotonClick(Boton As Integer)
    v_DtmFechaD = CampoFechaD.valor
    v_DtmFechaH = CampoFechaH.valor
    v_RutaD = CampoRutaD.ValorZ
    v_RutaH = CampoRutaH.ValorZ
    v_PreventistaD = CampoPreventistaD.ValorZ
    v_PreventistaH = CampoPreventistaH.ValorZ
    v_FamiliaD = CampoFamiliaD.valor
    v_FamiliaH = CampoFamiliaH.valor
    v_SubFamiliaD = CampoSubFamiliaD.valor
    v_SubFamiliaH = CampoSubFamiliaH.valor
    v_ArticuloD = CampoArticuloD.valor
    v_ArticuloH = CampoArticuloH.valor
    
    Screen.MousePointer = vbHourglass
    Dim strRuta As String
    Dim pos As Long
   
    Dim Objeto As Form
    Dim qimpresora As String
    qimpresora = Combo1.Text
    If Consulta Then
   
        Set Objeto = New frmVisorListados
        Load Objeto
        strRuta = App.Path
        pos = InStrRev(strRuta, "\")
        strRuta = Mid(strRuta, 1, pos)
        strRuta = strRuta & "INFORMES\Ventas\PrimerasCompras.rpt"
        Objeto.AbreInforme strRuta
        Objeto.AsignaRs rsPrincipal, 0, ""
        
        'Objeto.AsignaValorParam 0, v_DtmFechaD
        'Objeto.AsignaValorParam 2, v_DtmFechaH
        Objeto.AsignarFormula "fechaini", v_DtmFechaD
        Objeto.AsignarFormula "fechafin", v_DtmFechaH
        Objeto.AsignarFormula "rutaini", v_RutaD
        Objeto.AsignarFormula "rutafin", v_RutaH
        Objeto.AsignarFormula "preventaini", v_PreventistaD
        Objeto.AsignarFormula "preventafin", v_PreventistaH
        
        Objeto.AsignarFormula "familiaini", v_FamiliaD
        Objeto.AsignarFormula "familiafin", v_FamiliaH
        Objeto.AsignarFormula "subfamiliaini", v_SubFamiliaD
        Objeto.AsignarFormula "subfamiliafin", v_SubFamiliaH
        Objeto.AsignarFormula "articuloini", v_ArticuloD
        Objeto.AsignarFormula "articulofin", v_ArticuloH
        Objeto.AsignaImpresora qimpresora
        'Objeto.ImpresoraElegida = True
        Objeto.MuestraInforme
        
        
        Objeto.Show
        
        Set Objeto = Nothing
        
        Set rsPrincipal = Nothing
        Unload Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Boton2_BotonClick(Boton As Integer)
Unload Me
End Sub

Private Sub Form_Load()
'Me.Move 0, 0
CentrarFormulario Me
'''ALTER                 PROCEDURE     dbo.ListadoCliNuevosXArticulo
'''@FechaIni DATETIME =NULL,
'''@FechaFin DATETIME =NULL,
'''@RutaIni SMALLINT = NULL,
'''@RutaFin SMALLINT = NULL,
'''@PreventaIni SMALLINT = NULL,
'''@PreventaFin SMALLINT = NULL,
'''@FamiliaIni VARCHAR(3)='',
'''@FamiliaFin VARCHAR(3)='',
'''@SubfamiliaIni VARCHAR(3)='',
'''@SubfamiliaFin VARCHAR(3)='',
'''@ArticuloIni VARCHAR(4)='',
'''@ArticuloIni VARCHAR(4)=''
    dtmFechaIni = CDate("01/01/" & Year(Date))
    CampoFechaD.Texto = dtmFechaIni
    dtmFechaFin = CDate(Now)
    CampoFechaH.Texto = dtmFechaFin
    
    CampoRutaD.Texto = 1
    CampoRutaH.Texto = 32767
    
    CampoPreventistaD.Texto = 1
    CampoPreventistaH.Texto = 32767
    
    CampoFamiliaD.Texto = "A"
    CampoFamiliaH.Texto = "ZZZ"
    
    CampoSubFamiliaD.Texto = "A"
    CampoSubFamiliaH.Texto = "ZZZ"
    
    CampoArticuloD.Texto = "0"
    CampoArticuloH.Texto = "ZZZZ"
    '32,767
    
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
