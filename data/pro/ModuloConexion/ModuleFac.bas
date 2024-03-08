Attribute VB_Name = "Module1"
Option Explicit
'230207 GBlnTipoIvaImpPlasdiferent
Global GBlnTipoIvaImpPlasdiferente As Boolean

'221227 IdIvaPlastico
Global GIdIvaPlastico As Integer
Global GDblPorcentIvaPlastico As Double
Global GDblPorcentREPlastico As Double

'221228 BlnConImpuestoPlastico
Global BlnConImpuestoPlastico As Boolean

'si se crean instancias del form, cuando se nombran desde otros form se van
'creando mas instancias
'Public FormFact As frmFacturacion

'160808 blnRecorriendoGrid
Public blnRecorriendoGrid As Boolean

Public DB As ADODB.Connection

Public ValorMoneda As Currency

Public DBShape As ADODB.Connection

Public Const mIndExIva As Byte = 1
Public Const mIndExSIG As Byte = 2
Public Const mIndExPorReg As Byte = 4
Public Const mIndRecEqui As Byte = 8

Public NivelUsuario As Byte
Public NombreServidor As String
Public NombreBD As String

'26/05/14 Añadidas variables gUsuario y gClave
Public gUsuario As String
Public gClave As String
Global CadenaConexion As String

'220701 pasar idusuario a facturacion
Public IdUsuario As String

Sub RecuperaMoneda()
   Dim cmd As Command
   Dim rs As Recordset
   
   Set cmd = New ADODB.Command
   cmd.ActiveConnection = DB
   cmd.CommandText = "ValorMoneda"
   cmd.CommandType = adCmdStoredProc
   
   Set rs = cmd.Execute
   Set cmd.ActiveConnection = Nothing
   While (Not rs.EOF)
      ValorMoneda = rs(0)
      rs.MoveNext
   Wend
   Set rs = Nothing
   Set cmd = Nothing

End Sub




Public Sub Conecta()
'la conexión se realiza desde el from que abre éste
  Set DB = New ADODB.Connection
  
  Set DBShape = New ADODB.Connection

  DB.CursorLocation = adUseClient
  
'**********************************************************
' OJO SI UTILIZO UNA SOLA CONEXION PARA LOS SHAPE Y EL SQLOLEDB,
' NO SE PUEDE UTILIZAR LA GENERACION AUTOMATICA DE PARAMETROS CON
'        oCommand.Parameters("@IdArticulo") = vIdArticulo
' SINO QUE TENDRIA QUE CREAR EL PARAMETRO Y AGREGARLO A LA COLECCION
'        Set oParametro = oCommand.CreateParameter("IdArticulo", adVarChar, adParamInput, 4, vIdArticulo)
'        oCommand.Parameters.Append oParametro
'**********************************************************
    '26/05/14 Cambiada la conexion a la bd (para que funcione en maquinas virtuales)
    Dim strCad As String

    'strCad = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; password = ''; Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
    
    '05/08/14 Añadida la variable gVirtual
    If gVirtual Then
        strCad = "Provider=SQLOLEDB.1;"
        strCad = strCad & "Persist Security Info=False;"
        strCad = strCad & "User ID=" & gUsuario & ";"
        strCad = strCad & "password =" & gClave & ";"
        strCad = strCad & "Initial Catalog=" & NombreBD & ";"
        strCad = strCad & "Data Source=" & NombreServidor
    Else
        strCad = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
        'strCad = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; password = ''; Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
    End If

CadenaConexion = strCad
'RGN :Conexion a POLLO con la seguridad integrada de NT
'DB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=pollo"
'DB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=ROSELL"

'26/05/14 Cambiada la conexion a la bd (para que funcione en maquinas virtuales)
'DB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
DB.Open strCad


'Para poder utilizar el SHAPE
  DBShape.CursorLocation = adUseClient
  DBShape.Provider = "MSDataShape"
  'DBShape.ConnectionString = "Data Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=pollo"
  'DBShape.ConnectionString = "Data Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=ROSELL"
  
  '26/05/14 Cambiada la conexion a la bd (para que funcione en maquinas virtuales)
  'DBShape.ConnectionString = "Data Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
  DBShape.ConnectionString = "Data " & strCad
    
'RGN_200600:abro la conexion del shape
   DBShape.Open
   
End Sub

Sub Main()
    On Error GoTo Fallo
    '26/05/14 Añadidas variables gUsuario y gClave
    gUsuario = "sa"
    gClave = ""
    
    '05/08/14 Añadida la variable gVirtual
    Dim strArchivo As String
    strArchivo = "..\principal\virtual.txt"
    
    If ArchivoExistente(strArchivo) Then
        gVirtual = True
    End If
    
    '25/07/14 Añadido Accion
    Accion = Ninguna
    
    '28/05/14 añadida funcion QueSigno (para obtener el signo decimal del sistema)
    Call QueSigno
    
    vValM = "#,##0.0000"
    vValP = "#,##0.00"
    vValC = "#,##0.00"
    vValT = "#,##0.00"
    vValE = "###,###,##0"
    
    '29/05/14 añadida la variable vValPesetas
    vValPesetas = "###,###,##0"
    
    '24/04/14 Añadido el modulo ModIde (este modulo esta en la carpeta ModuloConexion)
    'si EsIde esta en tiempo de diseño, osea, ejecutando desde el entorno de visual
    gEsIde = EsIde
'    Conecta
'    RecuperaMoneda
''    Set FormFact = New frmFacturacion
''    Load FormFact
''    FormFact.Show
'      'frmFacturacion.Show
'      NivelUsuario = Command()
'   MDIForm1.Show
   
    Dim cadCommand As String
    Dim posInicial As Integer
    Dim posFinal As Integer
    '160630 Dim strFacturaInversa As String
    Dim strFacturaInversa As String
    gblnFacturaInversa = False
    'factura;9;(local);Rosell;SI
    cadCommand = Command()
    If cadCommand = "" Then
      End
    End If
    posInicial = 1
    '9;server2017;rosell200910mia;SI
    '9;server2017;rosell200910mia;SI
    
    posFinal = InStr(cadCommand, ";")
    
    '220701 pasar idusuario a facturacion
    IdUsuario = Mid(cadCommand, 1, posFinal - posInicial)
    posInicial = posFinal + 1
    posFinal = InStr(posInicial, cadCommand, ";")
    
    NivelUsuario = Mid(cadCommand, posInicial, posFinal - posInicial)
    posInicial = posFinal + 1
    posFinal = InStr(posInicial, cadCommand, ";")
    NombreServidor = Mid(cadCommand, posInicial, posFinal - posInicial)
    posInicial = posFinal + 1
    
    
    
    
    '160630 Dim strFacturaInversa As String
    'posFinal = Len(cadCommand) + 1
    posFinal = InStr(posInicial, cadCommand, ";")
    NombreBD = Mid(cadCommand, posInicial, posFinal - posInicial)
    
    '220526 Cambiar la base de datos cuando ESIDE
    strArchivo = "..\principal\principal.ini"
    gEsIde = EsIde
    If gEsIde Then
        NombreServidor = LeerINI(strArchivo, "SERVIDOR", "Nombre")
        NombreBD = LeerINI(strArchivo, "BASE DATOS", "BASEIDE")
    End If
    '220526 Cambiar la base de datos cuando ESIDE fin
    
    '160630 Dim strFacturaInversa As String
    posInicial = posFinal + 1
    posFinal = Len(cadCommand) + 1
    strFacturaInversa = Mid(cadCommand, posInicial, posFinal - posInicial)
    If strFacturaInversa = "SI" Then gblnFacturaInversa = True
'   NivelUsuario = Command()
    
    RutaConsultas = App.path & "\..\cfg\"
    RutaCFG = RutaConsultas
    Conecta
    RecuperaMoneda
    
    '190125 FechaINI
    
    'DATEADD(yy, DATEDIFF(yy, 0, GETDATE()), 0) AS PrimerDia,
    'DATEADD(yy, DateDiff(yy, 0, GETDATE()) + 1, -1) As UltimoDia
    
'    dtmFechaIni = DateSerial(Year(Date), Month(Date) + 0, 1)
'    dtmFechaFin = DateSerial(Year(Date), Month(Date) + 1, 0)
'    dtmFechaIni = DateSerial(Year(Date), Month(Date), 1)
'    dtmFechaFin = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    If UCase(NombreBD) = "ROSELL" Or UCase(NombreBD) = "REFRESKAS" Or UCase(NombreBD) = "ROSELLTMP" Then
        dtmFechaIni = CDate("01/01/" & Year(Date))
        dtmFechaFin = CDate("31/12/" & Year(Date))
    Else
        '190806 gEsIde en Module1
        'revisar
        If gEsIde Then
            dtmFechaIni = CDate("01/01/" & Year(Date))
            dtmFechaFin = CDate("31/12/" & Year(Date))
        Else
            MiSql = Right(NombreBD, 4)
            If Not IsNumeric(MiSql) Then
                MiSql = Year(Date)
            End If
            
            dtmFechaIni = CDate("01/01/" & MiSql)
            dtmFechaFin = CDate("31/12/" & MiSql)
            
            If dtmFechaIni < ("01/01/2006") Then
'                dtmFechaIni = CDate("01/01/" & MiSql)
'                dtmFechaFin = CDate("31/12/" & MiSql)
'            Else
                dtmFechaIni = CDate("01/01/" & Year(Date))
                dtmFechaFin = CDate("31/12/" & Year(Date))
            End If
        End If
    End If
'===
    Dim rsTemp As New ADODB.Recordset
    MiSql = "Select * from config"
    MiSql = "config"
    If Not ArSqlTXT(rsTemp, MiSql) Then GoTo Fallo
    If Not rsTemp Is Nothing Then
        MiSql = MiSql
        'vValM = rsTemp!vValM
        vValMC = rsTemp!vValMC
        vValMV = rsTemp!vValMV
        vValP = rsTemp!vValP
        vValC = rsTemp!vValC
        vValT = rsTemp!vValT
        vValE = rsTemp!vValE
        GIdIvaPlastico = rsTemp!TipoIvaImpPlas
        '230207 bln TipoIvaImpPlasdiferente
        GBlnTipoIvaImpPlasdiferente = rsTemp!TipoIvaImpPlasdiferente
    End If
    CRado rsTemp
    '221227 Obtener los porcentales de iva al plastico
    MiSql = "select * from iva Where idiva = " & GIdIvaPlastico
    If Not ArSqlTXT(rsTemp, MiSql) Then GoTo Fallo
    If Not rsTemp Is Nothing Then
        GDblPorcentIvaPlastico = rsTemp!PorcentIVA
        GDblPorcentREPlastico = rsTemp!PorcentRE
    End If
    CRado rsTemp
'===
     'Set fMainForm = New MDIForm1
'Msg mInformacion, "pasado"
    MDIForm1.Show
    Exit Sub
Fallo:
    Msg mError, "", Err, Erl
   
End Sub
