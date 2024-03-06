Attribute VB_Name = "Module1"
Public fMainForm As frmMain

Public MISql As String
Public Enum TipoMoneda
   PESETA
   EURO
End Enum

Public ValorMoneda As Currency
Public MonedaActiva As TipoMoneda
Public dB As Connection

'Lo pongo a nivel de modulo para poder obtener la moneda activa de cada usuario
Public NivelUsuario As Integer
Public IdUsuario As String
Public NombreServidor As String
Public NombreBD As String
'26/05/14 Añadidas variables gUsuario y gClave
Public gUsuario As String
Public gClave As String


Public Sub Conecta()
'la conexión se realiza desde el from que abre éste
  Set dB = New ADODB.Connection
  
  dB.CursorLocation = adUseClient
  
'RGN : Acceso al servidor POLLO con password y user
'  db.Open "Provider=SQLOLEDB.1;Password=regina;Persist Security Info=True;User ID=regina;Initial Catalog=rosell;Data Source=pollo"

'**********************************************************
' OJO SI UTILIZO UNA SOLA CONEXION PARA LOS SHAPE Y EL SQLOLEDB,
' NO SE PUEDE UTILIZAR LA GENERACION AUTOMATICA DE PARAMETROS CON
'        oCommand.Parameters("@IdArticulo") = vIdArticulo
' SINO QUE TENDRIA QUE CREAR EL PARAMETRO Y AGREGARLO A LA COLECCION
'        Set oParametro = oCommand.CreateParameter("IdArticulo", adVarChar, adParamInput, 4, vIdArticulo)
'        oCommand.Parameters.Append oParametro
'**********************************************************

'**********************************************************
' OJO 2 SI UTILIZO UNA SOLA CONEXION PARA LOS SHAPE Y EL SQLOLEDB,
' NO FUNCIONA BIEN LOS Requery Y LUEGO AddNew
'**********************************************************

''RGN :Conexion a POLLO con la seguridad integrada de NT
'db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=pollo"
'dB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=ROSELL"

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
    dB.Open strCad
'===dB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor

'RGN_290800 en lugar de utilizar dos conexiones utilizo una en la que puedo usar SHAPE
'  db.Provider = "MSDataShape"
'  db.ConnectionString = "Data Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=rosell;Data Source=pollo"
'RGN_200600:abro la conexion del shape
'   db.Open

  
End Sub
Sub Main()
    '26/05/14 Añadidas variables gUsuario y gClave
    gUsuario = "sa"
    gClave = ""
    
    '05/08/14 Añadida la variable gVirtual
    Dim strArchivo As String
    strArchivo = "..\principal\virtual.txt"
    
    If ArchivoExistente(strArchivo) Then
        gVirtual = True
    End If

'Dim cadUsuario As String
'Dim kk As Integer
'   'Tenemos que saber el idusuario para recuperar y guardar la moneda activa
''      NivelUsuario = Command()
'      cadUsuario = Command()
'      kk = InStr(cadUsuario, ";")
'      IdUsuario = Mid(cadUsuario, 1, kk - 1)
'      NivelUsuario = Mid(cadUsuario, kk + 1, Len(cadUsuario) - 1)
'
'    'frmSplash.Show
'    'frmSplash.Refresh
'    Conecta
'    RecuperaMoneda
'    Set fMainForm = New frmMain
'    Load fMainForm
'    'Unload frmSplash
'
'    fMainForm.Show
Dim cadCommand As String
Dim posInicial As Integer
Dim posFinal As Integer

    vValM = "#,##0.00"
    vValP = "#,##0.00"
    vValC = "#,##0.00"
    vValT = "#,##0.00"
    vValE = "###,###,##0"
    '28/05/14 añadida funcion QueSigno (para obtener el signo decimal del sistema)
    Call QueSigno

      cadCommand = Command()
      posInicial = 1
      posFinal = InStr(cadCommand, ";")
      'Tenemos que saber el idusuario para recuperar y guardar la moneda activa
      IdUsuario = Mid(cadCommand, 1, posFinal - posInicial)
      posInicial = posFinal + 1
      posFinal = InStr(posInicial, cadCommand, ";")
      NivelUsuario = Mid(cadCommand, posInicial, posFinal - posInicial)
      posInicial = posFinal + 1
      posFinal = InStr(posInicial, cadCommand, ";")
      NombreServidor = Mid(cadCommand, posInicial, posFinal - posInicial)
      posInicial = posFinal + 1
      posFinal = Len(cadCommand) + 1
      NombreBD = Mid(cadCommand, posInicial, posFinal - posInicial)

    Conecta
    RecuperaMoneda
    Set fMainForm = New frmMain
    Load fMainForm
    
    fMainForm.Show

End Sub

Public Sub ManejaError()
   Select Case Err.Number
     Case 6
        MsgBox "El valor introducido es demasiado grande"
     Case Else
        MsgBox "Se ha producido un error"
   End Select
End Sub

Sub RecuperaMoneda()
   Dim cmd As New Command
   Dim rs As New Recordset
   Dim prm As Parameter
   
   'definir el objeto command
   cmd.ActiveConnection = dB
   
   cmd.CommandText = "ValorMoneda"
   cmd.CommandType = adCmdStoredProc
   
   Set rs = cmd.Execute
   While (Not rs.EOF)
      ValorMoneda = rs(0)
      rs.MoveNext
   Wend
   Set rs = Nothing
   
   cmd.CommandText = "MonedaActiva"
   cmd.CommandType = adCmdStoredProc
   Set prm = cmd.CreateParameter("par1", adVarChar, adParamInput, 10, IdUsuario)
   cmd.Parameters.Append prm
   Set rs = cmd.Execute
   While (Not rs.EOF)
      MonedaActiva = rs(0)
      rs.MoveNext
   Wend
   Set rs = Nothing
   
End Sub


