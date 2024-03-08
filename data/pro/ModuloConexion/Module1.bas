Attribute VB_Name = "Module1"
Option Explicit
'221020 blnCancelado
Global blnCancelado As Boolean
Global bolContinuar As Boolean

Public DB As ADODB.Connection

Public ValorMoneda As Currency

Public NivelUsuario As Byte
Public NombreServidor As String
Public NombreBD As String

'26/05/14 Añadidas variables gUsuario y gClave en Gastos.vbp
Public gUsuario As String
Public gClave As String
'25/07/14 Añadido gdblPago
Global gdblPago As Double

'Global gVirtual As Boolean


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
    DB.CursorLocation = adUseClient
'''
'''    'RGN :Conexion a POLLO con la seguridad integrada de NT
'''    DB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor


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
    
    DB.Open strCad

End Sub

Sub Main()
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
'    NivelUsuario = Command()
'    'NivelUsuario = 9
'   MDIForm1.Show
Dim cadCommand As String
Dim posInicial As Integer
Dim posFinal As Integer

    cadCommand = Command()
    posInicial = 1
    posFinal = InStr(cadCommand, ";")
    NivelUsuario = Mid(cadCommand, 1, posFinal - posInicial)
    posInicial = posFinal + 1
    posFinal = InStr(posInicial, cadCommand, ";")
    NombreServidor = Mid(cadCommand, posInicial, posFinal - posInicial)
    posInicial = posFinal + 1
    posFinal = Len(cadCommand) + 1
    NombreBD = Mid(cadCommand, posInicial, posFinal - posInicial)
    
'    NivelUsuario = Command()
    
    '220526 Cambiar la base de datos cuando ESIDE
    strArchivo = "..\principal\principal.ini"
    gEsIde = EsIde
    If gEsIde Then
        NombreServidor = LeerINI(strArchivo, "SERVIDOR", "Nombre")
        NombreBD = LeerINI(strArchivo, "BASE DATOS", "BASEIDE")
    End If
    '220526 Cambiar la base de datos cuando ESIDE fin
   
    '26/05/14 Añadidas variables gUsuario y gClave
    gUsuario = "sa"
    gClave = ""
    
    Conecta
    RecuperaMoneda
    MDIForm1.Show

End Sub
