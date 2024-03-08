Attribute VB_Name = "Module1"
Option Explicit

Public DB As ADODB.Connection

Public ValorMoneda As Currency

Public NivelUsuario As Byte
Public NombreServidor As String
Public NombreBD As String
Public OrigenLLamada As String
Public pIdBanco As String
Public gdblImporteAPagar As Double

'27/07/14 Añadida la variable lngCtaBanco
Public lngCtaBanco As Long

'26/05/14 Añadidas variables gUsuario y gClave
Public gUsuario As String
Public gClave As String

'150602 dtFechaVto
Public dtFechaVto As Date

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
    ''''RGN :Conexion a POLLO con la seguridad integrada de NT
    '''DB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor

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
    On Error GoTo errMain
    
    '05/08/14 Añadida la variable gVirtual
    Dim strArchivo As String
    strArchivo = "..\principal\virtual.txt"
    
    If ArchivoExistente(strArchivo) Then
        gVirtual = True
    End If
    
    
    ''27/07/14 Añadida la variable lngCtaBanco
    'esto se leera desde un ini, de momento lo pongo a pelo
    lngCtaBanco = 570000000
    
    '28/05/14 añadida funcion QueSigno (para obtener el signo decimal del sistema)
    Call QueSigno
    '27/07/14 Aladida la variable lngCtaBanco
    vValT = "#,##0.00"
    
    '24/04/14 Añadido el modulo ModIde (este modulo esta en la carpeta ModuloConexion)
    'si EsIde esta en tiempo de diseño, osea, ejecutando desde el entorno de visual
    gEsIde = EsIde

'    Conecta
'    RecuperaMoneda
'    NivelUsuario = Command()
'    'NivelUsuario = 9
'   MDIForm1.Show

'25/07/14 Añadido gPagadoPor
gPagadoPor = pNinguno

Dim cadCommand As String
Dim posInicial As Integer
Dim posFinal As Integer
Dim intCuantosParametros As Integer

    'llamada desde el mdi de gastos
    '9;(LOCAL);rosell;2;0;1
    
   cadCommand = Command()
   intCuantosParametros = CuantosParametros(cadCommand)
   'nivelusuario ; servidor ; base datos ; origenllamada ; Banco ; Importe
   
   '9;(local);rosell;2;0
   '9;(LOCAL);rosell;2;0;1
   '..\pagos\pagos.exe 9;server2017;rosell;1;2;2;3246540;28/10/2019
   
   
   'OrigenLamada
   '0 entradas
   '1 pago de entradas
   '2 pago de gastos
   
   'parametro 1
   posInicial = 1
   posFinal = InStr(cadCommand, ";")
   NivelUsuario = Mid(cadCommand, 1, posFinal - posInicial)
   
   'parametro 2
   posInicial = posFinal + 1
   posFinal = InStr(posInicial, cadCommand, ";")
   NombreServidor = Mid(cadCommand, posInicial, posFinal - posInicial)
   
   'parametro 3
   posInicial = posFinal + 1
   posFinal = InStr(posInicial, cadCommand, ";")
   NombreBD = Mid(cadCommand, posInicial, posFinal - posInicial)
   
   'parametro 4
   posInicial = posFinal + 1
   posFinal = InStr(posInicial, cadCommand, ";")
   
   '9;(local);rosell;2;0
   'si viene de gastos, se le pasa un 2
   OrigenLLamada = Mid(cadCommand, posInicial, posFinal - posInicial)
   
   'parametro  5
   posInicial = posFinal + 1
   'posFinal = Len(cadCommand) + 1
   posFinal = InStr(posInicial, cadCommand, ";")
   
   'si viene de gastos, se le pasa el numero de banco
   pIdBanco = Mid(cadCommand, posInicial, posFinal - posInicial)
    
    'revisar
    'igual esto hay que ponerlo antes del banco (mirar si se pasa siempre o no)
    
    '25/07/14 Añadido gPagadoPor
    'gPagadoPor  =     pNinguno = 0;    pCaja = 1;    pBanco = 2;    pTalon = 3
    
    'el ultimo parametro es la forma de pago = gPagadoPor
    
    'OrigenLamada
    '0 entradas
    '1 pago de entradas
    '2 pago de gastos
    
    'esta cadema viene de entradas pagado con talon
    '9;manolo-pc;rosell;0;4;3
    
    'esta cadeona vienes de pagos de entradas
    '9;manolo-pc;rosell;1;2;3
    
    
    'esta cadena viene de gastos pagado con talon
    '9;(LOCAL);rosell;2;4;3
    
    'esta cadena viene de gastos pagado por caja
    '9;(LOCAL);rosell;2;0;1
    
    'esta cadena viene de gastos pagado por banco
    '9;(LOCAL);rosell;2;4;2
   
    'parametro 6
    '05/11/14 Añadida la funcion CuantosParametros
    'segun de donde venga, este puede ser el ultimo parametro
    
    posInicial = posFinal + 1
    'posinicial = 22
    

    If intCuantosParametros = 6 Then
        posFinal = Len(cadCommand) + 1
        'posInicial = posFinal + 1
    Else
        'posFinal = Len(cadCommand) + 1
        posFinal = InStr(posInicial, cadCommand, ";")
    End If
    
    '9;(LOCAL);rosell;2;0;1;24200
    '9;(LOCAL);rosell;2;0;1;24200
    
    'gPagadoPor  =     pNinguno = 0;    pCaja = 1;    pBanco = 2;    pTalon = 3
    gPagadoPor = Mid(cadCommand, posInicial, posFinal - posInicial)
    '..\pagos\pagos.exe 9;(LOCAL);rosell;2;0;1

    '..\pagos\pagos.exe 9;manolo-pc;rosell;0;0;1;64250
    '                   9;server2017;rosell;1;2;2;3246540;28/10/2019
    
    '                   9;server2017;rosell;1;2;2;219780;28/10/2019
    If intCuantosParametros > 6 Then
        If OrigenLLamada = 0 Or OrigenLLamada = 1 Then
            posInicial = posFinal + 1
            posFinal = InStr(posInicial, cadCommand, ";")
            'Public gdblImporteAPagar As Double
            gClave = Mid(cadCommand, posInicial, posFinal - posInicial)
            gdblImporteAPagar = DD(CDbl(gClave) / 100)
        
            posInicial = posFinal + 1
            posFinal = Len(cadCommand) + 1
            gClave = Mid(cadCommand, posInicial, posFinal - posInicial)
            dtFechaVto = CDate(gClave)
        Else
            posInicial = posFinal + 1
            posFinal = Len(cadCommand) + 1
            'Public gdblImporteAPagar As Double
            gClave = Mid(cadCommand, posInicial, posFinal - posInicial)
            gdblImporteAPagar = DD(CDbl(gClave) / 100)
        End If
    
    End If
'    NivelUsuario = Command()
    
    '26/05/14 Añadidas variables gUsuario y gClave
    gUsuario = "sa"
    gClave = ""
    
    Conecta
    RecuperaMoneda
    MDIForm1.Show
    
    On Error GoTo 0
    Exit Sub
errMain:
    MsgBox Err.Number & " " & Err.Description & " " & Erl
    'Resume
End Sub
