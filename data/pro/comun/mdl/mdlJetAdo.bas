Attribute VB_Name = "mdlJetAdo"
Option Explicit
Global BDado As ADODB.Connection
Global NombreBDAdo As String
Global NombreServidorADO As String

Public Function RegistroExistente(Tabla As String, Campo As String, valor As Variant, Optional Numero As Boolean = False) As Boolean
    '190717 RegistroExistente
    Dim strCad As String
    strCad = strCad
    strCad = "Select (" & Campo & ") as num from " & Tabla & " where " & Campo
    If Numero Then
        strCad = strCad & " = " & valor
    Else
        strCad = strCad & " = '" & valor & "'"
    End If
    
    Dim oComando As ADODB.Command
    Dim rsTmp As New ADODB.Recordset
    Set oComando = New ADODB.Command
    Set oComando.ActiveConnection = DB
    
    oComando.CommandText = strCad
    oComando.CommandType = adCmdText
    Set rsTmp = oComando.Execute
    
    If Not TVado(rsTmp) Then
        RegistroExistente = True
    End If
    CRado rsTmp

End Function

Public Sub IniciaValores()
    Dim archivo As String
    archivo = App.Path & "\leerexcel.ini"
    NombreServidorADO = LeerINI(archivo, "SERVIDOR", "Nombre")
    NombreBDAdo = LeerINI(archivo, "BASE DATOS", "Nombre")
End Sub

Public Function ConectaSql1(BD As ADODB.Connection) As Boolean
    On Error GoTo Fallo
    'IniciaValores
    BD.CursorLocation = adUseClient
    Dim strCad As String
    strCad = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
    BD.Open strCad
    'NombreBD
    'NombreServidor
    On Error GoTo 0
    ConectaSql1 = True
    Exit Function
Fallo:
    Err.Raise Err.Number, "ConectaSql " & Erl, Err.Description
End Function

Public Function ConectaSql(BD As ADODB.Connection) As Boolean
10        On Error GoTo Fallo
20        IniciaValores
30        BD.CursorLocation = adUseClient
          Dim strCad As String
40        strCad = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBDAdo & ";Data Source=" & NombreServidorADO
50        BD.Open strCad
          
60        On Error GoTo 0
70        ConectaSql = True
80        Exit Function
Fallo:
90        Err.Raise Err.Number, "ConectaSql " & Erl, Err.Description
End Function

'
'
'Function ConectaADObase(BD As ADODB.Connection) As Boolean
'    'esta funcion es similar a ConectaADObd
'    'a esta se le pasa ya una bd, que esta abierta en DAO
'
'    Dim Base As String
'    Dim Cad As String
'    Base = BD.Name
'    On Error GoTo FAllo
'    Set BD = New ADODB.Connection
'    Cad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
'    Cad = Cad & Base
'    Cad = Cad & ";Persist Security Info=False"
'    BD.Open Cad
'    ConectaADObase = True
'    On Error GoTo 0
'    Exit Function
'FAllo:
'    Err.Raise Err.Number, , Err.Description
'End Function
Public Sub DesconectaADObd(BD As ADODB.Connection)
    On Error GoTo Fallo
        BD.Close
        Set BD = Nothing
    On Error GoTo 0
    Exit Sub
Fallo:
    Err.Raise Err.Number, , Err.Description
End Sub

Function ConectaADObd(BD As ADODB.Connection, RutaBase As String) As Boolean
    'Dim Base As String
    Dim cad As String
    'Base = BD.Name
    On Error GoTo Fallo
    Set BD = New ADODB.Connection
    cad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    cad = cad & RutaBase
    cad = cad & ";Persist Security Info=False"
    BD.Open cad
    ConectaADObd = True
    On Error GoTo 0
    Exit Function
Fallo:
    Err.Raise Err.Number, , Err.Description
End Function
Public Function ARado(ByRef Base As ADODB.Connection, ByRef Tabla As ADODB.Recordset, sql As String) As Boolean
    On Error GoTo Fallo
    'Tabla.Open Sql, Base, adOpenKeyset, adLockBatchOptimistic, Tipo
    Tabla.Open sql, Base, adOpenKeyset, adLockBatchOptimistic, 1  ' 1 = CommandTypeEnum.adCmdText
    ARado = True
    On Error GoTo 0
    Exit Function
Fallo:
    MsgBox Err.Number & " " & Err.Description
    'Tabla.Fields(2).Type = adBigInt
    
End Function


Sub CRado(ByRef Tabla As ADODB.Recordset)
    'CR  = CerrarRecordset
    On Error Resume Next
    Tabla.Close
    Set Tabla = Nothing
    On Error GoTo 0
End Sub
Function TipoCampo(intTipo As Integer) As String
    Select Case intTipo
        Case dbBoolean
            TipoCampo = "dbBoolean"
        Case dbByte
            TipoCampo = "dbByte"
        Case dbInteger
            TipoCampo = "dbInteger"
        Case dbLong
            TipoCampo = "dbLong"
        Case dbCurrency
            TipoCampo = "dbCurrency"
        Case dbSingle
            TipoCampo = "dbSingle"
        Case dbDouble
            TipoCampo = "dbDouble"
        Case dbDate
            TipoCampo = "dbDate"
        Case dbText
            TipoCampo = "dbText"
        Case dbLongBinary
            TipoCampo = "dbLongBinary"
        Case dbMemo
            TipoCampo = "dbMemo"
        Case dbGUID
            TipoCampo = "dbGUID"
    End Select
End Function

Function TipoRecordset(intTipo As Integer) As String
   Select Case intTipo
      Case dbOpenTable
         TipoRecordset = "dbOpenTable"
      Case dbOpenDynaset
         TipoRecordset = "dbOpenDynaset"
      Case dbOpenSnapshot
         TipoRecordset = "dbOpenSnapshot"
      Case dbOpenForwardOnly
         TipoRecordset = "dbOpenForwardOnly"
   End Select
End Function

Function TVado(ByRef Tabla As ADODB.Recordset) As Boolean
     'Dim marcador As Variant
     On Error Resume Next
     'marcador = Tabla.Bookmark
     Tabla.MoveLast
     Tabla.MoveFirst
     TVado = Tabla.BOF And Tabla.EOF
     'Tabla.Bookmark = marcador
     On Error GoTo 0
End Function

Function BRado(ByRef Tabla As ADODB.Recordset, Valor1, Optional EsNumeroV1 As Boolean = False, _
                        Optional Valor2, Optional EsNumeroV2 As Boolean = False) As Boolean
'--> Se situa en el registro que cumple las condiciones de los valores de busqueda
'--> Busca en toda la tabla, desde el principio hasta el final
'--> y devuelve un valor booleano (true si encuentra el registro y false si no lo encuenta)
'--> @param Tabla       - Nombre del recordset en el que buscara el registro
'--> @param Valor1      - El valor que se quiere buscar en el primer campo de la tabla
'--> @param EsNumeroV1  - Indica si Valor1 es numerico, por defecto False = (no numerico)
'--> @param Valor2      - El valor que se quiere buscar en el segundo campo de la tabla
'--> @param EsNumeroV2  - Indica si Valor2 es numerico, por defecto False = (no numerico)
'--> @return Devuelve un valor booleano (True si existe, y False si no existe)

    Dim Seguir As Boolean
    On Error GoTo ErrorBuscando

    Dim Cadena As String
    'si no hay registros en la tabla, sale de la funcion
    '--> @sub TablaVacia
    If TVado(Tabla) Then BRado = False: Exit Function

    'se contruye la cadena para buscar el primer valor, dependiendo de EsNumeroV1
    If EsNumeroV1 Then
        Cadena = Tabla.Fields(0).Name & "=" & Valor1
    Else
        Cadena = Tabla.Fields(0).Name & "= '" & Valor1 & "'"
    End If

    'se situa en el primer registro de la tabla
    Tabla.MoveFirst
    'busca el primer valor

    'revisar
    'Tabla .Find Cadena

    'si el valor es EOF, no se ha encontrado el registro
    If Tabla.EOF Then Exit Function

    'si se le ha pasado el Valor2, se busca el valor en el segundo campo de la tabla
    If Not IsMissing(Valor2) Then
        'se contruye la cadena para buscar el segundo valor, dependiendo de EsNumeroV2
        If EsNumeroV2 Then
            Cadena = Tabla.Fields(1).Name & "=" & Valor2
        Else
            Cadena = Tabla.Fields(1).Name & "= '" & Valor2 & "'"
        End If

        'si Valor2 es igual al segundo campo de la tabla, ya hemos encontrado el registro
        'en caso contrario, seguimos buscando a partir del registro en que nos encontramos
        If Tabla.Fields(1) <> Valor2 Then
            'Tabla.Find Cadena, Tabla.AbsolutePosition, adSearchForward
            'Tabla.Find Cadena, Tabla.AbsolutePosition, adSearchBackward
            Seguir = True
                    If Not Tabla.EOF Then
                    Tabla.MoveNext
                    End If

            If Not Tabla.EOF Then
                While Not Tabla.EOF And Seguir
                    If Tabla.Fields(1) = Valor2 Then
                        Seguir = False
                    Else
                        If Not Tabla.EOF Then
                        Tabla.MoveNext
                        Else
                            Exit Function
                        End If
                    End If
                Wend
            End If

        End If

        'si el valor es EOF, no se ha encontrado el registro
        If Tabla.EOF Then Exit Function

        'si el valor del primer campo de la tabla es difernte a Valor1, no se ha encontrado el registro
        If Tabla.Fields(0) <> Valor1 Then Exit Function
    Else
        Tabla.Find Cadena
        If Tabla.EOF Then
            Exit Function
        End If
    End If
    On Error GoTo 0

    BRado = True
    Exit Function
ErrorBuscando:

MsgBox "Se ha producido el error " & Err.Number & " " & Err.Description
End Function


Function URado(ByRef Tabla As ADODB.Recordset, Optional Key1, Optional Key2, Optional KEY3) As Long
On Error GoTo AlgoFallO
    If TVado(Tabla) Then
        URado = 1
        Exit Function
    End If

    Tabla.MoveLast
    URado = Tabla.Fields(0) + 1

On Error GoTo 0
Exit Function

AlgoFallO:
MsgBox Err.Number & " " & Err.Description
End Function


Public Function ArSql(DB As ADODB.Connection, Rs As ADODB.Recordset, sql As String) As Boolean
10        On Error GoTo Fallo
          Dim oComando As ADODB.Command
          'Dim rsAux As ADODB.Recordset
         
20        Set oComando = New ADODB.Command
         
30        Set oComando.ActiveConnection = DB

40        oComando.CommandText = sql
50        oComando.CommandType = adCmdText
60        Set Rs = oComando.Execute
70        Set oComando = Nothing
80        ArSql = True
90        On Error GoTo 0
100       Exit Function
Fallo:
110       Err.Raise Err.Number, "ArSql " & Erl, Err.Description
End Function

'datatypeenum
Public Function ArSql_1(DB As ADODB.Connection, Rs As ADODB.Recordset, sql As String, Optional strParametro As String, Optional TipoParametro As DataTypeEnum, Optional vValorParametro As Variant) As Boolean
    On Error GoTo Fallo
    Dim strParamNombre As String
    Dim vntParamValor As Variant
    Dim oComando As ADODB.Command
    Dim oParametro As ADODB.Parameter
    
    Dim arrParam() As Variant
    Dim arrParam1()
    Dim intConta As Integer
    Dim strCad As String
    
    Set oComando = New ADODB.Command
   
    Set oComando.ActiveConnection = DB

    oComando.CommandText = sql
    'oComando.CommandType = adCmdText
    oComando.CommandType = adCmdStoredProc
    '
    If Not IsMissing(strParametro) Then
        'arrParam = Split(strParametro, "=")
'        For intConta = 0 To UBound(arrParam)
'            arrParam1 = Split(arrParam, "=")
'            strParamNombre = arrParam1(0)
'            vntParamValor = arrParam1(1)
'            oComando.Parameters(strParamNombre) = vntParamValor
'        Next
        
        
        oComando.Parameters(strParametro) = vValorParametro
'            Set oParametro = oComando.CreateParameter(strParametro, TipoParametro, adParamInput, , vValorParametro)
'            oComando.Parameters.Append oParametro
    End If
    
    Set Rs = oComando.Execute
    Set oComando = Nothing
    ArSql_1 = True

    On Error GoTo 0
    Exit Function
Fallo:
    Err.Raise Err.Number, "ArSql " & Erl, Err.Description
End Function
Public Function EjecutarSecuenciaSQL(sql As String) As Boolean
    On Error GoTo Fallo:
          Dim oComando As ADODB.Command
10        Set oComando = New ADODB.Command
20        Set oComando.ActiveConnection = DB
30        oComando.CommandText = sql
40        oComando.CommandType = adCmdText
50        oComando.Execute
60        On Error GoTo 0
70        EjecutarSecuenciaSQL = True
80        Exit Function
Fallo:
90        Msg mError, "", Err, Erl
End Function
Public Function DatosCambiados() As Boolean
          Dim oComando As ADODB.Command
          '190706 cambios en campos
          'Dim rsAux As ADODB.Recordset
10        On Error GoTo Fallo
20        Set oComando = New ADODB.Command
         
30        Set oComando.ActiveConnection = DB
40        oComando.CommandText = "InsertarCambios"
50        oComando.CommandType = adCmdStoredProc
60        oComando.Parameters("@tabla") = GstrTCtabla
70        oComando.Parameters("@Campo") = GstrTCcampo
80        oComando.Parameters("@Registro") = GstrTCRegistro
90        oComando.Parameters("@FECHA") = GstrTCFECHA
100       oComando.Parameters("@USUARIO") = GstrTCUSUARIO
110       oComando.Parameters("@Maquina") = GstrTCMaquina
120       oComando.Parameters("@Anterior") = GstrTCAnterior
130       oComando.Parameters("@Actual") = GstrTCActual
140       oComando.Execute
          DatosCambiados = True
150       On Error GoTo 0
160       Exit Function
Fallo:
170       Err.Raise Err.Number, "DatosCambiados " & Erl, Err.Description
End Function

Public Function EjecutarProcedimiento(Procedimiento As String) As Boolean
    Dim oComando As ADODB.Command
    'Dim rsAux As ADODB.Recordset
   
    Set oComando = New ADODB.Command
   
    Set oComando.ActiveConnection = DB
    oComando.CommandText = Procedimiento
    oComando.CommandType = adCmdStoredProc
    oComando.Execute
    
    Dim intResultado As Integer
    
    intResultado = oComando.Parameters(0).Value
    EjecutarProcedimiento = intResultado
End Function
'Optional strParametro As String, Optional TipoParametro As DataTypeEnum, Optional vValorParametro As Variant) As Boolean
Public Function EjecutarProcedimientoParametro(Procedimiento As String, Optional strParametro As String, Optional TipoParametro As DataTypeEnum, Optional vValorParametro As Variant) As Boolean
    Dim oComando As ADODB.Command
    'Dim rsAux As ADODB.Recordset
   
    Set oComando = New ADODB.Command
   
    Set oComando.ActiveConnection = DB
    oComando.CommandText = Procedimiento
    oComando.CommandType = adCmdStoredProc
    
    If Not IsMissing(strParametro) Then
        oComando.Parameters(strParametro) = vValorParametro
    End If
    
    oComando.Execute
    
    Dim intResultado As Integer
    
    intResultado = oComando.Parameters(0).Value
    EjecutarProcedimientoParametro = intResultado
End Function

Public Function RegistroSiguiente(Tabla As String, Campo As String, Numero As Long) As Long
    Dim strCad As String
    strCad = strCad
    strCad = "Select min(" & Campo & ") as num from " & Tabla & " where " & Campo & " > " & Numero
    RegistroSiguiente = Numero
    
    strCad = "Select min(" & Campo & ") as num from " & Tabla
    If Numero <> 0 Then
        'Beep
        'Exit Function
        strCad = strCad & " where " & Campo & " > " & Numero
    End If
    
    Dim oComando As ADODB.Command
    Dim rsTmp As New ADODB.Recordset
    Set oComando = New ADODB.Command
    Set oComando.ActiveConnection = DB
    
    oComando.CommandText = strCad
    oComando.CommandType = adCmdText
    Set rsTmp = oComando.Execute
    
    If Not TVado(rsTmp) Then
        If Not IsNull(rsTmp!num) Then
            RegistroSiguiente = rsTmp!num
        Else
             Beep
            RegistroSiguiente = Numero
        End If
    End If
    CRado rsTmp
End Function


Public Function RegistroAnterior(Tabla As String, Campo As String, Numero As Long) As Long
    Dim strCad As String
    strCad = strCad
    
    strCad = "Select max(" & Campo & ") as num from " & Tabla & " where " & Campo & " < " & Numero
    
    strCad = "Select max(" & Campo & ") as num from " & Tabla
    If Numero <> 0 Then
         strCad = strCad & " where " & Campo & " < " & Numero
    End If
    
    RegistroAnterior = Numero
    Dim oComando As ADODB.Command
    Dim rsTmp As New ADODB.Recordset
    Set oComando = New ADODB.Command
    Set oComando.ActiveConnection = DB
    
    oComando.CommandText = strCad
    oComando.CommandType = adCmdText
    Set rsTmp = oComando.Execute
    
    If Not TVado(rsTmp) Then
        If Not IsNull(rsTmp!num) Then
        RegistroAnterior = rsTmp!num
        Else
            Beep
            RegistroAnterior = Numero
        End If
    End If
    CRado rsTmp
End Function

