Attribute VB_Name = "mdlSQL"
'161010 Añadido el modulo mdlSQL
Option Explicit

'161010 Añadida la variable qOrigen en mdlSQL (para control de errores)
Public qOrigen As String
Public Function ArSqlTXT(qRS As ADODB.Recordset, sql As String) As Boolean
    On Error GoTo Fallo
    Dim cmd As New Command
    'Dim rs As New Recordset
    Dim intPos As Integer
    
    intPos = InStr(sql, " ")
    
    cmd.ActiveConnection = dB
    
    cmd.CommandText = sql
    
    
    If intPos = 0 Then
        'esto abriria una tabla
        cmd.CommandType = adCmdTable
    Else
        'esto abre un select
        cmd.CommandType = adCmdText
    End If
    
    'Set rs = cmd.Execute
    'Set ArSqlTXT = rs
    'Set rs = Nothing
    Set qRS = cmd.Execute
    On Error GoTo 0
    
    ArSqlTXT = True
    
    Exit Function
Fallo:
    Dim objError As ADODB.Error
    
    qOrigen = "ArSqlTXT" & " lin " & Erl
    
    If dB.Errors.Count > 0 Then
    For Each objError In dB.Errors
        'aqui habria que comprobar si siempre hay una sola linea o no
        Err.Raise objError.NativeError, qOrigen, objError.Description
    Next
    End If
End Function

