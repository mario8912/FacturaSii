Attribute VB_Name = "mdlFicheros"
Option Explicit

Public Function EsIndice(frm As String, ctrl As String) As Boolean
    '200625 EsIndice
    Dim Encontrado As Boolean
    Dim mivariable As String
    Dim Partes() As String
    
    Open RutaCFG & "indices.txt" For Input As #1
    
    While Not EOF(1) And Not Encontrado
    
        Line Input #1, mivariable
        Partes = Split(mivariable, ".")
        If Partes(0) = frm Then
            If Partes(1) = ctrl Then
                Encontrado = True
                
            End If
        End If
    Wend
    Close #1
    
    EsIndice = Not Encontrado
End Function

Public Function LeerFichero(fichero As String) As String
    Dim FNr As Integer, S As String

    If Not ArchivoExistente(fichero) Then
        LeerFichero = ""
        MsgBox ("no existe el fichero")
        Exit Function
    End If
    FNr = FreeFile
    Open fichero For Binary As #FNr
    S = Space$(LOF(FNr)): Get #FNr, , S
    Close #FNr
    LeerFichero = S
End Function

'Public Function ArchivoExistente(Archivo As String) As Boolean
'   Dim X
'   On Error Resume Next
'   X = GetAttr(Archivo)
'   If Err Then ArchivoExistente = False Else ArchivoExistente = True
'   On Error GoTo 0
'End Function
'
