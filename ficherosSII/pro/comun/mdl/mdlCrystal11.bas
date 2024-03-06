Attribute VB_Name = "mdlCrystal11"
'Global gstrficherolog As String


'190804 GRutaServidor en cargas
Global GRutaServidor As String

'190804 blnExisteCR11 an cargas
Global blnExisteCR11 As Boolean

Public Function hayCrystal11() As Boolean
    '190804 blnExisteCR11 an cargas
    Dim blnExiste As Boolean
    
    Dim path As String
    path = Environ("CommonProgramFiles")
    path = path & "\Business Objects\3.0\bin\craxdrt.dll"
    If ArchivoExistente(path) Then blnExiste = True
    hayCrystal11 = blnExiste
End Function

