Attribute VB_Name = "mdlINI"
'210607 añadido mdlINI

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Función Api para bloquear el repintado de una ventana
Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long


Function ObtenerCampo(Campo As Integer, tCadena As Variant, Separador As String) As String
   Dim Cadena As String
   Dim cad As String
   Dim I As Integer
   Dim j As Integer
   Dim ini As Integer
   Dim Fin As Integer
   Cadena = tCadena
   'Restricciones:
   If Campo = 0 Or Cadena = "" Or Separador = "" Then ObtenerCampo = "": Exit Function
   If Len(Separador) > 1 Then ObtenerCampo = "": Exit Function
   Dim CadTmp() As String
   CadTmp = Split(tCadena, Separador)
   If Campo - 1 > UBound(CadTmp) Then ObtenerCampo = "": Exit Function
   ObtenerCampo = CadTmp(Campo - 1)
End Function

Sub EscribirINI(fichero As String, Seccion As String, Variable As String, valor As String)
   Dim I As Integer
   I = WritePrivateProfileString(Seccion, Variable, valor, fichero)
End Sub

Function LeerINI(fichero As String, Seccion As String, Variable As String) As String
   Dim I As Integer
   Dim cad As String * 255
   Dim Dato As String
   I = GetPrivateProfileString(Seccion, Variable, "", cad, Len(cad), fichero)
   If I = 0 Then
      LeerINI = "-88888888"
      Exit Function
   Else
      Dato = ""
      For I = 1 To 255
         If Asc(Mid(cad, I, 1)) = 0 Then Exit For
         Dato = Dato + Mid(cad, I, 1)
      Next
      LeerINI = Trim(Dato)
   End If
End Function
