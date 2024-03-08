Attribute VB_Name = "ModIde"
Option Explicit
'24/04/14 Añadido el modulo ModIde (este modulo esta en la carpeta ModuloConexion)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Global gEsIde As Boolean

Public Function EsIde() As Boolean
    EsIde = GetModuleHandle("vb6.exe")
    Dim ReturnVal As Long
    ReturnVal = GetModuleHandle("vb6.exe")
    If ReturnVal <> 0 Then EsIde = True
End Function

