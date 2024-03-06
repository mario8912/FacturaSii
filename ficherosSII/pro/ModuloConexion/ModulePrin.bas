Attribute VB_Name = "Module1"
Option Explicit
'161108 Añadido blnCrystal11
Public blnCrystal11 As Boolean

'160630 Global gblnFacturaInversa  as Boolean
Global gblnFacturaInversa  As Boolean
'160630 Global gstrCad as string
Global gstrCad As String

Public db As Connection

Public MonedaActiva As Integer

Public IdUsuario As String

'RGN_250900 :cambio la tabla usuarios por la tabla UsuarioAplicacion para
'controlar el acceso a las distintas aplicaciones
Public rs As ADODB.Recordset

Public Const COLMonedaActiva = 0
Public Const COLPrincipal = 1
Public Const COLUsuarios = 2
Public Const COLMantenimiento = 3
Public Const COLEntradas = 4
Public Const COLGastos = 5
Public Const COLFacturacion = 6
Public Const COLCargas = 7
Public Const COLDiario = 8
Public Const COLActualizacion = 9
Public Const COLListados = 10

'################ APIS ########################
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public NombreServidor As String
Public NombreBD As String
'05/08/14 Añadida la variable gVirtual
Public gVirtual As Boolean

'26/05/14 Añadidas variables gUsuario y gClave
Public gUsuario As String
Public gClave As String


Function LeerIni(Seccion As String, entrada As String, fichero As String) As String
  
  Dim cadena As String * 255
  Dim valor As Long
  
  Screen.MousePointer = vbHourglass
  valor = GetPrivateProfileString(Seccion, ByVal entrada, "", cadena, Len(cadena), fichero)
  If valor = 0 Then
    LeerIni = ""
  Else
    LeerIni = Left(cadena, InStr(cadena, Chr(0)) - 1)
  End If
  Screen.MousePointer = vbDefault
  
End Function


Public Sub IniciaValores()
    '161108 Añadido blnCrystal11
    gstrCad = LeerIni("BASE DATOS", "cr11", App.Path & ".\Principal.ini")
    If UCase(gstrCad) = "SI" Then blnCrystal11 = True
    NombreServidor = LeerIni("SERVIDOR", "Nombre", App.Path & ".\Principal.ini")
    NombreBD = LeerIni("BASE DATOS", "Nombre", App.Path & ".\Principal.ini")
    '26/05/14 Añadidas variables gUsuario y gClave
    gUsuario = "sa"
    gClave = ""
      
    '05/08/14 Añadida la variable gVirtual
    Dim strArchivo As String
    strArchivo = App.Path & "\virtual.txt"
    If ArchivoExistente(strArchivo) Then
        gVirtual = True
    End If
End Sub

Public Function ArchivoExistente(Archivo As String) As Boolean
   Dim x
   On Error Resume Next
   x = GetAttr(Archivo)
   If Err Then ArchivoExistente = False Else ArchivoExistente = True
   On Error GoTo 0
End Function

