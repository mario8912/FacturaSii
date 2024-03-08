Attribute VB_Name = "mdlTiempo"
'05/05/14 Implantar el salvapantallas
Public gintTiempoaEsperar As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetLastInputInfo Lib "user32" (plii As Any) As Long
Public Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

'150910 cerrar programa por inactividad
Public gMinutosCierre As Long
Public gbolCerrar As Boolean

