Attribute VB_Name = "Module1"
Option Explicit


Public db As Connection

'################ APIS ########################
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public NombreServidor As String
Public NombreBD As String


Sub Main()
    IniciaValores
    Conecta
   
   frmAsistente.Show
End Sub
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
  NombreServidor = LeerIni("SERVIDOR", "Nombre", App.Path & ".\InicioAnual.ini")
  NombreBD = LeerIni("BASE DATOS", "Nombre", App.Path & ".\InicioAnual.ini")
End Sub



Public Sub Conecta()
    'la conexión se realiza desde el from que abre éste
    Set db = New ADODB.Connection
  
    db.CursorLocation = adUseClient
  
    'RGN :Conexion a ROSELL con la seguridad integrada de NT
    'db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=ROSELL"
    'RGN :Conexion a POLLO con la seguridad integrada de NT
    'db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=POLLO"
    'db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=ROSELL"
    
    'db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
    
    '26/05/14 Cambiada la conexion a la bd (para que funcione en maquinas virtuales)
    Dim strCad As String

    'strCad = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; password = ''; Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
    
    '05/08/14 Añadida la variable gVirtual
    Dim gVirtual As Boolean
    If gVirtual Then
''        strCad = "Provider=SQLOLEDB.1;"
''        strCad = strCad & "Persist Security Info=False;"
''        strCad = strCad & "User ID=" & gUsuario & ";"
''        strCad = strCad & "password =" & gClave & ";"
''        strCad = strCad & "Initial Catalog=" & NombreBD & ";"
''        strCad = strCad & "Data Source=" & NombreServidor
    Else
        strCad = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
        'strCad = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa; password = ''; Initial Catalog=" & NombreBD & ";Data Source=" & NombreServidor
    End If
    
    db.Open strCad
End Sub

