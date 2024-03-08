Attribute VB_Name = "ModuleCopia"
Option Explicit

Public DB As ADODB.Connection

Public Sub Conecta()
'la conexión se realiza desde el from que abre éste
  Set DB = New ADODB.Connection
  DB.CursorLocation = adUseClient
  DB.CommandTimeout = 0
  
'RGN :Conexion a POLLO con la seguridad integrada de NT
DB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ROSELL;Data Source=pollo"

End Sub

Sub Main()
    Conecta
   MDIForm1.Show
End Sub

