Attribute VB_Name = "ModuleUsu"
Option Explicit

'Public fUsuarioForm As frmUsuario

Public Db As Connection

Sub Main()
    Conecta
'    Set fUsuarioForm = New frmUsuario
'    Load fUsuarioForm
'
'    fUsuarioForm.Show
   MDIForm1.Show
End Sub


Public Sub Conecta()
'la conexi�n se realiza desde el from que abre �ste
  Set Db = New ADODB.Connection
  
  Db.CursorLocation = adUseClient
  
'RGN : Conexion a POLLO
'  db.Open "Provider=SQLOLEDB.1;Password=regina;Persist Security Info=True;User ID=regina;Initial Catalog=rosell;Data Source=pollo"
'RGN :Conexion a POLLO con la seguridad integrada de NT
Db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=rosell;Data Source=pollo"

'RGN : Conexion a ROSELL
'  db.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=rosell;Data Source=rosell"
End Sub

