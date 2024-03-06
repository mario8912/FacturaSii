Attribute VB_Name = "ModuleEntTal"

Public Sub BorraFichero(vFichero As String)
On Error GoTo ControlError
Dim fs

   Set fs = CreateObject("Scripting.FileSystemObject")
   
   If fs.FileExists(vFichero) Then
      fs.deletefile vFichero
   End If
   
Exit Sub
ControlError:
End Sub

Public Function ReferenciaEnTalon(pIdReferencia As Long, pCVG As String) As Integer
   ''16/06/14 Añadida la funcion ReferenciaEnTalon
   Dim oCadenaSQL As String
   Dim oComando As ADODB.Command
   Dim oParametro As ADODB.Parameter
   Dim rsAux As ADODB.Recordset
   
   Set oComando = New ADODB.Command
   
   Set oComando.ActiveConnection = DB
   
   oCadenaSQL = " SELECT T.IdTalon, TL.IdReferencia" & _
                  " FROM TalonLinea TL INNER JOIN Talon T ON TL.IdTalon = T.IdTalon " & _
                  " WHERE TL.IdReferencia = ? " & _
                  " AND UPPER(TL.CVG) = ? " & _
                  " AND T.Actualizado = 0 "
                  
   oComando.CommandText = oCadenaSQL
   oComando.CommandType = adCmdText
   
   Set oParametro = oComando.CreateParameter("IdReferencia", adInteger, adParamInput, , pIdReferencia)
   oComando.Parameters.Append oParametro
   Set oParametro = oComando.CreateParameter("CVG", adVarChar, adParamInput, 1, pCVG)
   oComando.Parameters.Append oParametro

   Set rsAux = oComando.Execute
   
   If rsAux.RecordCount > 0 Then
      rsAux.MoveFirst
      ReferenciaEnTalon = rsAux("IdTalon").Value
      'ReferenciaEnTalon = True
   Else
      'ReferenciaEnTalon = False
      ReferenciaEnTalon = 0
   End If
   
   Set oComando = Nothing
   Set rsAux = Nothing
End Function


