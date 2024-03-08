Attribute VB_Name = "mdlArticulos"
Option Explicit

'170126 añadido typoArticulo
Type tArticulo
    tidarticulo As String
    tdescripcion As String
    tidFamilia As String
    tEliminado As Boolean
End Type
'221124 blnQuitandoFiltro
Global blnQuitandoFiltro As Boolean

Public Function ObtenerDatosArticulo(vIdArticulo As Variant) As tArticulo
'    '161103 ObtenerFamiliaArticulo
'    ObtenerFamiliaArticulo = ""
'    Dim rsFamilia As New ADODB.Recordset
'    Dim oCommand As ADODB.Command
'    Dim oParametro As ADODB.Parameter
'
'    Set oCommand = New ADODB.Command
'    Set oCommand.ActiveConnection = dB
'    oCommand.CommandText = "ObtenerFamilia"
'    oCommand.CommandType = adCmdStoredProc
'    Set oParametro = oCommand.CreateParameter("IdArticulo", adVarChar, adParamInput, 4, vIdArticulo)
'    oCommand.Parameters.Append oParametro
'    Set rsFamilia = oCommand.Execute
'
'    Set oCommand = Nothing
'    If rsFamilia.RecordCount < 1 Then Exit Function
'    ObtenerFamiliaArticulo = rsFamilia.Fields("IdFamilia").Value
'    Set rsFamilia = Nothing
    
    ObtenerDatosArticulo.tidarticulo = ""
    ObtenerDatosArticulo.tdescripcion = ""
    ObtenerDatosArticulo.tidFamilia = ""
    ObtenerDatosArticulo.tEliminado = False
    Dim rs As Recordset
    Set rs = New ADODB.Recordset
    MiSql = "select idarticulo, descripcion, idFamilia, eliminado"
    MiSql = MiSql & " from articulo where idarticulo ='" & vIdArticulo & "'"
    On Error GoTo Fallo
    If Not AbrirRecordset(DB, rs, MiSql, adCmdText) Then GoTo Fallo
    If Not TablaVacia(rs) Then
        ObtenerDatosArticulo.tidarticulo = rs!IdArticulo
        ObtenerDatosArticulo.tdescripcion = rs!Descripcion
        ObtenerDatosArticulo.tidFamilia = rs!Idfamilia
        ObtenerDatosArticulo.tEliminado = rs!eliminado
    End If
    CerrarRecordset rs
    On Error GoTo 0
    
    Exit Function
Fallo:
    Dim qOrigen As String
    qOrigen = "ObtenerDatosArticulo " & Erl
    Err.Raise Err.Number, qOrigen, Err.Description
    
End Function

