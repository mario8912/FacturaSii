Attribute VB_Name = "mdlFechaIva"
Option Explicit
Global strCad As String
Global gdtnFechaIva As Date

Public Sub qFechaIva()
    On Error GoTo Fallo
    strCad = "fechaiva"
    
    Dim rsAux As New ADODB.Recordset
    If Not ArSql(DB, rsAux, strCad) Then GoTo Fallo
    
    If Not TVado(rsAux) Then
        gdtnFechaIva = rsAux!fechaiva
    Else
        strCad = "01/01/" & Format(CDate(Now), "YYYY")
        gdtnFechaIva = CDate(strCad)
    End If
    
    Set rsAux = Nothing
    
    On Error GoTo 0
    Exit Sub
Fallo:
    Err.Raise Err.Number, "qFechaIva " & Erl, Err.Description
End Sub

Public Function FechaIVABuena(qFecha As Date) As Boolean
    FechaIVABuena = False
    qFechaIva
    If qFecha <= gdtnFechaIva Then Exit Function
    FechaIVABuena = True
End Function

Public Function FechaHoy(qFecha As Date) As Boolean
    FechaHoy = False
    Dim dtmHoy As Date
    dtmHoy = Date
    If qFecha > dtmHoy Then Exit Function
    FechaHoy = True
End Function

Public Sub ComprobarFecha(Ctrl As Control, KeyAscii As Integer)
    'ComprobarFecha txtdesdeFecha, KeyAscii
    'si es la tecla borrar
    If KeyAscii = 8 Then Exit Sub
    'si se pulsa intro
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}", True
    End If
    'si el caracter es de 0 a 9
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        If Ctrl.SelStart = 0 And Ctrl.SelLength = Len(Ctrl) Then Exit Sub
        'If Ctrl.SelStart = 0 And Ctrl.SelLength = 10 Then Exit Sub
            If Len(Ctrl.Text) = 2 Or Len(Ctrl.Text) = 5 Then
               Ctrl.Text = Ctrl.Text & "/"
               Ctrl.SelStart = Len(Ctrl.Text)
            End If
        Exit Sub
    End If
    KeyAscii = 0
End Sub

Public Function ValidarFecha(Ctrl As Control) As Boolean
    'Cancel = ValidarFecha(txtdesdeFecha)
    If Len(Ctrl.Text) = 8 Then Ctrl.Text = Mid(Ctrl.Text, 1, 6) & Mid(Format(Date, "yyyy"), 1, 2) & Mid(Ctrl.Text, 7)
    If Len(Ctrl.Text) = 9 Then Ctrl.Text = Mid(Ctrl.Text, 1, 6) & Mid(Format(Date, "yyyy"), 1, 1) & Mid(Ctrl.Text, 7)
    If Len(Ctrl.Text) = 6 Then Ctrl.Text = Ctrl.Text & Format(Date, "yyyy")
    If Len(Ctrl) = 5 Then Ctrl.Text = Ctrl.Text & "/" & Format(Date, "yyyy")
    If Ctrl.Text <> "" And IsDate(Ctrl.Text) Then Ctrl.Text = CDate(Ctrl.Text)
    If Ctrl.Text <> "" And Not IsDate(Ctrl.Text) Then
        MsgBox "Introduzca una fecha válida"
        Ctrl.Text = ""
        Ctrl.SetFocus
        ValidarFecha = True
        Exit Function
    End If
End Function
