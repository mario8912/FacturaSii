Attribute VB_Name = "mdlMensajes"
Option Explicit

Enum enumTipoMsg
    mError = 0
    mAdvertencia = 1
    mPregunta = 2
    mInformacion = 3
End Enum

Function Msg(Tipo As enumTipoMsg, Mensaje As String, Optional ObjetoError As ErrObject, Optional LineaError As Integer = 0, Optional Origen As String = "", Optional Proceso As String = "") As Boolean
    'Tipo a mostrar
    'Mensaje a mostrar
    'ObjetoError , el error producido
    'LineaError , la linea del error (si procede)
    'Origen , el frm, modulo, etc que ha producido el error (si procede)
    'Proceso , el proceso que ha producido el error (si procede)
    Dim strCad As String
    
    
    If ObjetoError Is Nothing Then
        strCad = Mensaje
    Else
        If Mensaje <> "" Then
            strCad = Mensaje + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "[DETALLES: " + "(" + Trim(str(ObjetoError.Number)) + ") " + ObjetoError.Description + "]"
        Else
            strCad = "[DETALLES: " + "(" + Trim(str(ObjetoError.Number)) + ") " + ObjetoError.Description + "]"
        End If
    End If
    If Origen <> "" Then
        strCad = strCad & Chr(13) & Chr(10) & "Origen : " & Origen
    End If
    If Proceso <> "" Then
        strCad = strCad & Chr(13) & Chr(10) & "Proceso : " & Proceso
    End If
    If LineaError <> 0 Then
        strCad = strCad & Chr(13) & Chr(10) & "Linea : " & LineaError
    End If
    
    'dibujar el frm
    frmMsg.Tipo = Tipo
    frmMsg.Mensaje = strCad
    
    frmMsg.Line2.Y1 = frmMsg.Mensaje.Top + frmMsg.Mensaje.Height + 150
    frmMsg.Line2.Y2 = frmMsg.Mensaje.Top + frmMsg.Mensaje.Height + 150
    frmMsg.cmdAceptar.Top = frmMsg.Line2.Y1 + 120
    frmMsg.cmdSi.Top = frmMsg.Line2.Y1 + 120
    frmMsg.cmdNo.Top = frmMsg.Line2.Y1 + 120
    
    frmMsg.Shape1.Height = frmMsg.cmdAceptar.Top + 435
    frmMsg.Height = frmMsg.Shape1.Height + 510 + (frmMsg.Height - 2325)
    
    frmMsg.Width = frmMsg.Mensaje.Width + 480
    frmMsg.Shape1.Width = frmMsg.Mensaje.Width + 360
    frmMsg.Line1.X2 = frmMsg.Shape1.Left + frmMsg.Shape1.Width
    frmMsg.Line2.X2 = frmMsg.Shape1.Left + frmMsg.Shape1.Width
    
    If Tipo = mError Then
        Beep
    End If
    
    frmMsg.Show 1
    Msg = frmMsg.CONTESTACION
End Function
