Attribute VB_Name = "mdlValidarEmail"
Option Explicit

'Función que comprueba si una diección de email es válida
'*********************************************************************************
Public Function Validar_Email(ByVal Email As String) As Boolean
    
    Dim i As Integer, iLen As Integer, caracter As String
    Dim pos As Integer, bp As Boolean, iPos As Integer, iPos2 As Integer

    On Local Error GoTo Err_Sub
    If Len(Trim(Email)) = 0 Then Validar_Email = True: Exit Function
    Email = Trim$(Email)

    If Email = vbNullString Then
        Exit Function
    End If

    Email = LCase$(Email)
    iLen = Len(Email)

    
    For i = 1 To iLen
        caracter = Mid(Email, i, 1)

        If (Not (caracter Like "[a-z]")) And (Not (caracter Like "[0-9]")) Then
            
            If InStr(1, "_-" & "." & "@", caracter) > 0 Then
                If bp = True Then
                   Exit Function
                Else
                    bp = True
                   
                    If i = 1 Or i = iLen Then
                        Exit Function
                    End If
                    
                    If caracter = "@" Then
                        If iPos = 0 Then
                            iPos = i
                        Else
                            
                            Exit Function
                        End If
                    End If
                    If caracter = "." Then
                        iPos2 = i
                    End If
                    
                End If
            Else
                
                Exit Function
            End If
        Else
            bp = False
        End If
    Next i
    If iPos = 0 Or iPos2 = 0 Then
        Exit Function
    End If
    
    If iPos2 < iPos Then
        Exit Function
    End If

    
    Validar_Email = True

    Exit Function
Err_Sub:
    On Local Error Resume Next
    
    Validar_Email = False
End Function

