Attribute VB_Name = "mdlCifrado"

Function Cifrado(Cadena As String, Optional Cifrar As Boolean = False) As String
    Dim Vig(11) As Integer
    Dim N As Integer, p As Integer, Q As Integer, incr As Integer
    Dim Car As String
    Dim CadTmp As String

    If Cadena = "" Then Cifrado = "": Exit Function
    
    'dependiendo del valor que le des tu a vig() generara una cadena u otra
    
    'Constantes Vigenère:
    '3141592653531415926535314159265353141592653
    Vig(1) = 3: Vig(2) = 1: Vig(3) = 4: Vig(4) = 1: Vig(5) = 5
    Vig(6) = 9: Vig(7) = 2: Vig(8) = 6: Vig(9) = 5: Vig(10) = 3: Vig(11) = 5
    
    'Descifrar cadena.
    Q = 1: CadTmp = ""
    For p = 1 To Len(Cadena)
        Car = Mid(Cadena, p, 1)
        If Cifrar Then
            incr = Asc(Car) + Vig(Q)
        Else
            incr = Asc(Car) - Vig(Q)
        End If
        CadTmp = CadTmp + Chr(incr)
        Q = Q + 1
        If Q = 12 Then Q = 1
    Next
    
    Cifrado = CadTmp
End Function

