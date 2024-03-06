Attribute VB_Name = "ModFunciones"
Option Explicit



'190706 nombremaquina
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" ( _
    ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long) As Long

'190818 GetFileSpec
Public gcdg As CommonDialog



'190815 variables para los usuarios y maquinas
Global StrMaqAlta As String
Global StrMaqMod As String
Global StrUsAlta As String
Global StrUsMod As String

'190806 añadido A_PASSWORD
Global Contraseña As String
'240228 Variable strContrasenya
Global strContrasenya As String

'190718 blnNoValidar
Global BlnNoValidar As Boolean

'190706 nombremaquina
Global gStrNombreMaquina As String
  
  
'190706 TablaCambios
'[tabla] [varchar] (50)
'[Campo] [varchar] (50)
'[Registro] [varchar] (50)
'[FECHA] [datetime] NOT NULL ,
'[USUARIO] [varchar] (50)
'[Maquina] [varchar] (50)
'[Anterior] [varchar] (100)
'[Actual] [varchar] (100)
Global GstrTCtabla As String
Global GstrTCcampo As String
Global GstrTCRegistro As String
Global GstrTCFECHA As Date
Global GstrTCUSUARIO As String
Global GstrTCMaquina As String
Global GstrTCAnterior As String
Global GstrTCActual As String

  
  
Global gImpresoraDefecto As String
'190607 gNombreImpresora
Global gNombreImpresora As String

'190724 gstrFormato
Global gstrFormato As String
Global gblnVistaPrevia As Boolean
Global gNombreImpresoraContinuo As String

'190605 gblnCargandoDatos
Global gblnCargandoDatos As Boolean

Global LngAlturaMaxima As Long
Global G_Prn As Printer

'190125 FechaINI
Global dtmFechaIni As Date
Global dtmFechaFin As Date


'180730 strImpresoraCrystal
Global strImpresoraCrystal As String


'160630 Global gblnFacturaInversa as Boolean
Global gblnFacturaInversa As Boolean

'25/07/14 Añadido enumTipoPago
Enum enumTipoPago
    pNinguno = 0
    pCaja = 1
    pBanco = 2
    pTalon = 3
End Enum

'25/07/14 Añadido TAccion
Public Enum TAccion
     Ninguna = 0
     Alta = 1
     Modificacion = 2
     Consultando = 3
End Enum

'05/08/14 Añadida la variable gVirtual
Global gVirtual As Boolean

'25/07/14 Añadido gPagadoPor
Global gPagadoPor As enumTipoPago
'25/07/14 Añadido Accion
Global Accion As TAccion

'141230 añadida la variable bolCambioFormaPago
Global bolCambioFormaPago As Boolean

'161030 gStrCad
Global gStrCad As String

'25/05/14 Añadido el modulo ModFunciones(este modulo esta en la carpeta ModuloConexion)

'29/05/14 Variables para formateo decimales
'sin decimales = "#,##0"
'1 decimal = "#,##0.0"
'2 decimales = "#,##0.00"
'3 decimales = "#,##0.000"
'4 decimales = "#,##0.0000"

Global vValM As String 'formato moneda

'161010 variables vValMC y valMV para decimales en compras y ventas
Global vValMC As String 'formato moneda
Global vValMV As String 'formato moneda


Global vValP As String 'formato porcentaje
Global vValC As String 'formato cantidades
Global vValT As String 'formato totales
Global vValE As String 'formato enteros
'230116 añadido vValI para los decimales de impuestos
Global vValI As String 'formato impuestos



Global vValPesetas As String 'formato pesetas
'29/05/14 Variables para formateo decimales fin

'29/05/14 añadida variable SignoDecimal
Global SignoDecimal As String
Public Function RutaEscritorio() As String
    'Variable para usar WSH
    Dim Wscript As Object
    Dim cad As String
    
    'Creamos la referencia para usar Windows Scripting Host
    Set Wscript = CreateObject("WScript.Shell")
    cad = Wscript.SpecialFolders("Desktop")
    RutaEscritorio = cad & "\"
    If Not Wscript Is Nothing Then
       Set Wscript = Nothing
    End If
End Function

Public Function GetFileSpec(ByVal inExt As String) As String
    On Error GoTo errHandler
    GetFileSpec = ""
    With gcdg
         .DialogTitle = ""
         .DefaultExt = inExt
         .InitDir = RutaEscritorio
         .FileName = ""
         .Filter = "(*." & inExt & ")|*." & inExt & "|(*.*)|*.*"
         .FilterIndex = 0
         .CancelError = True
         .Flags = FileOpenConstants.cdlOFNHideReadOnly
         .ShowOpen
    End With
    GetFileSpec = gcdg.FileName
    Exit Function
errHandler:
    If Not Err = 32755 Then
         Msg mError, "", Err
         'MsgBox Err.Number & " " & Err.Description
    End If
End Function

Public Sub AbrirCarpeta(fichero As String)
    Dim ruta_matriz() As String
    Dim Ruta As String
    Dim Carpeta As String
    Dim fs As FileSystemObject
    Dim I As Integer
    'para el fichero fin
    'para el fichero
    Set fs = New FileSystemObject
    'asigno valores a la matriz ruta_matriz
    ruta_matriz = Split(fichero, "\")
    
    'aqui si es red
    Dim blnRed As Boolean
    Dim inicio As Integer
    Dim fichero1 As String
    inicio = 0
    
    If Mid(fichero, 1, 2) = "\\" Then blnRed = True
    
    If blnRed Then
        fichero1 = Mid(fichero, 3)
        ruta_matriz = Split(fichero1, "\")
        Ruta = "\\" & ruta_matriz(0) & "\" & ruta_matriz(1) & "\"
        inicio = 2
    End If
    'aqui si es red fin
    
    For I = inicio To UBound(ruta_matriz) - 1
        'recorro la matriz
        Carpeta = ruta_matriz(I) 'obtengo cada una de las careptas de la ruta
        If Carpeta <> "" And Right$(Carpeta, 1) <> ":" Then
            'si la variable carpeta no está vacía y no es la letra de la unidad
            If Dir(Ruta, vbDirectory) = "" Then
                'si no está creada la creo
                'mkdir (ruta)
                fs.CreateFolder (Ruta)
            End If
        End If
        Ruta = Ruta & Carpeta & "\"
    Next I
    Dim externa As String
    externa = Shell("explorer " & Ruta, vbNormalFocus)
End Sub

'170221 CrearCarpetas
Public Sub CrearCarpetas(fichero As String)
    'para el fichero
    Dim ruta_matriz() As String
    Dim Ruta As String
    Dim Carpeta As String
    Dim fs As FileSystemObject
    Dim I As Integer
    'para el fichero fin
    'para el fichero
    Set fs = New FileSystemObject
    'asigno valores a la matriz ruta_matriz
    
    ruta_matriz = Split(fichero, "\")
    Dim blnRed As Boolean
    Dim inicio As Integer
    
    
    Dim fichero1 As String
    
    
    inicio = 0
    
    If Mid(fichero, 1, 2) = "\\" Then blnRed = True
    
    If blnRed Then
        fichero1 = Mid(fichero, 3)
        ruta_matriz = Split(fichero1, "\")
        Ruta = "\\" & ruta_matriz(0) & "\" & ruta_matriz(1) & "\"
        inicio = 2
    End If
    
    '\\SERVER\datos contables v10\documentos\NEMP0002\Hacienda
    For I = inicio To UBound(ruta_matriz)
        'recorro la matriz
        Carpeta = ruta_matriz(I) 'obtengo cada una de las careptas de la ruta
        If Carpeta <> "" And Right$(Carpeta, 1) <> ":" Then
            'si la variable carpeta no está vacía y no es la letra de la unidad
            If Dir(Ruta, vbDirectory) = "" Then
                'si no está creada la creo
                'mkdir (ruta)
                fs.CreateFolder (Ruta)
            End If
        End If
        Ruta = Ruta & Carpeta & "\"
    Next I
    'para el fichero fin

End Sub


'141031 añadida la funcion DN (doble/nulo)
Public Function DN(IMPORTE As Variant) As Double
     If Not IsNumeric(IMPORTE) Then IMPORTE = 0#
     If IsNull(IMPORTE) Then IMPORTE = 0#
     DN = CDbl(IMPORTE)
End Function

Sub CopiarPorta(strCad As String)
    ''25/05/14 Añadido Sub CopiarPorta (para copiar un texto en el portapapeles)
    Clipboard.Clear
    Clipboard.SetText strCad
End Sub

'28/05/14 añadida funcion QueSigno (para obtener el signo decimal del sistema)
Public Sub QueSigno() 'Obtiene el separador decimal si es punto o coma
    Dim A As Double
    A = 1.1
    SignoDecimal = Mid(CStr(A), 2, 1)
End Sub

'29/05/14 añadido sub ValidaTXT (de momento para los campos de moneda)
Public Sub ValidaTXT(Txt As Control, KeyAscii As Integer, Optional Formato As String = "#,##0.0000")
    'en el KeyPress del txt
    'ValidaTXT txtBaseIVA, KeyAscii
    If KeyAscii = 8 Then Exit Sub
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Len(Txt) > 0 Then
            Txt.Text = Format(CCur(Txt.Text), Formato)
        End If
        MySendKeys "{TAB}", True
        Exit Sub
    End If
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
            If Len(Txt) > 0 Then
                If Mid(Txt, Txt.SelStart + 1, 1) <> "-" Then Exit Sub
            Else
                Exit Sub
            End If
        End If
    If SignoDecimal = "," Then
        If KeyAscii = Asc(".") Then KeyAscii = Asc(",") ':  Exit Sub
    End If
    If KeyAscii = Asc("-") Then
        If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
        If Txt.SelStart = 0 And InStr(1, Txt, "-") = 0 Then Exit Sub
    End If
    
    If SignoDecimal = "," Then
        If KeyAscii = Asc(",") Then
            If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
            If InStr(1, Txt, ",") = 0 Then
                If Len(Txt) > 0 Then
                    If Mid(Txt, Txt.SelStart + 1, 1) <> "-" Then Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    Else
        If KeyAscii = Asc(".") Then
            If Txt.SelStart = 0 And Txt.SelLength = Len(Txt) Then Exit Sub
            If InStr(1, Txt, ",") = 0 Then
                If Len(Txt) > 0 Then
                    If Mid(Txt, Txt.SelStart + 1, 1) <> "-" Then Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End If
    
    KeyAscii = 0
    
End Sub
Public Function DD(IMPORTE As Variant) As Double
     '29/05/14 Añadida la funcion DD (DosDecimales)
     If Not IsNumeric(IMPORTE) Then IMPORTE = 0#
     If IsNull(IMPORTE) Then IMPORTE = 0#
     DD = Format(CDbl(IMPORTE), vValT)
End Function

Public Function DDCur(IMPORTE As Variant) As Currency
     '12/06/14 añadida la funcion DDCur (dos decimales para datos currency)
     If Not IsNumeric(IMPORTE) Then IMPORTE = 0#
     If IsNull(IMPORTE) Then IMPORTE = 0#
     DDCur = Format(CCur(IMPORTE), vValT)
End Function
Public Function CCurVentas(IMPORTE As Variant) As Currency
     'REVISAR
     '190109 DECIMALES EN LINEAS DE FACTURAS
     If Not IsNumeric(IMPORTE) Then IMPORTE = 0#
     If IsNull(IMPORTE) Then IMPORTE = 0#
     CCurVentas = Format(CCur(IMPORTE), "#,##0.0000")
End Function
Public Sub ValidaCaracter(ctrl As Control, KeyAscii As Integer, Valido As String, Optional largo As Integer = 1)
    '161010 ValidaCaracter
    Dim Mayus As String
    Dim Minus As String
    'si es la tecla borrar
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then Exit Sub
    If KeyAscii = 32 Then ctrl.Text = "": Exit Sub
    If Len(Trim(ctrl.Text)) = largo Then KeyAscii = 0: Exit Sub
    'If Len(Trim(Ctrl.Text)) = largo Then Exit Sub
    Mayus = StrConv(Valido, vbUpperCase)
    Minus = StrConv(Valido, vbLowerCase)
        
    If Chr(KeyAscii) = Mayus Or Chr(KeyAscii) = Minus Then
        ctrl.Text = Mayus
        KeyAscii = Asc(Mayus)
    Else
        KeyAscii = 0
    End If
End Sub

Public Function ValidaCaracteres(ctrl As Control, KeyAscii As Integer, Valido As String, Optional largo As Integer = 1) As Integer

    '161030 ValidaCaracteres
    Dim Mayus As String
    Dim Minus As String
    Dim caracter As String
    
    On Error GoTo Fallo
    ValidaCaracteres = KeyAscii
    'si es la tecla borrar
    If KeyAscii = 8 Then Exit Function
    If KeyAscii = 13 Then Exit Function
    If KeyAscii = 32 Then Exit Function
    
    '170117 largo en ValidaCaracteres
    'If Len(Trim(Ctrl.Text)) = largo Then ValidaCaracteres = 0: Exit Function
    If Len(Trim(ctrl.Text)) > largo Then ValidaCaracteres = 0: Exit Function
    'If Len(Trim(Ctrl.Text)) = largo Then Exit Sub
    
    caracter = UCase(Chr(KeyAscii))
    
    If InStr(UCase(Valido), caracter) = 0 Then GoTo Terminar
    KeyAscii = Asc(caracter)
    ValidaCaracteres = Asc(caracter)
    Exit Function
    'Mayus = StrConv(Valido, vbUpperCase)
    'Minus = StrConv(Valido, vbLowerCase)
        
    'If Chr(KeyAscii) = Mayus Or Chr(KeyAscii) = Minus Then
    '    Ctrl.Text = Mayus
    '    KeyAscii = Asc(Mayus)
    'Else
    '    KeyAscii = 0
    'End If
    On Error GoTo 0
Fallo:
    'MSG mError, "", Err, Erl
    MsgBox Err.Number & " " & Err.Description
    Exit Function
Terminar:
   KeyAscii = 0
   ValidaCaracteres = 0
End Function
Public Function NumerosPositivos(pulsado As Integer) As Integer
    Dim caracter As Integer
    
    NumerosPositivos = pulsado
    'si es la tecla borrar
    If NumerosPositivos = 8 Then Exit Function
    'si el caracter es de 0 a 9
    If NumerosPositivos >= Asc("0") And NumerosPositivos <= Asc("9") Then Exit Function
    If NumerosPositivos = 13 Then Exit Function
    'If NumerosPositivos = 46 Then Exit Function
    If SignoDecimal = "," Then
        If NumerosPositivos = 46 Then
            NumerosPositivos = 44
            Exit Function
        ElseIf NumerosPositivos = 44 Then
            Exit Function
        End If
    Else
        If NumerosPositivos = 44 Then
            NumerosPositivos = 46
            Exit Function
        ElseIf NumerosPositivos = 46 Then
            Exit Function
        End If
    End If
    
    NumerosPositivos = 0
End Function


Public Function SoloNumeros(pulsado As Integer) As Integer
    '170810 añadida la funcion SoloNumeros
    SoloNumeros = pulsado
    'si es la tecla borrar
    If SoloNumeros = 8 Then Exit Function
    'si el caracter es de 0 a 9
    If SoloNumeros >= Asc("0") And SoloNumeros <= Asc("9") Then Exit Function
    If SoloNumeros = 13 Then Exit Function
'        If Ctrl.SelStart = 0 And Ctrl.SelLength = Len(Ctrl) Then Exit Sub
'        If Len(Ctrl) > 0 Then
'            If Mid(Ctrl, Ctrl.SelStart + 1, 1) <> "-" Then Exit Sub
'        Else
'            Exit Sub
'        End If
'        Exit Sub
'    End If
    SoloNumeros = 0
End Function


Public Sub ValidaNumero(ctrl As Control, KeyAscii As Integer, Optional Negativos As Boolean = True)
    '25/06/14 añadido el sub ValidaNumero
    'si es la tecla borrar
    If KeyAscii = 8 Then Exit Sub
    'si el caracter es de 0 a 9
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        If ctrl.SelStart = 0 And ctrl.SelLength = Len(ctrl) Then Exit Sub
        If Len(ctrl) > 0 Then
            If Mid(ctrl, ctrl.SelStart + 1, 1) <> "-" Then Exit Sub
        Else
            Exit Sub
        End If
        Exit Sub
    End If
    'return
    If KeyAscii = 13 Then Exit Sub
    
    If SignoDecimal = "," Then
        If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
    End If
    
    If KeyAscii = Asc("-") Then
        If Negativos Then
            If ctrl.SelStart = 0 And ctrl.SelLength = Len(ctrl.Text) Then Exit Sub
            If ctrl.SelStart = 0 And InStr(1, ctrl.Text, "-") = 0 Then Exit Sub
        End If
    End If
    
    If SignoDecimal = "," Then
        If KeyAscii = Asc(",") Then
            If ctrl.SelStart = 0 And ctrl.SelLength = Len(ctrl) Then Exit Sub
            If InStr(1, ctrl, ",") = 0 Then
                If Len(ctrl) > 0 Then
                    If Mid(ctrl, ctrl.SelStart + 1, 1) <> "-" Then Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    Else
        If KeyAscii = Asc(".") Then
            If ctrl.SelStart = 0 And ctrl.SelLength = Len(ctrl.Text) Then Exit Sub
            
            If InStr(1, ctrl.Text, ".") = 0 Then
                If Len(ctrl.Text) > 0 Then
                    Debug.Print Mid(ctrl.Text, ctrl.SelStart + 1, 1)
                    If Mid(ctrl.Text, ctrl.SelStart + 1, 1) <> "-" Then Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End If
    
    KeyAscii = 0
End Sub

Public Function strAsignaPagadoPor() As String
    '25/07/14 Function strAsignaPagadoPor
    Select Case gPagadoPor
        Case 0
            strAsignaPagadoPor = " "
        Case 1
            strAsignaPagadoPor = "C"
        Case 2
            strAsignaPagadoPor = "B"
        Case 3
            strAsignaPagadoPor = "T"
    End Select
    'frmTalones.txtFields(11).Text = strAsignaPagadoPor

End Function

Public Function ArchivoExistente(archivo As String) As Boolean
   Dim X
   On Error Resume Next
   X = GetAttr(archivo)
   If Err Then ArchivoExistente = False Else ArchivoExistente = True
   On Error GoTo 0
End Function
Public Sub CentrarFormulario(frm, Optional frmMDI As MDIForm)
'     If Not frmMDI Is Nothing Then
'        fRm.Move frmMDI.Left + (frmMDI.Width - fRm.Width) / 2, frmMDI.Top + (frmMDI.Height - fRm.Height) / 2
'     Else
'        fRm.Move (Screen.Width - fRm.Width) / 2, (Screen.Height - fRm.Height) / 2
'     End If
    
    '230117 Modificaca la funcion CentrarFormulario
    Dim vIz As Long
    Dim vAr As Long
    
    Dim vPosIz As Long
    Dim vPosAr As Long
    'mdichild
    
    If frmMDI.WindowState = 2 Then
        vIz = frmMDI.Left
        vAr = frmMDI.Top
        If frm.MDIChild Then
            vIz = 0
            vAr = 0
        End If
        vPosIz = ((frmMDI.ScaleWidth - frm.Width) / 2) + vIz
        vPosAr = ((frmMDI.ScaleHeight - frm.Height) / 2)
        'vPosAr = vPosAr + vAr + 720 + 255
        frm.Move vPosIz, vPosAr
    Else
        'Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
        vIz = 0
        vAr = 0
        
        'vIz = frmMDI.Left
        'vAr = frmMDI.Top
        
        'vPosIz = ((Screen.Width / 2) - (Frm.Width / 2)) + vIz
        'vPosAr = ((Screen.Height / 2) - (Frm.Height / 2))
        
        'vPosIz = ((Screen.Width - Frm.Width) / 2) + vIz
        'vPosAr = ((Screen.Height - Frm.Height) / 2)
        
        vPosIz = ((frmMDI.ScaleWidth - frm.Width) / 2) + vIz
        vPosAr = ((frmMDI.ScaleHeight - frm.Height) / 2)
        
        'vPosIz = ((frmMDI.Width - fRm.Width) / 2) + vIz
        'vPosAr = ((frmMDI.Height - fRm.Height) / 2)
        
        
        'vPosAr = vPosAr + vAr + 720 + 255
        'Frm.Move (Screen.Width - Screen.Height) \ 2
        'Frm.Move (vPosIz - vPosAr) \ 2
        
        'fRm.Move vPosIz, vPosAr
        frm.Left = vPosIz
        frm.Top = vPosAr
    End If

End Sub

Public Function CuantosParametros(strCadena As String) As Integer
    Dim strCad As String
    Dim intParametros As Integer
    Dim A As Integer
    Dim c As String
    
    intParametros = 0
    For A = 1 To Len(strCadena)
        c = Mid(strCadena, A, 1)
        If c = ";" Then
            intParametros = intParametros + 1
        End If
    Next
    If Len(strCadena) > 0 Then intParametros = intParametros + 1
    CuantosParametros = intParametros
End Function

Function ObtenerMes(periodo As Integer) As String
    '160403 añadida la funcion ObtenerMes en ModFunciones
    
    'el IdEti.Recuperar( es para los idiomas cuando los haya
    '02/04/11 Añadida la funcion ObtenerMes para mostrar el mes en los combos
    Dim strT As String
    Select Case periodo
        Case 0 ' Anteriores
            'strT = IdEti.Recuperar("Anteriores")
            strT = "Anteriores"
        Case 1 'Enero
            'strT = IdEti.Recuperar("Enero")
            strT = "Enero"
        Case 2 'Febero
            'strT = IdEti.Recuperar("Febrero")
            strT = "Febero"
        Case 3 'Marzo
            'strT = IdEti.Recuperar("Marzo")
            strT = "Marzo"
        Case 4 'Abril
            'strT = IdEti.Recuperar("Abril")
            strT = "Abril"
        Case 5 'Mayo
            'strT = IdEti.Recuperar("Mayo")
            strT = "Mayo"
        Case 6 'Junio
            'strT = IdEti.Recuperar("Junio")
            strT = "Junio"
        Case 7 'Julio
            'strT = IdEti.Recuperar("Julio")
            strT = "Julio"
        Case 8 'Agosto
            'strT = IdEti.Recuperar("Agosto")
            strT = "Agosto"
        Case 9 'Septiembre
            'strT = IdEti.Recuperar("Septiembre")
            strT = "Septiembre"
        Case 10 'Octubre
            'strT = IdEti.Recuperar("Octubre")
            strT = "Octubre"
        Case 11 'Noviembre
            'strT = IdEti.Recuperar("Noviembre")
            strT = "Noviembre"
        Case 12 'Diciembre
            'strT = IdEti.Recuperar("Diciembre")
            strT = "Diciembre"
    End Select
    ObtenerMes = strT
End Function
Public Function FormularioCargado(frm As String) As Boolean
    '160403 añadida la funcion FormularioCargado
    Dim intI As Integer
    For intI = 0 To Forms.Count - 1
        If Forms(intI).Name = frm Then
            FormularioCargado = True
            Exit For
        End If
    Next
End Function


Public Function NumText(Numero As Double) As String 'numero a texto
    Dim cad As String
    If Numero <> 0 Then
        Numero = Format(Numero, "#,##0.00")
    End If
    If Numero = 0 Then
        cad = "0.00"
    ElseIf Int(Numero) = 0 Then
        cad = cad
        cad = "0." & Format(Trim(str(Numero * 100)), "00")
    Else
        cad = Trim(str(Numero * 100))
        cad = Mid(cad, 1, Len(cad) - 2) & "." & Right(cad, 2)
    End If
    NumText = cad
End Function

Function Fdato(cad As String, longitud As Integer, Optional Relleno As String = "", Optional Derecha As Boolean = False) As String
    Dim PonerNegativo As Boolean
    'PonerNegativo = " "
'    If Mid(cad, 1, 1) = "?" Then
'        cad = Buscar(cad)
'    End If
    If Mid(cad, 1, 1) = "-" Then
        If IsNumeric(Mid(cad, 2)) Then
            'PonerNegativo = True
            cad = Mid(cad, 2)
            longitud = longitud - 1
        End If
    End If
    If Relleno = "" Then Relleno = " "
    If cad = "" Then
        Fdato = String(longitud, Relleno)
    Else
        cad = Trim(cad)
        If Len(cad) < longitud Then
            If Derecha Then
                If PonerNegativo Then
                    Fdato = "N" & String(longitud - Len(cad), Relleno) + cad
                Else
                    Fdato = String(longitud - Len(cad), Relleno) + cad
                End If
            Else
                If PonerNegativo Then
                    Fdato = "N" & cad & String(longitud - Len(cad), Relleno)
                Else
                    Fdato = cad & String(longitud - Len(cad), Relleno)
                End If
            End If
        Else
            Fdato = Mid(cad, 1, longitud)
        End If
    End If
End Function

Function FechaCrystal(Fechaini As Date, FechaFin As Date) As String
    Dim cad As String
    cad = "in datetime ("
    cad = cad & Format(Fechaini, "yyyy") & ", "
    cad = cad & Format(Fechaini, "mm") & ", "
    cad = cad & Format(Fechaini, "dd") & ",00, 00, 00) to datetime ("
    cad = cad & Format(FechaFin, "yyyy") & ", "
    cad = cad & Format(FechaFin, "mm") & ", "
    cad = cad & Format(FechaFin, "dd") & ",23, 59, 59)"
    FechaCrystal = cad
End Function
'190706 nombremaquina

Public Function NombreMaquina() As String
   '150917 aladida la funcion NumeroMaquina
   'esto solo es para si falla el codigo generado en SCocx
   Dim nPC As String
   Dim buffer As String
   Dim estado As Long
   buffer = String$(255, " ")
   estado = GetComputerName(buffer, 255)
   If estado <> 0 Then
      nPC = Left(buffer, 255)
   Else
      nPC = ""
   End If
   NombreMaquina = Left(Trim(nPC), Len(Trim(nPC)) - 1)
End Function

Function impresora() As String
  
    Dim buffer As String
    Dim Ret As Integer
  
    buffer = Space(255)
  
    Ret = GetProfileString("Windows", ByVal "device", "", _
                                 buffer, Len(buffer))
  
    If Ret Then
        impresora = UCase(Left(buffer, _
                                   InStr(buffer, ",") - 1))
    End If
  
End Function

Sub LogearFichero(fichero As String, cad As String)
    Open fichero For Append As #1
    Print #1, cad
    Close #1
End Sub

Public Function DosPalabras(Frase As String) As String
    DosPalabras = ""
    Dim I As Integer
    Dim ii As Integer
    Dim cad As String
    Dim cad1 As String
    
    Dim blnHayPrimeraPalabra As Boolean
    
    Dim CadTmp As String
    
    CadTmp = Frase
    
    cad = ""
    cad1 = ""
    
    I = InStr(Frase, " ")
    
    If I <> 0 Then
        cad = Mid(CadTmp, 1, I)
        cad = Trim(cad)
        blnHayPrimeraPalabra = True
        
    End If
    
    If blnHayPrimeraPalabra Then
        CadTmp = Trim(Mid(CadTmp, I))
        ii = InStr(CadTmp, " ")
        If ii <> 0 Then
            cad1 = Trim(Mid(CadTmp, 1, ii))
        Else
            cad1 = Trim(Mid(CadTmp, 1))
        cad1 = Trim(cad1)
        End If
    End If
    
    DosPalabras = cad & " " & cad1
End Function
Function HacerPausa()
   Dim TiempoPausa, inicio, Final, TiempoTotal
   TiempoPausa = 2   ' Asigna hora de inicio.
   inicio = Timer   ' Establece la hora de inicio.
   Do While Timer < inicio + TiempoPausa
      DoEvents   ' Cambia a otros procesos.
   Loop
   Final = Time   ' Asigna hora de finalización.
   TiempoTotal = Final - inicio   ' Calcula tiempo total.
End Function

