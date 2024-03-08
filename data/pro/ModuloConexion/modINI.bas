Attribute VB_Name = "modINI"
Option Explicit
'27/07/14 Añadido modulo modINI

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public gstrCadenaAComprobar As String
Public MiSql As String
Public T As String
Public gstrCatTmp As String
Public gstrFicheroLog As String

Public Enum FormatosCadenas
    Normal1 = 0
    ForzarMinusculas = 1 'vbLowerCase
    ForzarMayusculas = 2 'vbUpperCase
    PrimeraMayuscula = 3 'vbProperCase
End Enum
Public Enum Formatos 'Tipos de datos
    FCadena = 0 'Todos los caracteres
    FCodigo = 1 'Formatea con ceros a la izquierda
    FFecha = 2 'Fecha - No hace falta poner el separador / lo pone automaticamente
    FHora = 3 'Horas - Tampoco hay que poner el separador
    FSubcuenta = 4 'Datos de tipo subcuenta. Se puede utilizar el punto
    FCantidad = 5 'Cantidades - Utiliza la mascara para el formateo de decimales
    FMoneda = 6 'Modena - Utiliza la mascara para el formateo de decimales
    FPorcentaje = 7 'Porcentajes - Utiliza la mascara para el formateo de decimales
    FNumerico = 8 'Datos numéricos, sin decimales
    FTotales = 9 'Siempre formatea a dos decimales
    FChequeo = 10 'Actua como un chek box
    FOpcion = 11 'Actua como un option buton
    FCombo = 12 'Sin utilidad, no esta terminado, hay un control aparte
    FCodCli = 13 'Codigos de clientes/proveedores, la longitud sera la de longitudcuentas -4
    FCodObra = 14 'Sin uso
    FTelefono = 15 'Solo admite numeros
End Enum
Public Enum TiposAlineacion 'La alineacion de la etiqueta
    TAIzquierda = 0 'izquierda
    TADerecha = 1  'derecja
    TACentro = 2 'centro
End Enum

Public Enum TiposPosicion 'La posicion de la etiqueta
    TPIzquierda = 0 'izquiderda
    TPDerecha = 1 'derecha
    TPArriba = 2 'arriba
End Enum

Function ObtenerCampo(Campo As Integer, tCadena As Variant, Separador As String) As String
   Dim Cadena As String
   Dim cad As String
   Dim I As Integer
   Dim j As Integer
   Dim ini As Integer
   Dim Fin As Integer
   Cadena = tCadena
   'Restricciones:
   If Campo = 0 Or Cadena = "" Or Separador = "" Then ObtenerCampo = "": Exit Function
   If Len(Separador) > 1 Then ObtenerCampo = "": Exit Function
   Dim CadTmp() As String
   CadTmp = Split(tCadena, Separador)
   If Campo - 1 > UBound(CadTmp) Then ObtenerCampo = "": Exit Function
   ObtenerCampo = CadTmp(Campo - 1)
End Function
Sub EscribirIni(Fichero As String, Seccion As String, Variable As String, valor As String)
   Dim I As Integer
   Dim strCad As String
   I = WritePrivateProfileString(Seccion, Variable, valor, Fichero)
   If I = 0 Then
      strCad = "Error accediendo al fichero"
      Msg mError, strCad
   End If
End Sub


Function LeerINI(Fichero As String, Seccion As String, Variable As String) As String
   Dim I As Integer
   Dim cad As String * 255
   Dim Dato As String
   I = GetPrivateProfileString(Seccion, Variable, "", cad, Len(cad), Fichero)
   If I = 0 Then
      LeerINI = "-88888888"
      Exit Function
   Else
      Dato = ""
      For I = 1 To 255
         If Asc(Mid(cad, I, 1)) = 0 Then Exit For
         Dato = Dato + Mid(cad, I, 1)
      Next
      LeerINI = Trim(Dato)
   End If
End Function

