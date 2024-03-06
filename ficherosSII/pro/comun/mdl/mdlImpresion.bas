Attribute VB_Name = "mdlImpresion"
Option Explicit
Public Function ObtenerDriverImpresora(impresora As String) As String
   Dim pr As Printer
   For Each pr In Printers
      If pr.DeviceName = impresora Then
         ObtenerDriverImpresora = pr.DriverName
         Exit Function
      End If
   Next
   ObtenerDriverImpresora = ""
End Function

Public Function ObtenerPuertoImpresora(impresora As String) As String
   Dim pr As Printer
   For Each pr In Printers
      If pr.DeviceName = impresora Then
         ObtenerPuertoImpresora = pr.Port
         Exit Function
      End If
   Next
   ObtenerPuertoImpresora = ""
End Function

