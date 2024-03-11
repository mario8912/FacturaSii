Imports System
Imports System.Web

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'aeatNif2.VNifV2EntContribuyente = "18924245Y"
        Dim aa As New aeatNif2.VNifV2EntContribuyente

        aa.Nif = "1892424Y"
        aa.Nombre = "manuel peris"

    End Sub
End Class
