Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	
	'Dim XMLDoc As MSXML.DOMDocument
	
	Dim XMLDoc As MSXML2.DOMDocument60
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.SetBounds(0, 0, VB6.TwipsToPixelsX(5565), VB6.TwipsToPixelsY(4495))
		Command1.SetBounds(VB6.TwipsToPixelsX(120), VB6.TwipsToPixelsY(600), VB6.TwipsToPixelsX(1335), VB6.TwipsToPixelsY(375))
		Text1.SetBounds(VB6.TwipsToPixelsX(1560), VB6.TwipsToPixelsY(120), VB6.TwipsToPixelsX(3735), VB6.TwipsToPixelsY(285))
		Label1.SetBounds(VB6.TwipsToPixelsX(120), VB6.TwipsToPixelsY(120), VB6.TwipsToPixelsX(1575), VB6.TwipsToPixelsY(285))
		TreeView1.SetBounds(VB6.TwipsToPixelsX(120), VB6.TwipsToPixelsY(1080), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y) ', 5175, 2895
		Command1.Text = "Leer"
		Label1.Text = "Ruta del xml:"
		Text1.Text = My.Application.Info.DirectoryPath & "\archivo.XML"
		Me.Text = "Leer xml"
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		'Set XMLDoc = New DOMDocument
		XMLDoc = New MSXML2.DOMDocument60
		XMLDoc.async = False
		
		XMLDoc.Load(Text1.Text)
		
		If XMLDoc.parseError.errorCode = 0 Then
			MsgBox("Succeeded")
			If XMLDoc.readyState = 4 Then
				TreeView1.Nodes.Clear()
				'UPGRADE_WARNING: No se puede resolver la propiedad predeterminada del objeto XMLDoc.documentElement. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AddNode((XMLDoc.documentElement))
			End If
		Else
			MsgBox(XMLDoc.parseError.reason & vbCrLf & XMLDoc.parseError.Line & vbCrLf & XMLDoc.parseError.srcText)
		End If
		
	End Sub
	
	Private Sub AddNode(ByRef XML_Node As MSXML2.IXMLDOMNode, Optional ByRef TreeNode As System.Windows.Forms.TreeNode = Nothing)
		Dim xNode As System.Windows.Forms.TreeNode
		Dim xNodeList As MSXML2.IXMLDOMNodeList
		Dim i As Integer
		
		If TreeNode Is Nothing Then
			'UPGRADE_ISSUE: MSComctlLib.Nodes método TreeView1.Nodes.Add no se actualizó. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            'xNode = TreeView1.Nodes.Add()
            xNode = TreeView1.Nodes.Add("AQ")
		Else
			xNode = TreeView1.Nodes.Find(TreeNode.Name, True)(0).Nodes.Add("")
		End If
		
		xNode.Expand()
		xNode.Text = XML_Node.nodeName
		
		If xNode.Text = "#text" Then
			xNode.Text = XML_Node.nodeTypedValue
		Else
			xNode.Text = "<" & xNode.Text & ">"
		End If
		
		xNodeList = XML_Node.childNodes
		For i = 0 To xNodeList.length - 1
			AddNode(xNodeList.Item(i), xNode)
		Next 
	End Sub
End Class