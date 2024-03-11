VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10365
   ClientLeft      =   11235
   ClientTop       =   2700
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   9540
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8775
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   15478
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim XMLDoc As MSXML.DOMDocument

Dim XMLDoc As MSXML2.DOMDocument60
Private Sub Form_Load()
    Form1.Move 0, 0, 5565, 4495
    Command1.Move 120, 600, 1335, 375
    Text1.Move 1560, 120, 3735, 285
    Label1.Move 120, 120, 1575, 285
    TreeView1.Move 120, 1080 ', 5175, 2895
    Command1.Caption = "Leer"
    Label1.Caption = "Ruta del xml:"
    Text1.Text = App.Path & "\archivo.XML"
    Form1.Caption = "Leer xml"
End Sub

Private Sub Command1_Click()
    'Set XMLDoc = New DOMDocument
      Set XMLDoc = New MSXML2.DOMDocument60
    XMLDoc.async = False
    
    XMLDoc.Load Text1.Text
    
    If XMLDoc.parseError.errorCode = 0 Then
        MsgBox "Succeeded"
        If XMLDoc.readyState = 4 Then
            TreeView1.Nodes.Clear
            AddNode XMLDoc.documentElement
        End If
    Else
        MsgBox XMLDoc.parseError.reason & vbCrLf & _
        XMLDoc.parseError.Line & vbCrLf & _
        XMLDoc.parseError.srcText
    End If
    
End Sub

Private Sub AddNode(ByRef XML_Node As IXMLDOMNode, _
                    Optional ByRef TreeNode As Node)
    Dim xNode As Node
    Dim xNodeList As IXMLDOMNodeList
    Dim i As Long
    
    If TreeNode Is Nothing Then
        Set xNode = TreeView1.Nodes.Add
    Else
        Set xNode = TreeView1.Nodes.Add(TreeNode, tvwChild)
    End If
    
    xNode.Expanded = True
    xNode.Text = XML_Node.nodeName
    
    If xNode.Text = "#text" Then
        xNode.Text = XML_Node.nodeTypedValue
    Else
        xNode.Text = "<" + xNode.Text + ">"
    End If
    
    Set xNodeList = XML_Node.childNodes
    For i = 0 To xNodeList.length - 1
        AddNode xNodeList.Item(i), xNode
    Next
End Sub
                    

