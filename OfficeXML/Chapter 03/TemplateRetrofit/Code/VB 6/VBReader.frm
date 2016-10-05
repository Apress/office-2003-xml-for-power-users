VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "XML Reader"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.ListBox lstItems 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   280
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    ' Create the DOMDocument object.
    Dim Doc As DOMDocument
    Set Doc = New DOMDocument
    
    ' Load the ExpenseData.xml file into memory.
    If Doc.Load(App.Path & "\ExpenseData.xml") Then
    
        ' Scan all the nodes under the root element.
        Dim Node As IXMLDOMNode
        For Each Node In Doc.documentElement.childNodes
            
            ' Check if the node contains the metadata (header) section.
            If Node.nodeName = "Meta" Then
            
                ' Retrieve information about the expense report.
                txtName.Text = Node.childNodes(0).Text
                txtID.Text = Node.childNodes(2).Text
                
            ' Check if the node contains the expense list.
            ElseIf Node.nodeName = "ExpenseItem" Then
                 
                 ' Get the description and total of each expense item.
                 Dim Description As String
                 Description = Node.childNodes(1).Text
                 Dim Total As String
                 Total = FormatCurrency(Node.childNodes(10).Text)
                 
                 ' Show this information in a listbox.
                 lstItems.AddItem (Description & " (" & Total & ")")
            End If
        Next
    Else
        MsgBox ("File ExpenseData.xml not found in this directory.")
    End If
    
End Sub

