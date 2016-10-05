VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memo Generator"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
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
   ScaleHeight     =   4830
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   3840
      Width           =   4095
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtBody 
      Height          =   1845
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "CreateOfficeDoc.frx":0000
      Top             =   1800
      Width           =   5415
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "New Memo Generating Software"
      Top             =   1030
      Width           =   3375
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "Lisa Stitzwilliam"
      Top             =   570
      Width           =   3375
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "John Diagamore"
      Top             =   210
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Generated File: "
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3885
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "To:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtFileName.Text = App.Path & "\NewMemo.xml"
    
End Sub


Private Sub cmdGenerate_Click()
    
    ' Create the DOMDocument object for accessing the XML.
    Dim Doc As DOMDocument
    Set Doc = New DOMDocument
       
    ' Load the selected file into memory.
    If Doc.Load(App.Path & "\MemoTemplate.xml") Then
     
        Call Doc.setProperty("SelectionNamespaces", _
          "xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'")
          
        Dim TextNodes As IXMLDOMNodeList
        Set TextNodes = Doc.documentElement.selectNodes("//w:t")

        Dim TextNode As IXMLDOMNode
        For Each TextNode In TextNodes
        Select Case TextNode.Text
            Case "CompanyName"
                TextNode.Text = "ProseTech"
            Case "ToName"
                TextNode.Text = txtTo.Text
            Case "FromName"
                TextNode.Text = txtFrom.Text
            Case "DateInfo"
                TextNode.Text = Now
            Case "SubjectLine"
                TextNode.Text = txtSubject.Text
            Case "MemoBody"
                TextNode.Text = txtBody.Text
        End Select
        
        Next
    
        Doc.save txtFileName.Text
        MsgBox "File saved."
    Else
        MsgBox ("MemoTemplate.xml file is missing.")
    End If

End Sub
