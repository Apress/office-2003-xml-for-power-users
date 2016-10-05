VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Office Document Properties"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
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
   ScaleHeight     =   4725
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDumpData 
      Caption         =   "Dump Data"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get Info"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox lstFiles 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   3495
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "XML files in current directory:"
      Height          =   255
      Left            =   165
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    ' Create the FileSystemObject for accessing files and directories.
    Dim fso As Scripting.FileSystemObject
    Set fso = New FileSystemObject
    
    ' Get the startup folder.
    Dim Dir As Scripting.Folder
    Set Dir = fso.GetFolder(App.Path)
    
    ' Check all the files in the startup folder.
    Dim File As Scripting.File
    For Each File In Dir.Files
        ' If the file has the extension ".xml", add it to the list.
        If fso.GetExtensionName(File.Name) = "xml" Then
            lstFiles.AddItem (File.Name)
        End If
    Next

End Sub


Private Sub cmdGetInfo_Click()

    ' Use this string to store all the information we can retrieve.
    Dim Info As String

    ' Create the DOMDocument object for accessing the XML.
    Dim Doc As DOMDocument
    Set Doc = New DOMDocument
    
    ' Load the selected file into memory.
    If Doc.Load(App.Path & "\" & lstFiles.Text) Then
    
        ' Define the node objects that will be used to retrieve information.
        Dim PropertiesNode As IXMLDOMNode
        Dim PropertyNode As IXMLDOMNode
                
        ' Before performing the search, you must map the namespace that
        ' you will use. In this case, it's the Office namespace, which is
        ' mapped to the prefix "o". The prefix used here does NOT need to
        ' match the mapping used in the XML document, as long as the namespace
        ' URI is the same.
        Call Doc.setProperty("SelectionNamespaces", _
          "xmlns:o='urn:schemas-microsoft-com:office:office'")
  
        ' Search for a <DocumentProperties> node in the Office namespace.
        Set PropertiesNode = Doc.documentElement.selectSingleNode("//o:DocumentProperties")
        
        ' Check if we found it.
        If PropertiesNode Is Nothing Then
            Info = " This file is not an Office XML document."
        Else
            ' Get the name of the root node to determine the type of document
            ' being read. Excel uses <Workbook> and Word uses <wordDocument>.
            If Doc.documentElement.baseName = "Workbook" Then
                Info = " SpreadsheetML document." & vbNewLine
            ElseIf Doc.documentElement.baseName = "wordDocument" Then
                Info = " WordML document." & vbNewLine
            Else
               Info = " Unrecognizable Office document." & vbNewLine
            End If
            
            ' Drill-down in the <DocumentProperties> node to get additional
            ' information, like the author, creation date, etc.
            ' Every time a piece of informatio is successfully retrieved,
            ' add it to the Info string.
            Set PropertyNode = PropertiesNode.selectSingleNode("//o:Author")
            If Not PropertyNode Is Nothing Then
                Info = Info & "  Author: " & PropertyNode.Text & vbNewLine
            End If
            
            Set PropertyNode = PropertiesNode.selectSingleNode("//o:Created")
            If Not PropertyNode Is Nothing Then
                Info = Info & "  Created Date: " & Left(PropertyNode.Text, 10) & vbNewLine
                Info = Info & "  Created Time: " & Mid(PropertyNode.Text, 12, 8) & vbNewLine
            End If
            
            Set PropertyNode = PropertiesNode.selectSingleNode("//o:Version")
            If Not PropertyNode Is Nothing Then
                Info = Info & "  App Version: " & PropertyNode.Text & vbNewLine
            End If
            
            Set PropertyNode = PropertiesNode.selectSingleNode("//o:Pages")
            If Not PropertyNode Is Nothing Then
                Info = Info & "  Pages: " & PropertyNode.Text & vbNewLine
            End If
        End If
        
    Else
        MsgBox "XML file not found."
    End If

    ' Display all the retrieved data in a label.
    lblInfo.Caption = Info
    
End Sub


Private Sub cmdDumpData_Click()
    ' Use this string to store all the information we can retrieve.
    Dim Info As String

    ' Create the DOMDocument object for accessing the XML.
    Dim Doc As DOMDocument
    Set Doc = New DOMDocument
    
    Dim ContentNode As IXMLDOMNode
    
    ' Load the selected file into memory.
    If Doc.Load(App.Path & "\" & lstFiles.Text) Then
    
        ' Get the name of the root node to determine the type of document
        ' being read. Excel uses <Workbook> and Word uses <wordDocument>.
        If Doc.documentElement.baseName = "Workbook" Then
            ' Before performing the search, you must map the namespace that
            ' you will use. In this case, it's the spreadsheet namespace, which is
            ' mapped to the prefix "ss". The prefix used here does NOT need to
            ' match the mapping used in the XML document, as long as the namespace
            ' URI is the same.
            Call Doc.setProperty("SelectionNamespaces", _
              "xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'")
          
            ' Search for all tabular data.
            Dim ContentNodes As IXMLDOMNodeList
            Set ContentNodes = Doc.documentElement.selectNodes("//ss:Worksheet/ss:Table")
          
            ' Ignore all the elements, and just get the text.
            For Each ContentNode In ContentNodes
                If Not ContentNode Is Nothing Then
                    Info = Info & ContentNode.Text
                End If
            Next
        ElseIf Doc.documentElement.baseName = "wordDocument" Then
            ' Before performing the search, you must map the namespace that
            ' you will use. In this case, it's the WordML namespace, which is
            ' mapped to the prefix "w". The prefix used here does NOT need to
            ' match the mapping used in the XML document, as long as the namespace
            ' URI is the same.
            Call Doc.setProperty("SelectionNamespaces", _
              "xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'")
          
            ' Search for a <body> node.
            Set ContentNode = Doc.documentElement.selectSingleNode("//w:body")
          
            ' Ignore all the elements, and just get the text.
            If Not ContentNode Is Nothing Then
                Info = Info & ContentNode.Text
            End If
            
        Else
            Info = " Unrecognizable document." & vbNewLine
        End If
    
    Else
        MsgBox "XML file not found."
    End If

    ' Display all the retrieved data in a label.
    lblInfo.Caption = Info
    
End Sub

