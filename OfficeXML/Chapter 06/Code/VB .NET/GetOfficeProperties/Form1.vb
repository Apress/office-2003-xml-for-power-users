Imports System.IO
Imports System.Xml
Imports Microsoft.VisualBasic

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdGetInfo As System.Windows.Forms.Button
    Friend WithEvents lstFiles As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdDumpData As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdGetInfo = New System.Windows.Forms.Button
        Me.lstFiles = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmdDumpData = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblInfo = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdGetInfo
        '
        Me.cmdGetInfo.Location = New System.Drawing.Point(232, 48)
        Me.cmdGetInfo.Name = "cmdGetInfo"
        Me.cmdGetInfo.Size = New System.Drawing.Size(96, 24)
        Me.cmdGetInfo.TabIndex = 0
        Me.cmdGetInfo.Text = "Get Info"
        '
        'lstFiles
        '
        Me.lstFiles.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lstFiles.IntegralHeight = False
        Me.lstFiles.Location = New System.Drawing.Point(12, 44)
        Me.lstFiles.Name = "lstFiles"
        Me.lstFiles.Size = New System.Drawing.Size(208, 228)
        Me.lstFiles.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(188, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "XML files in current directory:"
        '
        'cmdDumpData
        '
        Me.cmdDumpData.Location = New System.Drawing.Point(336, 48)
        Me.cmdDumpData.Name = "cmdDumpData"
        Me.cmdDumpData.Size = New System.Drawing.Size(96, 24)
        Me.cmdDumpData.TabIndex = 3
        Me.cmdDumpData.Text = "Dump Data"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(308, 204)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(4, 4)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Label2"
        '
        'lblInfo
        '
        Me.lblInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblInfo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblInfo.Location = New System.Drawing.Point(232, 80)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(200, 192)
        Me.lblInfo.TabIndex = 6
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(448, 286)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmdDumpData)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstFiles)
        Me.Controls.Add(Me.cmdGetInfo)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.Text = "Office Document Properties"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdGetInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetInfo.Click

        ' Use this string to store all the information we can retrieve.
        Dim Info As String

        ' Load the selected file into memory.
        If File.Exists(Application.StartupPath & "\" & lstFiles.Text) Then

            ' Create the DOMDocument object for accessing the XML.
            Dim Doc As New XmlDocument
            Doc.Load(Application.StartupPath & "\" & lstFiles.Text)

            ' Define the node objects that will be used to retrieve information.
            Dim PropertiesNode As XmlNode
            Dim PropertyNode As XmlNode

            Dim NM As New XmlNamespaceManager(Doc.NameTable)
            NM.AddNamespace("o", "urn:schemas-microsoft-com:office:office")

            ' Search for a <DocumentProperties> node in the Office namespace.
            PropertiesNode = Doc.SelectSingleNode("//o:DocumentProperties", NM)

            ' Check if we found it.
            If PropertiesNode Is Nothing Then
                Info = " This file is not an Office XML document."
            Else
                ' Get the name of the root node to determine the type of document
                ' being read. Excel uses <Workbook> and Word uses <wordDocument>.
                If Doc.DocumentElement.LocalName = "Workbook" Then
                    Info = " SpreadsheetML document." & vbNewLine
                ElseIf Doc.DocumentElement.LocalName = "wordDocument" Then
                    Info = " WordML document." & vbNewLine
                Else
                    Info = " Unrecognizable Office document." & vbNewLine
                End If

                ' Drill-down in the <DocumentProperties> node to get additional
                ' information, like the author, creation date, etc.
                ' Every time a piece of informatio is successfully retrieved,
                ' add it to the Info string.
                PropertyNode = PropertiesNode.SelectSingleNode("//o:Author", NM)
                If Not PropertyNode Is Nothing Then
                    Info = Info & "  Author: " & PropertyNode.InnerText & vbNewLine
                End If

                PropertyNode = PropertiesNode.SelectSingleNode("//o:Created", NM)
                If Not PropertyNode Is Nothing Then
                    Info = Info & "  Created Date: " & Microsoft.VisualBasic.Left(PropertyNode.InnerText, 10) & vbNewLine
                    Info = Info & "  Created Time: " & Mid(PropertyNode.InnerText, 12, 8) & vbNewLine
                End If

                PropertyNode = PropertiesNode.SelectSingleNode("//o:Version", NM)
                If Not PropertyNode Is Nothing Then
                    Info = Info & "  App Version: " & PropertyNode.InnerText & vbNewLine
                End If

                PropertyNode = PropertiesNode.SelectSingleNode("//o:Pages", NM)
                If Not PropertyNode Is Nothing Then
                    Info = Info & "  Pages: " & PropertyNode.InnerText & vbNewLine
                End If
            End If

        Else
            MsgBox("XML file not found.")
        End If

        ' Display all the retrieved data in a label.
        lblInfo.Text = Info


    End Sub

    Private Sub cmdDumpData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDumpData.Click

        ' Use this string to store all the information we can retrieve.
        Dim Info As String

        ' Load the selected file into memory.
        If File.Exists(Application.StartupPath & "\" & lstFiles.Text) Then

            ' Create the DOMDocument object for accessing the XML.
            Dim Doc As New XmlDocument
            Doc.Load(Application.StartupPath & "\" & lstFiles.Text)
            
            Dim ContentNode As XmlNode

            ' Load the selected file into memory.
            ' Get the name of the root node to determine the type of document
            ' being read. Excel uses <Workbook> and Word uses <wordDocument>.
            If Doc.DocumentElement.LocalName = "Workbook" Then

                Dim NM As New XmlNamespaceManager(Doc.NameTable)
                NM.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet")

                ' Search for all tabular data.
                Dim ContentNodes As XmlNodeList
                ContentNodes = Doc.DocumentElement.SelectNodes("//ss:Worksheet/ss:Table", NM)

                ' Ignore all the elements, and just get the text.
                For Each ContentNode In ContentNodes
                    If Not ContentNode Is Nothing Then
                        Info = Info & ContentNode.InnerText
                    End If
                Next

            ElseIf Doc.DocumentElement.LocalName = "wordDocument" Then

                Dim NM As New XmlNamespaceManager(Doc.NameTable)
                NM.AddNamespace("w", "http://schemas.microsoft.com/office/word/2003/wordml")

                ' Search for a <body> node.
                ContentNode = Doc.DocumentElement.SelectSingleNode("//w:body", NM)

                ' Ignore all the elements, and just get the text.
                If Not ContentNode Is Nothing Then
                    Info = Info & ContentNode.InnerText
                End If

            Else
                Info = " Unrecognizable document." & vbNewLine
            End If

        Else
                MsgBox("XML file not found.")
            End If

            ' Display all the retrieved data in a label.
        lblInfo.Text = Info
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Get the startup folder.
        Dim Dir As New DirectoryInfo(Application.StartupPath)

        ' Check all the files in the startup folder.
        Dim File As FileInfo
        For Each File In Dir.GetFiles()
            ' If the file has the extension ".xml", add it to the list.
            If Path.GetExtension(File.Name) = ".xml" Then
                lstFiles.Items.Add(File.Name)
            End If
        Next
    End Sub
End Class
