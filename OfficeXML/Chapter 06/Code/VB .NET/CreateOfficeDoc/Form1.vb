Imports System.Xml
Imports System.IO

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtTo As System.Windows.Forms.TextBox
    Friend WithEvents txtFrom As System.Windows.Forms.TextBox
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents cmdGenerate As System.Windows.Forms.Button
    Friend WithEvents txtBody As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtTo = New System.Windows.Forms.TextBox
        Me.txtFrom = New System.Windows.Forms.TextBox
        Me.txtSubject = New System.Windows.Forms.TextBox
        Me.txtBody = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtFileName = New System.Windows.Forms.TextBox
        Me.cmdGenerate = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(20, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "To:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(20, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 12)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "From:"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(20, 76)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 12)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Subject:"
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(80, 16)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(196, 21)
        Me.txtTo.TabIndex = 3
        Me.txtTo.Text = "John Diagamore"
        '
        'txtFrom
        '
        Me.txtFrom.Location = New System.Drawing.Point(80, 40)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(196, 21)
        Me.txtFrom.TabIndex = 4
        Me.txtFrom.Text = "Lisa Stitzwilliam"
        '
        'txtSubject
        '
        Me.txtSubject.Location = New System.Drawing.Point(80, 68)
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(196, 21)
        Me.txtSubject.TabIndex = 5
        Me.txtSubject.Text = "New Memo Generating Software"
        '
        'txtBody
        '
        Me.txtBody.Location = New System.Drawing.Point(12, 108)
        Me.txtBody.Multiline = True
        Me.txtBody.Name = "txtBody"
        Me.txtBody.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtBody.Size = New System.Drawing.Size(348, 156)
        Me.txtBody.TabIndex = 6
        Me.txtBody.Text = "It's now possible to generate our memos using a new Visual Basic .NET application" & _
        ". The application outputs the same Word document, in XML format, and you can edi" & _
        "t it further using Word 2003."
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 284)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Generated File:"
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(120, 280)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(240, 21)
        Me.txtFileName.TabIndex = 8
        Me.txtFileName.Text = ""
        '
        'cmdGenerate
        '
        Me.cmdGenerate.Location = New System.Drawing.Point(268, 316)
        Me.cmdGenerate.Name = "cmdGenerate"
        Me.cmdGenerate.Size = New System.Drawing.Size(92, 28)
        Me.cmdGenerate.TabIndex = 9
        Me.cmdGenerate.Text = "Generate"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(372, 354)
        Me.Controls.Add(Me.cmdGenerate)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.txtBody)
        Me.Controls.Add(Me.txtSubject)
        Me.Controls.Add(Me.txtFrom)
        Me.Controls.Add(Me.txtTo)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.Text = "Memo Generator"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtFileName.Text = Application.StartupPath & "\NewMemo.xml"
    End Sub

    Private Sub cmdGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerate.Click
        ' Create the DOMDocument object for accessing the XML.
        Dim Doc As New XmlDocument

        ' Load the selected file into memory.
        If File.Exists(Application.StartupPath & "\MemoTemplate.xml") Then
            Doc.Load(Application.StartupPath & "\MemoTemplate.xml")


            Dim TextNodes As XmlNodeList
            TextNodes = Doc.GetElementsByTagName("t", "http://schemas.microsoft.com/office/word/2003/wordml")

            Dim TextNode As XmlNode
            For Each TextNode In TextNodes
                Select Case TextNode.InnerText
                    Case "CompanyName"
                        TextNode.InnerText = "ProseTech"
                    Case "ToName"
                        TextNode.InnerText = txtTo.Text
                    Case "FromName"
                        TextNode.InnerText = txtFrom.Text
                    Case "DateInfo"
                        TextNode.InnerText = Now
                    Case "SubjectLine"
                        TextNode.InnerText = txtSubject.Text
                    Case "MemoBody"
                        TextNode.InnerText = txtBody.Text
                End Select

            Next

            Doc.Save(txtFileName.Text)
            MsgBox("File saved.")

        Else
            MsgBox("MemoTemplate.xml file is missing.")
        End If
    End Sub
End Class
