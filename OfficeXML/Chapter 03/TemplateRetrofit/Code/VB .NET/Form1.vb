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
    Friend WithEvents lstItems As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lstItems = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.txtID = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'lstItems
        '
        Me.lstItems.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstItems.IntegralHeight = False
        Me.lstItems.Location = New System.Drawing.Point(8, 68)
        Me.lstItems.Name = "lstItems"
        Me.lstItems.Size = New System.Drawing.Size(292, 204)
        Me.lstItems.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Description:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "ID:"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(108, 8)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(188, 21)
        Me.txtName.TabIndex = 3
        Me.txtName.Text = ""
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(108, 36)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(188, 21)
        Me.txtID.TabIndex = 4
        Me.txtID.Text = ""
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(308, 282)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstItems)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.Text = "XML Reader"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Create the DOMDocument object.
        Dim Doc As New XmlDocument

        ' Load the ExpenseData.xml file into memory.
        If File.Exists(Application.StartupPath & "\ExpenseData.xml") Then

            Doc.Load(Application.StartupPath & "\ExpenseData.xml")

            ' Scan all the nodes under the root element.
            Dim Node As XmlNode
            For Each Node In Doc.DocumentElement.ChildNodes

                ' Check if the node contains the metadata (header) section.
                If Node.Name = "Meta" Then

                    ' Retrieve information about the expense report.
                    txtName.Text = Node.ChildNodes(0).InnerText
                    txtID.Text = Node.ChildNodes(2).InnerText

                    ' Check if the node contains the expense list.
                ElseIf Node.Name = "ExpenseItem" Then

                    ' Get the description and total of each expense item.
                    Dim Description As String
                    Description = Node.ChildNodes(1).InnerText
                    Dim Total As String
                    Total = FormatCurrency(Node.ChildNodes(10).InnerText)

                    ' Show this information in a listbox.
                    lstItems.Items.Add(Description & " (" & Total & ")")
                End If
            Next
        Else
            MsgBox("File ExpenseData.xml not found in this directory.")
        End If

    End Sub
End Class
