Imports System.Xml
Imports System.io

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtXmlDoc As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseXml As System.Windows.Forms.Button
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtOutput As System.Windows.Forms.TextBox
    Friend WithEvents cmdShowNamespaces As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtXmlDoc = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdBrowseXml = New System.Windows.Forms.Button
        Me.cmdShowNamespaces = New System.Windows.Forms.Button
        Me.txtOutput = New System.Windows.Forms.TextBox
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtXmlDoc
        '
        Me.txtXmlDoc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtXmlDoc.BackColor = System.Drawing.SystemColors.Window
        Me.txtXmlDoc.Location = New System.Drawing.Point(8, 24)
        Me.txtXmlDoc.Name = "txtXmlDoc"
        Me.txtXmlDoc.ReadOnly = True
        Me.txtXmlDoc.Size = New System.Drawing.Size(436, 21)
        Me.txtXmlDoc.TabIndex = 2
        Me.txtXmlDoc.TabStop = False
        Me.txtXmlDoc.Text = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.cmdBrowseXml)
        Me.GroupBox1.Controls.Add(Me.txtXmlDoc)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(524, 60)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "XML Document"
        '
        'cmdBrowseXml
        '
        Me.cmdBrowseXml.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseXml.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdBrowseXml.Location = New System.Drawing.Point(448, 24)
        Me.cmdBrowseXml.Name = "cmdBrowseXml"
        Me.cmdBrowseXml.Size = New System.Drawing.Size(68, 23)
        Me.cmdBrowseXml.TabIndex = 0
        Me.cmdBrowseXml.Text = "Browse"
        '
        'cmdShowNamespaces
        '
        Me.cmdShowNamespaces.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdShowNamespaces.Location = New System.Drawing.Point(8, 88)
        Me.cmdShowNamespaces.Name = "cmdShowNamespaces"
        Me.cmdShowNamespaces.Size = New System.Drawing.Size(112, 28)
        Me.cmdShowNamespaces.TabIndex = 2
        Me.cmdShowNamespaces.Text = "Show Namespaces"
        '
        'txtOutput
        '
        Me.txtOutput.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOutput.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOutput.Location = New System.Drawing.Point(8, 124)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.ReadOnly = True
        Me.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOutput.Size = New System.Drawing.Size(524, 208)
        Me.txtOutput.TabIndex = 3
        Me.txtOutput.Text = ""
        '
        'dlgOpen
        '
        Me.dlgOpen.DefaultExt = "xml"
        Me.dlgOpen.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
        '
        'Form1
        '
        Me.AcceptButton = Me.cmdShowNamespaces
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(540, 346)
        Me.Controls.Add(Me.txtOutput)
        Me.Controls.Add(Me.cmdShowNamespaces)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.Text = "Show Namespace Information"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdBrowseXml_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseXml.Click
        dlgOpen.FilterIndex = 1
        If dlgOpen.ShowDialog = DialogResult.OK Then
            txtXmlDoc.Text = dlgOpen.FileName
        End If
    End Sub

    Private Sub cmdShowNamespaces_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowNamespaces.Click
        If txtXmlDoc.Text = "" Then
            MessageBox.Show("No XML document has been specified", "Missing Information")
            Return
        End If

        txtOutput.Text = ""

        Dim Stream As FileStream
        Try
            Stream = New FileStream(txtXmlDoc.Text, FileMode.Open)
            Dim Reader As New XmlTextReader(Stream)

            Do While Reader.Read()
                If Reader.NodeType = XmlNodeType.Element Then
                    txtOutput.Text &= ("<" & Reader.Name & ">").PadRight(22)
                    txtOutput.Text &= "Namespace: " & Reader.NamespaceURI
                    txtOutput.Text &= vbNewLine
                    If Reader.HasAttributes Then
                        For i As Integer = 0 To Reader.AttributeCount - 1
                            Reader.MoveToAttribute(i)
                            txtOutput.Text &= ("ATTR: " & Reader.Name).PadRight(22)
                            txtOutput.Text &= "Namespace: " & Reader.NamespaceURI
                            txtOutput.Text &= vbNewLine
                        Next
                    End If
                End If
            Loop

        Catch Err As Exception
            MessageBox.Show(Err.Message, "Error")
        Finally
            If Not Stream Is Nothing Then Stream.Close()
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dlgOpen.InitialDirectory = Application.StartupPath
    End Sub

    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.Run(New Form1)
    End Sub

End Class
