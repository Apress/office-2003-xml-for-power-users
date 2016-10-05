Imports System.Xml
Imports System.Xml.Xsl
Imports System.Xml.XPath

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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtXmlDoc As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseXml As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseXsd As System.Windows.Forms.Button
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtXslDoc As System.Windows.Forms.TextBox
    Friend WithEvents cmdTransform As System.Windows.Forms.Button
    Friend WithEvents txtOutput As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtXmlDoc = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdBrowseXml = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdBrowseXsd = New System.Windows.Forms.Button
        Me.txtXslDoc = New System.Windows.Forms.TextBox
        Me.cmdTransform = New System.Windows.Forms.Button
        Me.txtOutput = New System.Windows.Forms.TextBox
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
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
        Me.txtXmlDoc.Size = New System.Drawing.Size(276, 21)
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
        Me.GroupBox1.Size = New System.Drawing.Size(364, 60)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "XML Document"
        '
        'cmdBrowseXml
        '
        Me.cmdBrowseXml.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseXml.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdBrowseXml.Location = New System.Drawing.Point(288, 24)
        Me.cmdBrowseXml.Name = "cmdBrowseXml"
        Me.cmdBrowseXml.Size = New System.Drawing.Size(68, 23)
        Me.cmdBrowseXml.TabIndex = 0
        Me.cmdBrowseXml.Text = "Browse"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.cmdBrowseXsd)
        Me.GroupBox2.Controls.Add(Me.txtXslDoc)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(8, 72)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(364, 60)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "XSLT Document"
        '
        'cmdBrowseXsd
        '
        Me.cmdBrowseXsd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseXsd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdBrowseXsd.Location = New System.Drawing.Point(288, 24)
        Me.cmdBrowseXsd.Name = "cmdBrowseXsd"
        Me.cmdBrowseXsd.Size = New System.Drawing.Size(68, 23)
        Me.cmdBrowseXsd.TabIndex = 1
        Me.cmdBrowseXsd.Text = "Browse"
        '
        'txtXslDoc
        '
        Me.txtXslDoc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtXslDoc.BackColor = System.Drawing.SystemColors.Window
        Me.txtXslDoc.Location = New System.Drawing.Point(8, 24)
        Me.txtXslDoc.Name = "txtXslDoc"
        Me.txtXslDoc.ReadOnly = True
        Me.txtXslDoc.Size = New System.Drawing.Size(276, 21)
        Me.txtXslDoc.TabIndex = 2
        Me.txtXslDoc.TabStop = False
        Me.txtXslDoc.Text = ""
        '
        'cmdTransform
        '
        Me.cmdTransform.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdTransform.Location = New System.Drawing.Point(8, 156)
        Me.cmdTransform.Name = "cmdTransform"
        Me.cmdTransform.Size = New System.Drawing.Size(88, 28)
        Me.cmdTransform.TabIndex = 2
        Me.cmdTransform.Text = "Transform Now "
        '
        'txtOutput
        '
        Me.txtOutput.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOutput.Location = New System.Drawing.Point(8, 188)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.ReadOnly = True
        Me.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOutput.Size = New System.Drawing.Size(364, 204)
        Me.txtOutput.TabIndex = 3
        Me.txtOutput.Text = ""
        '
        'dlgOpen
        '
        Me.dlgOpen.DefaultExt = "xml"
        Me.dlgOpen.Filter = "XML files (*.xml)|*.xml|XSLT files (*.xslt)|*.xslt|All files (*.*)|*.*"
        '
        'Form1
        '
        Me.AcceptButton = Me.cmdTransform
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(380, 406)
        Me.Controls.Add(Me.txtOutput)
        Me.Controls.Add(Me.cmdTransform)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.Text = "Transform Your XML Document"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdBrowseXml_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseXml.Click
        dlgOpen.FilterIndex = 1
        If dlgOpen.ShowDialog = DialogResult.OK Then
            txtXmlDoc.Text = dlgOpen.FileName
            txtXslDoc.Text = ""
        End If
    End Sub

    Private Sub cmdBrowseXsd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseXsd.Click
        dlgOpen.FilterIndex = 2
        If dlgOpen.ShowDialog = DialogResult.OK Then
            txtXslDoc.Text = dlgOpen.FileName
        End If
    End Sub

    Private Sub cmdTransform_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTransform.Click
        If txtXmlDoc.Text = "" Then
            MessageBox.Show("No XML document has been specified", "Missing Information")
            Return
        End If

        If txtXslDoc.Text = "" Then
            MessageBox.Show("No XSLT document has been specified", "Missing Information")
            Return
        End If

        Dim Transform As New XslTransform

        Try
            Transform.Load(txtXslDoc.Text)

            Dim Doc As New XPathDocument(txtXmlDoc.Text)
            Dim ms As New System.IO.MemoryStream
            Transform.Transform(Doc, Nothing, ms)
            
            txtOutput.Text = System.Text.Encoding.UTF8.GetString(ms.ToArray())
            txtOutput.Text = txtOutput.Text.Substring(1)
            'Dim Reader As XmlReader = Transform.Transform(Doc, Nothing)
            'Reader.MoveToContent()
            'txtOutput.Text = Reader.ReadOuterXml()
            'Reader.Close()
        Catch Err As Exception
            MessageBox.Show(Err.Message, "Error")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dlgOpen.FileName = Application.StartupPath
    End Sub
End Class
