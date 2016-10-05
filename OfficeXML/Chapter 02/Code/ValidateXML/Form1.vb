Imports System.Xml
Imports System.Xml.Schema

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
    Friend WithEvents cmdValidate As System.Windows.Forms.Button
    Friend WithEvents txtXmlDoc As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseXml As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseXsd As System.Windows.Forms.Button
    Friend WithEvents txtXsdDoc As System.Windows.Forms.TextBox
    Friend WithEvents txtValidationInfo As System.Windows.Forms.TextBox
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtXmlDoc = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdBrowseXml = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdBrowseXsd = New System.Windows.Forms.Button
        Me.txtXsdDoc = New System.Windows.Forms.TextBox
        Me.cmdValidate = New System.Windows.Forms.Button
        Me.txtValidationInfo = New System.Windows.Forms.TextBox
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
        Me.GroupBox2.Controls.Add(Me.txtXsdDoc)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(8, 72)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(364, 60)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "XSD Document"
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
        'txtXsdDoc
        '
        Me.txtXsdDoc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtXsdDoc.BackColor = System.Drawing.SystemColors.Window
        Me.txtXsdDoc.Location = New System.Drawing.Point(8, 24)
        Me.txtXsdDoc.Name = "txtXsdDoc"
        Me.txtXsdDoc.ReadOnly = True
        Me.txtXsdDoc.Size = New System.Drawing.Size(276, 21)
        Me.txtXsdDoc.TabIndex = 2
        Me.txtXsdDoc.TabStop = False
        Me.txtXsdDoc.Text = ""
        '
        'cmdValidate
        '
        Me.cmdValidate.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.cmdValidate.Location = New System.Drawing.Point(8, 156)
        Me.cmdValidate.Name = "cmdValidate"
        Me.cmdValidate.Size = New System.Drawing.Size(88, 28)
        Me.cmdValidate.TabIndex = 2
        Me.cmdValidate.Text = "Validate Now "
        '
        'txtValidationInfo
        '
        Me.txtValidationInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtValidationInfo.Location = New System.Drawing.Point(8, 188)
        Me.txtValidationInfo.Multiline = True
        Me.txtValidationInfo.Name = "txtValidationInfo"
        Me.txtValidationInfo.ReadOnly = True
        Me.txtValidationInfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtValidationInfo.Size = New System.Drawing.Size(364, 204)
        Me.txtValidationInfo.TabIndex = 3
        Me.txtValidationInfo.Text = ""
        '
        'dlgOpen
        '
        Me.dlgOpen.DefaultExt = "xml"
        Me.dlgOpen.Filter = "XML files (*.xml)|*.xml|XSD files (*.xsd)|*.xsd|All files (*.*)|*.*"
        '
        'Form1
        '
        Me.AcceptButton = Me.cmdValidate
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(380, 406)
        Me.Controls.Add(Me.txtValidationInfo)
        Me.Controls.Add(Me.cmdValidate)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.Text = "Validate Your XML Document"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdBrowseXml_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseXml.Click
        dlgOpen.FilterIndex = 1
        If dlgOpen.ShowDialog = DialogResult.OK Then
            txtXmlDoc.Text = dlgOpen.FileName
            txtXsdDoc.Text = ""
        End If
    End Sub

    Private Sub cmdBrowseXsd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseXsd.Click
        dlgOpen.FilterIndex = 2
        If dlgOpen.ShowDialog = DialogResult.OK Then
            txtXsdDoc.Text = dlgOpen.FileName
        End If
    End Sub

    Private Sub cmdValidate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValidate.Click
        If txtXmlDoc.Text = "" Then
            MessageBox.Show("No XML document has been specified", "Missing Information")
            Return
        End If

        txtValidationInfo.Text = ""
        Dim Validator As New WindowsValidator(txtValidationInfo)

        Dim Success As Boolean
        Try
            Success = Validator.ValidateXml(txtXmlDoc.Text, txtXsdDoc.Text)
            MessageBox.Show("Validation success: " & Success, "Validation Complete")
        Catch Err As Exception
            MessageBox.Show(Err.Message, "Error")
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dlgOpen.InitialDirectory = Application.StartupPath
    End Sub
End Class

Public Class WindowsValidator

    ' Set to True if at least one error exist.
    Private Failed As Boolean
    Private Control As Control

    Public Sub New(ByVal control As Control)
        Me.Control = control
    End Sub

    Public Function ValidateXml(ByVal Xmlfilename As String, _
      ByVal schemaFilename As String) As Boolean

        ' Create the validtor.
        Dim r As New XmlTextReader(Xmlfilename)
        Dim Validator As New XmlValidatingReader(r)
        Validator.ValidationType = ValidationType.Schema
        Dim Schema As New System.Xml.Schema.XmlSchema

        ' Load the schema file into the validator.
        If schemaFilename <> "" Then
            Dim Schemas As New XmlSchemaCollection
            Schemas.Add(Nothing, schemaFilename)
            Validator.Schemas.Add(Schemas)
        End If

        ' Set the validation event handler.
        AddHandler Validator.ValidationEventHandler, _
          AddressOf Me.ValidationEventHandler

        Failed = False

        Try
            ' Read all XML data.
            While Validator.Read()
            End While
            Validator.Close()
        Catch Err As Exception
            Failed = True
            Throw Err
        Finally
            r.Close()
        End Try

        Return Not Failed

    End Function

    Private Sub ValidationEventHandler(ByVal sender As Object, _
      ByVal args As System.Xml.Schema.ValidationEventArgs)
        Failed = True

        ' Display the validation error.
        Control.Text &= "Validation error: " & args.Message & vbNewLine & vbNewLine
    End Sub

    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.Run(New Form1)
    End Sub

End Class
