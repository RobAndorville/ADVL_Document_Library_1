<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmXmlDisplay
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim XmlHtmDisplaySettings1 As ADVL_Utilities_Library_1.XmlHtmDisplaySettings = New ADVL_Utilities_Library_1.XmlHtmDisplaySettings()
        Dim TextSettings1 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings2 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings3 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings4 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings5 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings6 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings7 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings8 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings9 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings10 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings11 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Dim TextSettings12 As ADVL_Utilities_Library_1.TextSettings = New ADVL_Utilities_Library_1.TextSettings()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnAddToLibrary = New System.Windows.Forms.Button()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.XmlHtmDisplay1 = New ADVL_Utilities_Library_1.XmlHtmDisplay(Me.components)
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(888, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnOpen
        '
        Me.btnOpen.Location = New System.Drawing.Point(12, 12)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(48, 22)
        Me.btnOpen.TabIndex = 52
        Me.btnOpen.Text = "Open"
        Me.btnOpen.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Location = New System.Drawing.Point(66, 12)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(48, 22)
        Me.btnEdit.TabIndex = 54
        Me.btnEdit.Text = "Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(120, 12)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(48, 22)
        Me.btnSave.TabIndex = 60
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnAddToLibrary
        '
        Me.btnAddToLibrary.Location = New System.Drawing.Point(222, 12)
        Me.btnAddToLibrary.Name = "btnAddToLibrary"
        Me.btnAddToLibrary.Size = New System.Drawing.Size(82, 22)
        Me.btnAddToLibrary.TabIndex = 227
        Me.btnAddToLibrary.Text = "Add to Library"
        Me.btnAddToLibrary.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(12, 43)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(38, 13)
        Me.Label14.TabIndex = 229
        Me.Label14.Text = "Name:"
        '
        'txtFileName
        '
        Me.txtFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFileName.Location = New System.Drawing.Point(56, 40)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.ReadOnly = True
        Me.txtFileName.Size = New System.Drawing.Size(880, 20)
        Me.txtFileName.TabIndex = 228
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(174, 12)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(42, 22)
        Me.btnClose.TabIndex = 237
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'XmlHtmDisplay1
        '
        Me.XmlHtmDisplay1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.XmlHtmDisplay1.Location = New System.Drawing.Point(12, 66)
        Me.XmlHtmDisplay1.Name = "XmlHtmDisplay1"
        TextSettings1.Bold = False
        TextSettings1.Color = System.Drawing.Color.Blue
        TextSettings1.ColorIndex = 4
        TextSettings1.FontIndex = 1
        TextSettings1.FontName = "Arial"
        TextSettings1.HalfPointSize = 20
        TextSettings1.Italic = False
        TextSettings1.PointSize = 10.0!
        XmlHtmDisplaySettings1.HAttribute = TextSettings1
        TextSettings2.Bold = False
        TextSettings2.Color = System.Drawing.Color.Gray
        TextSettings2.ColorIndex = 6
        TextSettings2.FontIndex = 1
        TextSettings2.FontName = "Arial"
        TextSettings2.HalfPointSize = 20
        TextSettings2.Italic = False
        TextSettings2.PointSize = 10.0!
        XmlHtmDisplaySettings1.HChar = TextSettings2
        TextSettings3.Bold = False
        TextSettings3.Color = System.Drawing.Color.Gray
        TextSettings3.ColorIndex = 6
        TextSettings3.FontIndex = 1
        TextSettings3.FontName = "Arial"
        TextSettings3.HalfPointSize = 20
        TextSettings3.Italic = False
        TextSettings3.PointSize = 10.0!
        XmlHtmDisplaySettings1.HComment = TextSettings3
        TextSettings4.Bold = False
        TextSettings4.Color = System.Drawing.Color.DarkRed
        TextSettings4.ColorIndex = 2
        TextSettings4.FontIndex = 1
        TextSettings4.FontName = "Arial"
        TextSettings4.HalfPointSize = 20
        TextSettings4.Italic = False
        TextSettings4.PointSize = 10.0!
        XmlHtmDisplaySettings1.HElement = TextSettings4
        TextSettings5.Bold = False
        TextSettings5.Color = System.Drawing.Color.Black
        TextSettings5.ColorIndex = 5
        TextSettings5.FontIndex = 1
        TextSettings5.FontName = "Arial"
        TextSettings5.HalfPointSize = 20
        TextSettings5.Italic = False
        TextSettings5.PointSize = 10.0!
        XmlHtmDisplaySettings1.HStyle = TextSettings5
        TextSettings6.Bold = False
        TextSettings6.Color = System.Drawing.Color.Black
        TextSettings6.ColorIndex = 5
        TextSettings6.FontIndex = 1
        TextSettings6.FontName = "Arial"
        TextSettings6.HalfPointSize = 20
        TextSettings6.Italic = False
        TextSettings6.PointSize = 10.0!
        XmlHtmDisplaySettings1.HValue = TextSettings6
        TextSettings7.Bold = False
        TextSettings7.Color = System.Drawing.Color.Red
        TextSettings7.ColorIndex = 3
        TextSettings7.FontIndex = 1
        TextSettings7.FontName = "Arial"
        TextSettings7.HalfPointSize = 20
        TextSettings7.Italic = False
        TextSettings7.PointSize = 10.0!
        XmlHtmDisplaySettings1.XAttributeKey = TextSettings7
        TextSettings8.Bold = False
        TextSettings8.Color = System.Drawing.Color.Blue
        TextSettings8.ColorIndex = 4
        TextSettings8.FontIndex = 1
        TextSettings8.FontName = "Arial"
        TextSettings8.HalfPointSize = 20
        TextSettings8.Italic = False
        TextSettings8.PointSize = 10.0!
        XmlHtmDisplaySettings1.XAttributeValue = TextSettings8
        TextSettings9.Bold = False
        TextSettings9.Color = System.Drawing.Color.Gray
        TextSettings9.ColorIndex = 6
        TextSettings9.FontIndex = 1
        TextSettings9.FontName = "Arial"
        TextSettings9.HalfPointSize = 20
        TextSettings9.Italic = False
        TextSettings9.PointSize = 10.0!
        XmlHtmDisplaySettings1.XComment = TextSettings9
        TextSettings10.Bold = False
        TextSettings10.Color = System.Drawing.Color.DarkRed
        TextSettings10.ColorIndex = 2
        TextSettings10.FontIndex = 1
        TextSettings10.FontName = "Arial"
        TextSettings10.HalfPointSize = 20
        TextSettings10.Italic = False
        TextSettings10.PointSize = 10.0!
        XmlHtmDisplaySettings1.XElement = TextSettings10
        XmlHtmDisplaySettings1.XIndentSpaces = 4
        XmlHtmDisplaySettings1.XmlLargeFileSizeLimit = 1000000
        TextSettings11.Bold = False
        TextSettings11.Color = System.Drawing.Color.Blue
        TextSettings11.ColorIndex = 1
        TextSettings11.FontIndex = 1
        TextSettings11.FontName = "Arial"
        TextSettings11.HalfPointSize = 20
        TextSettings11.Italic = False
        TextSettings11.PointSize = 10.0!
        XmlHtmDisplaySettings1.XTag = TextSettings11
        TextSettings12.Bold = False
        TextSettings12.Color = System.Drawing.Color.Black
        TextSettings12.ColorIndex = 5
        TextSettings12.FontIndex = 1
        TextSettings12.FontName = "Arial"
        TextSettings12.HalfPointSize = 20
        TextSettings12.Italic = False
        TextSettings12.PointSize = 10.0!
        XmlHtmDisplaySettings1.XValue = TextSettings12
        Me.XmlHtmDisplay1.Settings = XmlHtmDisplaySettings1
        Me.XmlHtmDisplay1.Size = New System.Drawing.Size(924, 427)
        Me.XmlHtmDisplay1.TabIndex = 238
        Me.XmlHtmDisplay1.Text = ""
        '
        'frmXmlDisplay
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(948, 505)
        Me.Controls.Add(Me.XmlHtmDisplay1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.btnAddToLibrary)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmXmlDisplay"
        Me.Text = "XML Display"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnOpen As Button
    Friend WithEvents btnEdit As Button
    Friend WithEvents btnSave As Button
    Friend WithEvents btnAddToLibrary As Button
    Friend WithEvents Label14 As Label
    Friend WithEvents txtFileName As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnClose As Button
    Friend WithEvents XmlHtmDisplay1 As ADVL_Utilities_Library_1.XmlHtmDisplay
End Class
