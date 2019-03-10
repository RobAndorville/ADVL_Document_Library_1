<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewRtf
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
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDirectory = New System.Windows.Forms.TextBox()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.chkAddToLibrary = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.btnCreate = New System.Windows.Forms.Button()
        Me.chkOpenInNewWindow = New System.Windows.Forms.CheckBox()
        Me.rbFileInFileSystem = New System.Windows.Forms.RadioButton()
        Me.rbFileInProject = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(496, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "File name:"
        '
        'txtFileName
        '
        Me.txtFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFileName.Location = New System.Drawing.Point(73, 40)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(471, 20)
        Me.txtFileName.TabIndex = 10
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 92)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Directory:"
        '
        'txtDirectory
        '
        Me.txtDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDirectory.Location = New System.Drawing.Point(73, 89)
        Me.txtDirectory.Multiline = True
        Me.txtDirectory.Name = "txtDirectory"
        Me.txtDirectory.Size = New System.Drawing.Size(471, 41)
        Me.txtDirectory.TabIndex = 12
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(12, 108)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(48, 22)
        Me.btnFind.TabIndex = 13
        Me.btnFind.Text = "Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'chkAddToLibrary
        '
        Me.chkAddToLibrary.AutoSize = True
        Me.chkAddToLibrary.Location = New System.Drawing.Point(12, 165)
        Me.chkAddToLibrary.Name = "chkAddToLibrary"
        Me.chkAddToLibrary.Size = New System.Drawing.Size(87, 17)
        Me.chkAddToLibrary.TabIndex = 14
        Me.chkAddToLibrary.Text = "Add to library"
        Me.chkAddToLibrary.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 191)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(115, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Document Description:"
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(12, 207)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(532, 41)
        Me.txtDescription.TabIndex = 16
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(12, 12)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(52, 22)
        Me.btnCreate.TabIndex = 17
        Me.btnCreate.Text = "Create"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'chkOpenInNewWindow
        '
        Me.chkOpenInNewWindow.AutoSize = True
        Me.chkOpenInNewWindow.Location = New System.Drawing.Point(12, 142)
        Me.chkOpenInNewWindow.Name = "chkOpenInNewWindow"
        Me.chkOpenInNewWindow.Size = New System.Drawing.Size(125, 17)
        Me.chkOpenInNewWindow.TabIndex = 18
        Me.chkOpenInNewWindow.Text = "Open in new window"
        Me.chkOpenInNewWindow.UseVisualStyleBackColor = True
        '
        'rbFileInFileSystem
        '
        Me.rbFileInFileSystem.AutoSize = True
        Me.rbFileInFileSystem.Location = New System.Drawing.Point(113, 66)
        Me.rbFileInFileSystem.Name = "rbFileInFileSystem"
        Me.rbFileInFileSystem.Size = New System.Drawing.Size(78, 17)
        Me.rbFileInFileSystem.TabIndex = 252
        Me.rbFileInFileSystem.TabStop = True
        Me.rbFileInFileSystem.Text = "File System"
        Me.rbFileInFileSystem.UseVisualStyleBackColor = True
        '
        'rbFileInProject
        '
        Me.rbFileInProject.AutoSize = True
        Me.rbFileInProject.Location = New System.Drawing.Point(12, 66)
        Me.rbFileInProject.Name = "rbFileInProject"
        Me.rbFileInProject.Size = New System.Drawing.Size(95, 17)
        Me.rbFileInProject.TabIndex = 251
        Me.rbFileInProject.TabStop = True
        Me.rbFileInProject.Text = "Current Project"
        Me.rbFileInProject.UseVisualStyleBackColor = True
        '
        'frmNewRtf
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(556, 286)
        Me.Controls.Add(Me.rbFileInFileSystem)
        Me.Controls.Add(Me.rbFileInProject)
        Me.Controls.Add(Me.chkOpenInNewWindow)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.chkAddToLibrary)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.txtDirectory)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmNewRtf"
        Me.Text = "New Rich Text Format (RTF) Document"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtFileName As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtDirectory As TextBox
    Friend WithEvents btnFind As Button
    Friend WithEvents chkAddToLibrary As CheckBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents btnCreate As Button
    Friend WithEvents chkOpenInNewWindow As CheckBox
    Friend WithEvents rbFileInFileSystem As RadioButton
    Friend WithEvents rbFileInProject As RadioButton
End Class
