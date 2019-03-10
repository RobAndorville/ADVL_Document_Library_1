<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewXml
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
        Me.chkOpenInNewWindow = New System.Windows.Forms.CheckBox()
        Me.btnCreate = New System.Windows.Forms.Button()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.chkAddToLibrary = New System.Windows.Forms.CheckBox()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.txtDirectory = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'chkOpenInNewWindow
        '
        Me.chkOpenInNewWindow.AutoSize = True
        Me.chkOpenInNewWindow.Location = New System.Drawing.Point(12, 113)
        Me.chkOpenInNewWindow.Name = "chkOpenInNewWindow"
        Me.chkOpenInNewWindow.Size = New System.Drawing.Size(125, 17)
        Me.chkOpenInNewWindow.TabIndex = 29
        Me.chkOpenInNewWindow.Text = "Open in new window"
        Me.chkOpenInNewWindow.UseVisualStyleBackColor = True
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(12, 12)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(52, 22)
        Me.btnCreate.TabIndex = 28
        Me.btnCreate.Text = "Create"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'txtDescription
        '
        Me.txtDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription.Location = New System.Drawing.Point(73, 159)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(631, 41)
        Me.txtDescription.TabIndex = 27
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 156)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(63, 13)
        Me.Label3.TabIndex = 26
        Me.Label3.Text = "Description:"
        '
        'chkAddToLibrary
        '
        Me.chkAddToLibrary.AutoSize = True
        Me.chkAddToLibrary.Location = New System.Drawing.Point(12, 136)
        Me.chkAddToLibrary.Name = "chkAddToLibrary"
        Me.chkAddToLibrary.Size = New System.Drawing.Size(87, 17)
        Me.chkAddToLibrary.TabIndex = 25
        Me.chkAddToLibrary.Text = "Add to library"
        Me.chkAddToLibrary.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(12, 85)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(48, 22)
        Me.btnFind.TabIndex = 24
        Me.btnFind.Text = "Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'txtDirectory
        '
        Me.txtDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDirectory.Location = New System.Drawing.Point(73, 66)
        Me.txtDirectory.Multiline = True
        Me.txtDirectory.Name = "txtDirectory"
        Me.txtDirectory.Size = New System.Drawing.Size(631, 41)
        Me.txtDirectory.TabIndex = 23
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Directory:"
        '
        'txtFileName
        '
        Me.txtFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFileName.Location = New System.Drawing.Point(73, 40)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(631, 20)
        Me.txtFileName.TabIndex = 21
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "File name:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(656, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 19
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'frmNewXml
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(716, 454)
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
        Me.Name = "frmNewXml"
        Me.Text = "New Extensible Markup Language (XML) Document"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chkOpenInNewWindow As CheckBox
    Friend WithEvents btnCreate As Button
    Friend WithEvents txtDescription As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents chkAddToLibrary As CheckBox
    Friend WithEvents btnFind As Button
    Friend WithEvents txtDirectory As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtFileName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnExit As Button
End Class
