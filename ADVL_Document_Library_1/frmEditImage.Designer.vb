<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditImage
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
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnPaste = New System.Windows.Forms.Button()
        Me.btnResize = New System.Windows.Forms.Button()
        Me.txtResize = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnShowImageSettings = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.rbNormal = New System.Windows.Forms.RadioButton()
        Me.rbStretchImage = New System.Windows.Forms.RadioButton()
        Me.rbAutoSize = New System.Windows.Forms.RadioButton()
        Me.rbCenterImage = New System.Windows.Forms.RadioButton()
        Me.rbZoom = New System.Windows.Forms.RadioButton()
        Me.btnInsert = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(681, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(12, 70)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(717, 460)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 9
        Me.PictureBox1.TabStop = False
        '
        'btnPaste
        '
        Me.btnPaste.Location = New System.Drawing.Point(12, 12)
        Me.btnPaste.Name = "btnPaste"
        Me.btnPaste.Size = New System.Drawing.Size(121, 22)
        Me.btnPaste.TabIndex = 10
        Me.btnPaste.Text = "Paste From Clipboard"
        Me.btnPaste.UseVisualStyleBackColor = True
        '
        'btnResize
        '
        Me.btnResize.Location = New System.Drawing.Point(266, 12)
        Me.btnResize.Name = "btnResize"
        Me.btnResize.Size = New System.Drawing.Size(48, 22)
        Me.btnResize.TabIndex = 11
        Me.btnResize.Text = "Resize"
        Me.btnResize.UseVisualStyleBackColor = True
        '
        'txtResize
        '
        Me.txtResize.Location = New System.Drawing.Point(320, 14)
        Me.txtResize.Name = "txtResize"
        Me.txtResize.Size = New System.Drawing.Size(48, 20)
        Me.txtResize.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(374, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(15, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "%"
        '
        'btnShowImageSettings
        '
        Me.btnShowImageSettings.Location = New System.Drawing.Point(395, 12)
        Me.btnShowImageSettings.Name = "btnShowImageSettings"
        Me.btnShowImageSettings.Size = New System.Drawing.Size(124, 22)
        Me.btnShowImageSettings.TabIndex = 14
        Me.btnShowImageSettings.Text = "Show Image Settings"
        Me.btnShowImageSettings.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "SizeMode:"
        '
        'rbNormal
        '
        Me.rbNormal.AutoSize = True
        Me.rbNormal.Location = New System.Drawing.Point(75, 40)
        Me.rbNormal.Name = "rbNormal"
        Me.rbNormal.Size = New System.Drawing.Size(58, 17)
        Me.rbNormal.TabIndex = 16
        Me.rbNormal.TabStop = True
        Me.rbNormal.Text = "Normal"
        Me.rbNormal.UseVisualStyleBackColor = True
        '
        'rbStretchImage
        '
        Me.rbStretchImage.AutoSize = True
        Me.rbStretchImage.Location = New System.Drawing.Point(139, 40)
        Me.rbStretchImage.Name = "rbStretchImage"
        Me.rbStretchImage.Size = New System.Drawing.Size(91, 17)
        Me.rbStretchImage.TabIndex = 17
        Me.rbStretchImage.TabStop = True
        Me.rbStretchImage.Text = "Stretch Image"
        Me.rbStretchImage.UseVisualStyleBackColor = True
        '
        'rbAutoSize
        '
        Me.rbAutoSize.AutoSize = True
        Me.rbAutoSize.Location = New System.Drawing.Point(236, 40)
        Me.rbAutoSize.Name = "rbAutoSize"
        Me.rbAutoSize.Size = New System.Drawing.Size(70, 17)
        Me.rbAutoSize.TabIndex = 18
        Me.rbAutoSize.TabStop = True
        Me.rbAutoSize.Text = "Auto Size"
        Me.rbAutoSize.UseVisualStyleBackColor = True
        '
        'rbCenterImage
        '
        Me.rbCenterImage.AutoSize = True
        Me.rbCenterImage.Location = New System.Drawing.Point(312, 40)
        Me.rbCenterImage.Name = "rbCenterImage"
        Me.rbCenterImage.Size = New System.Drawing.Size(88, 17)
        Me.rbCenterImage.TabIndex = 19
        Me.rbCenterImage.TabStop = True
        Me.rbCenterImage.Text = "Center Image"
        Me.rbCenterImage.UseVisualStyleBackColor = True
        '
        'rbZoom
        '
        Me.rbZoom.AutoSize = True
        Me.rbZoom.Location = New System.Drawing.Point(406, 40)
        Me.rbZoom.Name = "rbZoom"
        Me.rbZoom.Size = New System.Drawing.Size(52, 17)
        Me.rbZoom.TabIndex = 20
        Me.rbZoom.TabStop = True
        Me.rbZoom.Text = "Zoom"
        Me.rbZoom.UseVisualStyleBackColor = True
        '
        'btnInsert
        '
        Me.btnInsert.Location = New System.Drawing.Point(139, 12)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(121, 22)
        Me.btnInsert.TabIndex = 21
        Me.btnInsert.Text = "Insert In Document"
        Me.btnInsert.UseVisualStyleBackColor = True
        '
        'frmEditImage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(741, 542)
        Me.Controls.Add(Me.btnInsert)
        Me.Controls.Add(Me.rbZoom)
        Me.Controls.Add(Me.rbCenterImage)
        Me.Controls.Add(Me.rbAutoSize)
        Me.Controls.Add(Me.rbStretchImage)
        Me.Controls.Add(Me.rbNormal)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnShowImageSettings)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtResize)
        Me.Controls.Add(Me.btnResize)
        Me.Controls.Add(Me.btnPaste)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btnExit)
        Me.Name = "frmEditImage"
        Me.Text = "Edit Image"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btnPaste As Button
    Friend WithEvents btnResize As Button
    Friend WithEvents txtResize As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnShowImageSettings As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents rbNormal As RadioButton
    Friend WithEvents rbStretchImage As RadioButton
    Friend WithEvents rbAutoSize As RadioButton
    Friend WithEvents rbCenterImage As RadioButton
    Friend WithEvents rbZoom As RadioButton
    Friend WithEvents btnInsert As Button
End Class
