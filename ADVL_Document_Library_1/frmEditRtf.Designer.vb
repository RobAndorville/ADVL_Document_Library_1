<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditRtf
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
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnUnderline = New System.Windows.Forms.Button()
        Me.btnTextColor = New System.Windows.Forms.Button()
        Me.btnCut = New System.Windows.Forms.Button()
        Me.btnPaste = New System.Windows.Forms.Button()
        Me.btnCopy = New System.Windows.Forms.Button()
        Me.btnAlignRight = New System.Windows.Forms.Button()
        Me.btnAlignCenter = New System.Windows.Forms.Button()
        Me.btnAlighLeft = New System.Windows.Forms.Button()
        Me.btnDecrSize = New System.Windows.Forms.Button()
        Me.btnIncrSize = New System.Windows.Forms.Button()
        Me.btnFont = New System.Windows.Forms.Button()
        Me.btnBold = New System.Windows.Forms.Button()
        Me.btnItalic = New System.Windows.Forms.Button()
        Me.btnBackgroundColor = New System.Windows.Forms.Button()
        Me.btnRedo = New System.Windows.Forms.Button()
        Me.btnUndo = New System.Windows.Forms.Button()
        Me.btnHighlight = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.txtNewText = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbTextType = New System.Windows.Forms.ComboBox()
        Me.btnInsert = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.rbSelect = New System.Windows.Forms.RadioButton()
        Me.rbSecond = New System.Windows.Forms.RadioButton()
        Me.rbMinute = New System.Windows.Forms.RadioButton()
        Me.rbDegree = New System.Windows.Forms.RadioButton()
        Me.rbCopyright = New System.Windows.Forms.RadioButton()
        Me.rbRegTrademark = New System.Windows.Forms.RadioButton()
        Me.rbTrademark = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
        Me.txtCharCode = New System.Windows.Forms.TextBox()
        Me.btnInsertChar = New System.Windows.Forms.Button()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(381, 12)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(48, 22)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnUnderline
        '
        Me.btnUnderline.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.text_underline
        Me.btnUnderline.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnUnderline.Location = New System.Drawing.Point(164, 12)
        Me.btnUnderline.Name = "btnUnderline"
        Me.btnUnderline.Size = New System.Drawing.Size(32, 32)
        Me.btnUnderline.TabIndex = 24
        Me.ToolTip1.SetToolTip(Me.btnUnderline, "Underline selected text")
        Me.btnUnderline.UseVisualStyleBackColor = True
        '
        'btnTextColor
        '
        Me.btnTextColor.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.font_colors
        Me.btnTextColor.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnTextColor.Location = New System.Drawing.Point(126, 50)
        Me.btnTextColor.Name = "btnTextColor"
        Me.btnTextColor.Size = New System.Drawing.Size(32, 32)
        Me.btnTextColor.TabIndex = 23
        Me.ToolTip1.SetToolTip(Me.btnTextColor, "Select color of selected text")
        Me.btnTextColor.UseVisualStyleBackColor = True
        '
        'btnCut
        '
        Me.btnCut.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.cut
        Me.btnCut.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCut.Location = New System.Drawing.Point(290, 50)
        Me.btnCut.Name = "btnCut"
        Me.btnCut.Size = New System.Drawing.Size(32, 32)
        Me.btnCut.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.btnCut, "Cut")
        Me.btnCut.UseVisualStyleBackColor = True
        '
        'btnPaste
        '
        Me.btnPaste.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.page_paste
        Me.btnPaste.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnPaste.Location = New System.Drawing.Point(328, 50)
        Me.btnPaste.Name = "btnPaste"
        Me.btnPaste.Size = New System.Drawing.Size(32, 32)
        Me.btnPaste.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.btnPaste, "Paste")
        Me.btnPaste.UseVisualStyleBackColor = True
        '
        'btnCopy
        '
        Me.btnCopy.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.page_copy
        Me.btnCopy.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCopy.Location = New System.Drawing.Point(252, 50)
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.Size = New System.Drawing.Size(32, 32)
        Me.btnCopy.TabIndex = 20
        Me.ToolTip1.SetToolTip(Me.btnCopy, "Copy")
        Me.btnCopy.UseVisualStyleBackColor = True
        '
        'btnAlignRight
        '
        Me.btnAlignRight.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.text_align_right
        Me.btnAlignRight.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnAlignRight.Location = New System.Drawing.Point(290, 12)
        Me.btnAlignRight.Name = "btnAlignRight"
        Me.btnAlignRight.Size = New System.Drawing.Size(32, 32)
        Me.btnAlignRight.TabIndex = 19
        Me.ToolTip1.SetToolTip(Me.btnAlignRight, "Right justify selected text")
        Me.btnAlignRight.UseVisualStyleBackColor = True
        '
        'btnAlignCenter
        '
        Me.btnAlignCenter.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.text_align_center
        Me.btnAlignCenter.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnAlignCenter.Location = New System.Drawing.Point(252, 12)
        Me.btnAlignCenter.Name = "btnAlignCenter"
        Me.btnAlignCenter.Size = New System.Drawing.Size(32, 32)
        Me.btnAlignCenter.TabIndex = 18
        Me.ToolTip1.SetToolTip(Me.btnAlignCenter, "Center selected text")
        Me.btnAlignCenter.UseVisualStyleBackColor = True
        '
        'btnAlighLeft
        '
        Me.btnAlighLeft.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.text_align_left
        Me.btnAlighLeft.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnAlighLeft.Location = New System.Drawing.Point(214, 12)
        Me.btnAlighLeft.Name = "btnAlighLeft"
        Me.btnAlighLeft.Size = New System.Drawing.Size(32, 32)
        Me.btnAlighLeft.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me.btnAlighLeft, "Left justify selected text")
        Me.btnAlighLeft.UseVisualStyleBackColor = True
        '
        'btnDecrSize
        '
        Me.btnDecrSize.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.bullet_toggle_minus
        Me.btnDecrSize.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnDecrSize.Location = New System.Drawing.Point(88, 50)
        Me.btnDecrSize.Name = "btnDecrSize"
        Me.btnDecrSize.Size = New System.Drawing.Size(32, 32)
        Me.btnDecrSize.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.btnDecrSize, "Decrease font size of selected text")
        Me.btnDecrSize.UseVisualStyleBackColor = True
        '
        'btnIncrSize
        '
        Me.btnIncrSize.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.bullet_toggle_plus
        Me.btnIncrSize.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnIncrSize.Location = New System.Drawing.Point(50, 50)
        Me.btnIncrSize.Name = "btnIncrSize"
        Me.btnIncrSize.Size = New System.Drawing.Size(32, 32)
        Me.btnIncrSize.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.btnIncrSize, "Increase font size of selected text")
        Me.btnIncrSize.UseVisualStyleBackColor = True
        '
        'btnFont
        '
        Me.btnFont.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.font
        Me.btnFont.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnFont.Location = New System.Drawing.Point(12, 50)
        Me.btnFont.Name = "btnFont"
        Me.btnFont.Size = New System.Drawing.Size(32, 32)
        Me.btnFont.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.btnFont, "Select font for selected text")
        Me.btnFont.UseVisualStyleBackColor = True
        '
        'btnBold
        '
        Me.btnBold.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.text_bold
        Me.btnBold.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBold.Location = New System.Drawing.Point(88, 12)
        Me.btnBold.Name = "btnBold"
        Me.btnBold.Size = New System.Drawing.Size(32, 32)
        Me.btnBold.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.btnBold, "Make selected text bold")
        Me.btnBold.UseVisualStyleBackColor = True
        '
        'btnItalic
        '
        Me.btnItalic.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.text_italic
        Me.btnItalic.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnItalic.Location = New System.Drawing.Point(126, 12)
        Me.btnItalic.Name = "btnItalic"
        Me.btnItalic.Size = New System.Drawing.Size(32, 32)
        Me.btnItalic.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.btnItalic, "Make selected text italic")
        Me.btnItalic.UseVisualStyleBackColor = True
        '
        'btnBackgroundColor
        '
        Me.btnBackgroundColor.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.color_picker_alternative
        Me.btnBackgroundColor.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBackgroundColor.Location = New System.Drawing.Point(164, 50)
        Me.btnBackgroundColor.Name = "btnBackgroundColor"
        Me.btnBackgroundColor.Size = New System.Drawing.Size(32, 32)
        Me.btnBackgroundColor.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.btnBackgroundColor, "Select background color")
        Me.btnBackgroundColor.UseVisualStyleBackColor = True
        '
        'btnRedo
        '
        Me.btnRedo.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.arrow_redo
        Me.btnRedo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnRedo.Location = New System.Drawing.Point(50, 12)
        Me.btnRedo.Name = "btnRedo"
        Me.btnRedo.Size = New System.Drawing.Size(32, 32)
        Me.btnRedo.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.btnRedo, "Redo")
        Me.btnRedo.UseVisualStyleBackColor = True
        '
        'btnUndo
        '
        Me.btnUndo.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.arrow_undo
        Me.btnUndo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnUndo.Location = New System.Drawing.Point(12, 12)
        Me.btnUndo.Name = "btnUndo"
        Me.btnUndo.Size = New System.Drawing.Size(32, 32)
        Me.btnUndo.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.btnUndo, "Undo")
        Me.btnUndo.UseVisualStyleBackColor = True
        '
        'btnHighlight
        '
        Me.btnHighlight.BackgroundImage = Global.ADVL_Document_Library_1.My.Resources.Resources.highlighter_text
        Me.btnHighlight.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnHighlight.Location = New System.Drawing.Point(202, 50)
        Me.btnHighlight.Name = "btnHighlight"
        Me.btnHighlight.Size = New System.Drawing.Size(32, 32)
        Me.btnHighlight.TabIndex = 26
        Me.ToolTip1.SetToolTip(Me.btnHighlight, "Select highlight color of selected text")
        Me.btnHighlight.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(12, 88)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(417, 274)
        Me.TabControl1.TabIndex = 27
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtNewText)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.cmbTextType)
        Me.TabPage1.Controls.Add(Me.btnInsert)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(409, 248)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Insert Text"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'txtNewText
        '
        Me.txtNewText.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNewText.Location = New System.Drawing.Point(6, 35)
        Me.txtNewText.Multiline = True
        Me.txtNewText.Name = "txtNewText"
        Me.txtNewText.Size = New System.Drawing.Size(397, 207)
        Me.txtNewText.TabIndex = 20
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(34, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Type:"
        '
        'cmbTextType
        '
        Me.cmbTextType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbTextType.FormattingEnabled = True
        Me.cmbTextType.Location = New System.Drawing.Point(46, 8)
        Me.cmbTextType.Name = "cmbTextType"
        Me.cmbTextType.Size = New System.Drawing.Size(303, 21)
        Me.cmbTextType.TabIndex = 18
        '
        'btnInsert
        '
        Me.btnInsert.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInsert.Location = New System.Drawing.Point(355, 7)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(48, 22)
        Me.btnInsert.TabIndex = 17
        Me.btnInsert.Text = "Insert"
        Me.btnInsert.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.rbSelect)
        Me.TabPage2.Controls.Add(Me.rbSecond)
        Me.TabPage2.Controls.Add(Me.rbMinute)
        Me.TabPage2.Controls.Add(Me.rbDegree)
        Me.TabPage2.Controls.Add(Me.rbCopyright)
        Me.TabPage2.Controls.Add(Me.rbRegTrademark)
        Me.TabPage2.Controls.Add(Me.rbTrademark)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.RichTextBox2)
        Me.TabPage2.Controls.Add(Me.txtCharCode)
        Me.TabPage2.Controls.Add(Me.btnInsertChar)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(409, 248)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Insert Symbol"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'rbSelect
        '
        Me.rbSelect.AutoSize = True
        Me.rbSelect.Location = New System.Drawing.Point(71, 29)
        Me.rbSelect.Name = "rbSelect"
        Me.rbSelect.Size = New System.Drawing.Size(93, 17)
        Me.rbSelect.TabIndex = 35
        Me.rbSelect.TabStop = True
        Me.rbSelect.Text = "Select from list"
        Me.rbSelect.UseVisualStyleBackColor = True
        '
        'rbSecond
        '
        Me.rbSecond.AutoSize = True
        Me.rbSecond.Location = New System.Drawing.Point(71, 167)
        Me.rbSecond.Name = "rbSecond"
        Me.rbSecond.Size = New System.Drawing.Size(62, 17)
        Me.rbSecond.TabIndex = 34
        Me.rbSecond.TabStop = True
        Me.rbSecond.Text = "Second"
        Me.rbSecond.UseVisualStyleBackColor = True
        '
        'rbMinute
        '
        Me.rbMinute.AutoSize = True
        Me.rbMinute.Location = New System.Drawing.Point(71, 144)
        Me.rbMinute.Name = "rbMinute"
        Me.rbMinute.Size = New System.Drawing.Size(57, 17)
        Me.rbMinute.TabIndex = 33
        Me.rbMinute.TabStop = True
        Me.rbMinute.Text = "Minute"
        Me.rbMinute.UseVisualStyleBackColor = True
        '
        'rbDegree
        '
        Me.rbDegree.AutoSize = True
        Me.rbDegree.Location = New System.Drawing.Point(71, 121)
        Me.rbDegree.Name = "rbDegree"
        Me.rbDegree.Size = New System.Drawing.Size(60, 17)
        Me.rbDegree.TabIndex = 32
        Me.rbDegree.TabStop = True
        Me.rbDegree.Text = "Degree"
        Me.rbDegree.UseVisualStyleBackColor = True
        '
        'rbCopyright
        '
        Me.rbCopyright.AutoSize = True
        Me.rbCopyright.Location = New System.Drawing.Point(71, 98)
        Me.rbCopyright.Name = "rbCopyright"
        Me.rbCopyright.Size = New System.Drawing.Size(69, 17)
        Me.rbCopyright.TabIndex = 31
        Me.rbCopyright.TabStop = True
        Me.rbCopyright.Text = "Copyright"
        Me.rbCopyright.UseVisualStyleBackColor = True
        '
        'rbRegTrademark
        '
        Me.rbRegTrademark.AutoSize = True
        Me.rbRegTrademark.Location = New System.Drawing.Point(71, 75)
        Me.rbRegTrademark.Name = "rbRegTrademark"
        Me.rbRegTrademark.Size = New System.Drawing.Size(130, 17)
        Me.rbRegTrademark.TabIndex = 30
        Me.rbRegTrademark.TabStop = True
        Me.rbRegTrademark.Text = "Registered Trademark"
        Me.rbRegTrademark.UseVisualStyleBackColor = True
        '
        'rbTrademark
        '
        Me.rbTrademark.AutoSize = True
        Me.rbTrademark.Location = New System.Drawing.Point(71, 52)
        Me.rbTrademark.Name = "rbTrademark"
        Me.rbTrademark.Size = New System.Drawing.Size(76, 17)
        Me.rbTrademark.TabIndex = 29
        Me.rbTrademark.TabStop = True
        Me.rbTrademark.Text = "Trademark"
        Me.rbTrademark.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(71, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Character code:"
        '
        'RichTextBox2
        '
        Me.RichTextBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox2.Location = New System.Drawing.Point(6, 6)
        Me.RichTextBox2.Name = "RichTextBox2"
        Me.RichTextBox2.Size = New System.Drawing.Size(59, 236)
        Me.RichTextBox2.TabIndex = 27
        Me.RichTextBox2.Text = ""
        '
        'txtCharCode
        '
        Me.txtCharCode.Location = New System.Drawing.Point(160, 8)
        Me.txtCharCode.Name = "txtCharCode"
        Me.txtCharCode.Size = New System.Drawing.Size(37, 20)
        Me.txtCharCode.TabIndex = 26
        '
        'btnInsertChar
        '
        Me.btnInsertChar.Location = New System.Drawing.Point(203, 6)
        Me.btnInsertChar.Name = "btnInsertChar"
        Me.btnInsertChar.Size = New System.Drawing.Size(48, 22)
        Me.btnInsertChar.TabIndex = 24
        Me.btnInsertChar.Text = "Insert"
        Me.btnInsertChar.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.LinkLabel3)
        Me.TabPage3.Controls.Add(Me.Label6)
        Me.TabPage3.Controls.Add(Me.LinkLabel2)
        Me.TabPage3.Controls.Add(Me.LinkLabel1)
        Me.TabPage3.Controls.Add(Me.Label5)
        Me.TabPage3.Controls.Add(Me.Label4)
        Me.TabPage3.Controls.Add(Me.Label3)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(409, 248)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Icon License"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(173, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "The icons shown on this page are: "
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 41)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "All rights reserved."
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(366, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "These icons are licensed under a Creative Commons Attribution 3.0 License."
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(13, 101)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(239, 13)
        Me.LinkLabel1.TabIndex = 3
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "http://creativecommons.org/licenses/by/3.0/us/"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(13, 54)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(125, 13)
        Me.LinkLabel2.TabIndex = 4
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "https://www.fatcow.com"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(13, 28)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(224, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "© Copyright 2009-2014 FatCow Web Hosting."
        '
        'LinkLabel3
        '
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.Location = New System.Drawing.Point(13, 67)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(176, 13)
        Me.LinkLabel3.TabIndex = 6
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "https://www.fatcow.com/free-icons"
        '
        'frmEditRtf
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(441, 374)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnHighlight)
        Me.Controls.Add(Me.btnPaste)
        Me.Controls.Add(Me.btnCut)
        Me.Controls.Add(Me.btnTextColor)
        Me.Controls.Add(Me.btnUnderline)
        Me.Controls.Add(Me.btnCopy)
        Me.Controls.Add(Me.btnAlignRight)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnAlignCenter)
        Me.Controls.Add(Me.btnUndo)
        Me.Controls.Add(Me.btnAlighLeft)
        Me.Controls.Add(Me.btnRedo)
        Me.Controls.Add(Me.btnBackgroundColor)
        Me.Controls.Add(Me.btnBold)
        Me.Controls.Add(Me.btnItalic)
        Me.Controls.Add(Me.btnDecrSize)
        Me.Controls.Add(Me.btnFont)
        Me.Controls.Add(Me.btnIncrSize)
        Me.Name = "frmEditRtf"
        Me.Text = "Edit Rtf"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnUndo As Button
    Friend WithEvents btnBackgroundColor As Button
    Friend WithEvents btnRedo As Button
    Friend WithEvents btnCut As Button
    Friend WithEvents btnPaste As Button
    Friend WithEvents btnCopy As Button
    Friend WithEvents btnAlignRight As Button
    Friend WithEvents btnAlignCenter As Button
    Friend WithEvents btnAlighLeft As Button
    Friend WithEvents btnDecrSize As Button
    Friend WithEvents btnIncrSize As Button
    Friend WithEvents btnFont As Button
    Friend WithEvents btnBold As Button
    Friend WithEvents btnItalic As Button
    Friend WithEvents btnUnderline As Button
    Friend WithEvents btnTextColor As Button
    Friend WithEvents btnHighlight As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents txtNewText As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents cmbTextType As ComboBox
    Friend WithEvents btnInsert As Button
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents txtCharCode As TextBox
    Friend WithEvents btnInsertChar As Button
    Friend WithEvents RichTextBox2 As RichTextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents rbSelect As RadioButton
    Friend WithEvents rbSecond As RadioButton
    Friend WithEvents rbMinute As RadioButton
    Friend WithEvents rbDegree As RadioButton
    Friend WithEvents rbCopyright As RadioButton
    Friend WithEvents rbRegTrademark As RadioButton
    Friend WithEvents rbTrademark As RadioButton
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents Label6 As Label
    Friend WithEvents LinkLabel2 As LinkLabel
    Friend WithEvents LinkLabel1 As LinkLabel
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents LinkLabel3 As LinkLabel
End Class
