Public Class frmEditRtf
    'TForm used to edit rich text format text.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <!---->
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & ".xml"

        If Main.Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Main.Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            'Add code to read other saved setting here:
            If Settings.<FormSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedTabIndex>.Value

        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message) 'Save the form settings before the form is minimised:
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Display Methods - Code used to display this form." '============================================================================================================================

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        RestoreFormSettings()   'Restore the form settings
        GetTextTypeList()

        RichTextBox2.ZoomFactor = 2
        'Fill RichTextBox2:
        Dim I As Integer
        For I = 32 To 255
            RichTextBox2.AppendText(Chr(I) & vbCrLf)
        Next

        rbSelect.Checked = True

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub GetTextTypeList()
        'Update the list of Text Types in cmbTextType:

        Dim I As Integer

        cmbTextType.Items.Clear()
        cmbTextType.Items.Add("XML")

        For I = 0 To Main.XmlHtmDisplay1.Settings.TextType.Count - 1
            cmbTextType.Items.Add(Main.XmlHtmDisplay1.Settings.TextType.Keys(I))
        Next

    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click

        If cmbTextType.Text = "XML" Then
            Main.XmlHtmDisplay1.SelectedRtf = Main.XmlHtmDisplay1.XmlToRtf(txtNewText.Text, False)
        Else
            Main.XmlHtmDisplay1.SelectedRtf = Main.XmlHtmDisplay1.TextToRtf(txtNewText.Text, cmbTextType.Text)
        End If
    End Sub

    Private Sub btnUndo_Click(sender As Object, e As EventArgs) Handles btnUndo.Click
        RaiseEvent Undo()
    End Sub

    Private Sub btnRedo_Click(sender As Object, e As EventArgs) Handles btnRedo.Click
        RaiseEvent Redo()
    End Sub

    Private Sub btnBold_Click(sender As Object, e As EventArgs) Handles btnBold.Click
        RaiseEvent Bold()
    End Sub

    Private Sub btnItalic_Click(sender As Object, e As EventArgs) Handles btnItalic.Click
        RaiseEvent Italic()
    End Sub

    Private Sub btnUnderline_Click(sender As Object, e As EventArgs) Handles btnUnderline.Click
        RaiseEvent Underline()
    End Sub

    Private Sub btnAlighLeft_Click(sender As Object, e As EventArgs) Handles btnAlighLeft.Click
        RaiseEvent AlignLeft()
    End Sub

    Private Sub btnAlignCenter_Click(sender As Object, e As EventArgs) Handles btnAlignCenter.Click
        RaiseEvent AlignCenter()
    End Sub

    Private Sub btnAlignRight_Click(sender As Object, e As EventArgs) Handles btnAlignRight.Click
        RaiseEvent AlignRight()
    End Sub

    Private Sub btnFont_Click(sender As Object, e As EventArgs) Handles btnFont.Click
        RaiseEvent SelectFont()
    End Sub

    Private Sub btnIncrSize_Click(sender As Object, e As EventArgs) Handles btnIncrSize.Click
        RaiseEvent IncreaseFontSize()
    End Sub

    Private Sub btnDecrSize_Click(sender As Object, e As EventArgs) Handles btnDecrSize.Click
        RaiseEvent DecreaseFontSize()
    End Sub

    Private Sub btnTextColor_Click(sender As Object, e As EventArgs) Handles btnTextColor.Click
        RaiseEvent SelectTextColor()
    End Sub

    Private Sub btnBackgroundColor_Click(sender As Object, e As EventArgs) Handles btnBackgroundColor.Click
        RaiseEvent SelectBackgroundColor()
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        RaiseEvent Copy()
    End Sub

    Private Sub btnCut_Click(sender As Object, e As EventArgs) Handles btnCut.Click
        RaiseEvent Cut()
    End Sub

    Private Sub btnPaste_Click(sender As Object, e As EventArgs) Handles btnPaste.Click
        RaiseEvent Paste()
    End Sub

    Private Sub btnHighlight_Click(sender As Object, e As EventArgs) Handles btnHighlight.Click
        RaiseEvent SelectHighlightColor()
    End Sub

    Private Sub btnInsertChar_Click_1(sender As Object, e As EventArgs) Handles btnInsertChar.Click
        'Insert a character with the specified character code.

        If txtCharCode.Text = "" Then
            Main.Message.AddWarning("No character code specified." & vbCrLf)
            Exit Sub
        End If

        If Val(txtCharCode.Text) < 32 Then
            Main.Message.AddWarning("Invalid character code (less than 32)." & vbCrLf)
            Exit Sub
        ElseIf Val(txtCharCode.Text) > 255 Then
            Main.Message.AddWarning("Invalid character code (more than 255)." & vbCrLf)
            Exit Sub
        End If

        Try
            RaiseEvent InsertChar(Val(txtCharCode.Text))
        Catch ex As Exception
            Main.Message.AddWarning("Error: " & ex.Message & vbCrLf)
        End Try

    End Sub


    Private Sub RichTextBox2_Click(sender As Object, e As EventArgs) Handles RichTextBox2.Click
        'Get the line number and column number of the cursor position:
        Dim LineNo As Integer = RichTextBox2.GetLineFromCharIndex(RichTextBox2.SelectionStart) + 1
        Dim CharIndex As Integer = RichTextBox2.SelectionStart
        Dim LineStartIndex As Integer = RichTextBox2.GetFirstCharIndexOfCurrentLine
        Dim ColNo As Integer = CharIndex - LineStartIndex

        RichTextBox2.SelectionStart = LineStartIndex
        RichTextBox2.SelectionLength = 1
        Try
            txtCharCode.Text = Asc(RichTextBox2.SelectedText.Chars(0))
        Catch ex As Exception
            Main.Message.AddWarning("Error: " & ex.Message & vbCrLf)
        End Try

    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Event Message(ByVal Msg As String)
    Event ErrorMessage(ByVal Msg As String)

    Event Undo()
    Event Redo()
    Event Bold()
    Event Italic()
    Event Underline()
    Event AlignLeft()
    Event AlignCenter()
    Event AlignRight()
    Event SelectFont()
    Event IncreaseFontSize()
    Event DecreaseFontSize()
    Event SelectTextColor()
    Event SelectBackgroundColor()
    Event SelectHighlightColor()
    Event Copy()
    Event Cut()
    Event Paste()

    Event InsertChar(ByVal CharCode As Integer)

    Private Sub rbSelect_CheckedChanged(sender As Object, e As EventArgs) Handles rbSelect.CheckedChanged
        'Select symbol from the list.
        RichTextBox2.Focus()
    End Sub

    Private Sub rbTrademark_CheckedChanged(sender As Object, e As EventArgs) Handles rbTrademark.CheckedChanged
        'Trademark symbol selected.
        'Character code is 153
        txtCharCode.Text = "153"
        RichTextBox2.SelectionStart = RichTextBox2.Find(Chr(153), 0)
        RichTextBox2.SelectionLength = 1
        RichTextBox2.Focus()
    End Sub



    Private Sub rbRegTrademark_CheckedChanged(sender As Object, e As EventArgs) Handles rbRegTrademark.CheckedChanged
        'Registered Trademark symbol selected.
        'Character code is 174
        txtCharCode.Text = "174"
        RichTextBox2.SelectionStart = RichTextBox2.Find(Chr(174), 0)
        RichTextBox2.SelectionLength = 1
        RichTextBox2.Focus()
    End Sub

    Private Sub rbCopyright_CheckedChanged(sender As Object, e As EventArgs) Handles rbCopyright.CheckedChanged
        'Copyright symbol selected.
        'Character code is 169
        txtCharCode.Text = "169"
        RichTextBox2.SelectionStart = RichTextBox2.Find(Chr(169), 0)
        RichTextBox2.SelectionLength = 1
        RichTextBox2.Focus()
    End Sub

    Private Sub rbDegree_CheckedChanged(sender As Object, e As EventArgs) Handles rbDegree.CheckedChanged
        'Degree symbol selected.
        'Character code is 176
        txtCharCode.Text = "176"
        RichTextBox2.SelectionStart = RichTextBox2.Find(Chr(176), 0)
        RichTextBox2.SelectionLength = 1
        RichTextBox2.Focus()
    End Sub

    Private Sub rbMinute_CheckedChanged(sender As Object, e As EventArgs) Handles rbMinute.CheckedChanged
        'Minute symbol selected.
        'Character code is 39
        txtCharCode.Text = "39"
        RichTextBox2.SelectionStart = RichTextBox2.Find(Chr(39), 0)
        RichTextBox2.SelectionLength = 1
        RichTextBox2.Focus()
    End Sub

    Private Sub rbSecond_CheckedChanged(sender As Object, e As EventArgs) Handles rbSecond.CheckedChanged
        'TSecond symbol selected.
        'Character code is 34
        txtCharCode.Text = "34"
        RichTextBox2.SelectionStart = RichTextBox2.Find(Chr(34), 0)
        RichTextBox2.SelectionLength = 1
        RichTextBox2.Focus()
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        System.Diagnostics.Process.Start("https://www.fatcow.com")
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://creativecommons.org/licenses/by/3.0/us/")
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        System.Diagnostics.Process.Start("https://www.fatcow.com/free-icons")
    End Sub
End Class