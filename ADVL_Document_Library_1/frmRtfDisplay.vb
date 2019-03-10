Public Class frmRtfDisplay
    'Form used to display a rich text format (RTF) document.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    'Declare Forms used by the application:
    Public WithEvents EditRtf As frmEditRtf

    <Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="GetCursor")>
    Private Shared Function GetCursor() As IntPtr
    End Function

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _formNo As Integer = 0 'Multiple instances of this form can be displayed. FormNo is the index number of the form in RtfDisplayFormList.
    Public Property FormNo As Integer
        Get
            Return _formNo
        End Get
        Set(ByVal value As Integer)
            _formNo = value
        End Set
    End Property

    Private _libraryName As String = "" 'The name of the document library.
    Public Property LibraryName As String
        Get
            Return _libraryName
        End Get
        Set(value As String)
            _libraryName = value
        End Set
    End Property


    Private _libraryDocNo As Integer = -1 'The entry number in the document library.
    Public Property LibraryDocNo As Integer
        Get
            Return _libraryDocNo
        End Get
        Set(value As Integer)
            _libraryDocNo = value
        End Set
    End Property

    Private _useSavedSettings As Boolean = False 'If True, read the saved form settings.
    Public Property UseSavedSettings As Boolean
        Get
            Return _useSavedSettings
        End Get
        Set(value As Boolean)
            _useSavedSettings = value
        End Set
    End Property

    Private _fileName As String = "" 'The file name of the displayed document.
    Public Property FileName As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
            txtFileName.Text = _fileName
        End Set
    End Property

    Private _description As String = "" 'A description of the displayed document.
    Public Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Enum LocationTypes
        Project
        FileSystem
    End Enum

    Private _fileLocationType As LocationTypes = LocationTypes.Project 'The location type of the Document File. (Either the current project or the file system.)
    Property FileLocationType As LocationTypes
        Get
            Return _fileLocationType
        End Get
        Set(value As LocationTypes)
            _fileLocationType = value
        End Set
    End Property

    Private _fileDirectory As String = "" 'The path of the directory containing the current file.
    Property FileDirectory As String
        Get
            Return _fileDirectory
        End Get
        Set(value As String)
            _fileDirectory = value
        End Set
    End Property

    Private _docTextChanged As Boolean = False 'If True, the document text has been changed. Prompt to save the changes before they are lost.
    Property DocTextChanged As Boolean
        Get
            Return _docTextChanged
        End Get
        Set(value As Boolean)
            _docTextChanged = value
        End Set
    End Property

    Private _lastFileName As String = "" 'The name of the current file.
    Property LastFileName As String
        Get
            Return _lastFileName
        End Get
        Set(value As String)
            _lastFileName = value
        End Set
    End Property

    Private _lastFileLocationType As LocationTypes = LocationTypes.Project 'The location type of the Document File. (Either the current project or the file system.)
    Property LastFileLocationType As LocationTypes
        Get
            Return _lastFileLocationType
        End Get
        Set(value As LocationTypes)
            _lastFileLocationType = value
        End Set
    End Property

    Private _lastFileDirectory As String = "" 'The path of the directory containing the current file.
    Property LastFileDirectory As String
        Get
            Return _lastFileDirectory
        End Get
        Set(value As String)
            _lastFileDirectory = value
        End Set
    End Property

    Private _zoomFactor As Single = 1 'The RichTextBox zoom factor
    Property ZoomFactor As Single
        Get
            Return _zoomFactor
        End Get
        Set(value As Single)
            _zoomFactor = value
            XmlHtmDisplay1.ZoomFactor = _zoomFactor
            If txtZoomPercent.Focused Then
                'Dont change txtZoomPercent while it is being edited.
            Else
                txtZoomPercent.Text = Int(_zoomFactor * 100)
            End If
        End Set
    End Property

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
                               <LastFileName><%= LastFileName %></LastFileName>
                               <LastFileLocationType><%= LastFileLocationType %></LastFileLocationType>
                               <LastFileDirectory><%= LastFileDirectory %></LastFileDirectory>
                               <ZoomFactor><%= ZoomFactor %></ZoomFactor>
                               <WordWrap><%= chkWordWrap.Checked %></WordWrap>
                               <!---->
                           </FormSettings>

        'Add code to include other settings to save after the comment line <!---->

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"
        Main.Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        Dim SettingsFileName As String = "FormSettings_" & Main.ApplicationInfo.Name & "_" & Me.Text & "_" & FormNo & ".xml"

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
            If Settings.<FormSettings>.<LastFileName>.Value <> Nothing Then LastFileName = Settings.<FormSettings>.<LastFileName>.Value
            If Settings.<FormSettings>.<LastFileLocationType>.Value <> Nothing Then
                Select Case Settings.<FormSettings>.<LastFileLocationType>.Value
                    Case "FileSystem"
                        LastFileLocationType = LocationTypes.FileSystem
                    Case "Project"
                        LastFileLocationType = LocationTypes.Project
                End Select
            End If
            If Settings.<FormSettings>.<LastFileDirectory>.Value <> Nothing Then LastFileDirectory = Settings.<FormSettings>.<LastFileDirectory>.Value
            If Settings.<FormSettings>.<ZoomFactor>.Value <> Nothing Then
                ZoomFactor = Settings.<FormSettings>.<ZoomFactor>.Value
            End If
            If Settings.<FormSettings>.<WordWrap>.Value <> Nothing Then
                chkWordWrap.Checked = Settings.<FormSettings>.<WordWrap>.Value
                If chkWordWrap.Checked Then
                    XmlHtmDisplay1.WordWrap = True
                Else
                    XmlHtmDisplay1.WordWrap = False
                End If
            End If
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

        XmlHtmDisplay1.DetectUrls = True

    End Sub


    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form

        If DocTextChanged = True Then
            Dim result As Integer = MessageBox.Show("Save changes to the current document?", "Notice", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                Exit Sub
            ElseIf result = DialogResult.Yes Then
                SaveDocument()
            ElseIf result = DialogResult.No Then
                'Do not save the changes!
            End If
        End If

        Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the RtfDisplayFormClosed method to select the correct form to set to nothing.

        If FileName <> "" Then
            LastFileName = FileName
        End If

        If IsNothing(EditRtf) Then
            'The EditRtf form is already closed.
        Else
            'Close the EditRtf form:
            EditRtf.Close()
        End If

        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub

    Private Sub frmRtfDisplay_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Main.RtfDisplayFormClosed()
    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click

        If DocTextChanged = True Then
            Dim result As Integer = MessageBox.Show("Save changes to the current document?", "Notice", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                Exit Sub
            ElseIf result = DialogResult.Yes Then
                SaveDocument()
            ElseIf result = DialogResult.No Then
                'Do not save the changes!
            End If
        End If

        If Main.rbFileInProject.Checked = True Then 'Open an RTF file in the current Project.
            Dim SelectedFile As String = Main.Project.SelectDataFile("Rich text format", "rtf")
            If SelectedFile = "" Then
                'No file selected!
            Else
                FileName = SelectedFile
                FileLocationType = LocationTypes.Project
                FileDirectory = ""
                OpenDocument()
            End If

        Else 'Open an RTF file in the File System.
            If LastFileName = "" Then 'There is no last XML file saved.

            Else 'The last RTF file path was saved.
                OpenFileDialog1.InitialDirectory = LastFileDirectory
                OpenFileDialog1.FileName = LastFileName
            End If

            OpenFileDialog1.Filter = "Rich Text Format (*.rtf)| *.rtf"
            SendKeys.Send("{HOME}") 'To move the cursor to the start of the FileName
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                FileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                FileLocationType = LocationTypes.FileSystem
                FileDirectory = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
                OpenDocument()
            End If
        End If

    End Sub



#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If IsNothing(EditRtf) Then
            EditRtf = New frmEditRtf
            EditRtf.Show()
        Else
            EditRtf.Show()
        End If
    End Sub

    Private Sub EditRtf_FormClosed(sender As Object, e As FormClosedEventArgs) Handles EditRtf.FormClosed
        EditRtf = Nothing
    End Sub

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub XmlDisplay1_TextChanged(sender As Object, e As EventArgs)
        DocTextChanged = True
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SaveDocument()
    End Sub

    Public Sub SaveDocument()

        If FileName = "" Then
            Beep()
            'Call the File Save dialog...
        Else
            If FileLocationType = LocationTypes.FileSystem Then 'Save the document at the specified path in the File System.
                If XmlHtmDisplay1.SaveRtfFile(FileDirectory & "\" & FileName) = True Then
                    'File was saved OK.
                    LastFileName = FileName 'Update the LastFilePath.
                    LastFileLocationType = LocationTypes.FileSystem
                    LastFileDirectory = FileDirectory
                    DocTextChanged = False
                End If
            Else 'Save the document in the current project.
                Dim rtbData As New IO.MemoryStream
                XmlHtmDisplay1.SaveFile(rtbData, RichTextBoxStreamType.RichText)
                rtbData.Position = 0
                Main.Project.SaveData(FileName, rtbData)
                LastFileName = FileName 'Update the LastFileName.
                LastFileLocationType = LocationTypes.Project
                LastFileDirectory = ""
                DocTextChanged = False
            End If
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        If DocTextChanged = True Then
            Dim result As Integer = MessageBox.Show("Save changes to the current document?", "Notice", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                Exit Sub
            ElseIf result = DialogResult.Yes Then
                SaveDocument()
            ElseIf result = DialogResult.No Then
                'Do not save the changes!
            End If
        End If

        FileName = ""
        XmlHtmDisplay1.Clear()
        DocTextChanged = False

    End Sub



    Public Sub OpenDocument()
        'Open the document specified by FileName, FileLocationType and FileDirectory.

        If FileLocationType = LocationTypes.Project Then
            Dim rtbData As New IO.MemoryStream
            Main.Project.ReadData(FileName, rtbData)

            If rtbData Is Nothing Then
                Main.Message.AddWarning("Document is empty!" & vbCrLf)
                Exit Sub
            End If

            XmlHtmDisplay1.Clear()
            XmlHtmDisplay1.ZoomFactor = 1 'Set ZoomFactor to 1 after clearing the RichTextBox. (MS bug?)
            rtbData.Position = 0

            Try
                XmlHtmDisplay1.LoadFile(rtbData, RichTextBoxStreamType.RichText)
                XmlHtmDisplay1.ZoomFactor = _zoomFactor
                DocTextChanged = False
                LastFileName = FileName
                LastFileLocationType = LocationTypes.Project
                LastFileDirectory = ""
            Catch ex As Exception
                Main.Message.AddWarning("Error opening document: " & ex.Message & vbCrLf)
            End Try

        Else
            XmlHtmDisplay1.LoadFile(FileDirectory & "\" & FileName)
            XmlHtmDisplay1.ZoomFactor = _zoomFactor
            DocTextChanged = False
            LastFileName = FileName
            LastFileLocationType = LocationTypes.FileSystem
            LastFileDirectory = FileDirectory
        End If

    End Sub

    'Edit Rtf Code =====================================================================================

    Private Sub EditRtf_Undo() Handles EditRtf.Undo
        XmlHtmDisplay1.Undo()
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_Redo() Handles EditRtf.Redo
        XmlHtmDisplay1.Redo()
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_Bold() Handles EditRtf.Bold
        If XmlHtmDisplay1.SelectionFont.Bold = True Then 'No not apply bold

            Dim newStyle As FontStyle = FontStyle.Regular
            If XmlHtmDisplay1.SelectionFont.Italic Then newStyle = newStyle Or FontStyle.Italic
            If XmlHtmDisplay1.SelectionFont.Underline Then newStyle = newStyle Or FontStyle.Underline
            If XmlHtmDisplay1.SelectionFont.Strikeout Then newStyle = newStyle Or FontStyle.Strikeout

            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, newStyle)

        Else 'Apply bold
            Dim newStyle As FontStyle = FontStyle.Regular
            newStyle = newStyle Or FontStyle.Bold
            If XmlHtmDisplay1.SelectionFont.Italic Then newStyle = newStyle Or FontStyle.Italic
            If XmlHtmDisplay1.SelectionFont.Underline Then newStyle = newStyle Or FontStyle.Underline
            If XmlHtmDisplay1.SelectionFont.Strikeout Then newStyle = newStyle Or FontStyle.Strikeout

            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, newStyle)
        End If
        XmlHtmDisplay1.Focus()

    End Sub

    Private Sub EditRtf_Italic() Handles EditRtf.Italic
        If XmlHtmDisplay1.SelectionFont.Italic = True Then 'No not apply italic

            Dim newStyle As FontStyle = FontStyle.Regular
            If XmlHtmDisplay1.SelectionFont.Bold Then newStyle = newStyle Or FontStyle.Bold
            If XmlHtmDisplay1.SelectionFont.Underline Then newStyle = newStyle Or FontStyle.Underline
            If XmlHtmDisplay1.SelectionFont.Strikeout Then newStyle = newStyle Or FontStyle.Strikeout

            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, newStyle)

        Else 'Apply italic

            Dim newStyle As FontStyle = FontStyle.Regular
            newStyle = newStyle Or FontStyle.Italic
            If XmlHtmDisplay1.SelectionFont.Bold Then newStyle = newStyle Or FontStyle.Bold
            If XmlHtmDisplay1.SelectionFont.Underline Then newStyle = newStyle Or FontStyle.Underline
            If XmlHtmDisplay1.SelectionFont.Strikeout Then newStyle = newStyle Or FontStyle.Strikeout

            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, newStyle)
        End If
        XmlHtmDisplay1.Focus()

    End Sub

    Private Sub EditRtf_Underline() Handles EditRtf.Underline
        If XmlHtmDisplay1.SelectionFont.Underline = True Then 'No not apply underline
            Dim newStyle As FontStyle = FontStyle.Regular
            If XmlHtmDisplay1.SelectionFont.Bold Then newStyle = newStyle Or FontStyle.Bold
            If XmlHtmDisplay1.SelectionFont.Italic Then newStyle = newStyle Or FontStyle.Italic
            If XmlHtmDisplay1.SelectionFont.Strikeout Then newStyle = newStyle Or FontStyle.Strikeout

            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, newStyle)
        Else 'Add underline
            Dim newStyle As FontStyle = FontStyle.Regular
            newStyle = newStyle Or FontStyle.Underline
            If XmlHtmDisplay1.SelectionFont.Bold Then newStyle = newStyle Or FontStyle.Bold
            If XmlHtmDisplay1.SelectionFont.Italic Then newStyle = newStyle Or FontStyle.Italic
            If XmlHtmDisplay1.SelectionFont.Strikeout Then newStyle = newStyle Or FontStyle.Strikeout

            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, newStyle)
        End If
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_AlignLeft() Handles EditRtf.AlignLeft
        XmlHtmDisplay1.SelectionAlignment = HorizontalAlignment.Left
    End Sub

    Private Sub EditRtf_AlignCenter() Handles EditRtf.AlignCenter
        XmlHtmDisplay1.SelectionAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub EditRtf_AlignRight() Handles EditRtf.AlignRight
        XmlHtmDisplay1.SelectionAlignment = HorizontalAlignment.Right
    End Sub

    Private Sub EditRtf_SelectFont() Handles EditRtf.SelectFont
        Dim myFontDialog As New FontDialog
        myFontDialog.Font = XmlHtmDisplay1.SelectionFont
        myFontDialog.ShowDialog()
        XmlHtmDisplay1.SelectionFont = myFontDialog.Font
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_IncreaseFontSize() Handles EditRtf.IncreaseFontSize
        Try
            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont.FontFamily, XmlHtmDisplay1.SelectionFont.SizeInPoints + 1, XmlHtmDisplay1.SelectionFont.Style) 'Keep the original font style!
        Catch ex As Exception
        End Try
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_DecreaseFontSize() Handles EditRtf.DecreaseFontSize
        Try
            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont.FontFamily, XmlHtmDisplay1.SelectionFont.SizeInPoints - 1, XmlHtmDisplay1.SelectionFont.Style)
        Catch ex As Exception
        End Try
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_SelectTextColor() Handles EditRtf.SelectTextColor
        Dim colorDialog As New ColorDialog
        colorDialog.Color = XmlHtmDisplay1.SelectionColor
        colorDialog.ShowDialog()
        XmlHtmDisplay1.SelectionColor = colorDialog.Color
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_SelectBackgroundColor() Handles EditRtf.SelectBackgroundColor
        Dim colorDialog As New ColorDialog
        colorDialog.Color = XmlHtmDisplay1.BackColor
        colorDialog.ShowDialog()
        XmlHtmDisplay1.BackColor = colorDialog.Color
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_SelectHighlightColor() Handles EditRtf.SelectHighlightColor
        Dim colorDialog As New ColorDialog
        colorDialog.Color = XmlHtmDisplay1.SelectionBackColor
        colorDialog.ShowDialog()
        XmlHtmDisplay1.SelectionBackColor = colorDialog.Color
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_Copy() Handles EditRtf.Copy
        My.Computer.Clipboard.Clear()
        Try
            Clipboard.SetText(XmlHtmDisplay1.SelectedText)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub EditRtf_Cut() Handles EditRtf.Cut
        My.Computer.Clipboard.Clear()
        Try
            Clipboard.SetText(XmlHtmDisplay1.SelectedText)
            XmlHtmDisplay1.SelectedText = ""
        Catch ex As Exception

        End Try
    End Sub

    Private Sub EditRtf_Paste() Handles EditRtf.Paste
        If My.Computer.Clipboard.ContainsText Then
            XmlHtmDisplay1.Paste()
        End If
    End Sub

    Private Sub EditRtf_InsertChar(CharCode As Integer) Handles EditRtf.InsertChar
        'Insert the character with the specified Character Code in the text.
        XmlHtmDisplay1.SelectedText = Chr(CharCode)

    End Sub

    Private Sub XmlHtmDisplay1_LinkClicked(sender As Object, e As LinkClickedEventArgs) Handles XmlHtmDisplay1.LinkClicked
        txtLink.Text = e.LinkText
    End Sub

    Private Sub XmlHtmDisplay1_CursorChanged(sender As Object, e As EventArgs) Handles XmlHtmDisplay1.CursorChanged
        txtLink.Text = "Cursor Changed"
    End Sub

    Private Sub XmlHtmDisplay1_MouseMove(sender As Object, e As MouseEventArgs) Handles XmlHtmDisplay1.MouseMove

        If XmlHtmDisplay1.Cursor.Handle <> GetCursor() Then

            Dim indx As Integer = XmlHtmDisplay1.GetCharIndexFromPosition(e.Location)
            Dim si As Integer = indx
            Dim ei As Integer = indx

            For ss As Integer = indx To 0 Step -1
                If XmlHtmDisplay1.Text(ss) = " " Or XmlHtmDisplay1.Text(ss) = vbLf Then
                    si = ss + 1
                    Exit For
                End If
            Next

            For ee As Integer = indx To XmlHtmDisplay1.TextLength - 1
                If XmlHtmDisplay1.Text(ee) = " " Or XmlHtmDisplay1.Text(ee) = vbLf Then
                    ei = ee
                    Exit For
                ElseIf ee = XmlHtmDisplay1.TextLength - 1 Then
                    ei = ee + 1
                    Exit For
                End If
            Next

            If si < ei Then
                Dim str As String = XmlHtmDisplay1.Text.Substring(si, ei - si)
                txtLink.Text = str
            End If


        Else
            txtLink.Text = ""
        End If

    End Sub

    Private Sub XmlHtmDisplay1_TextChanged(sender As Object, e As EventArgs) Handles XmlHtmDisplay1.TextChanged
        DocTextChanged = True
    End Sub

    Private Sub btnZoomOut_Click(sender As Object, e As EventArgs) Handles btnZoomOut.Click
        'Zoom Out
        ZoomFactor = ZoomFactor / 1.1
    End Sub

    Private Sub btnZoomIn_Click(sender As Object, e As EventArgs) Handles btnZoomIn.Click
        'Zoom In
        ZoomFactor = ZoomFactor * 1.1
    End Sub

    Private Sub txtZoomPercent_TextChanged(sender As Object, e As EventArgs) Handles txtZoomPercent.TextChanged
        Dim ZoomPercent As Integer
        ZoomPercent = Int(Val(txtZoomPercent.Text))
        If ZoomPercent < 10 Then ZoomPercent = 10
        If ZoomPercent > 400 Then ZoomPercent = 400
        ZoomFactor = ZoomPercent / 100
    End Sub

    Private Sub txtZoomPercent_LostFocus(sender As Object, e As EventArgs) Handles txtZoomPercent.LostFocus

    End Sub

    Private Sub txtZoomPercent_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtZoomPercent.MouseDoubleClick
        'Use double-click to reset the Zoom Percent to 100.
        txtZoomPercent.Text = 100

    End Sub

    Private Sub chkWordWrap_CheckedChanged(sender As Object, e As EventArgs) Handles chkWordWrap.CheckedChanged
        If chkWordWrap.Checked Then
            XmlHtmDisplay1.WordWrap = True
        Else
            XmlHtmDisplay1.WordWrap = False
        End If
    End Sub

    Private Sub btnPasteImage_Click(sender As Object, e As EventArgs) Handles btnPasteImage.Click
        'Paste the image in the clipboard into the rich text box:

        Dim ClipImage As System.Drawing.Image = Nothing

        If My.Computer.Clipboard.ContainsImage() Then
            ClipImage = My.Computer.Clipboard.GetImage
            'ClipImage = System.Windows.Forms.Clipboard.GetImage
        End If

        If ClipImage IsNot Nothing Then
            'XmlHtmDisplay1.Paste()

            Dim myBitMap As Bitmap = ClipImage
            My.Computer.Clipboard.SetImage(myBitMap)
            XmlHtmDisplay1.Paste()
        End If

    End Sub


#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class