Public Class frmTextDisplay
    'Form used to display a text (TXT) document.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    'Declare Forms used by the application:

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
            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        'Dim MinWidthVisible As Integer = 48 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        'Dim MinHeightVisible As Integer = 48 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

        ''Check if the top of the form is less than zero:
        'If Me.Top < 0 Then Me.Top = 0

        'Check if the top of the form is above the top of the Working Area:
        If Me.Top < WARect.Top Then
            Me.Top = WARect.Top
        End If

        'Check if the top of the form is too close to the bottom of the Working Area:
        If (Me.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
            Me.Top = WARect.Top + WARect.Height - MinHeightVisible
        End If

        'Check if the left edge of the form is too close to the right edge of the Working Area:
        If (Me.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
            Me.Left = WARect.Left + WARect.Width - MinWidthVisible
        End If

        'Check if the right edge of the form is too close to the left edge of the Working Area:
        If (Me.Left + Me.Width - MinWidthVisible) < WARect.Left Then
            Me.Left = WARect.Left - Me.Width + MinWidthVisible
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

        XmlHtmDisplay1.WordWrap = False
        XmlHtmDisplay1.Settings = Main.XmlHtmDisplay1.Settings

        Dim FontName As String = XmlHtmDisplay1.Settings.PlainText.FontName
        Dim FontSize As Single = XmlHtmDisplay1.Settings.PlainText.PointSize
        Dim FontStyle As System.Drawing.FontStyle = Drawing.FontStyle.Regular
        If XmlHtmDisplay1.Settings.PlainText.Bold = True Then
            FontStyle = FontStyle Or Drawing.FontStyle.Bold
        End If
        If XmlHtmDisplay1.Settings.PlainText.Italic = True Then
            FontStyle = FontStyle Or Drawing.FontStyle.Italic
        End If

        XmlHtmDisplay1.Font = New System.Drawing.Font(FontName, FontSize, FontStyle)
        XmlHtmDisplay1.ForeColor = XmlHtmDisplay1.Settings.PlainText.Color
        Main.Message.Add("XmlHtmDisplay1.Settings.PlainText.Color" & XmlHtmDisplay1.Settings.PlainText.Color.Name & vbCrLf)

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form

        'If DocTextChanged = True Then
        '    Dim result As Integer = MessageBox.Show("Save changes to the current document?", "Notice", MessageBoxButtons.YesNoCancel)
        '    If result = DialogResult.Cancel Then
        '        Exit Sub
        '    ElseIf result = DialogResult.Yes Then
        '        SaveDocument()
        '    ElseIf result = DialogResult.No Then
        '        'Do not save the changes!
        '    End If
        'End If

        'Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the RtfDisplayFormClosed method to select the correct form to set to nothing.
        'If FileName <> "" Then
        '    LastFileName = FileName
        'End If
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If

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
    End Sub

    Private Sub frmTxtDisplay_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Main.TextDisplayFormClosed()
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

            OpenFileDialog1.Filter = "Text File (*.txt)| *.txt"
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

            If Main.ItemInfo.ContainsKey(FileName) Then
                Main.ItemInfo(FileName).LastEditDate = Now
            Else
                Main.Message.AddWarning("Document not found in the list: " & FileName & vbCrLf)
            End If

        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        'If DocumentTextChanged = True Then
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
            XmlHtmDisplay1.Clear()
            rtbData.Position = 0
            XmlHtmDisplay1.LoadFile(rtbData, RichTextBoxStreamType.PlainText)
            DocTextChanged = False
            LastFileName = FileName
            LastFileLocationType = LocationTypes.Project
            LastFileDirectory = ""
        Else
            XmlHtmDisplay1.LoadFile(FileDirectory & "\" & FileName)
            DocTextChanged = False
            LastFileName = FileName
            LastFileLocationType = LocationTypes.FileSystem
            LastFileDirectory = FileDirectory
        End If

    End Sub


#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class