Public Class frmNewRtf
    'Form used to create a new Rich Text Format (RTF) document.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _fileName As String = "" 'The name of the file to create.
    Property FileName As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
            If _fileName.EndsWith(".rtf") Then

            Else
                _fileName = System.IO.Path.GetFileNameWithoutExtension(_fileName) & ".rtf"
            End If
            txtFileName.Text = _fileName
        End Set
    End Property

    Enum LocationTypes
        Project
        FileSystem
    End Enum

    Private _fileLocation As LocationTypes = LocationTypes.Project 'The location to save the Document File. (Either the current project or the file system.)
    Property FileLocation As LocationTypes
        Get
            Return _fileLocation
        End Get
        Set(value As LocationTypes)
            _fileLocation = value
            If _fileLocation = LocationTypes.Project Then
                rbFileInProject.Checked = True
                txtDirectory.Enabled = False
                Label2.Enabled = False
                btnFind.Enabled = False
            Else
                rbFileInFileSystem.Checked = True
                txtDirectory.Enabled = True
                Label2.Enabled = True
                btnFind.Enabled = True
            End If
        End Set
    End Property

    Private _directory As String = "" 'The directory to save the new file.
    Property Directory As String
        Get
            Return _directory
        End Get
        Set(value As String)
            _directory = value
            txtDirectory.Text = _directory
        End Set
    End Property

    Private _description = "" 'A description of the rtf document.
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
            txtDescription = _description
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
                               <FileName><%= FileName %></FileName>
                               <FileLocation><%= FileLocation %></FileLocation>
                               <Directory><%= Directory %></Directory>
                               <OpenInNewWindow><%= chkOpenInNewWindow.Checked %></OpenInNewWindow>
                               <AddToLibrary><%= chkAddToLibrary.Checked %></AddToLibrary>
                               <Description><%= txtDescription.Text %></Description>
                               <!---->
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
            If Settings.<FormSettings>.<FileName>.Value <> Nothing Then FileName = Settings.<FormSettings>.<FileName>.Value
            If Settings.<FormSettings>.<FileLocation>.Value <> Nothing Then
                'FileLocation = Settings.<FormSettings>.<FileLocation>.Value
                Select Case Settings.<FormSettings>.<FileLocation>.Value
                    Case "Project"
                        FileLocation = LocationTypes.Project
                    Case "FileSystem"
                        FileLocation = LocationTypes.FileSystem
                End Select
            End If
            If Settings.<FormSettings>.<Directory>.Value <> Nothing Then Directory = Settings.<FormSettings>.<Directory>.Value
            If Settings.<FormSettings>.<OpenInNewWindow>.Value <> Nothing Then chkOpenInNewWindow.Checked = Settings.<FormSettings>.<OpenInNewWindow>.Value
            If Settings.<FormSettings>.<AddToLibrary>.Value <> Nothing Then
                chkAddToLibrary.Checked = Settings.<FormSettings>.<AddToLibrary>.Value
                If chkAddToLibrary.Checked = True Then
                    Label3.Enabled = True
                    txtDescription.Enabled = True
                Else
                    Label3.Enabled = False
                    txtDescription.Enabled = False
                End If
            End If
            If Settings.<FormSettings>.<Description>.Value <> Nothing Then txtDescription.Text = Settings.<FormSettings>.<Description>.Value
            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

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
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Me.Close() 'Close the form
    End Sub

    'Private Sub frmTemplate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click

        If Directory <> "" Then
            FolderBrowserDialog1.SelectedPath = Directory
        End If

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Directory = FolderBrowserDialog1.SelectedPath
        End If

    End Sub


    Private Sub txtFileName_LostFocus(sender As Object, e As EventArgs) Handles txtFileName.LostFocus
        FileName = txtFileName.Text
    End Sub

    Private Sub txtDirectory_LostFocus(sender As Object, e As EventArgs) Handles txtDirectory.LostFocus
        _directory = txtDirectory.Text
    End Sub

    Private Sub txtDescription_LostFocus(sender As Object, e As EventArgs) Handles txtDescription.LostFocus
        _description = txtDescription.Text
    End Sub

    Private Sub chkAddToLibrary_CheckedChanged(sender As Object, e As EventArgs) Handles chkAddToLibrary.CheckedChanged
        If chkAddToLibrary.Checked = True Then
            Label3.Enabled = True
            txtDescription.Enabled = True
        Else
            Label3.Enabled = False
            txtDescription.Enabled = False
        End If
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        'Create a new RTF file.

        'Check if file name is already in use: ===================================================================
        If FileLocation = LocationTypes.Project Then 'Document file will be saved in the current Project.
            If Main.Project.DataFileExists(FileName) Then
                Dim Result As Integer = MessageBox.Show("Overwrite existing file with the same name?", "Notice", MessageBoxButtons.YesNoCancel)
                If Result = DialogResult.Cancel Then
                    Exit Sub
                ElseIf Result = DialogResult.Yes Then
                    'Continue. New document will overwrite the old one.
                ElseIf Result = DialogResult.No Then
                    Exit Sub
                End If
            End If
        Else 'Document file will be saved in the File System.
            If System.IO.File.Exists(Directory & "\" & FileName) Then
                Dim Result As Integer = MessageBox.Show("Overwrite existing file with the same name?", "Notice", MessageBoxButtons.YesNoCancel)
                If Result = DialogResult.Cancel Then
                    Exit Sub
                ElseIf Result = DialogResult.Yes Then
                    'Continue. New document will overwrite the old one.
                ElseIf Result = DialogResult.No Then
                    Exit Sub
                End If
            End If
        End If
        '----------------------------------------------------------------------------------------------------------

        If chkOpenInNewWindow.Checked = True Then 'Open a new RTF window to display the new RTF file.
            Dim FormNo As Integer = Main.NewRtfDisplay
            If FileLocation = LocationTypes.Project Then 'Document file will be saved in the current Project.
                Main.RtfDisplayFormList(FormNo).FileName = FileName
                Main.RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                Main.RtfDisplayFormList(FormNo).FileDirectory = ""
                Main.RtfDisplayFormList(FormNo).SaveDocument
            Else 'Document file will be saved in the File System.
                Main.RtfDisplayFormList(FormNo).FileName = FileName
                Main.RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                Main.RtfDisplayFormList(FormNo).FileDirectory = Directory
                Main.RtfDisplayFormList(FormNo).SaveDocument
            End If

        Else 'Display the new RTF file in the main window.
            If Main.DocumentTextChanged = True Then
                Dim Result As Integer = MessageBox.Show("Save changes to the current document?", "Notice", MessageBoxButtons.YesNoCancel)
                If Result = DialogResult.Cancel Then
                    Exit Sub
                ElseIf Result = DialogResult.Yes Then
                    Main.SaveDocument()
                ElseIf Result = DialogResult.No Then
                    'Do not save the changes.
                End If
            End If
            If FileLocation = LocationTypes.Project Then 'Document file will be saved in the current Project.
                Main.XmlHtmDisplay1.Clear()
                Main.FileName = FileName
                Main.RtfFileName = FileName
                Main.FileLocationType = Main.LocationTypes.Project
                Main.FileDirectory = Directory
                Main.SaveDocument()
            Else 'Document file will be saved in the File System.
                Main.XmlHtmDisplay1.Clear()
                Main.FileName = FileName
                Main.RtfFileName = FileName
                Main.FileLocationType = Main.LocationTypes.FileSystem
                Main.FileDirectory = ""
                Main.SaveDocument()
            End If

        End If

    End Sub

    Private Sub rbFileInProject_CheckedChanged(sender As Object, e As EventArgs) Handles rbFileInProject.CheckedChanged
        If rbFileInProject.Checked = True Then
            _fileLocation = LocationTypes.Project
            txtDirectory.Enabled = False
            Label2.Enabled = False
            btnFind.Enabled = False
        Else
            _fileLocation = LocationTypes.FileSystem
            txtDirectory.Enabled = True
            Label2.Enabled = True
            btnFind.Enabled = True
        End If
    End Sub

    Private Sub txtFileName_TextChanged(sender As Object, e As EventArgs) Handles txtFileName.TextChanged

    End Sub

    Private Sub txtDirectory_TextChanged(sender As Object, e As EventArgs) Handles txtDirectory.TextChanged

    End Sub

    Private Sub txtDescription_TextChanged(sender As Object, e As EventArgs) Handles txtDescription.TextChanged

    End Sub






#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Events - Events that can be triggered by this form." '==========================================================================================================================

#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------



End Class