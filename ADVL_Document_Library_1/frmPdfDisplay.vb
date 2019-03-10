Public Class frmPdfDisplay
    'Form used to display a portable document format (PDF) document.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================


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

        Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the RtfDisplayFormClosed method to select the correct form to set to nothing.
        'If FilePath <> "" Then
        If FileName <> "" Then
            LastFileName = FileName
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


        If Main.rbFileInProject.Checked = True Then 'Open an RTF file in the current Project.
            'Dim SelectedFile As String = Main.Project.SelectDataFile("Rich text format", "rtf")
            Dim SelectedFile As String = Main.Project.SelectDataFile("Portable Document format", "pdf")
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

            OpenFileDialog1.Filter = "Portable Document Format (*.pdf)| *.pdf"
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


#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Public Sub OpenDocument()
        'Open the document specified by FileName, FileLocationType and FileDirectory.

        If FileLocationType = LocationTypes.Project Then
            If Main.Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then

            Else
                Dim DataLocnPath As String = Main.Project.DataLocn.Path
                AxAcroPDF1.LoadFile(DataLocnPath & "\" & FileName)
                AxAcroPDF1.Focus()
            End If
        Else
            AxAcroPDF1.LoadFile(FileDirectory & "\" & FileName)
        End If
    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------




End Class