Public Class frmHtmlDisplay
    'Form used to display a formatted HTML document.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

    'Declare Forms used by the application:
    Public WithEvents WebView As frmWebView

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _formNo As Integer = 0 'Multiple instances of this form can be displayed. FormNo is the index number of the form in XmlDisplayFormList.
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

    Private _plainTextDisplay As Boolean = False 'True if the file is displayed in plain text.
    Property PlainTextDisplay As Boolean
        Get
            Return _plainTextDisplay
        End Get
        Set(value As Boolean)
            _plainTextDisplay = value
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

        Debug.Print("Loading frmHtmlDisplay")

        RestoreFormSettings()   'Restore the form settings

        Timer1.Interval = 2500 '2.5 seconds interval. Wait time between typing in code and updating the formatting (if no more keys pressed).
        Timer1.Enabled = True

        XmlHtmDisplay1.WordWrap = False
        XmlHtmDisplay1.Settings = Main.XmlHtmDisplay1.Settings

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Form
        Main.ClosedFormNo = FormNo 'The Main form property ClosedFormNo is set to this form number. This is used in the RtfDisplayFormClosed method to select the correct form to set to nothing.
        Me.Close() 'Close the form
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        Else
            'Dont save settings if the form is minimised.
        End If
    End Sub

    Private Sub frmXmlDisplay_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Main.HtmlDisplayFormClosed()
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

    Private Sub btnWebView_Click(sender As Object, e As EventArgs) Handles btnWebView.Click
        'Open the Web View form:

        Dim FormNo As Integer = Main.NewWebView()
        Main.WebViewFormList(FormNo).FileName = FileName
        Main.WebViewFormList(FormNo).Description = Description
        Main.WebViewFormList(FormNo).FileDirectory = FileDirectory
        Main.WebViewFormList(FormNo).FileLocationType = FileLocationType
        Main.WebViewFormList(FormNo).OpenDocument

    End Sub

    Private Sub WebView_FormClosed(sender As Object, e As FormClosedEventArgs) Handles WebView.FormClosed
        WebView = Nothing
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
                If XmlHtmDisplay1.SaveXmlFile(FileDirectory & "\" & FileName) = True Then
                    'File was saved OK.
                    LastFileName = FileName 'Update the LastFilePath.
                    LastFileLocationType = LocationTypes.FileSystem
                    LastFileDirectory = FileDirectory
                    DocTextChanged = False
                End If
            Else 'Save the document in the current project.

                Dim htmData As New IO.MemoryStream
                XmlHtmDisplay1.SaveFile(htmData, RichTextBoxStreamType.PlainText)
                htmData.Position = 0
                Main.Project.SaveData(FileName, htmData)

                LastFileName = FileName 'Update the LastFileName.
                LastFileLocationType = LocationTypes.Project
                LastFileDirectory = ""
                DocTextChanged = False
            End If
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        'Close the current XML document.

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
        XmlHtmDisplay1.Clear() 'Test XmlHtmDisplay
        DocTextChanged = False

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

        If Main.rbFileInProject.Checked = True Then 'Open an HTML file in the current Project.
            Dim SelectedFile As String = Main.Project.SelectDataFile("HyperTtext markup language", "hml")
            If SelectedFile = "" Then
                'No file selected!
            Else
                FileName = SelectedFile
                FileLocationType = LocationTypes.Project
                FileDirectory = ""
                OpenDocument()
            End If

        Else 'Open an HTML file in the File System.
            If LastFileName = "" Then 'There is no last HTML file saved.

            Else 'The last HTML file path was saved.
                OpenFileDialog1.InitialDirectory = LastFileDirectory
                OpenFileDialog1.FileName = LastFileName
            End If

            OpenFileDialog1.Filter = "All Files (*.*)| *.*"
            SendKeys.Send("{HOME}") 'To move the cursor to the start of the FileName
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                FileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                FileLocationType = LocationTypes.FileSystem
                FileDirectory = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
                OpenDocument()
            End If
        End If

    End Sub

    Public Sub OpenDocument()
        'Open the document specified by FileName, FileLocationType and FileDirectory.

        If FileLocationType = LocationTypes.Project Then

            Dim rtbData As New IO.MemoryStream
            Main.Project.ReadData(FileName, rtbData)
            XmlHtmDisplay1.Clear()
            rtbData.Position = 0
            XmlHtmDisplay1.LoadFile(rtbData, RichTextBoxStreamType.PlainText)

            Dim htmText As String = XmlHtmDisplay1.Text

            XmlHtmDisplay1.Rtf = XmlHtmDisplay1.HmlToRtf(htmText)
            PlainTextDisplay = False

            DocTextChanged = False
            LastFileName = FileName
            LastFileLocationType = LocationTypes.Project
            LastFileDirectory = ""
        Else

            XmlHtmDisplay1.LoadFile(FileDirectory & "\" & FileName, RichTextBoxStreamType.PlainText)
            DocTextChanged = False
            LastFileName = FileName
            LastFileLocationType = LocationTypes.FileSystem
            LastFileDirectory = FileDirectory
        End If

    End Sub

    Public Sub OpenDocumentInPlainText()
        'Open the document specified by FileName, FileLocationType and FileDirectory.
        'The document will be displayed in Plain Text.

        If FileLocationType = LocationTypes.Project Then
            Dim htmData As New IO.MemoryStream
            Main.Project.ReadData(FileName, htmData)
            XmlHtmDisplay1.Clear()
            htmData.Position = 0
            XmlHtmDisplay1.LoadFile(htmData, RichTextBoxStreamType.PlainText)
            PlainTextDisplay = True

            DocTextChanged = False
            LastFileName = FileName
            LastFileLocationType = LocationTypes.Project
            LastFileDirectory = ""
        Else
            XmlHtmDisplay1.LoadFile(FileDirectory & "\" & FileName, RichTextBoxStreamType.PlainText)
            PlainTextDisplay = True
            DocTextChanged = False
            LastFileName = FileName
            LastFileLocationType = LocationTypes.FileSystem
            LastFileDirectory = FileDirectory
        End If
    End Sub

    Private Sub XmlDisplay1_Click(sender As Object, e As EventArgs)

        Dim Position As Integer
            Dim Line As Integer
            Dim PositionInLine As Integer

        Position = XmlHtmDisplay1.SelectionStart
        Line = XmlHtmDisplay1.GetLineFromCharIndex(Position)
        PositionInLine = Position - XmlHtmDisplay1.GetFirstCharIndexOfCurrentLine

    End Sub

    Private Sub XmlHtmDisplay1_Message(Msg As String) Handles XmlHtmDisplay1.Message
        Main.Message.Add(Msg)
    End Sub

    Private Sub btnShowRtf_Click(sender As Object, e As EventArgs) Handles btnShowRtf.Click
        'Show the RTF code in the Message window:
        Main.Message.Add(vbCrLf & "RTF Code --------------------------------" & vbCrLf & XmlHtmDisplay1.Rtf & vbCrLf & vbCrLf)
    End Sub

    Private Sub btnShowHtmlCols_Click(sender As Object, e As EventArgs) Handles btnShowHtmlCols.Click
        'SHow the HTML colors in the Message window:
        Main.Message.Add(vbCrLf & "HTML Colors ---------------------------------" & vbCrLf)
        Main.Message.Add("HTML Text      <Text> Index = " & XmlHtmDisplay1.Settings.HText.ColorIndex & " Name = " & XmlHtmDisplay1.Settings.HText.Color.Name & vbCrLf)
        Main.Message.Add("HTML Element   <Elem> Index = " & XmlHtmDisplay1.Settings.HElement.ColorIndex & " Name = " & XmlHtmDisplay1.Settings.HElement.Color.Name & vbCrLf)
        Main.Message.Add("HTML Attribute <Attr> Index = " & XmlHtmDisplay1.Settings.HAttribute.ColorIndex & " Name = " & XmlHtmDisplay1.Settings.HAttribute.Color.Name & vbCrLf)
        Main.Message.Add("HTML Comment   <Comm> Index = " & XmlHtmDisplay1.Settings.HComment.ColorIndex & " Name = " & XmlHtmDisplay1.Settings.HComment.Color.Name & vbCrLf)
        Main.Message.Add("HTML Spec Char <Char> Index = " & XmlHtmDisplay1.Settings.HChar.ColorIndex & " Name = " & XmlHtmDisplay1.Settings.HChar.Color.Name & vbCrLf)
        Main.Message.Add("HTML Value     <Valu> Index = " & XmlHtmDisplay1.Settings.HValue.ColorIndex & " Name = " & XmlHtmDisplay1.Settings.HValue.Color.Name & vbCrLf)
        Main.Message.Add("HTML Style     <Styl> Index = " & XmlHtmDisplay1.Settings.HStyle.ColorIndex & " Name = " & XmlHtmDisplay1.Settings.HStyle.Color.Name & vbCrLf & vbCrLf)

    End Sub

    Private Sub btnAddToLibrary_Click(sender As Object, e As EventArgs) Handles btnAddToLibrary.Click

    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        'Update the display.

        If PlainTextDisplay = True Then
            Main.Message.AddWarning("The HTML display form is set to display Plain Text only." & vbCrLf)
        Else
            Dim HtmText As String = XmlHtmDisplay1.Text
            XmlHtmDisplay1.Rtf = XmlHtmDisplay1.HmlToRtf(HtmText)
        End If
    End Sub



    Private Sub XmlHtmDisplay1_Click(sender As Object, e As EventArgs) Handles XmlHtmDisplay1.Click
        'Get the line number and column number of the cursor position:

        Dim LineNo As Integer = XmlHtmDisplay1.GetLineFromCharIndex(XmlHtmDisplay1.SelectionStart) + 1
        Dim CharIndex As Integer = XmlHtmDisplay1.SelectionStart
        Dim LineStartIndex As Integer = XmlHtmDisplay1.GetFirstCharIndexOfCurrentLine
        Dim ColNo As Integer = CharIndex - LineStartIndex

        txtInfo.Text = "L" & LineNo & " C" & ColNo

    End Sub

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class