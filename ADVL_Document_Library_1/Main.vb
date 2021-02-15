'==============================================================================================================================================================================================
'
'Copyright 2018 Signalworks Pty Ltd, ABN 26 066 681 598

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
''WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Imports System.ComponentModel
Imports System.Security.Permissions
<PermissionSet(SecurityAction.Demand, Name:="FullTrust")>
<System.Runtime.InteropServices.ComVisibleAttribute(True)>
Public Class Main
    'The ADVL_XML_Display_1 is used to display and edit XML documents.

#Region " Coding Notes - Notes on the code used in this class." '==============================================================================================================================

    'ADD THE SYSTEM UTILITIES REFERENCE: ==========================================================================================
    'The following references are required by this software: 
    'Project \ Add Reference... \ ADVL_Utilities_Library_1.dll
    'The Utilities Library is used for Project Management, Archive file management, running XSequence files and running XMessage files.
    'If there are problems with a reference, try deleting it from the references list and adding it again.

    'ADD THE SERVICE REFERENCE: ===================================================================================================
    'A service reference to the Message Service must be added to the source code before this service can be used.
    'This is used to connect to the Application Network.

    'Adding the service reference to a project that includes the WcfMsgServiceLib project: -----------------------------------------
    'Project \ Add Service Reference
    'Press the Discover button.
    'Expand the items in the Services window and select IMsgService.
    'Press OK.
    '------------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------------
    'Adding the service reference to other projects that dont include the WcfMsgServiceLib project: -------------------------------
    'Run the ADVL_Application_Network_1 application to start the Application Network message service.
    'In Microsoft Visual Studio select: Project \ Add Service Reference
    'Enter the address: http://localhost:8733/ADVLService
    'UPDATE: Enter the address: http://localhost:8734/ADVLService
    'Press the Go button.
    'MsgService is found.
    'Press OK to add ServiceReference1 to the project.
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'ADD THE MsgServiceCallback CODE: =============================================================================================
    'This is used to connect to the Application Network.
    'In Microsoft Visual Studio select: Project \ Add Class
    'MsgServiceCallback.vb
    'Add the following code to the class:
    'Imports System.ServiceModel
    'Public Class MsgServiceCallback
    '    Implements ServiceReference1.IMsgServiceCallback
    '    Public Sub OnSendMessage(message As String) Implements ServiceReference1.IMsgServiceCallback.OnSendMessage
    '        'A message has been received.
    '        'Set the InstrReceived property value to the message (usually in XMessage format). This will also apply the instructions in the XMessage.
    '        Main.InstrReceived = message
    '    End Sub
    'End Class
    '------------------------------------------------------------------------------------------------------------------------------
    '
    'Display a PDF with VB.NET: ===================================================================================================
    'Reference: https://www.thoughtco.com/display-a-pdf-with-vbnet-3424227
    'If not already installed, download and install the free Acrobat Reader.
    'Add the Adobe ActiveX COM control to the VB.NET Toolbox:
    'NOTE: This can ony be used in a Windows Forms application (WPF not yet supported).
    'Righ-click on any Toolbox tab (such as Common Controls) and select "Choose Items..."
    'Select the "COM Components" tab and click the checkbox beside "Adobe PDF Reader" and click OK.
    '------------------------------------------------------------------------------------------------------------------------------



#End Region 'Coding Notes ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Variable Declarations - All the variables and class objects used in this form and this application." '===============================================================================

    Public WithEvents ApplicationInfo As New ADVL_Utilities_Library_1.ApplicationInfo 'This object is used to store application information.
    Public WithEvents Project As New ADVL_Utilities_Library_1.Project 'This object is used to store Project information.
    Public WithEvents Message As New ADVL_Utilities_Library_1.Message 'This object is used to display messages in the Messages window.
    Public WithEvents ApplicationUsage As New ADVL_Utilities_Library_1.Usage 'This object stores application usage information.


    'Declare Forms used by the application:
    Public WithEvents WebPageList As frmWebPageList

    Public WithEvents NewWFHtmlDisplay As frmWFHtmlDisplay
    Public WFHtmlDisplayFormList As New ArrayList 'Used for displaying multiple HtmlDisplay forms.

    Public WithEvents NewWebPage As frmWebPage
    Public WebPageFormList As New ArrayList 'Used for displaying multiple WebView forms.

    Public WithEvents EditXml As frmEditXml
    Public WithEvents EditRtf As frmEditRtf
    Public WithEvents NewXml As frmNewXml
    Public WithEvents NewRtf As frmNewRtf
    Public WithEvents NewCollection As frmNewCollection_Old
    Public WithEvents NewLibrary As frmNewLibrary

    Public WithEvents XmlDisplay As frmXmlDisplay
    Public XmlDisplayFormList As New ArrayList 'Used for displaying multiple XmlDisplay forms.

    Public WithEvents RtfDisplay As frmRtfDisplay
    Public RtfDisplayFormList As New ArrayList 'Used for displaying multiple RtfDisplay forms.

    Public WithEvents HtmlDisplay As frmHtmlDisplay
    Public HtmlDisplayFormList As New ArrayList 'Used for displaying multiple HtmlDisplay forms.

    Public WithEvents WebView As frmWebView
    Public WebViewFormList As New ArrayList 'Used for displaying multiple WebView forms.

    Public WithEvents TextDisplay As frmTextDisplay
    Public TextDisplayFormList As New ArrayList 'Used for displaying multiple TextDisplay forms.

    Public WithEvents PdfDisplay As frmPdfDisplay
    Public PdfDisplayFormList As New ArrayList 'Used for displaying multiple PdfDisplay forms.


    'Declare objects used to connect to the Message Service:
    Public client As ServiceReference1.MsgServiceClient
    Public WithEvents XMsg As New ADVL_Utilities_Library_1.XMessage
    Dim XDoc As New System.Xml.XmlDocument
    Public Status As New System.Collections.Specialized.StringCollection
    Dim ClientProNetName As String = "" 'The name of the client Project Network requesting service. 
    Dim ClientAppName As String = "" 'The name of the client requesting service
    Dim ClientConnName As String = "" 'The name of the client connection requesting service
    Dim MessageXDoc As System.Xml.Linq.XDocument
    Dim xmessage As XElement 'This will contain the message. It will be added to MessageXDoc.
    Dim xlocns As New List(Of XElement) 'A list of locations. Each location forms part of the reply message. The information in the reply message will be sent to the specified location in the client application.
    Dim MessageText As String = "" 'The text of a message sent through the Application Network.

    'Dim CompletionInstruction As String = "Stop" 'The last instruction returned on completion of the processing of an XMessage.
    Public OnCompletionInstruction As String = "Stop" 'The last instruction returned in <EndInstruction> on completion of the processing of an XMessage.
    Public EndInstruction As String = "Stop" 'Another method of specifying the last instruction. This is processed in the EndOfSequence section of XMsg.Instructions.

    Public ConnectionName As String = "" 'The name of the connection used to connect this application to the ComNet (Message Service).

    Public ProNetName As String = "" 'The name of the Project Network
    Public ProNetPath As String = "" 'The path of the Project Network

    Public AdvlNetworkAppPath As String = "" 'The application path of the ADVL Network application (ComNet). This is where the "Application.Lock" file will be while ComNet is running
    Public AdvlNetworkExePath As String = "" 'The executable path of the ADVL Network.

    'Variable for local processing of an XMessage:
    Public WithEvents XMsgLocal As New ADVL_Utilities_Library_1.XMessage
    Dim XDocLocal As New System.Xml.XmlDocument
    Public StatusLocal As New System.Collections.Specialized.StringCollection

    Dim WithEvents Zip As ADVL_Utilities_Library_1.ZipComp

    'Dim ItemInfo As New Dictionary(Of String, clsItemInfo) 'Dictionary of information about each item (document or collection) in the library.
    Public ItemInfo As New Dictionary(Of String, clsItemInfo) 'Dictionary of information about each item (document or collection) in the library.


    'The following variables are used to run JavaScript in Web Pages loaded into the Document View: -------------------
    Public WithEvents XSeq As New ADVL_Utilities_Library_1.XSequence
    'To run an XSequence:
    '  XSeq.RunXSequence(xDoc, Status) 'ImportStatus in Import
    '    Handle events:
    '      XSeq.ErrorMsg
    '      XSeq.Instruction(Info, Locn)

    Private XStatus As New System.Collections.Specialized.StringCollection

    'Variables used to restore Item values on a web page.
    Private FormName As String
    Private ItemName As String
    Private SelectId As String
    '----------------------------------------------------------------------------------------------------------------------

    'Main.Load variables:
    Dim ProjectSelected As Boolean = False 'If True, a project has been selected using Command Arguments. Used in Main.Load.
    Dim StartupConnectionName As String = "" 'If not "" the application will be connected to the AppNet using this connection name in  Main.Load.

    'Search variables:
    Dim SearchFileName As String = "" 'File Name of the document to be searched.
    Dim SearchText As String = "" 'The string to tbe searched form
    Dim SearchHighlight As Boolean = True 'If True, the found string will be highlighted in the document view.
    Dim SearchFindFirst As Boolean = True 'If True, the first instance of the string will be found.

    Private WithEvents bgwComCheck As New System.ComponentModel.BackgroundWorker 'Used to perform communication checks on a separate thread.

    Public WithEvents bgwSendMessage As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service.
    Dim SendMessageParams As New clsSendMessageParams 'This hold the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message

    'Alternative SendMessage background worker - needed to send a message while instructions are being processed.
    Public WithEvents bgwSendMessageAlt As New System.ComponentModel.BackgroundWorker 'Used to send a message through the Message Service - alternative backgound worker.
    Dim SendMessageParamsAlt As New clsSendMessageParams 'This hold the Send Message parameters: .ProjectNetworkName, .ConnectionName & .Message - for the alternative background worker.

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _connectionHashcode As Integer 'The Application Network connection hashcode. This is used to identify a connection in the Application Netowrk when reconnecting.
    Property ConnectionHashcode As Integer
        Get
            Return _connectionHashcode
        End Get
        Set(value As Integer)
            _connectionHashcode = value
        End Set
    End Property

    Private _connectedToComNet As Boolean = False  'True if the application is connected to the Communication Network (Message Service).
    Property ConnectedToComNet As Boolean
        Get
            Return _connectedToComNet
        End Get
        Set(value As Boolean)
            _connectedToComNet = value
        End Set
    End Property


    Private _instrReceived As String = "" 'Contains Instructions received from the Application Network message service.
    Property InstrReceived As String
        Get
            Return _instrReceived
        End Get
        Set(value As String)
            If value = Nothing Then
                Message.Add("Empty message received!")
            Else
                _instrReceived = value
                ProcessInstructions(_instrReceived)
            End If
        End Set
    End Property

    Private Sub ProcessInstructions(ByVal Instructions As String)
        'Process the XMessage instructions.

        Dim MsgType As String
        If Instructions.StartsWith("<XMsg>") Then
            MsgType = "XMsg"
            If ShowXMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        ElseIf Instructions.StartsWith("<XSys>") Then
            MsgType = "XSys"
            If ShowSysMessages Then
                'Add the message header to the XMessages window:
                Message.XAddText("System Message received: " & vbCrLf, "XmlReceivedNotice")
            End If
        Else
            MsgType = "Unknown"
        End If

        'If ShowXMessages Then
        '    'Add the message header to the XMessages window:
        '    Message.XAddText("Message received: " & vbCrLf, "XmlReceivedNotice")
        'End If

        'If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
        If MsgType = "XMsg" Or MsgType = "XSys" Then 'This is an XMessage or XSystem set of instructions.
                Try
                    'Inititalise the reply message:
                    ClientProNetName = ""
                    ClientConnName = ""
                    ClientAppName = ""
                xlocns.Clear() 'Clear the list of locations in the reply message.

                Dim Decl As New XDeclaration("1.0", "utf-8", "yes")
                    MessageXDoc = New XDocument(Decl, Nothing) 'Reply message - this will be sent to the Client App.
                'xmessage = New XElement("XMsg")
                xmessage = New XElement(MsgType)
                xlocns.Add(New XElement("Main")) 'Initially set the location in the Client App to Main.

                    'Run the received message:
                    Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"

                    XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                'If ShowXMessages Then
                '    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddXml(XDoc)  'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                XMsg.Run(XDoc, Status)
                Catch ex As Exception
                    Message.Add("Error running XMsg: " & ex.Message & vbCrLf)
                End Try

                'XMessage has been run.
                'Reply to this message:
                'Add the message reply to the XMessages window:
                'Complete the MessageXDoc:
                xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the last location reply instructions to the message.
                MessageXDoc.Add(xmessage)
                MessageText = MessageXDoc.ToString

                If ClientConnName = "" Then
                'No client to send a message to - process the message locally.
                'If ShowXMessages Then
                '    Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                '    Message.XAddXml(MessageText)
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message processed locally:" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                ProcessLocalInstructions(MessageText)
                Else
                'If ShowXMessages Then
                '    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")
                '    Message.XAddXml(MessageText)
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If
                If (MsgType = "XMsg") And ShowXMessages Then
                    Message.XAddText("Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                ElseIf (MsgType = "XSys") And ShowSysMessages Then
                    Message.XAddText("System Message sent to [" & ClientProNetName & "]." & ClientConnName & ":" & vbCrLf, "XmlSentNotice")   'NOTE: There is no SendMessage code in the Message Service application!
                    Message.XAddXml(MessageText)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If

                'Send Message on a new thread:
                SendMessageParams.ProjectNetworkName = ClientProNetName
                    SendMessageParams.ConnectionName = ClientConnName
                    SendMessageParams.Message = MessageText
                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    End If
                End If

            Else 'This is not an XMessage!
                If Instructions.StartsWith("<XMsgBlk>") Then 'This is an XMessageBlock.
                'Process the received message:
                Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
                XDoc.LoadXml(XmlHeader & vbCrLf & Instructions.Replace("&", "&amp;")) 'Replace "&" with "&amp:" before loading the XML text.
                'NOTE: The message is an <XMsgBlk> - use the ShowXMessages property to determine if the message is shown:
                If ShowXMessages Then
                    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                End If
                'If (MsgType = "XMsg") And ShowXMessages Then
                '    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'ElseIf (MsgType = "XSys") And ShowSysMessages Then
                '    Message.XAddXml(XDoc)   'Add the message to the XMessages window.
                '    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                'End If

                'Process the XMessageBlock:
                Dim XMsgBlkLocn As String
                XMsgBlkLocn = XDoc.GetElementsByTagName("ClientLocn")(0).InnerText
                Select Case XMsgBlkLocn
                    Case "TestLocn" 'Replace this with the required location name.
                        Dim XInfo As Xml.XmlNodeList = XDoc.GetElementsByTagName("XInfo") 'Get the XInfo node list
                        Dim InfoXDoc As New Xml.Linq.XDocument 'Create an XDocument to hold the information contained in XInfo 
                        InfoXDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & vbCrLf & XInfo(0).InnerXml) 'Read the information into InfoXDoc
                        'Add processing instructions here - The information in the InfoXDoc is usually stored in an XDocument in the application or as an XML file in the project.

                    Case Else
                        Message.AddWarning("Unknown XInfo Message location: " & XMsgBlkLocn & vbCrLf)
                End Select
            Else
                Message.XAddText("The message is not an XMessage or XMessageBlock: " & vbCrLf & Instructions & vbCrLf & vbCrLf, "Normal")
            End If
            'Message.XAddText("The message is not an XMessage: " & Instructions & vbCrLf, "Normal")
        End If
    End Sub

    Private Sub ProcessLocalInstructions(ByVal Instructions As String)
        'Process the XMessage instructions locally.

        'If Instructions.StartsWith("<XMsg>") Then 'This is an XMessage set of instructions.
        If Instructions.StartsWith("<XMsg>") Or Instructions.StartsWith("<XSys>") Then 'This is an XMessage set of instructions.
            'Run the received message:
            Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
            XDocLocal.LoadXml(XmlHeader & vbCrLf & Instructions)
            XMsgLocal.Run(XDocLocal, StatusLocal)
        Else 'This is not an XMessage!
            Message.XAddText("The message is not an XMessage: " & Instructions & vbCrLf, "Normal")
        End If
    End Sub

    Private _showXMessages As Boolean = True 'If True, XMessages that are sent or received will be shown in the Messages window.
    Property ShowXMessages As Boolean
        Get
            Return _showXMessages
        End Get
        Set(value As Boolean)
            _showXMessages = value
        End Set
    End Property

    Private _showSysMessages As Boolean = True 'If True, System messages that are sent or received will be shown in the messages window.
    Property ShowSysMessages As Boolean
        Get
            Return _showSysMessages
        End Get
        Set(value As Boolean)
            _showSysMessages = value
        End Set
    End Property

#Region "Library Properties" '==========================================================================================

    Private _libraryFileName As String = "" 'The file name of the selected Document Library.
    Property LibraryFileName As String
        Get
            Return _libraryFileName
        End Get
        Set(value As String)
            _libraryFileName = value
            'txtLibraryFileName.Text = _libraryFileName
        End Set
    End Property

    Private _libraryName As String = "" 'The name of the Document Library.
    Property LibraryName As String
        Get
            Return _libraryName
        End Get
        Set(value As String)
            _libraryName = value
            txtLibraryName.Text = _libraryName
        End Set
    End Property

    Private _libraryDescription As String = "" 'A description of the Document Library.
    Property LibraryDescription As String
        Get
            Return _libraryDescription
        End Get
        Set(value As String)
            _libraryDescription = value
            'txtItemDescription.Text = _libraryDescription
        End Set
    End Property

    Private _libraryCreationDate As DateTime = Now ' Format(Now, "d-MMM-yyyy H:mm:ss") 'The creation date of the Library. 
    Property LibraryCreationDate As DateTime
        Get
            Return _libraryCreationDate
        End Get
        Set(value As DateTime)
            _libraryCreationDate = value
        End Set
    End Property

    Private _libraryLastEditDate As DateTime = Now 'The last edit date of the Library.
    Property LibraryLastEditDate As DateTime
        Get
            Return _libraryLastEditDate
        End Get
        Set(value As DateTime)
            _libraryLastEditDate = value
        End Set
    End Property

    Enum NewItemTypes
        Collection
        RTF
        XML
        TXT
        HTML
        PDF
        XLS
        FolderLink
        XMsg
        XSeq
    End Enum

    Private _newItemType As NewItemTypes = NewItemTypes.Collection
    Property NewItemType As NewItemTypes
        Get
            Return _newItemType
        End Get
        Set(value As NewItemTypes)
            _newItemType = value
            If _newItemType = NewItemTypes.Collection Then
                rbCollection.Checked = True
            ElseIf _newItemType = NewItemTypes.HTML Then
                rbHtml.Checked = True
            ElseIf _newItemType = NewItemTypes.RTF Then
                rbRtf.Checked = True
            ElseIf _newItemType = NewItemTypes.TXT Then
                rbTxt.Checked = True
            ElseIf _newItemType = NewItemTypes.XML Then
                rbXml.Checked = True
            ElseIf _newItemType = NewItemTypes.PDF Then
                rbPdf.Checked = True
            ElseIf _newItemType = NewItemTypes.XLS Then
                rbXls.Checked = True
            ElseIf _newItemType = NewItemTypes.FolderLink Then
                rbFolderLink.Checked = True
            ElseIf _newItemType = NewItemTypes.XMsg Then
                rbXMsg.Checked = True
            ElseIf _newItemType = NewItemTypes.XSeq Then
                rbXSeq.Checked = True
            End If
        End Set
    End Property

    Enum NewItemDisplayTypes
        DefaultDisplay
        NewWindow
        NoDisplay
    End Enum

    Private _newItemDisplay As NewItemDisplayTypes = NewItemDisplayTypes.NewWindow
    Property NewItemDisplay As NewItemDisplayTypes
        Get
            Return _newItemDisplay
        End Get
        Set(value As NewItemDisplayTypes)
            _newItemDisplay = value
            If _newItemDisplay = NewItemDisplayTypes.DefaultDisplay Then
                rbDefaultWindow.Checked = True
            ElseIf _newItemDisplay = NewItemDisplayTypes.NewWindow Then
                rbNewWindow.Checked = True
            ElseIf _newItemDisplay = NewItemDisplayTypes.NoDisplay Then
                rbNoDisplay.Checked = True
            End If
        End Set
    End Property

#End Region 'Library Properties ----------------------------------------------------------------------------------------


    Enum FileTypes
        RTF
        XML
        TXT
        HTML
        PDF
        XLS
        FolderLink
        XMsg
        XSeq
    End Enum

    'File Properties ======================================================================================================

    Private _fileTypeSelection As FileTypes = FileTypes.RTF 'The file type selection (RTF, XML, TXT, HTML ...). May be different from the currently opened file type.
    Property FileTypeSelection As FileTypes
        Get
            Return _fileTypeSelection
        End Get
        Set(value As FileTypes)
            _fileTypeSelection = value
            'cmbDocType.SelectedIndex = cmbDocType.FindStringExact(_fileTypeSelection.ToString)
            txtDocType.Text = _fileTypeSelection.ToString
        End Set
    End Property

    Private _fileType As FileTypes = FileTypes.RTF 'The current file type.
    Property FileType As FileTypes
        Get
            Return _fileType
        End Get
        Set(value As FileTypes)
            _fileType = value
            txtFileType.Text = _fileType.ToString
            txtFileType2.Text = _fileType.ToString
            'cmbDocType.SelectedIndex = cmbDocType.FindStringExact(_fileType.ToString)
            txtDocType.Text = _fileType.ToString
        End Set
    End Property

    Private _fileName As String = "" 'The name of the current file.
    Property FileName As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
            txtFileName2.Text = _fileName
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
            If _fileLocationType = LocationTypes.Project Then
                rbFileInProject.Checked = True
                txtFileDirectory.Enabled = False
                Label2.Enabled = False
            Else
                rbFileInFileSystem.Checked = True
                txtFileDirectory.Enabled = True
                Label2.Enabled = True
            End If
        End Set
    End Property

    Private _fileDirectory As String = "" 'The path of the directory containing the current file.
    Property FileDirectory As String
        Get
            Return _fileDirectory
        End Get
        Set(value As String)
            _fileDirectory = value
            txtFileDirectory.Text = _fileDirectory
        End Set
    End Property

    Private _documentTextChanged As Boolean = False 'If True, the document text has been changed. Prompt to save the changes before they are lost.
    Property DocumentTextChanged As Boolean
        Get
            Return _documentTextChanged
        End Get
        Set(value As Boolean)
            _documentTextChanged = value
        End Set
    End Property

    '----------------------------------------------------------------------------------------------------------------------

    'Last File Properties =================================================================================================

    Private _lastFileType As FileTypes = FileTypes.RTF 'The file type of the last document opened.
    Property LastFileType As FileTypes
        Get
            Return _lastFileType
        End Get
        Set(value As FileTypes)
            _lastFileType = value

        End Set
    End Property

    Private _lastFileName As String = "" 'The name of the last document opened.
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

    Private _lastFileDirectory As String = "" 'The path of the directory containing the last file.
    Property LastFileDirectory As String
        Get
            Return _lastFileDirectory
        End Get
        Set(value As String)
            _lastFileDirectory = value
        End Set
    End Property

    '----------------------------------------------------------------------------------------------------------------------

    'XML File Properties ==================================================================================================

    'FileName
    'FileLocationType
    'FileDirectory
    'LastFileName
    'LastFileLocationType
    'LastFileDirectory
    'TextChanged

    Private _xmlFileName As String = "" 'The file name of the displayed XML file.
    Property XmlFileName As String
        Get
            Return _xmlFileName
        End Get
        Set(value As String)
            _xmlFileName = value
            txtFileName2.Text = _xmlFileName
        End Set
    End Property

    Private _xmlFileLocationType As LocationTypes = LocationTypes.Project 'The location type of the XML Document File. (Either the current project or the file system.)
    Property XmlFileLocationType As LocationTypes
        Get
            Return _xmlFileLocationType
        End Get
        Set(value As LocationTypes)
            _xmlFileLocationType = value
        End Set
    End Property

    Private _xmlFileDirectory As String = "" 'The path of the directory containing the current file.
    Property XmlFileDirectory As String
        Get
            Return _xmlFileDirectory
        End Get
        Set(value As String)
            _xmlFileDirectory = value
            'txtFileDirectory.Text = _xmlFileDirectory
        End Set
    End Property

    Private _lastXmlFileName As String = "" 'The last used XML document name.
    Property LastXmlFileName As String
        Get
            Return _lastXmlFileName
        End Get
        Set(value As String)
            _lastXmlFileName = value
        End Set
    End Property

    Private _lastXmlFileLocationType As LocationTypes = LocationTypes.Project 'The location type of the last XML Document File. (Either the current project or the file system.)
    Property LastXmlFileLocationType As LocationTypes
        Get
            Return _lastXmlFileLocationType
        End Get
        Set(value As LocationTypes)
            _lastXmlFileLocationType = value
        End Set
    End Property

    Private _LastXmlFileDirectory As String = "" 'The path of the directory containing the last XML Document file.
    Property LastXmlFileDirectory As String
        Get
            Return _LastXmlFileDirectory
        End Get
        Set(value As String)
            _LastXmlFileDirectory = value
            'txtFileDirectory.Text = _xmlFileDirectory
        End Set
    End Property

    'TO BE DELETED?:
    Private _xmlTextChanged As Boolean = False 'If True, the XML text has been changed. Prompt to save the changes before they are lost.
    Property XmlTextChanged As Boolean
        Get
            Return _xmlTextChanged
        End Get
        Set(value As Boolean)
            _xmlTextChanged = value
        End Set
    End Property

    '----------------------------------------------------------------------------------------------------------------------

    'RTF File Properties ==================================================================================================
    'FileName
    'FileLocationType
    'FileDirectory
    'LastFileName
    'LastFileLocationType
    'LastFileDirectory
    'TextChanged

    Private _rtfFileName As String = "" 'The file name of the displayed RTF file.
    Property RtfFileName As String
        Get
            Return _rtfFileName
        End Get
        Set(value As String)
            _rtfFileName = value
            'txtRtfPath.Text = _rtfFilePath
        End Set
    End Property

    Private _rtfFileLocationType As LocationTypes = LocationTypes.Project 'The location type of the RTF Document File. (Either the current project or the file system.)
    Property RtfFileLocationType As LocationTypes
        Get
            Return _rtfFileLocationType
        End Get
        Set(value As LocationTypes)
            _rtfFileLocationType = value
        End Set
    End Property

    Private _rtfFileDirectory As String = "" 'The path of the directory containing the last RTF Document file.
    Property RtfFileDirectory As String
        Get
            Return _rtfFileDirectory
        End Get
        Set(value As String)
            _rtfFileDirectory = value
            'txtFileDirectory.Text = _xmlFileDirectory
        End Set
    End Property

    Private _lastRtfFileName As String = "" 'The last used RTF file name.
    Property LastRtfFileName As String
        Get
            Return _lastRtfFileName
        End Get
        Set(value As String)
            _lastRtfFileName = value
        End Set
    End Property

    Private _lastRtfFileLocationType As LocationTypes = LocationTypes.Project 'The location type of the last RTF Document File. (Either the current project or the file system.)
    Property LastRtfFileLocationType As LocationTypes
        Get
            Return _lastRtfFileLocationType
        End Get
        Set(value As LocationTypes)
            _lastRtfFileLocationType = value
        End Set
    End Property

    Private _LastRtfFileDirectory As String = "" 'The path of the directory containing the last RTF Document file.
    Property LastRtfFileDirectory As String
        Get
            Return _LastRtfFileDirectory
        End Get
        Set(value As String)
            _LastRtfFileDirectory = value
            'txtFileDirectory.Text = _xmlFileDirectory
        End Set
    End Property

    Private _rtfTextChanged As Boolean = False 'If True, the RTF text has been changed. Prompt to save the changes before they are lost.
    Property RtfTextChanged As Boolean
        Get
            Return _rtfTextChanged
        End Get
        Set(value As Boolean)
            _rtfTextChanged = value
        End Set
    End Property

    '----------------------------------------------------------------------------------------------------------------------


    'TXT File Properties ==================================================================================================
    'FileName
    'FileLocationType
    'FileDirectory
    'LastFileName
    'LastFileLocationType
    'LastFileDirectory
    'TextChanged

    Private _textFileName As String = "" 'The file name of the displayed RTF file.
    Property TextFileName As String
        Get
            Return _textFileName
        End Get
        Set(value As String)
            _textFileName = value
            'txtRtfPath.Text = _rtfFilePath
        End Set
    End Property

    Private _textFileLocationType As LocationTypes = LocationTypes.Project 'The location type of the RTF Document File. (Either the current project or the file system.)
    Property TextFileLocationType As LocationTypes
        Get
            Return _textFileLocationType
        End Get
        Set(value As LocationTypes)
            _textFileLocationType = value
        End Set
    End Property

    Private _textFileDirectory As String = "" 'The path of the directory containing the last RTF Document file.
    Property TextFileDirectory As String
        Get
            Return _textFileDirectory
        End Get
        Set(value As String)
            _textFileDirectory = value
            'txtFileDirectory.Text = _xmlFileDirectory
        End Set
    End Property

    Private _lastTextFileName As String = "" 'The last used TXT (or ASC) file name.
    Property LastTextFileName As String
        Get
            Return _lastTextFileName
        End Get
        Set(value As String)
            _lastTextFileName = value
        End Set
    End Property

    Private _lastTextFileLocationType As LocationTypes = LocationTypes.Project 'The location type of the last txt Document File. (Either the current project or the file system.)
    Property LastTextFileLocationType As LocationTypes
        Get
            Return _lastTextFileLocationType
        End Get
        Set(value As LocationTypes)
            _lastTextFileLocationType = value
        End Set
    End Property

    Private _LastTextFileDirectory As String = "" 'The path of the directory containing the last TXT Document file.
    Property LastTextFileDirectory As String
        Get
            Return _LastTextFileDirectory
        End Get
        Set(value As String)
            _LastTextFileDirectory = value
            'txtFileDirectory.Text = _xmlFileDirectory
        End Set
    End Property

    Private _txtTextChanged As Boolean = False 'If True, the RTF text has been changed. Prompt to save the changes before they are lost.
    Property TxtTextChanged As Boolean
        Get
            Return _txtTextChanged
        End Get
        Set(value As Boolean)
            _txtTextChanged = value
        End Set
    End Property

    '----------------------------------------------------------------------------------------------------------------------

    'HTML File Properties =================================================================================================
    'FileName
    'FileLocationType
    'FileDirectory
    'LastFileName
    'LastFileLocationType
    'LastFileDirectory
    'TextChanged

    Private _htmlFileName As String = "" 'The file name of the displayed RTF file.
    Property HtmlFileName As String
        Get
            Return _htmlFileName
        End Get
        Set(value As String)
            _htmlFileName = value
            'txtRtfPath.Text = _rtfFilePath
        End Set
    End Property

    Private _htmlFileLocationType As LocationTypes = LocationTypes.Project 'The location type of the RTF Document File. (Either the current project or the file system.)
    Property HtmlFileLocationType As LocationTypes
        Get
            Return _htmlFileLocationType
        End Get
        Set(value As LocationTypes)
            _htmlFileLocationType = value
        End Set
    End Property

    Private _htmlFileDirectory As String = "" 'The path of the directory containing the last RTF Document file.
    Property HtmlFileDirectory As String
        Get
            Return _htmlFileDirectory
        End Get
        Set(value As String)
            _htmlFileDirectory = value
            'txtFileDirectory.Text = _xmlFileDirectory
        End Set
    End Property

    Private _lastHtmlFileName As String = "" 'The last used RTF file name.
    Property LastHtmlFileName As String
        Get
            Return _lastHtmlFileName
        End Get
        Set(value As String)
            _lastHtmlFileName = value
        End Set
    End Property

    Private _lastHtmlFileLocationType As LocationTypes = LocationTypes.Project 'The location type of the last RTF Document File. (Either the current project or the file system.)
    Property LastHtmlFileLocationType As LocationTypes
        Get
            Return _lastHtmlFileLocationType
        End Get
        Set(value As LocationTypes)
            _lastHtmlFileLocationType = value
        End Set
    End Property

    Private _LastHtmlFileDirectory As String = "" 'The path of the directory containing the last RTF Document file.
    Property LastHtmlFileDirectory As String
        Get
            Return _LastHtmlFileDirectory
        End Get
        Set(value As String)
            _LastHtmlFileDirectory = value
            'txtFileDirectory.Text = _xmlFileDirectory
        End Set
    End Property

    Private _htmlTextChanged As Boolean = False 'If True, the RTF text has been changed. Prompt to save the changes before they are lost.
    Property HtmlTextChanged As Boolean
        Get
            Return _htmlTextChanged
        End Get
        Set(value As Boolean)
            _htmlTextChanged = value
        End Set
    End Property

    '----------------------------------------------------------------------------------------------------------------------

    Private _closedFormNo As Integer 'Temporarily holds the number of the form that is being closed. 
    Property ClosedFormNo As Integer
        Get
            Return _closedFormNo
        End Get
        Set(value As Integer)
            _closedFormNo = value
        End Set
    End Property


    'Private _startPageFileName As String = "" 'The file name of the html document displayed in the Start Page tab.
    'Public Property StartPageFileName As String
    '    Get
    '        Return _startPageFileName
    '    End Get
    '    Set(value As String)
    '        _startPageFileName = value
    '    End Set
    'End Property

    Private _workflowFileName As String = "" 'The file name of the html document displayed in the Workflow tab.
    Public Property WorkflowFileName As String
        Get
            Return _workflowFileName
        End Get
        Set(value As String)
            _workflowFileName = value
        End Set
    End Property


#End Region 'Properties -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Process XML Files - Read and write XML files." '=====================================================================================================================================

    Private Sub SaveFormSettings()
        'Save the form settings in an XML document.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Form settings for Main form.-->
                           <FormSettings>
                               <Left><%= Me.Left %></Left>
                               <Top><%= Me.Top %></Top>
                               <Width><%= Me.Width %></Width>
                               <Height><%= Me.Height %></Height>
                               <AdvlNetworkAppPath><%= AdvlNetworkAppPath %></AdvlNetworkAppPath>
                               <AdvlNetworkExePath><%= AdvlNetworkExePath %></AdvlNetworkExePath>
                               <ShowXMessages><%= ShowXMessages %></ShowXMessages>
                               <ShowSysMessages><%= ShowSysMessages %></ShowSysMessages>
                               <!---->
                               <LibraryFileName><%= LibraryFileName %></LibraryFileName>
                               <NewItemType><%= NewItemType.ToString %></NewItemType>
                               <NewItemDisplay><%= NewItemDisplay.ToString %></NewItemDisplay>
                               <LastXmlFileName><%= LastXmlFileName %></LastXmlFileName>
                               <LastXmlFileLocationType><%= LastXmlFileLocationType %></LastXmlFileLocationType>
                               <LastXmlFileDirectory><%= LastXmlFileDirectory %></LastXmlFileDirectory>
                               <LastRtfFileName><%= LastRtfFileName %></LastRtfFileName>
                               <LastRtfFileLocationType><%= LastRtfFileLocationType %></LastRtfFileLocationType>
                               <LastRtfFileDirectory><%= LastRtfFileDirectory %></LastRtfFileDirectory>
                               <LastTxtFileName><%= LastTextFileName %></LastTxtFileName>
                               <LastTxtFileLocationType><%= LastTextFileLocationType %></LastTxtFileLocationType>
                               <LastTxtFileDirectory><%= LastTextFileDirectory %></LastTxtFileDirectory>
                               <FileTypeSelection><%= FileTypeSelection.ToString %></FileTypeSelection>
                               <SettingsFile><%= txtSettingsFileName.Text %></SettingsFile>
                               <SelectedTabIndex><%= TabControl1.SelectedIndex %></SelectedTabIndex>
                               <SelectedLibraryTabIndex><%= TabControl3.SelectedIndex %></SelectedLibraryTabIndex>
                               <HtmlCodeView><%= rbCodeView.Checked %></HtmlCodeView>
                               <LibraryTabSplitDistance><%= SplitContainer1.SplitterDistance %></LibraryTabSplitDistance>
                               <!---->
                           </FormSettings>

        '<MsgServiceAppPath><%= MsgServiceAppPath %></MsgServiceAppPath>
        '<MsgServiceExePath><%= MsgServiceExePath %></MsgServiceExePath>

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)
    End Sub

    Private Sub RestoreFormSettings()
        'Read the form settings from an XML document.

        'Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & "_" & Me.Text & ".xml"
        Dim SettingsFileName As String = "FormSettings_" & ApplicationInfo.Name & " - Main.xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore form position and size:
            If Settings.<FormSettings>.<Left>.Value <> Nothing Then Me.Left = Settings.<FormSettings>.<Left>.Value
            If Settings.<FormSettings>.<Top>.Value <> Nothing Then Me.Top = Settings.<FormSettings>.<Top>.Value
            If Settings.<FormSettings>.<Height>.Value <> Nothing Then Me.Height = Settings.<FormSettings>.<Height>.Value
            If Settings.<FormSettings>.<Width>.Value <> Nothing Then Me.Width = Settings.<FormSettings>.<Width>.Value

            'If Settings.<FormSettings>.<MsgServiceAppPath>.Value <> Nothing Then MsgServiceAppPath = Settings.<FormSettings>.<MsgServiceAppPath>.Value
            'If Settings.<FormSettings>.<MsgServiceExePath>.Value <> Nothing Then MsgServiceExePath = Settings.<FormSettings>.<MsgServiceExePath>.Value
            If Settings.<FormSettings>.<AdvlNetworkAppPath>.Value <> Nothing Then AdvlNetworkAppPath = Settings.<FormSettings>.<AdvlNetworkAppPath>.Value
            If Settings.<FormSettings>.<AdvlNetworkExePath>.Value <> Nothing Then AdvlNetworkExePath = Settings.<FormSettings>.<AdvlNetworkExePath>.Value

            If Settings.<FormSettings>.<ShowXMessages>.Value <> Nothing Then ShowXMessages = Settings.<FormSettings>.<ShowXMessages>.Value
            If Settings.<FormSettings>.<ShowSysMessages>.Value <> Nothing Then ShowSysMessages = Settings.<FormSettings>.<ShowSysMessages>.Value

            'Add code to read other saved setting here:
            If Settings.<FormSettings>.<SelectedTabIndex>.Value <> Nothing Then TabControl1.SelectedIndex = Settings.<FormSettings>.<SelectedTabIndex>.Value

            If Settings.<FormSettings>.<SelectedLibraryTabIndex>.Value <> Nothing Then TabControl3.SelectedIndex = Settings.<FormSettings>.<SelectedLibraryTabIndex>.Value

            If Settings.<FormSettings>.<LibraryFileName>.Value <> Nothing Then
                LibraryFileName = Settings.<FormSettings>.<LibraryFileName>.Value
                OpenLibrary(LibraryFileName)
            End If

            If Settings.<FormSettings>.<NewItemType>.Value <> Nothing Then
                Select Case Settings.<FormSettings>.<NewItemType>.Value
                    Case "Collection"
                        rbCollection.Checked = True
                    Case "HTML"
                        rbHtml.Checked = True
                    Case "XML"
                        rbXml.Checked = True
                    Case "TXT"
                        rbTxt.Checked = True
                    Case "RTF"
                        rbRtf.Checked = True
                End Select
            End If

            If Settings.<FormSettings>.<NewItemDisplay>.Value <> Nothing Then
                Select Case Settings.<FormSettings>.<NewItemDisplay>.Value
                    Case "DefaultDisplay"
                        rbDefaultWindow.Checked = True
                    Case "NewWindow"
                        rbNewWindow.Checked = True
                    Case "NoDisplay"
                        rbNoDisplay.Checked = True

                End Select
            End If

            If Settings.<FormSettings>.<LastXmlFileName>.Value <> Nothing Then LastXmlFileName = Settings.<FormSettings>.<LastXmlFileName>.Value
            If Settings.<FormSettings>.<LastXmlFileLocationType>.Value <> Nothing Then
                Select Case Settings.<FormSettings>.<LastXmlFileLocationType>.Value
                    Case "FileSystem"
                        LastXmlFileLocationType = LocationTypes.FileSystem
                    Case "Project"
                        LastXmlFileLocationType = LocationTypes.Project
                End Select
            End If
            If Settings.<FormSettings>.<LastXmlFileDirectory>.Value <> Nothing Then LastXmlFileDirectory = Settings.<FormSettings>.<LastXmlFileDirectory>.Value

            If Settings.<FormSettings>.<LastRtfFileName>.Value <> Nothing Then LastRtfFileName = Settings.<FormSettings>.<LastRtfFileName>.Value
            If Settings.<FormSettings>.<LastRtfFileLocationType>.Value <> Nothing Then
                Select Case Settings.<FormSettings>.<LastRtfFileLocationType>.Value
                    Case "FileSystem"
                        LastRtfFileLocationType = LocationTypes.FileSystem
                    Case "Project"
                        LastRtfFileLocationType = LocationTypes.Project
                End Select
            End If
            If Settings.<FormSettings>.<LastRtfFileDirectory>.Value <> Nothing Then LastRtfFileDirectory = Settings.<FormSettings>.<LastRtfFileDirectory>.Value

            If Settings.<FormSettings>.<LastTxtFileName>.Value <> Nothing Then LastTextFileName = Settings.<FormSettings>.<LastTxtFileName>.Value
            If Settings.<FormSettings>.<LastTxtFileLocationType>.Value <> Nothing Then
                Select Case Settings.<FormSettings>.<LastTxtFileLocationType>.Value
                    Case "FileSystem"
                        LastTextFileLocationType = LocationTypes.FileSystem
                    Case "Project"
                        LastTextFileLocationType = LocationTypes.Project
                End Select
            End If
            If Settings.<FormSettings>.<LastTxtFileDirectory>.Value <> Nothing Then LastTextFileDirectory = Settings.<FormSettings>.<LastTxtFileDirectory>.Value

            If Settings.<FormSettings>.<FileTypeSelection>.Value <> Nothing Then 'FileType = Settings.<FormSettings>.<FileTypeSelection>.Value
                Select Case Settings.<FormSettings>.<FileTypeSelection>.Value
                    Case "XML"
                        FileTypeSelection = FileTypes.XML
                    Case "RTF"
                        FileTypeSelection = FileTypes.RTF
                    Case "TXT"
                        FileTypeSelection = FileTypes.TXT
                    Case "HTM"
                        FileTypeSelection = FileTypes.HTML
                    Case "HTML"
                        FileTypeSelection = FileTypes.HTML
                End Select
            End If

            If Settings.<FormSettings>.<SettingsFile>.Value <> Nothing Then
                txtSettingsFileName.Text = Settings.<FormSettings>.<SettingsFile>.Value
                Dim XmlDoc As System.Xml.Linq.XDocument
                Project.DataLocn.ReadXmlData(txtSettingsFileName.Text, XmlDoc)
                ReadSettings(XmlDoc)
            End If

            If Settings.<FormSettings>.<HtmlCodeView>.Value <> Nothing Then
                If Settings.<FormSettings>.<HtmlCodeView>.Value = "true" Then
                    rbCodeView.Checked = True
                Else
                    rbWebView.Checked = True
                End If
            End If

            If Settings.<FormSettings>.<LibraryTabSplitDistance>.Value <> Nothing Then SplitContainer1.SplitterDistance = Settings.<FormSettings>.<LibraryTabSplitDistance>.Value

            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        'Dim MinWidthVisible As Integer = 48 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        'Dim MinWidthVisible As Integer = 128 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        'Dim MinHeightVisible As Integer = 48 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.
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

    Private Sub ReadApplicationInfo()
        'Read the Application Information.

        If ApplicationInfo.FileExists Then
            ApplicationInfo.ReadFile()
        Else
            'There is no Application_Info_ADVL_2.xml file.
            DefaultAppProperties() 'Create a new Application Info file with default application properties.
            ApplicationInfo.WriteFile() 'Write the file now. The file information may be used by other applications.
        End If
    End Sub

    Private Sub DefaultAppProperties()
        'These properties will be saved in the Application_Info.xml file in the application directory.
        'If this file is deleted, it will be re-created using these default application properties.

        'Change this to show your application Name, Description and Creation Date.
        ApplicationInfo.Name = "ADVL_Document_Library_1"

        'ApplicationInfo.ApplicationDir is set when the application is started.
        ApplicationInfo.ExecutablePath = Application.ExecutablePath

        ApplicationInfo.Description = "Application for creating, storing and editing documents in different formats including text, rtf, xml and html."
        ApplicationInfo.CreationDate = "15-Jan-2018 12:00:00"

        'Author -----------------------------------------------------------------------------------------------------------
        'Change this to show your Name, Description and Contact information.
        ApplicationInfo.Author.Name = "Signalworks Pty Ltd"
        ApplicationInfo.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        ApplicationInfo.Author.Contact = "http://www.andorville.com.au/"

        'File Associations: -----------------------------------------------------------------------------------------------
        'Add any file associations here.
        'The file extension and a description of files that can be opened by this application are specified.
        'The example below specifies a coordinate system parameter file type with the file extension .ADVLCoord.
        'Dim Assn1 As New ADVL_System_Utilities.FileAssociation
        'Assn1.Extension = "ADVLCoord"
        'Assn1.Description = "Andorville™ software coordinate system parameter file"
        'ApplicationInfo.FileAssociations.Add(Assn1)

        'Version ----------------------------------------------------------------------------------------------------------
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        'Copyright --------------------------------------------------------------------------------------------------------
        'Add your copyright information here.
        ApplicationInfo.Copyright.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.Copyright.PublicationYear = "2018"

        'Trademarks -------------------------------------------------------------------------------------------------------
        'Add your trademark information here.
        Dim Trademark1 As New ADVL_Utilities_Library_1.Trademark
        Trademark1.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark1.Text = "Andorville"
        Trademark1.Registered = False
        Trademark1.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark1)
        Dim Trademark2 As New ADVL_Utilities_Library_1.Trademark
        Trademark2.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark2.Text = "AL-H7"
        Trademark2.Registered = False
        Trademark2.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark2)
        Dim Trademark3 As New ADVL_Utilities_Library_1.Trademark
        Trademark3.OwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        Trademark3.Text = "AL-M7"
        Trademark3.Registered = False
        Trademark3.GenericTerm = "software"
        ApplicationInfo.Trademarks.Add(Trademark3)

        'License -------------------------------------------------------------------------------------------------------
        'Add your license information here.
        ApplicationInfo.License.CopyrightOwnerName = "Signalworks Pty Ltd, ABN 26 066 681 598"
        ApplicationInfo.License.PublicationYear = "2018"

        'License Links:
        'http://choosealicense.com/
        'http://www.apache.org/licenses/
        'http://opensource.org/

        'Apache License 2.0 ---------------------------------------------
        ApplicationInfo.License.Code = ADVL_Utilities_Library_1.License.Codes.Apache_License_2_0
        ApplicationInfo.License.Notice = ApplicationInfo.License.ApacheLicenseNotice 'Get the pre-defined Aapche license notice.
        ApplicationInfo.License.Text = ApplicationInfo.License.ApacheLicenseText     'Get the pre-defined Apache license text.

        'Code to use other pre-defined license types is shown below:

        'GNU General Public License, version 3 --------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.GNU_GPL_V3_0
        'ApplicationInfo.License.Notice = 'Add the License Notice to ADVL_Utilities_Library_1 License class.
        'ApplicationInfo.License.Text = 'Add the License Text to ADVL_Utilities_Library_1 License class.

        'The MIT License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.MIT_License
        'ApplicationInfo.License.Notice = ApplicationInfo.License.MITLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.MITLicenseText

        'No License Specified -------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.None
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'The Unlicense --------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.The_Unlicense
        'ApplicationInfo.License.Notice = ApplicationInfo.License.UnLicenseNotice
        'ApplicationInfo.License.Text = ApplicationInfo.License.UnLicenseText

        'Unknown License ------------------------------------------------
        'ApplicationInfo.License.Type = ADVL_Utilities_Library_1.License.Types.Unknown
        'ApplicationInfo.License.Notice = ""
        'ApplicationInfo.License.Text = ""

        'Source Code: --------------------------------------------------------------------------------------------------
        'Add your source code information here if required.
        'THIS SECTION WILL BE UPDATED TO ALLOW A GITHUB LINK.
        ApplicationInfo.SourceCode.Language = "Visual Basic 2015"
        ApplicationInfo.SourceCode.FileName = ""
        ApplicationInfo.SourceCode.FileSize = 0
        ApplicationInfo.SourceCode.FileHash = ""
        ApplicationInfo.SourceCode.WebLink = ""
        ApplicationInfo.SourceCode.Contact = ""
        ApplicationInfo.SourceCode.Comments = ""

        'ModificationSummary: -----------------------------------------------------------------------------------------
        'Add any source code modification here is required.
        ApplicationInfo.ModificationSummary.BaseCodeName = ""
        ApplicationInfo.ModificationSummary.BaseCodeDescription = ""
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Major = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Minor = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Build = 0
        ApplicationInfo.ModificationSummary.BaseCodeVersion.Revision = 0
        ApplicationInfo.ModificationSummary.Description = "This is the first released version of the application. No earlier base code used."

        'Library List: ------------------------------------------------------------------------------------------------
        'Add the ADVL_Utilties_Library_1 library:
        Dim NewLib As New ADVL_Utilities_Library_1.LibrarySummary
        NewLib.Name = "ADVL_System_Utilities"
        NewLib.Description = "System Utility classes used in Andorville™ software development system applications"
        NewLib.CreationDate = "7-Jan-2016 12:00:00"
        NewLib.LicenseNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598" & vbCrLf &
                               vbCrLf &
                               "Licensed under the Apache License, Version 2.0 (the ""License"");" & vbCrLf &
                               "you may not use this file except in compliance with the License." & vbCrLf &
                               "You may obtain a copy of the License at" & vbCrLf &
                               vbCrLf &
                               "http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf &
                               vbCrLf &
                               "Unless required by applicable law or agreed to in writing, software" & vbCrLf &
                               "distributed under the License is distributed on an ""AS IS"" BASIS," & vbCrLf &
                               "WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied." & vbCrLf &
                               "See the License for the specific language governing permissions and" & vbCrLf &
                               "limitations under the License." & vbCrLf

        NewLib.CopyrightNotice = "Copyright 2016 Signalworks Pty Ltd, ABN 26 066 681 598"

        NewLib.Version.Major = 1
        NewLib.Version.Minor = 0
        NewLib.Version.Build = 1
        NewLib.Version.Revision = 0

        NewLib.Author.Name = "Signalworks Pty Ltd"
        NewLib.Author.Description = "Signalworks Pty Ltd" & vbCrLf &
            "Australian Proprietary Company" & vbCrLf &
            "ABN 26 066 681 598" & vbCrLf &
            "Registration Date 05/10/1994"

        NewLib.Author.Contact = "http://www.andorville.com.au/"

        Dim NewClass1 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass1.Name = "ZipComp"
        NewClass1.Description = "The ZipComp class is used to compress files into and extract files from a zip file."
        NewLib.Classes.Add(NewClass1)
        Dim NewClass2 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass2.Name = "XSequence"
        NewClass2.Description = "The XSequence class is used to run an XML property sequence (XSequence) file. XSequence files are used to record and replay processing sequences in Andorville™ software applications."
        NewLib.Classes.Add(NewClass2)
        Dim NewClass3 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass3.Name = "XMessage"
        NewClass3.Description = "The XMessage class is used to read an XML Message (XMessage). An XMessage is a simplified XSequence used to exchange information between Andorville™ software applications."
        NewLib.Classes.Add(NewClass3)
        Dim NewClass4 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass4.Name = "Location"
        NewClass4.Description = "The Location class consists of properties and methods to store data in a location, which is either a directory or archive file."
        NewLib.Classes.Add(NewClass4)
        Dim NewClass5 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass5.Name = "Project"
        NewClass5.Description = "An Andorville™ software application can store data within one or more projects. Each project stores a set of related data files. The Project class contains properties and methods used to manage a project."
        NewLib.Classes.Add(NewClass5)
        Dim NewClass6 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass6.Name = "ProjectSummary"
        NewClass6.Description = "ProjectSummary stores a summary of a project."
        NewLib.Classes.Add(NewClass6)
        Dim NewClass7 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass7.Name = "DataFileInfo"
        NewClass7.Description = "The DataFileInfo class stores information about a data file."
        NewLib.Classes.Add(NewClass7)
        Dim NewClass8 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass8.Name = "Message"
        NewClass8.Description = "The Message class contains text properties and methods used to display messages in an Andorville™ software application."
        NewLib.Classes.Add(NewClass8)
        Dim NewClass9 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass9.Name = "ApplicationSummary"
        NewClass9.Description = "The ApplicationSummary class stores a summary of an Andorville™ software application."
        NewLib.Classes.Add(NewClass9)
        Dim NewClass10 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass10.Name = "LibrarySummary"
        NewClass10.Description = "The LibrarySummary class stores a summary of a software library used by an application."
        NewLib.Classes.Add(NewClass10)
        Dim NewClass11 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass11.Name = "ClassSummary"
        NewClass11.Description = "The ClassSummary class stores a summary of a class contained in a software library."
        NewLib.Classes.Add(NewClass11)
        Dim NewClass12 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass12.Name = "ModificationSummary"
        NewClass12.Description = "The ModificationSummary class stores a summary of any modifications made to an application or library."
        NewLib.Classes.Add(NewClass12)
        Dim NewClass13 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass13.Name = "ApplicationInfo"
        NewClass13.Description = "The ApplicationInfo class stores information about an Andorville™ software application."
        NewLib.Classes.Add(NewClass13)
        Dim NewClass14 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass14.Name = "Version"
        NewClass14.Description = "The Version class stores application, library or project version information."
        NewLib.Classes.Add(NewClass14)
        Dim NewClass15 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass15.Name = "Author"
        NewClass15.Description = "The Author class stores information about an Author."
        NewLib.Classes.Add(NewClass15)
        Dim NewClass16 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass16.Name = "FileAssociation"
        NewClass16.Description = "The FileAssociation class stores the file association extension and description. An application can open files on its file association list."
        NewLib.Classes.Add(NewClass16)
        Dim NewClass17 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass17.Name = "Copyright"
        NewClass17.Description = "The Copyright class stores copyright information."
        NewLib.Classes.Add(NewClass17)
        Dim NewClass18 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass18.Name = "License"
        NewClass18.Description = "The License class stores license information."
        NewLib.Classes.Add(NewClass18)
        Dim NewClass19 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass19.Name = "SourceCode"
        NewClass19.Description = "The SourceCode class stores information about the source code for the application."
        NewLib.Classes.Add(NewClass19)
        Dim NewClass20 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass20.Name = "Usage"
        NewClass20.Description = "The Usage class stores information about application or project usage."
        NewLib.Classes.Add(NewClass20)
        Dim NewClass21 As New ADVL_Utilities_Library_1.ClassSummary
        NewClass21.Name = "Trademark"
        NewClass21.Description = "The Trademark class stored information about a trademark used by the author of an application or data."
        NewLib.Classes.Add(NewClass21)

        ApplicationInfo.Libraries.Add(NewLib)

        'Add other library information here: --------------------------------------------------------------------------

    End Sub

    'Save the form settings if the form is being minimised:
    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = &H112 Then 'SysCommand
            If m.WParam.ToInt32 = &HF020 Then 'Form is being minimised
                SaveFormSettings()
            End If
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub SaveProjectSettings()
        'Save the project settings in an XML file.
        'Add any Project Settings to be saved into the settingsData XDocument.
        Dim settingsData = <?xml version="1.0" encoding="utf-8"?>
                           <!---->
                           <!--Project settings for ADVL_Coordinates_1 application.-->
                           <ProjectSettings>
                           </ProjectSettings>

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & ".xml"
        Project.SaveXmlSettings(SettingsFileName, settingsData)

    End Sub

    Private Sub RestoreProjectSettings()
        'Restore the project settings from an XML document.

        Dim SettingsFileName As String = "ProjectSettings_" & ApplicationInfo.Name & "_" & ".xml"

        If Project.SettingsFileExists(SettingsFileName) Then
            Dim Settings As System.Xml.Linq.XDocument
            Project.ReadXmlSettings(SettingsFileName, Settings)

            If IsNothing(Settings) Then 'There is no Settings XML data.
                Exit Sub
            End If

            'Restore a Project Setting example:
            If Settings.<ProjectSettings>.<Setting1>.Value = Nothing Then
                'Project setting not saved.
                'Setting1 = ""
            Else
                'Setting1 = Settings.<ProjectSettings>.<Setting1>.Value
            End If

            'Continue restoring saved settings.

        End If

    End Sub

#End Region 'Process XML Files ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Display Methods - Code used to display this form." '============================================================================================================================

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Set the Application Directory path: ------------------------------------------------
        Project.ApplicationDir = My.Application.Info.DirectoryPath.ToString

        'Read the Application Information file: ---------------------------------------------
        ApplicationInfo.ApplicationDir = My.Application.Info.DirectoryPath.ToString 'Set the Application Directory property
        'Get the Application Version Information:
        ApplicationInfo.Version.Major = My.Application.Info.Version.Major
        ApplicationInfo.Version.Minor = My.Application.Info.Version.Minor
        ApplicationInfo.Version.Build = My.Application.Info.Version.Build
        ApplicationInfo.Version.Revision = My.Application.Info.Version.Revision

        If ApplicationInfo.ApplicationLocked Then
            MessageBox.Show("The application is locked. If the application is not already in use, remove the 'Application_Info.lock file from the application directory: " & ApplicationInfo.ApplicationDir, "Notice", MessageBoxButtons.OK)
            Dim dr As System.Windows.Forms.DialogResult
            dr = MessageBox.Show("Press 'Yes' to unlock the application", "Notice", MessageBoxButtons.YesNo)
            If dr = System.Windows.Forms.DialogResult.Yes Then
                ApplicationInfo.UnlockApplication()
            Else
                Application.Exit()
            End If
        End If

        ReadApplicationInfo()

        'Read the Application Usage information: --------------------------------------------
        ApplicationUsage.StartTime = Now
        ApplicationUsage.SaveLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        ApplicationUsage.SaveLocn.Path = Project.ApplicationDir
        ApplicationUsage.RestoreUsageInfo()

        'Restore Project information: -------------------------------------------------------
        Project.Application.Name = ApplicationInfo.Name

        'Set up Message object:
        Message.ApplicationName = ApplicationInfo.Name

        'Set up a temporary initial settings location:
        Dim TempLocn As New ADVL_Utilities_Library_1.FileLocation
        TempLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        TempLocn.Path = ApplicationInfo.ApplicationDir
        Message.SettingsLocn = TempLocn

        Me.Show() 'Show this form before showing the Message form - This will show the App icon on top in the TaskBar.

        Message.AddText("------------------- Starting Application: ADVL Document Library ---------------------- " & vbCrLf, "Heading")
        Message.AddText("Application usage: Total duration = " & Format(ApplicationUsage.TotalDuration.TotalHours, "#.##") & " hours" & vbCrLf, "Normal")

        'https://msdn.microsoft.com/en-us/library/z2d603cy(v=vs.80).aspx#Y550
        'Process any command line arguments:
        Try
            For Each s As String In My.Application.CommandLineArgs
                Message.Add("Command line argument: " & vbCrLf)
                Message.AddXml(s & vbCrLf & vbCrLf)
                InstrReceived = s
            Next
        Catch ex As Exception
            Message.AddWarning("Error processing command line arguments: " & ex.Message & vbCrLf)
        End Try

        If ProjectSelected = False Then
            'Read the Settings Location for the last project used:
            Project.ReadLastProjectInfo()
            'The Last_Project_Info.xml file contains:
            '  Project Name and Description. Settings Location Type and Settings Location Path.
            Message.Add("Last project details:" & vbCrLf)
            Message.Add("Project Type:  " & Project.Type.ToString & vbCrLf)
            Message.Add("Project Path:  " & Project.Path & vbCrLf)

            'At this point read the application start arguments, if any.
            'The selected project may be changed here.

            'Check if the project is locked:
            If Project.ProjectLocked Then
                Message.AddWarning("The project is locked: " & Project.Name & vbCrLf)
                Dim dr As System.Windows.Forms.DialogResult
                dr = MessageBox.Show("Press 'Yes' to unlock the project", "Notice", MessageBoxButtons.YesNo)
                If dr = System.Windows.Forms.DialogResult.Yes Then
                    Project.UnlockProject()
                    Message.AddWarning("The project has been unlocked: " & Project.Name & vbCrLf)
                    'Read the Project Information file: -------------------------------------------------
                    Message.Add("Reading project info." & vbCrLf)
                    Project.ReadProjectInfoFile()     'Read the file in the SettingsLocation: ADVL_Project_Info.xml
                    Project.ReadParameters()
                    Project.ReadParentParameters()
                    If Project.ParentParameterExists("ProNetName") Then
                        Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                        ProNetName = Project.Parameter("ProNetName").Value
                    Else
                        ProNetName = Project.GetParameter("ProNetName")
                    End If
                    If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                        Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                        ProNetPath = Project.Parameter("ProNetPath").Value
                    Else
                        ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
                    End If
                    Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

                    Project.LockProject() 'Lock the project while it is open in this application.
                    'Set the project start time. This is used to track project usage.
                    Project.Usage.StartTime = Now
                    ApplicationInfo.SettingsLocn = Project.SettingsLocn
                    'Set up the Message object:
                    Message.SettingsLocn = Project.SettingsLocn
                    Message.Show() 'Added 18May19
                Else
                    'Continue without any project selected.
                    Project.Name = ""
                    Project.Type = ADVL_Utilities_Library_1.Project.Types.None
                    Project.Description = ""
                    Project.SettingsLocn.Path = ""
                    Project.DataLocn.Path = ""
                End If

            Else
                'Read the Project Information file: -------------------------------------------------
                Message.Add("Reading project info." & vbCrLf)
                Project.ReadProjectInfoFile()    'Read the file in the SettingsLocation: ADVL_Project_Info.xml
                Project.ReadParameters()
                Project.ReadParentParameters()
                If Project.ParentParameterExists("ProNetName") Then
                    Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                    ProNetName = Project.Parameter("ProNetName").Value
                Else
                    ProNetName = Project.GetParameter("ProNetName")
                End If
                If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                    Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                    ProNetPath = Project.Parameter("ProNetPath").Value
                Else
                    ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
                End If
                Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

                Project.LockProject() 'Lock the project while it is open in this application.
                'Set the project start time. This is used to track project usage.
                Project.Usage.StartTime = Now
                ApplicationInfo.SettingsLocn = Project.SettingsLocn
                'Set up the Message object:
                Message.SettingsLocn = Project.SettingsLocn
                Message.Show() 'Added 18May19
            End If
        Else 'Project has been opened using Command Line arguments.
            Project.ReadParameters()
            Project.ReadParentParameters()
            If Project.ParentParameterExists("ProNetName") Then
                Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
                ProNetName = Project.Parameter("ProNetName").Value
            Else
                ProNetName = Project.GetParameter("ProNetName")
            End If
            If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
                Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
                ProNetPath = Project.Parameter("ProNetPath").Value
            Else
                ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
            End If
            Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

            Project.LockProject() 'Lock the project while it is open in this application.
            ProjectSelected = False 'Reset the Project Selected flag.

            'Set up the Message object:
            Message.SettingsLocn = Project.SettingsLocn
            Message.Show() 'Added 18May19
        End If

        'Initialise the form: ===============================================================
        Project.CreateDataDir() 'This project may need a Data Directory to store .pdf and .xlsx files.

        Me.WebBrowser2.ObjectForScripting = Me

        trvLibrary.ImageList = ImageList1

        rbDefaultWindow.Checked = True ' Default - Display new item in default window.

        XmlHtmDisplay1.WordWrap = False

        XmlHtmDisplay1.Settings.ClearAllTextTypes()
        'Default message display settings:
        XmlHtmDisplay1.Settings.AddNewTextType("Warning")
        XmlHtmDisplay1.Settings.TextType("Warning").FontName = "Arial"
        XmlHtmDisplay1.Settings.TextType("Warning").Bold = True
        XmlHtmDisplay1.Settings.TextType("Warning").Color = Color.Red
        XmlHtmDisplay1.Settings.TextType("Warning").PointSize = 12

        XmlHtmDisplay1.Settings.AddNewTextType("Default")
        XmlHtmDisplay1.Settings.TextType("Default").FontName = "Arial"
        XmlHtmDisplay1.Settings.TextType("Default").Bold = False
        XmlHtmDisplay1.Settings.TextType("Default").Color = Color.Black
        XmlHtmDisplay1.Settings.TextType("Default").PointSize = 10

        XmlHtmDisplay1.Settings.UpdateFontIndexes()
        XmlHtmDisplay1.Settings.UpdateColorIndexes()

        XmlHtmDisplay1.AllowDrop = True

        'Settings tab:
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.Columns.Clear()

        Dim TextBoxCol0 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol0)
        DataGridView1.Columns(0).HeaderText = "Text Type"

        Dim TextBoxCol1 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol1)
        DataGridView1.Columns(1).HeaderText = "Font Name"

        Dim TextBoxCol2 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol2)
        DataGridView1.Columns(2).HeaderText = "Font Index"

        Dim ComboBoxCol3 As New DataGridViewComboBoxColumn
        DataGridView1.Columns.Add(ComboBoxCol3)
        DataGridView1.Columns(3).HeaderText = "Color"

        For Each item As KnownColor In [Enum].GetValues(GetType(KnownColor))
            ComboBoxCol3.Items.Add([Enum].GetName(GetType(KnownColor), item))
        Next

        Dim TextBoxCol4 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol4)
        DataGridView1.Columns(4).HeaderText = "Color Index"

        Dim TextBoxCol5 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol5)
        DataGridView1.Columns(5).HeaderText = "Half Point Size"

        Dim TextBoxCol6 As New DataGridViewTextBoxColumn
        DataGridView1.Columns.Add(TextBoxCol6)
        DataGridView1.Columns(6).HeaderText = "Point Size"

        Dim ComboBoxCol7 As New DataGridViewComboBoxColumn
        DataGridView1.Columns.Add(ComboBoxCol7)
        DataGridView1.Columns(7).HeaderText = "Bold"
        ComboBoxCol7.Items.Add("True")
        ComboBoxCol7.Items.Add("False")

        Dim ComboBoxCol8 As New DataGridViewComboBoxColumn
        DataGridView1.Columns.Add(ComboBoxCol8)
        DataGridView1.Columns(8).HeaderText = "Italic"
        ComboBoxCol8.Items.Add("True")
        ComboBoxCol8.Items.Add("False")

        DataGridView1.AllowUserToAddRows = True

        'NOTE: cmbDocType is no longer used.
        'cmbDocType.Items.Clear()
        'cmbDocType.Items.Add("RTF")
        'cmbDocType.Items.Add("XML")
        'cmbDocType.Items.Add("TXT")
        'cmbDocType.Items.Add("HTML")
        'cmbDocType.Items.Add("PDF")
        'cmbDocType.Items.Add("XLS")
        'cmbDocType.Items.Add("FolderLink")

        rbFileInProject.Checked = True 'Default: open new file from Project.

        ShowSettings()   'Show the XML and other Text Type display settings

        XmlHtmDisplay2.Settings = XmlHtmDisplay1.Settings
        XmlHtmDisplay2.WordWrap = False
        WebBrowser1.Visible = False

        pbIconCollection.Image = ImageList1.Images(2)
        pbIconRtf.Image = ImageList1.Images(4)
        pbIconXml.Image = ImageList1.Images(6)
        pbIconHtml.Image = ImageList1.Images(8)
        pbIconTxt.Image = ImageList1.Images(10)
        pbIconPdf.Image = ImageList1.Images(12)
        pbIconFolder.Image = ImageList1.Images(16)
        pbIconXls.Image = ImageList1.Images(14)
        pbIconXmsg.Image = ImageList1.Images(18)
        pbIconXSeq.Image = ImageList1.Images(20)

        rbCollection.Checked = True 'Select collection in the Add Item groupbox by default.

        rbCodeView.Checked = True 'Default: Open HTML file in code view.

        'Disable Web View and Code View buttons:
        btnCodeView.Enabled = False
        btnWebView.Enabled = False

        'Set up dgvDocList
        dgvDocList.ColumnHeadersDefaultCellStyle.Font = New Font(dgvDocList.Font, FontStyle.Bold) 'Use bold font for the column headers
        dgvDocList.ColumnCount = 5
        dgvDocList.Columns(0).HeaderText = "Document Title"
        dgvDocList.Columns(1).HeaderText = "Type"
        'dgvDocList.Columns(2).HeaderText = "File Name"
        dgvDocList.Columns(2).HeaderText = "Creation Date"
        dgvDocList.Columns(2).ValueType = GetType(Date)
        dgvDocList.Columns(2).DefaultCellStyle.Format = "d-MMM-yyyy H:mm:ss"

        dgvDocList.Columns(3).HeaderText = "Last Edit Date"
        'dgvDocList.Columns(4).ValueType = GetType(DateTime)
        dgvDocList.Columns(3).ValueType = GetType(Date)
        dgvDocList.Columns(3).DefaultCellStyle.Format = "d-MMM-yyyy H:mm:ss"

        'dgvDocList.Columns(5).HeaderText = "Document Description"
        dgvDocList.Columns(4).HeaderText = "File Name"

        InitialiseForm() 'Initialise the form for a new project.

        '------------------------------------------------------------------------------------

        RestoreFormSettings() 'Restore the form settings
        Message.ShowXMessages = ShowXMessages
        Message.ShowSysMessages = ShowSysMessages
        RestoreProjectSettings() 'Restore the Project settings

        ShowProjectInfo()

        Message.AddText("------------------- Started OK -------------------------------------------------------------------------- " & vbCrLf & vbCrLf, "Heading")

        If StartupConnectionName = "" Then
            If Project.ConnectOnOpen Then
                ConnectToComNet() 'The Project is set to connect when it is opened.
            ElseIf ApplicationInfo.ConnectOnStartup Then
                ConnectToComNet() 'The Application is set to connect when it is started.
            Else
                'Don't connect to ComNet.
            End If
        Else
            'Connect to AppNet using the connection name StartupConnectionName.
            ConnectToComNet(StartupConnectionName)
        End If

    End Sub

    Private Sub InitialiseForm()
        'Initialise the form for a new project.
        OpenStartPage()
    End Sub

    Private Sub ShowProjectInfo()
        'Show the project information:

        txtParentProject.Text = Project.ParentProjectName
        txtProNetName.Text = Project.GetParameter("ProNetName")
        txtProjectName.Text = Project.Name
        txtProjectDescription.Text = Project.Description
        Select Case Project.Type
            Case ADVL_Utilities_Library_1.Project.Types.Directory
                txtProjectType.Text = "Directory"
            Case ADVL_Utilities_Library_1.Project.Types.Archive
                txtProjectType.Text = "Archive"
            Case ADVL_Utilities_Library_1.Project.Types.Hybrid
                txtProjectType.Text = "Hybrid"
            Case ADVL_Utilities_Library_1.Project.Types.None
                txtProjectType.Text = "None"
        End Select
        txtCreationDate.Text = Format(Project.Usage.FirstUsed, "d-MMM-yyyy H:mm:ss")
        txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        txtProjectPath.Text = Project.Path

        Select Case Project.SettingsLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSettingsLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSettingsLocationType.Text = "Archive"
        End Select
        txtSettingsPath.Text = Project.SettingsLocn.Path

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtDataLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtDataLocationType.Text = "Archive"
        End Select
        txtDataPath.Text = Project.DataLocn.Path

        Select Case Project.SystemLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                txtSystemLocationType.Text = "Directory"
            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                txtSystemLocationType.Text = "Archive"
        End Select
        txtSystemPath.Text = Project.SystemLocn.Path

        If Project.ConnectOnOpen Then
            chkConnect.Checked = True
        Else
            chkConnect.Checked = False
        End If

        txtTotalDuration.Text = Project.Usage.TotalDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                Project.Usage.TotalDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                Project.Usage.TotalDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                Project.Usage.TotalDuration.Seconds.ToString.PadLeft(2, "0"c)

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        'This code stops an unnecessary error dialog appearing.
        If e.Context = DataGridViewDataErrorContexts.Formatting Or e.Context = DataGridViewDataErrorContexts.PreferredSize Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exit the Application

        'Save any open documents:
        'For Each formItem In RtfDisplayFormList
        '    If formItem Is Nothing Then
        '        'The form is already closed
        '    Else
        '        formItem.Close
        '    End If
        'Next
        'For Each formItem In HtmlDisplayFormList
        '    If formItem Is Nothing Then
        '        'The form is already closed
        '    Else
        '        formItem.Close
        '    End If
        'Next
        'For Each formItem In TextDisplayFormList
        '    If formItem Is Nothing Then
        '        'The form is already closed
        '    Else
        '        formItem.Close
        '    End If
        'Next

        Dim I As Integer
        For I = RtfDisplayFormList.Count - 1 To 0 Step -1
            If RtfDisplayFormList(I) Is Nothing Then
                'The form is already closed
            Else
                RtfDisplayFormList(I).Close
            End If
        Next
        For I = HtmlDisplayFormList.Count - 1 To 0 Step -1
            If HtmlDisplayFormList(I) Is Nothing Then
                'The form is already closed
            Else
                HtmlDisplayFormList(I).Close
            End If
        Next
        For I = TextDisplayFormList.Count - 1 To 0 Step -1
            If TextDisplayFormList(I) Is Nothing Then
                'The form is already closed
            Else
                TextDisplayFormList(I).Close
            End If
        Next

        DisconnectFromComNet() 'Disconnect from the Application Network.

        SaveProjectSettings() 'Save project settings.

        SaveLibrary()

        ApplicationInfo.WriteFile() 'Update the Application Information file.

        Project.SaveLastProjectInfo() 'Save information about the last project used.
        Project.SaveParameters() 'ADDED 3Feb19

        Project.Usage.SaveUsageInfo() 'Save Project usage information.

        Project.UnlockProject() 'Unlock the project.

        ApplicationUsage.SaveUsageInfo() 'Save Application usage information.
        ApplicationInfo.UnlockApplication()

        Debug.Print("Exiting the application. --------------------------------")

        Debug.Print("==================================================================================================")

        Application.Exit()

    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'Save the form settings if the form state is normal. (A minimised form will have the incorrect size and location.)
        If WindowState = FormWindowState.Normal Then
            SaveFormSettings()
        End If
    End Sub


#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

    'Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

    '    Select Case FileTypeSelection
    '        Case FileTypes.RTF
    '            If IsNothing(EditRtf) Then
    '                EditRtf = New frmEditRtf
    '                EditRtf.Show()
    '            Else
    '                EditRtf.Show()
    '            End If

    '        Case FileTypes.TXT

    '        Case FileTypes.XML
    '            If IsNothing(EditXml) Then
    '                EditXml = New frmEditXml
    '                EditXml.Show()
    '            Else
    '                EditXml.Show()
    '            End If

    '    End Select

    'End Sub

    Private Sub EditXml_FormClosed(sender As Object, e As FormClosedEventArgs) Handles EditXml.FormClosed
        EditXml = Nothing
    End Sub

    Private Sub EditRtf_FormClosed(sender As Object, e As FormClosedEventArgs) Handles EditRtf.FormClosed
        EditRtf = Nothing
    End Sub

    'Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click

    '    Select Case FileTypeSelection
    '        Case FileTypes.RTF
    '            If IsNothing(NewRtf) Then
    '                NewRtf = New frmNewRtf
    '                NewRtf.Show()
    '            Else
    '                NewRtf.Show()
    '            End If
    '            If rbFileInProject.Checked = True Then
    '                NewRtf.FileLocation = frmNewRtf.LocationTypes.Project
    '            Else
    '                NewRtf.FileLocation = frmNewRtf.LocationTypes.FileSystem
    '            End If
    '        Case FileTypes.TXT

    '        Case FileTypes.XML
    '            If IsNothing(NewXml) Then
    '                NewXml = New frmNewXml
    '                NewXml.Show()
    '            Else
    '                NewXml.Show()
    '            End If

    '    End Select
    'End Sub

    Private Sub NewRtf_FormClosed(sender As Object, e As FormClosedEventArgs) Handles NewRtf.FormClosed
        NewRtf = Nothing
    End Sub

    Private Sub NewXml_FormClosed(sender As Object, e As FormClosedEventArgs) Handles NewXml.FormClosed
        NewXml = Nothing
    End Sub

    'Private Sub btnNewDisplay_Click(sender As Object, e As EventArgs) Handles btnNewDisplay.Click
    '    'Open a new document display window

    '    Select Case FileTypeSelection
    '        Case FileTypes.RTF
    '            'Code to display multiple instances of the RTF Display form:
    '            RtfDisplay = New frmRtfDisplay
    '            If RtfDisplayFormList.Count = 0 Then
    '                RtfDisplayFormList.Add(RtfDisplay)
    '                RtfDisplayFormList(0).FormNo = 0
    '                RtfDisplayFormList(0).Show
    '            Else
    '                Dim I As Integer
    '                Dim FormAdded As Boolean = False
    '                For I = 0 To RtfDisplayFormList.Count - 1 'Check if there are closed forms in RtfDisplayFormList. They can be re-used.
    '                    If IsNothing(RtfDisplayFormList(I)) Then
    '                        RtfDisplayFormList(I) = RtfDisplay
    '                        RtfDisplayFormList(I).FormNo = I
    '                        RtfDisplayFormList(I).Show
    '                        FormAdded = True
    '                        Exit For
    '                    End If
    '                Next
    '                If FormAdded = False Then 'Add a new form to XmlDisplayFormList
    '                    Dim FormNo As Integer
    '                    RtfDisplayFormList.Add(RtfDisplay)
    '                    FormNo = RtfDisplayFormList.Count - 1
    '                    RtfDisplayFormList(FormNo).FormNo = FormNo
    '                    RtfDisplayFormList(FormNo).Show
    '                End If
    '            End If

    '        Case FileTypes.TXT

    '        Case FileTypes.XML
    '            'Code to display multiple instances of the XML Display form:
    '            XmlDisplay = New frmXmlDisplay
    '            If XmlDisplayFormList.Count = 0 Then
    '                XmlDisplayFormList.Add(XmlDisplay)
    '                XmlDisplayFormList(0).FormNo = 0
    '                XmlDisplayFormList(0).Show
    '            Else
    '                Dim I As Integer
    '                Dim FormAdded As Boolean = False
    '                For I = 0 To XmlDisplayFormList.Count - 1 'Check if there are closed forms in XmlDisplayFormList. They can be re-used.
    '                    If IsNothing(XmlDisplayFormList(I)) Then
    '                        XmlDisplayFormList(I) = XmlDisplay
    '                        XmlDisplayFormList(I).FormNo = I
    '                        XmlDisplayFormList(I).Show
    '                        FormAdded = True
    '                        Exit For
    '                    End If
    '                Next
    '                If FormAdded = False Then 'Add a new form to XmlDisplayFormList
    '                    Dim FormNo As Integer
    '                    XmlDisplayFormList.Add(XmlDisplay)
    '                    FormNo = XmlDisplayFormList.Count - 1
    '                    XmlDisplayFormList(FormNo).FormNo = FormNo
    '                    XmlDisplayFormList(FormNo).Show
    '                End If
    '            End If

    '        Case FileTypes.HTML
    '            'Code to display multiple instances of the HTML Display form:
    '            HtmlDisplay = New frmHtmlDisplay
    '            If HtmlDisplayFormList.Count = 0 Then
    '                HtmlDisplayFormList.Add(HtmlDisplay)
    '                HtmlDisplayFormList(0).FormNo = 0
    '                HtmlDisplayFormList(0).Show
    '            Else
    '                Dim I As Integer
    '                Dim FormAdded As Boolean = False
    '                For I = 0 To HtmlDisplayFormList.Count - 1 'Check if there are closed forms in XmlDisplayFormList. They can be re-used.
    '                    If IsNothing(HtmlDisplayFormList(I)) Then
    '                        HtmlDisplayFormList(I) = HtmlDisplay
    '                        HtmlDisplayFormList(I).FormNo = I
    '                        HtmlDisplayFormList(I).Show
    '                        FormAdded = True
    '                        Exit For
    '                    End If
    '                Next
    '                If FormAdded = False Then 'Add a new form to XmlDisplayFormList
    '                    Dim FormNo As Integer
    '                    HtmlDisplayFormList.Add(HtmlDisplay)
    '                    FormNo = HtmlDisplayFormList.Count - 1
    '                    HtmlDisplayFormList(FormNo).FormNo = FormNo
    '                    HtmlDisplayFormList(FormNo).Show
    '                End If
    '            End If
    '    End Select
    'End Sub

    Public Function NewRtfDisplay() As Integer
        'Open a new RTF Display window, or reuse an existing one if avaiable.
        'The new forms index number in RtfDisplayFormList is returned.

        RtfDisplay = New frmRtfDisplay
        If RtfDisplayFormList.Count = 0 Then
            RtfDisplayFormList.Add(RtfDisplay)
            RtfDisplayFormList(0).FormNo = 0
            RtfDisplayFormList(0).Show
            Return 0 'The new RTF Display is at position 0 in RtfDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To RtfDisplayFormList.Count - 1 'Check if there are closed forms in RtfDisplayFormList. They can be re-used.
                If IsNothing(RtfDisplayFormList(I)) Then
                    RtfDisplayFormList(I) = RtfDisplay
                    RtfDisplayFormList(I).FormNo = I
                    RtfDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new RTF Display is at position I in RtfDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to RtfDisplayFormList
                Dim FormNo As Integer
                RtfDisplayFormList.Add(RtfDisplay)
                FormNo = RtfDisplayFormList.Count - 1
                RtfDisplayFormList(FormNo).FormNo = FormNo
                RtfDisplayFormList(FormNo).Show
                Return FormNo 'The new RTF Display is at position FormNo in RtfDisplayFormList()
            End If
        End If

    End Function

    Public Function NewXmlDisplay() As Integer
        'Open a new XML Display window, or reuse an existing on if avaiable.
        'The new forms index number in XmlDisplayFormList is returned.

        XmlDisplay = New frmXmlDisplay
        If XmlDisplayFormList.Count = 0 Then
            XmlDisplayFormList.Add(XmlDisplay)
            XmlDisplayFormList(0).FormNo = 0
            XmlDisplayFormList(0).Show
            Return 0 'The new XML Display is at position 0 in XmlDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To XmlDisplayFormList.Count - 1 'Check if there are closed forms in XmlDisplayFormList. They can be re-used.
                If IsNothing(XmlDisplayFormList(I)) Then
                    XmlDisplayFormList(I) = XmlDisplay
                    XmlDisplayFormList(I).FormNo = I
                    XmlDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new Xml Display is at position I in XmlDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to XmlDisplayFormList
                Dim FormNo As Integer
                XmlDisplayFormList.Add(XmlDisplay)
                FormNo = XmlDisplayFormList.Count - 1
                XmlDisplayFormList(FormNo).FormNo = FormNo
                XmlDisplayFormList(FormNo).Show
                Return FormNo 'The new XML Display is at position FormNo in XmlDisplayFormList()
            End If

        End If
    End Function

    Public Function NewHtmlDisplay() As Integer
        'Open a new HTML Display window, or reuse an existing on if avaiable.
        'The new forms index number in HtmlDisplayFormList is returned.

        HtmlDisplay = New frmHtmlDisplay
        If HtmlDisplayFormList.Count = 0 Then
            HtmlDisplayFormList.Add(HtmlDisplay)
            HtmlDisplayFormList(0).FormNo = 0
            HtmlDisplayFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in HtmlDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To HtmlDisplayFormList.Count - 1 'Check if there are closed forms in HtmlDisplayFormList. They can be re-used.
                If IsNothing(HtmlDisplayFormList(I)) Then
                    HtmlDisplayFormList(I) = HtmlDisplay
                    HtmlDisplayFormList(I).FormNo = I
                    HtmlDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in HtmlDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to HtmlDisplayFormList
                Dim FormNo As Integer
                HtmlDisplayFormList.Add(HtmlDisplay)
                FormNo = HtmlDisplayFormList.Count - 1
                HtmlDisplayFormList(FormNo).FormNo = FormNo
                HtmlDisplayFormList(FormNo).Show
                Return FormNo 'The new HTML Display is at position FormNo in HtmlDisplayFormList()
            End If

        End If
    End Function

    Public Function NewWebView() As Integer
        'Open a new HTML Web View window, or reuse an existing on if avaiable.
        'The new forms index number in WebViewFormList is returned.

        WebView = New frmWebView
        If WebViewFormList.Count = 0 Then
            WebViewFormList.Add(WebView)
            WebViewFormList(0).FormNo = 0
            WebViewFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in WebViewFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To WebViewFormList.Count - 1 'Check if there are closed forms in WebViewFormList. They can be re-used.
                If IsNothing(WebViewFormList(I)) Then
                    WebViewFormList(I) = WebView
                    WebViewFormList(I).FormNo = I
                    WebViewFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in WebViewFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to WebViewFormList
                Dim FormNo As Integer
                WebViewFormList.Add(WebView)
                FormNo = WebViewFormList.Count - 1
                WebViewFormList(FormNo).FormNo = FormNo
                WebViewFormList(FormNo).Show
                Return FormNo 'The new HTML Display is at position FormNo in WebViewFormList()
            End If

        End If
    End Function

    Public Function NewTextDisplay() As Integer
        'Open a new Text Display window, or reuse an existing on if avaiable.
        'The new forms index number in TextDisplayFormList is returned.

        TextDisplay = New frmTextDisplay
        If TextDisplayFormList.Count = 0 Then
            TextDisplayFormList.Add(TextDisplay)
            TextDisplayFormList(0).FormNo = 0
            TextDisplayFormList(0).Show
            Return 0 'The new Text Display is at position 0 in TextDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To TextDisplayFormList.Count - 1 'Check if there are closed forms in TextDisplayFormList. They can be re-used.
                If IsNothing(TextDisplayFormList(I)) Then
                    TextDisplayFormList(I) = TextDisplay
                    TextDisplayFormList(I).FormNo = I
                    TextDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new Text Display is at position I in TextDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to TextDisplayFormList
                Dim FormNo As Integer
                TextDisplayFormList.Add(TextDisplay)
                FormNo = TextDisplayFormList.Count - 1
                TextDisplayFormList(FormNo).FormNo = FormNo
                TextDisplayFormList(FormNo).Show
                Return FormNo 'The new Text Display is at position FormNo in TextDisplayFormList()
            End If

        End If
    End Function

    Public Function NewPdfDisplay() As Integer
        'Open a new PDF Display window, or reuse an existing one if available.
        'The new forms index number in PdfDisplayFormList is returned.

        PdfDisplay = New frmPdfDisplay
        If PdfDisplayFormList.Count = 0 Then
            PdfDisplayFormList.Add(PdfDisplay)
            PdfDisplayFormList(0).FormNo = 0
            PdfDisplayFormList(0).Show
            Return 0 'The new RPDF Display is at position 0 in PdfDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To PdfDisplayFormList.Count - 1 'Check if there are closed forms in PdfDisplayFormList. They can be re-used.
                If IsNothing(PdfDisplayFormList(I)) Then
                    PdfDisplayFormList(I) = PdfDisplay
                    PdfDisplayFormList(I).FormNo = I
                    PdfDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new PDF Display is at position I in PdfDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to PdfDisplayFormList
                Dim FormNo As Integer
                PdfDisplayFormList.Add(PdfDisplay)
                FormNo = PdfDisplayFormList.Count - 1
                PdfDisplayFormList(FormNo).FormNo = FormNo
                PdfDisplayFormList(FormNo).Show
                Return FormNo 'The new PDF Display is at position FormNo in PdfDisplayFormList()
            End If
        End If
    End Function

    Public Sub RtfDisplayFormClosed()
        'This subroutine is called when the RTF Display form has been closed.
        'The subroutine is usually called from the FormClosed event of the RtfDisplay form.
        'The RtfDisplay form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the RtfDisplay form.
        'This property should be updated by the RtfDisplay form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in RtfDisplayList should be set to Nothing.

        If RtfDisplayFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in RtfDisplayFormList
            Exit Sub
        End If

        If IsNothing(RtfDisplayFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            RtfDisplayFormList(ClosedFormNo) = Nothing
        End If

    End Sub

    Public Sub XmlDisplayFormClosed()
        'This subroutine is called when the XML Display form has been closed.
        'The subroutine is usually called from the FormClosed event of the XmlDisplay form.
        'The XmlDisplay form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the XmlDisplay form.
        'This property should be updated by the XmlDisplay form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in XmlDisplayList should be set to Nothing.

        If XmlDisplayFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in XmlDisplayFormList
            Exit Sub
        End If

        If IsNothing(XmlDisplayFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            XmlDisplayFormList(ClosedFormNo) = Nothing
        End If

    End Sub

    Public Sub TextDisplayFormClosed()
        'This subroutine is called when the TXT Display form has been closed.
        'The subroutine is usually called from the FormClosed event of the TxtDisplay form.
        'The TxtDisplay form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the TxtDisplay form.
        'This property should be updated by the TxtDisplay form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in TxtDisplayList should be set to Nothing.

        If TextDisplayFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in TxtDisplayFormList
            Exit Sub
        End If

        If IsNothing(TextDisplayFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            TextDisplayFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Public Sub HtmlDisplayFormClosed()
        'This subroutine is called when the HTML Display form has been closed.
        'The subroutine is usually called from the FormClosed event of the HtmlDisplay form.
        'The HtmlDisplay form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the HtmlDisplay form.
        'This property should be updated by the HtmlDisplay form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in HtmlDisplayList should be set to Nothing.

        If HtmlDisplayFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in HtmlDisplayFormList
            Exit Sub
        End If

        If IsNothing(HtmlDisplayFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            HtmlDisplayFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Public Sub WebViewFormClosed()
        'This subroutine is called when the Web View form has been closed.
        'The subroutine is usually called from the FormClosed event of the WebView form.
        'The WebView form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the WebView form.
        'This property should be updated by the WebView form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in WebViewList should be set to Nothing.

        If WebViewFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in WebViewFormList
            Exit Sub
        End If

        If IsNothing(WebViewFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            WebViewFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Private Sub btnNewCollection_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub NewCollection_FormClosed(sender As Object, e As FormClosedEventArgs) Handles NewCollection.FormClosed
        NewCollection = Nothing
    End Sub

    Private Sub btnNewLibrary_Click(sender As Object, e As EventArgs) Handles btnNewLibrary.Click
        'Open the New Library form:
        If IsNothing(NewLibrary) Then
            NewLibrary = New frmNewLibrary
            NewLibrary.Show()
        Else
            NewLibrary.Show()
        End If
    End Sub

    Private Sub NewLibrary_FormClosed(sender As Object, e As FormClosedEventArgs) Handles NewLibrary.FormClosed
        NewLibrary = Nothing
    End Sub

    Private Sub btnMessages_Click(sender As Object, e As EventArgs) Handles btnMessages.Click
        'Show the Messages form.
        Message.ApplicationName = ApplicationInfo.Name
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show()
        Message.ShowXMessages = ShowXMessages
        Message.MessageForm.BringToFront()
    End Sub

    Private Sub btnWebPages_Click(sender As Object, e As EventArgs) Handles btnWebPages.Click
        'Open the Web Pages form.

        If IsNothing(WebPageList) Then
            WebPageList = New frmWebPageList
            WebPageList.Show()
        Else
            WebPageList.Show()
            WebPageList.BringToFront()
        End If
    End Sub

    Private Sub WebPageList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles WebPageList.FormClosed
        WebPageList = Nothing
    End Sub

    Public Function OpenNewWebPage() As Integer
        'Open a new HTML Web View window, or reuse an existing one if avaiable.
        'The new forms index number in WebViewFormList is returned.

        NewWebPage = New frmWebPage
        If WebPageFormList.Count = 0 Then
            WebPageFormList.Add(NewWebPage)
            WebPageFormList(0).FormNo = 0
            WebPageFormList(0).Show
            Return 0 'The new HTML Display is at position 0 in WebViewFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To WebPageFormList.Count - 1 'Check if there are closed forms in WebViewFormList. They can be re-used.
                If IsNothing(WebPageFormList(I)) Then
                    WebPageFormList(I) = NewWebPage
                    WebPageFormList(I).FormNo = I
                    WebPageFormList(I).Show
                    FormAdded = True
                    Return I 'The new Html Display is at position I in WebViewFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to WebViewFormList
                Dim FormNo As Integer
                WebPageFormList.Add(NewWebPage)
                FormNo = WebPageFormList.Count - 1
                WebPageFormList(FormNo).FormNo = FormNo
                WebPageFormList(FormNo).Show
                Return FormNo 'The new WebPage is at position FormNo in WebPageFormList()
            End If
        End If
    End Function

    Public Sub WebPageFormClosed()
        'This subroutine is called when the Web Page form has been closed.
        'The subroutine is usually called from the FormClosed event of the WebPage form.
        'The WebPage form may have multiple instances.
        'The ClosedFormNumber property should contains the number of the instance of the WebPage form.
        'This property should be updated by the WebPage form when it is being closed.
        'The ClosedFormNumber property value is used to determine which element in WebPageList should be set to Nothing.

        If WebPageFormList.Count < ClosedFormNo + 1 Then
            'ClosedFormNo is too large to exist in WebPageFormList
            Exit Sub
        End If

        If IsNothing(WebPageFormList(ClosedFormNo)) Then
            'The form is already set to nothing
        Else
            WebPageFormList(ClosedFormNo) = Nothing
        End If
    End Sub

    Public Function OpenNewWFHtmlDisplayPage() As Integer
        'Open a new WorkFlow HTML display window, or reuse an existing one if avaiable.
        'The new forms index number in WFHtmlDisplayFormList is returned.

        NewWFHtmlDisplay = New frmWFHtmlDisplay
        If WFHtmlDisplayFormList.Count = 0 Then
            WFHtmlDisplayFormList.Add(NewWFHtmlDisplay)
            WFHtmlDisplayFormList(0).FormNo = 0
            WFHtmlDisplayFormList(0).Show
            Return 0 'The new WorkFlow HTML Display is at position 0 in WFHtmlDisplayFormList()
        Else
            Dim I As Integer
            Dim FormAdded As Boolean = False
            For I = 0 To WFHtmlDisplayFormList.Count - 1 'Check if there are closed forms in WFHtmlDisplayFormList. They can be re-used.
                If IsNothing(WFHtmlDisplayFormList(I)) Then
                    WFHtmlDisplayFormList(I) = NewWFHtmlDisplay()
                    WFHtmlDisplayFormList(I).FormNo = I
                    WFHtmlDisplayFormList(I).Show
                    FormAdded = True
                    Return I 'The new WorkFlow Html Display is at position I in WFHtmlDisplayFormList()
                    Exit For
                End If
            Next
            If FormAdded = False Then 'Add a new form to WFHtmlDisplayFormList
                Dim FormNo As Integer
                WFHtmlDisplayFormList.Add(NewWFHtmlDisplay)
                FormNo = WFHtmlDisplayFormList.Count - 1
                WFHtmlDisplayFormList(FormNo).FormNo = FormNo
                WFHtmlDisplayFormList(FormNo).Show
                Return FormNo 'The new WorkFlow HtmlDisplay is at position FormNo in WFHtmlDisplayFormList()
            End If
        End If
    End Function


#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

    Private Sub btnProject_Click(sender As Object, e As EventArgs) Handles btnProject.Click
        Project.SelectProject()
    End Sub

    Private Sub btnParameters_Click(sender As Object, e As EventArgs) Handles btnParameters.Click
        Project.ShowParameters()
    End Sub

    Private Sub btnOpenProject_Click(sender As Object, e As EventArgs) Handles btnOpenProject.Click
        If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then

        Else
            Process.Start(Project.Path)
        End If
    End Sub

    Private Sub btnOpenSettings_Click(sender As Object, e As EventArgs) Handles btnOpenSettings.Click
        If Project.SettingsLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SettingsLocn.Path)
        End If
    End Sub

    Private Sub btnOpenData_Click(sender As Object, e As EventArgs) Handles btnOpenData.Click
        If Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.DataLocn.Path)
        End If
    End Sub

    Private Sub btnOpenSystem_Click(sender As Object, e As EventArgs) Handles btnOpenSystem.Click
        If Project.SystemLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory Then
            Process.Start(Project.SystemLocn.Path)
        End If
    End Sub

    Private Sub btnOpenAppDir_Click(sender As Object, e As EventArgs) Handles btnOpenAppDir.Click

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Add the current project to the Message Service list.

        If Project.ParentProjectName <> "" Then
            Message.AddWarning("This project has a parent: " & Project.ParentProjectName & vbCrLf)
            Message.AddWarning("Child projects can not be added to the list." & vbCrLf)
            Exit Sub
        End If

        If ConnectedToComnet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", Project.Path)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to AppNet:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

    Private Sub btnAppInfo_Click(sender As Object, e As EventArgs) Handles btnAppInfo.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Private Sub btnAndorville_Click(sender As Object, e As EventArgs) Handles btnAndorville.Click
        ApplicationInfo.ShowInfo()
    End Sub

    Public Sub UpdateWebPage(ByVal FileName As String)
        'Update the web page in WebPageFormList if the Web file name is FileName.

        Dim NPages As Integer = WebPageFormList.Count
        Dim I As Integer

        Try
            For I = 0 To NPages - 1
                If IsNothing(WebPageFormList(I)) Then
                    'Web page has been deleted!
                Else
                    If WebPageFormList(I).FileName = FileName Then
                        WebPageFormList(I).OpenDocument
                    End If
                End If
            Next
        Catch ex As Exception
            Message.AddWarning(ex.Message & vbCrLf)
        End Try
    End Sub


#Region " Start Page Code" '=========================================================================================================================================

    Public Sub OpenStartPage()
        'Open the StartPage.html file and display in the Start Page tab.

        If Project.DataFileExists("StartPage.html") Then
            'StartPageFileName = "StartPage.html"
            WorkflowFileName = "StartPage.html"
            'DisplayStartPage()
            DisplayWorkflow()
        Else
            CreateStartPage()
            'StartPageFileName = "StartPage.html"
            WorkflowFileName = "StartPage.html"
            'DisplayStartPage()
            DisplayWorkflow()
        End If

    End Sub

    'Public Sub DisplayStartPage()
    '    'Display the StartPage.html file in the Start Page tab.

    '    If Project.DataFileExists(StartPageFileName) Then
    '        Dim rtbData As New IO.MemoryStream
    '        Project.ReadData(StartPageFileName, rtbData)
    '        rtbData.Position = 0
    '        Dim sr As New IO.StreamReader(rtbData)
    '        WebBrowser2.DocumentText = sr.ReadToEnd()
    '    Else
    '        Message.AddWarning("Web page file not found: " & StartPageFileName & vbCrLf)
    '    End If
    'End Sub

    Public Sub DisplayWorkflow()
        'Display the StartPage.html file in the Start Page tab.

        If Project.DataFileExists(WorkflowFileName) Then
            Dim rtbData As New IO.MemoryStream
            Project.ReadData(WorkflowFileName, rtbData)
            rtbData.Position = 0
            Dim sr As New IO.StreamReader(rtbData)
            'WebBrowser1.DocumentText = sr.ReadToEnd()
            WebBrowser2.DocumentText = sr.ReadToEnd()
        Else
            Message.AddWarning("Web page file not found: " & WorkflowFileName & vbCrLf)
        End If
    End Sub

    Private Sub CreateStartPage()
        'Create a new default StartPage.html file.

        Dim htmData As New IO.MemoryStream
        Dim sw As New IO.StreamWriter(htmData)
        sw.Write(AppInfoHtmlString("Application Information")) 'Create a web page providing information about the application.
        sw.Flush()
        Project.SaveData("StartPage.html", htmData)
    End Sub

    Public Function AppInfoHtmlString(ByVal DocumentTitle As String) As String
        'Create an Application Information Web Page.

        'This function should be edited to provide a brief description of the Application.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Application information."">" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & "Andorville&trade; Document Library" & "</h2>" & vbCrLf & vbCrLf) 'Add the page title.
        sb.Append("<hr>") 'Add a horizontal divider line.
        sb.Append("<p>The Document Library stores RTF, XML, HTML, Text and PDF documents in a tree structure.</p>" & vbCrLf) 'Add an application description.
        sb.Append("<p>RTF, XML, HTML and Text documents can be created and edited in this application.</p>" & vbCrLf) 'Add an application description.
        sb.Append("<hr>" & vbCrLf & vbCrLf) 'Add a horizontal divider line.

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

    Public Function DefaultJavaScriptString() As String
        'Generate the default JavaScript section of an Andorville(TM) Workflow Web Page.

        Dim sb As New System.Text.StringBuilder

        'Add JavaScript section:
        sb.Append("<script>" & vbCrLf & vbCrLf)

        'START: User defined JavaScript functions ==========================================================================
        'Add functions to implement the main actions performed by this web page.
        sb.Append("//START: User defined JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Add functions to implement the main actions performed by this web page." & vbCrLf & vbCrLf)

        sb.Append("//END:   User defined JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User defined JavaScript functions --------------------------------------------------------------------------


        'START: User modified JavaScript functions ==========================================================================
        'Modify these function to save all required web page settings and process all expected XMessage instructions.
        sb.Append("//START: User modified JavaScript functions ==========================================================================" & vbCrLf)
        sb.Append("//  Modify these function to save all required web page settings and process all expected XMessage instructions." & vbCrLf & vbCrLf)

        'Add the Start Up code section.
        sb.Append("//Code to execute on Start Up:" & vbCrLf)
        sb.Append("function StartUpCode() {" & vbCrLf)
        sb.Append("  RestoreSettings() ;" & vbCrLf)
        'sb.Append("  GetCalcsDbPath() ;" & vbCrLf)
        sb.Append("}" & vbCrLf & vbCrLf)

        'Add the SaveSettings function - This is used to save web page settings between sessions.
        sb.Append("//Save the web page settings." & vbCrLf)
        sb.Append("function SaveSettings() {" & vbCrLf)
        sb.Append("  var xSettings = ""<Settings>"" + "" \n"" ; //String containing the web page settings in XML format." & vbCrLf)
        sb.Append("  //Add xml lines to save each setting." & vbCrLf & vbCrLf)
        sb.Append("  xSettings +=    ""</Settings>"" + ""\n"" ; //End of the Settings element." & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("  //Save the settings as an XML file in the project." & vbCrLf)
        sb.Append("  window.external.SaveHtmlSettings(xSettings) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Process a single XMsg instruction (Information:Location pair)
        sb.Append("//Process an XMessage instruction:" & vbCrLf)
        sb.Append("function XMsgInstruction(Info, Locn) {" & vbCrLf)
        sb.Append("  switch(Locn) {" & vbCrLf)
        sb.Append("  //Insert case statements here." & vbCrLf)
        sb.Append("  case ""Status"" :" & vbCrLf)
        sb.Append("    if (Info = ""OK"") { " & vbCrLf)
        sb.Append("      //Instruction processing completed OK:" & vbCrLf)
        sb.Append("      } else {" & vbCrLf)
        sb.Append("      window.external.AddWarning(""Error: Unknown Status information: "" + "" Info: "" + Info + ""\r\n"") ;" & vbCrLf)
        sb.Append("     }" & vbCrLf)
        sb.Append("    break ;" & vbCrLf)
        sb.Append(vbCrLf)
        sb.Append("  default:" & vbCrLf)
        sb.Append("    window.external.AddWarning(""Unknown location: "" + Locn + ""\r\n"") ;" & vbCrLf)
        sb.Append("  }" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   User modified JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf)
        'END:   User modified JavaScript functions --------------------------------------------------------------------------

        'START: Required Document Library Web Page JavaScript functions ==========================================================================
        sb.Append("//START: Required Document Library Web Page JavaScript functions ==========================================================================" & vbCrLf & vbCrLf)

        'Add the AddText function - This sends a message to the message window using a named text type.
        sb.Append("//Add text to the Message window using a named txt type:" & vbCrLf)
        sb.Append("function AddText(Msg, TextType) {" & vbCrLf)
        sb.Append("  window.external.AddText(Msg, TextType) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddMessage function - This sends a message to the message window using default black text.
        sb.Append("//Add a message to the Message window using the default black text:" & vbCrLf)
        sb.Append("function AddMessage(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddMessage(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the AddWarning function - This sends a red, bold warning message to the message window.
        sb.Append("//Add a warning message to the Message window using bold red text:" & vbCrLf)
        sb.Append("function AddWarning(Msg) {" & vbCrLf)
        sb.Append("  window.external.AddWarning(Msg) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreSettings function - This is used to restore web page settings.
        sb.Append("//Restore the web page settings." & vbCrLf)
        sb.Append("function RestoreSettings() {" & vbCrLf)
        sb.Append("  window.external.RestoreHtmlSettings() " & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'This line runs the RestoreSettings function when the web page is loaded.
        sb.Append("//Restore the web page settings when the page loads." & vbCrLf)
        'sb.Append("window.onload = RestoreSettings; " & vbCrLf)
        sb.Append("window.onload = StartUpCode ; " & vbCrLf)
        sb.Append(vbCrLf)

        'Restores a single setting on the web page.
        sb.Append("//Restore a web page setting." & vbCrLf)
        sb.Append("  function RestoreSetting(FormName, ItemName, ItemValue) {" & vbCrLf)
        sb.Append("  document.forms[FormName][ItemName].value = ItemValue ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        'Add the RestoreOption function - This is used to add an option to a Select list.
        sb.Append("//Restore a Select control Option." & vbCrLf)
        sb.Append("function RestoreOption(SelectId, OptionText) {" & vbCrLf)
        sb.Append("  var x = document.getElementById(SelectId) ;" & vbCrLf)
        sb.Append("  var option = document.createElement(""Option"") ;" & vbCrLf)
        sb.Append("  option.text = OptionText ;" & vbCrLf)
        sb.Append("  x.add(option) ;" & vbCrLf)
        sb.Append("}" & vbCrLf)
        sb.Append(vbCrLf)

        sb.Append("//END:   Required Document Library Web Page JavaScript functions __________________________________________________________________________" & vbCrLf & vbCrLf)
        'END:   Required Document Library Web Page JavaScript functions --------------------------------------------------------------------------

        sb.Append("</script>" & vbCrLf & vbCrLf)

        Return sb.ToString

    End Function


    Public Function DefaultWFHtmlString(ByVal DocumentTitle As String) As String
        'Create a blank HTML Web Page.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h1>" & DocumentTitle & "</h1>" & vbCrLf & vbCrLf)

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

#End Region 'Start Page Code ------------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Project Events Code"

    Private Sub Project_Message(Msg As String) Handles Project.Message
        'Display the Project message:
        Message.Add(Msg & vbCrLf)
    End Sub

    Private Sub Project_ErrorMessage(Msg As String) Handles Project.ErrorMessage
        'Display the Project error message:
        Message.AddWarning(Msg & vbCrLf)
    End Sub

    Private Sub Project_Closing() Handles Project.Closing
        'The current project is closing.
        SaveFormSettings() 'Save the form settings - they are saved in the Project before is closes.
        SaveProjectSettings() 'Update this subroutine if project settings need to be saved.
        Project.Usage.SaveUsageInfo()   'Save the current project usage information.
        ClearCurrentProjectData()  'Clear the current project data before opening another project.    
        Project.UnlockProject() 'Unlock the current project before it Is closed.
        If ConnectedToComnet Then DisconnectFromComNet()
    End Sub

    Private Sub ClearCurrentProjectData()
        'Clear the current project data before opening another project.

        'Close any open forms:
        If IsNothing(EditXml) Then
        Else
            EditXml.Close()
        End If

        If IsNothing(EditRtf) Then
        Else
            EditRtf.Close()
        End If
        If IsNothing(NewXml) Then
        Else
            NewXml.Close()
        End If
        If IsNothing(NewRtf) Then
        Else
            NewRtf.Close()
        End If
        If IsNothing(NewLibrary) Then
        Else
            NewLibrary.Close()
        End If
        Dim I As Integer
        'Close all XML Display forms:
        For I = 0 To XmlDisplayFormList.Count - 1
            If IsNothing(XmlDisplayFormList(I)) Then
            Else
                XmlDisplayFormList(I).Close
            End If
        Next
        'Close all RTF Display forms:
        For I = 0 To RtfDisplayFormList.Count - 1
            If IsNothing(RtfDisplayFormList(I)) Then
            Else
                RtfDisplayFormList(I).Close
            End If
        Next
        'Close all HTML Display forms:
        For I = 0 To HtmlDisplayFormList.Count - 1
            If IsNothing(HtmlDisplayFormList(I)) Then
            Else
                HtmlDisplayFormList(I).Close
            End If
        Next
        'Close all Web View forms:
        For I = 0 To WebViewFormList.Count - 1
            If IsNothing(WebViewFormList(I)) Then
            Else
                WebViewFormList(I).Close
            End If
        Next
        'Close all Text Display forms:
        For I = 0 To TextDisplayFormList.Count - 1
            If IsNothing(TextDisplayFormList(I)) Then
            Else
                TextDisplayFormList(I).Close
            End If
        Next

        'Clear all data display controls:
        txtLibraryName.Text = ""
        trvLibrary.Nodes.Clear()
        txtNodeKey.Text = ""
        txtItemDescription.Text = ""
        txtNodePath.Text = ""
        txtNodeText.Text = ""
        txtNodeType.Text = ""
        txtSelNodeKey.Text = ""
        txtNodeIndex.Text = ""

        txtNewNodeFileName.Text = ""
        txtItemText.Text = ""
        txtNewNodeDescr.Text = ""

        txtNodeKey2.Text = ""
        txtEditNodeText.Text = ""
        txtEditDescription.Text = ""

        txtFileName.Text = ""
        txtFileType2.Text = ""
        txtDocLocation.Text = ""
        txtFileDirectory.Text = ""
        txtFileDescription.Text = ""

        WebBrowser1.DocumentText = ""
        XmlHtmDisplay2.Clear()

        txtFileName2.Text = ""
        XmlHtmDisplay1.Clear()

        txtSettingsFileName.Text = ""
        txtXmlIndentSpaces.Text = ""
        txtXmlFileSizeLimit.Text = ""
        DataGridView1.Rows.Clear()

        'Clear Properties:
        LibraryFileName = ""
        LibraryName = ""
        LibraryDescription = ""

        'Clear the ItemInfo() dictionary:
        ItemInfo.Clear()

    End Sub

    Private Sub Project_Selected() Handles Project.Selected
        'A new project has been selected.

        RestoreFormSettings()
        Project.ReadProjectInfoFile()

        Project.ReadParameters()
        Project.ReadParentParameters()
        If Project.ParentParameterExists("ProNetName") Then
            Project.AddParameter("ProNetName", Project.ParentParameter("ProNetName").Value, Project.ParentParameter("ProNetName").Description) 'AddParameter will update the parameter if it already exists.
            ProNetName = Project.Parameter("ProNetName").Value
        Else
            ProNetName = Project.GetParameter("ProNetName")
        End If
        If Project.ParentParameterExists("ProNetPath") Then 'Get the parent parameter value - it may have been updated.
            Project.AddParameter("ProNetPath", Project.ParentParameter("ProNetPath").Value, Project.ParentParameter("ProNetPath").Description) 'AddParameter will update the parameter if it already exists.
            ProNetPath = Project.Parameter("ProNetPath").Value
        Else
            ProNetPath = Project.GetParameter("ProNetPath") 'If the parameter does not exist, the value is set to ""
        End If
        Project.SaveParameters() 'These should be saved now - child projects look for parent parameters in the parameter file.

        Project.LockProject() 'Lock the project while it is open in this application.

        Project.Usage.StartTime = Now

        ApplicationInfo.SettingsLocn = Project.SettingsLocn
        Message.SettingsLocn = Project.SettingsLocn
        Message.Show()

        'Restore the new project settings:
        RestoreProjectSettings() 'Update this subroutine if project settings need to be restored.

        ShowProjectInfo()

        ''Show the project information:
        'txtProjectName.Text = Project.Name
        'txtProjectDescription.Text = Project.Description
        'Select Case Project.Type
        '    Case ADVL_Utilities_Library_1.Project.Types.Directory
        '        txtProjectType.Text = "Directory"
        '    Case ADVL_Utilities_Library_1.Project.Types.Archive
        '        txtProjectType.Text = "Archive"
        '    Case ADVL_Utilities_Library_1.Project.Types.Hybrid
        '        txtProjectType.Text = "Hybrid"
        '    Case ADVL_Utilities_Library_1.Project.Types.None
        '        txtProjectType.Text = "None"
        'End Select

        'txtCreationDate.Text = Format(Project.CreationDate, "d-MMM-yyyy H:mm:ss")
        'txtLastUsed.Text = Format(Project.Usage.LastUsed, "d-MMM-yyyy H:mm:ss")
        'Select Case Project.SettingsLocn.Type
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '        txtSettingsLocationType.Text = "Directory"
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
        '        txtSettingsLocationType.Text = "Archive"
        'End Select
        'txtSettingsPath.Text = Project.SettingsLocn.Path
        'Select Case Project.DataLocn.Type
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
        '        txtDataLocationType.Text = "Directory"
        '    Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
        '        txtDataLocationType.Text = "Archive"
        'End Select
        'txtDataPath.Text = Project.DataLocn.Path

        If Project.ConnectOnOpen Then
            ConnectToComNet() 'The Project is set to connect when it is opened.
        ElseIf ApplicationInfo.ConnectOnStartup Then
            ConnectToComNet() 'The Application is set to connect when it is started.
        Else
            'Don't connect to ComNet.
        End If

    End Sub

#End Region 'Project Events Code

#Region " Online/Offline Code"

    Private Sub btnOnline_Click(sender As Object, e As EventArgs) Handles btnOnline.Click
        'Connect to or disconnect from the Application Network.
        If ConnectedToComnet = False Then
            ConnectToComNet()
        Else
            DisconnectFromComNet()
        End If
    End Sub

    Private Sub ConnectToComNet()
        'Connect to the Message Service. (ComNet)

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        'UPDATE 14 Feb 2021 - If the VS2019 version of the ADVL Network is running it may not detected by ComNetRunning()!
        'Check if the Message Service is running by trying to open a connection:
        Try
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds (8 seconds is too short for a slow computer!)
            ConnectionName = ApplicationInfo.Name 'This name will be modified if it is already used in an existing connection.
            ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)
            If ConnectionName <> "" Then
                Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                btnOnline.Text = "Online"
                btnOnline.ForeColor = Color.ForestGreen
                ConnectedToComNet = True
                SendApplicationInfo()
                SendProjectInfo()
                client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                bgwComCheck.WorkerReportsProgress = True
                bgwComCheck.WorkerSupportsCancellation = True
                If bgwComCheck.IsBusy Then
                    'The ComCheck thread is already running.
                Else
                    bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                End If
                Exit Sub 'Connection made OK
            Else
                'Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                Message.Add("The Andorville™ Network was not found. Attempting to start it." & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            End If
        Catch ex As System.TimeoutException
            Message.Add("Message Service Check Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            Message.Add("Attempting to start the Message Service." & vbCrLf)
        Catch ex As Exception
            Message.Add("Error message: " & ex.Message & vbCrLf)
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            Message.Add("Attempting to start the Message Service." & vbCrLf)
        End Try

        If ComNetRunning() Then
            'The Message Service is Running.
        Else 'The Message Service is NOT running'
            'Start the Message Service:
            If AdvlNetworkAppPath = "" Then
                Message.AddWarning("Andorville™ Network application path is unknown." & vbCrLf)
            Else
                If System.IO.File.Exists(AdvlNetworkExePath) Then 'OK to start the Message Service application:
                    Shell(Chr(34) & AdvlNetworkExePath & Chr(34), AppWinStyle.NormalFocus) 'Start Message Service application with no argument
                Else
                    'Incorrect Message Service Executable path.
                    Message.AddWarning("Andorville™ Network exe file not found. Service not started." & vbCrLf)
                End If
            End If
        End If

        'Try to fix a faulted client state:
        If client.State = ServiceModel.CommunicationState.Faulted Then
            client = Nothing
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        If client.State = ServiceModel.CommunicationState.Faulted Then
            Message.AddWarning("Client state is faulted. Connection not made!" & vbCrLf)
        Else
            Try
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds  (8 seconds is too short for a slow computer!)

                ConnectionName = ApplicationInfo.Name 'This name will be modified if it is already used in an existing connection.
                ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)

                If ConnectionName <> "" Then
                    Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                    btnOnline.Text = "Online"
                    btnOnline.ForeColor = Color.ForestGreen
                    ConnectedToComnet = True
                    SendApplicationInfo()
                    SendProjectInfo()
                    client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                    bgwComCheck.WorkerReportsProgress = True
                    bgwComCheck.WorkerSupportsCancellation = True
                    If bgwComCheck.IsBusy Then
                        'The ComCheck thread is already running.
                    Else
                        bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                    End If

                Else
                    Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
                End If
            Catch ex As System.TimeoutException
                Message.Add("Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
            Catch ex As Exception
                Message.Add("Error message: " & ex.Message & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeaout to 1 hour
            End Try
        End If
    End Sub

    Private Sub ConnectToComNet(ByVal ConnName As String)
        'Connect to the Message Service (ComNet) with the connection name ConnName.

        'UPDATE 14 Feb 2021 - If the VS2019 version of the ADVL Network is running it may not be detected by ComNetRunning()!
        'Check if the Message Service is running by trying to open a connection:

        If IsNothing(client) Then
            client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
        End If

        Try
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds (8 seconds is too short for a slow computer!)
            ConnectionName = ConnName 'This name will be modified if it is already used in an existing connection.
            ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)
            If ConnectionName <> "" Then
                Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                btnOnline.Text = "Online"
                btnOnline.ForeColor = Color.ForestGreen
                ConnectedToComNet = True
                SendApplicationInfo()
                SendProjectInfo()
                client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                bgwComCheck.WorkerReportsProgress = True
                bgwComCheck.WorkerSupportsCancellation = True
                If bgwComCheck.IsBusy Then
                    'The ComCheck thread is already running.
                Else
                    bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                End If
                Exit Sub 'Connection made OK
            Else
                'Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                Message.Add("The Andorville™ Network was not found. Attempting to start it." & vbCrLf)
                client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            End If
        Catch ex As System.TimeoutException
            Message.Add("Message Service Check Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            Message.Add("Attempting to start the Message Service." & vbCrLf)
        Catch ex As Exception
            Message.Add("Error message: " & ex.Message & vbCrLf)
            client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
            Message.Add("Attempting to start the Message Service." & vbCrLf)
        End Try


        If ConnectedToComnet = False Then
            'Dim Result As Boolean

            If IsNothing(client) Then
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            'Try to fix a faulted client state:
            If client.State = ServiceModel.CommunicationState.Faulted Then
                client = Nothing
                client = New ServiceReference1.MsgServiceClient(New System.ServiceModel.InstanceContext(New MsgServiceCallback))
            End If

            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.SetWarningStyle()
                Message.Add("client state is faulted. Connection not made!" & vbCrLf)
            Else
                Try
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(0, 0, 16) 'Temporarily set the send timeaout to 16 seconds (8 seconds is too short for a slow computer!)
                    ConnectionName = ConnName 'This name will be modified if it is already used in an existing connection.
                    ConnectionName = client.Connect(ProNetName, ApplicationInfo.Name, ConnectionName, Project.Name, Project.Description, Project.Type, Project.Path, False, False)

                    If ConnectionName <> "" Then
                        Message.Add("Connected to the Andorville™ Network with Connection Name: [" & ProNetName & "]." & ConnectionName & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                        btnOnline.Text = "Online"
                        btnOnline.ForeColor = Color.ForestGreen
                        ConnectedToComnet = True
                        SendApplicationInfo()
                        SendProjectInfo()
                        client.GetAdvlNetworkAppInfoAsync() 'Update the Exe Path in case it has changed. This path may be needed in the future to start the ComNet (Message Service).

                        bgwComCheck.WorkerReportsProgress = True
                        bgwComCheck.WorkerSupportsCancellation = True
                        If bgwComCheck.IsBusy Then
                            'The ComCheck thread is already running.
                        Else
                            bgwComCheck.RunWorkerAsync() 'Start the ComCheck thread.
                        End If

                    Else
                        Message.Add("Connection to the Andorville™ Network failed!" & vbCrLf)
                        client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                    End If
                Catch ex As System.TimeoutException
                    Message.Add("Timeout error. Check if the Andorville™ Network (Message Service) is running." & vbCrLf)
                Catch ex As Exception
                    Message.Add("Error message: " & ex.Message & vbCrLf)
                    client.Endpoint.Binding.SendTimeout = New System.TimeSpan(1, 0, 0) 'Restore the send timeout to 1 hour
                End Try
            End If
        Else
            Message.AddWarning("Already connected to the Andorville™ Network (Message Service)." & vbCrLf)
        End If

    End Sub

    Private Sub DisconnectFromComNet()
        'Disconnect from the Communication Network (Message Service).

        If ConnectedToComnet = True Then
            If IsNothing(client) Then
                'Message.Add("Already disconnected from the Application Network." & vbCrLf)
                Message.Add("Already disconnected from the Andorville™ Network (Message Service)." & vbCrLf)
                btnOnline.Text = "Offline"
                btnOnline.ForeColor = Color.Red
                ConnectedToComnet = False
                ConnectionName = ""
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted." & vbCrLf)
                    ConnectionName = ""
                Else
                    Try
                        'client.Disconnect(AppNetName, ConnectionName)
                        client.Disconnect(ProNetName, ConnectionName)

                        btnOnline.Text = "Offline"
                        btnOnline.ForeColor = Color.Red
                        ConnectedToComnet = False
                        ConnectionName = ""
                        'Message.Add("Disconnected from the Application Network." & vbCrLf)
                        Message.Add("Disconnected from the Andorville™ Network (Message Service)." & vbCrLf)

                        If bgwComCheck.IsBusy Then
                            bgwComCheck.CancelAsync()
                        End If

                    Catch ex As Exception
                        Message.SetWarningStyle()
                        'Message.Add("Error disconnecting from Application Network: " & ex.Message & vbCrLf)
                        Message.AddWarning("Error disconnecting from Andorville™ Network (Message Service): " & ex.Message & vbCrLf)
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub SendApplicationInfo()
        'Send the application information to the Administrator connections.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to send application information.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                Dim applicationInfo As New XElement("ApplicationInfo")
                Dim name As New XElement("Name", Me.ApplicationInfo.Name)
                applicationInfo.Add(name)

                Dim text As New XElement("Text", "Document Library")
                applicationInfo.Add(text)

                Dim exePath As New XElement("ExecutablePath", Me.ApplicationInfo.ExecutablePath)
                applicationInfo.Add(exePath)

                Dim directory As New XElement("Directory", Me.ApplicationInfo.ApplicationDir)
                applicationInfo.Add(directory)
                Dim description As New XElement("Description", Me.ApplicationInfo.Description)
                applicationInfo.Add(description)
                xmessage.Add(applicationInfo)
                doc.Add(xmessage)

                'Show the message sent to AppNet:
                Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage("", "MessageService", doc.ToString)
            End If
        End If

    End Sub

    Private Sub SendProjectInfo()
        'Send the project information to the Network application.

        If ConnectedToComnet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    Dim Path As New XElement("Path", Project.Path)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to the Message Service:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

    Public Sub SendProjectInfo(ByVal ProjectPath As String)
        'Send the project information to the Network application.
        'This version of SendProjectInfo uses the ProjectPath argument.

        If ConnectedToComnet = False Then
            Message.AddWarning("The application is not connected to the Message Service." & vbCrLf)
        Else 'Connected to the Message Service (ComNet).
            If IsNothing(client) Then
                Message.Add("No client connection available!" & vbCrLf)
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("Client state is faulted. Message not sent!" & vbCrLf)
                Else
                    'Construct the XMessage to send to AppNet:
                    Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                    Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                    Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class
                    Dim projectInfo As New XElement("ProjectInfo")

                    'Dim Path As New XElement("Path", Project.Path)
                    Dim Path As New XElement("Path", ProjectPath)
                    projectInfo.Add(Path)
                    xmessage.Add(projectInfo)
                    doc.Add(xmessage)

                    'Show the message sent to the Message Service:
                    Message.XAddText("Message sent to " & "Message Service" & ":" & vbCrLf, "XmlSentNotice")
                    Message.XAddXml(doc.ToString)
                    Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    client.SendMessage("", "MessageService", doc.ToString)
                End If
            End If
        End If
    End Sub

    Private Function ComNetRunning() As Boolean
        'Return True if ComNet (Message Service) is running.
        ''If System.IO.File.Exists(MsgServiceAppPath & "\Application.Lock") Then
        'If System.IO.File.Exists(AdvlNetworkAppPath & "\Application.Lock") Then
        '    Return True
        'Else
        '    Return False
        'End If

        If AdvlNetworkAppPath = "" Then
            Message.Add("Andorville™ Network application path is not known." & vbCrLf)
            Message.Add("Run the Andorville™ Network before connecting to update the path." & vbCrLf)
            Return False
        Else
            If System.IO.File.Exists(AdvlNetworkAppPath & "\Application.Lock") Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

#End Region 'Online/Offline code

#Region " Process XMessages"

    Private Sub XMsg_Instruction(Data As String, Locn As String) Handles XMsg.Instruction
        'Process an XMessage instruction.
        'An XMessage is a simplified XSequence. It is used to exchange information between Andorville™ applications.
        '
        'An XSequence file is an AL-H7™ Information Sequence stored in an XML format.
        'AL-H7™ is the name of a programming system that uses sequences of data and location value pairs to store information or processing steps.
        'Any program, mathematical expression or data set can be expressed as an Information Sequence.

        'Add code here to process the XMessage instructions.
        'See other Andorville™ applciations for examples.
        If IsDBNull(Data) Then
            Data = ""
        End If

        'Intercept instructions with the prefix "WebPage_"
        If Locn.StartsWith("WebPage_") Then 'Send the Info, Location data to the correct Web Page:
            'Message.Add("Web Page Location: " & Locn & vbCrLf)
            If Locn.Contains(":") Then
                Dim EndOfWebPageNoString As Integer = Locn.IndexOf(":")
                If Locn.Contains("-") Then
                    Dim HyphenLocn As Integer = Locn.IndexOf("-")
                    If HyphenLocn < EndOfWebPageNoString Then 'Web Page Location contains a sub-location in the web page - WebPage_1-SubLocn:Locn - SubLocn:Locn will be sent to Web page 1
                        EndOfWebPageNoString = HyphenLocn
                    End If
                End If
                Dim PageNoLen As Integer = EndOfWebPageNoString - 8
                Dim WebPageNoString As String = Locn.Substring(8, PageNoLen)
                Dim WebPageNo As Integer = CInt(WebPageNoString)
                Dim WebPageData As String = Data
                Dim WebPageLocn As String = Locn.Substring(EndOfWebPageNoString + 1)

                'Message.Add("WebPageData = " & WebPageData & "  WebPageLocn = " & WebPageLocn & vbCrLf)

                WebPageFormList(WebPageNo).XMsgInstruction(WebPageData, WebPageLocn)
            Else
                Message.AddWarning("XMessage instruction location is not complete: " & Locn & vbCrLf)
            End If
        Else

            Select Case Locn

                'Case "ClientAppNetName"
                '    ClientAppNetName = Data 'The name of the Client Application Network requesting service. ADDED 2Feb19.
                Case "ClientProNetName"
                    ClientProNetName = Data 'The name of the Client Project Network requesting service.

                Case "ClientName"
                    ClientAppName = Data 'The name of the Client requesting service.

                Case "ClientConnectionName"
                    ClientConnName = Data 'The name of the client requesting service.

                Case "ClientLocn" 'The Location within the Client requesting service.
                    Dim statusOK As New XElement("Status", "OK") 'Add Status OK element when the Client Location is changed
                    xlocns(xlocns.Count - 1).Add(statusOK)
                    xmessage.Add(xlocns(xlocns.Count - 1)) 'Add the instructions for the last location to the reply xmessage
                    xlocns.Add(New XElement(Data)) 'Start the new location instructions

                'Case "OnCompletion" 'Specify the last instruction to be returned on completion of the XMessage processing.
                '    CompletionInstruction = Data

                    'UPDATE:
                Case "OnCompletion"
                    OnCompletionInstruction = Data


                Case "Main"
                 'Blank message - do nothing.

                'Case "Main:OnCompletion"
                '    Select Case "Stop"
                '        'Stop on completion of the instruction sequence.
                '    End Select

                Case "Main:EndInstruction"
                    Select Case Data
                        Case "Stop"
                            'Stop at the end of the instruction sequence.

                            'Add other cases here:
                    End Select


                Case "Main:Status"
                    Select Case Data
                        Case "OK"
                            'Main instructions completed OK
                    End Select

                Case "Command"
                    Select Case Data
                        Case "ConnectToAppNet" 'Startup Command
                            If ConnectedToComNet = False Then
                                ConnectToComNet()
                            End If

                        Case "AppComCheck"
                            'Add the Appplication Communication info to the reply message:
                            Dim clientProNetName As New XElement("ClientProNetName", ProNetName) 'The Project Network Name
                            xlocns(xlocns.Count - 1).Add(clientProNetName)
                            Dim clientName As New XElement("ClientName", "ADVL_Document_Library_1") 'The name of this application.
                            xlocns(xlocns.Count - 1).Add(clientName)
                            Dim clientConnectionName As New XElement("ClientConnectionName", ConnectionName)
                            xlocns(xlocns.Count - 1).Add(clientConnectionName)
                            '<Status>OK</Status> will be automatically appended to the XMessage before it is sent.

                    End Select


               'Startup Command Arguments ================================================
                Case "ProjectName"
                    If Project.OpenProject(Data) = True Then
                        ProjectSelected = True 'Project has been opened OK.
                    Else
                        ProjectSelected = False 'Project could not be opened.
                    End If

                Case "ProjectID"
                    Message.AddWarning("Add code to handle ProjectID parameter at StartUp!" & vbCrLf)

                Case "ProjectPath"
                    If Project.OpenProjectPath(Data) = True Then
                        ProjectSelected = True 'Project has been opened OK.
                    Else
                        ProjectSelected = False 'Project could not be opened.
                    End If

                Case "ConnectionName"
                    StartupConnectionName = Data
            '--------------------------------------------------------------------------

            'Application Information  =================================================
            'returned by client.GetAdvlNetworkAppInfoAsync()
                'Case "MessageServiceAppInfo:Name"
                ''The name of the Message Service Application. (Not used.)
                Case "AdvlNetworkAppInfo:Name"
                'The name of the Andorville™ Network Application. (Not used.)

                'Case "MessageServiceAppInfo:ExePath"
                '    'The executable file path of the Message Service Application.
                '    'MsgServiceExePath = Info
                Case "AdvlNetworkAppInfo:ExePath"
                    'The executable file path of the Andorville™ Network Application.
                    AdvlNetworkExePath = Data

                'Case "MessageServiceAppInfo:Path"
                '    'The path of the Message Service Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
                '    'MsgServiceAppPath = Info
                Case "AdvlNetworkAppInfo:Path"
                    'The path of the Andorville™ Network Application (ComNet). (This is where an Application.Lock file will be found while ComNet is running.)
                    AdvlNetworkAppPath = Data
           '---------------------------------------------------------------------------

           'Show Document =============================================================
                Case "ShowDocument:FileName"
                    SearchFileName = Data
                Case "ShowDocument:SearchText"
                    SearchText = Data
                Case "ShowDocument:Highlight"
                    If Data = "True" Then
                        SearchHighlight = True
                    Else
                        SearchHighlight = False
                    End If
                Case "ShowDocument:FindFirst"
                    If Data = "True" Then
                        SearchFindFirst = True
                    Else
                        SearchFindFirst = False
                    End If
                Case "ShowDocument:Command"
                    Select Case Data
                        Case "ShowInMainForm"
                            FindTextInMainRtf(SearchFileName, SearchText, SearchHighlight, SearchFindFirst)
                            TabControl1.SelectedIndex = 1 'Select the Library tab.
                            TabControl3.SelectedIndex = 4 'Select the Document View tab.
                        Case "ShowInDocumentWindow"
                            FindTextInRtf(SearchFileName, SearchText, SearchHighlight, SearchFindFirst)
                    End Select

           '---------------------------------------------------------------------------

               'Message Window Instructions  ==============================================
                Case "MessageWindow:Left"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Left = Data
                Case "MessageWindow:Top"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Top = Data
                Case "MessageWindow:Width"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Width = Data
                Case "MessageWindow:Height"
                    If IsNothing(Message.MessageForm) Then
                        Message.ApplicationName = ApplicationInfo.Name
                        Message.SettingsLocn = Project.SettingsLocn
                        Message.Show()
                    End If
                    Message.MessageForm.Height = Data
                Case "MessageWindow:Command"
                    Select Case Data
                        Case "BringToFront"
                            If IsNothing(Message.MessageForm) Then
                                Message.ApplicationName = ApplicationInfo.Name
                                Message.SettingsLocn = Project.SettingsLocn
                                Message.Show()
                            End If
                            'Message.MessageForm.BringToFront()
                            Message.MessageForm.Activate()
                            Message.MessageForm.TopMost = True
                            Message.MessageForm.TopMost = False
                        Case "SaveSettings"
                            Message.MessageForm.SaveFormSettings()
                    End Select

            '---------------------------------------------------------------------------

               'Command to bring the Application window to the front:
                Case "ApplicationWindow:Command"
                    Select Case Data
                        Case "BringToFront"
                            Me.Activate()
                            Me.TopMost = True
                            Me.TopMost = False
                    End Select

                Case "EndOfSequence"
                    'End of Information Vector Sequence reached.
                    'Add Status OK element at the end of the sequence:
                    Dim statusOK As New XElement("Status", "OK")
                    xlocns(xlocns.Count - 1).Add(statusOK)

                    'Clear the Search Document variables:
                    SearchFileName = ""
                    SearchText = ""
                    SearchHighlight = True
                    SearchFindFirst = True

                    Select Case EndInstruction
                        Case "Stop"
                            'No instructions.

                            'Add any other Cases here:

                        Case Else
                            Message.AddWarning("Unknown End Instruction: " & EndInstruction & vbCrLf)
                    End Select
                    EndInstruction = "Stop"

                    ''Add the final OnCompletion instruction:
                    'Dim onCompletion As New XElement("OnCompletion", CompletionInstruction) '
                    'xlocns(xlocns.Count - 1).Add(onCompletion)
                    'CompletionInstruction = "Stop" 'Reset the Completion Instruction

                    ''Final Version:
                    ''Add the final EndInstruction:
                    'Dim xEndInstruction As New XElement("EndInstruction", OnCompletionInstruction)
                    'xlocns(xlocns.Count - 1).Add(xEndInstruction)
                    'OnCompletionInstruction = "Stop" 'Reset the OnCompletion Instruction

                    'Add the final EndInstruction:
                    If OnCompletionInstruction = "Stop" Then
                        'Final EndInstruction is not required.
                    Else
                        Dim xEndInstruction As New XElement("EndInstruction", OnCompletionInstruction)
                        xlocns(xlocns.Count - 1).Add(xEndInstruction)
                        OnCompletionInstruction = "Stop" 'Reset the OnCompletion Instruction
                    End If


                Case Else
                    Message.AddWarning("Unknown location: " & Locn & vbCrLf)
                    Message.AddWarning("            data: " & Data & vbCrLf)
            End Select
        End If

    End Sub

    'Private Sub SendMessage()
    '    'Code used to send a message after a timer delay.
    '    'The message destination is stored in MessageDest
    '    'The message text is stored in MessageText
    '    Timer1.Interval = 100 '100ms delay
    '    Timer1.Enabled = True 'Start the timer.
    'End Sub

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    '    If IsNothing(client) Then
    '        Message.AddWarning("No client connection available!" & vbCrLf)
    '    Else
    '        If client.State = ServiceModel.CommunicationState.Faulted Then
    '            Message.AddWarning("client state is faulted. Message not sent!" & vbCrLf)
    '        Else
    '            Try
    '                'client.SendMessage(ClientAppNetName, ClientConnName, MessageText) 'Added 2Feb19
    '                client.SendMessage(ClientProNetName, ClientConnName, MessageText)
    '                MessageText = "" 'Clear the message after it has been sent.
    '                ClientAppName = "" 'Clear the Client Application Name after the message has been sent.
    '                ClientConnName = "" 'Clear the Client Application Name after the message has been sent.
    '                xlocns.Clear()
    '            Catch ex As Exception
    '                Message.AddWarning("Error sending message: " & ex.Message & vbCrLf)
    '            End Try
    '        End If
    '    End If

    '    'Stop timer:
    '    Timer1.Enabled = False
    'End Sub


#End Region 'Process XMessages

    'Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
    '    OpenDocumentFile()
    'End Sub

    Private Sub OpenDocumentFile()
        'Open a document.

        If DocumentTextChanged = True Then
            Dim result As Integer = MessageBox.Show("Save changes to the current document?", "Notice", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                Exit Sub
            ElseIf result = DialogResult.Yes Then
                SaveDocument()
            ElseIf result = DialogResult.No Then
                'Do not save the changes!
            End If
        End If

        Select Case FileTypeSelection
            Case FileTypes.RTF
                If rbFileInProject.Checked = True Then
                    'Look for the file in the Project.
                    Dim SelectedFile As String = Project.SelectDataFile("Rich text format", "rtf")
                    If SelectedFile = "" Then
                        'No file selected!
                    Else
                        RtfFileName = SelectedFile
                        RtfFileLocationType = LocationTypes.Project
                        Dim rtbData As New IO.MemoryStream
                        Project.ReadData(RtfFileName, rtbData)
                        If rtbData.Length = 0 Then
                            Message.AddWarning("No data read!" & vbCrLf)
                            Exit Sub
                        End If
                        XmlHtmDisplay1.Clear()
                        rtbData.Position = 0
                        XmlHtmDisplay1.LoadFile(rtbData, RichTextBoxStreamType.RichText)
                        DocumentTextChanged = False
                        FileName = RtfFileName
                        FileLocationType = LocationTypes.Project
                        FileType = FileTypes.RTF
                        LastRtfFileName = RtfFileName
                        LastRtfFileLocationType = LocationTypes.Project
                    End If
                Else
                    'Look for the file in the File System.
                    If LastRtfFileName = "" Then
                        'There is no last RTF file saved.
                    Else
                        'The last RTF file path was saved.
                        OpenFileDialog1.InitialDirectory = LastRtfFileDirectory
                        OpenFileDialog1.FileName = LastRtfFileName
                    End If
                    OpenFileDialog1.Filter = "Rich Text Format (*.rtf)| *.rtf"
                    SendKeys.Send("{HOME}") 'To move the cursor to the start of the FileName
                    If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                        RtfFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                        RtfFileLocationType = LocationTypes.FileSystem
                        RtfFileDirectory = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
                        XmlHtmDisplay1.LoadFile(RtfFileDirectory & "\" & RtfFileName)
                        DocumentTextChanged = False
                        FileName = RtfFileName
                        FileLocationType = LocationTypes.FileSystem
                        FileDirectory = RtfFileDirectory
                        FileType = FileTypes.RTF
                        LastRtfFileName = RtfFileName
                        LastRtfFileLocationType = LocationTypes.FileSystem
                        LastRtfFileDirectory = RtfFileDirectory
                    End If
                End If

            Case FileTypes.TXT

            Case FileTypes.XML
                If rbFileInProject.Checked = True Then
                    'Look for the file in the Project.
                    Dim SelectedFile As String = Project.SelectDataFile("Extensible markup language", "xml")
                    If SelectedFile = "" Then
                        'No file selected!
                    Else
                        XmlFileName = SelectedFile
                        XmlFileLocationType = LocationTypes.Project
                        Dim xmlDoc As New System.Xml.XmlDocument
                        Project.ReadXmlDocData(XmlFileName, xmlDoc)
                        XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlDoc, True)
                        DocumentTextChanged = False
                        FileName = XmlFileName
                        FileLocationType = LocationTypes.Project
                        FileType = FileTypes.XML
                        LastXmlFileName = XmlFileName
                        LastXmlFileLocationType = LocationTypes.Project
                    End If
                Else
                    'Look for the file in the File System.
                    If LastXmlFileName = "" Then
                        'There is no last XML file saved.
                    Else
                        'The last XML file path was saved.
                        OpenFileDialog1.InitialDirectory = LastXmlFileDirectory
                        OpenFileDialog1.FileName = LastXmlFileName
                    End If
                    OpenFileDialog1.Filter = "All Files (*.*)| *.*"
                    SendKeys.Send("{HOME}") 'To move the cursor to the start of the FileName
                    If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                        XmlFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                        XmlFileLocationType = LocationTypes.FileSystem
                        XmlFileDirectory = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
                        'Read file as XML:
                        XmlHtmDisplay1.ReadXmlFile(XmlFileDirectory & "\" & XmlFileName, False)
                        DocumentTextChanged = False
                        FileName = XmlFileName
                        FileLocationType = LocationTypes.FileSystem
                        FileType = FileTypes.XML
                        LastXmlFileName = XmlFileName
                        LastXmlFileLocationType = LocationTypes.FileSystem
                        LastXmlFileDirectory = XmlFileDirectory
                        'Message.Add("XmlTextChanged restored to False" & vbCrLf)
                    End If
                End If
            Case FileTypes.HTML
                If rbFileInProject.Checked = True Then
                    'Look for the file in the Project.
                    Dim SelectedFile As String = Project.SelectDataFile("HTML File", "html")
                    If SelectedFile = "" Then
                        'No file selected!
                    Else
                        HtmlFileName = SelectedFile
                        HtmlFileLocationType = LocationTypes.Project
                        Dim htmlData As New IO.MemoryStream
                        Project.ReadData(HtmlFileName, htmlData)
                        XmlHtmDisplay1.Clear()
                        htmlData.Position = 0
                        XmlHtmDisplay1.LoadFile(htmlData, RichTextBoxStreamType.PlainText)
                        Dim htmText As String = XmlHtmDisplay1.Text
                        XmlHtmDisplay1.Rtf = XmlHtmDisplay1.HmlToRtf(htmText)
                        DocumentTextChanged = False
                        FileName = HtmlFileName
                        FileLocationType = LocationTypes.Project
                        FileType = FileTypes.HTML
                        LastHtmlFileName = HtmlFileName
                        LastHtmlFileLocationType = LocationTypes.Project
                    End If
                Else
                    'Look for the file in the File System.
                    If LastHtmlFileName = "" Then
                        'There is no last HTML file saved.
                    Else
                        'The last HTML file path was saved.
                        OpenFileDialog1.InitialDirectory = LastHtmlFileDirectory
                        OpenFileDialog1.FileName = LastHtmlFileName
                    End If
                    OpenFileDialog1.Filter = "Html Files | *.html"
                    SendKeys.Send("{HOME}") 'To move the cursor to the start of the FileName
                    If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                        HtmlFileName = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                        HtmlFileLocationType = LocationTypes.FileSystem
                        HtmlFileDirectory = System.IO.Path.GetDirectoryName(OpenFileDialog1.FileName)
                        'Read file as HTML:
                        XmlHtmDisplay1.ReadXmlFile(XmlFileDirectory & "\" & XmlFileName, False)
                        XmlHtmDisplay1.LoadFile(XmlFileDirectory & "\" & XmlFileName, RichTextBoxStreamType.PlainText)
                        Dim htmText As String = XmlHtmDisplay1.Text
                        XmlHtmDisplay1.Rtf = XmlHtmDisplay1.HmlToRtf(htmText)
                        DocumentTextChanged = False
                        FileName = HtmlFileName
                        FileLocationType = LocationTypes.FileSystem
                        FileType = FileTypes.HTML
                        LastHtmlFileName = HtmlFileName
                        LastHtmlFileLocationType = LocationTypes.FileSystem
                        LastHtmlFileDirectory = HtmlFileDirectory
                    End If
                End If
        End Select
    End Sub

    Private Sub ShowSettings()
        'Show the XML, HTML and RTF display settings in the Settings tab.

        XmlHtmDisplay1.Settings.UpdateFontIndexes()
        XmlHtmDisplay1.Settings.UpdateColorIndexes()

        'Show the XML and other Text Type display settings
        txtXmlIndentSpaces.Text = XmlHtmDisplay1.Settings.XIndentSpaces
        txtXmlFileSizeLimit.Text = XmlHtmDisplay1.Settings.XmlLargeFileSizeLimit

        DataGridView1.Rows.Clear()
        'Show XML Document display settings:
        'DataGridView1.Rows.Add("Row1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8", "Col9")
        DataGridView1.Rows.Add("XML Tag", XmlHtmDisplay1.Settings.XTag.FontName, XmlHtmDisplay1.Settings.XTag.FontIndex, XmlHtmDisplay1.Settings.XTag.Color.Name, XmlHtmDisplay1.Settings.XTag.ColorIndex, XmlHtmDisplay1.Settings.XTag.HalfPointSize, XmlHtmDisplay1.Settings.XTag.PointSize, XmlHtmDisplay1.Settings.XTag.Bold.ToString, XmlHtmDisplay1.Settings.XTag.Italic.ToString)
        DataGridView1.Rows.Add("XML Value", XmlHtmDisplay1.Settings.XValue.FontName, XmlHtmDisplay1.Settings.XValue.FontIndex, XmlHtmDisplay1.Settings.XValue.Color.Name, XmlHtmDisplay1.Settings.XValue.ColorIndex, XmlHtmDisplay1.Settings.XValue.HalfPointSize, XmlHtmDisplay1.Settings.XValue.PointSize, XmlHtmDisplay1.Settings.XValue.Bold.ToString, XmlHtmDisplay1.Settings.XValue.Italic.ToString)
        DataGridView1.Rows.Add("XML Comment", XmlHtmDisplay1.Settings.XComment.FontName, XmlHtmDisplay1.Settings.XComment.FontIndex, XmlHtmDisplay1.Settings.XComment.Color.Name, XmlHtmDisplay1.Settings.XComment.ColorIndex, XmlHtmDisplay1.Settings.XComment.HalfPointSize, XmlHtmDisplay1.Settings.XComment.PointSize, XmlHtmDisplay1.Settings.XComment.Bold.ToString, XmlHtmDisplay1.Settings.XComment.Italic.ToString)
        DataGridView1.Rows.Add("XML Element", XmlHtmDisplay1.Settings.XElement.FontName, XmlHtmDisplay1.Settings.XElement.FontIndex, XmlHtmDisplay1.Settings.XElement.Color.Name, XmlHtmDisplay1.Settings.XElement.ColorIndex, XmlHtmDisplay1.Settings.XElement.HalfPointSize, XmlHtmDisplay1.Settings.XElement.PointSize, XmlHtmDisplay1.Settings.XElement.Bold.ToString, XmlHtmDisplay1.Settings.XElement.Italic.ToString)
        DataGridView1.Rows.Add("XML Attribute Key", XmlHtmDisplay1.Settings.XAttributeKey.FontName, XmlHtmDisplay1.Settings.XAttributeKey.FontIndex, XmlHtmDisplay1.Settings.XAttributeKey.Color.Name, XmlHtmDisplay1.Settings.XAttributeKey.ColorIndex, XmlHtmDisplay1.Settings.XAttributeKey.HalfPointSize, XmlHtmDisplay1.Settings.XAttributeKey.PointSize, XmlHtmDisplay1.Settings.XAttributeKey.Bold.ToString, XmlHtmDisplay1.Settings.XAttributeKey.Italic.ToString)
        'DataGridView1.Rows.Add("XML Attribute Value", XmlDisplay1.Settings.AttributeValue.FontName, XmlDisplay1.Settings.AttributeValue.FontIndex, XmlDisplay1.Settings.AttributeValue.Color.Name, XmlDisplay1.Settings.AttributeValue.ColorIndex, XmlDisplay1.Settings.AttributeValue.HalfPointSize, XmlDisplay1.Settings.AttributeValue.PointSize, XmlDisplay1.Settings.AttributeValue.Bold, XmlDisplay1.Settings.AttributeValue.Italic)
        DataGridView1.Rows.Add("XML Attribute Value", XmlHtmDisplay1.Settings.XAttributeValue.FontName, XmlHtmDisplay1.Settings.XAttributeValue.FontIndex, XmlHtmDisplay1.Settings.XAttributeValue.Color.Name, XmlHtmDisplay1.Settings.XAttributeValue.ColorIndex, XmlHtmDisplay1.Settings.XAttributeValue.HalfPointSize, XmlHtmDisplay1.Settings.XAttributeValue.PointSize, XmlHtmDisplay1.Settings.XAttributeValue.Bold.ToString, XmlHtmDisplay1.Settings.XAttributeValue.Italic.ToString)

        'Show HTML document display settings:
        DataGridView1.Rows.Add("HTML Text", XmlHtmDisplay1.Settings.HText.FontName, XmlHtmDisplay1.Settings.HText.FontIndex, XmlHtmDisplay1.Settings.HText.Color.Name, XmlHtmDisplay1.Settings.HText.ColorIndex, XmlHtmDisplay1.Settings.HText.HalfPointSize, XmlHtmDisplay1.Settings.HText.PointSize, XmlHtmDisplay1.Settings.HText.Bold.ToString, XmlHtmDisplay1.Settings.HText.Italic.ToString)
        DataGridView1.Rows.Add("HTML Element", XmlHtmDisplay1.Settings.HElement.FontName, XmlHtmDisplay1.Settings.HElement.FontIndex, XmlHtmDisplay1.Settings.HElement.Color.Name, XmlHtmDisplay1.Settings.HElement.ColorIndex, XmlHtmDisplay1.Settings.HElement.HalfPointSize, XmlHtmDisplay1.Settings.HElement.PointSize, XmlHtmDisplay1.Settings.HElement.Bold.ToString, XmlHtmDisplay1.Settings.HElement.Italic.ToString)
        DataGridView1.Rows.Add("HTML Attribute", XmlHtmDisplay1.Settings.HAttribute.FontName, XmlHtmDisplay1.Settings.HAttribute.FontIndex, XmlHtmDisplay1.Settings.HAttribute.Color.Name, XmlHtmDisplay1.Settings.HAttribute.ColorIndex, XmlHtmDisplay1.Settings.HAttribute.HalfPointSize, XmlHtmDisplay1.Settings.HAttribute.PointSize, XmlHtmDisplay1.Settings.HAttribute.Bold.ToString, XmlHtmDisplay1.Settings.HAttribute.Italic.ToString)
        DataGridView1.Rows.Add("HTML Comment", XmlHtmDisplay1.Settings.HComment.FontName, XmlHtmDisplay1.Settings.HComment.FontIndex, XmlHtmDisplay1.Settings.HComment.Color.Name, XmlHtmDisplay1.Settings.HComment.ColorIndex, XmlHtmDisplay1.Settings.HComment.HalfPointSize, XmlHtmDisplay1.Settings.HComment.PointSize, XmlHtmDisplay1.Settings.HComment.Bold.ToString, XmlHtmDisplay1.Settings.HComment.Italic.ToString)
        DataGridView1.Rows.Add("HTML Special Chars", XmlHtmDisplay1.Settings.HChar.FontName, XmlHtmDisplay1.Settings.HChar.FontIndex, XmlHtmDisplay1.Settings.HChar.Color.Name, XmlHtmDisplay1.Settings.HChar.ColorIndex, XmlHtmDisplay1.Settings.HChar.HalfPointSize, XmlHtmDisplay1.Settings.HChar.PointSize, XmlHtmDisplay1.Settings.HChar.Bold.ToString, XmlHtmDisplay1.Settings.HChar.Italic.ToString)
        DataGridView1.Rows.Add("HTML Value", XmlHtmDisplay1.Settings.HValue.FontName, XmlHtmDisplay1.Settings.HValue.FontIndex, XmlHtmDisplay1.Settings.HValue.Color.Name, XmlHtmDisplay1.Settings.HValue.ColorIndex, XmlHtmDisplay1.Settings.HValue.HalfPointSize, XmlHtmDisplay1.Settings.HValue.PointSize, XmlHtmDisplay1.Settings.HValue.Bold.ToString, XmlHtmDisplay1.Settings.HValue.Italic.ToString)
        DataGridView1.Rows.Add("HTML Style", XmlHtmDisplay1.Settings.HStyle.FontName, XmlHtmDisplay1.Settings.HStyle.FontIndex, XmlHtmDisplay1.Settings.HStyle.Color.Name, XmlHtmDisplay1.Settings.HStyle.ColorIndex, XmlHtmDisplay1.Settings.HStyle.HalfPointSize, XmlHtmDisplay1.Settings.HStyle.PointSize, XmlHtmDisplay1.Settings.HStyle.Bold.ToString, XmlHtmDisplay1.Settings.HStyle.Italic.ToString)

        DataGridView1.Rows.Add("TXT Plain Text", XmlHtmDisplay1.Settings.PlainText.FontName, XmlHtmDisplay1.Settings.PlainText.FontIndex, XmlHtmDisplay1.Settings.PlainText.Color.Name, XmlHtmDisplay1.Settings.PlainText.ColorIndex, XmlHtmDisplay1.Settings.PlainText.HalfPointSize, XmlHtmDisplay1.Settings.PlainText.PointSize, XmlHtmDisplay1.Settings.PlainText.Bold.ToString, XmlHtmDisplay1.Settings.PlainText.Italic.ToString)

        DataGridView1.Rows.Add("Default Text", XmlHtmDisplay1.Settings.DefaultText.FontName, XmlHtmDisplay1.Settings.DefaultText.FontIndex, XmlHtmDisplay1.Settings.DefaultText.Color.Name, XmlHtmDisplay1.Settings.DefaultText.ColorIndex, XmlHtmDisplay1.Settings.DefaultText.HalfPointSize, XmlHtmDisplay1.Settings.DefaultText.PointSize, XmlHtmDisplay1.Settings.DefaultText.Bold.ToString, XmlHtmDisplay1.Settings.DefaultText.Italic.ToString)

        Dim I As Integer

        For I = 0 To XmlHtmDisplay1.Settings.TextType.Count - 1
            DataGridView1.Rows.Add(XmlHtmDisplay1.Settings.TextType.ElementAt(I).Key, XmlHtmDisplay1.Settings.TextType.ElementAt(I).Value.FontName, XmlHtmDisplay1.Settings.TextType.ElementAt(I).Value.FontIndex, XmlHtmDisplay1.Settings.TextType.ElementAt(I).Value.Color.Name, XmlHtmDisplay1.Settings.TextType.ElementAt(I).Value.ColorIndex, XmlHtmDisplay1.Settings.TextType.ElementAt(I).Value.HalfPointSize, XmlHtmDisplay1.Settings.TextType.ElementAt(I).Value.PointSize, XmlHtmDisplay1.Settings.TextType.ElementAt(I).Value.Bold.ToString, XmlHtmDisplay1.Settings.TextType.ElementAt(I).Value.Italic.ToString)
        Next

        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

    End Sub


    Private Sub btnSaveCollection_Click(sender As Object, e As EventArgs)
        'SaveDocumentCollection()
    End Sub


    Private Sub btnOpenCollection_Click(sender As Object, e As EventArgs)
        'Find and open a document collection.

    End Sub


    Private Sub btnSaveSettingsFile_Click(sender As Object, e As EventArgs) Handles btnSaveSettingsFile.Click
        'Save the XML and RTF display settings:

        Dim SettingsFileName As String = Trim(txtSettingsFileName.Text)

        If SettingsFileName = "" Then
            Message.AddWarning("Please enter a file name for the settings!" & vbCrLf)
            Exit Sub
        End If

        If SettingsFileName.EndsWith(".XmlHmlRtfSettings") Then

        Else
            'Add file extension to the file name.
            SettingsFileName &= ".XmlHmlRtfSettings"
            txtSettingsFileName.Text = SettingsFileName
        End If

        'Create and save the XML and RTF settings:
        DataGridView1.AllowUserToAddRows = False

        'Find XML text types in DataGridView1:
        Dim XmlTagRowNo As Integer = -1
        Dim XmlValueRowNo As Integer = -1
        Dim XmlCommentRowNo As Integer = -1
        Dim XmlElementRowNo As Integer = -1
        Dim XmlAttributeKeyRowNo As Integer = -1
        Dim XmlAttributeValueRowNo As Integer = -1

        'Find HTML text types in DataGridView1:
        Dim HmlTextRowNo As Integer = -1
        Dim HmlElementRowNo As Integer = -1
        Dim HmlAttributeRowNo As Integer = -1
        Dim HmlCommentRowNo As Integer = -1
        Dim HmlCharsRowNo As Integer = -1
        Dim HmlValueRowNo As Integer = -1
        Dim HmlStyleRowNo As Integer = -1

        Dim TxtPlainTextRowNo As Integer = -1

        Dim DefaultTextRowNo As Integer = -1

        Dim RtfRowNo(0 To DataGridView1.RowCount) As Integer
        Dim NRtfSettings As Integer = 0

        Dim I As Integer

        For I = 0 To DataGridView1.RowCount - 1
            Select Case DataGridView1.Rows(I).Cells(0).Value
                Case "XML Tag"
                    If XmlTagRowNo = -1 Then
                        XmlTagRowNo = I
                    Else
                        Message.AddWarning("Duplicate XML Tag font settings defined!" & vbCrLf)
                    End If
                Case "XML Value"
                    If XmlValueRowNo = -1 Then
                        XmlValueRowNo = I
                    Else
                        Message.AddWarning("Duplicate XML Value font settings defined!" & vbCrLf)
                    End If
                Case "XML Comment"
                    If XmlCommentRowNo = -1 Then
                        XmlCommentRowNo = I
                    Else
                        Message.AddWarning("Duplicate XML Comment font settings defined!" & vbCrLf)
                    End If
                Case "XML Element"
                    If XmlElementRowNo = -1 Then
                        XmlElementRowNo = I
                    Else
                        Message.AddWarning("Duplicate XML Element font settings defined!" & vbCrLf)
                    End If
                Case "XML Attribute Key"
                    If XmlAttributeKeyRowNo = -1 Then
                        XmlAttributeKeyRowNo = I
                    Else
                        Message.AddWarning("Duplicate XML Attribute Key font settings defined!" & vbCrLf)
                    End If
                Case "XML Attribute Value"
                    If XmlAttributeValueRowNo = -1 Then
                        XmlAttributeValueRowNo = I
                    Else
                        Message.AddWarning("Duplicate XML Attribute Value font settings defined!" & vbCrLf)
                    End If
                Case "HTML Text"
                    If HmlTextRowNo = -1 Then
                        HmlTextRowNo = I
                    Else
                        Message.AddWarning("Duplicate HTML Text font settings defined!" & vbCrLf)
                    End If
                Case "HTML Element"
                    If HmlElementRowNo = -1 Then
                        HmlElementRowNo = I
                    Else
                        Message.AddWarning("Duplicate HTML Element font settings defined!" & vbCrLf)
                    End If
                Case "HTML Attribute"
                    If HmlAttributeRowNo = -1 Then
                        HmlAttributeRowNo = I
                    Else
                        Message.AddWarning("Duplicate HTML Attribute font settings defined!" & vbCrLf)
                    End If
                Case "HTML Comment"
                    If HmlCommentRowNo = -1 Then
                        HmlCommentRowNo = I
                    Else
                        Message.AddWarning("Duplicate HTML Comment font settings defined!" & vbCrLf)
                    End If
                Case "HTML Special Chars"
                    If HmlCharsRowNo = -1 Then
                        HmlCharsRowNo = I
                    Else
                        Message.AddWarning("Duplicate HTML Special Characters font settings defined!" & vbCrLf)
                    End If
                Case "HTML Value"
                    If HmlValueRowNo = -1 Then
                        HmlValueRowNo = I
                    Else
                        Message.AddWarning("Duplicate HTML Value font settings defined!" & vbCrLf)
                    End If
                Case "HTML Style"
                    If HmlStyleRowNo = -1 Then
                        HmlStyleRowNo = I
                    Else
                        Message.AddWarning("Duplicate HTML Style font settings defined!" & vbCrLf)
                    End If
                Case "TXT Plain Text"
                    If TxtPlainTextRowNo = -1 Then
                        TxtPlainTextRowNo = I
                    Else
                        Message.AddWarning("Duplicate TXT Plain Text font settings defined!" & vbCrLf)
                    End If
                Case "Default Text"
                    If DefaultTextRowNo = -1 Then
                        DefaultTextRowNo = I
                    Else
                        Message.AddWarning("Duplicate Default Text font settings defined!" & vbCrLf)
                    End If
                Case Else
                    'The Text type is not an XML setting.
                    'Add it the the list of RTF settings:
                    RtfRowNo(NRtfSettings) = I
                    NRtfSettings += 1
            End Select
        Next

        ReDim Preserve RtfRowNo(NRtfSettings - 1)

        Dim xmlRtfSettings = <?xml version="1.0" encoding="utf-8"?>
                             <!---->
                             <!--List of XML and RTF settings.-->
                             <ListOfXmlHmlRtfSettings>
                                 <XmlIndentSpaces><%= txtXmlIndentSpaces.Text %></XmlIndentSpaces>
                                 <XmlLargeFileSizeLimit><%= txtXmlFileSizeLimit.Text %></XmlLargeFileSizeLimit>
                                 <!--XML font settings:-->
                                 <XmlSettings>
                                     <%= If(XmlTagRowNo = -1,
                                     <XmlTag></XmlTag>,
                                     <XmlTag>
                                         <FontName><%= DataGridView1.Rows(XmlTagRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(XmlTagRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(XmlTagRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(XmlTagRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(XmlTagRowNo).Cells(8).Value %></Italic>
                                     </XmlTag>)
                                     %>
                                     <%= If(XmlValueRowNo = -1,
                                     <XmlValue></XmlValue>,
                                     <XmlValue>
                                         <FontName><%= DataGridView1.Rows(XmlValueRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(XmlValueRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(XmlValueRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(XmlValueRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(XmlValueRowNo).Cells(8).Value %></Italic>
                                     </XmlValue>)
                                     %>
                                     <%= If(XmlCommentRowNo = -1,
                                     <XmlComment></XmlComment>,
                                     <XmlComment>
                                         <FontName><%= DataGridView1.Rows(XmlCommentRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(XmlCommentRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(XmlCommentRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(XmlCommentRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(XmlCommentRowNo).Cells(8).Value %></Italic>
                                     </XmlComment>)
                                     %>
                                     <%= If(XmlElementRowNo = -1,
                                     <XmlElement></XmlElement>,
                                     <XmlElement>
                                         <FontName><%= DataGridView1.Rows(XmlElementRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(XmlElementRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(XmlElementRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(XmlElementRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(XmlElementRowNo).Cells(8).Value %></Italic>
                                     </XmlElement>)
                                     %>
                                     <%= If(XmlAttributeKeyRowNo = -1,
                                     <XmlAttributeKey></XmlAttributeKey>,
                                     <XmlAttributeKey>
                                         <FontName><%= DataGridView1.Rows(XmlAttributeKeyRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(XmlAttributeKeyRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(XmlAttributeKeyRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(XmlAttributeKeyRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(XmlAttributeKeyRowNo).Cells(8).Value %></Italic>
                                     </XmlAttributeKey>)
                                     %>
                                     <%= If(XmlAttributeValueRowNo = -1,
                                     <XmlAttributeValue></XmlAttributeValue>,
                                     <XmlAttributeValue>
                                         <FontName><%= DataGridView1.Rows(XmlAttributeValueRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(XmlAttributeValueRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(XmlAttributeValueRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(XmlAttributeValueRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(XmlAttributeValueRowNo).Cells(8).Value %></Italic>
                                     </XmlAttributeValue>)
                                     %>
                                 </XmlSettings>
                                 <HtmlSettings>
                                     <%= If(HmlTextRowNo = -1,
                                     <HtmlText></HtmlText>,
                                     <HtmlText>
                                         <FontName><%= DataGridView1.Rows(HmlTextRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(HmlTextRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(HmlTextRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(HmlTextRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(HmlTextRowNo).Cells(8).Value %></Italic>
                                     </HtmlText>)
                                     %>
                                     <%= If(HmlElementRowNo = -1,
                                     <HtmlElement></HtmlElement>,
                                     <HtmlElement>
                                         <FontName><%= DataGridView1.Rows(HmlElementRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(HmlElementRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(HmlElementRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(HmlElementRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(HmlElementRowNo).Cells(8).Value %></Italic>
                                     </HtmlElement>)
                                     %>
                                     <%= If(HmlAttributeRowNo = -1,
                                     <HtmlAttribute></HtmlAttribute>,
                                     <HtmlAttribute>
                                         <FontName><%= DataGridView1.Rows(HmlAttributeRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(HmlAttributeRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(HmlAttributeRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(HmlAttributeRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(HmlAttributeRowNo).Cells(8).Value %></Italic>
                                     </HtmlAttribute>)
                                     %>
                                     <%= If(HmlCommentRowNo = -1,
                                     <HtmlComment></HtmlComment>,
                                     <HtmlComment>
                                         <FontName><%= DataGridView1.Rows(HmlCommentRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(HmlCommentRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(HmlCommentRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(HmlCommentRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(HmlCommentRowNo).Cells(8).Value %></Italic>
                                     </HtmlComment>)
                                     %>
                                     <%= If(HmlCharsRowNo = -1,
                                     <HtmlChar></HtmlChar>,
                                     <HtmlChar>
                                         <FontName><%= DataGridView1.Rows(HmlCharsRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(HmlCharsRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(HmlCharsRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(HmlCharsRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(HmlCharsRowNo).Cells(8).Value %></Italic>
                                     </HtmlChar>)
                                     %>
                                     <%= If(HmlValueRowNo = -1,
                                     <HtmlValue></HtmlValue>,
                                     <HtmlValue>
                                         <FontName><%= DataGridView1.Rows(HmlValueRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(HmlValueRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(HmlValueRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(HmlValueRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(HmlValueRowNo).Cells(8).Value %></Italic>
                                     </HtmlValue>)
                                     %>
                                     <%= If(HmlStyleRowNo = -1,
                                     <HtmlStyle></HtmlStyle>,
                                     <HtmlStyle>
                                         <FontName><%= DataGridView1.Rows(HmlStyleRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(HmlStyleRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(HmlStyleRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(HmlStyleRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(HmlStyleRowNo).Cells(8).Value %></Italic>
                                     </HtmlStyle>)
                                     %>
                                 </HtmlSettings>
                                 <!--TXT Plain Text font settings:-->
                                 <%= If(TxtPlainTextRowNo = -1,
                                     <TxtPlainText></TxtPlainText>,
                                     <TxtPlainText>
                                         <FontName><%= DataGridView1.Rows(TxtPlainTextRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(TxtPlainTextRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(TxtPlainTextRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(TxtPlainTextRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(TxtPlainTextRowNo).Cells(8).Value %></Italic>
                                     </TxtPlainText>)
                                 %>
                                 <!--Default Text font settings:-->
                                 <%= If(DefaultTextRowNo = -1,
                                     <DefaultText></DefaultText>,
                                     <DefaultText>
                                         <FontName><%= DataGridView1.Rows(DefaultTextRowNo).Cells(1).Value %></FontName>
                                         <Color><%= DataGridView1.Rows(DefaultTextRowNo).Cells(3).Value %></Color>
                                         <HalfPointSize><%= DataGridView1.Rows(DefaultTextRowNo).Cells(5).Value %></HalfPointSize>
                                         <Bold><%= DataGridView1.Rows(DefaultTextRowNo).Cells(7).Value %></Bold>
                                         <Italic><%= DataGridView1.Rows(DefaultTextRowNo).Cells(8).Value %></Italic>
                                     </DefaultText>)
                                 %>
                                 <!--RTF font settings:-->
                                 <RtfSettings>
                                     <%= From item In RtfRowNo
                                         Select
                                             <TextType>
                                                 <Name><%= DataGridView1.Rows(item).Cells(0).Value %></Name>
                                                 <FontName><%= DataGridView1.Rows(item).Cells(1).Value %></FontName>
                                                 <Color><%= DataGridView1.Rows(item).Cells(3).Value %></Color>
                                                 <HalfPointSize><%= DataGridView1.Rows(item).Cells(5).Value %></HalfPointSize>
                                                 <Bold><%= DataGridView1.Rows(item).Cells(7).Value %></Bold>
                                                 <Italic><%= DataGridView1.Rows(item).Cells(8).Value %></Italic>
                                             </TextType>
                                     %>
                                 </RtfSettings>
                             </ListOfXmlHmlRtfSettings>

        Project.SaveXmlData(SettingsFileName, xmlRtfSettings)

        DataGridView1.AllowUserToAddRows = True

    End Sub

    Private Sub btnFindSettingsFile_Click(sender As Object, e As EventArgs) Handles btnFindSettingsFile.Click
        'Fin an XML and RTF settings file:

        Select Case Project.DataLocn.Type
            Case ADVL_Utilities_Library_1.FileLocation.Types.Directory
                'Select an XML and RTF settings file from the project directory:
                OpenFileDialog1.InitialDirectory = Project.DataLocn.Path
                OpenFileDialog1.Filter = "XML HML RTF Settings | *.XmlHmlRtfSettings"
                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim DataFileName As String = System.IO.Path.GetFileName(OpenFileDialog1.FileName)
                    txtSettingsFileName.Text = DataFileName
                    Dim XmlDoc As System.Xml.Linq.XDocument
                    Project.DataLocn.ReadXmlData(DataFileName, XmlDoc)
                    ReadSettings(XmlDoc)
                End If

            Case ADVL_Utilities_Library_1.FileLocation.Types.Archive
                'Select an XML and RTF settings file from the project archive:
                'Show the zip archive file selection form:
                Zip = New ADVL_Utilities_Library_1.ZipComp
                Zip.ArchivePath = Project.DataLocn.Path
                Zip.SelectFile()
                'Zip.SelectFileForm.ApplicationName = Project.ApplicationName
                Zip.SelectFileForm.ApplicationName = Project.Application.Name
                Zip.SelectFileForm.SettingsLocn = Project.SettingsLocn
                Zip.SelectFileForm.Show()
                Zip.SelectFileForm.RestoreFormSettings()
                Zip.SelectFileForm.FileExtension = ".XmlHmlRtfSettings"
                Zip.SelectFileForm.GetFileList()
        End Select

    End Sub

    Private Sub ReadSettings(ByRef XDoc As System.Xml.Linq.XDocument)
        'Read an XML and RTF settings file:

        If IsNothing(XDoc) Then
            Exit Sub
        End If

        DataGridView1.Rows.Clear()

        'Set XML indent spaces: 
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlIndentSpaces>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XIndentSpaces = XDoc.<ListOfXmlHmlRtfSettings>.<XmlIndentSpaces>.Value
        Else
            XmlHtmDisplay1.Settings.XIndentSpaces = 4
        End If

        'Set XML large file size limit.
        '  The software will show a warning if an attemps is made to open an XML file larger than this size. (The software may have difficulty displaying very large files.)
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlLargeFileSizeLimit>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XmlLargeFileSizeLimit = XDoc.<ListOfXmlHmlRtfSettings>.<XmlLargeFileSizeLimit>.Value
        Else
            XmlHtmDisplay1.Settings.XmlLargeFileSizeLimit = 1000000
        End If

        'Set the XML Tag font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XTag.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.XTag.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XTag.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.XTag.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XTag.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.XTag.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XTag.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.XTag.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XTag.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlTag>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.XTag.Italic = False
        End If

        'Set the XML Value font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XValue.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.XValue.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XValue.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.XValue.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XValue.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.XValue.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XValue.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.XValue.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XValue.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlValue>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.XValue.Italic = False
        End If

        'Set the XML Comment font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XComment.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.XComment.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XComment.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.XComment.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XComment.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.XComment.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XComment.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.XComment.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XComment.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlComment>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.XComment.Italic = False
        End If

        'Set the XML Element font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XElement.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.XElement.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XElement.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.XElement.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XElement.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.XElement.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XElement.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.XElement.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XElement.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlElement>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.XElement.Italic = False
        End If

        'Set the XML Attribute Key font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeKey.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.XAttributeKey.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeKey.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.XAttributeKey.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeKey.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.XAttributeKey.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeKey.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.XAttributeKey.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeKey.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeKey>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.XAttributeKey.Italic = False
        End If

        'Set the XML Attribute Value font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeValue.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.XAttributeValue.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeValue.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.XAttributeValue.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeValue.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.XAttributeValue.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeValue.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.XAttributeValue.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.XAttributeValue.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<XmlSettings>.<XmlAttributeValue>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.XAttributeValue.Italic = False
        End If

        'Set the HTML Text font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HText.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.HText.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HText.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.HText.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HText.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.HText.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HText.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.HText.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HText.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlText>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.HText.Italic = False
        End If

        'Set the HTML Element font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HElement.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.HElement.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HElement.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.HElement.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HElement.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.HElement.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HElement.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.HElement.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HElement.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlElement>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.HElement.Italic = False
        End If

        'Set the HTML Attribute font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HAttribute.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.HAttribute.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HAttribute.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.HAttribute.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HAttribute.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.HAttribute.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HAttribute.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.HAttribute.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HAttribute.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlAttribute>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.HAttribute.Italic = False
        End If

        'Set the HTML Comment font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HComment.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.HComment.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HComment.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.HComment.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HComment.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.HComment.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HComment.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.HComment.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HComment.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlComment>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.HComment.Italic = False
        End If

        'Set the HTML Special Character font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HChar.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.HChar.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HChar.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.HChar.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HChar.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.HChar.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HChar.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.HChar.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HChar.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlChar>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.HChar.Italic = False
        End If

        'Set the HTML Value font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HValue.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.HValue.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HValue.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.HValue.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HValue.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.HValue.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HValue.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.HValue.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HValue.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlValue>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.HValue.Italic = False
        End If

        'Set the HTML Style font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HStyle.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.HStyle.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HStyle.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.HStyle.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HStyle.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.HStyle.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HStyle.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.HStyle.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.HStyle.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<HtmlSettings>.<HtmlStyle>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.HStyle.Italic = False
        End If

        'Set the TXT Plain Text font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.PlainText.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.PlainText.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.PlainText.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.PlainText.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.PlainText.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.PlainText.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.PlainText.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.PlainText.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.PlainText.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<TxtPlainText>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.PlainText.Italic = False
        End If

        'Set the Default Text font:
        If XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<FontName>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.DefaultText.FontName = XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<FontName>.Value
        Else
            XmlHtmDisplay1.Settings.DefaultText.FontName = "Arial"
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<Color>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.DefaultText.Color = Color.FromName(XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<Color>.Value)
        Else
            XmlHtmDisplay1.Settings.DefaultText.Color = Color.Blue
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<HalfPointSize>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.DefaultText.HalfPointSize = XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<HalfPointSize>.Value
        Else
            XmlHtmDisplay1.Settings.DefaultText.HalfPointSize = 20
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<Bold>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.DefaultText.Bold = XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<Bold>.Value
        Else
            XmlHtmDisplay1.Settings.DefaultText.Bold = False
        End If
        If XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<Italic>.Value <> Nothing Then
            XmlHtmDisplay1.Settings.DefaultText.Italic = XDoc.<ListOfXmlHmlRtfSettings>.<DefaultText>.<Italic>.Value
        Else
            XmlHtmDisplay1.Settings.DefaultText.Italic = False
        End If

        'Set other RTF text type fonts:
        XmlHtmDisplay1.Settings.ClearAllTextTypes()

        Dim TextTypes = From item In XDoc.<ListOfXmlHmlRtfSettings>.<RtfSettings>.<TextType>
        Dim TextTypeName As String
        For Each item In TextTypes
            TextTypeName = item.<Name>.Value
            XmlHtmDisplay1.Settings.AddNewTextType(TextTypeName)
            XmlHtmDisplay1.Settings.TextType(TextTypeName).FontName = item.<FontName>.Value
            XmlHtmDisplay1.Settings.TextType(TextTypeName).Color = Color.FromName(item.<Color>.Value)
            XmlHtmDisplay1.Settings.TextType(TextTypeName).HalfPointSize = item.<HalfPointSize>.Value
            XmlHtmDisplay1.Settings.TextType(TextTypeName).Bold = item.<Bold>.Value
            XmlHtmDisplay1.Settings.TextType(TextTypeName).Italic = item.<Italic>.Value
        Next

        ShowSettings()

    End Sub


    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        'Apply the changes made to the settings in DataGridView1

        Dim I As Integer 'Loop index

        Dim TextTypeName As String
        XmlHtmDisplay1.Settings.ClearAllTextTypes()

        DataGridView1.AllowUserToAddRows = False

        XmlHtmDisplay1.Settings.XIndentSpaces = txtXmlIndentSpaces.Text
        XmlHtmDisplay1.Settings.XmlLargeFileSizeLimit = txtXmlFileSizeLimit.Text

        For I = 0 To DataGridView1.RowCount - 1
            Select Case DataGridView1.Rows(I).Cells(0).Value
                Case "XML Tag"
                    XmlHtmDisplay1.Settings.XTag.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.XTag.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.XTag.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.XTag.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.XTag.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "XML Value"
                    XmlHtmDisplay1.Settings.XValue.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.XValue.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.XValue.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.XValue.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.XValue.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "XML Comment"
                    XmlHtmDisplay1.Settings.XComment.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.XComment.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.XComment.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.XComment.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.XComment.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "XML Element"
                    XmlHtmDisplay1.Settings.XElement.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.XElement.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.XElement.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.XElement.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.XElement.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "XML Attribute Key"
                    XmlHtmDisplay1.Settings.XAttributeKey.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.XAttributeKey.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.XAttributeKey.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.XAttributeKey.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.XAttributeKey.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "XML Attribute Value"
                    XmlHtmDisplay1.Settings.XAttributeValue.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.XAttributeValue.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.XAttributeValue.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.XAttributeValue.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.XAttributeValue.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "HTML Text"
                    XmlHtmDisplay1.Settings.HText.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.HText.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.HText.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.HText.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.HText.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "HTML Element"
                    XmlHtmDisplay1.Settings.HElement.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.HElement.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.HElement.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.HElement.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.HElement.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "HTML Attribute"
                    XmlHtmDisplay1.Settings.HAttribute.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.HAttribute.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.HAttribute.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.HAttribute.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.HAttribute.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "HTML Comment"
                    XmlHtmDisplay1.Settings.HComment.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.HComment.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.HComment.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.HComment.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.HComment.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "HTML Special Chars"
                    XmlHtmDisplay1.Settings.HChar.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.HChar.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.HChar.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.HChar.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.HChar.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "HTML Value"
                    XmlHtmDisplay1.Settings.HValue.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.HValue.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.HValue.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.HValue.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.HValue.Italic = DataGridView1.Rows(I).Cells(8).Value
                Case "HTML Style"
                    XmlHtmDisplay1.Settings.HStyle.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.HStyle.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.HStyle.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.HStyle.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.HStyle.Italic = DataGridView1.Rows(I).Cells(8).Value

                Case "TXT Plain Text"
                    XmlHtmDisplay1.Settings.PlainText.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.PlainText.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.PlainText.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.PlainText.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.PlainText.Italic = DataGridView1.Rows(I).Cells(8).Value

                Case "Default Text"
                    XmlHtmDisplay1.Settings.DefaultText.FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.DefaultText.Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.DefaultText.HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.DefaultText.Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.DefaultText.Italic = DataGridView1.Rows(I).Cells(8).Value

                Case Else
                    TextTypeName = DataGridView1.Rows(I).Cells(0).Value
                    XmlHtmDisplay1.Settings.AddNewTextType(TextTypeName)
                    XmlHtmDisplay1.Settings.TextType(TextTypeName).FontName = DataGridView1.Rows(I).Cells(1).Value
                    XmlHtmDisplay1.Settings.TextType(TextTypeName).Color = Color.FromName(DataGridView1.Rows(I).Cells(3).Value)
                    XmlHtmDisplay1.Settings.TextType(TextTypeName).HalfPointSize = DataGridView1.Rows(I).Cells(5).Value
                    XmlHtmDisplay1.Settings.TextType(TextTypeName).Bold = DataGridView1.Rows(I).Cells(7).Value
                    XmlHtmDisplay1.Settings.TextType(TextTypeName).Italic = DataGridView1.Rows(I).Cells(8).Value
            End Select
        Next

        XmlHtmDisplay1.Settings.UpdateColorIndexes()
        XmlHtmDisplay1.Settings.UpdateFontIndexes()

        'The font and color indexes may have changed.
        'Update the DataGridView:
        ShowSettings()

        XmlHtmDisplay2.Settings = XmlHtmDisplay1.Settings

        DataGridView1.AllowUserToAddRows = True

    End Sub

    Private Sub btnOpenSelectedFile_Click_1(sender As Object, e As EventArgs)

    End Sub


    Private Sub btnShowTextTypes_Click_1(sender As Object, e As EventArgs) Handles btnShowTextTypes.Click
        'Show the list of text types:
        Message.Add("Text Types:" & vbCrLf)

        Dim I As Integer
        For I = 0 To XmlHtmDisplay1.Settings.TextType.Count
            Message.Add(XmlHtmDisplay1.Settings.TextType.Keys(I) & vbCrLf)
        Next
    End Sub

    Private Sub EditRtf_Message(Msg As String) Handles EditRtf.Message
        Message.Add(Msg)
    End Sub

    Private Sub EditRtf_ErrorMessage(Msg As String) Handles EditRtf.ErrorMessage
        Message.AddWarning(Msg)
    End Sub

    Private Sub XmlDisplay2_ErrorMessage(Msg As String)
        Message.AddWarning(Msg)
    End Sub

    Private Sub XmlDisplay2_Message(Msg As String)
        Message.Add(Msg)
    End Sub

    Private Sub XmlDisplay1_TextChanged(sender As Object, e As EventArgs)
        DocumentTextChanged = True
    End Sub

    'Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
    '    SaveDocument()
    'End Sub

    Public Sub SaveDocument()

        Select Case FileType
            Case FileTypes.RTF
                If FileLocationType = LocationTypes.FileSystem Then 'Save the document at the specified path.
                    If RtfFileName = "" Then
                        Beep()
                        'Call the File Save dialog...
                        Message.AddWarning("The RTF file name is blank!" & vbCrLf)
                    Else
                        If XmlHtmDisplay1.SaveRtfFile(RtfFileDirectory & "\" & RtfFileName) = True Then
                            'File was saved OK.
                            LastRtfFileName = RtfFileName 'Update the LastRtfFilePath.
                            LastRtfFileLocationType = LocationTypes.FileSystem
                            LastRtfFileDirectory = RtfFileDirectory
                            DocumentTextChanged = False
                        End If
                    End If
                Else 'Save the document in the current project.
                    If RtfFileName = "" Then
                        Beep()
                        Message.AddWarning("The RTF file name is blank!" & vbCrLf)
                    Else
                        Dim rtbData As New IO.MemoryStream
                        XmlHtmDisplay1.SaveFile(rtbData, RichTextBoxStreamType.RichText)
                        rtbData.Position = 0
                        Project.SaveData(RtfFileName, rtbData)
                        LastRtfFileName = RtfFileName 'Update the LastRtfFilePath.
                        LastRtfFileLocationType = LocationTypes.Project
                        DocumentTextChanged = False
                    End If
                End If

            Case FileTypes.TXT

            Case FileTypes.XML
                If FileLocationType = LocationTypes.FileSystem Then 'Save the document at the specified path.
                    If XmlFileName = "" Then
                        Beep()
                        'Call the File Save dialog...
                        Message.AddWarning("The XML file name is blank!" & vbCrLf)
                    Else
                        If XmlHtmDisplay1.SaveXmlFile(XmlFileDirectory & "\" & XmlFileName) = True Then
                            'File was saved OK.
                            LastXmlFileName = XmlFileName 'Update the LastXmlFilePath.
                            LastXmlFileLocationType = LocationTypes.FileSystem
                            LastXmlFileDirectory = XmlFileDirectory
                            DocumentTextChanged = False
                        End If
                    End If
                Else 'Save the document in the current project.
                    If XmlFileName = "" Then
                        Beep()
                        'Call the File Save dialog...
                        Message.AddWarning("The XML file name is blank!" & vbCrLf)
                    Else
                        Dim XDoc As System.Xml.Linq.XDocument = XDocument.Parse(XmlHtmDisplay1.Text)
                        Project.SaveXmlData(XmlFileName, XDoc)
                        LastXmlFileName = XmlFileName 'Update the LastXmlFilePath.
                        LastXmlFileLocationType = LocationTypes.Project
                        DocumentTextChanged = False
                    End If
                End If

            Case FileTypes.HTML
                If FileLocationType = LocationTypes.FileSystem Then 'Save the document at the specified path.
                    If HtmlFileName = "" Then
                        Beep()
                        'Call the file save dialog...
                        Message.AddWarning("The HTML file name is blank!" & vbCrLf)
                    Else
                        XmlHtmDisplay1.SaveFile(FileDirectory & "\" & HtmlFileName, RichTextBoxStreamType.PlainText)
                        LastHtmlFileName = HtmlFileName 'Update the LastXmlFilePath.
                        LastHtmlFileLocationType = LocationTypes.Project
                        DocumentTextChanged = False
                    End If
                Else 'Save the document in the current project.
                    If HtmlFileName = "" Then
                        Beep()
                        'Call the File Save dialog...
                        Message.AddWarning("The HTML file name is blank!" & vbCrLf)
                    Else
                        Dim htmData As New IO.MemoryStream
                        XmlHtmDisplay1.SaveFile(htmData, RichTextBoxStreamType.PlainText)
                        htmData.Position = 0
                        Project.SaveData(FileName, htmData)
                        LastHtmlFileName = HtmlFileName 'Update the LastXmlFilePath.
                        LastHtmlFileLocationType = LocationTypes.Project
                        DocumentTextChanged = False
                    End If

                End If
        End Select
    End Sub

    Private Sub txtFileName2_LostFocus(sender As Object, e As EventArgs) Handles txtFileName2.LostFocus
        _fileName = txtFileName2.Text 'CHECK IF THIS IS REQUIRED!!!
    End Sub

    Private Sub XmlDisplay1_MouseDown(sender As Object, e As MouseEventArgs)
        'Record the cursor line and position id if the EditXml window is open.

        If IsNothing(EditXml) Then
            'Exit XML window is not open
        Else
            EditXml.txtCursorLine.Text = XmlHtmDisplay1.GetLineFromCharIndex(XmlHtmDisplay1.SelectionStart) + 1
            EditXml.txtCursorPosn.Text = XmlHtmDisplay1.SelectionStart - XmlHtmDisplay1.GetFirstCharIndexOfCurrentLine + 1
        End If

    End Sub

    'Private Sub cmbDocType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDocType.SelectedIndexChanged
    '    'Select Case cmbDocType.SelectedText
    '    Select Case cmbDocType.Text
    '        Case "XML"
    '            _fileTypeSelection = FileTypes.XML
    '        Case "TXT"
    '            _fileTypeSelection = FileTypes.TXT
    '        Case "RTF"
    '            _fileTypeSelection = FileTypes.RTF
    '        Case "HTML"
    '            _fileTypeSelection = FileTypes.HTML
    '    End Select
    'End Sub

    'Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

    '    If DocumentTextChanged = True Then
    '        Dim result As Integer = MessageBox.Show("Save changes to the current document?", "Notice", MessageBoxButtons.YesNoCancel)
    '        If result = DialogResult.Cancel Then
    '            Exit Sub
    '        ElseIf result = DialogResult.Yes Then
    '            SaveDocument()
    '            FileName = ""
    '            XmlHtmDisplay1.Clear()
    '            DocumentTextChanged = False
    '        ElseIf result = DialogResult.No Then
    '            'Do not save the changes!
    '            FileName = ""
    '            XmlHtmDisplay1.Clear()
    '            DocumentTextChanged = False
    '        End If
    '    Else
    '        FileName = ""
    '        XmlHtmDisplay1.Clear()
    '        DocumentTextChanged = False
    '    End If

    'End Sub


    Private Sub btnSaveAs_Click(sender As Object, e As EventArgs) Handles btnSaveAs.Click

        Select Case FileType
            Case FileTypes.RTF
                If FileLocationType = LocationTypes.FileSystem Then 'NOTE: LOOK AT THE CHECKBOX FOR THE FILE TYPE!!!
                    SaveFileDialog1.InitialDirectory = FileDirectory
                    SaveFileDialog1.FileName = FileName
                    SaveFileDialog1.Filter = "Rich Text Format (*.rtf)| *.rtf"
                    SendKeys.Send("{HOME}") 'To move the cursor to the start of the FileName
                    If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                        FileName = System.IO.Path.GetFileName(SaveFileDialog1.FileName)
                        RtfFileName = FileName
                        FileLocationType = LocationTypes.FileSystem
                        FileDirectory = System.IO.Path.GetDirectoryName(SaveFileDialog1.FileName)
                        RtfFileDirectory = FileDirectory
                        If XmlHtmDisplay1.SaveRtfFile(RtfFileDirectory & "\" & RtfFileName) = True Then
                            'File was saved OK.
                            LastRtfFileName = FileName
                            LastRtfFileLocationType = LocationTypes.FileSystem
                            LastRtfFileDirectory = FileDirectory
                            DocumentTextChanged = False
                        End If
                    End If

                Else

                End If

            Case FileTypes.TXT

            Case FileTypes.XML
                If FileLocationType = LocationTypes.FileSystem Then
                    SaveFileDialog1.InitialDirectory = FileDirectory
                    SaveFileDialog1.FileName = FileName
                    SaveFileDialog1.Filter = "All Files (*.*)| *.*"
                    SendKeys.Send("{HOME}") 'To move the cursor to the start of the FileName
                    If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                        FileName = System.IO.Path.GetFileName(SaveFileDialog1.FileName)
                        XmlFileName = FileName
                        FileLocationType = LocationTypes.FileSystem
                        FileDirectory = System.IO.Path.GetDirectoryName(SaveFileDialog1.FileName)
                        XmlFileDirectory = FileDirectory
                        If XmlHtmDisplay1.SaveXmlFile(XmlFileDirectory & "\" & XmlFileName) = True Then
                            'File was saved OK.
                            LastXmlFileName = FileName
                            LastXmlFileLocationType = LocationTypes.FileSystem
                            LastXmlFileDirectory = FileDirectory
                            DocumentTextChanged = False
                        End If
                    End If
                Else

                End If
        End Select
    End Sub

    Private Sub btnAddDocument_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnOpenInNewWindow_Click_1(sender As Object, e As EventArgs)

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
        If XmlHtmDisplay1.SelectionFont.Bold = True Then
            If XmlHtmDisplay1.SelectionFont.Italic = True Then
                XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, FontStyle.Regular + FontStyle.Italic)
            Else
                XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, FontStyle.Regular)
            End If

        ElseIf XmlHtmDisplay1.SelectionFont.Bold = False Then
            If XmlHtmDisplay1.SelectionFont.Italic = True Then
                XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, FontStyle.Bold + FontStyle.Italic)
            Else
                XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, FontStyle.Bold)
            End If
        End If
        XmlHtmDisplay1.Focus()

    End Sub

    Private Sub EditRtf_Italic() Handles EditRtf.Italic
        If XmlHtmDisplay1.SelectionFont.Italic = True Then
            If XmlHtmDisplay1.SelectionFont.Bold = True Then
                XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, FontStyle.Regular + FontStyle.Bold)
            Else
                XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, FontStyle.Regular)
            End If

        ElseIf XmlHtmDisplay1.SelectionFont.Italic = False Then
            If XmlHtmDisplay1.SelectionFont.Bold = True Then
                XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, FontStyle.Italic + FontStyle.Bold)
            Else
                XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont, FontStyle.Italic)
            End If
        End If
        XmlHtmDisplay1.Focus()

    End Sub

    Private Sub EditRtf_Underline() Handles EditRtf.Underline

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
            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont.FontFamily, XmlHtmDisplay1.SelectionFont.SizeInPoints + 1)
        Catch ex As Exception
        End Try
        XmlHtmDisplay1.Focus()
    End Sub

    Private Sub EditRtf_DecreaseFontSize() Handles EditRtf.DecreaseFontSize
        Try
            XmlHtmDisplay1.SelectionFont = New Font(XmlHtmDisplay1.SelectionFont.FontFamily, XmlHtmDisplay1.SelectionFont.SizeInPoints - 1)
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

    'Edit Rtf Code -------------------------------------------------------------------------------------

    Public Sub SaveLibrary()
        'Save the current Document Library.

        If LibraryFileName = "" Then
            Message.AddWarning("The Library File Name is blank." & vbCrLf)
            Exit Sub
        End If

        Message.Add("START: SaveLibrary()" & vbCrLf)

        Dim decl As New XDeclaration("1.0", "utf-8", "yes")
        Dim XDoc As New XDocument(decl, Nothing)
        XDoc.Add(New XComment(""))
        XDoc.Add(New XComment("Document Library Information"))

        Dim myLibrary As New XElement("Library")

        SaveLibNode(myLibrary, "", trvLibrary.Nodes)

        XDoc.Add(myLibrary)

        Project.SaveXmlData(LibraryFileName, XDoc)

    End Sub

    Private Sub SaveLibNode(ByRef myElement As XElement, Parent As String, ByRef tnc As TreeNodeCollection)
        'Save the nodes in the TreeNodeCollection in the XElement
        'This method calls itself recursively to save all nodes in trvLibrary.

        Dim I As Integer

        If tnc.Count = 0 Then 'Leaf
        Else
            For I = 0 To tnc.Count - 1
                Dim NodeKey As String = tnc(I).Name
                Dim myNode As New XElement(System.Xml.XmlConvert.EncodeName(NodeKey)) 'A space character os not allowed in an XElement name. Replace spaces with &sp characters.
                Dim myNodeText As New XElement("Text", tnc(I).Text)
                myNode.Add(myNodeText)

                If tnc(I).Nodes.Count > 0 Then
                    Message.Add("Node name = " & tnc(I).Name & " IsExpanded: " & tnc(I).IsExpanded & vbCrLf)
                End If

                If NodeKey.EndsWith(".DocLib") Then 'This the root node containing information about the Library.
                    'Note: The LibraryName property is the same as the root node Text.
                    'Note: The LibraryFileName property is the same as the root node Name (or Key).
                    Dim myLibDescr As New XElement("Description", LibraryDescription)
                    myNode.Add(myLibDescr)
                    Dim myLibCreationDate As New XElement("CreationDate", Format(LibraryCreationDate, "d-MMM-yyyy H:mm:ss"))
                    myNode.Add(myLibCreationDate)
                    Dim myLibLastEditDate As New XElement("LastEditDate", Format(LibraryLastEditDate, "d-MMM-yyyy H:mm:ss"))
                    myNode.Add(myLibLastEditDate)
                    'Add IsExpanded element: ======================================
                    Dim isExpanded As New XElement("IsExpanded", tnc(I).IsExpanded)
                    myNode.Add(isExpanded)
                Else
                    'Non-root node. Save the node properties that are stored in the ItemInfo dictionary.
                    'Description   'String
                    'CreationDate  'DateTime Format(Now, "d-MMM-yyyy H:mm:ss") 
                    'LastEditDate  'DateTime Format(Now, "d-MMM-yyyy H:mm:ss") 
                    'Type          'String
                    'Directory     'String
                    'Left          'Integer
                    'Top           'Integer
                    'Width         'Integer
                    'Height        'Integer
                    'FormNo        'Integer
                    'Add IsExpanded element? =======================================================
                    If tnc(I).Nodes.Count > 0 Then
                        Dim isExpanded As New XElement("IsExpanded", tnc(I).IsExpanded)
                        myNode.Add(isExpanded)
                    End If

                    Dim myNodeDescr As New XElement("Description", ItemInfo(NodeKey).Description)
                    myNode.Add(myNodeDescr)
                    Dim myNodeCreationDate As New XElement("CreationDate", Format(ItemInfo(NodeKey).CreationDate, "d-MMM-yyyy H:mm:ss"))
                    myNode.Add(myNodeCreationDate)
                    Dim myNodeLastEditDate As New XElement("LastEditDate", Format(ItemInfo(NodeKey).LastEditDate, "d-MMM-yyyy H:mm:ss"))
                    myNode.Add(myNodeLastEditDate)
                    Dim myNodeType As New XElement("Type", ItemInfo(NodeKey).Type)
                    myNode.Add(myNodeType)
                    Dim myNodeDirectory As New XElement("Directory", ItemInfo(NodeKey).Directory)
                    myNode.Add(myNodeDirectory)
                    Dim myNodeLeft As New XElement("Left", ItemInfo(NodeKey).Left)
                    myNode.Add(myNodeLeft)
                    Dim myNodeTop As New XElement("Top", ItemInfo(NodeKey).Top)
                    myNode.Add(myNodeTop)
                    Dim myNodeWidth As New XElement("Width", ItemInfo(NodeKey).Width)
                    myNode.Add(myNodeWidth)
                    Dim myNodeHeight As New XElement("Height", ItemInfo(NodeKey).Height)
                    myNode.Add(myNodeHeight)
                End If
                SaveLibNode(myNode, tnc(I).Name, tnc(I).Nodes)
                myElement.Add(myNode)
            Next
        End If
    End Sub

    Private Sub SaveLibNode_Old(ByRef myElement As XElement, Parent As String, ByRef tnc As TreeNodeCollection)
        'Save the nodes in the TreeNodeCollection in the XElement
        'This method calls itself recursively to save all nodes in trvLibrary.

        Dim I As Integer

        If tnc.Count = 0 Then 'Leaf
        Else
            For I = 0 To tnc.Count - 1
                Dim myNode As New XElement(System.Xml.XmlConvert.EncodeName(tnc(I).Name)) 'A space character os not allowed in an XElement name. Replace spaces with &sp characters.
                Dim myNodeText As New XElement("Text", tnc(I).Text)
                myNode.Add(myNodeText)

                If tnc(I).Name.EndsWith(".DocLib") Then 'This the root node containing information about the Library.
                    'Note: The LibraryName property is the same as the root node Text.
                    'Note: The LibraryFileName property is the same as the root node Name (or Key).
                    Dim myLibDescr As New XElement("Description", LibraryDescription)
                    myNode.Add(myLibDescr)
                    Dim myLibCreationDate As New XElement("CreationDate", Format(LibraryCreationDate, "d-MMM-yyyy H:mm:ss"))
                    myNode.Add(myLibCreationDate)
                    Dim myLibLastEditDate As New XElement("LastEditDate", Format(LibraryLastEditDate, "d-MMM-yyyy H:mm:ss"))
                    myNode.Add(myLibLastEditDate)
                End If

                SaveLibNode(myNode, tnc(I).Name, tnc(I).Nodes)
                myElement.Add(myNode)
            Next
        End If
    End Sub

    Private Sub OpenLibrary(ByVal FileName As String)
        'Open the library with the specified file name.

        If FileName = "" Then
            Message.AddWarning("Library FileName is blank." & vbCrLf)
            Exit Sub
        End If


        LibraryFileName = FileName
        Dim XDocLib As XDocument
        Project.ReadXmlData(FileName, XDocLib)
        OpenLibraryXDoc(XDocLib)

        UpdateDocList()
    End Sub

    Private Sub OpenLibraryXDoc(ByVal myXDoc As XDocument)
        'Open the Document Library stored in the XDocument

        If myXDoc Is Nothing Then
            Message.AddWarning("Library file is empty." & vbCrLf)
            Exit Sub
        End If

        trvLibrary.Nodes.Clear() 'Clear the nodes in the Library TreeView

        Dim I As Integer

        'Need to convert the XDocument to an XmlDocument:
        Dim XDoc As New System.Xml.XmlDocument
        XDoc.LoadXml(myXDoc.ToString)

        ProcessLibChildNode(XDoc.DocumentElement, trvLibrary.Nodes, "", True)

    End Sub

    Private Sub ProcessLibChildNode(ByVal xml_Node As System.Xml.XmlNode, ByVal tnc As TreeNodeCollection, ByVal Spaces As String, ByVal ParentNodeIsExpanded As Boolean)
        'Opening the .DocLib library defintion file. Process the Child Nodes.
        'This subroutine calls itself to process the child node branches.

        Dim NodeInfo As System.Xml.XmlNode
        Dim NodeText As String = ""
        Dim NodeKey As String = ""
        Dim IsExpanded As Boolean = True
        Dim HasNodes As Boolean = True

        For Each ChildNode As System.Xml.XmlNode In xml_Node.ChildNodes
            Dim myNodeText As System.Xml.XmlNode
            myNodeText = ChildNode.SelectSingleNode("Text")
            If IsNothing(myNodeText) Then
                'Message.Add("/Text node not found. " & vbCrLf)
            Else
                Dim myNodeTextValue As String = myNodeText.InnerText 'eg My Library
                If ChildNode.Name.EndsWith(".DocLib") Then
                    NodeKey = System.Xml.XmlConvert.DecodeName(ChildNode.Name)
                    If ItemInfo.ContainsKey(NodeKey) Then
                        Message.AddWarning("The Library node is already listed in the ItemInfo dictionary: " & NodeKey & vbCrLf)
                    Else
                        ItemInfo.Add(NodeKey, New clsItemInfo) 'Add the Item Name to the ItemInfo dictionary.

                        'Add the Node Text to the ItemInfo dictionary: - ADDED 16 September 2020 - Used in the Document List tab.
                        ItemInfo(NodeKey).Text = myNodeTextValue

                        'Read Library description:
                        NodeInfo = ChildNode.SelectSingleNode("Description")
                        If IsNothing(NodeInfo) Then
                            LibraryDescription = ""
                        Else
                            LibraryDescription = NodeInfo.InnerText
                        End If
                        ItemInfo(NodeKey).Description = LibraryDescription

                        'Read Library creation date:
                        NodeInfo = ChildNode.SelectSingleNode("CreationDate")
                        If NodeInfo Is Nothing Then
                            LibraryCreationDate = Now
                        Else
                            LibraryCreationDate = NodeInfo.InnerText
                        End If
                        ItemInfo(NodeKey).CreationDate = LibraryCreationDate

                        'Read Library last edit date:
                        NodeInfo = ChildNode.SelectSingleNode("LastEditDate")
                        If NodeInfo Is Nothing Then
                            LibraryLastEditDate = Now
                        Else
                            LibraryLastEditDate = NodeInfo.InnerText
                        End If
                        ItemInfo(NodeKey).LastEditDate = LibraryLastEditDate

                        'Read Library IsExpanded:
                        NodeInfo = ChildNode.SelectSingleNode("IsExpanded")
                        If NodeInfo Is Nothing Then
                            IsExpanded = True
                        Else
                            IsExpanded = NodeInfo.InnerText
                        End If

                        ItemInfo(NodeKey).Type = "Library"

                        LibraryName = myNodeTextValue
                        Dim new_Node As TreeNode = tnc.Add(NodeKey, myNodeTextValue, 0, 1) 'Add a node to the tree node collection: Key, Text, ImageIndex, SelectedImageIndex.

                        ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)

                        If IsExpanded Then
                            new_Node.Expand()
                        End If

                    End If

                Else 'ChildNode.Name may end with: .DocColl, .rtf, .xml, .html, .txt or any other extension. Use the Type element to determine the node type.
                    'Description   'String
                    'CreationDate  'DateTime Format(Now, "d-MMM-yyyy H:mm:ss") 
                    'LastEditDate  'DateTime Format(Now, "d-MMM-yyyy H:mm:ss") 
                    'Type          'String
                    'Directory     'String
                    'Left          'Integer
                    'Top           'Integer
                    'Width         'Integer
                    'Height        'Integer
                    'FormNo        'Integer

                    NodeKey = System.Xml.XmlConvert.DecodeName(ChildNode.Name)
                    If ItemInfo.ContainsKey(NodeKey) Then
                        Message.AddWarning("The following item is already listed: " & NodeKey & vbCrLf)
                    Else
                        ItemInfo.Add(NodeKey, New clsItemInfo) 'Add the Item Name to the ItemInfo dictionary.

                        'Add the Node Text to the ItemInfo dictionary: - ADDED 16 September 2020 - Used in the Document List tab.
                        ItemInfo(NodeKey).Text = myNodeTextValue

                        'Read Item description:
                        NodeInfo = ChildNode.SelectSingleNode("Description")
                        If NodeInfo Is Nothing Then
                            ItemInfo(NodeKey).Description = ""
                        Else
                            ItemInfo(NodeKey).Description = NodeInfo.InnerText
                        End If

                        'Read Item Creation Date:
                        NodeInfo = ChildNode.SelectSingleNode("CreationDate")
                        If NodeInfo Is Nothing Then
                            ItemInfo(NodeKey).CreationDate = "1-Jan-2000 12:00:00"
                        Else
                            ItemInfo(NodeKey).CreationDate = NodeInfo.InnerText
                        End If

                        'Read Item Last Edit Date:
                        NodeInfo = ChildNode.SelectSingleNode("LastEditDate")
                        If NodeInfo Is Nothing Then
                            ItemInfo(NodeKey).LastEditDate = "1-Jan-2000 12:00:00"
                        Else
                            ItemInfo(NodeKey).LastEditDate = NodeInfo.InnerText
                        End If

                        'Read Item Directory:
                        NodeInfo = ChildNode.SelectSingleNode("Directory")
                        If NodeInfo Is Nothing Then
                            ItemInfo(NodeKey).Directory = ""
                        Else
                            ItemInfo(NodeKey).Directory = NodeInfo.InnerText
                        End If

                        'Read Item display Left position:
                        NodeInfo = ChildNode.SelectSingleNode("Left")
                        If NodeInfo Is Nothing Then
                            ItemInfo(NodeKey).Left = 0
                        Else
                            ItemInfo(NodeKey).Left = NodeInfo.InnerText
                        End If

                        'Read Item display Top position:
                        NodeInfo = ChildNode.SelectSingleNode("Top")
                        If NodeInfo Is Nothing Then
                            ItemInfo(NodeKey).Top = 0
                        Else
                            ItemInfo(NodeKey).Top = NodeInfo.InnerText
                        End If

                        'Read Item display Width:
                        NodeInfo = ChildNode.SelectSingleNode("Width")
                        If NodeInfo Is Nothing Then
                            ItemInfo(NodeKey).Width = 500
                        Else
                            ItemInfo(NodeKey).Width = NodeInfo.InnerText
                        End If

                        'Read Item display Height:
                        NodeInfo = ChildNode.SelectSingleNode("Height")
                        If NodeInfo Is Nothing Then
                            ItemInfo(NodeKey).Height = 400
                        Else
                            ItemInfo(NodeKey).Height = NodeInfo.InnerText
                        End If

                        'Item FormNo is not saved. It is only used while a document is displayed.

                        'Read Library IsExpanded:
                        NodeInfo = ChildNode.SelectSingleNode("IsExpanded")
                        If NodeInfo Is Nothing Then
                            HasNodes = False
                            IsExpanded = True
                        Else
                            IsExpanded = NodeInfo.InnerText
                        End If

                        'Read Item Type:
                        NodeInfo = ChildNode.SelectSingleNode("Type")
                        If NodeInfo Is Nothing Then
                            ItemInfo(NodeKey).Type = ""
                        Else
                            ItemInfo(NodeKey).Type = NodeInfo.InnerText

                            'Display the appropriate Node on the Library Tree:
                            Select Case NodeInfo.InnerText '(Collection, RTF, XML, HTML, TXT, etc)
                                Case "Collection"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 2, 3)
                                    If IsExpanded Then
                                        new_Node.EnsureVisible()
                                    Else
                                        new_Node.Collapse()
                                    End If
                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case "RTF"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 4, 5)
                                    If HasNodes Then
                                        If IsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    Else
                                        If ParentNodeIsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    End If

                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case "XML"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 6, 7)
                                    If HasNodes Then
                                        If IsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    Else
                                        If ParentNodeIsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    End If

                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case "HTML"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 8, 9)
                                    If HasNodes Then
                                        If IsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    Else
                                        If ParentNodeIsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    End If

                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case "TXT"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 10, 11)
                                    If HasNodes Then
                                        If IsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    Else
                                        If ParentNodeIsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    End If

                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case "PDF"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 12, 13)
                                    If HasNodes Then
                                        If IsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    Else
                                        If ParentNodeIsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    End If

                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case "XLS"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 14, 15)
                                    If HasNodes Then
                                        If IsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    Else
                                        If ParentNodeIsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    End If

                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case "FolderLink"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 16, 17)
                                    If HasNodes Then
                                        If IsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    Else
                                        If ParentNodeIsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    End If

                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case "XMsg"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 18, 19)
                                    If HasNodes Then
                                        If IsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    Else
                                        If ParentNodeIsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    End If

                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case "XSeq"
                                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeTextValue, 20, 21)
                                    If HasNodes Then
                                        If IsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    Else
                                        If ParentNodeIsExpanded Then
                                            new_Node.EnsureVisible()
                                        Else
                                            new_Node.Collapse()
                                        End If
                                    End If

                                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, IsExpanded)
                                Case Else
                                    Message.AddWarning("Unknown Item type: " & NodeInfo.InnerText & vbCrLf)
                                    Message.AddWarning("A corresponding node has not been added to the Library Tree. " & vbCrLf & vbCrLf)
                            End Select
                        End If
                    End If
                End If
            End If
        Next

        'Note: The Item Name property is the same as the node Text.
        'Note: The Item FileName property is the same as the  node Name (or Key).
        '      When a tree is saved in an XML file, the FileName is stored as an element name.
        '      Beacuse an element name cannot contain spaces, it is coded using: System.Xml.XmlConvert.EncodeName(ChildNode.Name)
        '                                     The encoded name is decoded using: System.Xml.XmlConvert.DecodeName(ChildNode.Name)

    End Sub

    Private Sub ProcessLibChildNode_Old(ByVal xml_Node As System.Xml.XmlNode, ByVal tnc As TreeNodeCollection, ByVal Spaces As String)
        'Opening the .DocLib library defintion file. Process the Child Nodes.
        'This subroutine calls itself to process the child node branches.

        Dim myInfo As System.Xml.XmlNode

        For Each ChildNode As System.Xml.XmlNode In xml_Node.ChildNodes
            Dim myName As System.Xml.XmlNode
            myName = ChildNode.SelectSingleNode("Text")
            If IsNothing(myName) Then
                Message.Add("/Text node not found. " & vbCrLf)
            Else
                Dim myNodeName As String = myName.InnerText 'eg My Library
                If ChildNode.Name.EndsWith(".DocLib") Then
                    'Read Library description:
                    Dim myDescr As System.Xml.XmlNode
                    myDescr = ChildNode.SelectSingleNode("Description")
                    If IsNothing(myDescr) Then
                        LibraryDescription = ""
                    Else
                        LibraryDescription = myDescr.InnerText
                    End If

                    'Read Library creation date:
                    Dim myCreationDate As System.Xml.XmlNode
                    myCreationDate = ChildNode.SelectSingleNode("CreationDate")
                    If myCreationDate Is Nothing Then
                        LibraryCreationDate = Now
                    Else
                        LibraryCreationDate = myCreationDate.InnerText
                    End If

                    'Read Library last edit date:
                    Dim myLastEditDate As System.Xml.XmlNode
                    myLastEditDate = ChildNode.SelectSingleNode("LastEditDate")
                    If myLastEditDate Is Nothing Then
                        LibraryLastEditDate = Now
                    Else
                        LibraryLastEditDate = myLastEditDate.InnerText
                    End If

                    LibraryName = myNodeName
                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeName, 0, 1) 'Add a node to the tree node collection: Key, Text, ImageIndex, SelectedImageIndex.
                    new_Node.EnsureVisible()

                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, True)

                ElseIf ChildNode.Name.EndsWith(".DocColl") Then 'Document collection.
                    Dim new_Node As TreeNode = tnc.Add(System.Xml.XmlConvert.DecodeName(ChildNode.Name), myNodeName, 2, 3)
                    new_Node.EnsureVisible()
                    ProcessLibChildNode(ChildNode, new_Node.Nodes, Spaces, True)

                ElseIf ChildNode.Name.EndsWith(".rtf") Then 'Rich text format file.

                ElseIf ChildNode.Name.EndsWith(".xml") Then 'XML file.

                ElseIf ChildNode.Name.EndsWith(".txt") Then 'Text file.

                Else 'Unknown file type. Check if there is a Type element.
                    'Description   'String
                    'CreationDate  'DateTime Format(Now, "d-MMM-yyyy H:mm:ss") 
                    'LastEditDate  'DateTime Format(Now, "d-MMM-yyyy H:mm:ss") 
                    'Type          'String
                    'Directory     'String
                    'Left          'Integer
                    'Top           'Integer
                    'Width         'Integer
                    'Height        'Integer
                    'FormNo        'Integer

                    If ItemInfo.ContainsKey(ChildNode.Name) Then
                        Message.AddWarning("The following item is already listed: " & ChildNode.Name & vbCrLf)
                    Else
                        ItemInfo.Add(ChildNode.Name, New clsItemInfo)

                        'Read Item description:
                        myInfo = ChildNode.SelectSingleNode("Description")
                        If myInfo Is Nothing Then
                            ItemInfo(ChildNode.Name).Description = ""
                        Else
                            ItemInfo(ChildNode.Name).Description = myInfo.InnerText
                        End If
                        'Read Item Creation Date:
                        myInfo = ChildNode.SelectSingleNode("CreationDate")
                        If myInfo Is Nothing Then
                            ItemInfo(ChildNode.Name).CreationDate = "1-Jan-2000 12:00:00"
                        Else
                            ItemInfo(ChildNode.Name).CreationDate = myInfo.InnerText
                        End If
                        'Read Item Last Edit Date:
                        myInfo = ChildNode.SelectSingleNode("LastEditDate")
                        If myInfo Is Nothing Then
                            ItemInfo(ChildNode.Name).LastEditDate = "1-Jan-2000 12:00:00"
                        Else
                            ItemInfo(ChildNode.Name).LastEditDate = myInfo.InnerText
                        End If
                        'Read Item Type:
                        myInfo = ChildNode.SelectSingleNode("Type")
                        If myInfo Is Nothing Then
                            ItemInfo(ChildNode.Name).Type = ""
                        Else
                            ItemInfo(ChildNode.Name).Type = myInfo.InnerText
                        End If
                        'Read Item Directory:
                        myInfo = ChildNode.SelectSingleNode("Directory")
                        If myInfo Is Nothing Then
                            ItemInfo(ChildNode.Name).Directory = ""
                        Else
                            ItemInfo(ChildNode.Name).Directory = myInfo.InnerText
                        End If
                        'Read Item display Left position:
                        myInfo = ChildNode.SelectSingleNode("Left")
                        If myInfo Is Nothing Then
                            ItemInfo(ChildNode.Name).Left = 0
                        Else
                            ItemInfo(ChildNode.Name).Left = myInfo.InnerText
                        End If
                        'Read Item display Top position:
                        myInfo = ChildNode.SelectSingleNode("Top")
                        If myInfo Is Nothing Then
                            ItemInfo(ChildNode.Name).Top = 0
                        Else
                            ItemInfo(ChildNode.Name).Top = myInfo.InnerText
                        End If
                        'Read Item display Width:
                        myInfo = ChildNode.SelectSingleNode("Width")
                        If myInfo Is Nothing Then
                            ItemInfo(ChildNode.Name).Width = 500
                        Else
                            ItemInfo(ChildNode.Name).Width = myInfo.InnerText
                        End If
                        'Read Item display Height:
                        myInfo = ChildNode.SelectSingleNode("Height")
                        If myInfo Is Nothing Then
                            ItemInfo(ChildNode.Name).Height = 400
                        Else
                            ItemInfo(ChildNode.Name).Height = myInfo.InnerText
                        End If
                        'Item FormNo is not saved. It is only used while a document is displayed.
                    End If
                End If
            End If
        Next

        'Note: The Item Name property is the same as the node Text.
        'Note: The Item FileName property is the same as the  node Name (or Key).

    End Sub

    Private Sub btnOpenFile_Click(sender As Object, e As EventArgs) Handles btnOpenFile.Click
        OpenDocumentFile()
    End Sub

    Private Sub btnCollectionSaveAs_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnOpenLibrary_Click(sender As Object, e As EventArgs) Handles btnOpenLibrary.Click
        'Open a library.

        Dim SelectedFile As String = Project.SelectDataFile("Document Library", "DocLib")

        If SelectedFile = "" Then
            'No file selected!
        Else
            SaveLibrary() 'Save the current library
            OpenLibrary(SelectedFile)
        End If

    End Sub


    Private Sub btnAddCollection_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnSaveLibrary_Click(sender As Object, e As EventArgs) Handles btnSaveLibrary.Click
        SaveLibrary()
    End Sub

    Private Sub trvLibrary_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles trvLibrary.AfterSelect
        txtNodePath.Text = e.Node.FullPath
        txtNodeText.Text = e.Node.Text
        txtEditNodeText.Text = e.Node.Text
        txtNodeKey.Text = e.Node.Name
        txtNodeIndex.Text = e.Node.Index


        Dim FileName As String
        FileName = trvLibrary.SelectedNode.Name
        If ItemInfo.ContainsKey(FileName) Then
            txtDocCreationDate.Text = Format(ItemInfo(FileName).CreationDate, "d-MMM-yyyy H:mm:ss")
            txtDocEditDate.Text = Format(ItemInfo(FileName).LastEditDate, "d-MMM-yyyy H:mm:ss")
            If FileName.EndsWith(".FolderLink") Then 'This is a link to a folder.
                txtFilePath.Text = ""
                txtFolderPath.Text = ItemInfo(FileName).Directory
            Else
                If ItemInfo(FileName).Directory = "" Then 'This item is located in the Project.
                    txtFilePath.Text = Project.DataLocn.Path & "\" & FileName
                    txtFolderPath.Text = ""
                Else 'The item in located in the file system.
                    txtFilePath.Text = ItemInfo(FileName).Directory & "\" & FileName
                    txtFolderPath.Text = ""
                End If
            End If
        Else
            txtFilePath.Text = ""
            txtFolderPath.Text = ""
        End If

        txtNodeKey2.Text = e.Node.Name
        If ItemInfo.ContainsKey(e.Node.Name) Then
            txtNodeType.Text = ItemInfo(e.Node.Name).Type
            txtEditDescription.Text = ItemInfo(e.Node.Name).Description
            txtItemDescription.Text = ItemInfo(e.Node.Name).Description
            If TabControl3.SelectedIndex = 4 Then 'Document View tab selected.
                'Show the document in this tab:
                DocumentView()
            End If
            If ItemInfo(e.Node.Name).Type = "HTML" Then
                btnCodeView.Enabled = True
                btnWebView.Enabled = True
            Else
                btnCodeView.Enabled = False
                btnWebView.Enabled = False
            End If
        Else
            Message.AddWarning("Node key not found in ItemInfo: " & e.Node.Name & vbCrLf)
            txtNodeType.Text = ""
            txtEditDescription.Text = ""
            txtItemDescription.Text = ""
        End If

    End Sub

    Private Sub DocumentView()
        'Show the document corresponding to the selected node in the Document View tab.
        'The document will be displayed in the Document tab.

        If trvLibrary.SelectedNode Is Nothing Then
            Message.AddWarning("No document has been selected in the tree view." & vbCrLf)
        Else
            Dim FileName As String
            FileName = trvLibrary.SelectedNode.Name
            If FileName.EndsWith(".DocLib") Then
                Message.AddWarning("The Document Library node has been selected, not a document." & vbCrLf)
            ElseIf FileName.EndsWith(".DocColl") Then
                Message.AddWarning("A Document Collection node has been selected, not a document." & vbCrLf)
            Else
                'A document has been selected.
                If ItemInfo.ContainsKey(FileName) Then
                    Select Case ItemInfo(FileName).Type
                        Case "RTF"
                            FileType = FileTypes.RTF
                            WebBrowser1.Visible = False
                            AxAcroPDF1.Visible = False
                            XmlHtmDisplay2.Visible = True
                            If ItemInfo(FileName).Directory = "" Then
                                Dim rtbData As New IO.MemoryStream
                                Project.ReadData(FileName, rtbData)
                                XmlHtmDisplay2.Clear()
                                rtbData.Position = 0
                                XmlHtmDisplay2.LoadFile(rtbData, RichTextBoxStreamType.RichText)
                            Else
                                XmlHtmDisplay2.LoadFile(ItemInfo(FileName).Directory & "\" & FileName)
                            End If
                        Case "XML"
                            FileType = FileTypes.XML
                            WebBrowser1.Visible = False
                            AxAcroPDF1.Visible = False
                            XmlHtmDisplay2.Visible = True
                            If ItemInfo(FileName).Directory = "" Then
                                Dim xmlDoc As New System.Xml.XmlDocument
                                Project.ReadXmlDocData(FileName, xmlDoc)
                                XmlHtmDisplay2.Clear()
                                XmlHtmDisplay2.Rtf = XmlHtmDisplay2.XmlToRtf(xmlDoc, True)
                            Else
                                XmlHtmDisplay2.ReadXmlFile(ItemInfo(FileName).Directory & "\" & FileName, False)
                            End If
                        Case "HTML"
                            FileType = FileTypes.HTML
                            If rbCodeView.Checked = True Then 'Show HTML code view.
                                WebBrowser1.Visible = False
                                AxAcroPDF1.Visible = False
                                XmlHtmDisplay2.Visible = True
                                If ItemInfo(FileName).Directory = "" Then
                                    Dim rtbData As New IO.MemoryStream
                                    Project.ReadData(FileName, rtbData)
                                    XmlHtmDisplay2.Clear()
                                    rtbData.Position = 0
                                    XmlHtmDisplay2.LoadFile(rtbData, RichTextBoxStreamType.PlainText)
                                    If chkPlainText.Checked = True Then
                                    Else
                                        Dim htmText As String = XmlHtmDisplay2.Text
                                        XmlHtmDisplay2.Rtf = XmlHtmDisplay2.HmlToRtf(htmText)
                                    End If
                                Else
                                    XmlHtmDisplay2.LoadFile(ItemInfo(FileName).Directory & "\" & FileName, RichTextBoxStreamType.PlainText)
                                    If chkPlainText.Checked = True Then
                                    Else
                                        Dim htmText As String = XmlHtmDisplay1.Text
                                        XmlHtmDisplay2.Rtf = XmlHtmDisplay2.HmlToRtf(htmText)
                                    End If
                                End If
                            Else 'Show web view.
                                WebBrowser1.Visible = True
                                AxAcroPDF1.Visible = False
                                XmlHtmDisplay2.Visible = False
                                If ItemInfo(FileName).Directory = "" Then
                                    Dim rtbData As New IO.MemoryStream
                                    Project.ReadData(FileName, rtbData)
                                    rtbData.Position = 0
                                    Dim sr As New IO.StreamReader(rtbData)
                                    WebBrowser1.DocumentText = sr.ReadToEnd()
                                Else
                                    WebBrowser1.Navigate("file:///" & ItemInfo(FileName).Directory & "\" & FileName)
                                End If
                            End If

                        Case "TXT"
                            FileType = FileTypes.TXT
                            WebBrowser1.Visible = False
                            AxAcroPDF1.Visible = False
                            XmlHtmDisplay2.Visible = True
                            If ItemInfo(FileName).Directory = "" Then
                                Dim rtbData As New IO.MemoryStream
                                Project.ReadData(FileName, rtbData)
                                XmlHtmDisplay2.Clear()
                                rtbData.Position = 0
                                XmlHtmDisplay2.Font = XmlHtmDisplay2.PlainTextFont

                                XmlHtmDisplay2.LoadFile(rtbData, RichTextBoxStreamType.PlainText)

                            Else
                                XmlHtmDisplay2.Font = XmlHtmDisplay2.PlainTextFont
                                XmlHtmDisplay2.LoadFile(ItemInfo(FileName).Directory & "\" & FileName, RichTextBoxStreamType.PlainText)

                            End If
                        Case "PDF"
                            FileType = FileTypes.PDF
                            WebBrowser1.Visible = False
                            XmlHtmDisplay2.Visible = False
                            AxAcroPDF1.Visible = True
                            If ItemInfo(FileName).Directory = "" Then
                                If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then
                                    'Write code later. Need to select a directory to temporarily store the extracted pdf file. (Maybe the Application directory?)
                                Else
                                    AxAcroPDF1.LoadFile(Project.DataDirLocn.Path & "\" & FileName)
                                    AxAcroPDF1.Focus()
                                End If
                                'OLD CODE:
                                'If Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then

                                'Else
                                '    Dim DataLocnPath As String = Project.DataLocn.Path
                                '    AxAcroPDF1.LoadFile(DataLocnPath & "\" & FileName)
                                '    AxAcroPDF1.Focus()
                                'End If
                            Else
                                AxAcroPDF1.LoadFile(ItemInfo(FileName).Directory & "\" & FileName)
                                AxAcroPDF1.Focus()
                            End If
                        Case "XMsg"
                            FileType = FileTypes.XML
                            WebBrowser1.Visible = False
                            AxAcroPDF1.Visible = False
                            XmlHtmDisplay2.Visible = True
                            If ItemInfo(FileName).Directory = "" Then
                                Dim xmlDoc As New System.Xml.XmlDocument
                                Project.ReadXmlDocData(FileName, xmlDoc)
                                XmlHtmDisplay2.Clear()
                                XmlHtmDisplay2.Rtf = XmlHtmDisplay2.XmlToRtf(xmlDoc, True)
                            Else
                                XmlHtmDisplay2.ReadXmlFile(ItemInfo(FileName).Directory & "\" & FileName, False)
                            End If

                        Case "XSeq"
                            FileType = FileTypes.XML
                            WebBrowser1.Visible = False
                            AxAcroPDF1.Visible = False
                            XmlHtmDisplay2.Visible = True
                            If ItemInfo(FileName).Directory = "" Then
                                Dim xmlDoc As New System.Xml.XmlDocument
                                Project.ReadXmlDocData(FileName, xmlDoc)
                                XmlHtmDisplay2.Clear()
                                XmlHtmDisplay2.Rtf = XmlHtmDisplay2.XmlToRtf(xmlDoc, True)
                            Else
                                XmlHtmDisplay2.ReadXmlFile(ItemInfo(FileName).Directory & "\" & FileName, False)
                            End If

                    End Select
                Else
                    Message.AddWarning("No information is available on the selected node :" & FileName & vbCrLf)
                End If
            End If
        End If
    End Sub

    Private Sub btnDeleteNode_Click(sender As Object, e As EventArgs) Handles btnDeleteNode.Click
        'Deleted the selected node.

        If trvLibrary.SelectedNode Is Nothing Then
            'No node has been selected.
        Else
            Dim Node As TreeNode
            Node = trvLibrary.SelectedNode
            If Node.Nodes.Count > 0 Then
                Message.AddWarning("The selected node has child nodes. Delete the child nodes before deleting this node." & vbCrLf)
            Else
                Dim Parent As TreeNode = Node.Parent
                If Node.Parent Is Nothing Then
                    Node.Remove()
                Else
                    Parent.Nodes.RemoveAt(Node.Index)
                End If

                'Delete the corresponding file:
                Dim FileName As String = Node.Name
                If ItemInfo.ContainsKey(FileName) Then
                    Dim ItemType As String = ItemInfo(FileName).Type
                    If ItemInfo(FileName).Directory = "" Then
                        'The file is located in the Project.
                        Project.DeleteData(FileName)
                    Else
                        'The file is located in the file system. DO NOT DELETE
                    End If
                    'Delete the ItemInfo entry:
                    ItemInfo.Remove(FileName)
                Else
                    'No information is available on the selected node.
                End If
            End If
        End If
    End Sub

    Private Sub btnMoveNodeUp_Click(sender As Object, e As EventArgs) Handles btnMoveNodeUp.Click
        'Move the selected item up in the Document Tree.

        If trvLibrary.SelectedNode Is Nothing Then
            'No node has been selected.
        Else
            Dim Node As TreeNode
            Node = trvLibrary.SelectedNode
            Dim index As Integer = Node.Index
            If index = 0 Then
                'Already at the first node.
                Node.TreeView.Focus()
            Else
                Dim Parent As TreeNode = Node.Parent
                Parent.Nodes.RemoveAt(index)
                Parent.Nodes.Insert(index - 1, Node)
                trvLibrary.SelectedNode = Node
                Node.TreeView.Focus()
            End If
        End If
    End Sub

    Private Sub btnMoveNodeDown_Click(sender As Object, e As EventArgs) Handles btnMoveNodeDown.Click
        'Move the selected item down in the Document Tree.

        If trvLibrary.SelectedNode Is Nothing Then
            'No node has been selected.
        Else
            Dim Node As TreeNode
            Node = trvLibrary.SelectedNode
            Dim index As Integer = Node.Index
            Dim Parent As TreeNode = Node.Parent
            If index < Parent.Nodes.Count - 1 Then
                Parent.Nodes.RemoveAt(index)
                Parent.Nodes.Insert(index + 1, Node)
                trvLibrary.SelectedNode = Node
                Node.TreeView.Focus()
            Else
                'Already at the last node.
                Node.TreeView.Focus()
            End If
        End If
    End Sub

    Private Sub rbCollection_CheckedChanged(sender As Object, e As EventArgs) Handles rbCollection.CheckedChanged
        If rbCollection.Checked = True Then
            _newItemType = NewItemTypes.Collection
            txtAddItemNotes.Text = "Enter the settings to create a new Collection."
        End If
    End Sub

    Private Sub rbRtf_CheckedChanged(sender As Object, e As EventArgs) Handles rbRtf.CheckedChanged
        If rbRtf.Checked = True Then
            _newItemType = NewItemTypes.RTF
            txtAddItemNotes.Text = "Enter the settings to create a new RTF file."
        End If
    End Sub

    Private Sub rbXml_CheckedChanged(sender As Object, e As EventArgs) Handles rbXml.CheckedChanged
        If rbXml.Checked = True Then
            _newItemType = NewItemTypes.XML
            txtAddItemNotes.Text = "Enter the settings to create a new XML file."
        End If
    End Sub

    Private Sub rbHtml_CheckedChanged(sender As Object, e As EventArgs) Handles rbHtml.CheckedChanged
        If rbHtml.Checked = True Then
            _newItemType = NewItemTypes.HTML
            txtAddItemNotes.Text = "Enter the settings to create a new Html file."
        End If
    End Sub

    Private Sub rbTxt_CheckedChanged(sender As Object, e As EventArgs) Handles rbTxt.CheckedChanged
        If rbTxt.Checked Then
            _newItemType = NewItemTypes.TXT
            txtAddItemNotes.Text = "Enter the settings to create a new Text file."
        End If
    End Sub

    Private Sub rbPdf_CheckedChanged(sender As Object, e As EventArgs) Handles rbPdf.CheckedChanged
        If rbPdf.Checked Then
            _newItemType = NewItemTypes.PDF
            txtAddItemNotes.Text = "Drag a PDF file onto a node in the tree view. The PDF file will also be copied into the project data location."
        End If
    End Sub

    Private Sub rbFolderLink_CheckedChanged(sender As Object, e As EventArgs) Handles rbFolderLink.CheckedChanged
        If rbFolderLink.Checked Then
            _newItemType = NewItemTypes.FolderLink
            txtAddItemNotes.Text = "Enter the folder link settings then press the New button."
        End If
    End Sub

    Private Sub rbXls_CheckedChanged(sender As Object, e As EventArgs) Handles rbXls.CheckedChanged
        If rbXls.Checked Then
            _newItemType = NewItemTypes.XLS
            txtAddItemNotes.Text = "Drag an Excel spreadsheet file onto a node in the tree view. The Excel file will also be copied into the project data location."
        End If
    End Sub

    Private Sub rbXMsg_CheckedChanged(sender As Object, e As EventArgs) Handles rbXMsg.CheckedChanged
        If rbXMsg.Checked Then
            _newItemType = NewItemTypes.XMsg
            txtAddItemNotes.Text = "Drag an XMessage file onto a node in the tree view. The XMessage file will also be copied into the project data location."
        End If
    End Sub

    Private Sub rbXSeq_CheckedChanged(sender As Object, e As EventArgs) Handles rbXSeq.CheckedChanged
        If rbXSeq.Checked Then
            _newItemType = NewItemTypes.XSeq
            txtAddItemNotes.Text = "Drag an XSequence spreadsheet file onto a node in the tree view. The XSequence file will also be copied into the project data location."
        End If
    End Sub

    Private Sub rbDefaultWindow_CheckedChanged(sender As Object, e As EventArgs) Handles rbDefaultWindow.CheckedChanged
        If rbDefaultWindow.Checked Then
            _newItemDisplay = NewItemDisplayTypes.DefaultDisplay
        End If
    End Sub

    Private Sub rbNewWindow_CheckedChanged(sender As Object, e As EventArgs) Handles rbNewWindow.CheckedChanged
        If rbNewWindow.Checked Then
            _newItemDisplay = NewItemDisplayTypes.NewWindow
        End If
    End Sub

    Private Sub rbNoDisplay_CheckedChanged(sender As Object, e As EventArgs) Handles rbNoDisplay.CheckedChanged
        If rbNoDisplay.Checked Then
            _newItemDisplay = NewItemDisplayTypes.NoDisplay
        End If
    End Sub

    Private Sub pbIconCollection_Click(sender As Object, e As EventArgs) Handles pbIconCollection.Click
        rbCollection.Checked = True
    End Sub

    Private Sub pbIconRtf_Click(sender As Object, e As EventArgs) Handles pbIconRtf.Click
        rbRtf.Checked = True
    End Sub

    Private Sub pbIconXml_Click(sender As Object, e As EventArgs) Handles pbIconXml.Click
        rbXml.Checked = True
    End Sub

    Private Sub pbIconHtml_Click(sender As Object, e As EventArgs) Handles pbIconHtml.Click
        rbHtml.Checked = True
    End Sub

    Private Sub pbIconTxt_Click(sender As Object, e As EventArgs) Handles pbIconTxt.Click
        rbTxt.Checked = True
    End Sub

    Private Sub pbIconPdf_Click(sender As Object, e As EventArgs) Handles pbIconPdf.Click
        rbPdf.Checked = True
    End Sub

    Private Sub pbIconFolder_Click(sender As Object, e As EventArgs) Handles pbIconFolder.Click
        rbFolderLink.Checked = True
    End Sub

    Private Sub pbIconXls_Click(sender As Object, e As EventArgs) Handles pbIconXls.Click
        rbXls.Checked = True
    End Sub

    Private Sub pbIconXmsg_Click(sender As Object, e As EventArgs) Handles pbIconXmsg.Click
        rbXMsg.Checked = True
    End Sub

    Private Sub pbIconXSeq_Click(sender As Object, e As EventArgs) Handles pbIconXSeq.Click
        rbXSeq.Checked = True
    End Sub

    Private Sub btnNewItem_Click(sender As Object, e As EventArgs) Handles btnCreateItem.Click
        'Add new item to the Document Library

        If LibraryName = "" Then
            Message.AddWarning("There is no library open." & vbCrLf)
            Message.AddWarning("Please open or create a new library." & vbCrLf)
            Exit Sub
        End If

        Dim NodeKey As String = Trim(txtNewNodeFileName.Text) '(Node Key) The name of the file. (Stored in the current Project.)
        If NodeKey = "" Then
            Message.AddWarning("The file name is blank!" & vbCrLf)
            Exit Sub
        End If

        Select Case NewItemType
            Case NewItemTypes.Collection '==============================================================================================================================
                'Add the current collection to the library.

                If ItemInfo.ContainsKey(NodeKey) Then
                    Message.AddWarning("The name of the new Collection is already used: " & NodeKey & vbCrLf)
                Else
                    ItemInfo.Add(NodeKey, New clsItemInfo)
                    'Dim NodeText As String = System.IO.Path.GetFileNameWithoutExtension(NodeKey) 'The text that appears on the node in the treeview.
                    Dim NodeText As String = Trim(txtItemText.Text) 'The text that appears on the node in the treeview.

                    If trvLibrary.SelectedNode Is Nothing Then
                        trvLibrary.Nodes.Add(NodeKey, NodeText, 2, 3) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                    Else
                        trvLibrary.SelectedNode.Nodes.Add(NodeKey, NodeText, 2, 3) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                    End If

                    'Add the Node Text to the ItemInfo dictionary: - ADDED 16 September 2020 - Used in the Document List tab.
                    ItemInfo(NodeKey).Text = NodeText

                    ItemInfo(NodeKey).Type = "Collection"
                    ItemInfo(NodeKey).Description = txtNewNodeDescr.Text
                    ItemInfo(NodeKey).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    ItemInfo(NodeKey).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    ItemInfo(NodeKey).Directory = ""
                End If

            Case NewItemTypes.HTML '==============================================================================================================================
                If NodeKey.EndsWith(".html") Then
                    'Correct extension.
                Else
                    'Check if extension is a variant of the HTML extension.
                    If LCase(NodeKey).EndsWith(".html") Then
                        'Replace last four characters with "html"
                        NodeKey = Strings.Left(NodeKey, Len(NodeKey) - 4) & "html"
                    ElseIf LCase(NodeKey).EndsWith(".htm") Then
                        'Replace last three characters with "html"
                        NodeKey = Strings.Left(NodeKey, Len(NodeKey) - 3) & "html"
                    Else
                        NodeKey = NodeKey & ".html"
                        txtNewNodeFileName.Text = NodeKey
                    End If
                End If

                If ItemInfo.ContainsKey(NodeKey) Then
                    Message.AddWarning("The name of the new HTML file is already used: " & NodeKey & vbCrLf)
                Else
                    ItemInfo.Add(NodeKey, New clsItemInfo)
                    Dim NodeText As String = Trim(txtItemText.Text) 'The text that appears on the node in the treeview.

                    If trvLibrary.SelectedNode Is Nothing Then
                        trvLibrary.Nodes.Add(NodeKey, NodeText, 8, 9) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                    Else
                        trvLibrary.SelectedNode.Nodes.Add(NodeKey, NodeText, 8, 9) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                    End If

                    'Add the Node Text to the ItemInfo dictionary: - ADDED 16 September 2020 - Used in the Document List tab.
                    ItemInfo(NodeKey).Text = NodeText

                    ItemInfo(NodeKey).Type = "HTML"
                    ItemInfo(NodeKey).Description = txtNewNodeDescr.Text
                    ItemInfo(NodeKey).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    ItemInfo(NodeKey).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    ItemInfo(NodeKey).Directory = ""

                    Dim HtmlString = DefaultHtmlString(NodeKey) 'Generate the default web page code.

                    Dim htmData As New IO.MemoryStream
                    Dim sw As New IO.StreamWriter(htmData)
                    sw.Write(HtmlString)
                    sw.Flush()

                    If FileLocationType = LocationTypes.FileSystem Then
                        'NOT SURE IF A NEW FILE WILL EVER BE WRITTEN TO THE FILE SYSTEM!
                        Dim fileStream As New IO.FileStream(FileDirectory & "\" & NodeKey, IO.FileMode.Create) 'This creates a new file or completely overwrites an existing file.
                        Dim streamReader As New IO.BinaryReader(htmData)
                        fileStream.Write(streamReader.ReadBytes(htmData.Length), 0, htmData.Length)
                    Else
                        Project.SaveData(NodeKey, htmData)
                    End If

                    If rbDefaultWindow.Checked = True Then 'Show the new HTML document in the Document tab on the Main form.
                        If DocumentTextChanged = True Then
                            Dim Result As Integer = MessageBox.Show("Save changes to the current HTML document?", "Notice", MessageBoxButtons.YesNoCancel)
                            If Result = DialogResult.Cancel Then
                                Exit Sub
                            ElseIf Result = DialogResult.Yes Then
                                SaveDocument()
                            ElseIf Result = DialogResult.No Then
                                'Do not save the changes.
                            End If
                        End If
                        'Clear the default view and create the blank HTML file:
                        XmlHtmDisplay1.Clear()
                        FileName = NodeKey
                        HtmlFileName = NodeKey
                        FileType = FileTypes.HTML
                        DocumentTextChanged = False
                        SaveDocument()
                        TabControl1.SelectedIndex = 1 'Open the document tab.
                    ElseIf rbNewWindow.Checked = True Then 'Show the new HTML document in a new window.
                        'Display the new HTML file in a new HTML Code window
                        Dim FormNo As Integer = NewHtmlDisplay()
                        HtmlDisplayFormList(FormNo).FileName = NodeKey
                        HtmlDisplayFormList(FormNo).Description = ItemInfo(NodeKey).Description
                        HtmlDisplayFormList(FormNo).FileDirectory = ItemInfo(NodeKey).Directory
                        If HtmlDisplayFormList(FormNo).FileDirectory = "" Then
                            HtmlDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                        Else
                            HtmlDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                        End If

                        If chkPlainText.Checked = True Then
                            HtmlDisplayFormList(FormNo).OpenDocumentInPlainText
                        Else
                            HtmlDisplayFormList(FormNo).OpenDocument
                        End If
                    Else
                        'Dont display the new HTML document.
                    End If
                End If

            Case NewItemTypes.RTF '==============================================================================================================================
                If NodeKey.EndsWith(".rtf") Then
                    'Correct extension.
                Else
                    'Check if extension is a variant of the rtf extension.
                    If LCase(NodeKey).EndsWith(".rtf") Then
                        'Replace last three characters with "rtf"
                        NodeKey = Strings.Left(NodeKey, Len(NodeKey) - 3) & "rtf"
                    Else
                        NodeKey = NodeKey & ".rtf"
                        txtNewNodeFileName.Text = NodeKey
                    End If
                End If

                If ItemInfo.ContainsKey(NodeKey) Then
                    Message.AddWarning("The name of the new rtf file is already used: " & NodeKey & vbCrLf)
                Else
                    ItemInfo.Add(NodeKey, New clsItemInfo)
                    Dim NodeText As String = Trim(txtItemText.Text) 'The text that appears on the node in the treeview.

                    If trvLibrary.SelectedNode Is Nothing Then
                        trvLibrary.Nodes.Add(NodeKey, NodeText, 4, 5) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                    Else
                        trvLibrary.SelectedNode.Nodes.Add(NodeKey, NodeText, 4, 5) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                    End If

                    'Add the Node Text to the ItemInfo dictionary: - ADDED 16 September 2020 - Used in the Document List tab.
                    ItemInfo(NodeKey).Text = NodeText

                    ItemInfo(NodeKey).Type = "RTF"
                    ItemInfo(NodeKey).Description = txtNewNodeDescr.Text
                    ItemInfo(NodeKey).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    ItemInfo(NodeKey).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    ItemInfo(NodeKey).Directory = ""
                    If rbDefaultWindow.Checked = True Then
                        If DocumentTextChanged = True Then
                            Dim Result As Integer = MessageBox.Show("Save changes to the current RTF document?", "Notice", MessageBoxButtons.YesNoCancel)
                            If Result = DialogResult.Cancel Then
                                Exit Sub
                            ElseIf Result = DialogResult.Yes Then
                                SaveDocument()
                            ElseIf Result = DialogResult.No Then
                                'Do not save the changes.
                            End If
                        End If
                        'Clear the default view and create the blank RTF file:
                        XmlHtmDisplay1.Clear()
                        FileName = NodeKey
                        RtfFileName = NodeKey
                        FileType = FileTypes.RTF
                        DocumentTextChanged = False
                        SaveDocument()
                        TabControl1.SelectedIndex = 1 'Open the document tab.
                    ElseIf rbNewWindow.Checked = True Then   'Display the new RTF file in a new RTF Code window
                        Dim FormNo As Integer = NewRtfDisplay()
                        RtfDisplayFormList(FormNo).FileName = NodeKey
                        RtfDisplayFormList(FormNo).Description = ItemInfo(NodeKey).Description
                        RtfDisplayFormList(FormNo).FileDirectory = ItemInfo(NodeKey).Directory
                        If RtfDisplayFormList(FormNo).FileDirectory = "" Then
                            RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                        Else
                            RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                        End If

                        RtfDisplayFormList(FormNo).SaveDocument 'This will create a blank RTF file.

                    Else
                        'Dont display the new RTF document.
                        Dim rtfString As String
                        rtfString = XmlHtmDisplay1.DefaultTextToRtf("") 'Create a blank rtf string using the default font settings.
                        Message.Add("rtfString = " & vbCrLf & rtfString & vbCrLf)
                        Dim rtfData As New IO.MemoryStream
                        Dim sw As New IO.StreamWriter(rtfData)
                        sw.Write(rtfString)
                        sw.Flush()

                        If FileLocationType = LocationTypes.FileSystem Then
                            'NOT SURE IF A NEW FILE WILL EVER BE WRITTEN TO THE FILE SYSTEM!
                            Dim fileStream As New IO.FileStream(FileDirectory & "\" & NodeKey, IO.FileMode.Create) 'This creates a new file or completely overwrites an existing file.
                            Dim streamReader As New IO.BinaryReader(rtfData)
                            fileStream.Write(streamReader.ReadBytes(rtfData.Length), 0, rtfData.Length)
                        Else
                            Project.SaveData(NodeKey, rtfData)
                        End If
                    End If
                End If

            Case NewItemTypes.TXT '==============================================================================================================================
                Message.AddWarning("New Text file code not yet written!" & vbCrLf)

            Case NewItemTypes.XML '==============================================================================================================================
                Message.AddWarning("New XML file code not yet written!" & vbCrLf)

            Case NewItemTypes.PDF  '==============================================================================================================================
                Message.AddWarning("New PDF file code not yet written! Currently you can only drag an existing PDF file onto a node." & vbCrLf)

            Case NewItemTypes.XLS  '==============================================================================================================================
                Message.AddWarning("New XLS file code not yet written! Currently you can only drag an existing Excel file onto a node." & vbCrLf)

            Case NewItemTypes.FolderLink  '==============================================================================================================================
                If NodeKey.EndsWith(".FolderLink") Then
                    'Correct extension.
                Else
                    'Check if extension is a variant of the rtf extension.
                    If LCase(NodeKey).EndsWith(".folderlink") Then
                        'Replace last ten characters with "FolderLink"
                        NodeKey = Strings.Left(NodeKey, Len(NodeKey) - 10) & "FolderLink"
                    Else
                        NodeKey = NodeKey & ".FolderLink"
                        txtNewNodeFileName.Text = NodeKey
                    End If
                End If

                If ItemInfo.ContainsKey(NodeKey) Then
                    Message.AddWarning("The name of the new rtf file is already used: " & NodeKey & vbCrLf)
                Else
                    'Check if the Folder Link in txtFolderLink is valid:
                    If System.IO.Directory.Exists(txtFolderLink.Text) Then
                        ItemInfo.Add(NodeKey, New clsItemInfo)
                        Dim NodeText As String = Trim(txtItemText.Text) 'The text that appears on the node in the treeview.
                        If trvLibrary.SelectedNode Is Nothing Then
                            'Image 16 is a closed folder, image 17 is an open folder.
                            trvLibrary.Nodes.Add(NodeKey, NodeText, 16, 17) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        Else
                            'Image 16 is a closed folder, image 17 is an open folder.
                            trvLibrary.SelectedNode.Nodes.Add(NodeKey, NodeText, 16, 17) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If

                        'Add the Node Text to the ItemInfo dictionary: - ADDED 16 September 2020 - Used in the Document List tab.
                        ItemInfo(NodeKey).Text = NodeText

                        ItemInfo(NodeKey).Type = "FolderLink"
                        ItemInfo(NodeKey).Description = txtNewNodeDescr.Text
                        ItemInfo(NodeKey).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NodeKey).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NodeKey).Directory = txtFolderLink.Text
                    Else
                        Message.AddWarning("The folder link is not valid: " & txtFolderLink.Text & vbCrLf)
                    End If

                End If
        End Select
    End Sub

    Private Function DefaultHtmlString(ByVal DocumentTitle As String) As String
        'Create a blank HTML Web Page.

        Dim sb As New System.Text.StringBuilder

        sb.Append("<!DOCTYPE html>" & vbCrLf)
        sb.Append("<html>" & vbCrLf)
        sb.Append("<!-- Andorville(TM) Workflow File -->" & vbCrLf)
        sb.Append("<!-- Application Name:    " & ApplicationInfo.Name & " -->" & vbCrLf)
        sb.Append("<!-- Application Version: " & My.Application.Info.Version.ToString & " -->" & vbCrLf)
        sb.Append("<!-- Creation Date:          " & Format(Now, "dd MMMM yyyy") & " -->" & vbCrLf)
        sb.Append("<head>" & vbCrLf)
        sb.Append("<title>" & DocumentTitle & "</title>" & vbCrLf)
        sb.Append("<meta name=""description"" content=""Workflow description."">" & vbCrLf)
        sb.Append("</head>" & vbCrLf)

        sb.Append("<body style=""font-family:arial;"">" & vbCrLf & vbCrLf)

        sb.Append("<h2>" & DocumentTitle & "</h2>" & vbCrLf & vbCrLf)

        sb.Append(DefaultJavaScriptString)

        sb.Append("</body>" & vbCrLf)
        sb.Append("</html>" & vbCrLf)

        Return sb.ToString

    End Function

    Private Sub btnFindItem_Click(sender As Object, e As EventArgs) Handles btnFindItem.Click
        'Find a new item to add the Document Library

        Select Case NewItemType
            Case NewItemTypes.Collection

            Case NewItemTypes.HTML
                'Look for a HTML file in the current project.
                Dim SelectedFile As String = Project.SelectDataFile("HTML format", "html")
                If SelectedFile = "" Then
                    'No document selected.
                Else
                    If ItemInfo.ContainsKey(SelectedFile) Then
                        Message.AddWarning("The Document Library already contains the selected document: " & SelectedFile & vbCrLf)
                    Else
                        ItemInfo.Add(SelectedFile, New clsItemInfo)
                        ItemInfo(SelectedFile).Directory = "" 'Files located in the current project have a blank Directory value.
                        ItemInfo(SelectedFile).Description = txtNewNodeDescr.Text
                        ItemInfo(SelectedFile).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss") 'UPDATE THIS LATER TO OBTAIN THE FILE CREATION DATE.
                        ItemInfo(SelectedFile).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss") 'UPDATE THIS LATER TO OBTAIN THE FILE CREATION DATE.
                        ItemInfo(SelectedFile).Type = "HTML"
                        If trvLibrary.SelectedNode Is Nothing Then
                            trvLibrary.Nodes.Add(SelectedFile, Trim(txtItemText.Text), 8, 9)
                        Else
                            trvLibrary.SelectedNode.Nodes.Add(SelectedFile, Trim(txtItemText.Text), 8, 9) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If
                    End If
                End If

            Case NewItemTypes.RTF
                'Look for a RTF file in the current project.
                Dim SelectedFile As String = Project.SelectDataFile("Rich text format", "rtf")
                If SelectedFile = "" Then
                    'No document selected.
                Else
                    If ItemInfo.ContainsKey(SelectedFile) Then
                        Message.AddWarning("The Document Library already contains the selected document: " & SelectedFile & vbCrLf)
                    Else
                        ItemInfo.Add(SelectedFile, New clsItemInfo)
                        ItemInfo(SelectedFile).Directory = "" 'Files located in the current project have a blank Directory value.
                        ItemInfo(SelectedFile).Description = txtNewNodeDescr.Text
                        ItemInfo(SelectedFile).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss") 'UPDATE THIS LATER TO OBTAIN THE FILE CREATION DATE.
                        ItemInfo(SelectedFile).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss") 'UPDATE THIS LATER TO OBTAIN THE FILE CREATION DATE.
                        ItemInfo(SelectedFile).Type = "RTF"
                        If trvLibrary.SelectedNode Is Nothing Then
                            trvLibrary.Nodes.Add(SelectedFile, Trim(txtItemText.Text), 4, 5)
                        Else
                            trvLibrary.SelectedNode.Nodes.Add(SelectedFile, Trim(txtItemText.Text), 4, 5) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If
                    End If
                End If

            Case NewItemTypes.TXT
                'Look for a Text file in the current project.
                Dim SelectedFile As String = Project.SelectDataFile("Text format", "txt")
                If SelectedFile = "" Then
                    'No document selected.
                Else
                    If ItemInfo.ContainsKey(SelectedFile) Then
                        Message.AddWarning("The Document Library already contains the selected document: " & SelectedFile & vbCrLf)
                    Else
                        ItemInfo.Add(SelectedFile, New clsItemInfo)
                        ItemInfo(SelectedFile).Directory = "" 'Files located in the current project have a blank Directory value.
                        ItemInfo(SelectedFile).Description = txtNewNodeDescr.Text
                        ItemInfo(SelectedFile).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss") 'UPDATE THIS LATER TO OBTAIN THE FILE CREATION DATE.
                        ItemInfo(SelectedFile).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss") 'UPDATE THIS LATER TO OBTAIN THE FILE CREATION DATE.
                        ItemInfo(SelectedFile).Type = "TXT"
                        If trvLibrary.SelectedNode Is Nothing Then
                            trvLibrary.Nodes.Add(SelectedFile, Trim(txtItemText.Text), 10, 11)
                        Else
                            trvLibrary.SelectedNode.Nodes.Add(SelectedFile, Trim(txtItemText.Text), 10, 11) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If
                    End If
                End If

            Case NewItemTypes.XML
                'Look for an XML file in the current project.
                Dim SelectedFile As String = Project.SelectDataFile("XML file", "*")
                If SelectedFile = "" Then
                    'No document selected.
                Else
                    If ItemInfo.ContainsKey(SelectedFile) Then
                        Message.AddWarning("The Document Library already contains the selected document: " & SelectedFile & vbCrLf)
                    Else
                        ItemInfo.Add(SelectedFile, New clsItemInfo)
                        ItemInfo(SelectedFile).Directory = "" 'Files located in the current project have a blank Directory value.
                        ItemInfo(SelectedFile).Description = txtNewNodeDescr.Text
                        ItemInfo(SelectedFile).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss") 'UPDATE THIS LATER TO OBTAIN THE FILE CREATION DATE.
                        ItemInfo(SelectedFile).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss") 'UPDATE THIS LATER TO OBTAIN THE FILE CREATION DATE.
                        ItemInfo(SelectedFile).Type = "XML"
                        If trvLibrary.SelectedNode Is Nothing Then
                            trvLibrary.Nodes.Add(SelectedFile, Trim(txtItemText.Text), 6, 7)
                        Else
                            trvLibrary.SelectedNode.Nodes.Add(SelectedFile, Trim(txtItemText.Text), 6, 7) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If

                    End If
                End If

        End Select

    End Sub

    Private Sub btnApplyNodeEdits_Click(sender As Object, e As EventArgs) Handles btnApplyNodeEdits.Click

        If trvLibrary.SelectedNode Is Nothing Then

        Else
            trvLibrary.SelectedNode.Text = Trim(txtEditNodeText.Text)
            txtNodeText.Text = Trim(txtEditNodeText.Text)
            If ItemInfo.ContainsKey(trvLibrary.SelectedNode.Name) Then
                ItemInfo(trvLibrary.SelectedNode.Name).Description = txtEditDescription.Text

            End If

        End If

    End Sub

    Private Sub btnOpenDocument_Click(sender As Object, e As EventArgs) Handles btnOpenDocument.Click
        'Open the document corresponding to the selected node.
        'The document will be displayed in the Document tab.

        If trvLibrary.SelectedNode Is Nothing Then
            Message.AddWarning("No document has been selected in the tree view." & vbCrLf)
        Else
            Dim FileName As String
            FileName = trvLibrary.SelectedNode.Name
            If FileName.EndsWith(".DocLib") Then
                Message.AddWarning("The Document Library node has been selected, not a document." & vbCrLf)
            ElseIf FileName.EndsWith(".DocColl") Then
                Message.AddWarning("A Document Collection node has been selected, not a document." & vbCrLf)
            Else
                'A document has been selected.
                If ItemInfo.ContainsKey(FileName) Then
                    Select Case ItemInfo(FileName).Type
                        Case "RTF"
                            'cmbDocType.SelectedIndex = cmbDocType.FindStringExact("RTF")
                            FileType = FileTypes.RTF 'Setting this property also displays the appropriate text in txtDocType
                            RtfFileName = FileName
                            txtFileName2.Text = FileName
                            txtFileType.Text = "RTF"
                            txtFileDescription.Text = ItemInfo(FileName).Description
                            If ItemInfo(FileName).Directory = "" Then
                                RtfFileLocationType = LocationTypes.Project
                                RtfFileDirectory = ""
                                Dim rtbData As New IO.MemoryStream
                                Project.ReadData(RtfFileName, rtbData)
                                XmlHtmDisplay1.Clear()
                                rtbData.Position = 0
                                XmlHtmDisplay1.LoadFile(rtbData, RichTextBoxStreamType.RichText)
                                DocumentTextChanged = False
                                TabControl1.SelectedIndex = 1 'Open the document tab.
                            Else
                                RtfFileLocationType = LocationTypes.FileSystem
                                RtfFileDirectory = ItemInfo(FileName).Directory
                                XmlHtmDisplay1.LoadFile(RtfFileDirectory & "\" & RtfFileName)
                                DocumentTextChanged = False
                                TabControl1.SelectedIndex = 1 'Open the document tab.
                            End If
                        Case "XML"
                            'cmbDocType.SelectedIndex = cmbDocType.FindStringExact("XML")
                            FileType = FileTypes.XML 'Setting this property also displays the appropriate text in txtDocType
                            XmlFileName = FileName
                            txtFileName2.Text = FileName
                            txtFileType.Text = "XML"
                            txtFileDescription.Text = ItemInfo(FileName).Description
                            If ItemInfo(FileName).Directory = "" Then
                                XmlFileLocationType = LocationTypes.Project
                                XmlFileDirectory = ""
                                Dim xmlDoc As New System.Xml.XmlDocument
                                Project.ReadXmlDocData(XmlFileName, xmlDoc)
                                XmlHtmDisplay1.Clear()
                                XmlHtmDisplay1.Rtf = XmlHtmDisplay1.XmlToRtf(xmlDoc, True)
                                DocumentTextChanged = False
                                TabControl1.SelectedIndex = 1 'Open the document tab.
                            Else
                                XmlFileLocationType = LocationTypes.FileSystem
                                XmlFileDirectory = ItemInfo(FileName).Directory
                                XmlHtmDisplay1.ReadXmlFile(XmlFileDirectory & "\" & FileName, False)
                                DocumentTextChanged = False
                                TabControl1.SelectedIndex = 1 'Open the document tab.
                            End If
                        Case "HTML"
                            FileType = FileTypes.HTML 'Setting this property also displays the appropriate text in txtDocType
                            'cmbDocType.SelectedIndex = cmbDocType.FindStringExact("HTML")
                            HtmlFileName = FileName
                            txtFileName2.Text = FileName
                            txtFileType.Text = "HTML"
                            txtFileDescription.Text = ItemInfo(FileName).Description
                            If ItemInfo(FileName).Directory = "" Then
                                HtmlFileLocationType = LocationTypes.Project
                                HtmlFileDirectory = ""
                                Dim rtbData As New IO.MemoryStream
                                Project.ReadData(FileName, rtbData)
                                XmlHtmDisplay1.Clear()
                                rtbData.Position = 0
                                XmlHtmDisplay1.LoadFile(rtbData, RichTextBoxStreamType.PlainText)

                                If chkPlainText.Checked = True Then
                                    PlainTextDisplay = True
                                Else
                                    Dim htmText As String = XmlHtmDisplay1.Text
                                    XmlHtmDisplay1.Rtf = XmlHtmDisplay1.HmlToRtf(htmText)
                                    PlainTextDisplay = False
                                End If

                                DocumentTextChanged = False
                                TabControl1.SelectedIndex = 1 'Open the document tab.
                            Else
                                HtmlFileLocationType = LocationTypes.FileSystem
                                HtmlFileDirectory = ItemInfo(FileName).Directory
                                XmlHtmDisplay1.LoadFile(HtmlFileDirectory & "\" & FileName, RichTextBoxStreamType.PlainText)

                                If chkPlainText.Checked = True Then
                                    PlainTextDisplay = True
                                Else
                                    Dim htmText As String = XmlHtmDisplay1.Text
                                    XmlHtmDisplay1.Rtf = XmlHtmDisplay1.HmlToRtf(htmText)
                                    PlainTextDisplay = False
                                End If


                                DocumentTextChanged = False
                                TabControl1.SelectedIndex = 1 'Open the document tab.
                            End If


                        Case "TXT"
                            FileType = FileTypes.TXT 'Setting this property also displays the appropriate text in txtDocType
                            'cmbDocType.SelectedIndex = cmbDocType.FindStringExact("TXT")
                            TextFileName = FileName
                            txtFileName2.Text = FileName
                            txtFileType.Text = "TXT"
                            txtFileDescription.Text = ItemInfo(FileName).Description
                            If ItemInfo(FileName).Directory = "" Then
                                TextFileLocationType = LocationTypes.Project
                                TextFileDirectory = ""

                                Dim rtbData As New IO.MemoryStream
                                Project.ReadData(FileName, rtbData)
                                XmlHtmDisplay1.Clear()
                                rtbData.Position = 0
                                XmlHtmDisplay1.LoadFile(rtbData, RichTextBoxStreamType.PlainText)

                                DocumentTextChanged = False
                                TabControl1.SelectedIndex = 1 'Open the document tab.
                            Else
                                HtmlFileLocationType = LocationTypes.FileSystem
                                HtmlFileDirectory = ItemInfo(FileName).Directory
                                XmlHtmDisplay1.LoadFile(HtmlFileDirectory & "\" & FileName, RichTextBoxStreamType.PlainText)

                                DocumentTextChanged = False
                                TabControl1.SelectedIndex = 1 'Open the document tab.
                            End If

                    End Select
                Else
                    Message.AddWarning("No information is available on the selected node :" & FileName & vbCrLf)
                End If
            End If
        End If

    End Sub

    Private Sub btnOpenDocInNewWindow_Click(sender As Object, e As EventArgs) Handles btnOpenDocInNewWindow.Click
        'Open the document corresponding to the selected node.
        'The document will be displayed in a new document window.

        OpenDocInNewWindow()

    End Sub

    Private Sub OpenDocInNewWindow()
        'Open the selected document in a new window.

        If trvLibrary.SelectedNode Is Nothing Then
            Message.AddWarning("No document has been selected in the tree view." & vbCrLf)
        Else
            Dim FileName As String
            FileName = trvLibrary.SelectedNode.Name
            If FileName.EndsWith(".DocLib") Then
                Message.AddWarning("The Document Library node has been selected, not a document." & vbCrLf)
            ElseIf FileName.EndsWith(".DocColl") Then
                Message.AddWarning("A Document Collection node has been selected, not a document." & vbCrLf)
            Else
                'A document has been selected.
                If ItemInfo.ContainsKey(FileName) Then
                    Select Case ItemInfo(FileName).Type
                        Case "RTF"
                            OpenRtfWindow(FileName)
                            'Dim FormNo As Integer = NewRtfDisplay()
                            'RtfDisplayFormList(FormNo).FileName = FileName
                            'RtfDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                            'RtfDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                            'If RtfDisplayFormList(FormNo).FileDirectory = "" Then
                            '    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                            'Else
                            '    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                            'End If
                            'RtfDisplayFormList(FormNo).OpenDocument
                        Case "XML"
                            Dim FormNo As Integer = NewXmlDisplay()
                            XmlDisplayFormList(FormNo).FileName = FileName
                            XmlDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                            XmlDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                            If XmlDisplayFormList(FormNo).FileDirectory = "" Then
                                XmlDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                            Else
                                XmlDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                            End If
                            XmlDisplayFormList(FormNo).OpenDocument
                        Case "HTML"
                            If rbCodeView.Checked = True Then 'Display HTML Code View.
                                Dim FormNo As Integer = NewHtmlDisplay()
                                HtmlDisplayFormList(FormNo).FileName = FileName
                                HtmlDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                                HtmlDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                                If HtmlDisplayFormList(FormNo).FileDirectory = "" Then
                                    HtmlDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                                Else
                                    HtmlDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                                End If

                                If chkPlainText.Checked = True Then
                                    HtmlDisplayFormList(FormNo).OpenDocumentInPlainText
                                Else
                                    HtmlDisplayFormList(FormNo).OpenDocument
                                End If
                            Else 'Display HTML Web View.
                                Dim FormNo As Integer = NewWebView()
                                WebViewFormList(FormNo).FileName = FileName
                                WebViewFormList(FormNo).Description = ItemInfo(FileName).Description
                                WebViewFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                                If WebViewFormList(FormNo).FileDirectory = "" Then
                                    WebViewFormList(FormNo).FileLocationType = LocationTypes.Project
                                Else
                                    WebViewFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                                End If
                                WebViewFormList(FormNo).OpenDocument
                            End If
                        Case "TXT"
                            Dim FormNo As Integer = NewTextDisplay()
                            TextDisplayFormList(FormNo).FileName = FileName
                            TextDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                            TextDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                            If TextDisplayFormList(FormNo).FileDirectory = "" Then
                                TextDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                            Else
                                TextDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                            End If
                            TextDisplayFormList(FormNo).OpenDocument
                        Case "PDF"
                            Dim FormNo As Integer = NewPdfDisplay()
                            PdfDisplayFormList(FormNo).FileName = FileName
                            PdfDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                            PdfDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                            If PdfDisplayFormList(FormNo).FileDirectory = "" Then
                                PdfDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                            Else
                                PdfDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                            End If
                            PdfDisplayFormList(FormNo).OpenDocument
                        Case "XLS"
                            'Open the Excel spreadsheet file using the Excel application.
                            If Project.Type = ADVL_Utilities_Library_1.Project.Types.Archive Then
                                'Add code to extract the file from the Archive to a temporary directory.

                            Else
                                If Project.DataLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then
                                    'NEW CODE:
                                    'Use the Project.DataDirLocn - this a directory that stores files that are not suitable for archive file storage.
                                    If System.IO.File.Exists(Project.DataDirLocn.Path & "\" & FileName) Then
                                        Process.Start("EXCEL.EXE", """" & Project.DataDirLocn.Path & "\" & FileName & """")
                                    Else
                                        Message.AddWarning("Excel file not found: " & FileName & vbCrLf)
                                    End If

                                    'OLD CODE:
                                    ''Copy the Excel file to the Extracted Files sub-directory:
                                    'Project.CopyCheckArchiveDataToProjectDir(FileName, "Extracted Files")
                                    ''Final check that the extracted file exists:
                                    'If System.IO.File.Exists(Project.Path & "\Extracted Files\" & FileName) Then
                                    '    Process.Start("EXCEL.EXE", """" & Project.Path & "\Extracted Files\" & FileName & """")
                                    'Else
                                    '    Message.AddWarning("Problem extracting the Excel file: " & FileName & vbCrLf)
                                    'End If

                                    'OLDER CODE:
                                    ''Copy the Excel file to the Project directory if it is not already there.
                                    'If Project.ProjectFileExists(FileName) Then
                                    '    Process.Start("EXCEL.EXE", """" & Project.Path & "\" & FileName & """")
                                    '    'Message.Add("Excel file path: " & """" & Project.Path & "\" & FileName & """" & vbCrLf)
                                    'Else
                                    '    'Extract the file to the Project directory:
                                    '    Project.CopyArchiveDataToProject(FileName)
                                    '    'https://stackoverflow.com/questions/14497006/how-to-open-a-xls-file-with-excel-in-vb
                                    '    Try
                                    '        Process.Start("EXCEL.EXE", """" & Project.Path & "\" & FileName & """")
                                    '        'Message.Add("Excel file path: " & "" & Project.Path & "\" & FileName & "" & vbCrLf)
                                    '    Catch ex As Exception
                                    '        Message.AddWarning("Error opening Excel file: " & ex.Message & vbCrLf)
                                    '    End Try

                                    'End If
                                Else
                                    'Final check that the file exists:
                                    If System.IO.File.Exists(Project.DataLocn.Path & "\" & FileName) Then
                                        Process.Start("EXCEL.EXE", """" & Project.DataLocn.Path & "\" & FileName & """")
                                    Else
                                        Message.AddWarning("Excel file not found: " & FileName & vbCrLf)
                                    End If
                                End If
                            End If

                        Case "FolderLink"
                            'Open the folder using Windows Explorer:
                            Dim FolderPath As String = ItemInfo(FileName).Directory
                            Process.Start(FolderPath)

                        Case "XMsg"
                            'NOTE: TO BE UPDATED TO ALLOW THE XMSG TO BE RUN
                            Dim FormNo As Integer = NewXmlDisplay()
                            XmlDisplayFormList(FormNo).FileName = FileName
                            XmlDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                            XmlDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                            If XmlDisplayFormList(FormNo).FileDirectory = "" Then
                                XmlDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                            Else
                                XmlDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                            End If
                            XmlDisplayFormList(FormNo).OpenDocument

                        Case "XSeq"
                            'NOTE: TO BE UPDATED TO ALLOW THE XSEQ TO BE RUN
                            Dim FormNo As Integer = NewXmlDisplay()
                            XmlDisplayFormList(FormNo).FileName = FileName
                            XmlDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                            XmlDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                            If XmlDisplayFormList(FormNo).FileDirectory = "" Then
                                XmlDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                            Else
                                XmlDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                            End If
                            XmlDisplayFormList(FormNo).OpenDocument

                    End Select
                Else
                    Message.AddWarning("No information is available on the selected node :" & FileName & vbCrLf)
                End If
            End If
        End If
    End Sub

    Private Sub OpenRtfWindow(ByRef FileName As String)
        'Open the RTF file with the specified File Name.
        'If the RTF file is already displayed, bring the window to the front.
        'If the RTF file is not already displayed, open it in a new window.

        If FileName = "" Then

        Else
            'First check if the RTF file is already open:
            Dim FileFound As Boolean = False
            If RtfDisplayFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To RtfDisplayFormList.Count - 1
                    If RtfDisplayFormList(I) Is Nothing Then

                    Else
                        If RtfDisplayFormList(I).FileName = FileName Then
                            FileFound = True
                            RtfDisplayFormList(I).BringToFront
                            Exit For 'ADDED 11Sep2020
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = NewRtfDisplay()
                RtfDisplayFormList(FormNo).FileName = FileName
                RtfDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                RtfDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                If RtfDisplayFormList(FormNo).FileDirectory = "" Then
                    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                Else
                    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                End If
                RtfDisplayFormList(FormNo).OpenDocument
            End If

        End If

    End Sub


    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.RowIndex = -1 Then 'No cell changed
        Else
            If e.ColumnIndex = 5 Then 'Leaving Half Point Size column
                If IsNothing(DataGridView1.Rows(e.RowIndex).Cells(5).Value) Then 'No value in the Half Point Size column
                ElseIf Trim(DataGridView1.Rows(e.RowIndex).Cells(5).Value) = "" Then 'No value in the Half Point Size column
                Else
                    Dim HalfPointSize As Integer = Val(DataGridView1.Rows(e.RowIndex).Cells(5).Value)
                    Dim PointSize As Single = HalfPointSize / 2
                    DataGridView1.Rows(e.RowIndex).Cells(6).Value = PointSize
                End If
            ElseIf e.ColumnIndex = 6 Then 'Leaving Point Size Column
                If IsNothing(DataGridView1.Rows(e.RowIndex).Cells(6).Value) Then 'No value in the Point Size column
                ElseIf Trim(DataGridView1.Rows(e.RowIndex).Cells(6).Value) = "" Then 'No value in the Point Size column
                Else
                    Dim PointSize As Single = Val(DataGridView1.Rows(e.RowIndex).Cells(6).Value)
                    Dim HalfPointSize As Integer = Math.Round(PointSize * 2)
                    DataGridView1.Rows(e.RowIndex).Cells(5).Value = HalfPointSize
                End If
            End If
        End If

    End Sub

    Private Sub btnRemoveEntry_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnDefaultXml_Click(sender As Object, e As EventArgs) Handles btnDefaultXml.Click
        'Set the XML display settings to default values.
        XmlHtmDisplay1.DefaultXmlSettings()
        ShowSettings()
    End Sub

    Private Sub btnDefaultHtml_Click(sender As Object, e As EventArgs) Handles btnDefaultHtml.Click
        'Set the HTML display settings to default values.
        XmlHtmDisplay1.DefaultHtmlSettings()
        ShowSettings()
    End Sub

    Private Sub btnDefaultText_Click(sender As Object, e As EventArgs) Handles btnDefaultText.Click
        'Set the Plain Text display settings to default values.
        XmlHtmDisplay1.DefaultPlainTextSettings()
        ShowSettings()
    End Sub

    Private Sub XmlHtmDisplay1_ErrorMessage(Msg As String) Handles XmlHtmDisplay1.ErrorMessage
        Message.AddWarning(Msg)
    End Sub

    Private Sub XmlHtmDisplay1_Message(Msg As String) Handles XmlHtmDisplay1.Message
        Message.Add(Msg)
    End Sub

    Private Sub btnCodeView_Click(sender As Object, e As EventArgs) Handles btnCodeView.Click
        'Display HTML file in a code view window:

        Dim FileName As String
        FileName = trvLibrary.SelectedNode.Name

        Dim FormNo As Integer = NewHtmlDisplay()
        HtmlDisplayFormList(FormNo).FileName = FileName
        HtmlDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
        HtmlDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
        If HtmlDisplayFormList(FormNo).FileDirectory = "" Then
            HtmlDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
        Else
            HtmlDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
        End If

        If chkPlainText.Checked = True Then
            HtmlDisplayFormList(FormNo).OpenDocumentInPlainText
        Else
            HtmlDisplayFormList(FormNo).OpenDocument
        End If

    End Sub

    Private Sub btnWebView_Click(sender As Object, e As EventArgs) Handles btnWebView.Click
        'Display HTML file in a web view window:

        Dim FileName As String
        FileName = trvLibrary.SelectedNode.Name

        Dim FormNo As Integer = NewWebView()
        WebViewFormList(FormNo).FileName = FileName
        WebViewFormList(FormNo).Description = ItemInfo(FileName).Description
        WebViewFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
        If WebViewFormList(FormNo).FileDirectory = "" Then
            WebViewFormList(FormNo).FileLocationType = LocationTypes.Project
        Else
            WebViewFormList(FormNo).FileLocationType = LocationTypes.FileSystem
        End If
        WebViewFormList(FormNo).OpenDocument

    End Sub


#Region " Methods Called by JavaScript - A collection of methods that can be called by JavaScript in a web page shown in WebBrowser2" '========================================================
    'These methods are used to display HTML pages in the Document tab.
    'The same methods can be found in the WebView form, which displays web pages on seprate forms.


    'Display Messages ==============================================================================================

    Public Sub AddMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub AddWarning(ByVal Msg As String)
        Message.AddWarning(Msg)
    End Sub

    Public Sub AddTextTypeMessage(ByVal Msg As String, ByVal TextType As String)
        'Add a message with the specified Text Type to the Message window.
        Message.AddText(Msg, TextType)
    End Sub

    Public Sub AddXmlMessage(ByVal XmlText As String)
        'Add an Xml message to the Message window.
        Message.AddXml(XmlText)
    End Sub

    'END Display Messages ------------------------------------------------------------------------------------------


    'Run an XSequence ==============================================================================================

    Public Sub RunClipboardXSeq()
        'Run the XSequence instructions in the clipboard.

        Dim XDocSeq As System.Xml.Linq.XDocument
        Try
            XDocSeq = XDocument.Parse(My.Computer.Clipboard.GetText)
        Catch ex As Exception
            Message.AddWarning("Error reading Clipboard data. " & ex.Message & vbCrLf)
            Exit Sub
        End Try

        If IsNothing(XDocSeq) Then
            Message.Add("No XSequence instructions were found in the clipboard.")
        Else
            Dim XmlSeq As New System.Xml.XmlDocument
            Try
                XmlSeq.LoadXml(XDocSeq.ToString) 'Convert XDocSeq to an XmlDocument to process with XSeq.
                'Run the sequence:
                XSeq.RunXSequence(XmlSeq, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub RunXSequence(ByVal XSequence As String)
        'Run the XMSequence
        Dim XmlSeq As New System.Xml.XmlDocument
        XmlSeq.LoadXml(XSequence)
        XSeq.RunXSequence(XmlSeq, Status)
    End Sub

    Private Sub XSeq_ErrorMsg(ErrMsg As String) Handles XSeq.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub

    Private Sub XSeq_Instruction(Data As String, Locn As String) Handles XSeq.Instruction
        'Execute each instruction produced by running the XSeq file.

        Select Case Locn

            Case "Settings:Form:Name"
                FormName = Data

            Case "Settings:Form:Item:Name"
                ItemName = Data

            Case "Settings:Form:Item:Value"
                RestoreSetting(FormName, ItemName, Data)

            Case "Settings:Form:SelectId"
                SelectId = Data

            Case "Settings:Form:OptionText"
                RestoreOption(SelectId, Data)

            Case "Settings"

            Case "EndOfSequence"
                'Main.Message.Add("End of processing sequence" & Data & vbCrLf)

            Case Else
                'Main.Message.AddWarning("Unknown location: " & Locn & "  Data: " & Data & vbCrLf)
                Message.AddWarning("Unknown location: " & Locn & "  Data: " & Data & vbCrLf)

        End Select
    End Sub

    'END Run an XSequence ------------------------------------------------------------------------------------------


    'Run an XMessage ===============================================================================================

    Public Sub RunXMessage(ByVal XMsg As String)
        'Run the XMessage by sending it to InstrReceived.
        InstrReceived = XMsg
    End Sub

    Public Sub SendXMessage(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMessage to the application with the connection name ConnName.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New Main.clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageExt(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetname.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                If bgwSendMessage.IsBusy Then
                    Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                Else
                    Dim SendMessageParams As New Main.clsSendMessageParams
                    SendMessageParams.ProjectNetworkName = ProNetName
                    SendMessageParams.ConnectionName = ConnName
                    SendMessageParams.Message = XMsg
                    bgwSendMessage.RunWorkerAsync(SendMessageParams)
                    If ShowXMessages Then
                        Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                        Message.XAddXml(XMsg)
                        Message.XAddText(vbCrLf, "Normal") 'Add extra line
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub SendXMessageWait(ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName.
        'Wait for the connection to be made.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            Try
                'Application.DoEvents() 'TRY THE METHOD WITHOUT THE DOEVENTS
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    Message.Add("client state is faulted. Message not sent!" & vbCrLf)
                Else
                    Dim StartTime As Date = Now
                    Dim Duration As TimeSpan
                    'Wait up to 16 seconds for the connection ConnName to be established
                    While client.ConnectionExists(ProNetName, ConnName) = False 'Wait until the required connection is made.
                        System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                        Duration = Now - StartTime
                        If Duration.Seconds > 16 Then Exit While
                    End While

                    If client.ConnectionExists(ProNetName, ConnName) = False Then
                        Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                    Else
                        If bgwSendMessage.IsBusy Then
                            Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                        Else
                            Dim SendMessageParams As New Main.clsSendMessageParams
                            SendMessageParams.ProjectNetworkName = ProNetName
                            SendMessageParams.ConnectionName = ConnName
                            SendMessageParams.Message = XMsg
                            bgwSendMessage.RunWorkerAsync(SendMessageParams)
                            If ShowXMessages Then
                                Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                                Message.XAddXml(XMsg)
                                Message.XAddText(vbCrLf, "Normal") 'Add extra line
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Message.AddWarning(ex.Message & vbCrLf)
            End Try
        End If
    End Sub

    Public Sub SendXMessageExtWait(ByVal ProNetName As String, ByVal ConnName As String, ByVal XMsg As String)
        'Send the XMsg to the application with the connection name ConnName and Project Network Name ProNetName.
        'Wait for the connection to be made.
        'This version can send the XMessage to a connection external to the current Project Network.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                Dim StartTime As Date = Now
                Dim Duration As TimeSpan
                'Wait up to 16 seconds for the connection ConnName to be established
                While client.ConnectionExists(ProNetName, ConnName) = False
                    System.Threading.Thread.Sleep(1000) 'Pause for 1000ms
                    Duration = Now - StartTime
                    If Duration.Seconds > 16 Then Exit While
                End While

                If client.ConnectionExists(ProNetName, ConnName) = False Then
                    Message.AddWarning("Connection not available: " & ConnName & " in application network: " & ProNetName & vbCrLf)
                Else
                    If bgwSendMessage.IsBusy Then
                        Message.AddWarning("Send Message backgroundworker is busy." & vbCrLf)
                    Else
                        Dim SendMessageParams As New Main.clsSendMessageParams
                        SendMessageParams.ProjectNetworkName = ProNetName
                        SendMessageParams.ConnectionName = ConnName
                        SendMessageParams.Message = XMsg
                        bgwSendMessage.RunWorkerAsync(SendMessageParams)
                        If ShowXMessages Then
                            Message.XAddText("Message sent to " & "[" & ProNetName & "]." & ConnName & ":" & vbCrLf, "XmlSentNotice")
                            Message.XAddXml(XMsg)
                            Message.XAddText(vbCrLf, "Normal") 'Add extra line
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub XMsgInstruction(ByVal Info As String, ByVal Locn As String)
        'Send the XMessage Instruction to the JavaScript function XMsgInstruction for processing.
        Me.WebBrowser1.Document.InvokeScript("XMsgInstruction", New String() {Info, Locn})
    End Sub

    'END Run an XMessage -------------------------------------------------------------------------------------------


    'Get Information ===============================================================================================

    Public Function GetFormNo() As String
        'Return FormNo.ToString
        Return "-1" 'The Main Form is not a Web Page form.
    End Function

    Public Function GetParentFormNo() As String
        'Return the Form Number of the Parent Form (that called this form).
        'Return ParentWebPageFormNo.ToString
        Return "-1" 'The Main Form does not have a Parent Web Page.
    End Function

    Public Function GetConnectionName() As String
        'Return the Connection Name of the Project.
        Return ConnectionName
    End Function

    Public Function GetProNetName() As String
        'Return the Project Network Name of the Project.
        Return ProNetName
    End Function

    Public Sub ParentProjectName(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project name:
        RestoreSetting(FormName, ItemName, Project.ParentProjectName)
    End Sub

    Public Sub ParentProjectPath(ByVal FormName As String, ByVal ItemName As String)
        'Return the Parent Project path:
        RestoreSetting(FormName, ItemName, Project.ParentProjectPath)
    End Sub

    Public Sub ParentProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Parent Project parameter value:
        RestoreSetting(FormName, ItemName, Project.ParentParameter(ParameterName).Value)
    End Sub

    Public Sub ProjectParameterValue(ByVal FormName As String, ByVal ItemName As String, ByVal ParameterName As String)
        'Return the specified Project parameter value:
        RestoreSetting(FormName, ItemName, Project.Parameter(ParameterName).Value)
    End Sub

    Public Sub ProjectNetworkName(ByVal FormName As String, ByVal ItemName As String)
        'Return the name of the Project Network:
        RestoreSetting(FormName, ItemName, Project.Parameter("ProNetName").Value)
    End Sub

    'END Get Information -------------------------------------------------------------------------------------------


    'Open a Web Page ===============================================================================================

    Public Sub OpenWebPage(ByVal FileName As String)
        'Open the web page with the specified File Name.

        If FileName = "" Then

        Else
            'First check if the HTML file is already open:
            Dim FileFound As Boolean = False
            If WebPageFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To WebPageFormList.Count - 1
                    If WebPageFormList(I) Is Nothing Then

                    Else
                        If WebPageFormList(I).FileName = FileName Then
                            FileFound = True
                            WebPageFormList(I).BringToFront
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = OpenNewWebPage()
                WebPageFormList(FormNo).FileName = FileName
                WebPageFormList(FormNo).OpenDocument
                WebPageFormList(FormNo).BringToFront
            End If
        End If
    End Sub

    'END Open a Web Page -------------------------------------------------------------------------------------------


    'Open and Close Projects =======================================================================================

    Public Sub OpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            client.StartProjectAtPath(ProjectPath, ConnectionName)
        End If
    End Sub

    Public Sub CheckOpenProjectAtRelativePath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Relative Path is open.
        'Open it if it is not already open.
        'Open the Project at the specified Relative Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            ProjectPath = Project.Path & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        Else
            ProjectPath = Project.Path & "\" & RelativePath
            If client.ProjectOpen(ProjectPath) Then
                'Project is already open.
            Else
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            End If
        End If
    End Sub

    Public Sub OpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Open the Project at the specified Path (relative to the ProNet Path) using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                client.StartProjectAtPath(ProjectPath, ConnectionName)
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CheckOpenProjectAtProNetPath(ByVal RelativePath As String, ByVal ConnectionName As String)
        'Check if the project at the specified Path (relative to the ProNet Path) is open.
        'Open it if it is not already open.
        'Open the Project at the specified Path using the specified Connection Name.

        Dim ProjectPath As String
        If RelativePath.StartsWith("\") Then
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & RelativePath
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        Else
            If Project.ParameterExists("ProNetPath") Then
                ProjectPath = Project.GetParameter("ProNetPath") & "\" & RelativePath
                If client.ProjectOpen(ProjectPath) Then
                    'Project is already open.
                Else
                    client.StartProjectAtPath(ProjectPath, ConnectionName)
                End If
            Else
                Message.AddWarning("The Project Network Path is not known." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub CloseProjectAtConnection(ByVal ProNetName As String, ByVal ConnectionName As String)
        'Close the application at the specified connection.

        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            If client.State = ServiceModel.CommunicationState.Faulted Then
                Message.Add("client state is faulted. Message not sent!" & vbCrLf)
            Else
                'Create the XML instructions to close the Project at the connection.
                Dim decl As New XDeclaration("1.0", "utf-8", "yes")
                Dim doc As New XDocument(decl, Nothing) 'Create an XDocument to store the instructions.
                Dim xmessage As New XElement("XMsg") 'This indicates the start of the message in the XMessage class

                'NOTE: No reply expected. No need to provide the following client information(?)
                'Dim clientConnName As New XElement("ClientConnectionName", Me.ConnectionName)
                'xmessage.Add(clientConnName)

                Dim command As New XDocument("Command", "Close")
                xmessage.Add(command)
                doc.Add(xmessage)

                'Show the message sent to:
                Message.XAddText("Message sent to [" & ProNetName & "]." & ConnectionName & ":" & vbCrLf, "XmlSentNotice")
                Message.XAddXml(doc.ToString)
                Message.XAddText(vbCrLf, "Normal") 'Add extra line

                client.SendMessage(ProNetName, ConnectionName, doc.ToString)
            End If
        End If
    End Sub

    'END Open and Close Projects -----------------------------------------------------------------------------------


    'System Methods ================================================================================================

    Public Sub SaveHtmlSettings(ByVal xSettings As String, ByVal FileName As String)
        'Save the Html settings for a web page.

        'Convert the XSettings to XML format:
        Dim XmlHeader As String = "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
        Dim XDocSettings As New System.Xml.Linq.XDocument

        Try
            XDocSettings = System.Xml.Linq.XDocument.Parse(XmlHeader & vbCrLf & xSettings)
        Catch ex As Exception
            Message.AddWarning("Error saving HTML settings file. " & ex.Message & vbCrLf)
        End Try

        Project.SaveXmlData(FileName, XDocSettings)
    End Sub

    Public Sub RestoreHtmlSettings()
        'Restore the Html settings for a web page.

        Dim SettingsFileName As String = txtNodeKey.Text & "Settings"
        Dim XDocSettings As New System.Xml.Linq.XDocument
        Project.ReadXmlData(SettingsFileName, XDocSettings)

        If XDocSettings Is Nothing Then
            'NOTE: THE FOLLOWING MESSAGE IS ALWAYS SHOWN. IS SETTINGS FILE USED HERE???
            'Message.Add("No HTML Settings file : " & SettingsFileName & vbCrLf)
        Else
            Dim XSettings As New System.Xml.XmlDocument
            Try
                XSettings.LoadXml(XDocSettings.ToString)
                'Run the Settings file:
                XSeq.RunXSequence(XSettings, Status)
            Catch ex As Exception
                Message.AddWarning("Error restoring HTML settings. " & ex.Message & vbCrLf)
            End Try

        End If
    End Sub

    Public Sub RestoreSetting(ByVal FormName As String, ByVal ItemName As String, ByVal ItemValue As String)
        'Restore the setting value with the specified Form Name and Item Name.
        Me.WebBrowser2.Document.InvokeScript("RestoreSetting", New String() {FormName, ItemName, ItemValue})
    End Sub

    Public Sub RestoreOption(ByVal SelectId As String, ByVal OptionText As String)
        'Restore the Option text in the Select control with the Id SelectId.
        Me.WebBrowser2.Document.InvokeScript("RestoreOption", New String() {SelectId, OptionText})
    End Sub

    Private Sub SaveWebPageSettings()
        'Call the SaveSettings JavaScript function:
        Try
            Me.WebBrowser2.Document.InvokeScript("SaveSettings")
        Catch ex As Exception
            Message.AddWarning("Web page settings not saved: " & ex.Message & vbCrLf)
        End Try
    End Sub

    'END System Methods --------------------------------------------------------------------------------------------


    'Legacy Code (These methods should no longer be used) ==========================================================

    Public Sub JSMethodTest1()
        'Test method that is called from JavaScript.
        Message.Add("JSMethodTest1 called OK." & vbCrLf)
    End Sub

    Public Sub JSMethodTest2(ByVal Var1 As String, ByVal Var2 As String)
        'Test method that is called from JavaScript.
        Message.Add("Var1 = " & Var1 & " Var2 = " & Var2 & vbCrLf)
    End Sub

    Public Sub JSDisplayXml(ByRef XDoc As XDocument)
        Message.Add(XDoc.ToString & vbCrLf & vbCrLf)
    End Sub

    Public Sub ShowMessage(ByVal Msg As String)
        Message.Add(Msg)
    End Sub

    Public Sub JSSendXMessage(ByVal ProNetName As String, ByVal XMsg As String, ByVal Destination As String)
        'Send an XMessage to the specified destination.
        If IsNothing(client) Then
            Message.Add("No client connection available!" & vbCrLf)
        Else
            client.SendMessage(ProNetName, Destination, XMsg)
        End If
    End Sub

    Public Sub AddText(ByVal Msg As String, ByVal TextType As String)
        Message.AddText(Msg, TextType)
    End Sub

    'END Legacy Code -----------------------------------------------------------------------------------------------


#End Region 'Methods Called by JavaScript -----------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub ApplicationInfo_UpdateExePath() Handles ApplicationInfo.UpdateExePath

    End Sub

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        'Update the current duration:

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)

        Timer2.Interval = 5000 '5 seconds
        Timer2.Enabled = True
        Timer2.Start()

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        'Update the current duration:

        txtCurrentDuration.Text = Project.Usage.CurrentDuration.Days.ToString.PadLeft(5, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Hours.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Minutes.ToString.PadLeft(2, "0"c) & ":" &
                                  Project.Usage.CurrentDuration.Seconds.ToString.PadLeft(2, "0"c)
    End Sub

    Private Sub TabPage2_Leave(sender As Object, e As EventArgs) Handles TabPage2.Leave
        Timer2.Enabled = False
    End Sub

    Private Sub btnSelectNode_Click(sender As Object, e As EventArgs) Handles btnSelectNode.Click
        'Select the node in trvLibrary with the key: txtSelNodeKey

        Dim node As TreeNode() = trvLibrary.Nodes.Find(txtSelNodeKey.Text, True)
        If node Is Nothing Then
            'Node key (name) not found.
        Else
            trvLibrary.SelectedNode = node(0)
            trvLibrary.Focus()
        End If
    End Sub

    Private Sub trvLibrary_DoubleClick(sender As Object, e As EventArgs) Handles trvLibrary.DoubleClick
        'The Library Tree View has been double-clicked.
        'Open any document corresponding to the selected node in a new window:
        OpenDocInNewWindow()
    End Sub

    Private Sub ToolStripMenuItem1_OpenDocument_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_OpenDocument.Click
        OpenDocInNewWindow()
    End Sub

    Private Sub trvLibrary_DragDrop(sender As Object, e As DragEventArgs) Handles trvLibrary.DragDrop
        'DragDrop:

        Dim Path As String()
        Path = e.Data.GetData(DataFormats.FileDrop)

        Message.Add(vbCrLf & "------------------------------------------------------------------------------------------------------------ " & vbCrLf) 'Add separator line.
        Message.Add("Path.Count: " & Path.Count & vbCrLf)

        Dim I As Integer
        Dim pt As Point
        Dim DestinationNode As TreeNode
        pt = CType(sender, TreeView).PointToClient(New Point(e.X, e.Y))
        DestinationNode = CType(sender, TreeView).GetNodeAt(pt)
        Message.Add("Destination node key: " & DestinationNode.Name & vbCrLf)

        For I = 0 To Path.Count - 1
            Message.Add(vbCrLf & "Path(" & I & "): " & Path(I) & vbCrLf)
            ProcessNewFile(Path(I), DestinationNode.Name)
        Next

    End Sub

    Private Sub trvLibrary_DragEnter(sender As Object, e As DragEventArgs) Handles trvLibrary.DragEnter
        'DragEnter: An object has been dragged into the trvLibrary.

        'This code is required to get the link to the item(s) being dragged into the trvLibrary:
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Link
        End If
    End Sub

    Private Sub ProcessNewFile(ByVal FilePath As String, ByVal NodeKey As String)
        'Process a file that has been dragged into the Library Tree View:

        Message.Add(vbCrLf & "Processing File:" & vbCrLf)
        Message.Add(FilePath & vbCrLf)

        'Check if ProjectPath is a File or a Directory:
        Dim Attr As System.IO.FileAttributes = IO.File.GetAttributes(FilePath)
        If Attr.HasFlag(IO.FileAttributes.Directory) Then
            Message.Add("The item is a Directory." & vbCrLf)
            Dim NewFolderLink As String = FilePath.Replace(":", " ") & ".FolderLink" 'Removing the ":" character allows it to be saved in an XElement. (The ":" character is not allowed in an XElement name.)

            If ItemInfo.ContainsKey(NewFolderLink) Then
                Message.AddWarning("A node is already has the NodeKey: " & NewFolderLink & vbCrLf)
            Else
                'Check if the Folder Link is valid:
                If System.IO.Directory.Exists(FilePath) Then
                    'ItemInfo.Add(NodeKey, New clsItemInfo)
                    ItemInfo.Add(NewFolderLink, New clsItemInfo)
                    Dim NodeText As String = Trim(FilePath) 'The text that appears on the node in the treeview.
                    If trvLibrary.SelectedNode Is Nothing Then
                        'Image 16 is a closed folder, image 17 is an open folder.
                        'trvLibrary.Nodes.Add(NodeKey, NodeText, 16, 17) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        trvLibrary.Nodes.Add(NewFolderLink, NodeText, 16, 17) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                    Else
                        'Image 16 is a closed folder, image 17 is an open folder.
                        'trvLibrary.SelectedNode.Nodes.Add(NodeKey, NodeText, 16, 17) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        trvLibrary.SelectedNode.Nodes.Add(NewFolderLink, NodeText, 16, 17) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                    End If

                    'Add the Node Text to the ItemInfo dictionary: - ADDED 16 September 2020 - Used in the Document List tab.
                    ItemInfo(NodeKey).Text = NodeText

                    'ItemInfo(NodeKey).Type = "FolderLink"
                    ItemInfo(NewFolderLink).Type = "FolderLink"
                    'ItemInfo(NodeKey).Description = "Link to folder: " & FilePath
                    ItemInfo(NewFolderLink).Description = "Link to folder: " & FilePath
                    'ItemInfo(NodeKey).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    ItemInfo(NewFolderLink).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    'ItemInfo(NodeKey).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    ItemInfo(NewFolderLink).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                    'ItemInfo(NodeKey).Directory = FilePath
                    ItemInfo(NewFolderLink).Directory = FilePath
                Else
                    Message.AddWarning("The folder link is not valid: " & NewFolderLink & vbCrLf)
                End If
            End If
                Else
            If LCase(FilePath).EndsWith(".pdf") Then
                Message.Add("The file is a PDF document." & vbCrLf)
                'Get the FileName of the new PDF file:
                Dim NewFileName As String
                NewFileName = System.IO.Path.GetFileName(FilePath)
                Message.Add("The file name is: " & NewFileName & vbCrLf)
                'Check if the file is already in the project:
                If Project.DataDirFileExists(NewFileName) Then 'NOTE: .pdf files are stored in the DataDirLocn.
                    Message.Add("The File Name already exists in the project." & vbCrLf)
                Else
                    If ItemInfo.ContainsKey(NewFileName) Then
                        Message.AddWarning("A node is already has the NodeKey: " & NewFileName & vbCrLf)
                    Else
                        'Copy the file to the Project Data Location.
                        If Project.DataDirLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then 'Use a filestream to write data to an archive.
                            Dim fileStream As New IO.FileStream(FilePath, System.IO.FileMode.Open)
                            Dim myData(fileStream.Length) As Byte
                            fileStream.Read(myData, 0, fileStream.Length)
                            Dim pdfData As New IO.MemoryStream
                            Dim streamWriter As New IO.BinaryWriter(pdfData)
                            streamWriter.Write(myData)
                            fileStream.Close()
                            Project.SaveData(NewFileName, pdfData)
                        Else 'Use a simple file copy to write data to a directory.
                            My.Computer.FileSystem.CopyFile(FilePath, Project.DataDirLocn.Path & "\" & NewFileName)
                        End If

                        'Add ItemInfo entry:
                        ItemInfo.Add(NewFileName, New clsItemInfo)
                        Dim NodeText As String = NewFileName 'The text that appears on the node in the treeview.
                        ItemInfo(NewFileName).Type = "PDF"
                        ItemInfo(NewFileName).Description = ""
                        ItemInfo(NewFileName).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).Directory = ""

                        'Add the PDF node to the tree.
                        'Select the Parent Node:
                        Dim node As TreeNode() = trvLibrary.Nodes.Find(NodeKey, True)
                        If node Is Nothing Then
                            'NodeKey not found
                            Message.AddWarning("Node Key not found: " & NodeKey & vbCrLf)
                        Else
                            trvLibrary.SelectedNode = node(0)
                            trvLibrary.SelectedNode.Nodes.Add(NewFileName, NewFileName, 12, 13) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If
                    End If
                End If
            ElseIf LCase(FilePath).EndsWith(".xls") Or LCase(FilePath).EndsWith(".xlsx") Then
                Message.Add("The file is an Excel spreadsheet." & vbCrLf)
                'Get the FileName of the new Excel file:
                Dim NewFileName As String = System.IO.Path.GetFileName(FilePath)
                'Check if the file is already in the project:
                If Project.DataFileExists(NewFileName) Then
                    Message.Add("The File Name already exists in the project." & vbCrLf)
                Else
                    If ItemInfo.ContainsKey(NewFileName) Then
                        Message.AddWarning("A node is already has the NodeKey: " & NewFileName & vbCrLf)
                    Else
                        'Copy the file to the Project Data Location.
                        If Project.DataDirLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then 'Use a filestream to write data to an archive.
                            'NOTE: This method appears to corrupt .xlsx files! To be checked! (The DataDirLocn will usually be a directory and use a simple file copy - only an issue for an Archive project.)
                            Dim fileStream As New IO.FileStream(FilePath, System.IO.FileMode.Open)
                            Dim myData(fileStream.Length) As Byte
                            fileStream.Read(myData, 0, fileStream.Length)
                            Dim pdfData As New IO.MemoryStream
                            Dim streamWriter As New IO.BinaryWriter(pdfData)
                            streamWriter.Write(myData)
                            fileStream.Close()
                            Project.SaveData(NewFileName, pdfData)
                        Else 'Use a simple file copy to write data to a directory.
                            My.Computer.FileSystem.CopyFile(FilePath, Project.DataDirLocn.Path & "\" & NewFileName)
                        End If

                        'Add ItemInfo entry:
                        ItemInfo.Add(NewFileName, New clsItemInfo)
                        Dim NodeText As String = NewFileName 'The text that appears on the node in the treeview.
                        ItemInfo(NewFileName).Type = "XLS"
                        ItemInfo(NewFileName).Description = ""
                        ItemInfo(NewFileName).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).Directory = ""

                        'Add the XLS node to the tree.
                        'Select the Parent Node:
                        Dim node As TreeNode() = trvLibrary.Nodes.Find(NodeKey, True)
                        If node Is Nothing Then
                            'NodeKey not found
                            Message.AddWarning("Node Key not found: " & NodeKey & vbCrLf)
                        Else
                            trvLibrary.SelectedNode = node(0)
                            trvLibrary.SelectedNode.Nodes.Add(NewFileName, NewFileName, 14, 15) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If
                    End If
                End If

            ElseIf LCase(FilePath).EndsWith(".rtf") Then
                Message.Add("The file is a Rich Text Format document." & vbCrLf)
                'Get the FileName of the new RTF file:
                Dim NewFileName As String = System.IO.Path.GetFileName(FilePath)
                'Check if the file is already in the project:
                If Project.DataFileExists(NewFileName) Then
                    Message.Add("The File Name already exists in the project." & vbCrLf)
                Else
                    If ItemInfo.ContainsKey(NewFileName) Then
                        Message.AddWarning("A node is already has the NodeKey: " & NewFileName & vbCrLf)
                    Else
                        'Copy the file to the Project Data Location.
                        If Project.DataDirLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then 'Use a filestream to write data to an archive.
                            Dim fileStream As New IO.FileStream(FilePath, System.IO.FileMode.Open)
                            Dim myData(fileStream.Length) As Byte
                            fileStream.Read(myData, 0, fileStream.Length)
                            Dim pdfData As New IO.MemoryStream
                            Dim streamWriter As New IO.BinaryWriter(pdfData)
                            streamWriter.Write(myData)
                            fileStream.Close()
                            Project.SaveData(NewFileName, pdfData)
                        Else 'Use a simple file copy to write data to a directory.
                            My.Computer.FileSystem.CopyFile(FilePath, Project.DataDirLocn.Path & "\" & NewFileName)
                        End If

                        'Add ItemInfo entry:
                        ItemInfo.Add(NewFileName, New clsItemInfo)
                        Dim NodeText As String = NewFileName 'The text that appears on the node in the treeview.
                        ItemInfo(NewFileName).Type = "RTF"
                        ItemInfo(NewFileName).Description = ""
                        ItemInfo(NewFileName).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).Directory = ""

                        'Add the XLS node to the tree.
                        'Select the Parent Node:
                        Dim node As TreeNode() = trvLibrary.Nodes.Find(NodeKey, True)
                        If node Is Nothing Then
                            'NodeKey not found
                            Message.AddWarning("Node Key not found: " & NodeKey & vbCrLf)
                        Else
                            trvLibrary.SelectedNode = node(0)
                            trvLibrary.SelectedNode.Nodes.Add(NewFileName, NewFileName, 4, 5) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If
                    End If
                End If

            ElseIf _newItemType = NewItemTypes.XMsg Then
                Message.Add("The file is an XMessage." & vbCrLf)
                'Get the FileName of the new XMessage file:
                Dim NewFileName As String = System.IO.Path.GetFileName(FilePath)
                'Check if the file is already in the project:
                If Project.DataFileExists(NewFileName) Then
                    Message.Add("The File Name already exists in the project." & vbCrLf)
                Else
                    If ItemInfo.ContainsKey(NewFileName) Then
                        Message.AddWarning("A node is already has the NodeKey: " & NewFileName & vbCrLf)
                    Else
                        'Copy the file to the Project Data Location.
                        If Project.DataDirLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then 'Use a filestream to write data to an archive.
                            'NOTE: This method appears to corrupt .xlsx files! To be checked! (The DataDirLocn will usually be a directory and use a simple file copy - only an issue for an Archive project.)
                            Dim fileStream As New IO.FileStream(FilePath, System.IO.FileMode.Open)
                            Dim myData(fileStream.Length) As Byte
                            fileStream.Read(myData, 0, fileStream.Length)
                            Dim pdfData As New IO.MemoryStream
                            Dim streamWriter As New IO.BinaryWriter(pdfData)
                            streamWriter.Write(myData)
                            fileStream.Close()
                            Project.SaveData(NewFileName, pdfData)
                        Else 'Use a simple file copy to write data to a directory.
                            My.Computer.FileSystem.CopyFile(FilePath, Project.DataDirLocn.Path & "\" & NewFileName)
                        End If

                        'Add ItemInfo entry:
                        ItemInfo.Add(NewFileName, New clsItemInfo)
                        Dim NodeText As String = NewFileName 'The text that appears on the node in the treeview.
                        ItemInfo(NewFileName).Type = "XMsg"
                        ItemInfo(NewFileName).Description = ""
                        ItemInfo(NewFileName).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).Directory = ""

                        'Add the XMsg node to the tree.
                        'Select the Parent Node:
                        Dim node As TreeNode() = trvLibrary.Nodes.Find(NodeKey, True)
                        If node Is Nothing Then
                            'NodeKey not found
                            Message.AddWarning("Node Key not found: " & NodeKey & vbCrLf)
                        Else
                            trvLibrary.SelectedNode = node(0)
                            trvLibrary.SelectedNode.Nodes.Add(NewFileName, NewFileName, 18, 19) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If
                    End If
                End If

            ElseIf _newItemType = NewItemTypes.XSeq Then
                Message.Add("The file is an XSequence." & vbCrLf)
                'Get the FileName of the new XMessage file:
                Dim NewFileName As String = System.IO.Path.GetFileName(FilePath)
                'Check if the file is already in the project:
                If Project.DataFileExists(NewFileName) Then
                    Message.Add("The File Name already exists in the project." & vbCrLf)
                Else
                    If ItemInfo.ContainsKey(NewFileName) Then
                        Message.AddWarning("A node is already has the NodeKey: " & NewFileName & vbCrLf)
                    Else
                        'Copy the file to the Project Data Location.
                        If Project.DataDirLocn.Type = ADVL_Utilities_Library_1.FileLocation.Types.Archive Then 'Use a filestream to write data to an archive.
                            'NOTE: This method appears to corrupt .xlsx files! To be checked! (The DataDirLocn will usually be a directory and use a simple file copy - only an issue for an Archive project.)
                            Dim fileStream As New IO.FileStream(FilePath, System.IO.FileMode.Open)
                            Dim myData(fileStream.Length) As Byte
                            fileStream.Read(myData, 0, fileStream.Length)
                            Dim pdfData As New IO.MemoryStream
                            Dim streamWriter As New IO.BinaryWriter(pdfData)
                            streamWriter.Write(myData)
                            fileStream.Close()
                            Project.SaveData(NewFileName, pdfData)
                        Else 'Use a simple file copy to write data to a directory.
                            My.Computer.FileSystem.CopyFile(FilePath, Project.DataDirLocn.Path & "\" & NewFileName)
                        End If

                        'Add ItemInfo entry:
                        ItemInfo.Add(NewFileName, New clsItemInfo)
                        Dim NodeText As String = NewFileName 'The text that appears on the node in the treeview.
                        ItemInfo(NewFileName).Type = "XSeq"
                        ItemInfo(NewFileName).Description = ""
                        ItemInfo(NewFileName).CreationDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).LastEditDate = Format(Now, "d-MMM-yyyy H:mm:ss")
                        ItemInfo(NewFileName).Directory = ""

                        'Add the XSeq node to the tree.
                        'Select the Parent Node:
                        Dim node As TreeNode() = trvLibrary.Nodes.Find(NodeKey, True)
                        If node Is Nothing Then
                            'NodeKey not found
                            Message.AddWarning("Node Key not found: " & NodeKey & vbCrLf)
                        Else
                            trvLibrary.SelectedNode = node(0)
                            trvLibrary.SelectedNode.Nodes.Add(NewFileName, NewFileName, 20, 21) 'key As String, text As String, imageIndex As Integer, selectedImageIndex As Integer
                        End If
                    End If
                End If

            Else
                Message.Add("The file is a not PDF or XLS document." & vbCrLf)
            End If
        End If
    End Sub


    Private Sub chkConnect_LostFocus(sender As Object, e As EventArgs) Handles chkConnect.LostFocus
        If chkConnect.Checked Then
            Project.ConnectOnOpen = True
        Else
            Project.ConnectOnOpen = False
        End If
        Project.SaveProjectInfoFile()
    End Sub

    'Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
    '    'Keet the connection awake with each tick:

    '    If ConnectedToComnet = True Then
    '        Try
    '            If client.IsAlive() Then
    '                Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
    '                Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '            Else
    '                Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf)
    '                Timer3.Interval = TimeSpan.FromMinutes(55).TotalMilliseconds '55 minute interval
    '            End If
    '        Catch ex As Exception
    '            Message.AddWarning(ex.Message & vbCrLf)
    '            'Set interval to five minutes - try again in five minutes:
    '            Timer3.Interval = TimeSpan.FromMinutes(5).TotalMilliseconds '5 minute interval
    '        End Try
    '    Else
    '        Message.Add(Format(Now, "HH:mm:ss") & " Not connected." & vbCrLf)
    '    End If
    'End Sub

    Private Sub chkConnect_CheckedChanged(sender As Object, e As EventArgs) Handles chkConnect.CheckedChanged

    End Sub

    Private Sub btnFindText_Click(sender As Object, e As EventArgs) Handles btnFindText.Click
        'Find the text shown in txtFindText.Text
        'Look in the document associated with the selected node

        If trvLibrary.SelectedNode Is Nothing Then
            Message.AddWarning("No document has been selected in the tree view." & vbCrLf)
        Else

            Dim FileName As String
            FileName = trvLibrary.SelectedNode.Name
            If FileName.EndsWith(".DocLib") Then
                Message.AddWarning("The Document Library node has been selected, not a document." & vbCrLf)
            ElseIf FileName.EndsWith(".DocColl") Then
                Message.AddWarning("A Document Collection node has been selected, not a document." & vbCrLf)
            Else
                'A document has been selected.
                If ItemInfo.ContainsKey(FileName) Then
                    Select Case ItemInfo(FileName).Type
                        Case "RTF"
                            'OpenRtfWindow(FileName) 'This shows the Rtf window if it already exists or creates a new one to display the file.
                            'FindFirstRtf(FileName, txtFindText.Text)
                            If chkHighlight.Checked Then
                                If chkFindFirst.Checked Then
                                    FindTextInRtf(FileName, txtFindText.Text, True, True)
                                Else
                                    FindTextInRtf(FileName, txtFindText.Text, True, False)
                                End If
                            Else
                                If chkFindFirst.Checked Then
                                    FindTextInRtf(FileName, txtFindText.Text, False, True)
                                Else
                                    FindTextInRtf(FileName, txtFindText.Text, False, False)
                                End If
                            End If
                        Case "XML"
                            Message.Add("Code not yet added to search for text in an XML file." & vbCrLf)

                        Case "HTML"
                            If rbCodeView.Checked = True Then 'Display HTML Code View.
                                Message.Add("Code not yet added to search for text in a HTML file." & vbCrLf)

                            Else 'Display HTML Web View.
                                Message.Add("Code not yet added to search for text in a Web page." & vbCrLf)

                            End If
                        Case "TXT"
                            Message.Add("Code not yet added to search for text in a Text file." & vbCrLf)

                        Case "PDF"
                            Message.Add("Code not yet added to search for text in a PDF file." & vbCrLf)

                    End Select
                Else
                    Message.AddWarning("No information is available on the selected node :" & FileName & vbCrLf)
                End If
            End If

        End If

    End Sub

    Private Sub FindFirstRtf(ByVal FileName As String, ByVal SearchText As String)
        'Find the first instance of the SearchText in the RTF window displaying FileName.

        If FileName = "" Then

        Else
            'First check if the RTF file is already open:
            Dim FileFound As Boolean = False
            If RtfDisplayFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To RtfDisplayFormList.Count - 1
                    If RtfDisplayFormList(I) Is Nothing Then

                    Else
                        If RtfDisplayFormList(I).FileName = FileName Then
                            FileFound = True
                            RtfDisplayFormList(I).BringToFront
                            RtfDisplayFormList(I).FindFirstText(SearchText) 'Search for the text
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = NewRtfDisplay()
                RtfDisplayFormList(FormNo).FileName = FileName
                RtfDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                RtfDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                If RtfDisplayFormList(FormNo).FileDirectory = "" Then
                    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                Else
                    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                End If
                RtfDisplayFormList(FormNo).OpenDocument
                RtfDisplayFormList(FormNo).FindFirstText(SearchText) 'Search for the text.
            End If

        End If

    End Sub

    Private Sub FindFirstRtf(ByVal FileName As String, ByVal SearchText As String, ByVal Highlight As Boolean)
        'Find the first instance of the SearchText in the RTF window displaying FileName.
        'A new RTF window is opened if required.
        'Highlight the SearchText when found.

        If FileName = "" Then

        Else
            'First check if the RTF file is already open:
            Dim FileFound As Boolean = False
            If RtfDisplayFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To RtfDisplayFormList.Count - 1
                    If RtfDisplayFormList(I) Is Nothing Then

                    Else
                        If RtfDisplayFormList(I).FileName = FileName Then
                            FileFound = True
                            RtfDisplayFormList(I).BringToFront
                            RtfDisplayFormList(I).FindFirstText(SearchText) 'Search for the text
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = NewRtfDisplay()
                RtfDisplayFormList(FormNo).FileName = FileName
                RtfDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                RtfDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                If RtfDisplayFormList(FormNo).FileDirectory = "" Then
                    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                Else
                    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                End If
                RtfDisplayFormList(FormNo).OpenDocument
                RtfDisplayFormList(FormNo).FindFirstText(SearchText) 'Search for the text.
            End If

        End If

    End Sub

    Private Sub FindTextInRtf(ByVal FileName As String, ByVal SearchText As String, ByVal Highlight As Boolean, ByVal FindFirst As Boolean)
        'Find the SearchText in the RTF window displaying FileName.
        'A new RTF window is opened if required.
        'Highlight the SearchText when found if Highlight is True.
        'Find the first instance of SearchText if FindFirst is True, else find the next instance.

        If FileName = "" Then
            Message.AddWarning("The file name is blank." & vbCrLf)
        Else
            'First check if the RTF file is already open:
            Dim FileFound As Boolean = False
            If RtfDisplayFormList.Count = 0 Then

            Else
                Dim I As Integer
                For I = 0 To RtfDisplayFormList.Count - 1
                    If RtfDisplayFormList(I) Is Nothing Then

                    Else
                        If RtfDisplayFormList(I).FileName = FileName Then
                            FileFound = True
                            RtfDisplayFormList(I).BringToFront
                            'RtfDisplayFormList(I).FindFirstText(SearchText) 'Search for the text
                            RtfDisplayFormList(I).FindText(SearchText, Highlight, FindFirst) 'Search for the text
                        End If
                    End If
                Next
            End If

            If FileFound = False Then
                Dim FormNo As Integer = NewRtfDisplay()
                RtfDisplayFormList(FormNo).FileName = FileName
                RtfDisplayFormList(FormNo).Description = ItemInfo(FileName).Description
                RtfDisplayFormList(FormNo).FileDirectory = ItemInfo(FileName).Directory
                If RtfDisplayFormList(FormNo).FileDirectory = "" Then
                    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.Project
                Else
                    RtfDisplayFormList(FormNo).FileLocationType = LocationTypes.FileSystem
                End If
                RtfDisplayFormList(FormNo).OpenDocument
                'RtfDisplayFormList(FormNo).FindFirstText(SearchText) 'Search for the text.
                RtfDisplayFormList(FormNo).FindText(SearchText, Highlight, FindFirst) 'Search for the text
            End If
        End If
    End Sub

    Private Sub FindTextInMainRtf(ByVal FileName As String, ByVal SearchText As String, ByVal Highlight As Boolean, ByVal FindFirst As Boolean)
        'Find the SearchText in the Document View tab (in the Main window) displaying FileName.
        'The RTF file is opened if required.
        'Highlight the SearchText when found if Highlight is True.
        'Find the first instance of SearchText if FindFirst is True, else find the next instance.

        If FileName = "" Then
            Message.AddWarning("The file name is blank." & vbCrLf)
        Else
            If ItemInfo.ContainsKey(FileName) Then
                If ItemInfo(FileName).Type = "RTF" Then
                    If Me.FileName = FileName Then
                        'The file is already open in the Document View tab.
                        FindTextInRtf(SearchText, Highlight, FindFirst)
                        'Me.BringToFront()
                        'Me.TopMost = True
                        Me.Activate()
                        'Me.Show()
                        'Me.Focus()
                        'Me.Refresh()
                        Me.TopMost = True
                        Me.TopMost = False
                    Else
                        'Open the file in the Document View tab:
                        WebBrowser1.Visible = False
                        AxAcroPDF1.Visible = False
                        XmlHtmDisplay2.Visible = True
                        If ItemInfo(FileName).Directory = "" Then
                            Dim rtbData As New IO.MemoryStream
                            Project.ReadData(FileName, rtbData)
                            XmlHtmDisplay2.Clear()
                            rtbData.Position = 0
                            XmlHtmDisplay2.LoadFile(rtbData, RichTextBoxStreamType.RichText)
                        Else
                            XmlHtmDisplay2.LoadFile(ItemInfo(FileName).Directory & "\" & FileName)
                        End If
                        FindTextInRtf(SearchText, Highlight, FindFirst)
                        'Me.BringToFront()
                        'Me.TopMost = True
                        Me.Activate()
                        'Me.Show()
                        'Me.Focus()
                        'Me.Refresh()
                        Me.TopMost = True
                        Me.TopMost = False
                    End If
                Else
                    Message.AddWarning("The file is not Rich Text Format." & vbCrLf)
                End If
            Else
                Message.AddWarning("The file name is not in the file list." & vbCrLf)
            End If
        End If
    End Sub

    Public Sub FindTextInRtf(ByVal myText As String, ByVal Highlight As Boolean, ByVal FindFirst As Boolean)
        'Find myText in the RichTextBox in the Document View tab (XmlHtmDisplay2).

        Dim StartPos As Integer

        If FindFirst Then
            StartPos = 0
        Else
            StartPos = XmlHtmDisplay1.SelectionStart + 1
        End If

        XmlHtmDisplay2.Focus()

        Dim FoundPos As Integer = XmlHtmDisplay2.Find(myText, StartPos, RichTextBoxFinds.MatchCase)

        If FoundPos < 0 Then
            Message.Add("String not found." & vbCrLf)
        Else
            XmlHtmDisplay2.SelectionStart = FoundPos
            If Highlight Then
                XmlHtmDisplay2.SelectionLength = myText.Length
            Else
                XmlHtmDisplay2.SelectionLength = 0
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem1_EditWorkflowTabPage_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_EditWorkflowTabPage.Click
        'Edit the Workflows Web Page:

        If WorkflowFileName = "" Then
            Message.AddWarning("No page to edit." & vbCrLf)
        Else
            Dim FormNo As Integer = OpenNewWFHtmlDisplayPage()
            WFHtmlDisplayFormList(FormNo).FileName = WorkflowFileName
            WFHtmlDisplayFormList(FormNo).OpenDocument

        End If

        '    Public WFHtmlDisplayFormList As New ArrayList 'Used for displaying multiple HtmlDisplay forms.
    End Sub

    Private Sub ToolStripMenuItem1_ShowStartPageInWorkflowTab_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1_ShowStartPageInWorkflowTab.Click
        'Show the Start Page in the Workflows Tab:
        OpenStartPage()

    End Sub

    Private Sub bgwSendMessage_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessage.DoWork
        'Send a message on a separate thread:
        Try
            If IsNothing(client) Then
                bgwSendMessage.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessage.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParams As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParams.ProjectNetworkName, SendMessageParams.ConnectionName, SendMessageParams.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessage.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessage_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessage.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessage message 
    End Sub

    Private Sub bgwSendMessageAlt_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwSendMessageAlt.DoWork
        'Alternative SendMessage background worker - used to send a message while instructions are being processed. 
        'Send a message on a separate thread
        Try
            If IsNothing(client) Then
                bgwSendMessageAlt.ReportProgress(1, "No Connection available. Message not sent!")
            Else
                If client.State = ServiceModel.CommunicationState.Faulted Then
                    bgwSendMessageAlt.ReportProgress(1, "Connection state is faulted. Message not sent!")
                Else
                    Dim SendMessageParamsAlt As clsSendMessageParams = e.Argument
                    client.SendMessage(SendMessageParamsAlt.ProjectNetworkName, SendMessageParamsAlt.ConnectionName, SendMessageParamsAlt.Message)
                End If
            End If
        Catch ex As Exception
            bgwSendMessageAlt.ReportProgress(1, ex.Message)
        End Try
    End Sub

    Private Sub bgwSendMessageAlt_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwSendMessageAlt.ProgressChanged
        'Display an error message:
        Message.AddWarning("Send Message error: " & e.UserState.ToString & vbCrLf) 'Show the bgwSendMessageAlt message 
    End Sub

    Private Sub XMsg_ErrorMsg(ErrMsg As String) Handles XMsg.ErrorMsg
        Message.AddWarning(ErrMsg & vbCrLf)
    End Sub

    Private Sub XMsgLocal_Instruction(Info As String, Locn As String) Handles XMsgLocal.Instruction

    End Sub

    Private Sub Message_ShowXMessagesChanged(Show As Boolean) Handles Message.ShowXMessagesChanged
        ShowXMessages = Show
    End Sub

    Private Sub Message_ShowSysMessagesChanged(Show As Boolean) Handles Message.ShowSysMessagesChanged
        ShowSysMessages = Show
    End Sub

    'Private Sub btnMove_Click(sender As Object, e As EventArgs) Handles btnMove.Click
    '    Message.MessageForm.Top = Val(txtYLocn.Text)
    '    Message.MessageForm.Left = Val(txtXLocn.Text)
    'End Sub

    'Private Sub chkCheckPosn_Click(sender As Object, e As EventArgs) Handles chkCheckPosn.Click
    '    'Check that the Message form can be seen on a screen.


    '    If Message.MessageForm Is Nothing Then
    '        Message.ApplicationName = ApplicationInfo.Name
    '        Message.SettingsLocn = Project.SettingsLocn
    '        Message.Show()
    '        Message.MessageForm.BringToFront()
    '    End If


    '    Dim MinWidthVisible As Integer = 48 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
    '    Dim MinHeightVisible As Integer = 48 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

    '    Dim FormRect As New Rectangle(Message.MessageForm.Left, Message.MessageForm.Top, Message.MessageForm.Width, Message.MessageForm.Height)
    '    Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

    '    'Check if the top of the form is less than zero:
    '    If Message.MessageForm.Top < 0 Then Message.MessageForm.Top = 0

    '    'Check if the top of the form is too close to the bottom of the Working Area:
    '    If (Message.MessageForm.Top + MinHeightVisible) > (WARect.Top + WARect.Height) Then
    '        Message.MessageForm.Top = WARect.Top + WARect.Height - MinHeightVisible
    '    End If

    '    'Check if the left edge of the form is too close to the right edge of the Working Area:
    '    If (Message.MessageForm.Left + MinWidthVisible) > (WARect.Left + WARect.Width) Then
    '        Message.MessageForm.Left = WARect.Left + WARect.Width - MinWidthVisible
    '    End If

    '    'Check if the right edge of the form is too close to the left edge of the Working Area:
    '    If (Message.MessageForm.Left + Message.MessageForm.Width - MinWidthVisible) < WARect.Left Then
    '        Message.MessageForm.Left = WARect.Left - Message.MessageForm.Width + MinWidthVisible
    '    End If
    'End Sub




    Private Sub Project_NewProjectCreated(ProjectPath As String) Handles Project.NewProjectCreated
        SendProjectInfo(ProjectPath) 'Send the path of the new project to the Network application. THe new project will be added to the list of projects.
    End Sub

    Private Sub bgwComCheck_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwComCheck.DoWork
        'The communications check thread.
        While ConnectedToComnet
            Try
                If client.IsAlive() Then
                    'Message.Add(Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf) 'This produces the error: Cross thread operation not valid.
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection OK." & vbCrLf)
                Else
                    'Message.Add(Format(Now, "HH:mm:ss") & " Connection Fault." & vbCrLf) 'This produces the error: Cross thread operation not valid.
                    bgwComCheck.ReportProgress(1, Format(Now, "HH:mm:ss") & " Connection Fault.")
                End If
            Catch ex As Exception
                bgwComCheck.ReportProgress(1, "Error in bgeComCheck_DoWork!" & vbCrLf)
                bgwComCheck.ReportProgress(1, ex.Message & vbCrLf)
            End Try

            'System.Threading.Thread.Sleep(60000) 'Sleep time in milliseconds (60 seconds) - For testing only.
            System.Threading.Thread.Sleep(1800000) 'Sleep time in milliseconds (30 minutes)
        End While

    End Sub

    Private Sub bgwComCheck_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwComCheck.ProgressChanged
        Message.Add(e.UserState.ToString) 'Show the ComCheck message 
    End Sub

    Private Sub btnShowItemList_Click(sender As Object, e As EventArgs) Handles btnShowItemList.Click
        'Show the list of items in the Message window:

        'Message.Add(vbCrLf & "List of items in the library: " & LibraryName & vbCrLf)
        Message.AddText(vbCrLf & "List of items in the library: " & LibraryName & vbCrLf, "Heading 11pt") 'Message displayed with bold text.
        Dim ItemCount As Integer = ItemInfo.Count
        Message.AddText("Number of items: " & ItemCount & vbCrLf, "Bold")

        For Each item In ItemInfo
            Message.AddText(vbCrLf & "Key: " & item.Key & vbCrLf, "Bold")
            Message.Add("Type: " & item.Value.Type & vbCrLf)
            Message.Add("Creation date: " & item.Value.CreationDate & vbCrLf)
            Message.Add("LastEditDate: " & item.Value.LastEditDate & vbCrLf)
            Message.Add("Description: " & item.Value.Description & vbCrLf)
            Message.Add("Directory: " & item.Value.Directory & vbCrLf)
            Message.Add("FormNo: " & item.Value.FormNo & vbCrLf)
        Next


    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        'This method copies information from the selected node to the Create Item tab.
        'This provides a template to be edited for the new item. 

        Select Case txtNodeType.Text
            Case "Collection"
                rbCollection.Checked = True
            Case "RTF"
                rbRtf.Checked = True
            Case "XML"
                rbXml.Checked = True
            Case "TXT"
                rbTxt.Checked = True
            Case "HTML"
                rbHtml.Checked = True
            Case "PDF"
                rbPdf.Checked = True
            Case "XLS"
                rbXls.Checked = True
            Case "FolderLink"
                rbFolderLink.Checked = True
            Case "XMsg"
                rbXMsg.Checked = True
            Case "XSeq"
                rbXSeq.Checked = True
            Case Else
                Message.AddWarning("Unknown node type: " & txtNodeType.Text & vbCrLf)
        End Select

        txtNewNodeFileName.Text = txtNodeKey.Text
        txtItemText.Text = txtNodeText.Text
        txtNewNodeDescr.Text = txtItemDescription.Text

    End Sub

    Private Sub btnUpdateDocList_Click(sender As Object, e As EventArgs) Handles btnUpdateDocList.Click
        'Update document list.
        UpdateDocList()
    End Sub

    Private Sub UpdateDocList()
        'Update document list.

        dgvDocList.Rows.Clear()

        For Each item In ItemInfo
            dgvDocList.Rows.Add(item.Value.Text, item.Value.Type, item.Value.CreationDate, item.Value.LastEditDate, item.Key)
        Next

        dgvDocList.AutoResizeColumns()
        dgvDocList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgvDocList.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        dgvDocList.Columns(4).Width = 300
        dgvDocList.Sort(dgvDocList.Columns(3), ListSortDirection.Descending)
    End Sub

    Private Sub dgvDocList_SelectionChanged(sender As Object, e As EventArgs) Handles dgvDocList.SelectionChanged

    End Sub

    Private Sub dgvDocList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDocList.CellContentClick

        'Dim RowNo As Integer = e.RowIndex

        'If RowNo = -1 Then
        '    txtDocFileName.Text = ""
        '    txtDocDescr.Text = ""
        'Else
        '    Dim FileName As String = dgvDocList.SelectedRows(0).Cells(4).Value
        '    If ItemInfo.ContainsKey(FileName) Then
        '        txtDocFileName.Text = FileName
        '        txtDocDescr.Text = ItemInfo(FileName).Description
        '    Else
        '        txtDocFileName.Text = FileName
        '        txtDocDescr.Text = ""
        '    End If
        'End If
    End Sub

    Private Sub dgvDocList_DoubleClick(sender As Object, e As EventArgs) Handles dgvDocList.DoubleClick
        OpenDocListItemInNewWindow()
    End Sub

    Private Sub dgvDocList_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDocList.RowEnter
        Dim RowNo As Integer = e.RowIndex

        If RowNo = -1 Then
            txtDocFileName.Text = ""
            txtDocDescr.Text = ""
        Else
            'If dgvDocList.Focused Then
            If dgvDocList.SelectedRows.Count > 0 Then
                Dim FileName As String = dgvDocList.SelectedRows(0).Cells(4).Value
                If ItemInfo.ContainsKey(FileName) Then
                    txtDocFileName.Text = FileName
                    txtDocDescr.Text = ItemInfo(FileName).Description
                Else
                    txtDocFileName.Text = FileName
                    txtDocDescr.Text = ""
                End If
            End If
            'End If
        End If
    End Sub

    Private Sub btnOpenDocInNewWindow2_Click(sender As Object, e As EventArgs) Handles btnOpenDocInNewWindow2.Click
        'Open the document selected in the document list (dgvDocList) in a new window.

        ''Select the node in the tree view:
        'Dim myNodes As TreeNode() = trvLibrary.Nodes.Find(txtDocFileName.Text, True)

        'If myNodes.Count = 0 Then
        '    Message.AddWarning("The document node cound not be found." & vbCrLf)
        'ElseIf myNodes.Count = 1 Then
        '    trvLibrary.SelectedNode = myNodes(0)
        '    OpenDocInNewWindow()
        'Else
        '    Message.AddWarning(myNodes.Count & " matching document nodes found. Te first one will be opened." & vbCrLf)
        '    trvLibrary.SelectedNode = myNodes(0)
        '    OpenDocInNewWindow()
        'End If
        OpenDocListItemInNewWindow()
    End Sub

    Private Sub OpenDocListItemInNewWindow()
        'Select the node in the tree view:
        Dim myNodes As TreeNode() = trvLibrary.Nodes.Find(txtDocFileName.Text, True)

        If myNodes.Count = 0 Then
            Message.AddWarning("The document node cound not be found." & vbCrLf)
        ElseIf myNodes.Count = 1 Then
            trvLibrary.SelectedNode = myNodes(0)
            OpenDocInNewWindow()
        Else
            Message.AddWarning(myNodes.Count & " matching document nodes found. Te first one will be opened." & vbCrLf)
            trvLibrary.SelectedNode = myNodes(0)
            OpenDocInNewWindow()
        End If
    End Sub


#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class clsSendMessageParams
        'Parameters used when sending a message using the Message Service.
        Public ProjectNetworkName As String
        Public ConnectionName As String
        Public Message As String
    End Class

    Private Sub btnAddToCollection_Click(sender As Object, e As EventArgs) Handles btnAddToCollection.Click

    End Sub

    Private Sub btnNewFile_Click(sender As Object, e As EventArgs) Handles btnNewFile.Click

    End Sub

    Private Sub txtNewNodeFileName_TextChanged(sender As Object, e As EventArgs) Handles txtNewNodeFileName.TextChanged

    End Sub


End Class

Public Class clsItemInfo
    'Information about each item in the Library.

    'NOTE: The name is the Key for the ItemInfo dictionary. It does not need to be repeated in the dictionary.
    '      For Rtf, Xml etc document items that are stored in a file, the name is also the FileName. (FileNames cannot be duplicated.)

    'Description   'String
    'CreationDate  'DateTime Format(Now, "d-MMM-yyyy H:mm:ss") 
    'LastEditDate  'DateTime Format(Now, "d-MMM-yyyy H:mm:ss") 
    'Type          'String
    'Directory     'String
    'Left          'Integer
    'Top           'Integer
    'Width         'Integer
    'Height        'Integer
    'FormNo        'Integer

    'The Text property was added 16 September 2020:
    Private _text As String = ""
    Property Text As String
        Get
            Return _text
        End Get
        Set(value As String)
            _text = value
        End Set
    End Property

    Private _description As String = "" 'A description of the item. (Default value is "".)
    Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property

    Private _creationDate As DateTime = Format(Now, "d-MMM-yyyy H:mm:ss") 'The creation date of the item. (Default value is Now.)
    Property CreationDate As DateTime
        Get
            Return _creationDate
        End Get
        Set(value As DateTime)
            _creationDate = value
        End Set
    End Property

    Private _lastEditDate As DateTime = Format(Now, "d-MMM-yyyy H:mm:ss") 'The last edit date of the item. (Default value is Now.)
    Property LastEditDate As DateTime
        Get
            Return _lastEditDate
        End Get
        Set(value As DateTime)
            _lastEditDate = value
        End Set
    End Property

    Private _type As String = "" 'The type of item (Collection, RTF, XML, HTML, TXT, etc).
    Property Type As String
        Get
            Return _type
        End Get
        Set(value As String)
            _type = value
        End Set
    End Property

    Private _directory As String = "" 'The directory in which the item is stored. This is blank if the item is stored in the current project.
    Property Directory As String
        Get
            Return _directory
        End Get
        Set(value As String)
            _directory = value
        End Set
    End Property

    Private _left As Integer = 0 'The Left position of the Document display form.
    Property Left As Integer
        Get
            Return _left
        End Get
        Set(value As Integer)
            _left = value
        End Set
    End Property

    Private _top As Integer = 0 'The Top position of the Document display form.
    Property Top As Integer
        Get
            Return _top
        End Get
        Set(value As Integer)
            _top = value
        End Set
    End Property

    Private _width As Integer = 400 'The Width of the Document display form.
    Property Width As Integer
        Get
            Return _width
        End Get
        Set(value As Integer)
            _width = value
        End Set
    End Property

    Private _height As Integer = 300 'The Height of the Document display form.
    Property Height As Integer
        Get
            Return _height
        End Get
        Set(value As Integer)
            _height = value
        End Set
    End Property

    Private _formNo As Integer = 0 'The Number of the Document display form.
    Property FormNo As Integer
        Get
            Return _formNo
        End Get
        Set(value As Integer)
            _formNo = value
        End Set
    End Property

End Class



