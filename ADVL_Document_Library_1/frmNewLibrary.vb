﻿Public Class frmNewLibrary
    'This form is used to create a new Document Library.
    'The Library is created in the current project.

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


    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        'Create a new Document Library.

        'Check new library file name:
        txtFileName.Text = Trim(txtFileName.Text) 'Remove leading and trailing spaces.

        If txtFileName.Text.EndsWith(".DocLib") Then
            txtFileName.Text = txtFileName.Text.Replace(" ", "_")
        Else
            txtFileName.Text = txtFileName.Text.Replace(" ", "_") & ".DocLib"
        End If

        'Check if the new library file name already exists:
        If Main.Project.DataFileExists(txtFileName.Text) = True Then
            Main.Message.AddWarning("The library file already exists: " & txtFileName.Text & vbCrLf)
            Exit Sub
        End If

        'Check the library name (label):
        txtLibraryLabel.Text = Trim(txtLibraryLabel.Text)
        If txtLibraryLabel.Text = "" Then
            Main.Message.AddWarning("The new library name is blank." & vbCrLf)
            Exit Sub
        End If

        Main.SaveLibrary() 'Save the current library.
        Main.trvLibrary.Nodes.Clear() 'Clear the Tree View
        Dim Node1 As TreeNode = New TreeNode(txtLibraryLabel.Text, 0, 1)
        Node1.Name = txtFileName.Text
        Main.trvLibrary.Nodes.Add(Node1)

        Main.LibraryName = txtLibraryLabel.Text
        Main.LibraryFileName = txtFileName.Text
        Main.LibraryDescription = txtDescription.Text
        Main.LibraryCreationDate = Now
        Main.LibraryLastEditDate = Now

        Main.SaveLibrary()

    End Sub


#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Region " Form Events - Events that can be triggered by this form." '==========================================================================================================================
#End Region 'Form Events ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

End Class