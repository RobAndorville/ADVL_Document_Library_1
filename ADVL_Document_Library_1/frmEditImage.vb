Public Class frmEditImage
    'This form is used to edit images for a rtf document.

#Region " Variable Declarations - All the variables used in this form and this application." '=================================================================================================

#End Region 'Variable Declarations ------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Properties - All the properties used in this form and this application" '============================================================================================================

    Private _sizeMode As String = "Normal" 'The SizeMode property of the PictureBox1 (Normal, StretchImage, AutoSize, CenterImage, Zoom).
    Property SizeMode As String
        Get
            Return _sizeMode
        End Get
        Set(value As String)
            _sizeMode = value
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
                               <!---->
                               <SizeMode><%= SizeMode %></SizeMode>
                               <ResizePercent><%= txtResize.Text %></ResizePercent>
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
            If Settings.<FormSettings>.<SizeMode>.Value <> Nothing Then
                SizeMode = Settings.<FormSettings>.<SizeMode>.Value
                Select Case SizeMode
                    Case "Normal"
                        rbNormal.Checked = True
                    Case "StretchImage"
                        rbStretchImage.Checked = True
                    Case "AutoSize"
                        rbAutoSize.Checked = True
                    Case "CenterImage"
                        rbCenterImage.Checked = True
                    Case "Zoom"
                        rbZoom.Checked = True
                    Case Else
                        SizeMode = "Normal"
                        rbNormal.Checked = True
                End Select
            Else
                SizeMode = "Normal"
                rbNormal.Checked = True
            End If

            If Settings.<FormSettings>.<ResizePercent>.Value <> Nothing Then
                txtResize.Text = Settings.<FormSettings>.<ResizePercent>.Value
            Else
                txtResize.Text = "60"
            End If

            CheckFormPos()
        End If
    End Sub

    Private Sub CheckFormPos()
        'Check that the form can be seen on a screen.

        Dim MinWidthVisible As Integer = 192 'Minimum number of X pixels visible. The form will be moved if this many form pixels are not visible.
        Dim MinHeightVisible As Integer = 64 'Minimum number of Y pixels visible. The form will be moved if this many form pixels are not visible.

        Dim FormRect As New Rectangle(Me.Left, Me.Top, Me.Width, Me.Height)
        Dim WARect As Rectangle = Screen.GetWorkingArea(FormRect) 'The Working Area rectangle - the usable area of the screen containing the form.

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

        'Set defaults: (These may be updated by RestoreFormSettings())
        rbNormal.Checked = True
        txtResize.Text = 60

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

    Private Sub btnPaste_Click(sender As Object, e As EventArgs) Handles btnPaste.Click
        'Paste the image in the Clipboard to PictureBox1
        If Clipboard.ContainsImage Then
            PictureBox1.Image = Clipboard.GetImage
        ElseIf Clipboard.ContainsFileDropList Then
            Try
                'PictureBox1.Image = Clipboard.GetData(DataFormats.Bitmap)
                PictureBox1.Image = Image.FromFile(Clipboard.GetFileDropList(0))
            Catch ex As Exception
                Main.Message.AddWarning(ex.Message & vbCrLf)
            End Try
        End If

        Exit Sub 'Skip the messages (only display for debugging)

        Main.Message.Add("Clipboard.ContainsAudio: " & Clipboard.ContainsAudio & vbCrLf)

        Main.Message.Add("Clipboard.ContainsData(DataFormats.Bitmap): " & Clipboard.ContainsData(DataFormats.Bitmap) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.CommaSeparatedValue): " & Clipboard.ContainsData(DataFormats.CommaSeparatedValue) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Dib): " & Clipboard.ContainsData(DataFormats.Dib) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Dif): " & Clipboard.ContainsData(DataFormats.Dif) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.EnhancedMetafile): " & Clipboard.ContainsData(DataFormats.EnhancedMetafile) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.FileDrop): " & Clipboard.ContainsData(DataFormats.FileDrop) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Html): " & Clipboard.ContainsData(DataFormats.Html) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Locale): " & Clipboard.ContainsData(DataFormats.Locale) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.MetafilePict): " & Clipboard.ContainsData(DataFormats.MetafilePict) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.OemText): " & Clipboard.ContainsData(DataFormats.OemText) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Palette): " & Clipboard.ContainsData(DataFormats.Palette) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.PenData): " & Clipboard.ContainsData(DataFormats.PenData) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Riff): " & Clipboard.ContainsData(DataFormats.Riff) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Rtf): " & Clipboard.ContainsData(DataFormats.Rtf) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Serializable): " & Clipboard.ContainsData(DataFormats.Serializable) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.StringFormat): " & Clipboard.ContainsData(DataFormats.StringFormat) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.SymbolicLink): " & Clipboard.ContainsData(DataFormats.SymbolicLink) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Text): " & Clipboard.ContainsData(DataFormats.Text) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.Tiff): " & Clipboard.ContainsData(DataFormats.Tiff) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.UnicodeText): " & Clipboard.ContainsData(DataFormats.UnicodeText) & vbCrLf)
        Main.Message.Add("Clipboard.ContainsData(DataFormats.WaveAudio): " & Clipboard.ContainsData(DataFormats.WaveAudio) & vbCrLf)

        Main.Message.Add("Clipboard.ContainsFileDropList: " & Clipboard.ContainsFileDropList & vbCrLf)
        Main.Message.Add("Clipboard.ContainsImage: " & Clipboard.ContainsImage & vbCrLf)
        Main.Message.Add("Clipboard.ContainsText: " & Clipboard.ContainsText & vbCrLf)



    End Sub

    Private Sub btnResize_Click(sender As Object, e As EventArgs) Handles btnResize.Click
        ResizeV2()
    End Sub

    Private Sub ResizeV1()
        'Resize the image
        Dim Scale As Single
        Scale = Val(txtResize.Text) / 100
        If Scale < 0 Then
            Main.Message.AddWarning("Scale is too negative.")
            If Scale = 0 Then
                Main.Message.AddWarning("Scale is too zero.")
            End If
        ElseIf Scale > 1000 Then
            Main.Message.AddWarning("Scale is too large.")
        Else
            'Dim OldWidth As Integer = PictureBox1.Width
            'Dim OldHeight As Integer = PictureBox1.Height
            Dim OldWidth As Integer = PictureBox1.ClientSize.Width
            Dim OldHeight As Integer = PictureBox1.ClientSize.Height
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            Dim NewSize As Drawing.Size
            NewSize.Width = Int(OldWidth * Scale)
            NewSize.Height = Int(OldHeight * Scale)
            'PictureBox1.ClientSize.Width = Int(OldWidth * Scale)
            PictureBox1.ClientSize = NewSize

        End If
    End Sub

    Private Sub ResizeV2()
        'Resize the image in PictureBox1:

        Dim myImage As Image
        'myImage = PictureBox1.Image

        Dim ResizePercent As Single = Val(txtResize.Text)
        Dim ResizeWidth As Integer = Int(PictureBox1.Image.Width * ResizePercent / 100)
        Dim ResizeHeight As Integer = Int(PictureBox1.Image.Height * ResizePercent / 100)

        myImage = PictureBox1.Image.GetThumbnailImage(ResizeWidth, ResizeHeight, Nothing, IntPtr.Zero)

        PictureBox1.Image = myImage

    End Sub

    Private Sub btnShowImageSettings_Click(sender As Object, e As EventArgs) Handles btnShowImageSettings.Click
        'Show the Image settings in PictureBox1:

        Main.Message.Add("PictureBox1 settings:" & vbCrLf)
        Main.Message.Add("PictureBox1.ClientSize.Height:" & PictureBox1.ClientSize.Height & vbCrLf)
        Main.Message.Add("PictureBox1.ClientSize.Width:" & PictureBox1.ClientSize.Width & vbCrLf)
        Main.Message.Add("PictureBox1.Height:" & PictureBox1.Height & vbCrLf)
        Main.Message.Add("PictureBox1.Width:" & PictureBox1.Width & vbCrLf)
        Main.Message.Add("PictureBox1.Bottom:" & PictureBox1.Bottom & vbCrLf)
        Main.Message.Add("PictureBox1.Left:" & PictureBox1.Left & vbCrLf)
        Main.Message.Add("PictureBox1.Right:" & PictureBox1.Right & vbCrLf)
        Main.Message.Add("PictureBox1.Top:" & PictureBox1.Top & vbCrLf)
        Main.Message.Add("PictureBox1.SizeMode.ToString :" & PictureBox1.SizeMode.ToString & vbCrLf)
        Main.Message.Add(vbCrLf)

    End Sub

    Private Sub rbNormal_CheckedChanged(sender As Object, e As EventArgs) Handles rbNormal.CheckedChanged
        If rbNormal.Checked Then PictureBox1.SizeMode = PictureBoxSizeMode.Normal
        SizeMode = "Normal"
    End Sub

    Private Sub rbStretchImage_CheckedChanged(sender As Object, e As EventArgs) Handles rbStretchImage.CheckedChanged
        If rbStretchImage.Checked Then PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        SizeMode = "StretchImage"
    End Sub

    Private Sub rbAutoSize_CheckedChanged(sender As Object, e As EventArgs) Handles rbAutoSize.CheckedChanged
        If rbAutoSize.Checked Then PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize
        SizeMode = "AutoSize"
    End Sub

    Private Sub rbCenterImage_CheckedChanged(sender As Object, e As EventArgs) Handles rbCenterImage.CheckedChanged
        If rbCenterImage.Checked Then PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
        SizeMode = "CenterImage"
    End Sub

    Private Sub rbZoom_CheckedChanged(sender As Object, e As EventArgs) Handles rbZoom.CheckedChanged
        If rbZoom.Checked Then PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        SizeMode = "Zoom"
    End Sub

    Private Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click
        RaiseEvent InsertImage(PictureBox1.Image)
    End Sub

#End Region 'Form Display Methods -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Open and Close Forms - Code used to open and close other forms." '===================================================================================================================

#End Region 'Open and Close Forms -------------------------------------------------------------------------------------------------------------------------------------------------------------


#Region " Form Methods - The main actions performed by this form." '===========================================================================================================================

#End Region 'Form Methods ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Event InsertImage(ByRef myImage As Image)

End Class