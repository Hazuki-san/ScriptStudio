Public NotInheritable Class ScriptStudio_About

    Private Sub AboutDog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Set the title of the form.
        Dim ApplicationTitle As String
        If My.Application.Info.Title <> "" Then
            ApplicationTitle = My.Application.Info.Title
        Else
            ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If
        Me.Text = String.Format("About {0}", ApplicationTitle)
        ' Initialize all of the text displayed on the About Box.
        ' TODO: Customize the application's assembly information in the "Application" pane of the project 
        '    properties dialog (under the "Project" menu).
        Me.LabelProductName.Text = My.Application.Info.ProductName
        Me.LabelVersion.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = My.Application.Info.CompanyName
        Me.TextBoxDescription.Text = My.Application.Info.Description & vbNewLine & vbNewLine & "(Thanks ""Thad Van den Bosch"" for the Encrypt/Decrypt code!" & vbNewLine & "https://www.codeproject.com/Articles/12092/Encrypt-Decrypt-Files-in-VB-NET-Using-Rijndael" & ")" & vbNewLine & vbNewLine & "Icon by MannMitDerTarnjacke (" & "http://mannmitdertarnjacke.deviantart.com/art/Aurora-287069107" & ")"
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub
End Class
