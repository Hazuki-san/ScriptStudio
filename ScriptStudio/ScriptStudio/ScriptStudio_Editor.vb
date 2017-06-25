Imports System.IO

Public Class ScriptStudio_Editor

    Dim ReleaseType As String = "Beta"
    Dim TextBuild As String = "Snapshot"
    Dim VerMajor As String = My.Application.Info.Version.Major.ToString
    Dim VerMinor As String = My.Application.Info.Version.Minor.ToString
    Dim VerMajorRe As String = My.Application.Info.Version.MajorRevision.ToString
    Dim VerMinorRe As String = My.Application.Info.Version.MinorRevision.ToString
    Dim VersionDef As String = My.Application.Info.Version.ToString
    Dim VersionAll As String = VerMajor & "." & VerMinor & "." & VerMajorRe
    Dim VersionCus As String = ReleaseType & " " & VerMajor & "." & VerMinor & "." & VerMajorRe & " " & "(" & TextBuild & " " & VerMinorRe & ")"
    '
    Dim DistroName As String = "Demza Script Studio - " & VersionCus
    Dim ProjectName As String = Nothing

    Private Sub DogHome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ProjectName = "" Then
            Me.Text = DistroName
            RichEditor.Hide()
            RichEditor.ReadOnly = True
            StatusLabel.Text = "Please click, ""File -> New"" to create new Script"
            RichEditor.Text = "Welcome to " & DistroName & "!"
            RichEditor.AppendText(vbNewLine & "To create new Script,")
            RichEditor.AppendText(vbNewLine & "Please click, Files -> New (or press ""CTRL + N"")")
        Else
            Me.Text = DistroName & " - " & ProjectName
            RichEditor.Show()
            RichEditor.ReadOnly = False
            StatusLabel.Text = ProjectName
        End If
        Me.AllowDrop = True
        RichEditor.AllowDrop = True
        VersionX.Text = VersionCus
    End Sub
    Private Sub DogHome_Closing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        If ProjectName = "" Then
            If MessageBox.Show("Are you sure you want to exit ?", "Exit " & DistroName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then

                ' Cancel the Closing event from closing the form.

                e.Cancel = True

            End If
        Else
            If MessageBox.Show("Do you want to save this file before closing ?", "Exit " & DistroName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                SaveToolStripMenuItem.PerformClick()
            End If
        End If
    End Sub

    Private Sub DogHome_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In files
            Dim oReader As StreamReader
            If path.EndsWith(".dsproj") = True OrElse path.EndsWith(".ds") = True OrElse path.EndsWith(".txt") = True Then
                oReader = New StreamReader(path, True)
                RichEditor.Text = oReader.ReadToEnd
                ProjectName = path
                Me.Text = DistroName & " - " & ProjectName
                StatusLabel.Text = ProjectName
                oReader.Close()
                RichEditor.ReadOnly = False
                RichEditor.Show()
            Else
                MessageBox.Show("This format is not supported.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Next
    End Sub

    Private Sub DogHome_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub RichEditor_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles RichEditor.DragDrop
        If MessageBox.Show("Do you want to save this file before opening another files ?", "Dropped!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            SaveToolStripMenuItem.PerformClick()
        End If
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In files
            Dim oReader As StreamReader
            If path.EndsWith(".dsproj") = True OrElse path.EndsWith("ds") = True OrElse path.EndsWith(".txt") OrElse path.EndsWith(".ds") Then
                oReader = New StreamReader(path, True)
                RichEditor.Text = oReader.ReadToEnd
                ProjectName = path
                Me.Text = DistroName & " - " & ProjectName
                StatusLabel.Text = ProjectName
                oReader.Close()
                RichEditor.ReadOnly = False
                RichEditor.Show()
            Else
                MessageBox.Show("This format is not supported.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Next
    End Sub

    Private Sub RichEditor_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles RichEditor.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub RichEditor_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RichEditor.MouseUp
        If e.Button <> Windows.Forms.MouseButtons.Right Then Return
        RC_Sel.Tag = 4
        AddHandler RC_Sel.Click, AddressOf menuChoice
        '-- etc
        '..
        StudioContext.Show(RichEditor, e.Location)
    End Sub

    Private Sub menuChoice(ByVal sender As Object, ByVal e As EventArgs)
        Dim item = CType(sender, ToolStripMenuItem)
        Dim selection = CInt(item.Tag)
        If selection = 4 Then
            RichEditor.SelectAll()
        Else
        End If
        '-- etc
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Dim saveFileDialog1 As New SaveFileDialog()

        saveFileDialog1.Filter = "Demza Script Project (*.dsproj)|*.dsproj|All files (*.*)|*.*"
        saveFileDialog1.FilterIndex = 1
        saveFileDialog1.DefaultExt = "dsproj"
        saveFileDialog1.FileName = ""
        saveFileDialog1.RestoreDirectory = True

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            If ProjectName = "" = False Then
                If MessageBox.Show("Do you want to save this file before opening another files ?", "Dropped!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    SaveToolStripMenuItem.PerformClick()
                End If
            Else
            End If
            If RichEditor.ReadOnly = True Then
                RichEditor.ReadOnly = False
                RichEditor.Show()
            End If
            RichEditor.Text = ""
            System.IO.File.WriteAllText(saveFileDialog1.FileName, RichEditor.Text)
            ProjectName = saveFileDialog1.FileName
            Me.Text = DistroName & " - " & ProjectName
            StatusLabel.Text = ProjectName
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim oReader As StreamReader
        Dim OpenFileDialog1 As New OpenFileDialog()

        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.CheckPathExists = True
        OpenFileDialog1.DefaultExt = "dsproj"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Demza Script Project (*.dsproj)|*.dsproj|All files (*.*)|*.*"
        OpenFileDialog1.Multiselect = False

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            If ProjectName = "" = False Then
                If MessageBox.Show("Do you want to save this file before opening another files ?", "Dropped!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    SaveToolStripMenuItem.PerformClick()
                End If
            Else
            End If
            If RichEditor.ReadOnly = True Then
                RichEditor.ReadOnly = False
                RichEditor.Show()
            End If
            oReader = New StreamReader(OpenFileDialog1.FileName, True)
            RichEditor.Text = oReader.ReadToEnd
            ProjectName = OpenFileDialog1.FileName
            Me.Text = DistroName & " - " & ProjectName
            StatusLabel.Text = ProjectName
            oReader.Close()
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        If RichEditor.ReadOnly = True Then
            MessageBox.Show("To create new file, Please click ""File -> New""", "Bark!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If My.Computer.FileSystem.FileExists(ProjectName) = True Then
                My.Computer.FileSystem.WriteAllText(ProjectName, RichEditor.Text, False)
            ElseIf My.Computer.FileSystem.FileExists(ProjectName) = False Then
                Dim saveFileDialog1 As New SaveFileDialog()

                saveFileDialog1.Filter = "Demza Script Project (*.dsproj)|*.dsproj|All files (*.*)|*.*"
                saveFileDialog1.FilterIndex = 1
                saveFileDialog1.DefaultExt = "dsproj"
                saveFileDialog1.FileName = ""
                saveFileDialog1.RestoreDirectory = True

                If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                    System.IO.File.WriteAllText(saveFileDialog1.FileName, RichEditor.Text)
                    ProjectName = saveFileDialog1.FileName
                    Me.Text = DistroName & " - " & ProjectName
                    StatusLabel.Text = ProjectName
                End If
            End If
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        If RichEditor.ReadOnly = True Then
            MessageBox.Show("To create new file, Please click ""File -> New""", "Bark!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim saveFileDialog1 As New SaveFileDialog()

            saveFileDialog1.Filter = "Demza Script Project (*.dsproj)|*.dsproj|All files (*.*)|*.*"
            saveFileDialog1.FilterIndex = 1
            saveFileDialog1.DefaultExt = "dsproj"
            saveFileDialog1.FileName = ""
            saveFileDialog1.RestoreDirectory = True

            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                System.IO.File.WriteAllText(saveFileDialog1.FileName, RichEditor.Text)
                ProjectName = saveFileDialog1.FileName
                Me.Text = DistroName & " - " & ProjectName
                StatusLabel.Text = ProjectName
            End If
        End If
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        If MessageBox.Show("Printing?", "Developer", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error) = DialogResult.Retry Then
            MessageBox.Show("Save this as *.txt and print, It should be Ok!", "Developer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        If MessageBox.Show("Now, you cannot Preview anything", "Developer", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error) = DialogResult.Retry Then
            MessageBox.Show("Sorry :(", "Developer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        RichEditor.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RedoToolStripMenuItem.Click
        RichEditor.Redo()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        If RichEditor.SelectedText <> "" Then
            Clipboard.SetText(RichEditor.SelectedText)
            RichEditor.SelectedText = ""
        Else
        End If
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        If RichEditor.SelectedText <> "" Then
            Clipboard.SetText(RichEditor.SelectedText)
        Else
        End If
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        Dim iData As IDataObject = Clipboard.GetDataObject()
        If iData.GetDataPresent(DataFormats.Text) Then
            RichEditor.SelectedText = CType(iData.GetData(DataFormats.Text), String)
        Else
        End If
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        RichEditor.SelectAll()
    End Sub

    Private Sub CloseProjectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseProjectToolStripMenuItem.Click
        If ProjectName = "" = False Then
            If MessageBox.Show("Do you want to save this file before closing ?", "Exit " & DistroName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                SaveToolStripMenuItem.PerformClick()
            End If
        End If
        ProjectName = ""
        Me.Text = DistroName
        RichEditor.Hide()
        RichEditor.ReadOnly = True
        StatusLabel.Text = "Please click, ""File -> New"" to create new Script"
        RichEditor.Text = "Welcome to " & DistroName & "!"
        RichEditor.AppendText(vbNewLine & "To create new Script,")
        RichEditor.AppendText(vbNewLine & "Please click, Files -> New (or press ""CTRL + N"")")
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        ScriptStudio_About.Show()
    End Sub

    Private Sub DogcryptToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DogcryptToolStripMenuItem.Click
        ScriptStudio_Cryptor.Show()
    End Sub

    Private Sub CustomizeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomizeToolStripMenuItem.Click
        If MessageBox.Show("Customize this app?", "Developer", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error) = DialogResult.Retry Then
            MessageBox.Show("NOPE!", "Developer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        If MessageBox.Show("Too many options!", "Developer", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error) = DialogResult.Retry Then
            MessageBox.Show("Sorry :(", "Developer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub ContentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContentsToolStripMenuItem.Click
        If MessageBox.Show("Think creative!", "Developer", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error) = DialogResult.Retry Then
            MessageBox.Show("To make your contents!!!", "Developer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub IndexToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IndexToolStripMenuItem.Click
        If MessageBox.Show("How about index.html ?", "Developer", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error) = DialogResult.Retry Then
            MessageBox.Show("Don't like this joke?", "Developer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub SearchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SearchToolStripMenuItem.Click
        If MessageBox.Show("Google It!", "Developer", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error) = DialogResult.Retry Then
            MessageBox.Show("What you can get?", "Developer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class
