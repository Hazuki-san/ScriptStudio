Imports System.IO

Public Class DogHome

    Dim DogTitle = ""

    Private Sub DogHome_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If DogTitle = "" Then
            Me.Text = "Dogcument"
            RichEditor.Hide()
            RichEditor.ReadOnly = True
            StatusLabel.Text = "Please click, ""File -> New"" to create new Dogcument"
            RichEditor.Text = "Welcome to Dogcument!"
            RichEditor.AppendText(vbNewLine & "To create new Dogcument,")
            RichEditor.AppendText(vbNewLine & "Please click, Files -> New (or press ""CTRL + N"")")
        Else
            Me.Text = "Dogcument - " & DogTitle
            RichEditor.Show()
            RichEditor.ReadOnly = False
            StatusLabel.Text = DogTitle
        End If
        Me.AllowDrop = True
        RichEditor.AllowDrop = True
    End Sub
    Private Sub DogHome_Closing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles MyBase.FormClosing
        If DogTitle = "" Then
            If MessageBox.Show("Are you sure you want to exit ?", "Exit Dogcument", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then

                ' Cancel the Closing event from closing the form.

                e.Cancel = True

            End If
        Else
            If MessageBox.Show("Do you want to save this file before closing ?", "Exit Dogcument", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                SaveToolStripMenuItem.PerformClick()
            End If
        End If
    End Sub

    Private Sub DogHome_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In files
            Dim oReader As StreamReader
            If path.EndsWith(".dog") = True OrElse path.EndsWith(".txt") = True Then
                oReader = New StreamReader(path, True)
                RichEditor.Text = oReader.ReadToEnd
                DogTitle = path
                Me.Text = "Dogcument - " & DogTitle
                StatusLabel.Text = DogTitle
                oReader.Close()
                RichEditor.ReadOnly = False
                RichEditor.Show()
            ElseIf path.EndsWith(".dogapp") = True OrElse path.EndsWith(".ds") OrElse path.EndsWith(".scriptx") = True Then
                Dim dogveri As String
                Dim apptype As String
                Dim dogapp As String
                Dim dogarg1 As String
                Dim dogend As String
                Dim lines() As String = IO.File.ReadAllLines(path)
                dogveri = lines(0)
                apptype = lines(1)
                dogapp = lines(2)
                dogarg1 = lines(3)
                dogend = lines(lines.Length - 1)
                'dog
                If dogarg1 = "" Then
                    MsgBox("Error: dogarg1 cannot be null.", MsgBoxStyle.Critical, "Failed to start " & apptype & " on the Dogscript Host")
                ElseIf dogveri = "[DogApp]" And apptype.StartsWith("Apptype: ") And apptype.EndsWith("dogapp.popup") And dogend.StartsWith("[DogEnd]") Then
                    MsgBox(dogapp, MsgBoxStyle.OkOnly, "Dogscript Host")
                ElseIf dogveri = "[DogApp]" And apptype.StartsWith("Apptype: ") And apptype.EndsWith("dogapp.popup.advanced") And dogend.StartsWith("[DogEnd]") Then
                    MessageBox.Show(dogapp, dogarg1)
                ElseIf dogveri = "[DogApp]" And apptype.StartsWith("Apptype: ") And apptype.EndsWith("dogapp.startshell") And dogend.StartsWith("[DogEnd]") Then
                    Process.Start(dogapp)
                ElseIf path.EndsWith(".ds") And dogveri = "[Dogscript.Start]" And apptype.StartsWith("Apptype: ") And apptype.EndsWith("dogscript.script.directaccess") And dogend.StartsWith("[Dogscript.End]") Then
                    If dogapp = "dogcument.stop" Then
                    ElseIf dogapp = "dogcument.error" OrElse dogapp = "dogcument.bark" And dogarg1.StartsWith("enabled: ") And dogarg1.EndsWith("true") Then
                        MessageBox.Show("Error: Cursed by Dogscript", "Bark!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                ElseIf path.EndsWith(".ds") OrElse path.EndsWith(".scriptx") And dogveri = "[Dogscript.Start]" And apptype.StartsWith("Apptype: ") And apptype.EndsWith("x.dogapp.popup.advanced") And dogend.StartsWith("[Dogscript.End]") Then
                    MessageBox.Show(dogapp, "ScriptX - " & dogarg1)
                ElseIf path.EndsWith(".dogapp") And dogveri = "[Dogscript.Start]" Then
                    MsgBox("Error: DogApp is cannot execute the Dogscript file!", MsgBoxStyle.Critical, "Failed to start " & apptype & " on the Dogscript Host")
                Else
                    MsgBox("Error: Script Error, Please Check your script again.", MsgBoxStyle.Critical, "Failed to start " & apptype & " on the Dogscript Host")
                End If
            Else
                MessageBox.Show("What is this?", "Bark!", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            If path.EndsWith(".dog") = True OrElse path.EndsWith(".txt") Then
                oReader = New StreamReader(path, True)
                RichEditor.Text = oReader.ReadToEnd
                DogTitle = path
                Me.Text = "Dogcument - " & DogTitle
                StatusLabel.Text = DogTitle
                oReader.Close()
                RichEditor.ReadOnly = False
                RichEditor.Show()
            Else
                MessageBox.Show("What is this?", "Bark!", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
        DogContext.Show(RichEditor, e.Location)
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
        If RichEditor.ReadOnly = True = True Then
            RichEditor.ReadOnly = False
            RichEditor.Show()
            NewToolStripMenuItem.PerformClick()
        Else
            Dim myStream As Stream
            Dim saveFileDialog1 As New SaveFileDialog()

            saveFileDialog1.Filter = "Dogcument (*.dog)|*.dog|DogApp (*.dogapp)|*.dogapp|Dogscript (*.ds)|*.ds|All files (*.*)|*.*"
            saveFileDialog1.FilterIndex = 1
            saveFileDialog1.DefaultExt = "dog"
            saveFileDialog1.FileName = ""
            saveFileDialog1.RestoreDirectory = True

            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                RichEditor.Text = ""
                System.IO.File.WriteAllText(saveFileDialog1.FileName, RichEditor.Text)
                DogTitle = saveFileDialog1.FileName
                Me.Text = "Dogcument - " & DogTitle
                StatusLabel.Text = DogTitle
            End If
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        If RichEditor.ReadOnly = True = True Then
            RichEditor.ReadOnly = False
            RichEditor.Show()
            OpenToolStripMenuItem.PerformClick()
        Else
            Dim oReader As StreamReader
            Dim OpenFileDialog1 As New OpenFileDialog()

            OpenFileDialog1.CheckFileExists = True
            OpenFileDialog1.CheckPathExists = True
            OpenFileDialog1.DefaultExt = "dog"
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Dogcument (*.dog)|*.dog|DogApp (*.dogapp)|*.dogapp|Dogscript (*.ds)|*.ds|All files (*.*)|*.*"
            OpenFileDialog1.Multiselect = False

            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                oReader = New StreamReader(OpenFileDialog1.FileName, True)
                RichEditor.Text = oReader.ReadToEnd
                DogTitle = OpenFileDialog1.FileName
                Me.Text = "Dogcument - " & DogTitle
                StatusLabel.Text = DogTitle
                oReader.Close()
            End If
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        If RichEditor.ReadOnly = True Then
            MessageBox.Show("To create new file, Please click ""File -> New""", "Bark!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If My.Computer.FileSystem.FileExists(DogTitle) = True Then
                My.Computer.FileSystem.WriteAllText(DogTitle, RichEditor.Text, False)
            ElseIf My.Computer.FileSystem.FileExists(DogTitle) = False Then
                Dim myStream As Stream
                Dim saveFileDialog1 As New SaveFileDialog()

                saveFileDialog1.Filter = "Dogcument (*.dog)|*.dog|DogApp (*.dogapp)|*.dogapp|Dogscript (*.ds)|*.ds|All files (*.*)|*.*"
                saveFileDialog1.FilterIndex = 1
                saveFileDialog1.DefaultExt = "dog"
                saveFileDialog1.FileName = ""
                saveFileDialog1.RestoreDirectory = True

                If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                    System.IO.File.WriteAllText(saveFileDialog1.FileName, RichEditor.Text)
                    DogTitle = saveFileDialog1.FileName
                    Me.Text = "Dogcument - " & DogTitle
                    StatusLabel.Text = DogTitle
                End If
            End If
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        If RichEditor.ReadOnly = True Then
            MessageBox.Show("To create new file, Please click ""File -> New""", "Bark!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim myStream As Stream
            Dim saveFileDialog1 As New SaveFileDialog()

            saveFileDialog1.Filter = "Dogcument (*.dog)|*.dog|DogApp (*.dogapp)|*.dogapp|Dogscript (*.ds)|*.ds|All files (*.*)|*.*"
            saveFileDialog1.FilterIndex = 1
            saveFileDialog1.DefaultExt = "dog"
            saveFileDialog1.FileName = ""
            saveFileDialog1.RestoreDirectory = True

            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                System.IO.File.WriteAllText(saveFileDialog1.FileName, RichEditor.Text)
                DogTitle = saveFileDialog1.FileName
                Me.Text = "Dogcument - " & DogTitle
                StatusLabel.Text = DogTitle
            End If
        End If
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click

    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click

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
        DogTitle = ""
        Me.Text = "Dogcument"
        RichEditor.Hide()
        RichEditor.ReadOnly = True
        StatusLabel.Text = "Please click, ""File -> New"" to create new Dogcument"
        RichEditor.Text = "Welcome to Dogcument!"
        RichEditor.AppendText(vbNewLine & "To create new Dogcument,")
        RichEditor.AppendText(vbNewLine & "Please click, Files -> New (or press ""CTRL + N"")")
    End Sub
End Class
