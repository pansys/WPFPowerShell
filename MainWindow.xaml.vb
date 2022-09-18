Imports System.Collections.ObjectModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Text

Class MainWindow

    Dim powershell As PowerShell = powershell.Create()

    Dim prevCommands As List(Of String) = New List(Of String)()
    Dim prevCommandsPos As Integer = -1
    Dim prevKey As Key

    Dim shellPrompt As String = ""

    Public Sub New()
        InitializeComponent()

        getWorkingDir()

        tbCommands.Text = shellPrompt
    End Sub

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        tbCommands.Text = "WPF PS> "
        tbCommands.CaretIndex = tbCommands.Text.Length
        tbCommands.Focus()
    End Sub

    Private Sub ExecuteCommand()
        Debug.Print("WPFPowerShellSample : Executing PowerShell Command...")
        If (tbCommands.Text.Length > shellPrompt.Length) Then

            Dim lastIndex As Int16 = tbCommands.Text.LastIndexOf(shellPrompt) + shellPrompt.Length
            Dim lastLine As String = tbCommands.Text.Substring(lastIndex, tbCommands.Text.Length - lastIndex)

            Dim command As String = lastLine.Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, "").Replace("WPF PS> ", "")

            If (command.Length > 0) Then
                powershell.AddScript(command)
                'powershell.AddScript("1..5 | % { Write-Progress -Activity 'My Important Activity' -PercentComplete ($PSItem*20) -Status 'This is my status'; Start-Sleep -Milliseconds 200; }")
                powershell.AddCommand("Out-String")
                AddHandler powershell.Streams.Progress.DataAdded, AddressOf ProcessOutput

                Try
                    Dim results = powershell.Invoke()
                    If (results.Count > 0) Then
                        tbCommands.AppendText(vbLf)
                    End If

                    For Each item As PSObject In results
                        tbCommands.AppendText(item.ToString)
                    Next

                    prevCommands.Add(command)
                    prevCommandsPos = prevCommands.Count - 1

                Catch ex As Exception
                    tbCommands.AppendText(ex.Message)
                End Try

            End If

            getWorkingDir()
        End If
    End Sub

    Private Sub getWorkingDir()
        powershell.AddScript("pwd")
        powershell.AddCommand("Out-String")

        Dim strOut As String = ""
        Try
            Dim results = powershell.Invoke()
            For Each item As PSObject In results
                strOut = item.ToString
            Next

            shellPrompt = "WPF PS " + strOut.Replace("Path", "").Replace("----", "").Replace(vbCrLf, "") + "> "
        Catch ex As Exception
            'tbCommands.AppendText(ex.Message)
        End Try
    End Sub
    Private Sub ProcessOutput(sender As Object, data As DataAddedEventArgs)
        For Each item As ProgressRecord In sender
            'tbOutput.AppendText(item.ToString)
            Debug.Print("WPFPowerShellSample : " & item.ToString)
        Next
    End Sub

    Private Sub tbCommands_KeyDown(sender As Object, e As KeyEventArgs) Handles tbCommands.KeyDown
        If e.Key = Key.[Return] Then
            If prevKey = Key.LeftCtrl Or prevKey = Key.RightCtrl Then
                tbCommands.AppendText(vbLf)
                tbCommands.CaretIndex = tbCommands.Text.Length
                tbCommands.Focus()
            Else
                'tbCommands.AppendText(vbLf)
                'tbCommands.AppendText("WPF PS > ")
                ExecuteCommand()
                tbCommands.AppendText(vbLf)
                'tbCommands.AppendText(powershell.HistoryString)
                tbCommands.AppendText(shellPrompt)
                tbCommands.CaretIndex = tbCommands.Text.Length
                tbCommands.Focus()
            End If
        ElseIf e.Key = Key.Back Then

        End If

        prevKey = e.Key

    End Sub

    Private Sub tbCommands_GotFocus(sender As Object, e As RoutedEventArgs) Handles tbCommands.GotFocus
        tbCommands.CaretIndex = tbCommands.Text.Length
        tbCommands.Focus()
    End Sub

    Private Sub tbCommands_GotMouseCapture(sender As Object, e As MouseEventArgs) Handles tbCommands.GotMouseCapture
        tbCommands.CaretIndex = tbCommands.Text.Length
        tbCommands.Focus()
    End Sub

    Private Sub tbCommands_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles tbCommands.PreviewKeyDown
        If e.Key = Key.Up Then
            e.Handled = True

            Dim validPart = tbCommands.Text.Substring(0, tbCommands.Text.LastIndexOf("> ") + 2)
            If prevCommandsPos > -1 Then
                Dim prevCommand As String = prevCommands(prevCommandsPos)
                tbCommands.Text = validPart & prevCommand
                prevCommandsPos = prevCommandsPos - 1

                tbCommands.CaretIndex = tbCommands.Text.Length
                tbCommands.Focus()
            End If
        ElseIf e.Key = Key.Down Then
            e.Handled = True

            Dim validPart = tbCommands.Text.Substring(0, tbCommands.Text.LastIndexOf("> ") + 2)

            If prevCommandsPos < (prevCommands.Count - 1) Then
                If prevCommandsPos = -1 Then
                    prevCommandsPos = 0
                End If

                prevCommandsPos = prevCommandsPos + 1
                Dim prevCommand As String = prevCommands(prevCommandsPos)
                tbCommands.Text = validPart & prevCommand

                tbCommands.CaretIndex = tbCommands.Text.Length
                tbCommands.Focus()
            End If
        ElseIf e.Key = Key.Back Or e.Key = Key.Left Then
            Dim p As Int16 = tbCommands.SelectionStart
            If (tbCommands.Text.Substring(tbCommands.SelectionStart - 2, 2) = "> ") Then
                e.Handled = True
            End If
        End If
    End Sub
End Class

