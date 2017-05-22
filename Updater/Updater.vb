Imports Ionic.Zip
Imports System.ComponentModel

Public Class Updater

    Private WithEvents _worker As New BackgroundWorker

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value < ProgressBar1.Maximum Then
            ProgressBar1.Value += 1
        Else
            ProgressBar1.Value = ProgressBar1.Minimum
        End If
    End Sub

    Private Sub start()
        Dim p As List(Of Process) = Process.GetProcessesByName("VMM").ToList
        If p.Count > 0 Then
            Debug.Print("wait!")
        Else
            If Not _worker.IsBusy Then
                Timer2.Enabled = False
                _worker.RunWorkerAsync()
            End If
        End If
    End Sub

    Private Sub _worker_DoWork(sender As Object, e As DoWorkEventArgs) Handles _worker.DoWork
        Dim CommandLine As List(Of String) = Environment.GetCommandLineArgs.ToList

        If CommandLine.Count >= 2 Then
            Dim Path As String = CommandLine(1)
            If My.Computer.FileSystem.FileExists(Path) Then
                Dim ModManagerFolder As String = Application.StartupPath
                Try
                    Using zip As ZipFile = ZipFile.Read(Path)
                        For Each entry As ZipEntry In zip.Entries
                            Dim File As String = String.Format("{0}\{1}", ModManagerFolder, entry.FileName)
                            entry.Extract(ModManagerFolder, ExtractExistingFileAction.DoNotOverwrite)
                            If Not My.Computer.FileSystem.FileExists(File) Then
                                entry.Extract(ModManagerFolder)
                            Else
                                Try
                                    My.Computer.FileSystem.DeleteFile(File)
                                    If Not My.Computer.FileSystem.FileExists(File) Then
                                        entry.Extract(ModManagerFolder)
                                    End If
                                Catch ex As Exception
                                    Debug.Print(ex.Message)
                                End Try
                            End If
                        Next
                        'zip.ExtractAll(ModManagerFolder, ExtractExistingFileAction.OverwriteSilently)
                    End Using
                Catch ex As Exception
                    Dim Message As String = String.Format("An error occured when extracting the update:{0}{1}", vbCrLf + vbCrLf, ex.Message)
                    MsgBox(Message, MsgBoxStyle.OkOnly, "Error")
                End Try
            End If
        End If
    End Sub

    Private Sub _worker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles _worker.RunWorkerCompleted
        Dim VMM As String = String.Format("{0}{1}", Application.StartupPath, "\VMM.exe")
        Process.Start(VMM)
        Application.Exit()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        start()
    End Sub

End Class
