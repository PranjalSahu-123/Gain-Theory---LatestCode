Imports System
Imports System.Windows.Forms
Imports System.IO

Public Class ChangeLogDir

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub ChangeLogDir_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            TextBox1.Text = Globals.Ribbons.MSprintExRibbon.LogDirectoryPath
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim planFolderPath As FolderBrowserDialog = New FolderBrowserDialog()

     
        If planFolderPath.ShowDialog = DialogResult.OK Then
            Try
                Dim path As String = planFolderPath.SelectedPath

                    If Not (IO.Directory.Exists(path + "\Logs")) Then
                        IO.Directory.CreateDirectory(path + "\Logs")
                    End If
                    Globals.Ribbons.MSprintExRibbon.LogDirectoryPath = path + "\Logs\"
                    Globals.Ribbons.MSprintExRibbon.logDirectoryXML = <LogDirPath></LogDirPath>
                    Globals.Ribbons.MSprintExRibbon.logDirectoryXML.Value = Globals.Ribbons.MSprintExRibbon.LogDirectoryPath
                    Globals.Ribbons.MSprintExRibbon.logDirectoryXML.Save(Globals.Ribbons.MSprintExRibbon.MasterFolderPath + "\LogDirecPath")
                    TextBox1.Text = Globals.Ribbons.MSprintExRibbon.LogDirectoryPath
            Catch ex As Exception
                    If Not (Directory.Exists(Globals.Ribbons.MSprintExRibbon.MasterFolderPath + "\Logs")) Then
                        Directory.CreateDirectory(Globals.Ribbons.MSprintExRibbon.MasterFolderPath + "\Logs")
                    End If
                    Globals.Ribbons.MSprintExRibbon.LogDirectoryPath = Globals.Ribbons.MSprintExRibbon.MasterFolderPath + "\Logs\"
                    Globals.Ribbons.MSprintExRibbon.logDirectoryXML = <LogDirPath></LogDirPath>
                    Globals.Ribbons.MSprintExRibbon.logDirectoryXML.Value = Globals.Ribbons.MSprintExRibbon.LogDirectoryPath
                    Globals.Ribbons.MSprintExRibbon.logDirectoryXML.Save(Globals.Ribbons.MSprintExRibbon.MasterFolderPath + "\LogDirecPath")
            End Try

            ' InputXMLFolderPath = planFolderPath.FileName
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class