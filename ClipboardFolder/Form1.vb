Imports System.IO
Imports System.Windows.Forms

Public Class Form1
    Private watchFolder As String = System.IO.Path.Combine(Application.StartupPath, "Clipboard")
    Private WithEvents checkFileTimer As New Timer()
    Private fileToCheck As String
    Private trayIcon As New NotifyIcon()
    Private contxtMenu As New ContextMenuStrip()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize FileSystemWatcher
        FileSystemWatcher1.EnableRaisingEvents = False

        ' Initialize Timer
        checkFileTimer.Interval = 500 ' Check every 500 milliseconds

        ' Initialize System Tray Icon
        trayIcon.Icon = My.Resources.icon_ClipboardFolder

        trayIcon.Visible = True
        trayIcon.Text = "Clipboard Folder"

        ' Initialize Context Menu with Options
        ' Settings Button
        'Dim settingsOption As New ToolStripMenuItem("Settings")
        'AddHandler settingsOption.Click, AddressOf ShowForm
        'contxtMenu.Items.Add(settingsOption)

        ' Exit Button
        Dim exitOption As New ToolStripMenuItem("Exit")
        AddHandler exitOption.Click, AddressOf ExitApplication
        contxtMenu.Items.Add(exitOption)

        trayIcon.ContextMenuStrip = contxtMenu

        ' Load the saved folder path
        If Directory.Exists(watchFolder) Then
            FileSystemWatcher1.Path = watchFolder
            FileSystemWatcher1.EnableRaisingEvents = True
            'MessageBox.Show("Watching folder: " & watchFolder)
        Else
            MessageBox.Show("The 'Clipboard' folder is missing. Please re-add it the Clipboard App Folder.")
        End If

        ' Hide the main form and prevent it from appearing when starting the app
        Me.Visible = False
        Me.ShowInTaskbar = False
    End Sub

    Private Sub FileSystemWatcher1_Created(sender As Object, e As FileSystemEventArgs) Handles FileSystemWatcher1.Created
        ' When a new file is created, set a timer to routinely check when the file can be safely copied to the clipboard.
        fileToCheck = e.FullPath
        checkFileTimer.Start()
    End Sub

    Private Sub CheckFileTimer_Tick(sender As Object, e As EventArgs) Handles checkFileTimer.Tick
        ' Whenever the timer rings, check to see if the file is ready for copying.
        Dim fileReady As Boolean = False

        Try
            Using fileStream As FileStream = File.Open(fileToCheck, FileMode.Open, FileAccess.Read, FileShare.None)
                fileReady = True
            End Using
        Catch ex As IOException
            fileReady = False
        End Try

        If fileReady Then
            checkFileTimer.Stop()
            CopyFileToClipboard(fileToCheck)
        End If
    End Sub

    Private Sub CopyFileToClipboard(filePath As String)
        Try
            Dim fileExtension As String = Path.GetExtension(filePath).ToLower()

            Select Case fileExtension
                Case ".txt"
                    ' Read the file content and copy to clipboard
                    Dim textContent As String = File.ReadAllText(filePath)
                    Clipboard.SetText(textContent)

                Case ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".webp", ".jfif"
                    ' Load the image, set it to the clipboard, and dispose of the image.
                    Using image As Image = Image.FromFile(filePath)
                        Clipboard.SetImage(image)
                    End Using

                Case Else
                    ' Handle other file types by adding them to the clipboard as a file drop list
                    Clipboard.SetFileDropList(New Specialized.StringCollection From {filePath})
            End Select

            ' Notify the user that the file was copied to the clipboard
            'MessageBox.Show("File copied to clipboard: " & filePath)

            ' Delete the file after copying
            If File.Exists(filePath) Then
                File.Delete(filePath)
                'MessageBox.Show("File deleted: " & filePath)
            End If

        Catch ex As Exception
            MessageBox.Show("Error copying file to clipboard: " & ex.Message)
        End Try
    End Sub


    'Menu Related Functions ---------------------------------------------------------
    Private Sub ShowForm(sender As Object, e As EventArgs)
        ' Show the form and bring it to the front
        Me.Visible = True
        Me.WindowState = FormWindowState.Normal
        Me.BringToFront()
        Me.ShowInTaskbar = True
    End Sub

    Private Sub ExitApplication(sender As Object, e As EventArgs)
        trayIcon.Visible = False
        ' Delete all files within the clipboard folder
        If Directory.Exists(watchFolder) Then
            For Each item In Directory.EnumerateFiles(watchFolder)
                File.Delete(item)
            Next
        End If
        Application.Exit()
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Me.Hide()
        End If
        MyBase.OnFormClosing(e)
    End Sub

End Class
