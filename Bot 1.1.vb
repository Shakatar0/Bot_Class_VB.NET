'Created by SHAKATAR
'Youtube Channel : https://www.youtube.com/channel/UC6lbgL5LYO0f4-4DH8JNU4Q
'Please Subscribe
Public Module Bot
    Private Declare Function ShowWindow Lib "user32" (ByVal handle As IntPtr, ByVal nCmdShow As Integer) As Integer
    Declare Sub mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Public Sub Click()
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Public Sub ClickPos(ByVal x As Integer, ByVal y As Integer)
        Cursor.Position = New Point(x, y)
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Public Sub MoveCursor(ByVal x As Integer, ByVal y As Integer)
        Cursor.Position = New Point(x, y)
    End Sub

    Public Function getCursorPosition()
        Return Cursor.Position
    End Function

    Public Function getCursorXPosition()
        Return Cursor.Position.X
    End Function

    Public Function getCursorYPosition()
        Return Cursor.Position.Y
    End Function

    Public Sub HideApplication(ByVal NameProcess As String)
        If NameProcess.EndsWith(".exe") Then
            NameProcess = NameProcess.Replace(".exe", Nothing)
        End If

        If Process.GetProcessesByName(NameProcess).Length > 0 Then
            For Each p As Process In Process.GetProcessesByName(NameProcess)
                ShowWindow(p.MainWindowHandle, 0)
            Next
        Else
            MsgBox("Le processus " & NameProcess & ".exe" & " n'a pas était trouvé !", vbCritical, "Erreur !")
        End If
    End Sub

    Public Sub MinimizeApplication(ByVal NameProcess As String)
        If NameProcess.EndsWith(".exe") Then
            NameProcess = NameProcess.Replace(".exe", Nothing)
        End If

        If Process.GetProcessesByName(NameProcess).Length > 0 Then
            For Each p As Process In Process.GetProcessesByName(NameProcess)
                ShowWindow(p.MainWindowHandle, 2)
            Next
        Else
            MsgBox("Le processus " & NameProcess & ".exe" & " n'a pas était trouvé !", vbCritical, "Erreur !")
        End If
    End Sub

    Public Sub MaximizeApplication(ByVal NameProcess As String)
        If NameProcess.EndsWith(".exe") Then
            NameProcess = NameProcess.Replace(".exe", Nothing)
        End If

        If Process.GetProcessesByName(NameProcess).Length > 0 Then
            For Each p As Process In Process.GetProcessesByName(NameProcess)
                ShowWindow(p.MainWindowHandle, 3)
            Next
        Else
            MsgBox("Le processus " & NameProcess & ".exe" & " n'a pas était trouvé !", vbCritical, "Erreur !")
        End If
    End Sub

    Public Sub FocusApplication(ByVal NameProcess As String)
        If NameProcess.EndsWith(".exe") Then
            NameProcess = NameProcess.Replace(".exe", Nothing)
        End If

        If Process.GetProcessesByName(NameProcess).Length > 0 Then
            For Each p In Process.GetProcessesByName(NameProcess)
                AppActivate(p.MainWindowTitle)
            Next
        Else
            MsgBox("Le processus " & NameProcess & ".exe" & " n'a pas était trouvé !", vbCritical, "Erreur !")
        End If
    End Sub

    Public Sub KillApplication(ByVal NameProcess As String)
        If NameProcess.EndsWith(".exe") Then
            NameProcess = NameProcess.Replace(".exe", Nothing)
        End If

        Dim myProcesses As Process() = Process.GetProcessesByName(NameProcess)
        Dim myProcess As Process
        For Each myProcess In myProcesses
            myProcess.Kill()
        Next myProcess
    End Sub

    Public Sub OpenApplication(ByVal Path As String)
        Dim App As New Process()

        App.StartInfo.FileName = Path
        App.Start()
    End Sub

    Public Sub Batch(ByVal Commands As String, Optional ByVal WindowStyle As ProcessWindowStyle = ProcessWindowStyle.Normal)
        Dim Batch As New Process()

        Batch.StartInfo.FileName = "cmd.exe"
        Batch.StartInfo.Arguments = "/c " & Commands
        Batch.StartInfo.WindowStyle = WindowStyle

        Batch.Start()
    End Sub

    Public Function getPixelColor(ByVal x As Integer, y As Integer)
        Dim Bitmap As New Bitmap(1, 1)
        Using gr As Graphics = Graphics.FromImage(Bitmap)
            gr.CopyFromScreen(New Point(x, y), New Point(0, 0), New Size(1, 1))
        End Using
        Dim Pixel As Drawing.Color = Bitmap.GetPixel(0, 0)

        Return Pixel
    End Function


    Public Sub WaitMilli(ByVal Milliseconds As Integer)
        System.Threading.Thread.Sleep(Milliseconds)
    End Sub

    Public Sub WaitSec(ByVal Seconds As Integer)
        System.Threading.Thread.Sleep(Seconds * 1000)
    End Sub

    Public Sub WaitMin(ByVal Minutes As Integer)
        System.Threading.Thread.Sleep(Minutes * 60000)
    End Sub

    Public Sub WaitHour(ByVal Hours As Integer)
        System.Threading.Thread.Sleep(Hours * 3600000)
    End Sub
End Module
