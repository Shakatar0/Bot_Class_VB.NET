'Created by SHAKATAR
'Youtube Channel : https://www.youtube.com/channel/UC6lbgL5LYO0f4-4DH8JNU4Q
'Please Subscribe
Public Class Bot
    Declare Sub mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)
    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Public Shared Sub Click()
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Public Shared Sub ClickPos(ByVal x As Integer, ByVal y As Integer)
        Cursor.Position = New Point(x, y)
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Public Shared Sub MoveCursor(ByVal x As Integer, ByVal y As Integer)
        Cursor.Position = New Point(x, y)
    End Sub

    Public Shared Function getCursorPosition()
        Return Cursor.Position
    End Function

    Public Shared Function getCursorXPosition()
        Return Cursor.Position.X
    End Function

    Public Shared Function getCursorYPosition()
        Return Cursor.Position.Y
    End Function

    Public Shared Sub FocusApplication(ByVal NameProcess As String)
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

    Public Shared Sub KillApplication(ByVal NameProcess As String)
        If NameProcess.EndsWith(".exe") Then
            NameProcess = NameProcess.Replace(".exe", Nothing)
        End If

        Dim myProcesses As Process() = Process.GetProcessesByName(NameProcess)
        Dim myProcess As Process
        For Each myProcess In myProcesses
            myProcess.Kill()
        Next myProcess
    End Sub

    Public Shared Sub OpenApplication(ByVal Path As String)
        Dim App As New Process()

        App.StartInfo.FileName = Path
        App.Start()
    End Sub

    Public Shared Sub Batch(ByVal Commands As String, Optional ByVal WindowStyle As ProcessWindowStyle = ProcessWindowStyle.Normal)
        Dim Batch As New Process()

        Batch.StartInfo.FileName = "cmd.exe"
        Batch.StartInfo.Arguments = "/c " & Commands
        Batch.StartInfo.WindowStyle = WindowStyle

        Batch.Start()
    End Sub

    Public Shared Sub WaitMilli(ByVal Milliseconds As Integer)
        System.Threading.Thread.Sleep(Milliseconds)
    End Sub

    Public Shared Sub WaitSec(ByVal Seconds As Integer)
        System.Threading.Thread.Sleep(Seconds * 1000)
    End Sub

    Public Shared Sub WaitMin(ByVal Minutes As Integer)
        System.Threading.Thread.Sleep(Minutes * 60000)
    End Sub

    Public Shared Sub WaitHour(ByVal Hours As Integer)
        System.Threading.Thread.Sleep(Hours * 3600000)
    End Sub
End Class
