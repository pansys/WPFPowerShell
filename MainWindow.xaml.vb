Imports System.Collections.ObjectModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Text
Imports System.Windows.Interop

Class MainWindow
    Declare Auto Function SetParent Lib "user32.dll" (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Integer
    Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Private Const WM_SYSCOMMAND As Integer = 274
    Private Const SC_MAXIMIZE As Integer = 61488

    Private Sub Grid_Loaded(sender As Object, e As RoutedEventArgs)
        Dim startInfo As New ProcessStartInfo("powershell.exe")
        startInfo.WindowStyle = ProcessWindowStyle.Hidden 'this is not working, need fine tuning

        Dim p As System.Diagnostics.Process = Process.Start(startInfo)
        'p.WaitForInputIdle() 'If this works below sleep is not required and embedding external window is quicker
        Threading.Thread.Sleep(1000) 'Change this to 100 to avoid poping up the external powershell
        SetParent(p.MainWindowHandle, New WindowInteropHelper(Application.Current.MainWindow).Handle)
        SendMessage(p.MainWindowHandle, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
        p.StartInfo.WindowStyle = ProcessWindowStyle.Normal
    End Sub
End Class

