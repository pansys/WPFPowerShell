Imports System.Collections.ObjectModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Text
Imports System.Windows.Interop

Class MainWindow
    Declare Auto Function SetParent Lib "user32.dll" (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Integer
    Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Declare Auto Function MoveWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    Declare Auto Function PostMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean

    Private Const WM_CLOSE As Integer = 16
    Private Const WM_SYSCOMMAND As Integer = 274
    Private Const SC_MAXIMIZE As Integer = 61488

    Private Sub Grid_Loaded(sender As Object, e As RoutedEventArgs)
        Dim startInfo As New ProcessStartInfo("powershell.exe")
        startInfo.WindowStyle = ProcessWindowStyle.Hidden 'this is not working, need fine tuning

        Dim p As System.Diagnostics.Process = Process.Start(startInfo)
        'p.WaitForInputIdle() 'If this works below sleep is not required and embedding external window is quicker
        Threading.Thread.Sleep(1000) 'Change this to 100 to avoid poping up the external powershell
        'SetParent(p.MainWindowHandle, New WindowInteropHelper(Application.Current.MainWindow).Handle)

        Dim psPanel As Forms.Panel = New System.Windows.Forms.Panel()
        psPanel.Width = 1180
        psPanel.Height = 500
        winFormsHost.Child = psPanel

        SetParent(p.MainWindowHandle, psPanel.Handle)
        'p.StartInfo.WindowStyle = ProcessWindowStyle.Maximized
        'SendMessage(p.MainWindowHandle, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
        MoveWindow(p.MainWindowHandle, 0, 0, psPanel.Width, psPanel.Height, True)
    End Sub
End Class

