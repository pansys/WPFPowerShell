Imports System.Collections.ObjectModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Interop

Class MainWindow
    Declare Auto Function SetParent Lib "user32.dll" (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Integer
    Declare Auto Function GetParent Lib "user32.dll" (ByVal hWnd As IntPtr) As IntPtr
    Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Declare Auto Function MoveWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    Declare Auto Function PostMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    Declare Auto Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Integer) As Integer
    Declare Auto Function GetDesktopWindow Lib "user32.dll" () As Long

    Declare Auto Function GetWindowLong Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    Declare Auto Function SetWindowLong Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Declare Auto Function SetWindowPos Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As Integer) As Boolean
    Declare Auto Function ShowWindow Lib "user32.dll" (ByVal hWnd, ByVal nCmdShow) As Boolean

    Const WS_BORDER As Integer = 8388608
    Const WS_DLGFRAME As Integer = 4194304
    Const WS_CAPTION As Integer = WS_BORDER Or WS_DLGFRAME
    Const WS_SYSMENU As Integer = 524288
    Const WS_THICKFRAME As Integer = 262144
    Const WS_MINIMIZE As Integer = 536870912
    Const WS_MAXIMIZEBOX As Integer = 65536
    Const GWL_STYLE As Integer = -16&
    Const GWL_EXSTYLE As Integer = -20&
    Const WS_EX_DLGMODALFRAME As Integer = &H1L

    Const SW_HIDE As Integer = 0
    Const SW_RESTORE As Integer = 9
    Const SW_MINIMIZE As Integer = 6
    Const SW_SHOWMINIMIZED As Integer = 2

    Private Const WM_SYSCOMMAND As Integer = 274
    Private Const SC_MAXIMIZE As Integer = 61488

    Dim psHandle As IntPtr
    Dim psPanel As Forms.Panel
    Private Sub Grid_Loaded(sender As Object, e As RoutedEventArgs)
        Dim startInfo As New ProcessStartInfo("powershell.exe")
        startInfo.WindowStyle = ProcessWindowStyle.Hidden
        'startInfo.CreateNoWindow = True
        'startInfo.UseShellExecute = True

        Dim p As System.Diagnostics.Process = Process.Start(startInfo)

        psPanel = New System.Windows.Forms.Panel()
        psPanel.Dock = DockStyle.Fill
        winFormsHost.Child = psPanel

        Threading.Thread.Sleep(1000)
        Dim success As Integer = SetParent(p.MainWindowHandle, psPanel.Handle)
        MakeExternalWindowBorderless(p.MainWindowHandle)
        psHandle = p.MainWindowHandle
    End Sub

    Public Sub MakeExternalWindowBorderless(ByVal MainWindowHandle As IntPtr)
        Dim Style As Integer
        Style = GetWindowLong(MainWindowHandle, GWL_STYLE)
        Style = Style And Not WS_CAPTION
        Style = Style And Not WS_SYSMENU
        Style = Style And Not WS_THICKFRAME
        Style = Style And Not WS_MINIMIZE
        Style = Style And Not WS_MAXIMIZEBOX
        SetWindowLong(MainWindowHandle, GWL_STYLE, Style)
        Style = GetWindowLong(MainWindowHandle, GWL_EXSTYLE)
        SetWindowLong(MainWindowHandle, GWL_EXSTYLE, Style Or WS_EX_DLGMODALFRAME)
    End Sub

    Private Sub winFormsHost_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles winFormsHost.SizeChanged
        MoveWindow(psHandle, 0, 0, psPanel.Width, psPanel.Height, True)
    End Sub
End Class

