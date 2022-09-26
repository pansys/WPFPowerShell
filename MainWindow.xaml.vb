Imports System.Collections.ObjectModel
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Text
Imports System.Windows.Interop

Class MainWindow
    Declare Auto Function SetParent Lib "user32.dll" (ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As Integer
    Declare Auto Function GetParent Lib "user32.dll" (ByVal hWnd As IntPtr) As IntPtr
    Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Declare Auto Function MoveWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    Declare Auto Function PostMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean

    Declare Auto Function GetWindowLong Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    Declare Auto Function SetWindowLong Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Declare Auto Function SetWindowPos Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As Integer) As Boolean


    Private Const WM_CLOSE As Integer = 16
    Private Const WM_SYSCOMMAND As Integer = 274
    Private Const SC_MAXIMIZE As Integer = 61488

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
    Const SWP_NOMOVE As Integer = &H2
    Const SWP_NOSIZE As Integer = &H1
    Const SWP_FRAMECHANGED As Integer = &H20
    Const MF_BYPOSITION As UInteger = &H400
    Const MF_REMOVE As UInteger = &H1000

    'Private Const GWL_STYLE As Integer = -16
    'Private Const WS_CHILD As uint = 0x40000000 
    'Private Const WS_BORDER = 0x00800000 'window With border
    'Private Const WS_DLGFRAME = 0x00400000 'window With Double border but no title

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

        Dim prnt As IntPtr = GetParent(p.MainWindowHandle)
        'Dim pt As System.Drawing.Point = New System.Drawing.Point(100, 100)
        'Do While psPanel.Controls.Count = 0
        SetParent(p.MainWindowHandle, psPanel.Handle)
        Dim prnt2 As IntPtr = GetParent(p.MainWindowHandle)
        'p.StartInfo.WindowStyle = ProcessWindowStyle.Maximized
        'SendMessage(p.MainWindowHandle, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
        MakeExternalWindowBorderless(p.MainWindowHandle)
        MoveWindow(p.MainWindowHandle, 0, 0, psPanel.Width, psPanel.Height, True)
        'SetWindowLong(p.MainWindowHandle, GWL_EXSTYLE, Style Or WS_EX_DLGMODALFRAME)


        'Threading.Thread.Sleep(100)
        'Exit Do
        'Loop

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
        'SetWindowPos(MainWindowHandle, New IntPtr(0), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED)
    End Sub
End Class

