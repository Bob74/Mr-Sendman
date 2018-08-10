Imports System.Security.Principal

Module VistaSecurity

    ' Declare API
    Private Declare Ansi Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Private Const BCM_FIRST As Int32 = &H1600
    Private Const BCM_SETSHIELD As Int32 = (BCM_FIRST + &HC)

    Public Function IsVistaOrHigher() As Boolean
        Return Environment.OSVersion.Version.Major >= 6
    End Function

    ' Checks if the process is elevated
    Public Function IsAdmin() As Boolean
        Dim id As WindowsIdentity = WindowsIdentity.GetCurrent()
        Dim p As WindowsPrincipal = New WindowsPrincipal(id)
        Return p.IsInRole(WindowsBuiltInRole.Administrator)
    End Function


    ' Restart the current process with administrator credentials
    Public Sub RestartElevated()
        Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
        startInfo.UseShellExecute = True
        startInfo.WorkingDirectory = Environment.CurrentDirectory
        startInfo.FileName = System.Reflection.Assembly.GetExecutingAssembly().Location
        startInfo.Verb = "runas"
        Try
            Dim p As Process = Process.Start(startInfo)
        Catch ex As Exception   'Si annulé (pas d'appli avec droits) ...
            'System.Windows.Application.Current.Shutdown()  'Ferme appli sans droits
            Return
        End Try
        'Appli avec droits ouverte
        System.Windows.Application.Current.Shutdown()      'Ferme appli sans droits
    End Sub

End Module
