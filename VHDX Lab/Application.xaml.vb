Imports System.Diagnostics
Imports System.Security.Principal
Imports System.Threading.Tasks
Imports System.Windows

Class Application

    Private Async Sub Application_Startup(sender As Object, e As StartupEventArgs)
        If Not IsAdministrator() Then
            Try
                Dim psi As New ProcessStartInfo With {
                    .FileName = Process.GetCurrentProcess().MainModule.FileName,
                    .UseShellExecute = True,
                    .Verb = "runas"
                }
                Process.Start(psi)
            Catch ex As Exception
                MessageBox.Show("Failed to restart with elevated privileges: " & ex.Message, "Elevation Error",
                                MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Shutdown()
            Return
        End If

        Dim previousMode = ShutdownMode
        ShutdownMode = ShutdownMode.OnExplicitShutdown

        Dim splash As New SplashWindow()
        splash.Show()

        Try
            Await Task.Delay(TimeSpan.FromSeconds(1.5))
            splash.Close()

            Dim main = New MainWindow()
            Me.MainWindow = main
            main.Show()
        Finally
            ShutdownMode = previousMode
        End Try
    End Sub

    Private Shared Function IsAdministrator() As Boolean
        Try
            Dim wi = WindowsIdentity.GetCurrent()
            Dim wp = New WindowsPrincipal(wi)
            Return wp.IsInRole(WindowsBuiltInRole.Administrator)
        Catch
            Return False
        End Try
    End Function

End Class
