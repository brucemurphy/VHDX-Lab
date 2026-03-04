Imports System.IO
Imports System.Net.Http
Imports System.Security.Principal
Imports System.Threading
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Effects
Imports System.Windows.Media.Imaging
Imports System.Windows.Threading
Imports System.Xml.Linq
Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.Linq
Imports Forms = System.Windows.Forms
Imports System.Collections.Generic
Imports System.Windows.Controls

Partial Class MainWindow
    Private Shared ReadOnly httpClient As New HttpClient()
    Private bcdEntries As New List(Of BcdEntry)()

    ' HDD indicator (from DISM Lab)
    Private _dismBlinkTimer As DispatcherTimer
    Private _dismActive As Boolean = False
    Private ReadOnly _rand As New Random()

    ' Progress tracking
    Private _lastReportedBytes As Long = -1
    Private _lastReportedPercentage As Integer = -1

    ' Track driver export path for cleanup
    Private _driverExportPath As String = Nothing

    Private Class BcdEntry
        Public Property Identifier As String
        Public Property DisplayName As String
        Public Property FullDetails As String

        Public Overrides Function ToString() As String
            Return If(Not String.IsNullOrWhiteSpace(DisplayName), DisplayName, Identifier)
        End Function

    End Class

    ' ==================== CREATE VHDX FROM FOLDER ====================

    Private Async Sub BrowseSourceFolderButton_Click(sender As Object, e As RoutedEventArgs) Handles BrowseSourceFolderButton.Click
        Dim folderDialog As New Forms.FolderBrowserDialog With {.Description = "Select the folder to package into a VHDX"}
        If folderDialog.ShowDialog() = Forms.DialogResult.OK Then
            Await SetSourceFolderAsync(folderDialog.SelectedPath)
        End If
    End Sub

    Private Sub BrowseSavePathButton_Click(sender As Object, e As RoutedEventArgs) Handles BrowseSavePathButton.Click
        Dim defaultName = "Image.vhdx"
        If Not String.IsNullOrWhiteSpace(SourceFolderTextBox.Text) Then
            defaultName = Path.GetFileName(SourceFolderTextBox.Text.TrimEnd(Path.DirectorySeparatorChar)) & ".vhdx"
        End If

        Dim saveDialog As New SaveFileDialog With {
            .Filter = "VHDX Files (*.vhdx)|*.vhdx",
            .Title = "Save VHDX",
            .FileName = defaultName,
            .OverwritePrompt = True
        }

        If saveDialog.ShowDialog() = True Then
            SavePathTextBox.Text = saveDialog.FileName
        End If
    End Sub

    Private Async Function SetSourceFolderAsync(sourcePath As String) As Task
        If String.IsNullOrWhiteSpace(sourcePath) OrElse Not Directory.Exists(sourcePath) Then
            MessageBox.Show("Please select a valid source folder.", "Invalid Source", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        SourceFolderTextBox.Text = sourcePath

        UpdateProgress("Scanning...", "Calculating folder size")
        StartDismIndicator()
        Dim folderSize As Long = Await Task.Run(Function() GetFolderSizeSafe(sourcePath))
        StopDismIndicator()

        Dim suggestedSizeGb As Double = Math.Max(1, Math.Ceiling((folderSize * 1.2) / (1024.0 * 1024 * 1024)))
        VhdxSizeTextBox.Text = suggestedSizeGb.ToString("0")

        If String.IsNullOrWhiteSpace(SavePathTextBox.Text) Then
            Dim trimmed = sourcePath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            Dim defaultName = System.IO.Path.GetFileName(trimmed) & ".vhdx"
            Dim defaultDir = System.IO.Path.GetDirectoryName(trimmed)
            If Not String.IsNullOrWhiteSpace(defaultDir) Then
                SavePathTextBox.Text = System.IO.Path.Combine(defaultDir, defaultName)
            Else
                SavePathTextBox.Text = defaultName
            End If
        End If
    End Function

    Private Sub CreateVhdxFromFolderButton_Click(sender As Object, e As RoutedEventArgs) Handles CreateVhdxFromFolderButton.Click
        If Not IsAdministrator() Then
            MessageBox.Show("Administrator privileges are required to create a VHDX.", "Administrator Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Only show the pane; actions happen via Start
        If CreatePane.Visibility <> Visibility.Visible Then
            CreatePane.Visibility = Visibility.Visible
        End If
    End Sub

    Private Async Sub StartCreateVhdxButton_Click(sender As Object, e As RoutedEventArgs) Handles StartCreateVhdxButton.Click
        Await RunCreateVhdxWorkflowAsync()
    End Sub

    Private Sub CancelCreateVhdxButton_Click(sender As Object, e As RoutedEventArgs) Handles CancelCreateVhdxButton.Click
        CreatePane.Visibility = Visibility.Collapsed
    End Sub

    Private Async Function RunCreateVhdxWorkflowAsync() As Task
        Dim workflowStarted As Boolean = False

        Dim sourceFolder As String = SourceFolderTextBox.Text
        If String.IsNullOrWhiteSpace(sourceFolder) OrElse Not Directory.Exists(sourceFolder) Then
            BrowseSourceFolderButton_Click(Nothing, Nothing)
            Return
        End If

        Dim folderSize As Long = Await Task.Run(Function() GetFolderSizeSafe(sourceFolder))

        If String.IsNullOrWhiteSpace(VhdxSizeTextBox.Text) Then
            Dim suggestedSizeGb As Double = Math.Max(1, Math.Ceiling((folderSize * 1.2) / (1024.0 * 1024 * 1024)))
            VhdxSizeTextBox.Text = suggestedSizeGb.ToString("0")
        End If

        Dim sizeGb As Double
        If Not Double.TryParse(VhdxSizeTextBox.Text, sizeGb) OrElse sizeGb <= 0 Then
            MessageBox.Show("Invalid size entered.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim sizeBytes As Double = sizeGb * 1024 * 1024 * 1024
        Dim sizeMb As Long = CLng(Math.Ceiling(sizeBytes / (1024 * 1024)))
        If sizeBytes < folderSize Then
            MessageBox.Show("The VHDX size must be greater than or equal to the folder size.", "Invalid Size", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' File system choice
        If FileSystemCombo.SelectedIndex < 0 Then FileSystemCombo.SelectedIndex = 0
        Dim fsChoice As String = TryCast((TryCast(FileSystemCombo.SelectedItem, ComboBoxItem))?.Content, String)
        fsChoice = If(String.IsNullOrWhiteSpace(fsChoice), "FAT32", fsChoice.Trim().ToUpperInvariant())
        Select Case fsChoice
            Case "NTFS", "FAT32", "EXFAT"
            Case Else
                MessageBox.Show("Invalid file system selection. Choose NTFS, FAT32, or exFAT.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
        End Select

        ' Windows cannot format FAT32 volumes larger than 32GB
        If fsChoice = "FAT32" AndAlso sizeGb > 32 Then
            Dim res = MessageBox.Show("Windows cannot format FAT32 volumes larger than 32GB. Switch to exFAT instead?", "FAT32 Size Limit", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.Yes)
            If res = MessageBoxResult.Yes Then
                fsChoice = "EXFAT"
            Else
                Return
            End If
        End If

        ' Partition style choice
        If PartitionStyleCombo.SelectedIndex < 0 Then PartitionStyleCombo.SelectedIndex = 0
        Dim partStyle As String = TryCast((TryCast(PartitionStyleCombo.SelectedItem, ComboBoxItem))?.Content, String)
        partStyle = If(String.IsNullOrWhiteSpace(partStyle), "GPT", partStyle.Trim().ToUpperInvariant())
        Dim convertCmd As String
        Dim useMsr As Boolean = False
        If partStyle = "MBR" Then
            convertCmd = "convert mbr"
        Else
            partStyle = "GPT"
            convertCmd = "convert gpt"
            useMsr = True
            If sizeMb <= 32 Then
                MessageBox.Show("GPT layout requires some overhead (MSR). Increase the size above 32MB or choose MBR.", "Size Too Small for GPT", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If
        End If

        Dim vhdxPath As String = SavePathTextBox.Text
        If String.IsNullOrWhiteSpace(vhdxPath) Then
            BrowseSavePathButton_Click(Nothing, Nothing)
            vhdxPath = SavePathTextBox.Text
        End If

        If String.IsNullOrWhiteSpace(vhdxPath) Then
            MessageBox.Show("Please choose a destination path.", "Missing Destination", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        Dim driveLetter As String = GetAvailableDriveLetter()

        If String.IsNullOrWhiteSpace(driveLetter) Then
            MessageBox.Show("No available drive letters to mount the VHDX.", "Drive Letter Unavailable", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End If

        workflowStarted = True

        Dim createScriptPath As String = Nothing
        Dim caughtEx As Exception = Nothing
        Try
            UpdateProgress("Creating...", "Building virtual disk")
            StartDismIndicator()

            createScriptPath = Path.GetTempFileName()
            Dim scriptLines As New List(Of String) From {
                $"create vdisk file=""{vhdxPath}"" maximum={sizeMb} type=expandable",
                $"select vdisk file=""{vhdxPath}""",
                "attach vdisk",
                convertCmd
            }

            If useMsr Then
                scriptLines.Add("create partition msr size=16")
            End If

            scriptLines.Add("create partition primary")
            scriptLines.Add($"format fs={fsChoice} quick label=""VHDXFolder""")
            scriptLines.Add($"assign letter={driveLetter}")
            scriptLines.Add("exit")

            File.WriteAllLines(createScriptPath, scriptLines)

            Dim dpResult = Await RunDiskpartScriptAsync(createScriptPath)
            If dpResult.ExitCode <> 0 Then
                StopDismIndicator()
                UpdateProgressError("DiskPart Error", dpResult.StdErr)
                Await DetachVhdxAsync(vhdxPath)
                Await Task.Delay(2000)
                ResetProgress()
                MessageBox.Show($"Failed to create VHDX:{vbCrLf}{dpResult.StdErr}", "DiskPart Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            UpdateProgress("Copying...", "Transferring files to VHDX")

            Dim targetRoot = driveLetter & ":\"
            Await Task.Run(Sub() CopyDirectoryContents(sourceFolder, targetRoot, folderSize))

            UpdateProgress("Detaching...", "Finalizing VHDX")
            Await DetachVhdxAsync(vhdxPath)

            StopDismIndicator()
            UpdateProgressSuccess("Complete", $"Saved to {vhdxPath} ({partStyle}/{fsChoice})")
            Await Task.Delay(2000)
            ResetProgress()

            MessageBox.Show($"VHDX created successfully!{vbCrLf}Source: {sourceFolder}{vbCrLf}Destination: {vhdxPath}{vbCrLf}Partition: {partStyle}{vbCrLf}File System: {fsChoice}", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        Catch ex As Exception
            caughtEx = ex
        Finally
            If Not String.IsNullOrWhiteSpace(createScriptPath) AndAlso File.Exists(createScriptPath) Then
                Try
                    File.Delete(createScriptPath)
                Catch
                End Try
            End If
        End Try

        If caughtEx IsNot Nothing Then
            StopDismIndicator()
            UpdateProgressError("Error", caughtEx.Message)
            Await Task.Delay(2000)
            ResetProgress()
            MessageBox.Show($"Failed to create VHDX: {caughtEx.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Await DetachVhdxAsync(vhdxPath)
        End If

        If workflowStarted Then
            CreatePane.Visibility = Visibility.Collapsed
        End If
    End Function

    Private Function GetFolderSizeSafe(path As String) As Long
        Dim total As Long = 0
        For Each file In EnumerateSafeFiles(path)
            Try
                total += New FileInfo(file).Length
            Catch
            End Try
        Next
        Return total
    End Function

    Private Function GetAvailableDriveLetter() As String
        Dim usedLetters As New HashSet(Of Char)(DriveInfo.GetDrives().Select(Function(d) Char.ToUpperInvariant(d.Name(0))))
        For letterCode As Integer = Asc("Z"c) To Asc("D"c) Step -1
            Dim letter = Chr(letterCode)
            If Not usedLetters.Contains(letter) Then
                Return letter
            End If
        Next
        Return Nothing
    End Function

    Private Async Function RunDiskpartScriptAsync(scriptPath As String) As Task(Of (ExitCode As Integer, StdOut As String, StdErr As String))
        Try
            Dim psi As New ProcessStartInfo("diskpart.exe", $"/s ""{scriptPath}""") With {
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }

            Using p As Process = Process.Start(psi)
                Dim stdOutTask = p.StandardOutput.ReadToEndAsync()
                Dim stdErrTask = p.StandardError.ReadToEndAsync()
                Await p.WaitForExitAsync()
                Return (p.ExitCode, Await stdOutTask, Await stdErrTask)
            End Using
        Catch ex As Exception
            Return (-1, String.Empty, ex.Message)
        End Try
    End Function

    Private Async Function DetachVhdxAsync(vhdxPath As String) As Task
        Dim detachScript As String = Nothing
        Try
            detachScript = Path.GetTempFileName()
            File.WriteAllLines(detachScript, New String() {
                               $"select vdisk file=""{vhdxPath}""",
                               "detach vdisk",
                               "exit"})
            Await RunDiskpartScriptAsync(detachScript)
        Catch
        Finally
            If Not String.IsNullOrWhiteSpace(detachScript) AndAlso File.Exists(detachScript) Then
                Try
                    File.Delete(detachScript)
                Catch
                End Try
            End If
        End Try
    End Function

    Private Function ShouldSkipSystemFolder(folderName As String) As Boolean
        If String.IsNullOrWhiteSpace(folderName) Then Return False
        Return folderName.Equals("System Volume Information", StringComparison.OrdinalIgnoreCase) OrElse
               folderName.Equals("$RECYCLE.BIN", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Iterator Function EnumerateSafeDirectories(root As String) As IEnumerable(Of String)
        Dim stack As New Stack(Of String)()
        stack.Push(root)

        While stack.Count > 0
            Dim current = stack.Pop()
            Dim dirs() As String = {}
            Try
                dirs = System.IO.Directory.GetDirectories(current)
            Catch ex As UnauthorizedAccessException
                Continue While
            Catch ex As IOException
                Continue While
            End Try

            For Each directoryPath In dirs
                Dim name = Path.GetFileName(directoryPath)
                If ShouldSkipSystemFolder(name) Then Continue For
                Yield directoryPath
                stack.Push(directoryPath)
            Next
        End While
    End Function

    Private Iterator Function EnumerateSafeFiles(root As String) As IEnumerable(Of String)
        Dim stack As New Stack(Of String)()
        stack.Push(root)

        While stack.Count > 0
            Dim current = stack.Pop()
            Dim dirs() As String = {}
            Try
                For Each f In System.IO.Directory.GetFiles(current)
                    Yield f
                Next
                dirs = System.IO.Directory.GetDirectories(current)
            Catch ex As UnauthorizedAccessException
                Continue While
            Catch ex As IOException
                Continue While
            End Try

            For Each directoryPath In dirs
                Dim name = Path.GetFileName(directoryPath)
                If ShouldSkipSystemFolder(name) Then Continue For
                stack.Push(directoryPath)
            Next
        End While
    End Function

    Private Sub CopyDirectoryContents(sourceFolder As String, targetRoot As String, totalSize As Long)
        If Not System.IO.Directory.Exists(targetRoot) Then
            System.IO.Directory.CreateDirectory(targetRoot)
        End If

        For Each subDir In EnumerateSafeDirectories(sourceFolder)
            Dim relativeDir = subDir.Substring(sourceFolder.Length).TrimStart(Path.DirectorySeparatorChar)
            Dim targetDir = Path.Combine(targetRoot, relativeDir)
            If Not System.IO.Directory.Exists(targetDir) Then
                System.IO.Directory.CreateDirectory(targetDir)
            End If
        Next

        Dim files = EnumerateSafeFiles(sourceFolder)
        Dim copiedBytes As Long = 0

        For Each file In files
            Dim relativePath = file.Substring(sourceFolder.Length).TrimStart(Path.DirectorySeparatorChar)
            Dim destination = Path.Combine(targetRoot, relativePath)
            Dim destDir = Path.GetDirectoryName(destination)
            If Not System.IO.Directory.Exists(destDir) Then
                System.IO.Directory.CreateDirectory(destDir)
            End If

            System.IO.File.Copy(file, destination, True)
            copiedBytes += New FileInfo(file).Length

            Dim pct As Integer = If(totalSize > 0, CInt(Math.Min(100, (copiedBytes * 100.0) / totalSize)), 0)

            Dispatcher.Invoke(Sub()
                                   MountSizeProgressText.Text = If(totalSize > 0, $"{pct}%", "Copying...")
                                   MountSizeProgressTextDetail.Text = $"{FormatBytes(copiedBytes)} / {FormatBytes(totalSize)}"
                               End Sub)
        Next
    End Sub

    Private Async Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ' Enable dark title bar (Windows 10 1809+)
        EnableDarkTitleBar()

        If Not IsAdministrator() Then
            Dim result = MessageBox.Show("This application must be run as Administrator." & vbCrLf &
                                         "Restart as Administrator now?", "Administrator Required",
                                         MessageBoxButton.YesNo, MessageBoxImage.Warning)
            If result = MessageBoxResult.Yes Then
                Try
                    Dim psi As New ProcessStartInfo With {
                        .FileName = Process.GetCurrentProcess().MainModule.FileName,
                        .UseShellExecute = True,
                        .Verb = "runas"
                    }
                    Process.Start(psi)
                Catch ex As Exception
                    MessageBox.Show("Failed to restart as Administrator: " & ex.Message, "Error",
                                    MessageBoxButton.OK, MessageBoxImage.Error)
                End Try
                Application.Current.Shutdown()
                Return
            End If
        End If

        InitDismIndicators()
        Await SetBingWallpaperAsync()
        Await LoadBcdEntriesAsync()
    End Sub

    ' ==================== DISM INDICATOR METHODS (from DISM Lab) ====================

    Private Sub InitDismIndicators()
        If _dismBlinkTimer Is Nothing Then
            _dismBlinkTimer = New DispatcherTimer(TimeSpan.FromMilliseconds(120),
                                                  DispatcherPriority.Background,
                                                  AddressOf DismBlinkTick,
                                                  Dispatcher.CurrentDispatcher)
        End If
        UpdateDismLightOff()
    End Sub

    Private Sub DismBlinkTick(sender As Object, e As EventArgs)
        If Not _dismActive Then Return
        Dim sample = _rand.NextDouble()
        If sample < 0.55 Then
            DismActivityLight.Fill = New SolidColorBrush(Color.FromRgb(0, 220, 120))
            DismActivityLight.Effect = New DropShadowEffect With {
                .Color = Colors.Lime,
                .BlurRadius = 8,
                .ShadowDepth = 0,
                .Opacity = 0.75
            }
        ElseIf sample < 0.8 Then
            UpdateDismLightDim()
        Else
            UpdateDismLightOff()
        End If
    End Sub

    Private Sub StartDismIndicator()
        _dismActive = True
        If _dismBlinkTimer IsNot Nothing AndAlso Not _dismBlinkTimer.IsEnabled Then
            _dismBlinkTimer.Start()
        End If
    End Sub

    Private Sub StopDismIndicator()
        _dismActive = False
        _dismBlinkTimer?.Stop()
        UpdateDismLightOff()
    End Sub

    Private Sub UpdateDismLightOff()
        If DismActivityLight IsNot Nothing Then
            DismActivityLight.Fill = New SolidColorBrush(Color.FromRgb(50, 50, 50))
            DismActivityLight.Effect = Nothing
        End If
    End Sub

    Private Sub UpdateDismLightDim()
        If DismActivityLight IsNot Nothing Then
            DismActivityLight.Fill = New SolidColorBrush(Color.FromRgb(25, 80, 55))
            DismActivityLight.Effect = Nothing
        End If
    End Sub

    ' ==================== PROGRESS PANEL METHODS (from DISM Lab) ====================

    Private Sub ResetProgressTracking()
        _lastReportedBytes = -1
        _lastReportedPercentage = -1
    End Sub

    Private Sub UpdateProgress(mainText As String, detailText As String)
        If MountSizeProgressText IsNot Nothing Then
            MountSizeProgressText.Text = mainText
            MountSizeProgressText.Visibility = Visibility.Visible
        End If
        If MountSizeProgressTextDetail IsNot Nothing Then
            MountSizeProgressTextDetail.Text = detailText
            MountSizeProgressTextDetail.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub UpdateProgressSuccess(mainText As String, detailText As String)
        If MountSizeProgressText IsNot Nothing Then
            MountSizeProgressText.Text = mainText
            MountSizeProgressText.Foreground = New SolidColorBrush(Colors.LimeGreen)
            MountSizeProgressText.Visibility = Visibility.Visible
        End If
        If MountSizeProgressTextDetail IsNot Nothing Then
            MountSizeProgressTextDetail.Text = detailText
            MountSizeProgressTextDetail.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub UpdateProgressError(mainText As String, detailText As String)
        If MountSizeProgressText IsNot Nothing Then
            MountSizeProgressText.Text = mainText
            MountSizeProgressText.Foreground = New SolidColorBrush(Colors.Red)
            MountSizeProgressText.Visibility = Visibility.Visible
        End If
        If MountSizeProgressTextDetail IsNot Nothing Then
            MountSizeProgressTextDetail.Text = detailText
            MountSizeProgressTextDetail.Visibility = Visibility.Visible
        End If
    End Sub

    Private Async Sub ResetProgress()
        Await Task.Delay(1000)
        If MountSizeProgressText IsNot Nothing Then
            MountSizeProgressText.Text = ""
            MountSizeProgressText.Foreground = New SolidColorBrush(Color.FromRgb(&HCC, &HCC, &HCC))
            MountSizeProgressText.Visibility = Visibility.Collapsed
        End If
        If MountSizeProgressTextDetail IsNot Nothing Then
            MountSizeProgressTextDetail.Text = ""
            MountSizeProgressTextDetail.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Function FormatBytes(value As Long) As String
        Dim units = {"B", "KB", "MB", "GB", "TB"}
        Dim dbl = CDbl(value)
        Dim idx = 0
        While dbl >= 1024 AndAlso idx < units.Length - 1
            dbl /= 1024
            idx += 1
        End While
        Return $"{dbl:0.##} {units(idx)}"
    End Function

    ' ==================== UPDATED ADD VHDX WITH DRIVER INJECTION ====================

    Private Async Sub AddVHDXButton_Click(sender As Object, e As RoutedEventArgs) Handles AddVHDXButton.Click
        ' Open file dialog to select VHDX file
        Dim openFileDialog As New OpenFileDialog With {
            .Filter = "VHDX Files (*.vhdx)|*.vhdx|All Files (*.*)|*.*",
            .Title = "Select VHDX File",
            .CheckFileExists = True
        }

        If openFileDialog.ShowDialog() <> True Then
            Return
        End If

        Dim sourcePath As String = openFileDialog.FileName
        Dim vhdxFileName As String = Path.GetFileName(sourcePath)

        ' Prompt for boot entry description
        Dim description As String = InputBox("Enter a description for the boot entry:", "Boot Entry Description", Path.GetFileNameWithoutExtension(sourcePath))
        If String.IsNullOrWhiteSpace(description) Then
            description = Path.GetFileNameWithoutExtension(sourcePath)
        End If

        ' Prompt user to select destination drive
        Dim drives As DriveInfo() = DriveInfo.GetDrives().Where(Function(d) d.DriveType = DriveType.Fixed AndAlso d.IsReady).ToArray()

        If drives.Length = 0 Then
            MessageBox.Show("No available fixed drives found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End If

        ' Build drive selection message
        Dim driveOptions As New Text.StringBuilder()
        driveOptions.AppendLine("Select a drive to save the VHDX file:")
        driveOptions.AppendLine()
        For i As Integer = 0 To drives.Length - 1
            Dim totalGB As Long = drives(i).TotalSize \ (1024 * 1024 * 1024)
            Dim freeGB As Long = drives(i).AvailableFreeSpace \ (1024 * 1024 * 1024)
            driveOptions.AppendLine($"{i + 1}. {drives(i).Name} - {drives(i).VolumeLabel} ({freeGB} GB free of {totalGB} GB)")
        Next
        driveOptions.AppendLine()
        driveOptions.Append($"Enter drive number (1-{drives.Length}):")

        Dim driveSelection As String = InputBox(driveOptions.ToString(), "Select Destination Drive", "1")
        Dim selectedDriveIndex As Integer

        If Not Integer.TryParse(driveSelection, selectedDriveIndex) OrElse selectedDriveIndex < 1 OrElse selectedDriveIndex > drives.Length Then
            MessageBox.Show("Invalid drive selection.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim selectedDrive As DriveInfo = drives(selectedDriveIndex - 1)
        Dim vhdxFolder As String = Path.Combine(selectedDrive.RootDirectory.FullName, "VHDX")

        Try
            ' Step 1: Ensure drive:\VHDX folder exists
            If Not Directory.Exists(vhdxFolder) Then
                Directory.CreateDirectory(vhdxFolder)
                Debug.WriteLine($"Created directory: {vhdxFolder}")
            End If

            ' Step 2: Copy VHDX file to selected drive:\VHDX
            Dim destinationPath As String = Path.Combine(vhdxFolder, vhdxFileName)

            ' Check if file already exists
            If File.Exists(destinationPath) Then
                Dim overwriteResult = MessageBox.Show($"File '{vhdxFileName}' already exists in {vhdxFolder}.{vbCrLf}Do you want to overwrite it?",
                                                      "File Exists",
                                                      MessageBoxButton.YesNo,
                                                      MessageBoxImage.Question)
                If overwriteResult <> MessageBoxResult.Yes Then
                    Return
                End If
            End If

            ' Check available space
            Dim sourceFileSize As Long = New FileInfo(sourcePath).Length
            If selectedDrive.AvailableFreeSpace < sourceFileSize Then
                MessageBox.Show($"Insufficient space on drive {selectedDrive.Name}.{vbCrLf}" &
                               $"Required: {sourceFileSize \ (1024 * 1024)} MB{vbCrLf}" &
                               $"Available: {selectedDrive.AvailableFreeSpace \ (1024 * 1024)} MB",
                               "Insufficient Space",
                               MessageBoxButton.OK,
                               MessageBoxImage.Error)
                Return
            End If

            ' Show progress: Copying VHDX with activity light
            UpdateProgress("Copying...", $"Copying VHDX to {vhdxFolder}")
            StartDismIndicator()

            ' Monitor copy progress with cancellation support
            Dim copyMonitorCts As New CancellationTokenSource()
            Dim copyMonitorTask = MonitorFileCopyAsync(sourcePath, destinationPath, sourceFileSize, copyMonitorCts.Token)

            ' Perform the file copy
            Dim copyTask = Task.Run(Sub()
                                        File.Copy(sourcePath, destinationPath, True)
                                    End Sub)

            ' Wait for copy to complete
            Await copyTask

            ' Stop monitoring
            copyMonitorCts.Cancel()
            Try
                Await Task.WhenAny(copyMonitorTask, Task.Delay(1000))
            Catch
                ' Ignore cancellation exceptions
            End Try

            StopDismIndicator()

            Debug.WriteLine($"VHDX copied to: {destinationPath}")

            UpdateProgressSuccess("Copied", $"VHDX file ready at {vhdxFolder}")
            Await Task.Delay(1500)

            ' Prompt to add drivers from this PC
            Dim addDriversPrompt = MessageBox.Show("Add drivers from this PC to the VHDX image?",
                                                   "Add Drivers",
                                                   MessageBoxButton.YesNo,
                                                   MessageBoxImage.Question,
                                                   MessageBoxResult.Yes)

            If addDriversPrompt = MessageBoxResult.Yes Then
                ' Generate unique driver export path and store it
                _driverExportPath = Path.Combine(Path.GetTempPath(), "VHDXLab_Drivers_" & guid.NewGuid().ToString("N"))

                ' Export and inject drivers
                Await ExportAndInjectDriversAsync(destinationPath)
            Else
                ResetProgress()
            End If

            ' Step 3: Create boot entry
            UpdateProgress("Creating...", "Creating BCD boot entry")
            StartDismIndicator()

            Dim guuid As String = Await CreateBootEntryAsync(description)

            StopDismIndicator()

            If String.IsNullOrWhiteSpace(guuid) Then
                UpdateProgressError("Failed", "No GUID returned from bcdedit")
                Await Task.Delay(3000)
                ' Cleanup driver folder even on failure
                CleanupDriverExportFolder()
                ResetProgress()
                MessageBox.Show("Failed to create boot entry. No GUID returned.", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            Debug.WriteLine($"Created boot entry with GUID: {guuid}")

            ' Step 4: Configure the boot entry with VHDX information
            UpdateProgress("Configuring...", "Setting VHDX boot parameters")
            StartDismIndicator()

            Dim success As Boolean = Await ConfigureVhdxBootEntryAsync(guuid, destinationPath)

            StopDismIndicator()

            If success Then
                UpdateProgressSuccess("Complete", $"Boot entry created: {description}")
                Await Task.Delay(3000)

                ' ✅ Cleanup driver export folder AFTER successful BCD entry creation
                CleanupDriverExportFolder()

                ResetProgress()

                MessageBox.Show($"Boot entry created successfully!{vbCrLf}GUID: {guuid}{vbCrLf}VHDX: {destinationPath}{vbCrLf}Description: {description}",
                                "Success", MessageBoxButton.OK, MessageBoxImage.Information)

                ' Refresh the BCD entries list
                Await LoadBcdEntriesAsync()
            Else
                UpdateProgressError("Config Failed", "Check debug output for details")
                Await Task.Delay(3000)

                ' ✅ Cleanup driver folder even if BCD configuration failed
                CleanupDriverExportFolder()

                ResetProgress()

                MessageBox.Show($"Boot entry was created but failed to configure VHDX settings.{vbCrLf}GUID: {guuid}{vbCrLf}Please check the debug output for details.",
                                "Configuration Error",
                                MessageBoxButton.OK, MessageBoxImage.Warning)

                ' Still refresh to show the created entry
                Await LoadBcdEntriesAsync()
            End If

        Catch ex As UnauthorizedAccessException
            StopDismIndicator()
            UpdateProgressError("Access Denied", ex.Message)
            Task.Delay(3000)
            ' ✅ Cleanup on error
            CleanupDriverExportFolder()
            ResetProgress()
            MessageBox.Show($"Access denied creating folder or copying file: {ex.Message}", "Access Denied",
                            MessageBoxButton.OK, MessageBoxImage.Error)
        Catch ex As IOException
            StopDismIndicator()
            UpdateProgressError("I/O Error", ex.Message)
            Task.Delay(3000)
            ' ✅ Cleanup on error
            CleanupDriverExportFolder()
            ResetProgress()
            MessageBox.Show($"Failed to copy VHDX file: {ex.Message}", "File Copy Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)
        Catch ex As Exception
            StopDismIndicator()
            UpdateProgressError("Error", ex.Message)
            Task.Delay(3000)
            ' ✅ Cleanup on error
            CleanupDriverExportFolder()
            ResetProgress()
            MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' ==================== FILE COPY MONITORING WITH PROGRESS ====================

    Private Async Function MonitorFileCopyAsync(sourcePath As String, destPath As String, expectedSize As Long, ct As CancellationToken) As Task
        Dim lastSize As Long = 0
        Dim unchangedCount As Integer = 0

        While Not ct.IsCancellationRequested
            Dim currentSize As Long = 0

            If File.Exists(destPath) Then
                Try
                    currentSize = Await Task.Run(Function()
                                                     Try
                                                         Return New FileInfo(destPath).Length
                                                     Catch
                                                         Return 0L
                                                     End Try
                                                 End Function, ct)
                Catch ex As OperationCanceledException
                    Exit While
                Catch
                    ' Continue monitoring even if file access fails
                End Try
            End If

            ' Detect stalled copy
            If currentSize = lastSize AndAlso currentSize > 0 Then
                unchangedCount += 1
            Else
                unchangedCount = 0
                lastSize = currentSize
            End If

            ' Update progress UI
            Await Dispatcher.InvokeAsync(Sub() UpdateFileCopyProgress(currentSize, expectedSize),
                                          DispatcherPriority.Background)

            ' Adaptive delay
            Dim delayMs = If(unchangedCount > 5, 2000, 500)

            Try
                Await Task.Delay(delayMs, ct)
            Catch ex As OperationCanceledException
                Exit While
            End Try
        End While

        ' Final update
        If Not ct.IsCancellationRequested AndAlso File.Exists(destPath) Then
            Dim finalSize = Await Task.Run(Function() New FileInfo(destPath).Length)
            Await Dispatcher.InvokeAsync(Sub() UpdateFileCopyProgress(finalSize, expectedSize))
        End If
    End Function

    Private Sub UpdateFileCopyProgress(bytesCopied As Long, totalBytes As Long)
        If MountSizeProgressText Is Nothing OrElse MountSizeProgressTextDetail Is Nothing Then
            Return
        End If

        If totalBytes <= 0 Then
            ' Unknown total size
            If Math.Abs(bytesCopied - _lastReportedBytes) > 1048576 Then ' Update every 1MB
                MountSizeProgressTextDetail.Text = $"{FormatBytes(bytesCopied)} copied"
                _lastReportedBytes = bytesCopied
            End If
            Return
        End If

        Dim pct = CInt(Math.Min(100.0, (bytesCopied * 100.0) / totalBytes))

        ' Debounce updates
        If pct <> _lastReportedPercentage OrElse
           Math.Abs(bytesCopied - _lastReportedBytes) > (totalBytes \ 100) Then

            MountSizeProgressText.Text = $"{pct}%"
            MountSizeProgressTextDetail.Text = $"{FormatBytes(bytesCopied)} / {FormatBytes(totalBytes)}"

            _lastReportedPercentage = pct
            _lastReportedBytes = bytesCopied
        End If
    End Sub

    ' ==================== EXPORT AND INJECT DRIVERS ====================

    Private Async Function ExportAndInjectDriversAsync(vhdxPath As String) As Task
        If Not IsAdministrator() Then
            MessageBox.Show("Administrator privileges required to export and inject drivers.", "Error",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Const mountPath As String = "C:\mount"

        Try
            ' ✅ Create driver export directory BEFORE DISM export
            If Not Directory.Exists(_driverExportPath) Then
                Directory.CreateDirectory(_driverExportPath)
                Debug.WriteLine($"Created driver export directory: {_driverExportPath}")
            End If

            ' Create mount directory if it doesn't exist
            If Not Directory.Exists(mountPath) Then
                Directory.CreateDirectory(mountPath)
            End If

            ' Step 1: Export drivers from current system
            UpdateProgress("Exporting...", "Capturing drivers from this PC")
            StartDismIndicator()

            Dim exportArgs = $"/online /export-driver /destination:""{_driverExportPath}"""
            Dim exportResult = Await RunDismCommandAsync(exportArgs)

            StopDismIndicator()

            If exportResult.ExitCode <> 0 Then
                UpdateProgressError("Export Failed", "Could not capture system drivers")
                Await Task.Delay(3000)
                ResetProgress()
                MessageBox.Show($"Failed to export drivers:{vbCrLf}{exportResult.StdErr}",
                               "Driver Export Failed",
                               MessageBoxButton.OK,
                               MessageBoxImage.Error)
                Return
            End If

            ' Count exported drivers
            Dim exportedDriverCount As Integer = 0
            If Directory.Exists(_driverExportPath) Then
                exportedDriverCount = Directory.GetFiles(_driverExportPath, "*.inf", SearchOption.AllDirectories).Length
            End If

            UpdateProgressSuccess("Exported", $"{exportedDriverCount} driver(s) captured")
            Await Task.Delay(1500)

            ' Step 2: Mount VHDX
            UpdateProgress("Mounting...", "Attaching VHDX image")
            StartDismIndicator()

            ' Ensure mount directory is empty
            If Directory.Exists(mountPath) AndAlso Directory.EnumerateFileSystemEntries(mountPath).Any() Then
                Try
                    For Each f In Directory.GetFiles(mountPath, "*", SearchOption.AllDirectories)
                        Try : File.Delete(f) : Catch : End Try
                    Next
                    For Each d In Directory.GetDirectories(mountPath, "*", SearchOption.AllDirectories)
                        Try : Directory.Delete(d, True) : Catch : End Try
                    Next
                Catch
                End Try
            End If

            ' Mount VHDX using DISM
            Dim mountArgs = $"/Mount-Image /ImageFile""{vhdxPath}"" /Index:1 /MountDir:""{mountPath}"""
            Dim mountResult = Await RunDismCommandAsync(mountArgs)

            StopDismIndicator()

            If mountResult.ExitCode <> 0 Then
                UpdateProgressError("Mount Failed", "Could Not attach VHDX")
                Await Task.Delay(3000)
                ResetProgress()
                MessageBox.Show($"Failed to mount VHDX{vbCrLf}{mountResult.StdErr}",
                               "Mount Failed",
                               MessageBoxButton.OK,
                               MessageBoxImage.Error)
                Return
            End If

            UpdateProgressSuccess("Mounted", "VHDX image attached")
            Await Task.Delay(1000)

            ' Step 3: Inject drivers into mounted image
            UpdateProgress("Injecting...", $"Adding {exportedDriverCount} driver(s)")
            StartDismIndicator()

            Dim addDriverArgs = $"/Image""{mountPath}"" /Add-Driver /Driver:""{_driverExportPath}"" /Recurse"
            Dim addDriverResult = Await RunDismCommandAsync(addDriverArgs)

            StopDismIndicator()

            If addDriverResult.ExitCode <> 0 Then
                UpdateProgressError("Injection Warning", "Some drivers may have failed")
                Await Task.Delay(2000)
                MessageBox.Show($"Warning: Some drivers may have failed to inject:{vbCrLf}{addDriverResult.StdErr}",
                               "Driver Injection Warning",
                               MessageBoxButton.OK,
                               MessageBoxImage.Warning)
            End If

            UpdateProgressSuccess("Injected", "Drivers added successfully")
            Await Task.Delay(1500)

            ' Step 4: Unmount and commit changes
            UpdateProgress("Committing...", "Saving changes to VHDX")
            StartDismIndicator()

            Dim unmountArgs = $"/Unmount-Image /MountDir:""{mountPath}"" /Commit"
            Dim unmountResult = Await RunDismCommandAsync(unmountArgs)

            StopDismIndicator()

            If unmountResult.ExitCode <> 0 Then
                UpdateProgressError("Unmount Failed", "Could not save changes")
                Await Task.Delay(3000)
                ResetProgress()
                MessageBox.Show($"Failed to unmount VHDX:{vbCrLf}{unmountResult.StdErr}",
                               "Unmount Failed",
                               MessageBoxButton.OK,
                               MessageBoxImage.Error)
                Return
            End If

            ' Cleanup mountpoints
            Await RunDismCommandAsync("/Cleanup-Mountpoints")

            ' Show success
            UpdateProgressSuccess("Complete", $"{exportedDriverCount} driver(s) integrated")
            Await Task.Delay(3000)
            ResetProgress()

        Catch ex As Exception
            UpdateProgressError("Error", ex.Message)
            Task.Delay(3000)
            ResetProgress()
            MessageBox.Show($"Error during driver export/injection:{vbCrLf}{ex.Message}",
                           "Error",
                           MessageBoxButton.OK,
                           MessageBoxImage.Error)
        End Try
    End Function

    ' Helper: Run DISM command with activity indicator
    Private Async Function RunDismCommandAsync(arguments As String) As Task(Of (ExitCode As Integer, StdOut As String, StdErr As String))
        Try
            Dim psi As New ProcessStartInfo("dism.exe", arguments) With {
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }

            Using p As Process = Process.Start(psi)
                Dim stdOutTask = p.StandardOutput.ReadToEndAsync()
                Dim stdErrTask = p.StandardError.ReadToEndAsync()
                Await p.WaitForExitAsync()
                Return (p.ExitCode, Await stdOutTask, Await stdErrTask)
            End Using

        Catch ex As Exception
            Return (-1, "", "Exception: " & ex.Message)
        End Try
    End Function

    ' Helper: Cleanup driver export folder
    Private Sub CleanupDriverExportFolder()
        Try
            If Not String.IsNullOrWhiteSpace(_driverExportPath) AndAlso Directory.Exists(_driverExportPath) Then
                UpdateProgress("Cleaning up...", "Removing temporary driver files")
                Directory.Delete(_driverExportPath, True)
                Debug.WriteLine($"Deleted driver export folder: {_driverExportPath}")
                UpdateProgressSuccess("Cleanup Complete", "Temporary files removed")
                _driverExportPath = Nothing ' Reset the path
            End If
        Catch ex As Exception
            Debug.WriteLine($"Failed to cleanup driver export folder: {ex.Message}")
            ' Don't show error to user - this is non-critical cleanup
        End Try
    End Sub

    ' ==================== EXISTING METHODS ====================

    Private Async Sub DeleteEntryButton_Click(sender As Object, e As RoutedEventArgs) Handles DeleteEntryButton.Click
        ' Check if an entry is selected
        If BCDEntrys.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a boot entry to delete.", "No Selection",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim selectedEntry = CType(BCDEntrys.SelectedItem, BcdEntry)

        ' Prevent deletion of system entries
        Dim protectedIdentifiers As String() = {"{bootmgr}", "{current}", "{default}", "{ntldr}", "{memdiag}"}
        If protectedIdentifiers.Contains(selectedEntry.Identifier.ToLower()) Then
            MessageBox.Show("Cannot delete system boot entries. This entry is protected.", "Protected Entry",
                            MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End If

        ' Confirm deletion
        Dim result = MessageBox.Show($"Are you sure you want to delete this boot entry?{vbCrLf}{vbCrLf}" &
                                     $"Description: {selectedEntry.DisplayName}{vbCrLf}" &
                                     $"GUID: {selectedEntry.Identifier}{vbCrLf}{vbCrLf}" &
                                     "This action cannot be undone.",
                                     "Confirm Deletion",
                                     MessageBoxButton.YesNo,
                                     MessageBoxImage.Warning)

        If result <> MessageBoxResult.Yes Then
            Return
        End If

        Try
            UpdateProgress("Deleting...", $"Removing boot entry {selectedEntry.Identifier}")
            StartDismIndicator()

            Dim success As Boolean = Await DeleteBootEntryAsync(selectedEntry.Identifier)

            StopDismIndicator()

            If success Then
                UpdateProgressSuccess("Deleted", "Boot entry removed")
                Await Task.Delay(2000)
                ResetProgress()

                MessageBox.Show($"Boot entry deleted successfully!{vbCrLf}GUID: {selectedEntry.Identifier}",
                                "Success", MessageBoxButton.OK, MessageBoxImage.Information)

                ' Refresh the BCD entries list
                Await LoadBcdEntriesAsync()
            Else
                UpdateProgressError("Delete Failed", "Could not remove entry")
                Await Task.Delay(3000)
                ResetProgress()
                MessageBox.Show("Failed to delete boot entry.", "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error)
            End If

        Catch ex As Exception
            StopDismIndicator()
            UpdateProgressError("Error", ex.Message)
            Task.Delay(3000)
            ResetProgress()
            MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Async Function DeleteBootEntryAsync(guid As String) As Task(Of Boolean)
        Return Await Task.Run(Function() As Boolean
                                  Try
                                      Dim psi As New ProcessStartInfo With {
                                          .FileName = "bcdedit.exe",
                                          .Arguments = $"/delete {guid}",
                                          .RedirectStandardOutput = True,
                                          .RedirectStandardError = True,
                                          .UseShellExecute = False,
                                          .CreateNoWindow = True
                                      }

                                      Using proc As Process = Process.Start(psi)
                                          Dim output As String = proc.StandardOutput.ReadToEnd()
                                          Dim errorOutput As String = proc.StandardError.ReadToEnd()
                                          proc.WaitForExit()

                                          If proc.ExitCode = 0 Then
                                              Debug.WriteLine($"Successfully deleted entry {guid}")
                                              Return True
                                          Else
                                              Debug.WriteLine($"bcdedit /delete failed: {errorOutput}")
                                              Return False
                                          End If
                                      End Using
                                  Catch ex As Exception
                                      Debug.WriteLine($"DeleteBootEntryAsync error: {ex.Message}")
                                      Return False
                                  End Try
                              End Function)
    End Function

    Private Async Function CreateBootEntryAsync(description As String) As Task(Of String)
        Return Await Task.Run(Function() As String
                                  Try
                                      Dim psi As New ProcessStartInfo With {
                                          .FileName = "bcdedit.exe",
                                          .Arguments = $"/copy {{current}} /d ""{description}""",
                                          .RedirectStandardOutput = True,
                                          .RedirectStandardError = True,
                                          .UseShellExecute = False,
                                          .CreateNoWindow = True
                                      }

                                      Using proc As Process = Process.Start(psi)
                                          Dim output As String = proc.StandardOutput.ReadToEnd()
                                          Dim errorOutput As String = proc.StandardError.ReadToEnd()
                                          proc.WaitForExit()

                                          Debug.WriteLine($"bcdedit /copy output: {output}")
                                          If Not String.IsNullOrWhiteSpace(errorOutput) Then
                                              Debug.WriteLine($"bcdedit /copy error: {errorOutput}")
                                          End If

                                          If proc.ExitCode = 0 Then
                                              Dim match = Regex.Match(output, "\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}")
                                              If match.Success Then
                                                  Debug.WriteLine($"Extracted GUID: {match.Value}")
                                                  Return match.Value
                                              Else
                                                  Debug.WriteLine("Failed to extract GUID from output")
                                              End If
                                          Else
                                              Debug.WriteLine($"bcdedit /copy failed with exit code: {proc.ExitCode}")
                                          End If
                                      End Using
                                  Catch ex As Exception
                                      Debug.WriteLine($"CreateBootEntryAsync error: {ex.Message}")
                                  End Try
                                  Return String.Empty
                              End Function)
    End Function

    Private Async Function ConfigureVhdxBootEntryAsync(guid As String, vhdxPath As String) As Task(Of Boolean)
        Return Await Task.Run(Function() As Boolean
                                  Try
                                      Dim driveLetter As String = Path.GetPathRoot(vhdxPath).TrimEnd("\"c)
                                      Dim pathWithoutDrive As String = vhdxPath.Substring(Path.GetPathRoot(vhdxPath).Length)
                                      Dim vhdxFormatted As String = $"[{driveLetter}]\{pathWithoutDrive}"

                                      Debug.WriteLine($"Formatted VHDX path: vhd={vhdxFormatted}")

                                      Dim commands As String() = {
                                          $"/set {guid} device vhd={vhdxFormatted}",
                                          $"/set {guid} osdevice vhd={vhdxFormatted}",
                                          $"/set {guid} detecthal on"
                                      }

                                      For i As Integer = 0 To commands.Length - 1
                                          Dim command As String = commands(i)
                                          Debug.WriteLine($"Executing command {i + 1}/{commands.Length}: bcdedit {command}")

                                          Dim psi As New ProcessStartInfo With {
                                              .FileName = "bcdedit.exe",
                                              .Arguments = command,
                                              .RedirectStandardOutput = True,
                                              .RedirectStandardError = True,
                                              .UseShellExecute = False,
                                              .CreateNoWindow = True
                                          }

                                          Using proc As Process = Process.Start(psi)
                                              Dim output As String = proc.StandardOutput.ReadToEnd()
                                              Dim errorOutput As String = proc.StandardError.ReadToEnd()
                                              proc.WaitForExit()

                                              Debug.WriteLine($"Command {i + 1} output: {output}")
                                              If Not String.IsNullOrWhiteSpace(errorOutput) Then
                                                  Debug.WriteLine($"Command {i + 1} error: {errorOutput}")
                                              End If

                                              If proc.ExitCode <> 0 Then
                                                  Debug.WriteLine($"Command {i + 1} failed with exit code: {proc.ExitCode}")
                                                  Debug.WriteLine($"Full command: bcdedit {command}")
                                                  Return False
                                              Else
                                                  Debug.WriteLine($"Command {i + 1} completed successfully")
                                              End If
                                          End Using
                                      Next

                                      Debug.WriteLine("All VHDX configuration commands completed successfully")
                                      Return True
                                  Catch ex As Exception
                                      Debug.WriteLine($"ConfigureVhdxBootEntryAsync error: {ex.Message}")
                                      Debug.WriteLine($"Stack trace: {ex.StackTrace}")
                                      Return False
                                  End Try
                              End Function)
    End Function

    Private Async Function LoadBcdEntriesAsync() As Task
        UpdateProgress("Loading...", "Reading BCD boot entries")
        StartDismIndicator()

        Await Task.Run(Sub()
                           Try
                               Dim psi As New ProcessStartInfo With {
                                   .FileName = "bcdedit.exe",
                                   .Arguments = "/enum",
                                   .RedirectStandardOutput = True,
                                   .RedirectStandardError = True,
                                   .UseShellExecute = False,
                                   .CreateNoWindow = True
                               }

                               Using proc As Process = Process.Start(psi)
                                   Dim output As String = proc.StandardOutput.ReadToEnd()
                                   proc.WaitForExit()

                                   If proc.ExitCode = 0 Then
                                       ParseBcdOutput(output)
                                   End If
                               End Using
                           Catch ex As Exception
                               Dispatcher.Invoke(Sub()
                                                     MessageBox.Show("Failed to load BCD entries: " & ex.Message, "Error",
                                                                     MessageBoxButton.OK, MessageBoxImage.Error)
                                                 End Sub)
                           End Try
                       End Sub)

        StopDismIndicator()

        Dispatcher.Invoke(Sub()
                              BCDEntrys.ItemsSource = Nothing
                              BCDEntrys.ItemsSource = bcdEntries
                              If bcdEntries.Count > 0 Then
                                  BCDEntrys.SelectedIndex = 0
                              End If
                          End Sub)

        UpdateProgressSuccess("Loaded", $"{bcdEntries.Count} boot entries found")
        Await Task.Delay(1500)
        ResetProgress()
    End Function

    Private Sub ParseBcdOutput(output As String)
        bcdEntries.Clear()
        Dim entries = output.Split(New String() {vbCrLf & vbCrLf, vbLf & vbLf}, StringSplitOptions.RemoveEmptyEntries)

        For Each entryBlock As String In entries
            If entryBlock.Contains("identifier") OrElse entryBlock.Contains("Identifier") Then
                Dim entry As New BcdEntry()
                Dim lines = entryBlock.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                For Each line As String In lines
                    Dim trimmedLine = line.Trim()

                    If trimmedLine.StartsWith("identifier", StringComparison.OrdinalIgnoreCase) Then
                        Dim parts = trimmedLine.Split(New Char() {" "c, vbTab}, StringSplitOptions.RemoveEmptyEntries)
                        If parts.Length >= 2 Then
                            entry.Identifier = parts(1)
                        End If
                    End If

                    If trimmedLine.StartsWith("description", StringComparison.OrdinalIgnoreCase) Then
                        Dim descIndex = trimmedLine.IndexOf(" ", StringComparison.OrdinalIgnoreCase)
                        If descIndex > 0 AndAlso descIndex < trimmedLine.Length - 1 Then
                            entry.DisplayName = trimmedLine.Substring(descIndex + 1).Trim()
                        End If
                    End If
                Next

                entry.FullDetails = entryBlock
                If Not String.IsNullOrWhiteSpace(entry.Identifier) Then
                    bcdEntries.Add(entry)
                End If
            End If
        Next
    End Sub

    Private Sub BCDEntrys_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles BCDEntrys.SelectionChanged
        If BCDEntrys.SelectedItem IsNot Nothing Then
            Dim selectedEntry = CType(BCDEntrys.SelectedItem, BcdEntry)
            GUIDDetail.Text = selectedEntry.FullDetails
        Else
            GUIDDetail.Text = String.Empty
        End If
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

    Private Async Function SetBingWallpaperAsync(Optional ct As CancellationToken = Nothing) As Task
        If Not Await IsInternetAvailableAsync(ct) Then Return
        Const xmlUrl As String = "https://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1&mkt=en-US"
        Dim xmlContent As String
        Try
            Using resp = Await httpClient.GetAsync(xmlUrl, ct)
                resp.EnsureSuccessStatusCode()
                xmlContent = Await resp.Content.ReadAsStringAsync(ct)
            End Using
        Catch
            Return
        End Try

        Dim doc As XDocument
        Try
            doc = XDocument.Parse(xmlContent)
        Catch
            Return
        End Try

        Dim imageElement = doc.Root?.Element("image")
        If imageElement Is Nothing Then Return

        Dim relativeUrl = imageElement.Element("url")?.Value
        If String.IsNullOrWhiteSpace(relativeUrl) Then Return

        Dim fullImageUrl = If(relativeUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase),
                              relativeUrl,
                              "https://www.bing.com" & relativeUrl)

        Dim imageBytes As Byte()
        Try
            imageBytes = Await httpClient.GetByteArrayAsync(fullImageUrl, ct)
        Catch
            Return
        End Try

        Dim bmp As New BitmapImage()
        Try
            Using ms As New MemoryStream(imageBytes)
                bmp.BeginInit()
                bmp.CacheOption = BitmapCacheOption.OnLoad
                bmp.StreamSource = ms
                bmp.EndInit()
                bmp.Freeze()
            End Using
        Catch
            Return
        End Try

        Me.Background = New ImageBrush(bmp) With {.Stretch = Stretch.UniformToFill, .AlignmentX = AlignmentX.Center, .AlignmentY = AlignmentY.Center}

        Dim headline = imageElement.Element("headline")?.Value
        Dim copyright = imageElement.Element("copyright")?.Value

        Await Dispatcher.InvokeAsync(Sub()
                                         HeadingTextBlock.Text = If(String.IsNullOrWhiteSpace(headline), "Description", headline)
                                         CopyrightTextBlock.Text = If(String.IsNullOrWhiteSpace(copyright), "Detail", copyright)
                                     End Sub)
    End Function

    Private Shared Async Function IsInternetAvailableAsync(Optional ct As CancellationToken = Nothing) As Task(Of Boolean)
        Try
            Using resp = Await httpClient.GetAsync("https://www.bing.com", HttpCompletionOption.ResponseHeadersRead, ct)
                Return resp.IsSuccessStatusCode
            End Using
        Catch
            Return False
        End Try
    End Function

    Private Sub EnableDarkTitleBar()
        Try
            Dim hwnd = New System.Windows.Interop.WindowInteropHelper(Me).Handle
            Dim useImmersiveDarkMode As Integer = 20 ' DWMWA_USE_IMMERSIVE_DARK_MODE
            Dim value As Integer = 1
            DwmSetWindowAttribute(hwnd, useImmersiveDarkMode, value, Marshal.SizeOf(value))
        Catch ex As Exception
            Debug.WriteLine($"Failed to set dark title bar: {ex.Message}")
        End Try
    End Sub

    <Runtime.InteropServices.DllImport("dwmapi.dll", PreserveSig:=True)>
    Private Shared Function DwmSetWindowAttribute(hwnd As IntPtr, attr As Integer, ByRef attrValue As Integer, attrSize As Integer) As Integer
    End Function
End Class