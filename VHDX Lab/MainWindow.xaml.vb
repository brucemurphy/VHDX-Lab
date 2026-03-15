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
    Private _bootButtonHooked As Boolean = False
    Private _commandLogTarget As TextBox
    Private Const EnableCommandLogPanel As Boolean = False

    Private Enum CreationMode
        Folder
        BootableWim
    End Enum

    Private _currentCreationMode As CreationMode = CreationMode.Folder

    Private Function EnsureUniqueVhdxPath(desiredPath As String) As String
        Dim fallbackName = "Image"
        If String.IsNullOrWhiteSpace(desiredPath) Then
            Return fallbackName & ".vhdx"
        End If

        Dim directoryPart = Path.GetDirectoryName(desiredPath)
        Dim nameWithoutExt = Path.GetFileNameWithoutExtension(desiredPath)
        Dim extension = Path.GetExtension(desiredPath)

        If String.IsNullOrWhiteSpace(nameWithoutExt) Then
            nameWithoutExt = fallbackName
        End If

        If Not extension.Equals(".vhdx", StringComparison.OrdinalIgnoreCase) Then
            extension = ".vhdx"
        End If

        Dim candidate = nameWithoutExt & extension
        If String.IsNullOrWhiteSpace(directoryPart) Then
            Return candidate
        End If

        Dim fullCandidate = Path.Combine(directoryPart, candidate)
        Dim counter = 1
        While File.Exists(fullCandidate)
            Dim numberedName = $"{nameWithoutExt} ({counter}){extension}"
            fullCandidate = Path.Combine(directoryPart, numberedName)
            counter += 1
        End While

        Return fullCandidate
    End Function

    Private Function PromptForFreshVhdxPath(suggestedPath As String) As String
        Dim dialog As New SaveFileDialog With {
            .Filter = "VHDX Files (*.vhdx)|*.vhdx",
            .Title = "Save VHDX",
            .OverwritePrompt = False,
            .CheckFileExists = False
        }

        Dim seedPath = EnsureUniqueVhdxPath(suggestedPath)
        Dim seedDir = Path.GetDirectoryName(seedPath)
        Dim seedName = Path.GetFileName(seedPath)

        If Not String.IsNullOrWhiteSpace(seedDir) Then
            dialog.InitialDirectory = seedDir
        End If

        If Not String.IsNullOrWhiteSpace(seedName) Then
            dialog.FileName = seedName
        End If

        Do
            Dim result = dialog.ShowDialog()
            If result <> True Then
                Return Nothing
            End If

            Dim selectedPath = dialog.FileName
            If String.IsNullOrWhiteSpace(selectedPath) Then
                Return Nothing
            End If

            If Not selectedPath.EndsWith(".vhdx", StringComparison.OrdinalIgnoreCase) Then
                selectedPath &= ".vhdx"
            End If

            If File.Exists(selectedPath) Then
                MessageBox.Show("VHDX files cannot overwrite existing files. Please choose a new file name.", "Destination Exists", MessageBoxButton.OK, MessageBoxImage.Warning)
                dialog.FileName = Path.GetFileName(selectedPath)
                Continue Do
            End If

            Return selectedPath
        Loop
    End Function

    Private Function GetFreshVhdxDestinationPath() As String
        Dim currentPath = SavePathTextBox.Text

        If Not String.IsNullOrWhiteSpace(currentPath) AndAlso File.Exists(currentPath) Then
            MessageBox.Show("The selected VHDX already exists. Please choose a new file name before continuing.", "Destination Exists", MessageBoxButton.OK, MessageBoxImage.Warning)
            currentPath = Nothing
        End If

        If String.IsNullOrWhiteSpace(currentPath) Then
            BrowseSavePathButton_Click(Nothing, Nothing)
            currentPath = SavePathTextBox.Text
        End If

        If String.IsNullOrWhiteSpace(currentPath) Then
            Return Nothing
        End If

        If File.Exists(currentPath) Then
            MessageBox.Show("The selected VHDX already exists. Please choose a new file name before continuing.", "Destination Exists", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return Nothing
        End If

        Return currentPath
    End Function

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
        If _currentCreationMode = CreationMode.BootableWim Then
            Dim wimDialog As New OpenFileDialog With {
                .Filter = "WIM Files (*.wim)|*.wim|All Files (*.*)|*.*",
                .Title = "Select WIM Image",
                .CheckFileExists = True
            }

            If wimDialog.ShowDialog() = True Then
                SetWimSource(wimDialog.FileName)
            End If
            Return
        End If

        Dim folderDialog As New Forms.FolderBrowserDialog With {.Description = "Select the folder to package into a VHDX"}
        If folderDialog.ShowDialog() = Forms.DialogResult.OK Then
            Await SetSourceFolderAsync(folderDialog.SelectedPath)
        End If
    End Sub

    Private Sub BrowseSavePathButton_Click(sender As Object, e As RoutedEventArgs) Handles BrowseSavePathButton.Click
        Dim defaultName = "Image.vhdx"
        Dim initialDir As String = Nothing

        If _currentCreationMode = CreationMode.BootableWim AndAlso File.Exists(SourceFolderTextBox.Text) Then
            defaultName = Path.GetFileNameWithoutExtension(SourceFolderTextBox.Text) & "_bootable.vhdx"
            initialDir = Path.GetDirectoryName(SourceFolderTextBox.Text)
        ElseIf Not String.IsNullOrWhiteSpace(SourceFolderTextBox.Text) Then
            Dim trimmed = SourceFolderTextBox.Text.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            If Directory.Exists(trimmed) Then
                defaultName = Path.GetFileName(trimmed) & ".vhdx"
                initialDir = Path.GetDirectoryName(trimmed)
            End If
        End If

        Dim suggestedPath = defaultName
        If Not String.IsNullOrWhiteSpace(initialDir) Then
            suggestedPath = Path.Combine(initialDir, defaultName)
        End If

        Dim selectedPath = PromptForFreshVhdxPath(suggestedPath)
        If Not String.IsNullOrWhiteSpace(selectedPath) Then
            SavePathTextBox.Text = selectedPath
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

        If String.IsNullOrWhiteSpace(SavePathTextBox.Text) OrElse File.Exists(SavePathTextBox.Text) Then
            Dim trimmed = sourcePath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            Dim defaultName = System.IO.Path.GetFileName(trimmed) & ".vhdx"
            Dim defaultDir = System.IO.Path.GetDirectoryName(trimmed)
            Dim candidate = If(Not String.IsNullOrWhiteSpace(defaultDir), System.IO.Path.Combine(defaultDir, defaultName), defaultName)
            SavePathTextBox.Text = EnsureUniqueVhdxPath(candidate)
        End If
    End Function

    Private Sub SetWimSource(wimPath As String)
        If String.IsNullOrWhiteSpace(wimPath) OrElse Not File.Exists(wimPath) Then
            MessageBox.Show("Please select a valid WIM file.", "Invalid Source", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        If Not wimPath.EndsWith(".wim", StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show("The selected file is not a WIM image.", "Invalid Source", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        SourceFolderTextBox.Text = wimPath

        Dim imageBytes As Long = 0
        Try
            imageBytes = New FileInfo(wimPath).Length
        Catch
        End Try

        Dim suggestedSizeGb As Double = Math.Max(20, Math.Ceiling((imageBytes * 2.5) / (1024.0 * 1024 * 1024)))
        VhdxSizeTextBox.Text = suggestedSizeGb.ToString("0")

        If String.IsNullOrWhiteSpace(SavePathTextBox.Text) OrElse File.Exists(SavePathTextBox.Text) Then
            Dim defaultName = Path.GetFileNameWithoutExtension(wimPath) & "_bootable.vhdx"
            Dim defaultDir = Path.GetDirectoryName(wimPath)
            Dim candidate = If(Not String.IsNullOrWhiteSpace(defaultDir), Path.Combine(defaultDir, defaultName), defaultName)
            SavePathTextBox.Text = EnsureUniqueVhdxPath(candidate)
        End If
    End Sub

    Private Sub CreateVhdxFromFolderButton_Click(sender As Object, e As RoutedEventArgs) Handles CreateVhdxFromFolderButton.Click
        If Not IsAdministrator() Then
            MessageBox.Show("Administrator privileges are required to create a VHDX.", "Administrator Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ConfigureCreatePaneForMode(CreationMode.Folder)
        CreatePane.Visibility = Visibility.Visible
    End Sub

    Private Sub ConfigureCreatePaneForMode(mode As CreationMode)
        Dim modeChanged = _currentCreationMode <> mode
        _currentCreationMode = mode

        If SourcePathLabel IsNot Nothing Then
            SourcePathLabel.Text = If(mode = CreationMode.Folder, "Source Folder", "WIM Image")
        End If
        Dim sourceLabel = TryCast(Me.FindName("SourcePathLabel"), TextBlock)
        If sourceLabel IsNot Nothing Then
            sourceLabel.Text = If(mode = CreationMode.Folder, "Source Folder", "WIM Image")
        End If

        If modeChanged Then
            SourceFolderTextBox?.Clear()
            If mode = CreationMode.BootableWim Then
                SavePathTextBox?.Clear()
            End If
        End If

        If FileSystemCombo IsNot Nothing Then
            If mode = CreationMode.BootableWim Then
                FileSystemCombo.SelectedIndex = 1 ' NTFS
            ElseIf FileSystemCombo.SelectedIndex < 0 Then
                FileSystemCombo.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub CreateBootableVhdxButton_Click(sender As Object, e As RoutedEventArgs)
        If Not IsAdministrator() Then
            MessageBox.Show("Administrator privileges are required to create a bootable VHDX.", "Administrator Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ConfigureCreatePaneForMode(CreationMode.BootableWim)
        CreatePane.Visibility = Visibility.Visible
    End Sub

    Private Async Sub StartCreateVhdxButton_Click(sender As Object, e As RoutedEventArgs) Handles StartCreateVhdxButton.Click
        If _currentCreationMode = CreationMode.BootableWim Then
            Await RunBootableVhdxWorkflowAsync()
        Else
            Await RunCreateVhdxWorkflowAsync()
        End If
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

        Dim vhdxPath As String = GetFreshVhdxDestinationPath()
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

    Private Async Function RunBootableVhdxWorkflowAsync() As Task
        Dim wimPath = SourceFolderTextBox.Text
        ResetCommandLog()
        AppendCommandLog("Bootable VHDX workflow started.")

        If String.IsNullOrWhiteSpace(wimPath) OrElse Not File.Exists(wimPath) Then
            AppendCommandLog("No valid WIM path provided. Prompting user to select one.")
            BrowseSourceFolderButton_Click(Nothing, Nothing)
            Return
        End If

        AppendCommandLog($"Source WIM: {wimPath}")

        If Not wimPath.EndsWith(".wim", StringComparison.OrdinalIgnoreCase) Then
            AppendCommandLog("Selected file does not have .wim extension. Aborting.")
            MessageBox.Show("Please select a valid WIM image.", "Invalid Source", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim vhdxPath As String = GetFreshVhdxDestinationPath()
        If String.IsNullOrWhiteSpace(vhdxPath) Then
            AppendCommandLog("Destination path is empty or invalid after prompting user.")
            MessageBox.Show("Please choose a destination path.", "Missing Destination", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        AppendCommandLog($"Destination VHDX: {vhdxPath}")

        Dim sizeGb As Double
        AppendCommandLog($"Requested size: {VhdxSizeTextBox.Text} GB")
        If Not Double.TryParse(VhdxSizeTextBox.Text, sizeGb) OrElse sizeGb < 10 Then
            AppendCommandLog("Invalid size supplied. Requires at least 10 GB.")
            MessageBox.Show("Enter a valid VHDX size (minimum 10 GB).", "Invalid Size", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim sizeBytes As Double = sizeGb * 1024 * 1024 * 1024
        Dim sizeMb As Long = CLng(Math.Ceiling(sizeBytes / (1024 * 1024)))
        If sizeMb < 2048 Then
            AppendCommandLog("Calculated size below 2 GB. Cannot continue.")
            MessageBox.Show("Bootable images require at least 2 GB.", "Size Too Small", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        AppendCommandLog($"Calculated size: {sizeMb} MB")

        If PartitionStyleCombo.SelectedIndex < 0 Then PartitionStyleCombo.SelectedIndex = 0
        Dim partStyle As String = TryCast((TryCast(PartitionStyleCombo.SelectedItem, ComboBoxItem))?.Content, String)
        partStyle = If(String.IsNullOrWhiteSpace(partStyle), "GPT", partStyle.Trim().ToUpperInvariant())
        AppendCommandLog($"Partition style requested: {partStyle}")
        If partStyle <> "GPT" AndAlso partStyle <> "MBR" Then
            AppendCommandLog("Unsupported partition style. Aborting.")
            MessageBox.Show("Invalid partition style. Choose GPT or MBR.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim letters = GetAvailableDriveLetters(2)
        If letters.Count < 2 Then
            AppendCommandLog("Insufficient free drive letters for system and OS partitions.")
            MessageBox.Show("Two free drive letters are required to build a bootable image.", "Drive Letter Unavailable", MessageBoxButton.YesNo, MessageBoxImage.Warning)
            Return
        End If

        Dim osLetter = letters(0)
        Dim systemLetter = letters(1)
        AppendCommandLog($"Assigned letters -> System: {systemLetter}: OS: {osLetter}:")

        Dim imageIndexInput = InputBox("Enter the WIM image index to apply:", "WIM Image Index", "1")
        If String.IsNullOrWhiteSpace(imageIndexInput) Then
            imageIndexInput = "1"
        End If

        Dim imageIndex As Integer
        If Not Integer.TryParse(imageIndexInput, imageIndex) OrElse imageIndex < 1 Then
            AppendCommandLog("Invalid WIM index entered. Stopping workflow.")
            MessageBox.Show("Please enter a valid WIM index (1 or greater).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If
        AppendCommandLog($"Using WIM index {imageIndex}")

        Dim createScriptPath As String = Nothing
        Dim caughtException As Exception = Nothing
        Try
            createScriptPath = Path.GetTempFileName()
            Dim scriptLines = BuildBootableDiskpartScript(vhdxPath, sizeMb, partStyle, osLetter, systemLetter)
            File.WriteAllLines(createScriptPath, scriptLines)
            AppendCommandLog($"DiskPart script saved to {createScriptPath}")
            AppendCommandLog("DiskPart commands:")
            For Each line In scriptLines
                AppendCommandLog($"  {line}")
            Next

            UpdateProgress("Creating...", "Preparing bootable VHDX")
            StartDismIndicator()
            AppendCommandLog($"Executing diskpart.exe /s ""{createScriptPath}""")
            Dim dpResult = Await RunDiskpartScriptAsync(createScriptPath)
            StopDismIndicator()
            AppendCommandLog($"DiskPart exit code: {dpResult.ExitCode}")
            LogProcessOutput("diskpart.exe", dpResult.StdOut, dpResult.StdErr)

            If dpResult.ExitCode <> 0 Then
                AppendCommandLog($"DiskPart stderr: {dpResult.StdErr}")
                UpdateProgressError("DiskPart Error", dpResult.StdErr)
                Await DetachVhdxAsync(vhdxPath)
                Await Task.Delay(2000)
                ResetProgress()
                MessageBox.Show($"Failed to create VHDX:{vbCrLf}{dpResult.StdErr}", "DiskPart Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            Dim osDrive = $"{osLetter}:"
            Dim systemDrive = $"{systemLetter}:"
            Dim applyDir = osDrive & Path.DirectorySeparatorChar

            UpdateProgress("Applying...", "Applying WIM image")
            StartDismIndicator()
            Dim applyArgs = String.Join(" ",
                                        "/Apply-Image",
                                        $"/ImageFile:{QuoteArgument(wimPath)}",
                                        $"/Index:{imageIndex}",
                                        $"/ApplyDir:{QuoteArgument(applyDir)}")
            AppendCommandLog($"Executing dism.exe {applyArgs}")
            Dim applyResult = Await RunDismCommandAsync(applyArgs)
            StopDismIndicator()
            AppendCommandLog($"DISM exit code: {applyResult.ExitCode}")
            LogProcessOutput("dism.exe", applyResult.StdOut, applyResult.StdErr)

            If applyResult.ExitCode <> 0 Then
                AppendCommandLog($"DISM stderr: {applyResult.StdErr}")
                UpdateProgressError("Apply Failed", applyResult.StdErr)
                Await DetachVhdxAsync(vhdxPath)
                Await Task.Delay(2000)
                ResetProgress()
                MessageBox.Show($"Failed to apply WIM image:{vbCrLf}{applyResult.StdErr}", "DISM Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            UpdateProgress("Boot Files...", "Configuring boot files")
            StartDismIndicator()
            Dim firmwareSwitch = If(partStyle = "GPT", "/f UEFI", "/f BIOS")
            Dim windowsDir = osDrive & "\Windows"
            Dim bcdArgs = $"""{windowsDir}"" /s {systemDrive} {firmwareSwitch}"
            AppendCommandLog($"Executing bcdboot.exe {bcdArgs}")
            Dim bcdResult = Await RunProcessAsync("bcdboot.exe", bcdArgs)
            StopDismIndicator()
            AppendCommandLog($"BCDBoot exit code: {bcdResult.ExitCode}")
            LogProcessOutput("bcdboot.exe", bcdResult.StdOut, bcdResult.StdErr)

            If bcdResult.ExitCode <> 0 Then
                AppendCommandLog($"BCDBoot stderr: {bcdResult.StdErr}")
                UpdateProgressError("BCDBoot Failed", bcdResult.StdErr)
                Await DetachVhdxAsync(vhdxPath)
                Await Task.Delay(2000)
                ResetProgress()
                MessageBox.Show($"Failed to configure boot files:{vbCrLf}{bcdResult.StdErr}", "BCDBoot Error", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            UpdateProgress("Detaching...", "Finalizing VHDX")
            StartDismIndicator()
            AppendCommandLog("Detaching VHDX via DiskPart.")
            Await DetachVhdxAsync(vhdxPath)
            StopDismIndicator()

            UpdateProgressSuccess("Complete", $"Bootable VHDX saved to {vhdxPath}")
            Await Task.Delay(2000)
            ResetProgress()

            AppendCommandLog("Bootable VHDX workflow completed successfully.")
            MessageBox.Show($"Bootable VHDX created successfully!{vbCrLf}WIM: {wimPath}{vbCrLf}Destination: {vhdxPath}{vbCrLf}Partition: {partStyle}", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            CreatePane.Visibility = Visibility.Collapsed
        Catch ex As Exception
            AppendCommandLog($"Workflow exception: {ex.Message}")
            caughtException = ex
        Finally
            If Not String.IsNullOrWhiteSpace(createScriptPath) AndAlso File.Exists(createScriptPath) Then
                Try
                    File.Delete(createScriptPath)
                    AppendCommandLog($"Deleted temporary DiskPart script {createScriptPath}")
                Catch
                    AppendCommandLog($"Failed to delete temporary script {createScriptPath}")
                End Try
            End If
        End Try

        If caughtException IsNot Nothing Then
            StopDismIndicator()
            UpdateProgressError("Error", caughtException.Message)
            Await Task.Delay(2000)
            ResetProgress()
            Try
                AppendCommandLog("Attempting to detach VHDX after failure.")
                Await DetachVhdxAsync(vhdxPath)
            Catch ex As Exception
                AppendCommandLog($"Detach attempt failed: {ex.Message}")
            End Try
            MessageBox.Show($"Failed to create bootable VHDX: {caughtException.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Function

    Private Function BuildBootableDiskpartScript(vhdxPath As String, sizeMb As Long, partStyle As String, osLetter As Char, systemLetter As Char) As List(Of String)
        Dim lines As New List(Of String) From {
            $"create vdisk file=""{vhdxPath}"" maximum={sizeMb} type=expandable",
            $"select vdisk file=""{vhdxPath}""",
            "attach vdisk",
            If(partStyle = "MBR", "convert mbr", "convert gpt")
        }

        If partStyle = "GPT" Then
            lines.AddRange(New String() {
                             "create partition efi size=100",
                             "format quick fs=fat32 label=""SYSTEM""",
                             $"assign letter={systemLetter}",
                             "create partition msr size=16",
                             "create partition primary",
                             "format quick fs=ntfs label=""Windows""",
                             $"assign letter={osLetter}"
                         })
        Else
            lines.AddRange(New String() {
                             "create partition primary size=500",
                             "format quick fs=ntfs label=""System""",
                             $"assign letter={systemLetter}",
                             "active",
                             "create partition primary",
                             "format quick fs=ntfs label=""Windows""",
                             $"assign letter={osLetter}"
                         })
        End If

        lines.Add("exit")
        Return lines
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

    Private Function GetAvailableDriveLetters(count As Integer) As List(Of Char)
        Dim available As New List(Of Char)()
        Dim usedLetters As New HashSet(Of Char)(DriveInfo.GetDrives().Select(Function(d) Char.ToUpperInvariant(d.Name(0))))

        For letterCode As Integer = Asc("Z"c) To Asc("D"c) Step -1
            Dim letter = Char.ToUpperInvariant(Chr(letterCode))
            If Not usedLetters.Contains(letter) Then
                available.Add(letter)
                If available.Count = count Then Exit For
            End If
        Next

        Return available
    End Function

    Private Function GetAvailableDriveLetter() As String
        Dim letters = GetAvailableDriveLetters(1)
        If letters.Count > 0 Then
            Return letters(0).ToString()
        End If
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

        If Not _bootButtonHooked Then
            Dim bootButton = TryCast(Me.FindName("CreateBootableVhdxButton"), Button)
            If bootButton IsNot Nothing Then
                AddHandler bootButton.Click, AddressOf CreateBootableVhdxButton_Click
                _bootButtonHooked = True
            End If
        End If

        _commandLogTarget = TryCast(Me.FindName("CommandLogTextBox"), TextBox)

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

    Private Sub ResetCommandLog()
        Dispatcher.Invoke(Sub()
                              EnsureCommandLogVisible()
                              If _commandLogTarget IsNot Nothing Then
                                  _commandLogTarget.Clear()
                              End If
                          End Sub)
    End Sub

    Private Sub AppendCommandLog(message As String)
        Dispatcher.Invoke(Sub()
                              EnsureCommandLogVisible()
                              If _commandLogTarget Is Nothing Then Return
                              _commandLogTarget.AppendText($"[{Date.Now:HH:mm:ss}] {message}{Environment.NewLine}")
                              _commandLogTarget.ScrollToEnd()
                          End Sub)
    End Sub

    Private Sub EnsureCommandLogVisible()
        If Not EnableCommandLogPanel Then Return
        If CommandLogPanel IsNot Nothing AndAlso CommandLogPanel.Visibility <> Visibility.Visible Then
            CommandLogPanel.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub LogProcessOutput(toolName As String, stdOut As String, stdErr As String)
        Dim trimmedOut = If(stdOut, String.Empty).Trim()
        If trimmedOut.Length > 0 Then
            AppendCommandLog($"{toolName} stdout:{Environment.NewLine}{trimmedOut}")
        End If

        Dim trimmedErr = If(stdErr, String.Empty).Trim()
        If trimmedErr.Length > 0 Then
            AppendCommandLog($"{toolName} stderr:{Environment.NewLine}{trimmedErr}")
        End If
    End Sub

    Private Function QuoteArgument(value As String) As String
        Dim safeValue = If(value, String.Empty)
        safeValue = safeValue.Replace(ChrW(34), ChrW(34) & ChrW(34))
        If safeValue.EndsWith("\", StringComparison.Ordinal) Then
            safeValue &= "\"
        End If
        Return """" & safeValue & """"
    End Function

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

    Private Async Function RunProcessAsync(fileName As String, arguments As String) As Task(Of (ExitCode As Integer, StdOut As String, StdErr As String))
        Try
            Dim psi As New ProcessStartInfo(fileName, arguments) With {
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

    Private Async Function ExportDriversAsync(destination As String) As Task(Of Boolean)
        Dim args = String.Join(" ", "/Online", "/Export-Driver", $"/Destination:{QuoteArgument(destination)}")
        AppendCommandLog($"Executing dism.exe {args}")
        Dim result = Await RunDismCommandAsync(args)
        AppendCommandLog($"DISM exit code: {result.ExitCode}")
        LogProcessOutput("dism.exe", result.StdOut, result.StdErr)
        Return result.ExitCode = 0
    End Function

    Private Async Function InjectDriversIntoVhdxAsync(vhdxPath As String, driverFolder As String) As Task(Of Boolean)
        Dim initialLetters As New HashSet(Of Char)(DriveInfo.GetDrives().Select(Function(d) Char.ToUpperInvariant(d.Name(0))))
        Dim mountScript As String = Nothing
        Dim mounted As Boolean = False
        Dim operationSuccess As Boolean = False

        Try
            mountScript = Path.GetTempFileName()
            File.WriteAllLines(mountScript, New String() {
                               $"select vdisk file=""{vhdxPath}""",
                               "attach vdisk",
                               "exit"})

            AppendCommandLog($"Executing diskpart.exe /s ""{mountScript}"" to mount VHDX for driver injection")
            Dim mountResult = Await RunDiskpartScriptAsync(mountScript)
            LogProcessOutput("diskpart.exe", mountResult.StdOut, mountResult.StdErr)
            If mountResult.ExitCode <> 0 Then
                Return False
            End If
            mounted = True

            Dim targetDrive As DriveInfo = Nothing
            For attempt As Integer = 1 To 20
                Await Task.Delay(500)
                Try
                    Dim drives = DriveInfo.GetDrives()
                    targetDrive = drives.FirstOrDefault(Function(d)
                                                            Dim letter = Char.ToUpperInvariant(d.Name(0))
                                                            If initialLetters.Contains(letter) Then Return False
                                                            Dim windowsPath = Path.Combine(d.RootDirectory.FullName, "Windows")
                                                            Return Directory.Exists(windowsPath)
                                                        End Function)
                    If targetDrive IsNot Nothing Then Exit For
                Catch
                End Try
            Next

            If targetDrive Is Nothing Then
                AppendCommandLog("Failed to locate a Windows volume inside the attached VHDX.")
            Else
                Dim imagePath = targetDrive.RootDirectory.FullName
                Dim addArgs = String.Join(" ",
                                          $"/Image:{QuoteArgument(imagePath)}",
                                          "/Add-Driver",
                                          $"/Driver:{QuoteArgument(driverFolder)}",
                                          "/Recurse")
                AppendCommandLog($"Executing dism.exe {addArgs}")
                Dim addResult = Await RunDismCommandAsync(addArgs)
                AppendCommandLog($"DISM exit code: {addResult.ExitCode}")
                LogProcessOutput("dism.exe", addResult.StdOut, addResult.StdErr)
                operationSuccess = addResult.ExitCode = 0
            End If
        Finally
            If Not String.IsNullOrWhiteSpace(mountScript) AndAlso File.Exists(mountScript) Then
                Try
                    File.Delete(mountScript)
                Catch
                End Try
            End If
        End Try

        If mounted Then
            Await DetachVhdxAsync(vhdxPath)
        End If

        Return operationSuccess
    End Function

    ' ==================== EXISTING METHODS ====================

    Private Async Sub AddVhdxButton_Click(sender As Object, e As RoutedEventArgs) Handles AddVHDXButton.Click
        If Not IsAdministrator() Then
            MessageBox.Show("Administrator privileges are required to add a VHDX boot entry.", "Administrator Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim vhdxDialog As New OpenFileDialog With {
            .Filter = "VHDX Files (*.vhdx)|*.vhdx|All Files (*.*)|*.*",
            .Title = "Select VHDX Image",
            .CheckFileExists = True
        }

        If vhdxDialog.ShowDialog() <> True Then
            Return
        End If

        Dim vhdxPath = vhdxDialog.FileName
        If Not File.Exists(vhdxPath) Then
            MessageBox.Show("Selected VHDX file could not be found.", "Invalid File", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End If

        Dim includeDriversPrompt = MessageBox.Show("Include drivers from this device inside the selected VHDX before adding it to the boot menu?", "Include Drivers", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No)
        Dim shouldInjectDrivers = includeDriversPrompt = MessageBoxResult.Yes

        Dim defaultDescription = Path.GetFileNameWithoutExtension(vhdxPath)
        Dim description = InputBox("Enter a description for the new boot entry:", "Boot Entry Description", defaultDescription)
        If String.IsNullOrWhiteSpace(description) Then
            description = defaultDescription
        End If

        Dim driverExportFolder As String = Nothing
        Dim newGuid As String = Nothing
        Dim errorMessage As String = Nothing

        StartDismIndicator()
        Try
            If shouldInjectDrivers Then
                driverExportFolder = Path.Combine(Path.GetTempPath(), "VHDXLabDrivers_" & Guid.NewGuid().ToString("N"))
                Directory.CreateDirectory(driverExportFolder)
                _driverExportPath = driverExportFolder

                UpdateProgress("Drivers...", "Exporting installed drivers")
                Dim exportSuccess = Await ExportDriversAsync(driverExportFolder)
                If Not exportSuccess Then
                    Throw New InvalidOperationException("Failed to export drivers from this device.")
                End If

                UpdateProgress("Drivers...", "Injecting drivers into VHDX")
                Dim injectSuccess = Await InjectDriversIntoVhdxAsync(vhdxPath, driverExportFolder)
                If Not injectSuccess Then
                    Throw New InvalidOperationException("Failed to inject drivers into the selected VHDX.")
                End If
            End If

            UpdateProgress("Creating...", "Adding boot entry")

            newGuid = Await CreateBootEntryAsync(description)
            If String.IsNullOrWhiteSpace(newGuid) Then
                Throw New InvalidOperationException("Failed to create boot entry (no GUID returned).")
            End If

            UpdateProgress("Configuring...", "Linking VHDX image")
            Dim configured = Await ConfigureVhdxBootEntryAsync(newGuid, vhdxPath)
            If Not configured Then
                Throw New InvalidOperationException("Failed to configure the VHDX boot entry.")
            End If

            UpdateProgressSuccess("Added", Path.GetFileName(vhdxPath))
            Await Task.Delay(1500)
            ResetProgress()

            MessageBox.Show($"Boot entry created successfully!{vbCrLf}Description: {description}{vbCrLf}VHDX: {vhdxPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Await LoadBcdEntriesAsync()
        Catch ex As Exception
            errorMessage = ex.Message
        Finally
            StopDismIndicator()

            If Not String.IsNullOrWhiteSpace(driverExportFolder) AndAlso Directory.Exists(driverExportFolder) Then
                Try
                    Directory.Delete(driverExportFolder, True)
                Catch
                End Try
                _driverExportPath = Nothing
            End If
        End Try

        If errorMessage IsNot Nothing Then
            UpdateProgressError("Error", errorMessage)
            Await Task.Delay(2000)
            ResetProgress()

            If Not String.IsNullOrWhiteSpace(newGuid) Then
                Await DeleteBootEntryAsync(newGuid)
            End If

            MessageBox.Show($"Failed to add VHDX boot entry:{vbCrLf}{errorMessage}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End If
    End Sub

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

    Private Async Function SetBingWallpaperAsync() As Task
        If Not Await IsInternetAvailableAsync() Then Return

        Const xmlUrl As String = "https://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1&mkt=en-US"
        Dim xmlContent As String

        Try
            Using resp = Await httpClient.GetAsync(xmlUrl)
                resp.EnsureSuccessStatusCode()
                xmlContent = Await resp.Content.ReadAsStringAsync()
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
            imageBytes = Await httpClient.GetByteArrayAsync(fullImageUrl)
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

        Me.Background = New ImageBrush(bmp) With {
            .Stretch = Stretch.UniformToFill,
            .AlignmentX = AlignmentX.Center,
            .AlignmentY = AlignmentY.Center
        }

        Dim headline = imageElement.Element("headline")?.Value
        Dim copyright = imageElement.Element("copyright")?.Value

        Await Dispatcher.InvokeAsync(Sub()
                                         HeadingTextBlock.Text = If(String.IsNullOrWhiteSpace(headline), "Description", headline)
                                         CopyrightTextBlock.Text = If(String.IsNullOrWhiteSpace(copyright), "Detail", copyright)
                                     End Sub)
    End Function

    Private Shared Async Function IsInternetAvailableAsync() As Task(Of Boolean)
        Try
            Using resp = Await httpClient.GetAsync("https://www.bing.com", HttpCompletionOption.ResponseHeadersRead)
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

