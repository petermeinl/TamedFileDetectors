Imports System.Collections.Concurrent
Imports System.IO
Imports System.Threading
Imports Meinl.LeanWork.FileSystem

Module _Main
    Const _inDir = "\\gogo\test\in"
    'Const _inDir = "d:\temp\in"
    Const _filter = "*.*"
    'Const _filter = "*.xml"
    Const _outDir = "d:\temp\OUT"
    Private _cts As New CancellationTokenSource

    Private _trace As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Sub Main(cmdArgs() As String)
        log4net.Config.XmlConfigurator.ConfigureAndWatch(New FileInfo(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile))
        Try
            Dim workerTask As Task = Nothing
            Dim cmd = String.Concat(cmdArgs).ToLower()

            'If String.IsNullOrWhiteSpace(cmd) Then cmd = "-poller"
            If String.IsNullOrWhiteSpace(cmd) Then cmd = "-watcher"

            Select Case cmd
                Case "-poller", "-p"
                    workerTask = RunPoller()
                Case "-watcher", "-w"
                    workerTask = RunWatcher()
                Case Else
                    WriteLineInColor($"Usage: {My.Application.Info.AssemblyName} [-p | -w]", ConsoleColor.Red)
            End Select
            Console.WindowWidth = 130
            PromptUser("Processing...")

            _cts.Cancel()
            Console.WriteLine("Stopping...")
            workerTask.Wait()
            PromptUser("Stopped.")

        Catch ex As Exception
            WriteLineInColor($"Unexpected error: {ex.ToString}", ConsoleColor.Red)
            PromptUser("")
        End Try
    End Sub

    Function RunPoller() As Task
        Console.Title = "-- File Poller --"
        Dim CT = _cts.Token
        Dim poller As New FilePoller(_inDir, _filter, TimeSpan.FromSeconds(3)) With {
                .OrderByOldestFirst = False,
                .IncludeSubdirectories = True
                }
        'AddHandler poller.Error, Sub(s, e) HandleError(s, e) 'Override auto robustness
        Return Task.Run(Sub() poller.StartProcessing(AddressOf ProcessFile, CT),
                                  CT)
    End Function

    Function RunWatcher() As Task
        Console.Title = "-- File Watcher --"
        Dim CT = _cts.Token

        Dim watcher As New FileWatcher(_inDir, _filter)
        Dim workerTask = Task.Run(Sub() ProcessFilesWait(watcher.DetectedFiles, CT),
                                  CT)
        'watcher.ChangeTypes = WatcherChangeTypes.All
        'watcher.ChangeTypes = WatcherChangeTypes.Renamed
        watcher.ChangeTypes = WatcherChangeTypes.Created
        watcher.DirectoryMonitorInterval = TimeSpan.FromMinutes(1)
        watcher.InternalBufferSize = watcher.GetMaxInternalBuffersize
        watcher.Start(CT)
        'AddHandler watcher.Error, Sub(s, e) HandleError(s, e) 'Override auto robustness
        Return workerTask
    End Function

#Region "Processors"
    'Only used by FilteWatcher not by FilePoller
    Sub ProcessFilesWait(detectedFiles As BlockingCollection(Of FileChangedInfo),
                         CT As CancellationToken)
        _trace.Debug("")
        Try
            'Waits for new items if BlockingCollection is empty until cancelled
            For Each fileChangedInfo In detectedFiles.GetConsumingEnumerable(CT)
                ProcessFile(fileChangedInfo, CT)
            Next
        Catch ex As OperationCanceledException
            _trace.Debug("Obeying cancel request")
        End Try
    End Sub

    Sub ProcessFile(fileChangedInfo As FileChangedInfo, CT As CancellationToken)
        CT.ThrowIfCancellationRequested()
        Dim filePath = fileChangedInfo.FullPath

        Try
            _trace.Info(fileChangedInfo)
            '_trace.Warn(fileChangedInfo)

            Dim toFilePath = Path.Combine(_outDir, Path.GetFileName(filePath))
            My.Computer.FileSystem.MoveFile(filePath, toFilePath, overwrite:=True)
            'Microsoft.VisualBasic.FileSystem.Rename()
            'Microsoft.VisualBasic.FileIO.FileSystem.RenameFile()
            'My.Computer.FileSystem.RenameFile()

        Catch ex As IOException When IsFileInUse(ex)
            _trace.Warn($"Skipping file in use {filePath}")
        Catch ex As UnauthorizedAccessException
            _trace.Warn($"Skipping file not authorized for {filePath}")
        Catch ex As FileNotFoundException
            _trace.Warn($"Skipping file no longer existing {filePath}")
        End Try
    End Sub

    Function IsFileInUse(ByVal ex As IOException) As Boolean
        'https://msdn.microsoft.com/en-us/library/ms681382%28v=VS.85%29.aspx
        Const ERROR_SHARING_VIOLATION = &H80070020, ERROR_LOCK_VIOLATION = &H21
        Return (ex.HResult = ERROR_SHARING_VIOLATION OrElse ex.HResult = ERROR_LOCK_VIOLATION)
    End Function

    Private Sub HandleError(sender As Object, e As FileWatcherErrorEventArgs)
        e.Handled = True
        WriteLineInColor($"Unexpected error: {e.Error.ToString}", ConsoleColor.Red)
        'TODO: Implement custom recovery action or abort (fail fast)
        PromptUser("Will fail fast without trying to cancel.")
    End Sub
#End Region

End Module
