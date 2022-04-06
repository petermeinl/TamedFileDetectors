Imports System.ComponentModel
Imports System.IO
Imports System.Threading

''' <remarks>
''' Files are not sorted. For sorted files use Dataflow ActionBlocks instead of BlockingCollection.
''' Duplicate files are ignored using TryAdd().
''' Keeps working even when the watched directory is renamed.
''' Automatically recovers after a watch directory accessability problem was resolved.
''' Can be configured to monitor WatchPath availability via a slow poll using DirectoryMonitorPollInterval
''' </remarks>
Public Class FileWatcher
    Public WatchPath = ""
    Public Filter = "*.*"
    Public Notifyfilter? As NotifyFilters
    Public InternalBufferSize As Integer
    Public ChangeTypes As WatcherChangeTypes = WatcherChangeTypes.All
    Public IncludeSubdirectories = False

    Public DirectoryMonitorInterval = TimeSpan.FromMinutes(5)
    Public DirectoryRetryInterval = TimeSpan.FromSeconds(5)
    Public DetectedFiles As New Concurrent.BlockingCollection(Of FileChangedInfo)

    'To allow consumter to cancel default error handling
    Event [Error](ByVal sender As Object, ByVal e As FileWatcherErrorEventArgs)

    Private _fileSystenWatcher As FileSystemWatcher
    Private _CT As CancellationToken
    Private _monitorTimer As New System.Threading.Timer(AddressOf _monitorTimer_Elapsed)
    Private _isRecovering = False

    Private Shared _trace As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Sub New()
    End Sub
    Sub New(watchPath As String, filter As String)
        Me.WatchPath = watchPath
        Me.Filter = filter
    End Sub

    Sub Start(CT)
        _trace.Debug("")
        _trace.Info($"Will auto-detect unavailability of watched directory.
           - Windows timeout accessing network shares: ~110 sec on start,
             ~45 sec while watching.")
        Try
            _CT = CT
            _CT.Register(Sub()
                             _monitorTimer.Dispose()
                             _trace.Info("Obeying cancel request")
                             If _fileSystenWatcher IsNot Nothing Then _fileSystenWatcher.Dispose()
                         End Sub)

            ReStartIfNeccessary(TimeSpan.Zero)

        Catch ex As Exception
            _trace.Error($"Unexpected error: {ex}")
        End Try
    End Sub

    Private Sub _monitorTimer_Elapsed(state As Object)
        _trace.Debug("!!")
        _trace.Info($"Watching:{WatchPath}")

        Try
            If Not Directory.Exists(WatchPath) Then
                Throw New DirectoryNotFoundException($"Directory not found {WatchPath}")
            Else
                _trace.Info($"Directory {WatchPath} accessibility is OK.")
                If _fileSystenWatcher Is Nothing Then
                    StartWatching()
                    If _isRecovering Then _trace.Warn("<= Watcher recovered")
                End If

                ReStartIfNeccessary(DirectoryMonitorInterval)
            End If

        Catch ex As DirectoryNotFoundException
            If ExceptionWasHandledByCaller(ex) Then Exit Sub

            If _isRecovering Then
                _trace.Warn("...retrying")
            Else
                _trace.Warn($"=> Directory {WatchPath} Is Not accessible.
                        - Will try to recover automatically in {DirectoryRetryInterval}!")
                _isRecovering = True
            End If

            If _fileSystenWatcher IsNot Nothing Then _fileSystenWatcher.Dispose()
            _fileSystenWatcher = Nothing
            _isRecovering = True
            ReStartIfNeccessary(DirectoryRetryInterval)

        Catch ex As Exception
            _trace.Error($"Unexpected error: {ex}")
        End Try
    End Sub

    Private Sub ReStartIfNeccessary(delay As TimeSpan)
        _trace.Debug("")
        Try
            _monitorTimer.Change(delay, Timeout.InfiniteTimeSpan)
        Catch ex As ObjectDisposedException
            'ignore timer disposed
        End Try
    End Sub

    Private Sub StartWatching()
        _trace.Debug("")

        Try
            _fileSystenWatcher = New FileSystemWatcher()
            With _fileSystenWatcher
                .Path = WatchPath
                .Filter = Filter
                .IncludeSubdirectories = IncludeSubdirectories
                .InternalBufferSize = InternalBufferSize
                'NotifyFilter lacks .None
                If Notifyfilter IsNot Nothing Then .NotifyFilter = Notifyfilter

                AddHandler _fileSystenWatcher.Error, AddressOf _fileWatcher_Error
                If ChangeTypes.HasFlag(WatcherChangeTypes.Changed) Then AddHandler _fileSystenWatcher.Changed,
                    Sub(__, e) BufferDetectedFile(New FileChangedInfo(WatcherChangeTypes.Changed, _fileSystenWatcher.Path, e.Name, ""))
                If ChangeTypes.HasFlag(WatcherChangeTypes.Created) Then AddHandler _fileSystenWatcher.Created,
                    Sub(__, e) BufferDetectedFile(New FileChangedInfo(WatcherChangeTypes.Created, _fileSystenWatcher.Path, e.Name, ""))
                If ChangeTypes.HasFlag(WatcherChangeTypes.Deleted) Then AddHandler _fileSystenWatcher.Deleted,
                    Sub(__, e) BufferDetectedFile(New FileChangedInfo(WatcherChangeTypes.Deleted, _fileSystenWatcher.Path, e.Name, ""))
                If ChangeTypes.HasFlag(WatcherChangeTypes.Renamed) Then AddHandler _fileSystenWatcher.Renamed,
                    Sub(__, e) BufferDetectedFile(New FileChangedInfo(WatcherChangeTypes.Renamed, _fileSystenWatcher.Path, e.Name, e.OldName))

                'NOTE: Don't use a single handler for all change types. This would increase the possibility for InternalBufferOverflowExceptions. 
                'Private Sub _fileWatcher_AllChanges(sender As Object, e As Object) Handles _fileWatcher.Created, _fileWatcher.Changed, _fileWatcher.Deleted, _fileWatcher.Renamed
                '    If TypeOf e Is FileSystemEventArgs Then
                '        BufferFileChange(New FileChangedInfo(e.ChangeType, _fileWatcher.Path, e.Name, ""))
                '    Else
                '        BufferFileChange(New FileChangedInfo(e.ChangeType, _fileWatcher.Path, e.Name, e.OldName))
                '    End If
                'End Sub
            End With

            _fileSystenWatcher.EnableRaisingEvents = True
            BufferExistingFiles()
        Catch ex As OperationCanceledException
            _trace.Info("Obeying cancel request")
        Catch ex As Exception When TypeOf ex Is FileNotFoundException Or TypeOf ex Is DirectoryNotFoundException
            'For race conditions: Path not accessible between .Exists() and .EnableRaisingEvents 
            ReStartIfNeccessary(DirectoryRetryInterval)
        End Try
    End Sub

    Private Sub BufferExistingFiles()
        _trace.Debug("->")
        For Each filePath In Directory.EnumerateFiles(_fileSystenWatcher.Path, _fileSystenWatcher.Filter)
            BufferDetectedFile(New FileChangedInfo(WatcherChangeTypes.All, _fileSystenWatcher.Path, IO.Path.GetFileName(filePath), ""))
            _CT.ThrowIfCancellationRequested()
        Next
        _trace.Debug("<-")
    End Sub

    Private Sub BufferDetectedFile(fileChangedInfo As FileChangedInfo)
        If DetectedFiles.TryAdd(fileChangedInfo) Then
            _trace.Debug(fileChangedInfo)
        Else
            _trace.Debug($"Ingnoring duplicate file {fileChangedInfo.FullPath}")
        End If
    End Sub

    Private Sub _fileWatcher_Error(sender As Object, e As ErrorEventArgs)
        'FSW does set .EnableRaisingEvents=False AFTER raising OnError()
        '  but we don't use this flag for clarity 

        'These exceptions have the same HResult
        Const NetworkNameNoLongerAvailable = -2147467259 'ocurrs on network outage
        Const AccessIsDenied = -2147467259 'occurs after directory was deleted

        _trace.Debug("")
        Dim ex = e.GetException
        If ExceptionWasHandledByCaller(e.GetException) Then Exit Sub

        Select Case True
            Case TypeOf ex Is IO.InternalBufferOverflowException
                _trace.Warn(ex.Message)
                _trace.Error($"Will recover automatically!
                          - This should Not happen with short event handlers. Consider using MaxDefaultInternalBuffersizeKB().")
                ReStartIfNeccessary(DirectoryRetryInterval)
            Case TypeOf ex Is Win32Exception AndAlso
                        (ex.HResult = NetworkNameNoLongerAvailable Or ex.HResult = AccessIsDenied)
                _trace.Warn(ex.Message)
                _trace.Warn("Will try to recover automatically!")
                ReStartIfNeccessary(DirectoryRetryInterval)
            Case Else
                _trace.Error($"Unexpected error: {ex}
                             - Watcher is disabled!")
                Throw ex
        End Select
    End Sub

    Function GetMaxInternalBuffersize() As Integer
        'NOTE: Only increase FSW InternalBuffersize after evaluation other options:
        '  http://msdn.microsoft.com/en-us/library/ded0dc5s.aspx
        '  http://msdn.microsoft.com/en-us/library/aa366778(VS.85).aspx
        Dim maxInternalBufferSize64BitOS = ByteSize.ByteSize.FromKiloBytes(16 * 4)
        Dim maxInternalBufferSize32BitOS = ByteSize.ByteSize.FromKiloBytes(2 * 4)
        If Environment.Is64BitOperatingSystem Then
            Return maxInternalBufferSize64BitOS.Bytes
        Else
            Return maxInternalBufferSize32BitOS.Bytes
        End If
    End Function

    Private Function ExceptionWasHandledByCaller(ByVal ex As Exception)
        'Allow consumer to handle error
        Dim fileWatcherErrorEventArgs As New FileWatcherErrorEventArgs(ex)
        RaiseEvent Error(Me, fileWatcherErrorEventArgs)
        Return fileWatcherErrorEventArgs.Handled
    End Function
End Class




