Imports System.IO
Imports System.Threading

''' <remarks>
''' Files are not sorted sorted.
''' Automatically resumes after a watch directory accessability problem was resolved.
''' </remarks>
Public Class FilePoller
    Public WatchPath = ""
    Public Filter = "*.*"
    Public PollInterval = TimeSpan.FromSeconds(5)
    Public IncludeSubdirectories = False
    Public OrderByOldestFirst As Boolean = False
    'To allow consumter to cancel on error
    Event [Error](ByVal sender As Object, ByVal e As FileWatcherErrorEventArgs)

    Private _wasWatchPathAccessible As Boolean = True
    Private _trace As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Sub New()
    End Sub
    Sub New(watchPath As String, filter As String, pollInterval As TimeSpan)
        Me.WatchPath = watchPath
        Me.Filter = filter
        Me.PollInterval = pollInterval
    End Sub

    Sub StartProcessing(processFile As Action(Of FileChangedInfo, CancellationToken), CT As CancellationToken)
        'These exceptions have the same HResult
        Const TheNetworkPathWasNotFound = -2147024843 'occurs on network outage
        Const CoulNotFindaPartofPath = -2147024843 'occurs after directory was deleted

        _trace.Debug("")
        _trace.Info($"Watching:{WatchPath}")

        _trace.Info($"Will auto-detect unavailability of watched directory.
           - Windows timeout accessing network shares: ~110 sec on start, ~45 sec while watching.")
        Do
            Try
                CT.ThrowIfCancellationRequested()

                If Directory.Exists(WatchPath) And Not _wasWatchPathAccessible Then
                    _trace.Warn($"<= WatchPath {WatchPath} is accessible again")
                    _wasWatchPathAccessible = True
                End If

                Dim searchSubDirectoriesOption = If(IncludeSubdirectories, SearchOption.AllDirectories, SearchOption.TopDirectoryOnly)
                If OrderByOldestFirst Then
                    'Gets all fileInfos first, sorts oldest first and then starts working
                    For Each fileInfo In From fi In New DirectoryInfo(WatchPath).GetFiles(Filter, searchSubDirectoriesOption)
                                         Order By fi.LastWriteTime Ascending
                                         Select fi
                        CT.ThrowIfCancellationRequested()
                        processFile(New FileChangedInfo(WatcherChangeTypes.All, WatchPath, fileInfo.Name, ""), CT)
                    Next
                Else
                    'Works on unsorted files as soon as the first file is detected
                    '- Good option if there are very many files
                    For Each fileInfo In New DirectoryInfo(WatchPath).EnumerateFiles(Filter, searchSubDirectoriesOption)
                        CT.ThrowIfCancellationRequested()
                        processFile(New FileChangedInfo(WatcherChangeTypes.All, WatchPath, fileInfo.Name, ""), CT)
                    Next
                End If

            Catch ex As OperationCanceledException
                _trace.Debug("Obeying cancel request")
                Exit Do
            Catch ex As Exception When TypeOf ex Is DirectoryNotFoundException _
                    Or TypeOf ex Is IOException And ex.HResult = TheNetworkPathWasNotFound _
                    Or TypeOf ex Is IOException And ex.HResult = CoulNotFindaPartofPath
                If ExceptionWasHandledByCaller(ex) Then Exit Sub

                If _wasWatchPathAccessible Then
                    _trace.Warn($"=> WatchPath {WatchPath} is not accessible.
                            - Will try to recover automatically in {PollInterval}!")
                Else
                    _trace.Warn("...retrying")
                End If
            Catch ex As Exception
                _trace.Error($"Unexpected error: {ex}")
                Throw
            End Try

            CT.WaitHandle.WaitOne(PollInterval)
        Loop
    End Sub

    Private Function ExceptionWasHandledByCaller(ByVal ex As Exception)
        'Allow consumer to handle error
        Dim fileWatcherErrorEventArgs As New FileWatcherErrorEventArgs(ex)
        RaiseEvent Error(Me, fileWatcherErrorEventArgs)
        Return fileWatcherErrorEventArgs.Handled
    End Function
End Class
