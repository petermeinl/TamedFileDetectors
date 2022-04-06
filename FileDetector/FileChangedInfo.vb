Imports System.IO

Public Class FileChangedInfo : Inherits RenamedEventArgs
    Public DetectedTime = DateTime.Now

    Sub New(changeType As WatcherChangeTypes, directory As String, name As String, oldName As String)
        MyBase.New(changeType, directory, name, oldName)
    End Sub

    Public Overrides Function ToString() As String
        Dim changeType = If(Me.ChangeType = WatcherChangeTypes.All, "Polled", Me.ChangeType)
        Return $"{DetectedTime} -{changeType}- {FullPath}"
    End Function
End Class