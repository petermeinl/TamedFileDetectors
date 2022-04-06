Imports System.ComponentModel

Public Class FileWatcherErrorEventArgs : Inherits HandledEventArgs
    Public ReadOnly [Error] As Exception

    Public Sub New(ByVal exception As Exception)
        Me.Error = exception
    End Sub
End Class
