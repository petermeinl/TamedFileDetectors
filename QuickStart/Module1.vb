Imports System.Threading
Imports Meinl.LeanWork.FileSystem

Module Module1
    Sub Main()
        Dim cts As New CancellationTokenSource
        Dim poller As New FilePoller("d:\temp", "*.*",
                                     TimeSpan.FromSeconds(5))
        poller.StartProcessing(AddressOf ProcessFile, cts.Token)
    End Sub

    Sub ProcessFile(fileInfo As FileChangedInfo, CT As CancellationToken)
        Console.WriteLine(fileInfo)
    End Sub
End Module
