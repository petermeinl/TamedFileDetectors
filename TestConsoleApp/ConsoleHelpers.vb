Module ConsoleHelpers
    Public Sub PromptUser(message As String, Optional foregroundColor As ConsoleColor = ConsoleColor.White)
        WriteLineInColor($"? {message} |Press <Enter> to contine.", foregroundColor)
        Do Until (Console.ReadKey(True).Key = ConsoleKey.Enter) : Loop
    End Sub

    Public Sub WriteLineInColor(message As String, foregroundColor As ConsoleColor)
        Console.ForegroundColor = foregroundColor
        Console.WriteLine(message)
        Console.ResetColor()
    End Sub
End Module
