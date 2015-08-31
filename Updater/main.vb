Imports System.IO

Module main
    Dim AppAddress As String = My.Application.Info.DirectoryPath
    Sub Main()
        Threading.Thread.Sleep(1000) 'ne olur ne olmaz
        For Each prog As Process In Process.GetProcesses 'eğer hata olursa gswin64 işlemini durdur emri versin.
            If prog.ProcessName = "autorun" Then prog.Kill()
        Next
        If File.Exists(AppAddress & "\update.zip") Then Shell(AppAddress & "\Autorun\7za.exe x """ & AppAddress & "\update.zip"" -y -o""" & AppAddress & """", AppWinStyle.Hide, True)
        File.SetAttributes(AppAddress & "\Autorun", FileAttribute.Hidden)
        File.Delete(AppAddress & "\update.zip")
    End Sub
End Module
