Module main
    Dim AppAddress As String = My.Application.Info.DirectoryPath
    Sub Main()
        Threading.Thread.Sleep(1000) 'ne olur ne olmaz
        For Each prog As Process In Process.GetProcesses 'eğer hata olursa gswin64 işlemini durdur emri versin.
            If prog.ProcessName = "autorun" Then prog.Kill()
        Next
        'son klasörü silip bir üst klasörde işlem yapacağız
        Dim tmp() As String = AppAddress.Split("\")
        AppAddress = ""
        For i = 0 To UBound(tmp) - 1
            AppAddress &= tmp(i) & "\"
        Next
        Shell(AppAddress & "\Autorun\7za.exe x """ & AppAddress & "update.zip"" -y -o""" & AppAddress & """") '-pSECRET
    End Sub
End Module
