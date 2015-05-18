Imports System.IO
Imports System.Xml

Module main
    Sub Main()
        Dim klasor As New IO.DirectoryInfo(My.Application.Info.DirectoryPath) : Dim Dosya As IO.DirectoryInfo() = klasor.GetDirectories() : Dim FileNames As IO.DirectoryInfo
        Dim html As String = "" : Dim FolderCount As Integer = 0 : Dim Yayinevi, KitapIsmi As String
        For Each FileNames In Dosya
            If File.Exists(FileNames.FullName & "\atu.exe") And File.Exists(FileNames.FullName & "\Book\Data\info.xml") Then
                Dim xmldoc As New XmlDocument : Dim xmlnode As XmlNodeList
                Dim fs As New FileStream(FileNames.FullName & "\Book\Data\info.xml", FileMode.Open, FileAccess.Read)
                xmldoc.Load(fs)
                xmlnode = xmldoc.GetElementsByTagName("info")
                Yayinevi = xmlnode(0).Item("YayineviAdi").InnerText.Replace(" Yayınları", "")
                KitapIsmi = xmlnode(0).Item("KitapIsmi").InnerText
                html &= WriteBook(FileNames.FullName & "\atu.exe", Yayinevi & " " & KitapIsmi, FileNames.Name)
                FolderCount += 1
            End If
        Next
        If FolderCount > 0 Then
            html = "<!DOCTYPE html><html><head><meta charset=""utf-8""><title>AutoRun</title>" &
                    "<HTA:APPLICATION ID=""oMyApp"" APPLICATIONNAME = ""Application Executer"" BORDER = ""no"" CAPTION = ""yes"" ShowInTaskbar = ""yes"" SINGLEINSTANCE = ""yes"" SYSMENU = ""yes"" Scroll = ""no"" WINDOWSTATE=""normal"" MAXIMIZEBUTTON=""no"" SELECTION=""no"" ICON=""Autorun.ico"" />" &
                    "<style>body{background:url(Background.jpg) no-repeat center center fixed;background-size:cover;min-height:500px;padding:200px 0 0 0;margin:0;}h3{max-width:150px;margin:5px;}" &
                    "#panel{width:" & 300 * FolderCount & "px;margin:0 auto;}.cerceve{width:280px;height:500px;padding:0;float:left;margin:0 10px;text-align:center;background:url(Background.png) no-repeat;background-size:100%;}img{border:0;}</style>" &
                    "</head><body><div id=""panel"">" & html & "</div>"
            html &= "<script type=""text/javascript"" language=""javascript"">window.resizeTo(1024,768);window.onresize=function(){window.resizeTo(1024,768);};function RunFile(src) {WshShell = new ActiveXObject(""WScript.Shell"");WshShell.Run(src, 1, false);window.close();}</script>"
            html &= "</body></html>"
            File.Delete(My.Application.Info.DirectoryPath & "\index.hta")
            Dim objWriter As New System.IO.StreamWriter(My.Application.Info.DirectoryPath & "\index.hta")
            objWriter.Write(html)
            objWriter.Close()
            Shell("explorer.exe " & My.Application.Info.DirectoryPath & "\index.hta")
        End If
    End Sub
    Private Function WriteBook(ExeFile As String, title As String, foldername As String) As String
        Return "<div class=""cerceve""><a href=""#"" onclick=""RunFile('file:///" & ExeFile.Replace("\", "/") & "');""><img src=""" & foldername & "/Book/Images/Pages/Thumbs/Sayfa1.etf"" height=""280"" /></a><h3>" & title & "<h3></div>"
    End Function
End Module
