Imports System.IO
Imports System.Xml

Module main
    Sub Main()
        Dim tmp() As String = My.Application.Info.DirectoryPath.Split("\") : Dim klasoryolu As String = "" : For i = 0 To UBound(tmp) - 1 : klasoryolu &= tmp(i) & "\" : Next
        Dim klasor As New IO.DirectoryInfo(klasoryolu) : Dim Dosya As IO.DirectoryInfo() = klasor.GetDirectories() : Dim FileNames As IO.DirectoryInfo
        Dim html, Yayinevi, KitapIsmi, Web, SonSayfa, Sinif As String : Dim FolderCount, j As Integer : Dim isLibrary As Boolean = False
        html = "" : Yayinevi = "" : Web = "" : SonSayfa = "" : Sinif = "" : FolderCount = 0 : j = 0
        For Each FileNames In Dosya
            If File.Exists(FileNames.FullName & "\atu.exe") And File.Exists(FileNames.FullName & "\Book\Data\info.xml") Then FolderCount += 1
            If FolderCount > 5 Then isLibrary = True : Exit For
        Next
        For Each FileNames In Dosya
            If File.Exists(FileNames.FullName & "\atu.exe") And File.Exists(FileNames.FullName & "\Book\Data\info.xml") Then
                Dim xmldoc As New XmlDocument : Dim xmlnode As XmlNodeList
                Dim fs As New FileStream(FileNames.FullName & "\Book\Data\info.xml", FileMode.Open, FileAccess.Read)
                Try
                    xmldoc.Load(fs) : xmlnode = xmldoc.GetElementsByTagName("info")
                    Yayinevi = xmlnode(0).Item("YayineviAdi").InnerText
                    KitapIsmi = xmlnode(0).Item("KitapIsmi").InnerText
                    Web = xmlnode(0).Item("InternetSitesi").InnerText
                    SonSayfa = xmlnode(0).Item("SonSayfa").InnerText
                    SonSayfa = xmlnode(0).Item("SonSayfa").InnerText
                    j += 1
                    If isLibrary = True Then
                        If html <> "" Then html &= ","
                        html &= WriteBook2(j, KitapIsmi, FileNames.FullName & "\atu.exe", SonSayfa, Sinif)
                    Else
                        html &= WriteBook(FileNames.FullName & "\atu.exe", Yayinevi.Replace(" Yayınları", "") & " " & KitapIsmi, FileNames.Name)
                    End If
                Catch ex As Exception
                End Try
            End If
        Next
        If FolderCount > 0 Then
            Dim genislik As Integer = 200 * FolderCount : If genislik > 1000 Then genislik = 1000
            html = WriteHta(isLibrary, True, My.Application.Info.DirectoryPath & "\Autorun.exe", genislik, Yayinevi, Web) & html & WriteHta(isLibrary, False, My.Application.Info.DirectoryPath & "\Autorun.exe")
            File.Delete(klasoryolu & "Autorun.hta")
            Dim objWriter As New System.IO.StreamWriter(klasoryolu & "Autorun.hta")
            objWriter.Write(html)
            objWriter.Close()
            Shell("explorer.exe " & klasoryolu & "Autorun.hta")
        End If
    End Sub
    Private Function WriteBook(ExeFile As String, title As String, foldername As String) As String
        Return "<div class=""cerceve""><a href=""#"" onclick=""RunFile('file:///" & ExeFile.Replace("\", "/") & "');""><img src=""" & foldername & "/Book/Images/Pages/Thumbs/Sayfa1.etf"" height=""180"" /></a><h3>" & title & "<h3></div>"
    End Function
    Private Function WriteBook2(id As Integer, title As String, ExeFile As String, SonSayfa As String, Sinif As String) As String
        Return "{""id"":""" & id & """,""title"":""" & title & """,""description"":""" & title & """,""url"":""file:///" & ExeFile.Replace("\", "/") & """,""pages"":""" & SonSayfa & """,""newTime"":"""",""categoryname"":"""",""categoryid"":"""",""brans"":"""",""bransid"":"""",""sinif"":"""",""sinifid"":""" & Sinif & """,""label"":""0""}"
    End Function
    Private Function WriteHta(isLibrary As Boolean, isBeginning As Boolean, ExeFile As String, Optional genislik As Integer = 1000, Optional yayinevi As String = "", Optional web As String = "") As String
        Dim html As String = ""
        If isLibrary = True Then
            Dim tmpDokuman As String = "[{""value"": ""5"",""label"": ""Çalışma Kitabı""},{""value"": ""10"",""label"": ""Çıkmış Sorular""},{""value"": ""4"",""label"": ""Çözümlü Soru Bankası""},{""value"": ""12"",""label"": ""Defter""},{""value"": ""8"",""label"": ""Deneme""},{""value"": ""6"",""label"": ""Dergi""},{""value"": ""7"",""label"": ""Fasikül""},{""value"": ""1"",""label"": ""Konu Anlatımı""},{""value"": ""2"",""label"": ""Konu Anlatımlı Soru Bankası""},{""value"": ""9"",""label"": ""Konu Testi""},{""value"": ""3"",""label"": ""Soru Bankası""},{""value"": ""11"",""label"": ""Yaprak test""}]"
            Dim tmpSinif As String = "[{""value"": ""1"",""label"": "" 1. Sınıf""},{""value"": ""2"",""label"": "" 2. Sınıf""},{""value"": ""3"",""label"": "" 3. Sınıf""},{""value"": ""4"",""label"": "" 4. Sınıf""},{""value"": ""5"",""label"": "" 5. Sınıf""},{""value"": ""6"",""label"": "" 6. Sınıf""},{""value"": ""7"",""label"": "" 7. Sınıf""},{""value"": ""8"",""label"": "" 8. Sınıf""},{""value"": ""9"",""label"": "" 9. Sınıf""},{""value"": ""10"",""label"": ""10. Sınıf""},{""value"": ""11"",""label"": ""11. Sınıf""},{""value"": ""12"",""label"": ""12. Sınıf""},{""value"": ""19"",""label"": ""Ales""},{""value"": ""16"",""label"": ""Genel""},{""value"": ""18"",""label"": ""Hazırlık""},{""value"": ""20"",""label"": ""KPDS-UDS""},{""value"": ""15"",""label"": ""KPSS""},{""value"": ""14"",""label"": ""LYS""},{""value"": ""21"",""label"": ""Okul Öncesi""},{""value"": ""13"",""label"": ""YGS""},{""value"": ""17"",""label"": ""YGS-LYS""}]"
            Dim tmpBrans As String = "[{""value"": ""3"",""label"": ""Biyoloji""},{""value"": ""7"",""label"": ""Coğrafya""},{""value"": ""12"",""label"": ""Din Kültürü ve Ahlak Bilgisi""},{""value"": ""10"",""label"": ""Eğitim Bilimleri""},{""value"": ""8"",""label"": ""Felsefe""},{""value"": ""1"",""label"": ""Fizik""},{""value"": ""14"",""label"": ""Genel""},{""value"": ""5"",""label"": ""Geometri""},{""value"": ""13"",""label"": ""İlkögretim 1-4""},{""value"": ""11"",""label"": ""İngilizce""},{""value"": ""2"",""label"": ""Kimya""},{""value"": ""4"",""label"": ""Matematik""},{""value"": ""6"",""label"": ""Tarih""},{""value"": ""9"",""label"": ""Türkçe""},{""value"": ""15"",""label"": ""Yok""}]"
            If isBeginning = True Then
                html = "<!DOCTYPE html><html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge;"" /><meta charset=""utf-8""><meta http-equiv=""content-type"" content=""text/html; charset=utf-8"" /><meta content=""tr_TR"" http-equiv=""Content-Language""><title>AutoRun</title>"
                html &= "<HTA:APPLICATION ID=""oMyApp"" APPLICATIONNAME = ""Application Executer"" BORDER = ""no"" CAPTION = ""yes"" ShowInTaskbar = ""yes"" SINGLEINSTANCE = ""yes"" SYSMENU = ""yes"" Scroll = ""no"" WINDOWSTATE=""normal"" MAXIMIZEBUTTON=""no"" SELECTION=""no"" ICON=""Autorun/Autorun.ico"" />"
                html &= "<LINK href=""Autorun/Autorun.ico"" rel=""SHORTCUT ICON""/><link href=""Autorun/Includes/Bookcase.min.css"" rel=""stylesheet""/></head><body>"
                html &= "<div class=""main-container""></div>"
                html &= "<script src=""Autorun/Includes/jquery.min.js""></script>"
                html &= "<script src=""Autorun/Includes/md5.js""></script>"
                html &= "<script src=""Autorun/Includes/hammer.min.js""></script>"
                html &= "<script src=""Autorun/Includes/jquery.hammer.min.js""></script>"
                html &= "<script src=""Autorun/Includes/jquery.qrcode.min.js""></script>"
                html &= "<script src=""Autorun/Includes/bookcase.min.js""></script>"
                html &= "<script type=""text/javascript"" language=""javascript"">window.resizeTo(1024,768);window.moveTo(0,0);window.onresize=function(){window.resizeTo(1024,768);};function RunFile(src){WshShell = new ActiveXObject(""WScript.Shell"");WshShell.Run(src, 1, false);window.close();};"

                html &= "(function ($) {$(function () {"
                html &= "var userData = {title: """ & yayinevi & """,name: """ & yayinevi & """, email: ""info@surat.gen.tr"",website: """ & web & """,logoLink: """ & web & """,logoAddress: ""Autorun/Logo.png"",accountLogo: ""Autorun/Logo.png"",skin: """","
                html &= "isShowShare: ""1"",isShowContact: ""1"",isShowSearch: ""1"",isShowSkin: ""0"",isShowCategory: ""1"",isShowLogo: ""1"",openType: ""0"",isSelf: ""0"",about: ""İlkokuldan Üniversiteye Kadar"",updatelink:""file:///" & ExeFile.Replace("\", "/") & """};"
                html &= "new Bookcase($("".main-container""),userData,["
            Else
                html = "]," & tmpDokuman & "," & tmpSinif & "," & tmpBrans & ");"
                html &= "});})(window.jQuery);"
                html &= "</script></body></html>"
            End If

        Else
            If isBeginning = True Then
                html = "<!DOCTYPE html><html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge;"" /><meta charset=""utf-8""><meta http-equiv=""content-type"" content=""text/html; charset=utf-8"" /><meta content=""tr_TR"" http-equiv=""Content-Language""><title>AutoRun</title>"
                html &= "<HTA:APPLICATION ID=""oMyApp"" APPLICATIONNAME = ""Application Executer"" BORDER = ""no"" CAPTION = ""yes"" ShowInTaskbar = ""yes"" SINGLEINSTANCE = ""yes"" SYSMENU = ""yes"" Scroll = ""no"" WINDOWSTATE=""normal"" MAXIMIZEBUTTON=""no"" SELECTION=""no"" ICON=""Autorun/Autorun.ico"" />"
                html &= "<style>body{background:url(Autorun/Background.jpg) no-repeat center center fixed;background-size:cover;min-height:500px;padding:200px 0 0 0;margin:0;}h3{max-width:170px;margin:5px;}"
                html &= "#panel{width:" & genislik & "px;margin:0 auto;}.cerceve{width:180px;height:300px;padding:0;float:left;margin:0 10px;text-align:center;background:url(Autorun/Background.png) no-repeat;background-size:100%;}img{border:0;}"
                html &= "#update{position:absolute;right:0;top:0;}a{color:#000;font-weight:bold;text-decoration:none;}</style>"
                html &= "<LINK href=""Autorun/Autorun.ico"" rel=""SHORTCUT ICON""/></head><body><div id=""update""><a href=""#"" onclick=""RunFile('file:///" & ExeFile.Replace("\", "/") & "');window.close;"">Güncelle</a></div><div id=""panel"">"
            Else
                html = "</div><script type=""text/javascript"" language=""javascript"">window.resizeTo(1024,768);window.moveTo(0,0);window.onresize=function(){window.resizeTo(1024,768);};function RunFile(src){WshShell = new ActiveXObject(""WScript.Shell"");WshShell.Run(src, 1, false);window.close();};</script>"
                html &= "</body></html>"
            End If
        End If
        Return html
    End Function
End Module
