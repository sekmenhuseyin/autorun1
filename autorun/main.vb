Imports System.IO
Imports System.Xml

Module main
    Sub Main()
        'tanımlamalar
        Dim tmp() As String = My.Application.Info.DirectoryPath.Split("\") : Dim klasoryolu As String = "" : For i = 0 To UBound(tmp) - 1 : klasoryolu &= tmp(i) & "\" : Next
        Dim klasor As New IO.DirectoryInfo(My.Application.Info.DirectoryPath) : Dim Dosya As IO.DirectoryInfo() = klasor.GetDirectories() : Dim FileNames As IO.DirectoryInfo
        Dim title, web, html As String : Dim FolderCount As Integer : Dim isLibrary As Boolean = False
        html = "" : title = "" : web = "" : FolderCount = 0
        'kitap sayısını bul
        For Each FileNames In Dosya
            If File.Exists(FileNames.FullName & "\atu.exe") And File.Exists(FileNames.FullName & "\Book\Data\info.xml") Then
                FolderCount += 1
                'eğer xmlde ilk satır indo değilse hata veriyordu
                'bu yüzden infoda öncesini sildiriyorum
                Dim objWriter As New StreamReader(FileNames.FullName & "\Book\Data\info.xml")
                Dim alltext = objWriter.ReadToEnd
                Dim pos = InStr(alltext, "<info>")
                If pos <> 1 Then
                    objWriter.Close() 'önce açık olanı kapat
                    alltext = Mid(alltext, pos) 'info yazısından itibaren yazdır
                    Dim objWriter2 As New StreamWriter(FileNames.FullName & "\Book\Data\info.xml")
                    objWriter2.Write(alltext)
                    objWriter2.Close()
                Else
                    objWriter.Close()
                End If
            End If
            If FolderCount > 5 Then isLibrary = True : FolderCount = 0 : Exit For
        Next
        For Each FileNames In Dosya
            'değerleri sıfırlar
            Dim Yayinevi, KitapIsmi, InternetSitesi, SonSayfa, Sinif(1), Tur(1), Brans(1) As String
            Yayinevi = "" : KitapIsmi = "" : InternetSitesi = "" : SonSayfa = "" : Sinif(0) = "" : Tur(0) = "" : Brans(0) = "" : Sinif(1) = "" : Tur(1) = "" : Brans(1) = ""
            'xml varsa oku
            If File.Exists(FileNames.FullName & "\atu.exe") And File.Exists(FileNames.FullName & "\Book\Data\info.xml") Then
                'xmlden okuma bmlünü
                Dim xmldoc As New XmlDocument : Dim xmlnode As XmlNodeList
                Dim fs As New FileStream(FileNames.FullName & "\Book\Data\info.xml", FileMode.Open, FileAccess.Read)
                'xml yükle
                xmldoc.Load(fs) : xmlnode = xmldoc.GetElementsByTagName("info")
                'bilgileri al
                If IsNothing(xmlnode(0).Item("YayineviAdi")) = False Then Yayinevi = xmlnode(0).Item("YayineviAdi").InnerText
                If IsNothing(xmlnode(0).Item("KitapIsmi")) = False Then KitapIsmi = xmlnode(0).Item("KitapIsmi").InnerText
                If IsNothing(xmlnode(0).Item("InternetSitesi")) = False Then InternetSitesi = xmlnode(0).Item("InternetSitesi").InnerText
                If IsNothing(xmlnode(0).Item("SonSayfa")) = False Then SonSayfa = xmlnode(0).Item("SonSayfa").InnerText
                If IsNothing(xmlnode(0).Item("Sinif")) = False Then Sinif(0) = xmlnode(0).Item("Sinif").InnerText
                If IsNothing(xmlnode(0).Item("SinifAd")) = False Then Sinif(1) = xmlnode(0).Item("SinifAd").InnerText
                If IsNothing(xmlnode(0).Item("Tur")) = False Then Tur(0) = xmlnode(0).Item("Tur").InnerText
                If IsNothing(xmlnode(0).Item("TurAd")) = False Then Tur(1) = xmlnode(0).Item("TurAd").InnerText
                If IsNothing(xmlnode(0).Item("Brans")) = False Then Brans(0) = xmlnode(0).Item("Brans").InnerText
                If IsNothing(xmlnode(0).Item("BransAd")) = False Then Brans(1) = xmlnode(0).Item("BransAd").InnerText
                'genel bilgiler
                If Yayinevi <> "" And title = "" Then title = Yayinevi
                If InternetSitesi <> "" And web = "" Then web = InternetSitesi
                'her kitap için bir satır yaz
                If isLibrary = True Then
                    If html <> "" Then html &= ","
                    html &= WriteBook2(FolderCount, KitapIsmi, FileNames.Name, FileNames.FullName & "\atu.exe", SonSayfa, Tur(1), Tur(0), Brans(1), Brans(0), Sinif(1), Sinif(0))
                Else
                    html &= WriteBook(FileNames.FullName & "\atu.exe", Yayinevi.Replace(" Yayınları", "") & " " & KitapIsmi, FileNames.Name)
                End If
                Try
                Catch ex As Exception
                End Try
                'kapat
                fs.Close() : fs.Dispose() : xmlnode = Nothing : xmldoc = Nothing : FolderCount += 1
            End If
        Next
        'kitap varsa
        If FolderCount > 0 Then
            'genişliği bul
            Dim genislik As Integer = 200 * FolderCount : If genislik > 1000 Then genislik = 1000
            'sayfayı yazdır
            html = WriteHta(isLibrary, True, klasoryolu & "Autorun.exe", genislik, title, web) & html & WriteHta(isLibrary, False, klasoryolu & "Autorun.exe")
            'hta yazdır
            File.Delete(My.Application.Info.DirectoryPath & "\Autorun\Autorun.hta")
            Dim objWriter As New System.IO.StreamWriter(My.Application.Info.DirectoryPath & "\Autorun\Autorun.hta")
            objWriter.Write(html)
            objWriter.Close()
            've başlat
            Shell("explorer.exe " & My.Application.Info.DirectoryPath & "\Autorun\Autorun.hta")
        End If
    End Sub
    Private Function WriteBook(ExeFile As String, title As String, foldername As String) As String
        Dim klasor As String = ExeFile.Replace("atu.exe", "Book\Images\Pages\Thumbs\Cover.etf")
        Dim kapak As String = IIf(File.Exists(klasor) = True, "../" & foldername & "/Book/Images/Pages/Thumbs/cover.etf", "../" & foldername & "/Book/Images/Pages/Thumbs/Sayfa1.etf")
        Return "<div class=""cerceve""><a href=""#"" onclick=""RunFile('file:///" & ExeFile.Replace("\", "/") & "');""><img src=""" & kapak & """ height=""180"" /></a><h3>" & title & "</h3></div>"
    End Function
    Private Function WriteBook2(id As Integer, title As String, foldername As String, ExeFile As String, SonSayfa As String, category As String, categoryid As String, brans As String, bransid As String, sinif As String, sinifid As String) As String
        Dim aciklama = "Çözümlü Sorular ve cevap anahtarları eksiklerinizi göstermez Konu sütunlarına yerleştirilen anahtar bilgiler ile başka bir kaynağa ihtiyaç duymadan cevabı kendiniz bularak konuyu pekiştirebilirsiniz."
        Dim klasor As String = ExeFile.Replace("atu.exe", "Book\Images\Pages\Thumbs\Cover.etf")
        Dim kapak As String = IIf(File.Exists(klasor) = True, "../" & foldername & "/Book/Images/Pages/Thumbs/cover.etf", "../" & foldername & "/Book/Images/Pages/Thumbs/Sayfa1.etf")
        Return "{""id"":""" & id & """,""title"":""" & title & """,""description"":""" & aciklama & """,""url"":""file:///" & ExeFile.Replace("\", "/") & """,""pages"":""" & SonSayfa & """,""kapak"":""" & kapak & """,""categoryname"":""" & category & """,""categoryid"":""" & categoryid & """,""brans"":""" & brans & """,""bransid"":""" & bransid & """,""sinif"":""" & sinif & """,""sinifid"":""" & sinifid & """,""label"":""0""}"
    End Function
    Private Function WriteHta(isLibrary As Boolean, isBeginning As Boolean, ExeFile As String, Optional genislik As Integer = 1000, Optional yayinevi As String = "", Optional web As String = "") As String
        Dim html As String = ""
        If isLibrary = True Then
            Dim tmpDokuman As String = "[{""value"": ""5"",""label"": ""Çalışma Kitabı""},{""value"": ""10"",""label"": ""Çıkmış Sorular""},{""value"": ""4"",""label"": ""Çözümlü Soru Bankası""},{""value"": ""12"",""label"": ""Defter""},{""value"": ""8"",""label"": ""Deneme""},{""value"": ""6"",""label"": ""Dergi""},{""value"": ""7"",""label"": ""Fasikül""},{""value"": ""1"",""label"": ""Konu Anlatımı""},{""value"": ""2"",""label"": ""Konu Anlatımlı Soru Bankası""},{""value"": ""9"",""label"": ""Konu Testi""},{""value"": ""3"",""label"": ""Soru Bankası""},{""value"": ""11"",""label"": ""Yaprak test""}]"
            Dim tmpSinif As String = "[{""value"": ""1"",""label"": "" 1. Sınıf""},{""value"": ""2"",""label"": "" 2. Sınıf""},{""value"": ""3"",""label"": "" 3. Sınıf""},{""value"": ""4"",""label"": "" 4. Sınıf""},{""value"": ""5"",""label"": "" 5. Sınıf""},{""value"": ""6"",""label"": "" 6. Sınıf""},{""value"": ""7"",""label"": "" 7. Sınıf""},{""value"": ""8"",""label"": "" 8. Sınıf""},{""value"": ""9"",""label"": "" 9. Sınıf""},{""value"": ""10"",""label"": ""10. Sınıf""},{""value"": ""11"",""label"": ""11. Sınıf""},{""value"": ""12"",""label"": ""12. Sınıf""},{""value"": ""19"",""label"": ""Ales""},{""value"": ""16"",""label"": ""Genel""},{""value"": ""18"",""label"": ""Hazırlık""},{""value"": ""20"",""label"": ""KPDS-UDS""},{""value"": ""15"",""label"": ""KPSS""},{""value"": ""14"",""label"": ""LYS""},{""value"": ""21"",""label"": ""Okul Öncesi""},{""value"": ""13"",""label"": ""YGS""},{""value"": ""17"",""label"": ""YGS-LYS""}]"
            Dim tmpBrans As String = "[{""value"": ""3"",""label"": ""Biyoloji""},{""value"": ""7"",""label"": ""Coğrafya""},{""value"": ""12"",""label"": ""Din Kültürü ve Ahlak Bilgisi""},{""value"": ""10"",""label"": ""Eğitim Bilimleri""},{""value"": ""8"",""label"": ""Felsefe""},{""value"": ""14"",""label"": ""Fen Bilimleri""},{""value"": ""15"",""label"": ""Fen ve Teknoloji""},{""value"": ""1"",""label"": ""Fizik""},{""value"": ""5"",""label"": ""Geometri""},{""value"": ""11"",""label"": ""İngilizce""},{""value"": ""2"",""label"": ""Kimya""},{""value"": ""18"",""label"": ""Kur'an-ı Kerim""},{""value"": ""4"",""label"": ""Matematik""},{""value"": ""19"",""label"": ""Siyeri Nebi""},{""value"": ""16"",""label"": ""Sosyal Bilgiler""},{""value"": ""6"",""label"": ""Tarih""},{""value"": ""13"",""label"": ""Tüm Dersler""},{""value"": ""17"",""label"": ""Türk Dili ve Edebiyatı""},{""value"": ""9"",""label"": ""Türkçe""}]"
            If isBeginning = True Then
                html = "<!DOCTYPE html><html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge;"" /><meta charset=""utf-8""><meta http-equiv=""content-type"" content=""text/html; charset=utf-8"" /><meta content=""tr_TR"" http-equiv=""Content-Language""><title>AutoRun</title>"
                html &= "<HTA:APPLICATION ID=""oMyApp"" APPLICATIONNAME = ""Application Executer"" BORDER = ""no"" CAPTION = ""yes"" ShowInTaskbar = ""yes"" SINGLEINSTANCE = ""yes"" SYSMENU = ""yes"" Scroll = ""no"" WINDOWSTATE=""normal"" MAXIMIZEBUTTON=""no"" SELECTION=""no"" ICON=""Autorun.ico"" />"
                html &= "<style>body{background:url(Background2.jpg) no-repeat center center fixed;background-size:cover;padding:0;margin:0;}</style>"
                html &= "<LINK href=""Autorun.ico"" rel=""SHORTCUT ICON""/><link href=""Includes/Bookcase.min.css"" rel=""stylesheet""/></head><body>"
                html &= "<div class=""main-container""></div>"
                html &= "<script src=""Includes/jquery.min.js""></script>"
                html &= "<script src=""Includes/md5.js""></script>"
                html &= "<script src=""Includes/hammer.min.js""></script>"
                html &= "<script src=""Includes/jquery.hammer.min.js""></script>"
                html &= "<script src=""Includes/jquery.qrcode.min.js""></script>"
                html &= "<script src=""Includes/bookcase.min.js""></script>"
                html &= "<script type=""text/javascript"" language=""javascript"">window.resizeTo(1024,768);window.moveTo(0,0);window.onresize=function(){window.resizeTo(1024,768);};function RunFile(src){WshShell = new ActiveXObject(""WScript.Shell"");WshShell.Run(src, 1, false);window.close();};"

                html &= "(function ($) {$(function () {"
                html &= "var userData = {title: """ & yayinevi & """,name: """ & yayinevi & """, email: ""info@surat.gen.tr"",website: """ & web & """,logoLink: """ & web & """,logoAddress: ""Logo.png"",accountLogo: ""Logo.png"",skin: """","
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
                html &= "<HTA:APPLICATION ID=""oMyApp"" APPLICATIONNAME = ""Application Executer"" BORDER = ""no"" CAPTION = ""yes"" ShowInTaskbar = ""yes"" SINGLEINSTANCE = ""yes"" SYSMENU = ""yes"" Scroll = ""no"" WINDOWSTATE=""normal"" MAXIMIZEBUTTON=""no"" SELECTION=""no"" ICON=""Autorun.ico"" />"
                html &= "<style>body{background:url(Background.jpg) no-repeat center center fixed;background-size:cover;min-height:500px;padding:200px 0 0 0;margin:0;}h3{max-width:170px;margin:5px;}"
                html &= "#panel{width:" & genislik & "px;margin:0 auto;}.cerceve{width:180px;height:300px;padding:0;float:left;margin:0 10px;text-align:center;background:url(Background.png) no-repeat;background-size:100%;}img{border:0;}</style>"
                html &= "<LINK href=""Autorun.ico"" rel=""SHORTCUT ICON""/></head><body><div id=""panel"">"
            Else
                html = "</div><script type=""text/javascript"" language=""javascript"">window.resizeTo(1024,768);window.moveTo(0,0);window.onresize=function(){window.resizeTo(1024,768);};function RunFile(src){WshShell = new ActiveXObject(""WScript.Shell"");WshShell.Run(src, 1, false);window.close();};</script>"
                html &= "</body></html>"
            End If
        End If
        Return html
    End Function
End Module
