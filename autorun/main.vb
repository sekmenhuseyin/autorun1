Imports System.IO
Imports System.Net
Imports System.Xml

Module main
    Dim lstSinif As List(Of String) = New List(Of String) : Dim lstBrans As List(Of String) = New List(Of String) : Dim lstTur As List(Of String) = New List(Of String)
    Dim lstSinifID As List(Of String) = New List(Of String) : Dim lstBransID As List(Of String) = New List(Of String) : Dim lstTurID As List(Of String) = New List(Of String)
    Dim HtaFolder As String = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\Sürat\"
    Dim AppAddress As String = My.Application.Info.DirectoryPath : Dim isler As List(Of Integer) = New List(Of Integer)
    Sub Main()
        Dim isLibrary As Boolean = True
        Dim xmldoc As New XmlDocument : Dim xmlnode As XmlNodeList
        Dim No, Yayinevi, KitapIsmi, InternetSitesi, SonSayfa, Sinif(1), Tur(1), Brans(1), title, web, html, version As String : Dim i As Integer = 0
        html = "" : title = "" : web = ""
        'control
        If Directory.Exists(HtaFolder) = False Then Directory.CreateDirectory(HtaFolder)
        If File.Exists(HtaFolder & "setting.ini") = False Then CreateSettings()
        Dim CheckForUpdates As Boolean = AskUpdates()
        'search in directories
        Dim dirs As String() = Directory.GetDirectories(AppAddress)
        For Each Dir As String In dirs
            'eğer atu klasörü ise işleme devam et
            If File.Exists(Dir & "\atu.exe") And File.Exists(Dir & "\Book\Data\info.xml") Then
                Dim tmp() As String = Dir.Split("\") : Dim klasoryolu As String = tmp(UBound(tmp))
                No = "" : Yayinevi = "" : KitapIsmi = "" : InternetSitesi = "" : SonSayfa = "" : Sinif(0) = "" : Tur(0) = "" : Brans(0) = "" : Sinif(1) = "" : Tur(1) = "" : Brans(1) = "" : version = "1.0" : i += 1
                ClearXml(Dir)
                'her kitap için htmlde bir satır oluştur
                'xmlden okuma bmlünü
                Dim fs As New FileStream(Dir & "\Book\Data\info.xml", FileMode.Open, FileAccess.Read)
                'xml yükle
                xmldoc.Load(fs) : xmlnode = xmldoc.GetElementsByTagName("info")
                'bilgileri al
                If IsNothing(xmlnode(0).Item("No")) = False Then No = xmlnode(0).Item("No").InnerText
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
                If IsNothing(xmlnode(0).Item("Surum")) = False Then version = xmlnode(0).Item("Surum").InnerText
                'genel bilgiler
                If No = "" Then No = i Else writeIni(HtaFolder & "setting.ini", "Versions", No, version) : isler.Add(No) 'eğer işler no varsa sürümünü yaz
                If title = "" Then title = Yayinevi
                If web = "" Then web = InternetSitesi
                If lstSinif.Contains(Sinif(1)) = False Then lstSinif.Add(Sinif(1)) : lstSinifID.Add(Sinif(0))
                If lstBrans.Contains(Brans(1)) = False Then lstBrans.Add(Brans(1)) : lstBransID.Add(Brans(0))
                If lstTur.Contains(Tur(1)) = False Then lstTur.Add(Tur(1)) : lstTurID.Add(Tur(0))
                'her kitap için bir satır yaz
                If isLibrary = True Then
                    If html <> "" Then html &= ","
                    html &= WriteBook2(No, KitapIsmi, klasoryolu, Dir & "\atu.exe", SonSayfa, Tur(1), Tur(0), Brans(1), Brans(0), Sinif(1), Sinif(0))
                Else
                    html &= WriteBook(Dir & "\atu.exe", Yayinevi.Replace(" Yayınları", "").Replace(" Yayın", "").Replace(" Publishing", "") & " " & KitapIsmi, klasoryolu)
                End If
                'kapat
                fs.Close() : fs.Dispose()
            End If
        Next
        xmlnode = Nothing : xmldoc = Nothing : dirs = Nothing
        'kitap varsa programı çalıştır
        If i > 0 Then
            'genişliği bul
            Dim genislik As Integer = 200 * i : If genislik > 1000 Then genislik = 1000
            'sayfayı yazdır
            html = WriteHta(isLibrary, True, genislik, title, web) & html & WriteHta(isLibrary, False)
            'hta yazdır
            File.Delete(HtaFolder & "Autorun.hta")
            Dim objWriter As New StreamWriter(HtaFolder & "Autorun.hta")
            objWriter.Write(html)
            objWriter.Close()
            've başlat
            Shell("explorer.exe " & HtaFolder & "Autorun.hta")
        End If
        'güncelleme zamanı
        If CheckForUpdates = True Then CheckUpdate()
    End Sub
    'az kitap için sadece resmi ve adı yazılıyor
    Private Function WriteBook(ExeFile As String, title As String, foldername As String) As String
        Dim kapak As String = IIf(File.Exists(AppAddress & "\" & foldername & "\Book\Images\Pages\Thumbs\Cover.etf") = True, "file:///" & AppAddress.Replace("\", "/") & "/" & foldername & "/Book/Images/Pages/Thumbs/cover.etf", "file:///" & AppAddress.Replace("\", "/") & "/" & foldername & "/Book/Images/Pages/Thumbs/Sayfa1.etf")
        Return "<div class=""cerceve""><a href=""#"" onclick=""RunFile('file:///" & ExeFile.Replace("\", "/") & "');"" title=""" & title & """><img src=""" & kapak & """ height=""180"" /></a><h3>" & title & "</h3></div>"
    End Function
    'kitaplık için tüm bilgileri yazılıyor
    Private Function WriteBook2(id As Integer, title As String, foldername As String, ExeFile As String, SonSayfa As String, category As String, categoryid As String, brans As String, bransid As String, sinif As String, sinifid As String) As String
        Dim aciklama = "Çözümlü Sorular ve cevap anahtarları eksiklerinizi göstermez Konu sütunlarına yerleştirilen anahtar bilgiler ile başka bir kaynağa ihtiyaç duymadan cevabı kendiniz bularak konuyu pekiştirebilirsiniz."
        Dim kapak As String = "file:///" & AppAddress.Replace("\", "/") & "/" & foldername & "/Book/Images/Pages/Thumbs/cover.etf"
        Dim url As String = "file:///" & ExeFile.Replace("\", "/")
        Return "{""id"":""" & id & """,""title"":""" & title & """,""description"":""" & aciklama & """,""url"":""" & url & """,""pages"":""" & SonSayfa & """,""kapak"":""" & kapak & """,""categoryname"":""" & category & """,""categoryid"":""" & categoryid & """,""brans"":""" & brans & """,""bransid"":""" & bransid & """,""sinif"":""" & sinif & """,""sinifid"":""" & sinifid & """,""label"":""0""}"
    End Function
    'web sayfasını oluşturur
    Private Function WriteHta(isLibrary As Boolean, isBeginning As Boolean, Optional genislik As Integer = 1000, Optional yayinevi As String = "", Optional web As String = "") As String
        Dim html As String = ""
        If isLibrary = True Then
            If isBeginning = True Then
                html = "<!DOCTYPE html><html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge;"" /><meta charset=""utf-8""><meta http-equiv=""content-type"" content=""text/html; charset=utf-8"" /><meta content=""tr_TR"" http-equiv=""Content-Language""><title>" & yayinevi & "</title>"
                html &= "<HTA:APPLICATION ID=""oMyApp"" APPLICATIONNAME = ""Application Executer"" BORDER = ""no"" CAPTION = ""yes"" ShowInTaskbar = ""yes"" SINGLEINSTANCE = ""yes"" SYSMENU = ""yes"" Scroll = ""no"" WINDOWSTATE=""normal"" MAXIMIZEBUTTON=""Yes"" SELECTION=""no"" ICON=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Autorun.ico"" />"
                html &= "<style>body{background:url(file:///" & AppAddress.Replace("\", "/") & "/Autorun/Background2.jpg) no-repeat center center fixed;background-size:cover;padding:0;margin:0;}</style>"
                html &= "<LINK href=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Autorun.ico"" rel=""SHORTCUT ICON""/><link href=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Includes/Bookcase.min.css"" rel=""stylesheet""/></head><body>"
                html &= "<div class=""main-container""></div>"
                html &= "<script src=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Includes/jquery.min.js""></script>"
                html &= "<script src=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Includes/md5.js""></script>"
                html &= "<script src=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Includes/hammer.min.js""></script>"
                html &= "<script src=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Includes/jquery.hammer.min.js""></script>"
                html &= "<script src=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Includes/jquery.qrcode.min.js""></script>"
                html &= "<script src=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Includes/bookcase.min.js""></script>"
                html &= "<script type=""text/javascript"" language=""javascript"">window.resizeTo(1100,770);window.moveTo(0,0);function RunFile(src){WshShell = new ActiveXObject(""WScript.Shell"");WshShell.Run(src, 1, false);window.close();};"

                html &= "(function ($) {$(function () {"
                html &= "var userData = {title:""" & yayinevi & """,name:""" & yayinevi & """, email:""info@surat.gen.tr"",website:""" & web & """,logoLink:""" & web & """,logoAddress:""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Logo.png"",accountLogo:""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Logo.jpg"",skin:"""",path:""file:///" & AppAddress.Replace("\", "/") & ""","
                html &= "isShowShare:""1"",isShowContact:""1"",isShowSearch:""1"",isShowSkin:""0"",isShowCategory:""1"",isShowLogo:""1"",openType:""0"",isSelf:""0"",about:""İlkokuldan Üniversiteye Kadar""};"
                'function Bookcase($container, userData, data, cateData, classData, departmentData)
                html &= "new Bookcase($("".main-container""),userData,["
            Else
                'sorting
                Dim tmp As String
                For i = 0 To lstBrans.Count - 2
                    For j = i To lstBrans.Count - 1
                        If lstBrans(i) > lstBrans(j) Then
                            tmp = lstBrans(j) : lstBrans(j) = lstBrans(i) : lstBrans(i) = tmp
                            tmp = lstBransID(j) : lstBransID(j) = lstBransID(i) : lstBransID(i) = tmp
                        End If
                    Next
                Next
                For i = 0 To lstTur.Count - 2
                    For j = i To lstTur.Count - 1
                        If lstTur(i) > lstTur(j) Then
                            tmp = lstTur(j) : lstTur(j) = lstTur(i) : lstTur(i) = tmp
                            tmp = lstTurID(j) : lstTurID(j) = lstTurID(i) : lstTurID(i) = tmp
                        End If
                    Next
                Next
                'auto list combobox contents
                Dim tmpTur As String = "[" : Dim tmpSinif As String = "[" : Dim tmpBrans As String = "["
                For i = 0 To lstBrans.Count - 1 : tmpBrans &= "{""value"":""" & lstBransID(i) & """,""label"":""" & lstBrans(i) & """}," : Next
                For i = 0 To lstSinif.Count - 1 : tmpSinif &= "{""value"":""" & lstSinifID(i) & """,""label"":""" & lstSinif(i) & """}," : Next
                For i = 0 To lstTur.Count - 1 : tmpTur &= "{""value"":""" & lstTurID(i) & """,""label"":""" & lstTur(i) & """}," : Next
                tmpBrans &= "]" : tmpSinif &= "]" : tmpTur &= "]"
                'contiune writing
                html = "]," & tmpTur & "," & tmpSinif & "," & tmpBrans & ");"
                html &= "});})(window.jQuery);"
                html &= "</script></body></html>"
            End If

        Else
            If isBeginning = True Then
                html = "<!DOCTYPE html><html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge;"" /><meta charset=""utf-8""><meta http-equiv=""content-type"" content=""text/html; charset=utf-8"" /><meta content=""tr_TR"" http-equiv=""Content-Language""><title>" & yayinevi & "</title>"
                html &= "<HTA:APPLICATION ID=""oMyApp"" APPLICATIONNAME = ""Application Executer"" BORDER = ""no"" CAPTION = ""yes"" ShowInTaskbar = ""yes"" SINGLEINSTANCE = ""yes"" SYSMENU = ""yes"" Scroll = ""no"" WINDOWSTATE=""normal"" MAXIMIZEBUTTON=""no"" SELECTION=""no"" ICON=""file:///" & AppAddress.Replace("\", "/") & "/Autorun/Autorun.ico"" />"
                html &= "<style>body{background:url(file:///" & AppAddress.Replace("\", "/") & "/Autorun/Background.jpg) no-repeat center center fixed;background-size:cover;min-height:500px;padding:200px 0 0 0;margin:0;}h3{max-width:170px;margin:5px;}"
                html &= "#panel{width:" & genislik & "px;margin:0 auto;}.cerceve{width:180px;height:300px;padding:0;float:left;margin:0 10px;text-align:center;background:url(file:///" & AppAddress.Replace("\", "/") & "/Autorun/Background.png) no-repeat;background-size:100%;}img{border:0;}</style>"
                html &= "<LINK href=""" & AppAddress.Replace("\", "/") & "/Autorun/Autorun.ico"" rel=""SHORTCUT ICON""/></head><body><div id=""panel"">"
            Else
                html = "</div><script type=""text/javascript"" language=""javascript"">window.resizeTo(1024,768);window.moveTo(0,0);window.onresize=function(){window.resizeTo(1024,768);};function RunFile(src){WshShell = new ActiveXObject(""WScript.Shell"");WshShell.Run(src, 1, false);window.close();};</script>"
                html &= "</body></html>"
            End If
        End If
        Return html
    End Function
    'Info.xml kontrol
    Private Sub ClearXml(dir As String)
        'eğer xmlde ilk satır indo değilse hata veriyordu
        'bu yüzden infoda öncesini sildiriyorum
        Dim objWriter As New StreamReader(dir & "\Book\Data\info.xml")
        Dim alltext = objWriter.ReadToEnd
        Dim pos = InStr(alltext, "<info>")
        If pos <> 1 Then
            objWriter.Close() : objWriter.Dispose() 'önce açık olanı kapat
            alltext = Mid(alltext, pos) 'info yazısından itibaren yazdır
            Dim objWriter2 As New StreamWriter(dir & "\Book\Data\info.xml")
            objWriter2.Write(alltext)
            objWriter2.Close() : objWriter2.Dispose()
        Else
            objWriter.Close() : objWriter.Dispose()
        End If
    End Sub
    'create default settings file
    Private Sub CreateSettings()
        writeIni(HtaFolder & "setting.ini", "General", "LastCheck", Now())
        writeIni(HtaFolder & "setting.ini", "General", "CountOfNos", "0")
    End Sub
    'güncelleme yapsın mı
    Private Function AskUpdates() As Boolean
        Dim CheckForUpdates As Boolean = False
        Dim LastCheck As DateTime = Convert.ToDateTime(ReadIni(HtaFolder & "setting.ini", "General", "LastCheck"))
        Dim CountOfNos As Int16 = Convert.ToInt16(ReadIni(HtaFolder & "setting.ini", "General", "CountOfNos"))
        If LastCheck < Now().AddDays(-30) Then
            If MsgBox("Güncellemeleri kontrol etmek ister misiniz?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.MsgBoxSetForeground, "Autorun") = MsgBoxResult.Yes Then
                CheckForUpdates = True
                writeIni(HtaFolder & "setting.ini", "General", "CountOfNos", "0")
            Else
                writeIni(HtaFolder & "setting.ini", "General", "CountOfNos", CountOfNos + 1)
                If CountOfNos > 3 Then writeIni(HtaFolder & "setting.ini", "General", "LastCheck", Now().AddMonths(6)) : writeIni(HtaFolder & "setting.ini", "General", "CountOfNos", "0")
            End If
        End If
        Return CheckForUpdates
    End Function
    'ne kadar güncelleme var?
    Private Sub CheckUpdate()
        Dim webClient As New WebClient : Dim KitapSayisi As Integer = 0 : Dim version As String
        'autorun.exe güncelleme kontrol
        version = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor
        Dim updateAutorun As String = webClient.DownloadString("http://kaynakkatalog.com/API/CheckUpdates").Replace("""", "")
        If updateAutorun <> version Then updateAutorun = "yes"
        'atu.exe güncelleme kontrol
        'tüm kitapları tek tek kontrol et
        For Each item In isler
            version = ReadIni(HtaFolder & "setting.ini", "General", item) : If version = "" Then version = "1.0"
            Dim updateKitap As String = webClient.DownloadString("http://kaynakkatalog.com/API/CheckUpdates/" & item).Replace("""", "") : If updateKitap = "" Then updateKitap = "1.0"
            If updateKitap <> version Then KitapSayisi += 1
        Next
        Dim tmp As String = ""
        If updateAutorun = "yes" Then tmp = "Program güncellenecek." & vbCrLf
        If KitapSayisi > 0 Then tmp &= KitapSayisi & " adet kitap güncellenecek."
        If tmp <> "" Then
            If MsgBox(tmp, MsgBoxStyle.YesNo + MsgBoxStyle.Question + MsgBoxStyle.MsgBoxSetForeground, "Autorun") = MsgBoxResult.Yes Then
                Update()
                writeIni(HtaFolder & "setting.ini", "General", "LastCheck", Now())
            End If
        End If
        Exit Sub
    End Sub
    'güncelleme
    Private Sub Update()

    End Sub
End Module