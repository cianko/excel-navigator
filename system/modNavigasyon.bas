Attribute VB_Name = "modNavigasyon"
'====================================================================
' Excel Navigator  -  (c) 2026 Ahmet Zan  -  MIT License
' Developed by Ahmet Zan
'====================================================================
' MODUL: modNavigasyon  (v6)
'  - Path: HER ZAMAN kokten su ana TUM zincir (gecmis + canli konum)
'  - Panel konumu hatirlanir (TEMP\nav_panel.txt), Excel'e goreli
'  - Olaylar ile panel Excel penceresi ICINDE kalir (clamp)
'  - X butonu: sadece paneli kapatir (Ctrl+Shift+N ile tekrar acilir)
'  - v6: "Scan for New Files" ozelligi kaldirildi (yeni dosyalar kur.ps1 ile
'        eklenir); NavKonumUygula'ya gPanelAcik guard'i eklendi.
'====================================================================
Option Explicit

Private Const HARITA_ADI  As String = "nav_map.txt"
Private Const GECMIS_ADI  As String = "nav_history.txt"
Private Const KONUM_ADI   As String = "nav_panel.txt"
Private Const KAYDET      As Boolean = True
Private Const KOK_ANAHTAR As String = "__ROOT__"
Private Const NAV_ONEK    As String = "nav_"

' --- panel durum degiskenleri (her dosyada kendi kopyasi; TEMP ile koprulenir) ---
Private gPanelAcik As Boolean
Private gDx As Double, gDy As Double            ' istenen goreli ofset (App'e gore)
Private gLastL As Double, gLastT As Double      ' son uygulanan form konumu
Private gLastAppL As Double, gLastAppT As Double, gLastAppW As Double, gLastAppH As Double
Private gTimerPlanli As Boolean, gTimerZaman As Double
Public gSecilenSayfa As String                  ' sayfa secim formu sonucu
Private gZincir As String                        ' path zinciri ONBELLEGI (dosya acilisinda 1 kez okunur)
Private gNavCalisiyor As Boolean                 ' navigasyon suruyor mu (re-entrant/cift-tetik korumasi)

Private Function AY() As String
    AY = " " & ChrW$(8250) & " "        ' path ayirici:  ">"  (U+203A)
End Function

'====================================================================
' PLATFORM / YOL
'====================================================================
Private Function Sep() As String
    Sep = Application.PathSeparator
End Function
Private Function WindowsMi() As Boolean
    WindowsMi = (InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0)
End Function
Private Function BaseAd(ByVal p As String) As String
    Dim s As String, k As Long
    s = Replace(p, "/", "\")
    k = InStrRev(s, "\")
    If k > 0 Then BaseAd = Mid$(s, k + 1) Else BaseAd = s
End Function
Private Function UzantisizAd(ByVal d As String) As String
    Dim k As Long
    k = InStrRev(d, ".")
    If k > 1 Then UzantisizAd = Left$(d, k - 1) Else UzantisizAd = d
End Function

Private Function HaritaDosyasiniBul() As String
    Dim klasor As String, aday As String, ust As String, p As Long
    klasor = ThisWorkbook.Path
    If klasor = "" Then Exit Function
    Do
        aday = klasor & Sep() & HARITA_ADI
        If Dir(aday) <> "" Then HaritaDosyasiniBul = aday: Exit Function
        p = InStrRev(klasor, Sep())
        If p <= 0 Then Exit Do
        ust = Left$(klasor, p - 1)
        If InStr(ust, Sep()) = 0 Then
            aday = ust & Sep() & HARITA_ADI
            If Dir(aday) <> "" Then HaritaDosyasiniBul = aday
            Exit Function
        End If
        If ust = klasor Or Len(ust) = 0 Then Exit Do
        klasor = ust
    Loop
End Function

Private Function YolCoz(ByVal anahtar As String, ByRef hataMsj As String) As String
    Dim haritaYolu As String, kok As String, ff As Integer, satir As String
    Dim esit As Long, a As String, sag As String, b As Long
    haritaYolu = HaritaDosyasiniBul()
    If haritaYolu = "" Then hataMsj = "Harita (" & HARITA_ADI & ") bulunamadi.": Exit Function
    kok = Left$(haritaYolu, InStrRev(haritaYolu, Sep()) - 1)
    ff = FreeFile
    On Error GoTo OkumaHata
    Open haritaYolu For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, satir
        satir = Trim$(satir)
        If Len(satir) > 0 And Left$(satir, 1) <> "#" And Left$(satir, 1) <> "@" Then
            esit = InStr(satir, "=")
            If esit > 0 Then
                a = Trim$(Left$(satir, esit - 1))
                If StrComp(a, anahtar, vbTextCompare) = 0 Then
                    sag = Trim$(Mid$(satir, esit + 1))
                    b = InStr(sag, "|")
                    If b > 0 Then sag = Trim$(Left$(sag, b - 1))
                    Close #ff
                    sag = Replace(sag, "\", Sep()): sag = Replace(sag, "/", Sep())
                    YolCoz = kok & Sep() & sag
                    Exit Function
                End If
            End If
        End If
    Loop
    Close #ff
    hataMsj = "Secili kutucuk haritada tanimli degil (" & anahtar & ")."
    Exit Function
OkumaHata:
    hataMsj = "Harita okunamadi: " & Err.Description
    On Error Resume Next
    Close #ff
End Function

Private Function DosyaEtiketi(ByVal tamYol As String) As String
    Dim haritaYolu As String, ff As Integer, satir As String
    Dim govde As String, esit As Long, sol As String, saga As String, hedefAd As String
    hedefAd = BaseAd(tamYol)
    DosyaEtiketi = UzantisizAd(hedefAd)
    haritaYolu = HaritaDosyasiniBul()
    If haritaYolu = "" Then Exit Function
    ff = FreeFile
    On Error GoTo Son
    Open haritaYolu For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, satir
        satir = Trim$(satir)
        If LCase$(Left$(satir, 6)) = "@file " Then
            govde = Trim$(Mid$(satir, 7))
            esit = InStr(govde, "=")
            If esit > 0 Then
                sol = Trim$(Left$(govde, esit - 1))
                saga = Trim$(Mid$(govde, esit + 1))
                If StrComp(BaseAd(sol), hedefAd, vbTextCompare) = 0 And saga <> "" Then
                    DosyaEtiketi = saga: Close #ff: Exit Function
                End If
            End If
        End If
    Loop
    Close #ff
    Exit Function
Son:
    On Error Resume Next
    Close #ff
End Function

'====================================================================
' NAVIGASYON GECMISI (Geri icin) - satir: dosya|Tab|sayfa|Tab|hucre|Tab|etiket
'====================================================================
Private Function TempKlasor() As String
    If WindowsMi() Then
        TempKlasor = Environ$("TEMP")
        If TempKlasor = "" Then TempKlasor = Environ$("TMP")
    Else
        TempKlasor = Environ$("TMPDIR")
    End If
End Function
Private Function GecmisYolu() As String
    Dim t As String: t = TempKlasor()
    If t <> "" Then GecmisYolu = t & Sep() & GECMIS_ADI
End Function
Private Function GecmiseEkle(ByVal dosya As String, ByVal sayfa As String, _
                             ByVal hucre As String, ByVal etiket As String) As Boolean
    Dim gy As String, ff As Integer
    gy = GecmisYolu()
    If gy = "" Then Exit Function
    On Error GoTo Hata
    ff = FreeFile
    Open gy For Append As #ff
    Print #ff, dosya & vbTab & sayfa & vbTab & hucre & vbTab & etiket
    Close #ff
    GecmiseEkle = True
    Exit Function
Hata:
    On Error Resume Next
    Close #ff
End Function
Private Function GecmisOku() As Collection
    Dim gy As String, ff As Integer, satir As String, c As Collection
    Set c = New Collection: Set GecmisOku = c
    gy = GecmisYolu()
    If gy = "" Then Exit Function
    If Dir(gy) = "" Then Exit Function
    ff = FreeFile
    On Error GoTo Hata
    Open gy For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, satir
        If Trim$(satir) <> "" Then c.Add satir
    Loop
    Close #ff
    Exit Function
Hata:
    On Error Resume Next
    Close #ff
End Function
Private Function GecmistenCek(ByRef bulundu As Boolean) As String
    Dim gy As String, ff As Integer, i As Long, c As Collection
    bulundu = False
    Set c = GecmisOku()
    If c.Count = 0 Then Exit Function
    GecmistenCek = c.Item(c.Count): bulundu = True
    gy = GecmisYolu()
    On Error GoTo Hata
    ff = FreeFile
    Open gy For Output As #ff
    For i = 1 To c.Count - 1
        Print #ff, c.Item(i)
    Next i
    Close #ff
    Exit Function
Hata:
    On Error Resume Next
    Close #ff
End Function
Private Sub GecmisiSil()
    Dim gy As String: gy = GecmisYolu()
    On Error Resume Next
    If gy <> "" Then If Dir(gy) <> "" Then Kill gy
    On Error GoTo 0
End Sub

' Cift-tetik korumasi: son navigasyonun TAMAMLANMASINDAN <0.4 sn gectiyse True
' (spurious ikinci tetigi yok say). Zaman damgasi hedef acilinca (NavInit) yazilir;
' insan bir navigasyon bitip goruntuyu gorup 0.4 sn'den once tekrar tiklayamaz,
' ama makine hizindaki ikinci OnTime bu pencere icinde kalir -> engellenir.
Private Function NavCokHizli() As Boolean
    On Error Resume Next
    Dim lp As String, ff As Integer, t As String, fark As Double
    lp = TempKlasor() & Sep() & "nav_lock.txt"
    If Dir(lp) <> "" Then
        ff = FreeFile
        Open lp For Input As #ff
        Line Input #ff, t
        Close #ff
        fark = Timer - Val(t)
        If fark >= 0 And fark < 0.4 Then NavCokHizli = True
    End If
    On Error GoTo 0
End Function

' Navigasyon TAMAMLANDI damgasi (hedef dosya acilinca NavInit'te yazilir).
Private Sub NavTamamlandiIsaretle()
    On Error Resume Next
    Dim lp As String, ff As Integer
    lp = TempKlasor() & Sep() & "nav_lock.txt"
    ff = FreeFile
    Open lp For Output As #ff
    Print #ff, Trim$(Str$(Timer))
    Close #ff
    On Error GoTo 0
End Sub

'====================================================================
' SECILI HUCRE / SECME
'====================================================================
Private Function AktifNavAdi() As String
    Dim nm As Name, r As Range, hc As Range
    On Error Resume Next
    Set hc = ActiveCell
    On Error GoTo 0
    If hc Is Nothing Then Exit Function
    For Each nm In ActiveWorkbook.Names
        If LCase$(Left$(nm.Name, Len(NAV_ONEK))) = NAV_ONEK Then
            Set r = Nothing
            On Error Resume Next
            Set r = nm.RefersToRange
            On Error GoTo 0
            If Not r Is Nothing Then
                If r.Worksheet.Name = hc.Worksheet.Name And r.Row = hc.Row And r.Column = hc.Column Then
                    AktifNavAdi = nm.Name: Exit Function
                End If
            End If
        End If
    Next nm
End Function
Private Sub HucreSec(ByVal wb As Object, ByVal sayfa As String, ByVal hucreAdi As String)
    On Error Resume Next
    wb.Worksheets(sayfa).Activate
    wb.Names(hucreAdi).RefersToRange.Select
    On Error GoTo 0
End Sub

'====================================================================
' CANLI PATH  (kokten su ana TUM zincir)
'====================================================================
Private Function GecmisZinciri() As String
    Dim c As Collection, i As Long, s As String, alan() As String
    Set c = GecmisOku()
    For i = 1 To c.Count
        alan = Split(c.Item(i), vbTab)
        If UBound(alan) >= 3 Then
            If s <> "" Then s = s & AY()
            s = s & DosyaEtiketi(alan(0)) & AY() & alan(1) & AY() & alan(3)
        End If
    Next i
    GecmisZinciri = s
End Function

Public Sub NavPathGuncelle()
    On Error Resume Next
    If Not gPanelAcik Then Exit Sub
    If ActiveWorkbook Is Nothing Then Exit Sub
    If ActiveSheet Is Nothing Then Exit Sub
    Dim s As String, zincir As String, guncel As String, nm As String
    zincir = gZincir                ' ONBELLEK'ten (secim degisince dosyaya DOKUNMAZ)
    guncel = DosyaEtiketi(ActiveWorkbook.FullName) & AY() & ActiveSheet.Name
    nm = AktifNavAdi()
    If nm <> "" Then guncel = guncel & AY() & ActiveCell.Text
    If zincir <> "" Then s = zincir & AY() & guncel Else s = guncel
    frmNav.lblPath.Caption = Kisalt(s)
    On Error GoTo 0
End Sub

Private Function Kisalt(ByVal s As String) As String
    Dim p() As String, n As Long, i As Long, r As String
    Const MAKS As Long = 8
    p = Split(s, AY())
    n = UBound(p)
    If (n + 1) <= MAKS Then Kisalt = s: Exit Function
    r = ChrW$(8230)   ' ...
    For i = n - MAKS + 1 To n
        r = r & AY() & p(i)
    Next i
    Kisalt = r
End Function

'====================================================================
' PANEL KONUMU (hatirlama + Excel icinde tutma)
'====================================================================
Private Function KonumDosyaYolu() As String
    Dim t As String: t = TempKlasor()
    If t <> "" Then KonumDosyaYolu = t & Sep() & KONUM_ADI
End Function
Private Function KonumYukle() As Boolean
    Dim fp As String, ff As Integer, satir As String, p() As String
    fp = KonumDosyaYolu()
    If fp = "" Then Exit Function
    If Dir(fp) = "" Then Exit Function
    On Error GoTo H
    ff = FreeFile
    Open fp For Input As #ff
    Line Input #ff, satir
    Close #ff
    p = Split(satir, "|")
    If UBound(p) >= 1 Then
        gDx = Val(p(0)): gDy = Val(p(1))
        KonumYukle = True
    End If
    Exit Function
H:
    On Error Resume Next
    Close #ff
End Function
Private Sub KonumKaydet()
    Dim fp As String, ff As Integer
    fp = KonumDosyaYolu()
    If fp = "" Then Exit Sub
    On Error GoTo H
    ff = FreeFile
    Open fp For Output As #ff
    Print #ff, Trim$(Str$(gDx)) & "|" & Trim$(Str$(gDy))
    Close #ff
    Exit Sub
H:
    On Error Resume Next
    Close #ff
End Sub

' Formu istenen ofsetten, Excel penceresi icinde kalacak sekilde konumlar.
Public Sub NavKonumUygula()
    On Error Resume Next
    If Not gPanelAcik Then Exit Sub      ' panel kapaliyken frmNav'i yeniden yukleme
    Dim L As Double, T As Double, minL As Double, maxL As Double, minT As Double, maxT As Double
    L = Application.Left + gDx
    T = Application.Top + gDy
    minL = Application.Left + 2
    maxL = Application.Left + Application.Width - frmNav.Width - 2
    If maxL < minL Then maxL = minL
    minT = Application.Top + 2
    maxT = Application.Top + Application.Height - frmNav.Height - 2
    If maxT < minT Then maxT = minT
    If L < minL Then L = minL
    If L > maxL Then L = maxL
    If T < minT Then T = minT
    If T > maxT Then T = maxT
    frmNav.Left = L: frmNav.Top = T
    gLastL = frmNav.Left: gLastT = frmNav.Top
    gLastAppL = Application.Left: gLastAppT = Application.Top
    gLastAppW = Application.Width: gLastAppH = Application.Height
    On Error GoTo 0
End Sub

' Kullanici surukledi mi / Excel boyutlandi mi -> ofseti guncelle ve uygula.
Public Sub NavKonumIzle()
    On Error Resume Next
    If Not gPanelAcik Then Exit Sub
    Dim appDegisti As Boolean, formDegisti As Boolean
    appDegisti = (Application.Left <> gLastAppL Or Application.Top <> gLastAppT _
                  Or Application.Width <> gLastAppW Or Application.Height <> gLastAppH)
    formDegisti = (Abs(frmNav.Left - gLastL) > 0.6 Or Abs(frmNav.Top - gLastT) > 0.6)
    If formDegisti And Not appDegisti Then
        gDx = frmNav.Left - Application.Left
        gDy = frmNav.Top - Application.Top
        KonumKaydet
    End If
    NavKonumUygula
    On Error GoTo 0
End Sub

Public Sub NavPanelKonumKaydet()
    On Error Resume Next
    If gPanelAcik Then
        gDx = frmNav.Left - Application.Left
        gDy = frmNav.Top - Application.Top
        KonumKaydet
    End If
    On Error GoTo 0
End Sub

'====================================================================
' PANEL ACMA / KAPAMA
'====================================================================
Public Sub NavInit()
    On Error Resume Next
    If gPanelAcik Then NavKonumUygula: NavPathGuncelle: Exit Sub
    ' Kok dosya acilinca eski gecmisi temizle (path zinciri sifirdan baslasin)
    If BuDosyaKokMu() Then GecmisiSil
    gZincir = GecmisZinciri()      ' zinciri dosya acilisinda BIR KEZ onbellege al
    Load frmNav
    If Not KonumYukle() Then
        gDx = Application.Width - frmNav.Width - 24   ' varsayilan: sag ust
        gDy = 72
    End If
    frmNav.Show vbModeless
    gPanelAcik = True
    NavKonumUygula
    NavPathGuncelle
    NavTamamlandiIsaretle            ' navigasyon/acilis TAMAMLANDI damgasi (debounce icin)
    Application.OnKey "^+n", "NavPaneliGoster"    ' Ctrl+Shift+N ile tekrar ac
    On Error GoTo 0
End Sub

' X butonu (QueryClose): konumu kaydet + timer durdur (form kendi unload olur).
Public Sub NavPanelKapat()
    On Error Resume Next
    NavPanelKonumKaydet
    gPanelAcik = False
    On Error GoTo 0
End Sub

' Workbook kapanirken (BeforeClose): formu da unload et ki kapanma engellenmesin.
Public Sub NavPanelUnload()
    On Error Resume Next
    NavPanelKapat
    Unload frmNav
    On Error GoTo 0
End Sub

' Kapali paneli tekrar acar (Ctrl+Shift+N).
Public Sub NavPaneliGoster()
    NavInit
End Sub

'====================================================================
' KAPATMA / KOK
'====================================================================
Private Sub BuDosyayiKapat()
    NavPanelKonumKaydet
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=KAYDET
    Application.DisplayAlerts = True
End Sub
Private Function BuDosyaKokMu() As Boolean
    Dim kokYol As String, h As String
    kokYol = YolCoz(KOK_ANAHTAR, h)
    If kokYol = "" Then Exit Function
    BuDosyaKokMu = (StrComp(kokYol, ThisWorkbook.FullName, vbTextCompare) = 0)
End Function
Private Sub AcVeKapat(ByVal hedefYol As String, ByVal sayfa As String, ByVal hucreAdi As String)
    Dim wbYeni As Object
    NavPanelKonumKaydet          ' yeni dosya guncel paneli okusun
    On Error GoTo AcmaHata
    Set wbYeni = Workbooks.Open(Filename:=hedefYol)
    On Error GoTo 0
    If sayfa <> "" Then HucreSec wbYeni, sayfa, hucreAdi
    If WindowsMi() Then
        BuDosyayiKapat
    Else
        On Error Resume Next
        BuDosyayiKapat
        On Error GoTo 0
    End If
    Exit Sub
AcmaHata:
    gNavCalisiyor = False        ' acilamadi -> kilidi coz ki tekrar denenebilsin
    MsgBox "Could not open file:" & vbCrLf & hedefYol & vbCrLf & vbCrLf & Err.Description, vbCritical, "Navigation"
End Sub

'====================================================================
' BAGLANTI KUR SIHIRBAZI (Link Cell)
'====================================================================
' Sadece harf/rakam/altcizgi birak (gecerli tanimli ad icin)
Private Function San(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            r = r & ch
        Else
            r = r & "_"
        End If
    Next i
    If r = "" Then r = "x"
    San = r
End Function

Private Function AdVarMi(ByVal nm As String) As Boolean
    Dim x As Object
    On Error Resume Next
    Set x = ThisWorkbook.Names(nm)
    On Error GoTo 0
    AdVarMi = Not (x Is Nothing)
End Function

' Kok klasor (nav_map.txt'nin klasoru; yoksa bu dosyanin klasoru)
Private Function KokKlasor() As String
    Dim m As String
    m = HaritaDosyasiniBul()
    If m <> "" Then KokKlasor = Left$(m, InStrRev(m, Sep()) - 1) Else KokKlasor = ThisWorkbook.Path
End Function

Private Function HaritaHedefYolu() As String
    Dim m As String
    m = HaritaDosyasiniBul()
    If m <> "" Then HaritaHedefYolu = m Else HaritaHedefYolu = ThisWorkbook.Path & Sep() & HARITA_ADI
End Function

' bazKlasor -> hedefTam icin goreceli yol (\ ile). disari=True ise koktan disari cikiyor.
Private Function GoreceliYol(ByVal bazKlasor As String, ByVal hedefTam As String, ByRef disari As Boolean) As String
    disari = False
    Dim hKlasor As String, hDosya As String, k As Long
    hedefTam = Replace(hedefTam, "/", "\")
    bazKlasor = Replace(bazKlasor, "/", "\")
    k = InStrRev(hedefTam, "\")
    If k = 0 Then GoreceliYol = hedefTam: Exit Function
    hKlasor = Left$(hedefTam, k - 1)
    hDosya = Mid$(hedefTam, k + 1)
    If Right$(bazKlasor, 1) = "\" Then bazKlasor = Left$(bazKlasor, Len(bazKlasor) - 1)
    If Right$(hKlasor, 1) = "\" Then hKlasor = Left$(hKlasor, Len(hKlasor) - 1)

    Dim bp() As String, hp() As String
    bp = Split(bazKlasor, "\")
    hp = Split(hKlasor, "\")
    If LCase$(bp(0)) <> LCase$(hp(0)) Then      ' farkli surucu -> goreceli imkansiz
        disari = True: GoreceliYol = "": Exit Function
    End If
    Dim ortak As Long, i As Long
    ortak = 0
    Do While ortak <= UBound(bp) And ortak <= UBound(hp)
        If LCase$(bp(ortak)) = LCase$(hp(ortak)) Then ortak = ortak + 1 Else Exit Do
    Loop
    Dim r As String
    For i = ortak To UBound(bp)
        r = r & "..\"
    Next i
    For i = ortak To UBound(hp)
        r = r & hp(i) & "\"
    Next i
    r = r & hDosya
    If InStr(r, "..\") > 0 Then disari = True   ' kok klasorun disina cikiyor
    GoreceliYol = r
End Function

' Hedef dosyanin sayfa adlarini dizi olarak doner.
Private Function SayfaAdlariniAl(ByVal path As String, ByRef hata As String) As Variant
    Dim wb As Object, ws As Object, arr() As String, n As Long
    Dim zatenAcik As Boolean, eskiEv As Boolean
    On Error Resume Next
    Set wb = Application.Workbooks(BaseAd(path))
    On Error GoTo 0
    If wb Is Nothing Then
        eskiEv = Application.EnableEvents
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        On Error GoTo AcHata
        Set wb = Application.Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo 0
        zatenAcik = False
    Else
        zatenAcik = True
    End If
    ReDim arr(0 To wb.Worksheets.Count - 1)
    n = 0
    For Each ws In wb.Worksheets
        arr(n) = ws.Name: n = n + 1
    Next ws
    If Not zatenAcik Then
        wb.Close SaveChanges:=False
        Application.EnableEvents = eskiEv
        Application.ScreenUpdating = True
    End If
    SayfaAdlariniAl = arr
    Exit Function
AcHata:
    hata = Err.Description
    Application.EnableEvents = eskiEv
    Application.ScreenUpdating = True
End Function

' frmSheetSec ile sayfa sec (bos = iptal)
Private Function SayfaSec(ByVal sayfalar As Variant) As String
    On Error Resume Next
    gSecilenSayfa = ""
    frmSheetSec.lstSheets.Clear
    Dim i As Long
    For i = LBound(sayfalar) To UBound(sayfalar)
        frmSheetSec.lstSheets.AddItem sayfalar(i)
    Next i
    If frmSheetSec.lstSheets.ListCount > 0 Then frmSheetSec.lstSheets.ListIndex = 0
    frmSheetSec.Show 1        ' vbModal
    SayfaSec = gSecilenSayfa
    On Error GoTo 0
End Function

' Haritanin hedefSayfa alanini (| sonrasi) doner
Private Function HedefSayfa(ByVal anahtar As String) As String
    Dim mp As String, ff As Integer, satir As String, a As String, esit As Long, sag As String, b As Long
    mp = HaritaDosyasiniBul()
    If mp = "" Then Exit Function
    ff = FreeFile
    On Error GoTo H
    Open mp For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, satir
        satir = Trim$(satir)
        If Len(satir) > 0 And Left$(satir, 1) <> "#" And Left$(satir, 1) <> "@" Then
            esit = InStr(satir, "=")
            If esit > 0 Then
                a = Trim$(Left$(satir, esit - 1))
                If StrComp(a, anahtar, vbTextCompare) = 0 Then
                    sag = Trim$(Mid$(satir, esit + 1))
                    b = InStr(sag, "|")
                    If b > 0 Then HedefSayfa = Trim$(Mid$(sag, b + 1))
                    Close #ff: Exit Function
                End If
            End If
        End If
    Loop
    Close #ff
    Exit Function
H:
    On Error Resume Next
    Close #ff
End Function

' Haritaya satir ekle/guncelle (ayni anahtar varsa degistirir; harita yoksa olusturur)
Private Sub HaritaSatirYaz(ByVal anahtar As String, ByVal rel As String, ByVal sheet As String)
    Dim mp As String, yeniMi As Boolean, c As Collection
    Dim ff As Integer, satir As String, s2 As String, a As String, esit As Long, i As Long, yeniSatir As String
    mp = HaritaHedefYolu()
    yeniMi = (Dir(mp) = "")
    Set c = New Collection
    If Not yeniMi Then
        ff = FreeFile
        Open mp For Input As #ff
        Do While Not EOF(ff)
            Line Input #ff, satir
            s2 = Trim$(satir)
            If Len(s2) > 0 And Left$(s2, 1) <> "#" And Left$(s2, 1) <> "@" Then
                esit = InStr(s2, "=")
                If esit > 0 Then
                    a = Trim$(Left$(s2, esit - 1))
                    If StrComp(a, anahtar, vbTextCompare) = 0 Then GoTo Devam
                End If
            End If
            c.Add satir
Devam:
        Loop
        Close #ff
    End If
    ff = FreeFile
    Open mp For Output As #ff
    If yeniMi Then
        Print #ff, "# Navigasyon haritasi (Link Cell sihirbazi ile olusturuldu)"
        Print #ff, "__ROOT__ = " & BaseAd(ThisWorkbook.FullName)
        Print #ff, "@FILE " & BaseAd(ThisWorkbook.FullName) & " = " & UzantisizAd(BaseAd(ThisWorkbook.FullName))
        Print #ff, ""
    Else
        For i = 1 To c.Count
            Print #ff, c.Item(i)
        Next i
    End If
    yeniSatir = anahtar & " = " & rel
    If sheet <> "" Then yeniSatir = yeniSatir & " | " & sheet
    Print #ff, yeniSatir
    Close #ff
End Sub

Private Sub HaritaSatirSil(ByVal anahtar As String)
    Dim mp As String, c As Collection, ff As Integer, satir As String, s2 As String, a As String, esit As Long, i As Long
    mp = HaritaDosyasiniBul()
    If mp = "" Then Exit Sub
    Set c = New Collection
    ff = FreeFile
    Open mp For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, satir
        s2 = Trim$(satir)
        If Len(s2) > 0 And Left$(s2, 1) <> "#" And Left$(s2, 1) <> "@" Then
            esit = InStr(s2, "=")
            If esit > 0 Then
                a = Trim$(Left$(s2, esit - 1))
                If StrComp(a, anahtar, vbTextCompare) = 0 Then GoTo Devam
            End If
        End If
        c.Add satir
Devam:
    Loop
    Close #ff
    ff = FreeFile
    Open mp For Output As #ff
    For i = 1 To c.Count
        Print #ff, c.Item(i)
    Next i
    Close #ff
End Sub

' ---- Panel "Link Cell" butonu ----
Public Sub NavBaglantiKur()
    On Error GoTo GenelHata
    Dim hc As Range
    On Error Resume Next
    Set hc = ActiveCell
    On Error GoTo GenelHata
    If hc Is Nothing Then
        MsgBox "Please select the cell you want to link first.", vbInformation, "Link Cell"
        Exit Sub
    End If

    Dim mevcut As String: mevcut = AktifNavAdi()
    If mevcut <> "" Then
        If MsgBox("This cell already has a link." & vbCrLf & _
                  "Do you want to remove the link?", vbYesNo + vbQuestion, "Link Cell") = vbYes Then
            NavBaglantiKaldir
        End If
        Exit Sub
    End If

    Dim f As Variant
    f = Application.GetOpenFilename("Excel files (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , "Select target file")
    If VarType(f) = vbBoolean Then Exit Sub
    Dim hedefTam As String: hedefTam = CStr(f)

    Dim hata As String, sayfalar As Variant
    sayfalar = SayfaAdlariniAl(hedefTam, hata)
    If hata <> "" Then
        MsgBox "Could not read the target file's sheets:" & vbCrLf & hata, vbExclamation, "Link Cell"
        Exit Sub
    End If
    Dim secilen As String
    If (UBound(sayfalar) - LBound(sayfalar)) = 0 Then
        secilen = sayfalar(LBound(sayfalar))
    Else
        secilen = SayfaSec(sayfalar)
        If secilen = "" Then Exit Sub
    End If

    Dim disari As Boolean, rel As String
    rel = GoreceliYol(KokKlasor(), hedefTam, disari)
    If rel = "" Then
        MsgBox "Could not compute a relative path (target may be on a different drive).", vbExclamation, "Link Cell"
        Exit Sub
    End If
    If disari Then
        If MsgBox("The target file is OUTSIDE the project folder." & vbCrLf & _
                  "The link may break if you move the files." & vbCrLf & vbCrLf & _
                  "Continue anyway?", vbYesNo + vbExclamation, "Link Cell") <> vbYes Then Exit Sub
    End If

    Dim nm As String, taban As String, kk As Long
    ' Anahtar GLOBAL benzersiz olmali (nav_map anahtarlari tum dosyalarda ortak):
    ' kaynak DOSYA ADINI da kat, yoksa ayni adresteki hucreler cakisir.
    nm = "nav_" & San(UzantisizAd(BaseAd(ThisWorkbook.FullName))) & "_" & _
         San(hc.Worksheet.Name) & "_" & San(hc.Address(False, False))
    taban = nm: kk = 1
    Do While AdVarMi(nm)
        kk = kk + 1: nm = taban & "_" & kk
    Loop

    On Error Resume Next
    hc.Hyperlinks.Delete
    On Error GoTo GenelHata
    ThisWorkbook.Names.Add Name:=nm, _
        RefersTo:="='" & Replace(hc.Worksheet.Name, "'", "''") & "'!" & hc.Address(True, True)
    hc.Font.Color = RGB(0, 102, 204)
    hc.Font.Underline = xlUnderlineStyleSingle

    HaritaSatirYaz nm, rel, secilen
    NavPathGuncelle
    MsgBox "Link created:" & vbCrLf & _
           "'" & hc.Text & "'  ->  " & UzantisizAd(BaseAd(hedefTam)) & "  >  " & secilen, _
           vbInformation, "Link Cell"
    Exit Sub
GenelHata:
    MsgBox "Could not create link:" & vbCrLf & Err.Description, vbCritical, "Link Cell"
End Sub

' ---- Bagli hucreden baglantiyi kaldir ----
Public Sub NavBaglantiKaldir()
    On Error GoTo GenelHata
    Dim hc As Range: Set hc = ActiveCell
    If hc Is Nothing Then Exit Sub
    Dim nm As String: nm = AktifNavAdi()
    If nm = "" Then
        MsgBox "There is no link to remove on this cell.", vbInformation, "Link Cell"
        Exit Sub
    End If
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo GenelHata
    hc.Font.Underline = xlUnderlineStyleNone
    hc.Font.ColorIndex = xlAutomatic
    HaritaSatirSil nm
    NavPathGuncelle
    MsgBox "Link removed.", vbInformation, "Link Cell"
    Exit Sub
GenelHata:
    MsgBox "Could not remove link:" & vbCrLf & Err.Description, vbCritical, "Link Cell"
End Sub

'====================================================================
' ANA MAKROLAR (panel butonlari OnTime ile bunlari cagirir)
'====================================================================
Public Sub NavIleri()
    Dim navName As String, hedef As String, hataMsj As String
    Dim oDosya As String, oSayfa As String, oEtiket As String
    If gNavCalisiyor Then Exit Sub              ' re-entrant/cift-tetik: ikinciyi yok say
    On Error GoTo GenelHata
    If NavCokHizli() Then Exit Sub
    gNavCalisiyor = True
    navName = AktifNavAdi()
    If navName = "" Then
        gNavCalisiyor = False
        MsgBox "Please select a menu cell first, then click Forward.", vbInformation, "Forward"
        Exit Sub
    End If
    hedef = YolCoz(navName, hataMsj)
    If hedef = "" Then gNavCalisiyor = False: MsgBox hataMsj, vbExclamation, "Forward": Exit Sub
    If Dir(hedef) = "" Then gNavCalisiyor = False: MsgBox "Target file not found:" & vbCrLf & hedef, vbExclamation, "Forward": Exit Sub
    oDosya = ThisWorkbook.FullName
    oSayfa = ActiveSheet.Name
    oEtiket = ""
    On Error Resume Next
    oEtiket = ActiveCell.Text
    On Error GoTo GenelHata
    If oEtiket = "" Then oEtiket = navName
    If BuDosyaKokMu() Then GecmisiSil
    GecmiseEkle oDosya, oSayfa, navName, oEtiket
    AcVeKapat hedef, HedefSayfa(navName), ""     ' basarili -> bu dosya kapanir (bayrak onunla gider)
    Exit Sub
GenelHata:
    gNavCalisiyor = False
    MsgBox "An unexpected error occurred:" & vbCrLf & Err.Description, vbCritical, "Forward"
End Sub

Public Sub NavGeri()
    Dim bulundu As Boolean, satir As String, alan() As String
    If gNavCalisiyor Then Exit Sub              ' re-entrant/cift-tetik: ikinciyi yok say
    On Error GoTo GenelHata
    If NavCokHizli() Then Exit Sub
    gNavCalisiyor = True
    satir = GecmistenCek(bulundu)
    If Not bulundu Then gNavCalisiyor = False: MsgBox "There is no parent location to go back to.", vbInformation, "Back": Exit Sub
    alan = Split(satir, vbTab)
    If UBound(alan) < 3 Then gNavCalisiyor = False: MsgBox "Could not read the history entry.", vbExclamation, "Back": Exit Sub
    If Dir(alan(0)) = "" Then gNavCalisiyor = False: MsgBox "Parent file not found:" & vbCrLf & alan(0), vbExclamation, "Back": Exit Sub
    AcVeKapat alan(0), alan(1), alan(2)         ' basarili -> bu dosya kapanir (bayrak onunla gider)
    Exit Sub
GenelHata:
    gNavCalisiyor = False
    MsgBox "An unexpected error occurred:" & vbCrLf & Err.Description, vbCritical, "Back"
End Sub

Public Sub NavGecmisiSifirla()
    GecmisiSil
    MsgBox "Navigation history reset.", vbInformation, "Navigation"
End Sub
