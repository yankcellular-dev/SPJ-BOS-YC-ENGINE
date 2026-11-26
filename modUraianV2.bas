Attribute VB_Name = "Moduraian"
Option Explicit

Public Sub BuatDaftarNamaDropdown()
    Dim Wb As Workbook
    Dim k As Long, s As Long, e As Long, t As Long
    Dim shD As String

    Set Wb = ThisWorkbook

    For t = 1 To 2
        shD = "dbekstra-t" & t
        For k = 1 To 10
            s = 4 + 13 * (k - 1)
            e = s + 9
            Wb.Names.Add Name:="EksT" & t & "_" & Format(k, "00"), _
                RefersTo:="=OFFSET('" & shD & "'!$D$" & s & ",0,0,MAX(1,COUNTA('" & shD & "'!$D$" & s & ":$D$" & e & ")),1)"
        Next k
    Next t

    Wb.Names.Add Name:="HonT1", RefersTo:="=OFFSET('dbhonor-t1'!$D$3,0,0,MAX(1,COUNTA('dbhonor-t1'!$D$3:$D$12)),1)"
    Wb.Names.Add Name:="HonT2", RefersTo:="=OFFSET('dbhonor-t2'!$D$3,0,0,MAX(1,COUNTA('dbhonor-t2'!$D$3:$D$12)),1)"
    Wb.Names.Add Name:="TukT1", RefersTo:="=OFFSET('dbtukang-t1'!$D$3,0,0,MAX(1,COUNTA('dbtukang-t1'!$D$3:$D$12)),1)"
    Wb.Names.Add Name:="TukT2", RefersTo:="=OFFSET('dbtukang-t2'!$D$3,0,0,MAX(1,COUNTA('dbtukang-t2'!$D$3:$D$12)),1)"
End Sub

Private Function NamaAda(ByVal nm As String) As Boolean
    Dim x As Name
    On Error Resume Next
    Set x = ThisWorkbook.Names(nm)
    NamaAda = (Err.Number = 0)
    Err.Clear
End Function

Private Function TahapKe(ws As Worksheet) As Long
    Select Case LCase$(Right$(ws.Name, 3))
        Case "jan", "feb", "mar", "apr", "mei", "jun": TahapKe = 1
        Case Else: TahapKe = 2
    End Select
End Function

Private Function NamaDaftarUntukBaris(ws As Worksheet, r As Long) As String
    Dim t As String, tahap As Long, m3 As Worksheet
    Dim i As Long, nm As String, bestIdx As Long, bestLen As Long

    ws.Cells(r, 6).Calculate                 ' pastikan nilai F sudah terhitung
    tahap = TahapKe(ws)
    t = LCase$(ws.Cells(r, 6).text)

    If InStr(t, "honor") > 0 Then
        NamaDaftarUntukBaris = "Hon" & "T" & tahap
        Exit Function
    End If

    If InStr(t, "upah ekstra") > 0 Then
        Set m3 = ThisWorkbook.Worksheets("menu3")
        For i = 7 To 16
            If tahap = 1 Then
                nm = Trim$(CStr(m3.Cells(i, 5).Value))     ' kolom E
            Else
                nm = Trim$(CStr(m3.Cells(i, 14).Value))    ' kolom N
            End If
            If Len(nm) > 0 Then
                If InStr(t, LCase$(nm)) > 0 And Len(nm) > bestLen Then
                    bestLen = Len(nm)
                    bestIdx = i - 6
                End If
            End If
        Next i
        If bestIdx > 0 Then NamaDaftarUntukBaris = "EksT" & tahap & "_" & Format(bestIdx, "00")
    End If
End Function

Public Sub PasangSemuaDropdown()
    Dim nama As Variant, ws As Worksheet, r As Long

    If Not NamaAda("EksT1_01") Then BuatDaftarNamaDropdown

    For Each nama In Array("U-jan", "U-feb", "U-mar", "U-apr", "U-mei", "U-jun", _
                           "U-jul", "U-ags", "U-sep", "U-okt", "U-nov", "U-des")
        Set ws = ThisWorkbook.Worksheets(nama)
        If BukaSheetAman(ws) Then
            Application.EnableEvents = False
            For r = 3 To 52
                PasangDropdownBaris ws, r
            Next r
            Application.EnableEvents = True
            KunciSheetAman ws
        End If
    Next nama
End Sub
Public Function IsPenerimaCol(ByVal col As Long) As Boolean
    Select Case col
        Case 10, 14, 18, 22, 26, 30, 34, 38, 42, 46   ' J,N,R,V,Z,AD,AH,AL,AP,AT
            IsPenerimaCol = True
    End Select
End Function

Public Sub PasangDropdownBaris(ws As Worksheet, r As Long)
    Dim cols As Variant, c As Variant, nm As String

    cols = Array(10, 14, 18, 22, 26, 30, 34, 38, 42, 46)

    For Each c In cols
        With ws.Cells(r, c)
            On Error Resume Next
            .Validation.Delete
            On Error GoTo 0

            nm = NamaDaftarUntukBaris(ws, r)
            If nm <> "" Then
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
                    Formula1:="=" & nm
                .Validation.ErrorTitle = "Nama tidak ada di daftar"
                .Validation.ErrorMessage = "Nama ini tidak ada di database. Lanjutkan?"
                .Validation.InCellDropdown = True
            End If
        End With
    Next c
End Sub

Private Function BlokSumber(ws As Worksheet, r As Long, ByRef shName As String, _
                            ByRef s As Long, ByRef e As Long) As Boolean
    Dim t As String, tahap As Long, m3 As Worksheet
    Dim i As Long, nm As String, bestIdx As Long, bestLen As Long

    ws.Cells(r, 6).Calculate
    tahap = TahapKe(ws)
    t = LCase$(ws.Cells(r, 6).text)

    If InStr(t, "honor") > 0 Then
        shName = "dbhonor-t" & tahap
        s = 3: e = 12
        BlokSumber = True
        Exit Function
    End If

    If InStr(t, "upah ekstra") > 0 Then
        Set m3 = ThisWorkbook.Worksheets("menu3")
        For i = 7 To 16
            If tahap = 1 Then
                nm = Trim$(CStr(m3.Cells(i, 5).Value))
            Else
                nm = Trim$(CStr(m3.Cells(i, 14).Value))
            End If
            If Len(nm) > 0 Then
                If InStr(t, LCase$(nm)) > 0 And Len(nm) > bestLen Then
                    bestLen = Len(nm)
                    bestIdx = i - 6
                End If
            End If
        Next i
        If bestIdx > 0 Then
            shName = "dbekstra-t" & tahap
            s = 4 + 13 * (bestIdx - 1)
            e = s + 9
            BlokSumber = True
        End If
        Exit Function
    End If

    If InStr(t, "upah pekerja") > 0 Then
        shName = "dbtukang-t" & tahap
        s = 3: e = 12
        BlokSumber = True
    End If
End Function
Public Sub IsiJabatanTarif(ws As Worksheet, r As Long, colNama As Long)
    Dim shName As String, s As Long, e As Long, nm As String
    Dim db As Worksheet, idx As Variant

    nm = Trim$(CStr(ws.Cells(r, colNama).Value))

    If nm = "" Then
        ws.Cells(r, colNama + 1).ClearContents    ' jabatan
        ws.Cells(r, colNama + 2).ClearContents    ' tarif
        Exit Sub
    End If

    If Not BlokSumber(ws, r, shName, s, e) Then Exit Sub

    Set db = ThisWorkbook.Worksheets(shName)
    idx = Application.Match(nm, db.Range("D" & s & ":D" & e), 0)
    If IsError(idx) Then Exit Sub                 ' tidak ada di daftar: biarkan isian manual

    ws.Cells(r, colNama + 1).Value = db.Cells(s + idx - 1, 5).Value   ' kolom E = jabatan
    ws.Cells(r, colNama + 2).Value = db.Cells(s + idx - 1, 6).Value   ' kolom F = tarif
End Sub

Public Sub IsiSemuaJabatanTarif()
    Dim nama As Variant, ws As Worksheet, r As Long, c As Variant

    For Each nama In Array("U-jan", "U-feb", "U-mar", "U-apr", "U-mei", "U-jun", _
                           "U-jul", "U-ags", "U-sep", "U-okt", "U-nov", "U-des")
        Set ws = ThisWorkbook.Worksheets(nama)
        If BukaSheetAman(ws) Then
            Application.EnableEvents = False
            For r = 3 To 52
                For Each c In Array(10, 14, 18, 22, 26, 30, 34, 38, 42, 46)
                    IsiJabatanTarif ws, r, CLng(c)
                Next c
            Next r
            Application.EnableEvents = True
            KunciSheetAman ws
        End If
    Next nama
End Sub
Public Sub IsiMLLainnya(ws As Worksheet, ByVal r As Long)
    Dim t As String, bulanKey As String, nmT As Variant, sb As Worksheet
    Dim i As Long, s As Long, baris As Long, tahapSh As String, jenis As String
    Dim isEkstra As Boolean

    If Len(Trim$(CStr(ws.Cells(r, 6).text))) = 0 Then
        BersihkanRumusML ws, r
        Exit Sub
    End If

    t = NormKeyU(CStr(ws.Cells(r, 6).text))
    isEkstra = (InStr(LCase$(t), "upah ekstra") > 0)
    If Not isEkstra And IsKhususUraian(t) Then
        BersihkanRumusML ws, r
        Exit Sub
    End If

    bulanKey = BulanKeyU(ws.Name)
    baris = 0: tahapSh = "": jenis = ""

    For Each nmT In Array("TAHAP I", "TAHAP II")
        Set sb = ThisWorkbook.Worksheets(CStr(nmT))
        For i = 12 To 1254
            If i = 877 Then
                ' baris 877 hanya pelengkap, tidak pernah dipakai di sheet U
            ElseIf SetAwal(i) > 0 Then
                ' Blok set 527-876 dan 878-1227: hanya judul (s+1) dan Upah Pekerja (s+31)
                s = SetAwal(i)
                If i = s + 1 Or i = s + 31 Then
                    If UCase$(Trim$(CStr(sb.Cells(s, 15).text))) = bulanKey Then
                        If NormKeyU(CStr(sb.Cells(i, 1).text)) = t Then
                            baris = i: tahapSh = CStr(nmT)
                            If i = s + 1 Then jenis = "JUDUL" Else jenis = "UPAH"
                            Exit For
                        End If
                    End If
                End If
            Else
                ' Di luar blok set (baris 12-526 dan 1228-1254)
                If BulanBarisTahapU(sb, i) = bulanKey Then
                    If NormKeyU(CStr(sb.Cells(i, 3).text)) = t Then
                        baris = i: tahapSh = CStr(nmT)
                        If i >= 1230 And i <= 1254 Then jenis = "ASET" Else jenis = "ITEM"
                        Exit For
                    ElseIf NormKeyU(CStr(sb.Cells(i, 1).text)) = t Then
                        baris = i: tahapSh = CStr(nmT): jenis = "JUDUL"
                        Exit For
                    End If
                End If
            End If
        Next i
        If baris > 0 Then Exit For
    Next nmT
    If baris = 0 Then Exit Sub

    Set sb = ThisWorkbook.Worksheets(tahapSh)
    Select Case jenis
        Case "ASET"
            ' Aset: Jml Hari = Volume (J), Nominal = Harga (L)
            ws.Cells(r, 13).Formula = "='" & tahapSh & "'!$J$" & baris
            ws.Cells(r, 12).Formula = "='" & tahapSh & "'!$L$" & baris
        Case "JUDUL"
            ' Judul: Nominal = L, kalau kosong pakai M. Jml Hari = 1
            ws.Cells(r, 13).Formula = "=1"
            ws.Cells(r, 12).Formula = "=IF(ISNUMBER('" & tahapSh & "'!$L$" & baris & "),'" & tahapSh & "'!$L$" & baris & ",'" & tahapSh & "'!$M$" & baris & ")"
        Case "UPAH"
            ' Upah Pekerja: K dan L diisi dari nama tukang (IsiJabatanTarif). Jml Hari dari H.
            PasangMDariH ws, r, tahapSh, "UPAH", baris
        Case "ITEM"
            If isEkstra Then
                ' Ekstra: L (tarif) diisi dari dbekstra oleh IsiJabatanTarif. Jml Hari dari H.
                PasangMDariH ws, r, tahapSh, "EKSTRA", baris
            Else
                ws.Cells(r, 13).Formula = "=1"
                ws.Cells(r, 12).Formula = "='" & tahapSh & "'!$M$" & baris
            End If
    End Select
End Sub
' Isi kolom jumlah hari (M, Q, U, ..., AW) dari kolom H TAHAP, untuk semua 10 blok penerima.
' EKSTRA : H baris uraian (per orang).
' UPAH    : Tukang = H baris (s+32), Kuli = H baris (s+33), dipilih dari jabatan di K.
Private Sub PasangMDariH(ws As Worksheet, ByVal r As Long, ByVal tahapSh As String, _
                         ByVal jenis As String, ByVal baris As Long)
    Dim k As Long, nmA As String, jbA As String, shRef As String, rumus As String

    shRef = "'" & tahapSh & "'!"
    For k = 0 To 9
        nmA = ws.Cells(r, 10 + 4 * k).Address(False, False)   ' nama penerima
        jbA = ws.Cells(r, 11 + 4 * k).Address(False, False)   ' jabatan penerima

        If jenis = "EKSTRA" Then
            rumus = "=IF(" & nmA & "="""",""""," & shRef & "$H$" & baris & ")"
        Else
            rumus = "=IF(" & nmA & "="""","""",IF(" & jbA & "=""Tukang""," & _
                    shRef & "$H$" & (baris + 1) & ",IF(" & jbA & "=""Kuli""," & _
                    shRef & "$H$" & (baris + 2) & ","""")))"
        End If

        ws.Cells(r, 13 + 4 * k).Formula = rumus
    Next k
End Sub
' Label bulan baris TAHAP r: O di baris itu; kalau kosong, cari ke bawah (detail) sampai judul berikutnya;
' kalau masih kosong, cari ke atas (header set, mis. O527).
Private Function BulanBarisTahapU(sb As Worksheet, ByVal r As Long) As String
    Dim k As Long, b As String
    b = UCase$(Trim$(CStr(sb.Cells(r, 15).text)))
    If b = "" Then
        For k = 1 To 15
            If Trim$(CStr(sb.Cells(r + k, 1).text)) <> "" Then Exit For
            b = UCase$(Trim$(CStr(sb.Cells(r + k, 15).text)))
            If b <> "" Then Exit For
        Next k
    End If
    If b = "" Then
        For k = 1 To 2
            b = UCase$(Trim$(CStr(sb.Cells(r - k, 15).text)))
            If b <> "" Then Exit For
        Next k
    End If
    BulanBarisTahapU = b
End Function
' Baris awal set untuk baris TAHAP r (0 kalau bukan blok set).
' Blok 527-876: set mulai 527, 562, ..., 842. Blok 878-1227: set mulai 878, 913, ..., 1193.
Private Function SetAwal(ByVal r As Long) As Long
    If r >= 527 And r <= 876 Then
        SetAwal = 527 + 35 * ((r - 527) \ 35)
    ElseIf r >= 878 And r <= 1227 Then
        SetAwal = 878 + 35 * ((r - 878) \ 35)
    Else
        SetAwal = 0
    End If
End Function

' Nilai angka sel (teks atau kosong dianggap 0)
Private Function NilaiAngka(sb As Worksheet, ByVal r As Long, ByVal kol As Long) As Double
    Dim v As Variant
    v = sb.Cells(r, kol).Value
    If Not IsEmpty(v) And IsNumeric(v) Then NilaiAngka = CDbl(v) Else NilaiAngka = 0
End Function
Private Function BulanKeyU(ByVal namaSheet As String) As String
    Dim kode As String
    kode = LCase$(Mid$(namaSheet, 3))      ' "U-okt" -> "okt"
    Select Case kode
        Case "jan": BulanKeyU = "TUNAI JANUARI"
        Case "feb": BulanKeyU = "TUNAI FEBRUARI"
        Case "mar": BulanKeyU = "TUNAI MARET"
        Case "apr": BulanKeyU = "TUNAI APRIL"
        Case "mei": BulanKeyU = "TUNAI MEI"
        Case "jun": BulanKeyU = "TUNAI JUNI"
        Case "jul": BulanKeyU = "TUNAI JULI"
        Case "ags": BulanKeyU = "TUNAI AGUSTUS"
        Case "sep": BulanKeyU = "TUNAI SEPTEMBER"
        Case "okt": BulanKeyU = "TUNAI OKTOBER"
        Case "nov": BulanKeyU = "TUNAI NOVEMBER"
        Case "des": BulanKeyU = "TUNAI DESEMBER"
        Case Else: BulanKeyU = ""
    End Select
End Function
Private Sub BersihkanRumusML(ws As Worksheet, ByVal r As Long)
    ' Hanya hapus rumus, isian ketikan tidak disentuh
    If ws.Cells(r, 12).HasFormula Then ws.Cells(r, 12).ClearContents
    If ws.Cells(r, 13).HasFormula Then ws.Cells(r, 13).ClearContents
End Sub

Private Function IsKhususUraian(ByVal t As String) As Boolean
    Dim k As Variant
    t = LCase$(t)
    For Each k In Array("honor", "upah ekstra", "upah pekerja")
        If InStr(t, k) > 0 Then
            IsKhususUraian = True
            Exit Function
        End If
    Next k
End Function

Private Function IsBarisPekerjaTahap(sb As Worksheet, ByVal baris As Long) As Boolean
    Dim d As String
    d = UCase$(Trim$(CStr(sb.Cells(baris, 4).Value)))
    IsBarisPekerjaTahap = (Left$(d, 5) = "UPAH " And (InStr(d, "TUKANG") > 0 Or InStr(d, "KULI") > 0))
End Function
Public Sub ProsesUbahUraian(ByVal ws As Worksheet, ByVal Target As Range)
    Dim c As Range, f As String
    Dim adaKolomF As Boolean
    Dim evAsal As Boolean, layarAsal As Boolean, kalkAsal As XlCalculation

    If Target.Row < 3 Or Target.Row > 52 Then Exit Sub
    If Target.Column <> 6 And Not IsPenerimaCol(Target.Column) Then Exit Sub

    adaKolomF = (Target.Column = 6)

    ' Simpan status apa adanya, lalu kembalikan persis seperti semula
    evAsal = Application.EnableEvents
    layarAsal = Application.ScreenUpdating
    kalkAsal = Application.Calculation

    On Error GoTo Gagal
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If adaKolomF Then
        For Each c In Target.Cells
            If Not c.HasFormula And Trim$(CStr(c.Value)) <> "" Then
                f = CariMaster(Trim$(CStr(c.Value)))
                If f <> "" Then c.Formula = f
            End If
        Next c
    End If

    If BukaSheetAman(ws) Then
        If adaKolomF Then
            For Each c In Target.Cells
                                If c.Row >= 3 And c.Row <= 52 Then
                    PasangDropdownBaris ws, c.Row
                    IsiMLLainnya ws, c.Row
                End If
            Next c
        ElseIf IsPenerimaCol(Target.Column) Then
            For Each c In Target.Cells
                If c.Row >= 3 And c.Row <= 52 Then IsiJabatanTarif ws, c.Row, c.Column
            Next c
        End If
        KunciSheetAman ws
    End If

    If adaKolomF Then SiapkanUraianSatuSheet ws
    PasangDropdownUpahU ws

Selesai:
    On Error Resume Next
    Application.Calculation = kalkAsal
    Application.EnableEvents = evAsal
    Application.ScreenUpdating = layarAsal
    On Error GoTo 0
    Exit Sub

Gagal:
    MsgBox "Gagal memperbarui uraian: " & Err.Description & " (" & Err.Number & ")", vbExclamation
    Resume Selesai
End Sub

Private Function NormKeyU(ByVal s As String) As String
    Dim t As String
    t = Replace(s, Chr(160), " ")
    t = UCase$(Trim$(t))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormKeyU = t
End Function
