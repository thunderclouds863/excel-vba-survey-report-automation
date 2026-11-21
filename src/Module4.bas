Attribute VB_Name = "Module4"
Sub ExportPDF_SatuFile_BoldSaja_HinggaBulanTerakhir()
    Dim ws As Worksheet
    Dim tempWB As Workbook
    Dim tempWs As Worksheet
    Dim reportPath As String, filePath As String
    Dim i As Long, lastRow As Long, j As Long
    Dim userInput As String, cutOffText As String
    Dim tableHeaderRow As Long
    Dim pasteRow As Long
    Dim lastCol As Long
    Dim dataStartRow As Long
    Dim colGrandTotal As Variant
    
    ' --- Input user ---
    userInput = InputBox("Jenis Media Edukasi:", "PT Toyota Astra Motor")
    If userInput = "" Then Exit Sub
    cutOffText = InputBox("Input Cut off (Ex: Cut Off W3 Oct 2025):", "PT Toyota Astra Motor")
    If cutOffText = "" Then Exit Sub
    
    ' --- Setup awal ---
    Set ws = ThisWorkbook.Sheets("SURVEY RATE")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    
    ' --- Buat Folder ---
    reportPath = ThisWorkbook.Path & "\Report_All Dealer"
    If Dir(reportPath, vbDirectory) = "" Then MkDir reportPath
    
    ' --- File PDF tunggal ---
    filePath = reportPath & "\" & "ALL_DEALERS_Meduk_" & userInput & ".pdf"
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Buat workbook temporary untuk semua dealer ---
    Set tempWB = Workbooks.Add(xlWBATWorksheet)
    Set tempWs = tempWB.Sheets(1)
    tempWs.Name = "Monthly Survey Achievement"
    
    tempWs.Cells.Clear
    pasteRow = 1
    
    ' --- Judul Utama untuk semua dealer ---
    With tempWs.Cells(pasteRow, "A")
        .Value = "MONTHLY SURVEY ACHIEVEMENT" & vbCrLf & _
                 "MEDUK " & userInput & vbCrLf & _
                 "SURVEY RATE - ALL DEALERS"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 14
        .WrapText = True
    End With
    ' Merge judul utama
    tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(pasteRow, lastCol)).Merge
    tempWs.Rows(pasteRow).RowHeight = 60
    
    pasteRow = pasteRow + 2
    
    ' --- Rumus Survey Rate ---
    tempWs.Cells(pasteRow, "A").Value = "Survey Rate = " & ChrW(931) & " Cabang yang sudah melakukan survey / " & ChrW(931) & " Cabang di dealer"
    tempWs.Cells(pasteRow, "A").Font.Bold = True
    tempWs.Cells(pasteRow, "A").Font.Color = RGB(0, 0, 0)
    tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(pasteRow, lastCol)).Merge
    
    pasteRow = pasteRow + 2
    
    ' --- Cut Off Text ---
    tempWs.Cells(pasteRow, "A").Value = cutOffText
    tempWs.Cells(pasteRow, "A").Font.Italic = True
    tempWs.Cells(pasteRow, "A").Font.Color = RGB(150, 150, 150)
    tempWs.Cells(pasteRow, "A").HorizontalAlignment = xlCenter
    tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(pasteRow, lastCol)).Merge
    
    pasteRow = pasteRow + 1 ' Spasi sebelum tabel
    tableHeaderRow = pasteRow
    
    ' --- Copy struktur tabel dari SURVEY RATE ---
    ' Copy baris 1 dari SURVEY RATE (Dealer, Number of Outlet, SURVEY RATE)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy
    tempWs.Cells(pasteRow, "A").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Copy baris 2 dari SURVEY RATE (header bulan)
    ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol)).Copy
    tempWs.Cells(pasteRow + 1, "A").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Format header sesuai dengan aslinya
    With tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(pasteRow + 1, lastCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' Warna background header
    tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(pasteRow + 1, lastCol)).Interior.Color = RGB(140, 181, 249)
    
    ' Merge sel header sesuai dengan SURVEY RATE asli
    ' Cek merge area di sheet asli
    If ws.Range("A1").MergeCells Then
        tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(pasteRow + 1, "A")).Merge
    End If
    
    If ws.Range("B1").MergeCells Then
        tempWs.Range(tempWs.Cells(pasteRow, "B"), tempWs.Cells(pasteRow + 1, "B")).Merge
    End If
    
    ' Merge untuk "SURVEY RATE" di baris 1
    If ws.Range("C1").MergeCells Then
        Dim surveyRateRange As Range
        Set surveyRateRange = ws.Range("C1").MergeArea
        tempWs.Range(tempWs.Cells(pasteRow, 3), tempWs.Cells(pasteRow, lastCol)).Merge
    End If
    
    pasteRow = pasteRow + 2 ' Pindah ke baris data
    dataStartRow = pasteRow
    
    ' --- Loop untuk semua baris bold (dealer) ---
    For i = 3 To lastRow
        If ws.Cells(i, "A").Font.Bold And Trim(ws.Cells(i, "A").Value) <> "" Then
            ' Copy seluruh baris dealer
            ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Copy
            tempWs.Cells(pasteRow, "A").PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            
            ' Format baris dealer
            With tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(pasteRow, lastCol))
                .Font.Bold = True
                .Interior.Color = RGB(240, 240, 240)
                
                ' Format persentase (kolom C dan seterusnya) tanpa desimal
                For j = 3 To lastCol
                    If tempWs.Cells(pasteRow, j).Value <> "" Then
                        ' Konversi ke persentase bulat
                        If IsNumeric(tempWs.Cells(pasteRow, j).Value) Then
                            tempWs.Cells(pasteRow, j).NumberFormat = "0%"
                            ' Bulatkan ke angka terdekat
                            Dim cellValue As Double
                            cellValue = tempWs.Cells(pasteRow, j).Value
                            tempWs.Cells(pasteRow, j).Value = Round(cellValue * 100, 0) / 100
                        ElseIf InStr(1, tempWs.Cells(pasteRow, j).Value, "%") > 0 Then
                            ' Jika sudah format persentase, ubah ke tanpa desimal
                            Dim percValue As String
                            percValue = tempWs.Cells(pasteRow, j).Value
                            ' Hapus % dan ubah ke angka
                            percValue = Replace(percValue, "%", "")
                            If IsNumeric(percValue) Then
                                tempWs.Cells(pasteRow, j).Value = Round(CDbl(percValue) / 100, 2)
                                tempWs.Cells(pasteRow, j).NumberFormat = "0%"
                            End If
                        End If
                    End If
                Next j
                
                ' Format Number of Outlet (kolom B)
                If IsNumeric(tempWs.Cells(pasteRow, "B").Value) Then
                    tempWs.Cells(pasteRow, "B").NumberFormat = "0"
                    tempWs.Cells(pasteRow, "B").HorizontalAlignment = xlCenter
                End If
                
                ' Format Dealer Name (kolom A)
                tempWs.Cells(pasteRow, "A").HorizontalAlignment = xlLeft
            End With
            
            pasteRow = pasteRow + 1
        End If
    Next i
    
    ' --- Terapkan borders untuk seluruh tabel ---
    Dim lastPasteRow As Long
    lastPasteRow = tempWs.Cells(tempWs.Rows.Count, "A").End(xlUp).Row
    
    If lastPasteRow >= tableHeaderRow Then
        Dim dataRange As Range
        Set dataRange = tempWs.Range(tempWs.Cells(tableHeaderRow, "A"), tempWs.Cells(lastPasteRow, lastCol))
        
        With dataRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    
    ' --- AutoFit columns sesuai dengan layout SURVEY RATE ---
    ' Kolom A (Dealer) - Lebarkan sedikit
    tempWs.Columns("A").ColumnWidth = 30
    
    ' Kolom B (Number of Outlet)
    tempWs.Columns("B").ColumnWidth = 15
    tempWs.Columns("B").HorizontalAlignment = xlCenter
    
    ' Kolom C dan seterusnya (bulan)
    If lastCol > 2 Then
        For j = 3 To lastCol
            tempWs.Columns(j).ColumnWidth = 10
            tempWs.Columns(j).HorizontalAlignment = xlCenter
        Next j
    End If
    
    ' AutoFit tinggi baris
    tempWs.Rows(tableHeaderRow).RowHeight = 25
    tempWs.Rows(tableHeaderRow + 1).RowHeight = 20
    
    For i = dataStartRow To lastPasteRow
        tempWs.Rows(i).RowHeight = 20
    Next i
    
    ' --- Format halaman untuk PDF ---
    With tempWs.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .CenterVertically = False
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintTitleRows = "$" & tableHeaderRow & ":$" & (tableHeaderRow + 1)
    End With
    
    ' --- Set Print Area ---
    tempWs.PageSetup.PrintArea = tempWs.Range(tempWs.Cells(1, "A"), tempWs.Cells(lastPasteRow, lastCol)).Address
    
    ' --- Export ke PDF tunggal ---
    On Error Resume Next
    Kill filePath ' Hapus file jika sudah ada
    On Error GoTo 0
    
    tempWs.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=filePath, Quality:=xlQualityStandard, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    ' --- Cleanup ---
    Application.DisplayAlerts = False
    tempWB.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Semua dealer berhasil diekspor ke 1 file PDF:" & vbCrLf & filePath, vbInformation, "Export Complete"
End Sub

