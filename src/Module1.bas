Attribute VB_Name = "Module1"
Sub ExportPDF()
    Dim ws As Worksheet
    Dim tempWB As Workbook
    Dim tempWS1 As Worksheet, tempWS2 As Worksheet
    Dim reportPath As String, filePath As String
    Dim i As Long, lastRow As Long, startRow As Long, lastPasteRow As Long
    Dim brandName As String
    Dim userInput As String, cutOffText As String
    Dim table1HeaderRow As Long, table2HeaderRow As Long
    Dim dataRange As Range
    Dim r As Long
    
    ' --- Input user ---
    userInput = InputBox("Jenis Media Edukasi:", "PT Toyota Astra Motor")
    If userInput = "" Then Exit Sub
    cutOffText = InputBox("Input Cut off (Ex: Cut Off W3 Oct 2025):", "PT Toyota Astra Motor")
    If cutOffText = "" Then Exit Sub
    
    ' --- Setup awal ---
    Set ws = ThisWorkbook.Sheets("export")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' --- Buat folder "Report" kalau belum ada ---
    reportPath = ThisWorkbook.path & "\Report"
    If Dir(reportPath, vbDirectory) = "" Then MkDir reportPath
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Loop tiap grup/brand ---
    For i = 3 To lastRow
        DoEvents ' biar bisa stop paksa
        If ws.Cells(i, "A").Font.Bold And Trim(ws.Cells(i, "A").value) <> "" Then
            brandName = Trim(ws.Cells(i, "A").value)
            ' Ganti karakter ilegal
            brandName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(brandName, "/", "_"), "\", "_"), ":", "_"), "*", "_"), "?", "_"), """", "_"), "<", "_"), ">", "_"), "|", "_")
            filePath = reportPath & "\" & brandName & ".pdf"
            
            ' --- Buat workbook sementara ---
            Set tempWB = Workbooks.Add(xlWBATWorksheet)
            Set tempWS1 = tempWB.Sheets(1)
            tempWS1.Name = "Customer Survey Rate"
            Set tempWS2 = tempWB.Sheets.Add(After:=tempWS1)
            tempWS2.Name = "Monthly Survey Achievement"
            
            ' --- HALAMAN 1: Customer Survey Rate ---
            tempWS1.Cells.Clear
            tempWS1.Columns.ColumnWidth = 8.43
            Dim pasteRow As Long: pasteRow = 1
            
            ' Judul
            With tempWS1.Cells(pasteRow, "A")
                .value = "CUSTOMER SURVEY RATE" & vbCrLf & _
                         "MEDUK " & userInput & vbCrLf & _
                         "DEALER " & brandName
                With tempWS1.Range(tempWS1.Cells(pasteRow, "A"), tempWS1.Cells(pasteRow + 2, "D"))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.Bold = True
                    .Font.Size = 10
                    .WrapText = True
                End With
            End With
            tempWS1.Rows(pasteRow).AutoFit
            
            ' Cut off
            With tempWS1.Cells(4, "D")
                .value = cutOffText
                .Font.Italic = True
                .Font.Color = RGB(150, 150, 150)
                .HorizontalAlignment = xlRight
            End With
            
            ' Header tabel
            pasteRow = 5
            table1HeaderRow = pasteRow
            tempWS1.Cells(pasteRow, "A").value = "DEALER"
            tempWS1.Cells(pasteRow, "B").value = "NUMBER OF OUTLET"
            tempWS1.Cells(pasteRow, "C").value = "NUMBER OF OUTLET ALREADY EXECUTED SURVEY"
            tempWS1.Cells(pasteRow, "D").value = "SURVEY RATE"
            With tempWS1.Range(tempWS1.Cells(pasteRow, "A"), tempWS1.Cells(pasteRow, "D"))
                .Font.Bold = True
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Interior.Color = RGB(68, 114, 196)
                .Font.Color = vbWhite
                .HorizontalAlignment = xlCenter
            End With
            tempWS1.Rows(pasteRow).RowHeight = 30
            
            ' Data tabel
            pasteRow = pasteRow + 1
            startRow = i + 1
            Do While startRow <= lastRow _
                And Not ws.Cells(startRow, "A").Font.Bold _
                And Not IsEmpty(ws.Cells(startRow, "A"))
                startRow = startRow + 1
            Loop
            startRow = startRow - 1
            
            ws.Range(ws.Cells(i, "A"), ws.Cells(startRow, "A")).Copy Destination:=tempWS1.Cells(pasteRow, "A")
            ws.Range(ws.Cells(i, "K"), ws.Cells(startRow, "K")).Copy Destination:=tempWS1.Cells(pasteRow, "B")
            ws.Range(ws.Cells(i, "L"), ws.Cells(startRow, "L")).Copy Destination:=tempWS1.Cells(pasteRow, "C")
            ws.Range(ws.Cells(i, "M"), ws.Cells(startRow, "M")).Copy
            tempWS1.Cells(pasteRow, "D").PasteSpecial xlPasteValuesAndNumberFormats
            Application.CutCopyMode = False
            
            ' Grand Total
            lastPasteRow = tempWS1.Cells(tempWS1.Rows.Count, "A").End(xlUp).Row
            pasteRow = lastPasteRow + 1
            tempWS1.Cells(pasteRow, "A").value = "Grand Total"
            tempWS1.Cells(pasteRow, "A").Font.Bold = True
            tempWS1.Cells(pasteRow, "D").Formula = "=IFERROR(SUM(C" & table1HeaderRow + 1 & ":C" & lastPasteRow & ")/SUM(B" & table1HeaderRow + 1 & ":B" & lastPasteRow & "),0)"
            tempWS1.Cells(pasteRow, "D").NumberFormat = "0%"
            tempWS1.Cells(pasteRow, "D").Font.Bold = True
            
            ' Border & AutoFit
            Set dataRange = tempWS1.Range(tempWS1.Cells(table1HeaderRow, "A"), tempWS1.Cells(pasteRow, "D"))
            With dataRange.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            tempWS1.Columns("A:D").AutoFit
            For r = table1HeaderRow To pasteRow
                tempWS1.Rows(r).AutoFit
            Next r
            
            ' Page setup halaman 1 (Customer Survey Rate)
            With tempWS1.PageSetup
                .Orientation = xlPortrait
                .Zoom = 100
                .CenterHorizontally = True
                .TopMargin = Application.InchesToPoints(0.3)
                .BottomMargin = Application.InchesToPoints(0.3)
                .LeftMargin = Application.InchesToPoints(0.3)
                .RightMargin = Application.InchesToPoints(0.3)
            End With
            
            ' --- HALAMAN 2: Monthly Survey Achievement ---
            tempWS2.Cells.Clear
            tempWS2.Columns.ColumnWidth = 8.43
            pasteRow = 1
            
            ' Judul
            With tempWS2.Cells(pasteRow, "A")
                .value = "MONTHLY SURVEY ACHIEVEMENT" & vbCrLf & _
                         "MEDUK " & userInput & vbCrLf & _
                         "DEALER " & brandName
                With tempWS2.Range(tempWS2.Cells(pasteRow, "A"), tempWS2.Cells(pasteRow + 2, "K"))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.Bold = True
                    .Font.Size = 10
                    .WrapText = True
                End With
            End With
            tempWS2.Rows(pasteRow).AutoFit
            
            ' Header tabel
            pasteRow = pasteRow + 4
            table2HeaderRow = pasteRow
            ws.Range("A1:J2").Copy
            tempWS2.Cells(pasteRow, "A").PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            
            ' Data tabel
            pasteRow = pasteRow + 2
            ws.Range(ws.Cells(i, "A"), ws.Cells(startRow, "J")).Copy
            tempWS2.Cells(pasteRow, "A").PasteSpecial xlPasteAll
            
            ' Border & AutoFit
            lastPasteRow = tempWS2.Cells(tempWS2.Rows.Count, "A").End(xlUp).Row
            Set dataRange = tempWS2.Range(tempWS2.Cells(table2HeaderRow, "A"), tempWS2.Cells(lastPasteRow, "J"))
            With dataRange.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            tempWS2.Columns("A:J").AutoFit
            For r = table2HeaderRow To lastPasteRow
                tempWS2.Rows(r).AutoFit
            Next r
            
            ' Page setup
            With tempWS2.PageSetup
                .Orientation = xlPortrait
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .CenterHorizontally = True
                .TopMargin = Application.InchesToPoints(0.5)
                .BottomMargin = Application.InchesToPoints(0.5)
                .LeftMargin = Application.InchesToPoints(0.4)
                .RightMargin = Application.InchesToPoints(0.4)
            End With
            
            ' --- Export kedua sheet menjadi 1 PDF ---
            Dim wsArray As Variant
            wsArray = Array(tempWS1.Name, tempWS2.Name) ' urut: 1 = Customer, 2 = Monthly
            tempWB.Sheets(wsArray).Select
            tempWB.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=filePath, Quality:=xlQualityStandard, _
                IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            ' Tutup workbook sementara tanpa konfirmasi
            Application.DisplayAlerts = False
            tempWB.Close SaveChanges:=False
            Application.DisplayAlerts = True
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "? Semua branch berhasil diekspor ke folder:" & vbCrLf & reportPath, vbInformation, "Export Complete"
End Sub


Sub AddData()
    Dim wsDest As Worksheet
    Dim fDialog As FileDialog
    Dim filePath As String
    Dim lastRowDest As Long
    Dim destData As Variant, sourceData As Variant
    Dim dict As Object
    Dim i As Long, n As Long
    Dim key As String
    Dim duplicates As Collection
    Dim response As VbMsgBoxResult
    Dim replaceAll As Boolean
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsDest = ThisWorkbook.Sheets("Data")
    Set dict = CreateObject("Scripting.Dictionary")
    Set duplicates = New Collection
    
    ' Load existing destination data into dictionary
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    If lastRowDest < 2 Then
        destData = wsDest.Range("A1").CurrentRegion.value
    Else
        destData = wsDest.Range("A1:A" & lastRowDest).Resize(lastRowDest, wsDest.Cells(1, wsDest.Columns.Count).End(xlToLeft).Column).value
    End If
    
    ' Build dictionary for duplicates
    For i = 2 To UBound(destData, 1)
        key = destData(i, 1) & "|" & destData(i, 2) & "|" & destData(i, 3)
        dict(key) = i
    Next i
    
    ' Open file dialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select Excel or CSV file to import"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        .Filters.Add "CSV Files", "*.csv"
        If .Show = -1 Then
            filePath = .selectedItems(1)
        Else
            MsgBox "No file selected.", vbExclamation
            GoTo Cleanup
        End If
    End With
    
    ' Load source data into array
    If LCase(Right(filePath, 3)) = "csv" Then
        ' Open CSV temporarily
        Workbooks.OpenText Filename:=filePath, DataType:=xlDelimited, Comma:=True
        sourceData = ActiveSheet.UsedRange.value
        ActiveWorkbook.Close False
    Else
        ' Open Excel file temporarily
        Dim wbSource As Workbook
        Set wbSource = Workbooks.Open(filePath)
        sourceData = wbSource.Sheets(1).UsedRange.value
        wbSource.Close False
    End If
    
    ' Collect duplicates
    For i = 2 To UBound(sourceData, 1)
        key = sourceData(i, 1) & "|" & sourceData(i, 2) & "|" & sourceData(i, 3)
        If dict.Exists(key) Then duplicates.Add Array(key, i)
    Next i
    
    ' Ask user once if duplicates exist
    replaceAll = False
    If duplicates.Count > 0 Then
        response = MsgBox(duplicates.Count & " duplicate(s) found. Replace all existing duplicates?", vbYesNo + vbQuestion, "Duplicate Found")
        If response = vbYes Then replaceAll = True
    End If
    
    ' Create a new array to merge data
    Dim merged() As Variant
    Dim totalRows As Long, totalCols As Long
    totalRows = UBound(destData, 1) - 1 + UBound(sourceData, 1) - 1
    totalCols = UBound(destData, 2)
    ReDim merged(1 To totalRows, 1 To totalCols)
    
    ' Copy existing destination data
    For i = 2 To UBound(destData, 1)
        For n = 1 To totalCols
            merged(i - 1, n) = destData(i, n)
        Next n
    Next i
    
    ' Copy source data (check duplicates)
    Dim currentRow As Long
    currentRow = UBound(destData, 1) - 1
    For i = 2 To UBound(sourceData, 1)
        key = sourceData(i, 1) & "|" & sourceData(i, 2) & "|" & sourceData(i, 3)
        If dict.Exists(key) Then
            If replaceAll Then
                ' Replace in merged array
                Dim destRow As Long
                destRow = dict(key) - 1
                For n = 1 To totalCols
                    merged(destRow, n) = sourceData(i, n)
                Next n
            End If
        Else
            currentRow = currentRow + 1
            For n = 1 To totalCols
                merged(currentRow, n) = sourceData(i, n)
            Next n
        End If
    Next i
    
    ' Write merged data back to sheet (starting row 2)
    wsDest.Range("A2").Resize(currentRow, totalCols).value = merged
    
    ' Sort by Timestamp (column A)
    wsDest.Sort.SortFields.Clear
    wsDest.Sort.SortFields.Add key:=wsDest.Range("A2:A" & wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsDest.Sort
        .SetRange wsDest.Range("A1").CurrentRegion
        .Header = xlYes
        .Apply
    End With
    
    MsgBox "Data import complete and sorted by Timestamp.", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Sub DeleteDataMenu()
    Dim choice As Variant
    choice = Application.InputBox("Choose delete method:" & vbCrLf & _
                                  "1 = Timestamp Range" & vbCrLf & _
                                  "2 = Media Edukasi" & vbCrLf & _
                                  "3 = Keyword Search" & vbCrLf & _
                                  "4 = Manual Selection", Type:=1)
    
    Select Case choice
        Case 1
            DeleteByTimestamp
        Case 2
            DeleteByMediaEdukasi
        Case 3
            DeleteByKeyword
        Case 4
            ManualDelete
        Case Else
            MsgBox "Invalid choice", vbExclamation
    End Select
End Sub
Sub DeleteByTimestamp()
    Dim ws As Worksheet
    Dim startDate As Date, endDate As Date
    Dim lastRow As Long, i As Long
    
    Set ws = ThisWorkbook.Sheets("Data")
    
    startDate = Application.InputBox("Enter start date (yyyy-mm-dd):", Type:=2)
    endDate = Application.InputBox("Enter end date (yyyy-mm-dd):", Type:=2)
    
    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        MsgBox "Invalid date.", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "A").value >= startDate And ws.Cells(i, "A").value <= endDate Then
            ws.Rows(i).Delete
        End If
    Next i
    
    MsgBox "Data deleted based on timestamp range.", vbInformation
End Sub

Sub ManualDelete()
    MsgBox "Go to the 'Data' sheet, select rows you want to delete, then run 'DeleteSelectedRows' macro.", vbInformation
End Sub

Sub DeleteSelectedRows()
    Dim rng As Range
    On Error Resume Next
    Set rng = Selection.EntireRow
    If rng Is Nothing Then
        MsgBox "No rows selected", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete the selected rows?", vbYesNo + vbQuestion) = vbYes Then
        rng.Delete
        MsgBox "Selected rows deleted.", vbInformation
    End If
End Sub


Sub DeleteByMediaEdukasi()
    Dim colF As Long
    colF = 6 ' Kolom F
    frmDeleteDataFast.colFilter = colF
    frmDeleteDataFast.keyword = "" ' Semua media edukasi
    frmDeleteDataFast.Show
End Sub

Sub DeleteByKeyword()
    Dim kw As String
    kw = Application.InputBox("Enter keyword to search:", Type:=2)
    If kw = "" Then Exit Sub
    
    frmDeleteDataFast.colFilter = 6 ' Bisa ganti sesuai kolom yang mau dicari
    frmDeleteDataFast.keyword = kw
    frmDeleteDataFast.Show
End Sub

Sub ClearAllData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' Cek apakah ada data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data to clear.", vbInformation
        Exit Sub
    End If
    
    ' Tampilkan konfirmasi
    If MsgBox("Are you sure you want to delete ALL data? This action cannot be undone.", _
              vbYesNo + vbExclamation, "Confirm Clear Data") = vbYes Then
        ' Hapus semua baris kecuali header (baris 1)
        ws.Rows("2:" & ws.Rows.Count).ClearContents
        MsgBox "All data cleared.", vbInformation
    End If
End Sub


