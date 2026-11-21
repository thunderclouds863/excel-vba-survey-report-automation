Attribute VB_Name = "Module1"
Sub ExportPDF_Dinamis_SatuHalaman()
    Dim ws As Worksheet
    Dim tempWB As Workbook
    Dim tempWs As Worksheet
    Dim reportPath As String, filePath As String
    Dim i As Long, lastRow As Long, startRow As Long, lastPasteRow As Long
    Dim brandName As String
    Dim userInput As String, cutOffText As String
    Dim tableHeaderRow As Long
    Dim dataRange As Range, cfDataRange As Range
    Dim r As Long
    
    Dim colOutlet As Variant, colExecuted As Variant, colRate As Variant
    Dim colLastDynamic As Long
    
    ' --- Input user ---
    userInput = InputBox("Jenis Media Edukasi:", "PT Toyota Astra Motor")
    If userInput = "" Then Exit Sub
    cutOffText = InputBox("Input Cut off (Ex: Cut Off W3 Oct 2025):", "PT Toyota Astra Motor")
    If cutOffText = "" Then Exit Sub
    
    ' --- Setup awal ---
    Set ws = ThisWorkbook.Sheets("CUSTOMER SURVEY RATE")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    On Error Resume Next
    colOutlet = Application.Match("NUMBER OF OUTLET", ws.Rows(1), 0)
    colExecuted = Application.Match("NUMBER OF OUTLET ALREADY EXECUTED SURVEY", ws.Rows(1), 0)
    colRate = Application.Match("SURVEY RATE", ws.Rows(1), 0)
    On Error GoTo 0
    
    If IsError(colOutlet) Or IsError(colExecuted) Or IsError(colRate) Then
        MsgBox "Kolom 'NUMBER OF OUTLET', '...EXECUTED SURVEY', atau 'SURVEY RATE' tidak ditemukan di Baris 1.", vbCritical
        Exit Sub
    End If
    
    colLastDynamic = colOutlet - 1

    ' --- Buat Folder ---
    reportPath = ThisWorkbook.Path & "\Report" & brandName & "_Meduk_" & userInput
    If Dir(reportPath, vbDirectory) = "" Then MkDir reportPath
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = 3 To lastRow - 1
        DoEvents
        If ws.Cells(i, "A").Font.Bold And Trim(ws.Cells(i, "A").Value) <> "" Then
            brandName = Trim(ws.Cells(i, "A").Value)
            brandName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(brandName, "/", "_"), "\", "_"), ":", "_"), "*", "_"), "?", "_"), """", "_"), "<", "_"), ">", "_"), "|", "_")
            filePath = reportPath & "\" & brandName & "_Meduk_" & userInput & ".pdf"
            
            startRow = i + 1
            Do While startRow <= lastRow _
                And Not ws.Cells(startRow, "A").Font.Bold _
                And Not IsEmpty(ws.Cells(startRow, "A"))
                startRow = startRow + 1
            Loop
            startRow = startRow - 1
            
            Set tempWB = Workbooks.Add(xlWBATWorksheet)
            Set tempWs = tempWB.Sheets(1)
            tempWs.Name = "Monthly Survey Achievement"
            
            tempWs.Cells.Clear
            Dim pasteRow As Long: pasteRow = 1
            
            ' Judul Utama
            With tempWs.Cells(pasteRow, "A")
                .Value = "MONTHLY SURVEY ACHIEVEMENT" & vbCrLf & _
                         "MEDUK " & userInput & vbCrLf & _
                         "DEALER " & brandName
                With tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(pasteRow + 2, colLastDynamic))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.Bold = True
                    .Font.Size = 14
                    .WrapText = True
                End With
            End With
            tempWs.Rows(pasteRow & ":" & pasteRow + 2).AutoFit
            tempWs.Rows(pasteRow & ":" & pasteRow + 2).RowHeight = tempWs.Rows(pasteRow).RowHeight + 5

            With tempWs.Cells(4, "A")
                .Value = "Target survei per bulan = 5 pelanggan"
                .Font.Bold = True
                .Font.Color = RGB(0, 0, 0)
                .HorizontalAlignment = xlLeft
            End With
            
            With tempWs.Range(tempWs.Cells(6, "A"), tempWs.Cells(6, colLastDynamic))
                .Merge
                .Value = cutOffText
                .Font.Italic = True
                .Font.Color = RGB(150, 150, 150)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            With tempWs.Cells(5, "A")
                .Value = ChrW(&H25A0) & "  Belum memenuhi target survei"
                .Font.Bold = True
                .Font.Color = vbBlack
                .HorizontalAlignment = xlLeft
                .Characters(Start:=1, Length:=1).Font.Color = vbYellow
            End With
            
            pasteRow = 7
            tableHeaderRow = pasteRow
            
            ' --- COPY HEADER (Hybrid: Format + Values) ---
            ws.Range(ws.Cells(1, "A"), ws.Cells(2, colLastDynamic)).Copy
            
            With tempWs.Cells(pasteRow, "A")
                .PasteSpecial xlPasteColumnWidths ' Copy Lebar Kolom
                .PasteSpecial xlPasteFormats      ' Copy Format Cantik (Warna, Bold, dll)
                .PasteSpecial xlPasteValues       ' Copy Isi Data
            End With
            
            ' [KRUSIAL] Langsung Hapus Conditional Formatting Bawaan Header
            ' Ini agar coretan (strikethrough) yang menempel di rule lama hilang
            With tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(pasteRow + 1, colLastDynamic))
                .FormatConditions.Delete
                .Font.Strikethrough = False
            End With
            
            Application.CutCopyMode = False
            pasteRow = pasteRow + 2

            ' --- COPY DATA (Hybrid: Format + Values) ---
            ws.Range(ws.Cells(i, "A"), ws.Cells(startRow, colLastDynamic)).Copy
            
            With tempWs.Cells(pasteRow, "A")
                .PasteSpecial xlPasteFormats      ' Copy Format (Border, Font, dll)
                .PasteSpecial xlPasteValues       ' Copy Isi Data
            End With
            
            lastPasteRow = tempWs.Cells(tempWs.Rows.Count, "A").End(xlUp).Row
            
            ' [KRUSIAL] Langsung Hapus Conditional Formatting Bawaan Data
            If lastPasteRow >= pasteRow Then
                With tempWs.Range(tempWs.Cells(pasteRow, "A"), tempWs.Cells(lastPasteRow, colLastDynamic))
                    .FormatConditions.Delete     ' Hapus biang kerok coretan
                    .Font.Strikethrough = False  ' Pastikan font bersih
                End With
            End If
            
            Application.CutCopyMode = False

            ' --- PASANG CF BARU (Hanya Warna Kuning) ---
            If lastPasteRow >= (pasteRow + 1) And colLastDynamic >= 2 Then
            
                Set cfDataRange = tempWs.Range(tempWs.Cells(pasteRow + 1, 2), tempWs.Cells(lastPasteRow, colLastDynamic - 1))
                
                With cfDataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=5)
                    .Interior.Color = RGB(255, 255, 0)
                End With
                
                Dim firstCellAddr As String
                firstCellAddr = cfDataRange.Cells(1, 1).Address(False, False)
                
                With cfDataRange.FormatConditions.Add( _
                    Type:=xlExpression, _
                    Formula1:="=OR(TRIM(" & firstCellAddr & ")=""-"";" & _
                             "TRIM(" & firstCellAddr & ")="""";" & _
                             "ISBLANK(" & firstCellAddr & "))")
                    .Interior.Color = RGB(255, 255, 0)
                End With
                            
            End If

            ' Rapikan Border Ulang (Jaga-jaga kalau format border asli ada yang putus)
            Set dataRange = tempWs.Range(tempWs.Cells(tableHeaderRow, "A"), tempWs.Cells(lastPasteRow, colLastDynamic))
            With dataRange.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            
            tempWs.Columns("A").AutoFit
            If colLastDynamic > 2 Then
                tempWs.Range(tempWs.Cells(1, 2), tempWs.Cells(1, colLastDynamic - 1)).EntireColumn.ColumnWidth = 12
            End If
            tempWs.Columns(colLastDynamic).AutoFit
            
            For r = tableHeaderRow To lastPasteRow
                tempWs.Rows(r).AutoFit
            Next r
            
            With tempWs.PageSetup
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
            
            tempWs.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=filePath, Quality:=xlQualityStandard, _
                IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            Application.DisplayAlerts = False
            tempWB.Close SaveChanges:=False
            Application.DisplayAlerts = True
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Semua branch berhasil diekspor ke folder:" & vbCrLf & reportPath, vbInformation, "Export Complete"
End Sub

Sub AddData()
    Dim wsDest As Worksheet
    Dim fDialog As FileDialog
    Dim filePath As String
    Dim lastRowDest As Long, lastColDest As Long
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
    
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    lastColDest = wsDest.Cells(1, wsDest.Columns.Count).End(xlToLeft).Column
    
    If lastRowDest < 2 Then
        ReDim destData(1 To 1, 1 To lastColDest)
    Else
        destData = wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(lastRowDest, lastColDest)).Value
    End If
    
    If lastRowDest >= 2 Then
        For i = 1 To UBound(destData, 1)
            key = destData(i, 1) & "|" & destData(i, 2) & "|" & destData(i, 3)
            dict(key) = i
        Next i
    End If
    
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
    
    If LCase(Right(filePath, 3)) = "csv" Then
        Workbooks.OpenText Filename:=filePath, DataType:=xlDelimited, Comma:=True
        sourceData = ActiveSheet.UsedRange.Value
        ActiveWorkbook.Close False
    Else
        Dim wbSource As Workbook
        Set wbSource = Workbooks.Open(filePath)
        sourceData = wbSource.Sheets(1).UsedRange.Value
        wbSource.Close False
    End If
    
    For i = 2 To UBound(sourceData, 1)
        key = sourceData(i, 1) & "|" & sourceData(i, 2) & "|" & sourceData(i, 3)
        If dict.exists(key) Then duplicates.Add Array(key, i)
    Next i
    
    replaceAll = False
    If duplicates.Count > 0 Then
        response = MsgBox(duplicates.Count & " duplicate(s) found. Replace all existing duplicates?", vbYesNo + vbQuestion, "Duplicate Found")
        If response = vbYes Then replaceAll = True
    End If
    
    Dim merged() As Variant
    Dim totalRows As Long, totalCols As Long
    Dim currentRow As Long
    
    totalCols = lastColDest
    totalRows = IIf(lastRowDest < 2, 0, UBound(destData, 1)) + UBound(sourceData, 1) - 1
    ReDim merged(1 To totalRows, 1 To totalCols)
    
    If lastRowDest >= 2 Then
        For i = 1 To UBound(destData, 1)
            For n = 1 To totalCols
                merged(i, n) = destData(i, n)
            Next n
        Next i
        currentRow = UBound(destData, 1)
    Else
        currentRow = 0
    End If
    
    For i = 2 To UBound(sourceData, 1)
        key = sourceData(i, 1) & "|" & sourceData(i, 2) & "|" & sourceData(i, 3)
        If dict.exists(key) Then
            If replaceAll Then
                Dim destRow As Long
                destRow = dict(key)
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
    
    wsDest.Range("A2").Resize(currentRow, totalCols).Value = merged
    
    wsDest.Sort.SortFields.Clear
    wsDest.Sort.SortFields.Add key:=wsDest.Range("A2:A" & wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsDest.Sort
        .SetRange wsDest.Range("A1").Resize(wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row, lastColDest)
        .Header = xlYes
        .Apply
    End With
    
    MsgBox "Data import complete and sorted by Timestamp.", vbInformation
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Sub ClearAllData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = ThisWorkbook.Sheets("Data")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data to clear.", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete ALL data? This action cannot be undone.", _
              vbYesNo + vbExclamation, "Confirm Clear Data") = vbYes Then
        ws.Rows("2:" & ws.Rows.Count).ClearContents
        MsgBox "All data cleared.", vbInformation
    End If
End Sub

Sub SinkronkanSemuaData()
    Dim wb As Workbook
    Dim wsPivot As Worksheet, wsSurvey As Worksheet
    Dim lastColPivot As Long, lastRowPivot As Long
    Dim lastRowSurvey As Long
    
    Dim sourceHeaderRange As Range
    Dim pivotDataRange As Range
    Dim pivotHeaderRowRange As Range
    
    Dim vlookupFormulaString As String, sumFormulaString As String
    
    Dim colGrandTotal As Long
    Dim colOutlet As Variant, colExecuted As Variant, colRate As Variant
    Dim oldDynamicColCount As Long, newDynamicColCount As Long
    Dim diff As Long
    
    Dim r As Long, c As Long, j As Long
    Dim k As Long, c_new As Long
    Dim vlookupValue As String, matchHeader As String
    Dim sumRange As String
    Dim cfRange As Range
    Dim firstCellAddress As String
    Dim cfFormula As String

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Set wb = ThisWorkbook
    Set wsPivot = wb.Sheets("MONTHLY SURVEY ACHIEVEMENT")
    Set wsSurvey = wb.Sheets("CUSTOMER SURVEY RATE")
    On Error GoTo 0
    
    If wsPivot Is Nothing Or wsSurvey Is Nothing Then
        MsgBox "Pastikan Anda memiliki sheet 'MONTHLY SURVEY ACHIEVEMENT' dan 'CUSTOMER SURVEY RATE'.", vbCritical
        Exit Sub
    End If

    With wsPivot
        lastColPivot = .Cells(4, .Columns.Count).End(xlToLeft).Column
        lastRowPivot = .Cells(.Rows.Count, "A").End(xlUp).Row
        Set sourceHeaderRange = .Range(.Cells(4, 2), .Cells(4, lastColPivot))
        Set pivotDataRange = .Range(.Cells(7, 1), .Cells(lastRowPivot, lastColPivot))
        Set pivotHeaderRowRange = .Range(.Cells(4, 1), .Cells(4, lastColPivot))
        colGrandTotal = lastColPivot
        newDynamicColCount = sourceHeaderRange.Columns.Count
    End With

    With wsSurvey
        lastRowSurvey = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRowSurvey < 3 Then lastRowSurvey = 3
        
        On Error Resume Next
        colOutlet = Application.Match("NUMBER OF OUTLET", .Rows(1), 0)
        colExecuted = Application.Match("NUMBER OF OUTLET ALREADY EXECUTED SURVEY", .Rows(1), 0)
        colRate = Application.Match("SURVEY RATE", .Rows(1), 0)
        On Error GoTo 0
        
        If IsError(colOutlet) Or IsError(colExecuted) Or IsError(colRate) Then
            MsgBox "Kolom 'NUMBER OF OUTLET', '...EXECUTED SURVEY', atau 'SURVEY RATE' tidak ditemukan di Baris 1.", vbCritical
            Exit Sub
        End If

        oldDynamicColCount = colOutlet - 2
        diff = newDynamicColCount - oldDynamicColCount
        
        If diff > 0 Then
            .Range(.Cells(1, colOutlet), .Cells(1, colOutlet + diff - 1)).EntireColumn.Insert
            
            Dim SourceFormatCol As Range
            Set SourceFormatCol = .Range(.Cells(1, colOutlet - 1), .Cells(2, colOutlet - 1))
            
            SourceFormatCol.EntireColumn.Copy
            .Range(.Cells(1, colOutlet), .Cells(1, colOutlet + diff - 1)).EntireColumn.PasteSpecial xlPasteColumnWidths
            
            For c_new = colOutlet To colOutlet + diff - 1
                Dim DestFormatRange As Range
                Set DestFormatRange = .Range(.Cells(1, c_new), .Cells(2, c_new))
                
                SourceFormatCol.Copy
                DestFormatRange.PasteSpecial xlPasteFormats
            Next c_new
            
            Application.CutCopyMode = False
            
        ElseIf diff < 0 Then
            .Range(.Cells(1, colOutlet + diff), .Cells(1, colOutlet - 1)).EntireColumn.Delete
        End If
        
        .Range(.Cells(1, 2), .Cells(2, colOutlet + diff - 1)).ClearContents
        .Range(.Cells(3, 2), .Cells(lastRowSurvey, colOutlet + diff - 1)).ClearContents
        
        Dim hCell As Range, i As Long: i = 0
        For Each hCell In sourceHeaderRange.Cells
            .Range("B1").Offset(0, i).Value = hCell.Value
            i = i + 1
        Next hCell
        
        On Error Resume Next
        colOutlet = Application.Match("NUMBER OF OUTLET", .Rows(1), 0)
        colExecuted = Application.Match("NUMBER OF OUTLET ALREADY EXECUTED SURVEY", .Rows(1), 0)
        colRate = Application.Match("SURVEY RATE", .Rows(1), 0)
        On Error GoTo 0
        
        If IsError(colOutlet) Then
             MsgBox "Gagal menemukan 'NUMBER OF OUTLET' setelah digeser.", vbCritical
             Exit Sub
        End If
    End With

    With wsSurvey
        r = 3
        
        While r <= lastRowSurvey
            
            If .Cells(r, "A").Font.Bold = True Then
                
                startSumRow = r + 1
                k = startSumRow
                While k <= lastRowSurvey And .Cells(k, "A").Font.Bold = False
                    k = k + 1
                Wend
                endSumRow = k - 1
                
                If startSumRow <= endSumRow Then

                    For j = startSumRow To endSumRow
                        vlookupValue = .Cells(j, "A").Address(False, True)
                        For c = 2 To colGrandTotal
                            matchHeader = .Cells(1, c).Address(True, False)
                            vlookupFormulaString = "=IFERROR(VLOOKUP(" & vlookupValue & ", " & _
                                                 pivotDataRange.Address(External:=True) & ", " & _
                                                 "MATCH(" & matchHeader & ", " & pivotHeaderRowRange.Address(External:=True) & ", 0), " & _
                                                 "FALSE), ""-"")"
                            .Cells(j, c).Formula = vlookupFormulaString
                        Next c
                        
                        Set cfRange = .Range(.Cells(j, 2), .Cells(j, colGrandTotal))
                        cfRange.FormatConditions.Delete
                        
                        Set cfRange = .Range(.Cells(j, 2), .Cells(j, colGrandTotal - 1))
                        With cfRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=5)
                            .Interior.Color = RGB(255, 255, 0)
                        End With
                        
                        With cfRange.FormatConditions.Add(Type:=xlTextString, String:="-", TextOperator:=xlEqual)
                           .Interior.Color = RGB(255, 255, 0)
                        End With

                        .Cells(j, colExecuted).Formula = "=IF(" & .Cells(j, colGrandTotal).Address(False, False) & "=""-"", 0, 1)"
                        .Cells(j, colRate).Formula = "=IFERROR(" & .Cells(j, colExecuted).Address(False, False) & "/" & .Cells(j, colOutlet).Address(False, False) & ", 0)"
                        .Cells(j, colRate).NumberFormat = "0.0%"
                        .Range(.Cells(j, 2), .Cells(j, colRate)).Font.Bold = False
                    Next j
                    
                    For c = 2 To colGrandTotal - 1
                        sumRange = .Range(.Cells(startSumRow, c), .Cells(endSumRow, c)).Address(False, False)
                        .Cells(r, c).Formula = "=SUM(" & sumRange & ")"
                        .Cells(r, c).NumberFormat = "General;General;""-"""
                    Next c
                    
                    sumRange = .Range(.Cells(startSumRow, colExecuted), .Cells(endSumRow, colExecuted)).Address(False, False)
                    .Cells(r, colExecuted).Formula = "=SUM(" & sumRange & ")"
                    .Cells(r, colExecuted).NumberFormat = "General;General;""-"""
                    
                    .Cells(r, colRate).Formula = "=IFERROR(" & .Cells(r, colExecuted).Address(False, False) & "/" & .Cells(r, colOutlet).Address(False, False) & ", 0)"
                    .Cells(r, colRate).NumberFormat = "0.0%"
                    .Range(.Cells(r, 2), .Cells(r, colRate)).Font.Bold = True
                    

                ElseIf r = lastRowSurvey Then
                    
                    For c = colOutlet To colExecuted
                        sumFormulaString = ""
                        For j = 3 To r - 1
                            If .Cells(j, "A").Font.Bold = True Then
                                sumFormulaString = sumFormulaString & .Cells(j, c).Address(False, False) & ","
                            End If
                        Next j
                        
                        If sumFormulaString <> "" Then
                            sumFormulaString = "=SUM(" & Left(sumFormulaString, Len(sumFormulaString) - 1) & ")"
                            .Cells(r, c).Formula = sumFormulaString
                        Else
                            .Cells(r, c).Value = 0
                        End If
                        .Cells(r, c).NumberFormat = "General;General;""-"""
                    Next c
                    
                    .Cells(r, colRate).Formula = "=IFERROR(" & .Cells(r, colExecuted).Address(False, False) & "/" & .Cells(r, colOutlet).Address(False, False) & ", 0)"
                    .Cells(r, colRate).NumberFormat = "0.0%"
                    .Range(.Cells(r, 2), .Cells(r, colRate)).Font.Bold = True
                    
                End If
                
                r = k
            
            Else
                vlookupValue = .Cells(r, "A").Address(False, True)
                For c = 2 To colGrandTotal
                    matchHeader = .Cells(1, c).Address(True, False)
                    vlookupFormulaString = "=IFERROR(VLOOKUP(" & vlookupValue & ", " & _
                                         pivotDataRange.Address(External:=True) & ", " & _
                                         "MATCH(" & matchHeader & ", " & pivotHeaderRowRange.Address(External:=True) & ", 0), " & _
                                         "FALSE), ""-"")"
                    .Cells(r, c).Formula = vlookupFormulaString
                Next c
                
                        Set cfRange = .Range(.Cells(j, 2), .Cells(j, colGrandTotal))
                        cfRange.FormatConditions.Delete
                        
                        Set cfRange = .Range(.Cells(j, 2), .Cells(j, colGrandTotal - 1))
                        With cfRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=5)
                            .Interior.Color = RGB(255, 255, 0)
                        End With
                        
                        With cfRange.FormatConditions.Add(Type:=xlTextString, String:="-", TextOperator:=xlEqual)
                           .Interior.Color = RGB(255, 255, 0)
                        End With

                .Cells(r, colExecuted).Formula = "=IF(" & .Cells(r, colGrandTotal).Address(False, False) & ">0, 1, 0)"
                .Cells(r, colRate).Formula = "=IFERROR(" & .Cells(r, colExecuted).Address(False, False) & "/" & .Cells(r, colOutlet).Address(False, False) & ", 0)"
                .Cells(r, colRate).NumberFormat = "0.0%"
                .Range(.Cells(r, 2), .Cells(r, colRate)).Font.Bold = False
                
                r = r + 1
            End If
        Wend
    End With
    
    Application.Calculate
    Application.ScreenUpdating = True
    wsSurvey.Select
    MsgBox "Sinkronisasi dan pengolahan data selesai!"
End Sub

