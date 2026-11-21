Attribute VB_Name = "Module5"
Sub OlahDataSurveyRate_Fixed()
    Dim wb As Workbook
    Dim wsSurvey As Worksheet, wsPivot As Worksheet, wsSource As Worksheet
    Dim lastRowSurvey As Long, lastRowPivot As Long, lastColPivot As Long
    Dim lastRowSource As Long
    Dim pivotDataRange As Range, pivotHeaderRowRange As Range
    Dim r As Long, c As Long
    Dim vlookupFormulaString As String, matchHeader As String, vlookupValue As String
    Dim lastColSurvey As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' --- Set worksheet ---
    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsSurvey = wb.Sheets("SURVEY RATE")
    On Error GoTo 0
    
    ' --- Create SURVEY RATE sheet if it doesn't exist ---
    If wsSurvey Is Nothing Then
        Set wsSurvey = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsSurvey.Name = "SURVEY RATE"
    End If
    
    Set wsPivot = wb.Sheets("MONTHLY SURVEY ACHIEVEMENT")
    Set wsSource = wb.Sheets("CUSTOMER SURVEY RATE") ' Sheet sumber data dealer-cabang
    
    ' --- Clear existing data in SURVEY RATE ---
    wsSurvey.Cells.Clear
    
    ' --- Ambil struktur dealer-cabang dari CUSTOMER SURVEY RATE ---
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row - 1
    
    ' --- Setup pivot data range ---
    With wsPivot
        lastRowPivot = .Cells(.Rows.Count, "A").End(xlUp).Row
        lastColPivot = .Cells(4, .Columns.Count).End(xlToLeft).Column - 1
        Set pivotDataRange = .Range(.Cells(7, 1), .Cells(lastRowPivot, lastColPivot))
        Set pivotHeaderRowRange = .Range(.Cells(4, 1), .Cells(4, lastColPivot))
    End With
    
    Dim headerCol As Long
    headerCol = 3 ' Mulai dari kolom C karena kolom B untuk Number of Outlet
    
    ' --- Tulis header ---
    ' Header untuk kolom A
    wsSurvey.Range(wsSurvey.Cells(1, 1), wsSurvey.Cells(2, 1)).Merge
    wsSurvey.Cells(1, 1).Value = "Dealer"
    wsSurvey.Cells(1, 1).Font.Bold = True
    wsSurvey.Cells(1, 1).Interior.Color = RGB(140, 181, 249)
    wsSurvey.Cells(1, 1).VerticalAlignment = xlCenter
    wsSurvey.Cells(1, 1).HorizontalAlignment = xlCenter
    
    ' Header untuk kolom B (Number of Outlet)
    wsSurvey.Range(wsSurvey.Cells(1, 2), wsSurvey.Cells(2, 2)).Merge
    wsSurvey.Cells(1, 2).Value = "Number of Outlet"
    wsSurvey.Cells(1, 2).Font.Bold = True
    wsSurvey.Cells(1, 2).Interior.Color = RGB(140, 181, 249)
    wsSurvey.Cells(1, 2).VerticalAlignment = xlCenter
    wsSurvey.Cells(1, 2).HorizontalAlignment = xlCenter
    
    ' Header untuk SURVEY RATE (mulai dari kolom C)
    wsSurvey.Range(wsSurvey.Cells(1, headerCol), wsSurvey.Cells(1, lastColPivot + 1)).Merge
    wsSurvey.Cells(1, headerCol).Value = "SURVEY RATE"
    wsSurvey.Cells(1, headerCol).Font.Bold = True
    wsSurvey.Cells(1, headerCol).Interior.Color = RGB(140, 181, 249)
    wsSurvey.Cells(1, headerCol).HorizontalAlignment = xlCenter
    wsSurvey.Cells(1, headerCol).VerticalAlignment = xlCenter
    
    ' Write month headers from pivot (mulai dari kolom C)
    For c = 2 To lastColPivot
        wsSurvey.Cells(2, headerCol).Value = wsPivot.Cells(4, c).Value
        wsSurvey.Cells(2, headerCol).Font.Bold = True
        wsSurvey.Cells(2, headerCol).Interior.Color = RGB(140, 181, 249)
        wsSurvey.Cells(2, headerCol).HorizontalAlignment = xlCenter
        headerCol = headerCol + 1
    Next c
    
    ' --- Salin struktur dealer-cabang dari CUSTOMER SURVEY RATE ---
    Dim writeRow As Long
    writeRow = 3
    
    For r = 3 To lastRowSource
        If Trim(wsSource.Cells(r, "A").Value) <> "" Then
            ' Salin nama dealer/cabang ke kolom A
            wsSurvey.Cells(writeRow, 1).Value = wsSource.Cells(r, "A").Value
            
            ' Salin Number of Outlet dari kolom J ke kolom B (hanya untuk dealer)
            If wsSource.Cells(r, "A").Font.Bold = True Then
                ' Ini adalah dealer, ambil dari kolom J
                If wsSource.Cells(r, "J").Value <> "" Then
                    wsSurvey.Cells(writeRow, 2).Value = wsSource.Cells(r, "J").Value
                Else
                    wsSurvey.Cells(writeRow, 2).Value = 0
                End If
            Else
                ' Ini adalah cabang, kosongkan atau isi dengan 0
                wsSurvey.Cells(writeRow, 2).Value = ""
            End If
            
            ' Tentukan apakah ini dealer (bold) atau cabang (tidak bold)
            If wsSource.Cells(r, "A").Font.Bold = True Then
                wsSurvey.Cells(writeRow, 1).Font.Bold = True
                wsSurvey.Cells(writeRow, 1).Interior.Color = RGB(240, 240, 240)
                wsSurvey.Cells(writeRow, 2).Font.Bold = True
                wsSurvey.Cells(writeRow, 2).Interior.Color = RGB(240, 240, 240)
            Else
                wsSurvey.Cells(writeRow, 1).Font.Bold = False
                wsSurvey.Cells(writeRow, 2).Font.Bold = False
            End If
            
            ' Format angka untuk kolom B
            wsSurvey.Cells(writeRow, 2).HorizontalAlignment = xlCenter
            
            writeRow = writeRow + 1
        End If
    Next r
    
    ' --- Process VLOOKUP untuk setiap baris (mulai dari kolom C) ---
    lastRowSurvey = wsSurvey.Cells(wsSurvey.Rows.Count, "A").End(xlUp).Row
    
    For r = 3 To lastRowSurvey
        vlookupValue = wsSurvey.Cells(r, "A").Address(False, True)
        
        For c = 3 To lastColPivot + 1 ' Mulai dari kolom C
            matchHeader = wsSurvey.Cells(2, c).Address(True, False)
            
            ' VLOOKUP: returns 1 if finds any number >=1, otherwise 0
            vlookupFormulaString = "=IF(IFERROR(VLOOKUP(" & vlookupValue & ", " & _
                                 pivotDataRange.Address(External:=True) & ", " & _
                                 "MATCH(" & matchHeader & ", " & pivotHeaderRowRange.Address(External:=True) & ", 0), " & _
                                 "FALSE), 0)>=1, 1, 0)"
            
            wsSurvey.Cells(r, c).Formula = vlookupFormulaString
        Next c
    Next r
    
    ' --- Hitung persentase langsung di sini ---
    Call HitungPersentaseDealerLangsung_Updated(wsSurvey)
    
    ' --- Format akhir ---
    lastColSurvey = wsSurvey.Cells(2, wsSurvey.Columns.Count).End(xlToLeft).Column
    
    ' Auto-fit columns
    wsSurvey.Columns.AutoFit
    
    ' Add borders
    With wsSurvey.Range(wsSurvey.Cells(1, 1), wsSurvey.Cells(lastRowSurvey, lastColSurvey)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Conditional formatting untuk 0
    If lastRowSurvey >= 3 And lastColSurvey >= 3 Then
        Dim dataRange As Range
        Set dataRange = wsSurvey.Range(wsSurvey.Cells(3, 3), wsSurvey.Cells(lastRowSurvey, lastColSurvey))
        
        dataRange.FormatConditions.Delete
        
        ' Format untuk angka 0
        With dataRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
            .Interior.Color = RGB(255, 255, 0)
        End With
        
        ' Format untuk persentase 0% - FIXED VERSION
        With dataRange.FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=A1=""0.0%""")
            .Interior.Color = RGB(255, 255, 0)
        End With
    End If
    
    ' Format untuk kolom B (Number of Outlet) - beri border juga
    If lastRowSurvey >= 3 Then
        With wsSurvey.Range(wsSurvey.Cells(3, 2), wsSurvey.Cells(lastRowSurvey, 2))
            .HorizontalAlignment = xlCenter
            .NumberFormat = "0"
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Data berhasil diolah di sheet SURVEY RATE dengan persentase dan Number of Outlet!", vbInformation
End Sub

Sub HitungPersentaseDealerLangsung_Updated(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, startRow As Long, endRow As Long
    Dim col As Long
    
    Application.Calculation = xlCalculationManual
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    
    i = 3
    Do While i <= lastRow
        If ws.Cells(i, 1).Font.Bold = True Then
            startRow = i + 1
            
            ' Cari baris bold berikutnya atau akhir data
            endRow = startRow
            Do While endRow <= lastRow And ws.Cells(endRow, 1).Font.Bold = False
                endRow = endRow + 1
            Loop
            endRow = endRow - 1
            
            ' Jika ada cabang di bawah dealer ini
            If startRow <= endRow Then
                ' Hitung jumlah cabang (tidak ditambah 1)
                Dim outletCount As Long
                outletCount = (endRow - startRow + 1) ' Hanya jumlah cabang
                
                ' Update Number of Outlet dengan jumlah cabang
                ws.Cells(i, 2).Value = outletCount
                ws.Cells(i, 2).Font.Bold = True
                ws.Cells(i, 2).Interior.Color = RGB(240, 240, 240)
                
                For col = 3 To lastCol ' Mulai dari kolom C
                    ' Hitung persentase: (jumlah 1) / (total cabang)
                    Dim sumRange As String
                    Dim countRange As String
                    
                    sumRange = ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)).Address(False, False)
                    countRange = ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)).Address(False, False)
                    
                    ' Formula: SUM / COUNT (COUNT hanya menghitung angka, jadi 0 dan 1)
                    ws.Cells(i, col).Formula = "=SUM(" & sumRange & ")/COUNT(" & countRange & ")"
                    ws.Cells(i, col).NumberFormat = "0.0%"
                    ws.Cells(i, col).Font.Bold = True
                    ws.Cells(i, col).Interior.Color = RGB(240, 240, 240)
                Next col
            Else
                ' Jika tidak ada cabang (dealer saja)
                ws.Cells(i, 2).Value = 0 ' Tidak ada outlet
                ws.Cells(i, 2).Font.Bold = True
                ws.Cells(i, 2).Interior.Color = RGB(240, 240, 240)
                
                For col = 3 To lastCol ' Mulai dari kolom C
                    ' Gunakan nilai dari dealer itu sendiri (0 atau 1)
                    If ws.Cells(i, col).Value = 1 Then
                        ws.Cells(i, col).Value = "100.0%"
                    Else
                        ws.Cells(i, col).Value = "0.0%"
                    End If
                    ws.Cells(i, col).NumberFormat = "0.0%"
                    ws.Cells(i, col).Font.Bold = True
                    ws.Cells(i, col).Interior.Color = RGB(240, 240, 240)
                Next col
            End If
            
            i = endRow + 1
        Else
            i = i + 1
        End If
    Loop
    
    ' HAPUS GRAND TOTAL - tidak ditambahkan
    
    Application.Calculate
    Application.Calculation = xlCalculationAutomatic
End Sub

' Macro untuk menghitung persentase saja (jika sudah ada data 0/1)
Sub HitungPersentasePerDealer_Fixed()
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    
    Set ws = ActiveSheet
    
    ' Panggil sub HitungPersentaseDealerLangsung dengan worksheet aktif
    Call HitungPersentaseDealerLangsung_Updated(ws)
    
    Application.ScreenUpdating = True
    
    MsgBox "Persentase per dealer berhasil dihitung!", vbInformation
End Sub

