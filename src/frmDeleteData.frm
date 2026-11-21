Attribute VB_Name = "frmDeleteData"
Attribute VB_Base = "0{F142B7F9-6187-48B2-AA95-F42A930787C4}{267166C1-4BBE-4E19-865B-6917A887FC8E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub btnSearch_Click()
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim key As String
    key = txtKeyword.text
    If key = "" Then
        MsgBox "Please enter keyword.", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("Data")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    lstResult.Clear
    For i = 2 To lastRow
        If InStr(1, ws.Cells(i, 1).Value & " " & ws.Cells(i, 2).Value & " " & ws.Cells(i, 6).Value, key, vbTextCompare) > 0 Then
            lstResult.AddItem ws.Cells(i, 1).Value & " | " & ws.Cells(i, 6).Value
        End If
    Next i
    
    If lstResult.ListCount = 0 Then MsgBox "No matching data found.", vbInformation
End Sub

Private Sub btnDelete_Click()
    Dim ws As Worksheet, lastRow As Long, i As Long, r As Long
    Dim selectedItems As Collection
    Set selectedItems = New Collection
    Set ws = ThisWorkbook.Sheets("Data")
    
    For i = 0 To lstResult.ListCount - 1
        If lstResult.Selected(i) Then selectedItems.Add lstResult.List(i)
    Next i
    
    If selectedItems.Count = 0 Then
        MsgBox "No item selected.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Delete selected keyword data?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = lastRow To 2 Step -1
        For i = 1 To selectedItems.Count
            If InStr(1, ws.Cells(r, 1).Value & " | " & ws.Cells(r, 6).Value, selectedItems(i), vbTextCompare) > 0 Then
                ws.Rows(r).Delete
                Exit For
            End If
        Next i
    Next r
    
    MsgBox "Selected keyword data deleted.", vbInformation
    Unload Me
End Sub

