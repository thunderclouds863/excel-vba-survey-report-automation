Attribute VB_Name = "frmDeleteMedia"
Attribute VB_Base = "0{403B788E-1505-4740-8551-919888730131}{236FF019-2169-494A-A1E0-D721BC54951A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub btnDelete_Click()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, lastRow As Long, r As Long, i As Long
    Dim selectedItems As Collection
    Set selectedItems = New Collection
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' Ambil media yang dicentang
    For i = 0 To lstMedia.ListCount - 1
        If lstMedia.Selected(i) Then selectedItems.Add lstMedia.List(i)
    Next i
    
    If selectedItems.Count = 0 Then
        MsgBox "No media selected.", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete all data for selected media?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    
    ' Hapus dari bawah ke atas
    For r = lastRow To 2 Step -1
        For i = 1 To selectedItems.Count
            If ws.Cells(r, 6).Value = selectedItems(i) Then
                ws.Rows(r).Delete
                Exit For
            End If
        Next i
    Next r
    
    MsgBox "Selected media data deleted.", vbInformation
    Unload Me
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
