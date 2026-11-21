Attribute VB_Name = "frmDeleteData"
Attribute VB_Base = "0{4AFAB9B1-C97E-4FB0-BFCE-AC556BC4189E}{BFF2022C-655B-4C9A-A4AB-CE2A04BB652D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public ws As Worksheet
Public colFilter As Long   ' Kolom untuk filter data
Public keyword As String   ' Keyword search
Private dataArr() As Variant ' Array untuk menampung data cepat

Private Sub UserForm_Initialize()
    Dim lastRow As Long, i As Long, dict As Object
    Set ws = ThisWorkbook.Sheets("Data")
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = ws.Cells(ws.Rows.Count, colFilter).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' Load seluruh kolom ke array
    dataArr = ws.Range(ws.Cells(2, colFilter), ws.Cells(lastRow, colFilter)).value
    
    ' Populate ListBox dengan dictionary untuk unik + filter keyword
    lstData.Clear
    For i = 1 To UBound(dataArr, 1)
        If keyword = "" Or InStr(1, dataArr(i, 1), keyword, vbTextCompare) > 0 Then
            If Not dict.Exists(dataArr(i, 1)) Then
                dict(dataArr(i, 1)) = 1
                lstData.AddItem dataArr(i, 1)
            End If
        End If
    Next i
End Sub

Private Sub btnDelete_Click()
    Dim i As Long, lastRow As Long
    Dim cell As Range
    Dim selectedItems As Collection
    Set selectedItems = New Collection
    
    ' Ambil data yang dicentang
    For i = 0 To lstData.ListCount - 1
        If lstData.Selected(i) Then selectedItems.Add lstData.List(i)
    Next i
    
    If selectedItems.Count = 0 Then
        MsgBox "No item selected.", vbExclamation
        Exit Sub
    End If
    
    ' Loop dari bawah ke atas untuk hapus baris di sheet
    lastRow = ws.Cells(ws.Rows.Count, colFilter).End(xlUp).Row
    Dim r As Long
    For r = lastRow To 2 Step -1
        For i = 1 To selectedItems.Count
            If ws.Cells(r, colFilter).value = selectedItems(i) Then
                ws.Rows(r).Delete

