Attribute VB_Name = "Module2"
Sub RefreshSemuaPivot()

    ThisWorkbook.RefreshAll
    
    MsgBox "Semua Pivot Table telah di-refresh!"
    
    SinkronkanSemuaData
End Sub

Sub Save()
    Dim filePath As Variant
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="Report.xlsm", _
        FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm")

    If filePath <> False Then
        ThisWorkbook.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        MsgBox "File macro berhasil disimpan di: " & filePath
    End If
End Sub
