Attribute VB_Name = "deletemenu"
Attribute VB_Base = "0{C2D22419-E5D4-405A-9541-D5F231C520C9}{FCAC4423-2CAD-4C83-A6EF-2FD8217F1A06}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Sub DeleteDataMenu()
Private Sub btnOK_Click()
    Dim choice As Long
    If ListBox1.ListIndex = -1 Then
        MsgBox "Please select a method.", vbExclamation
        Exit Sub
    End If
    
    choice = ListBox1.ListIndex + 1
    Me.Hide
    
    Select Case choice
        Case 1
            Call DeleteByTimestamp
        Case 2
            Call DeleteByMediaEdukasi
        Case 3
            Call DeleteByKeyword
        Case 4
            Call ManualDelete
    End Select
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ListBox1.Clear
    ListBox1.AddItem "Delete by Timestamp Range"
    ListBox1.AddItem "Delete by Media Edukasi"
    ListBox1.AddItem "Delete by Keyword / Search"
    ListBox1.AddItem "Manual Selection"
End Sub
