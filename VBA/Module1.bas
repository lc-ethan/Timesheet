Attribute VB_Name = "Module1"
Sub ADD_ROW(Optional ByVal FIRST_ROW As Boolean = False)
    Application.ScreenUpdating = False
    '' log the input data into the source data table
    ' object definition
    Dim tbl_input As ListObject
    Set tbl_input = Sheets("Input Page").ListObjects("TABLE_INPUT")

    Sheets("Input Page").Select
    tbl_input.ListRows.Add (1)

    tbl_input.ListRows(2).Range.Select
    Selection.Copy
    tbl_input.ListRows(1).Range.Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    
    tbl_input.DataBodyRange(1, tbl_input.ListColumns("Index").Index).Select
    ActiveCell.FormulaR1C1 = "=R[1]C+1"
    
    Dim SELECTED_CELL As Range
    Set SELECTED_CELL = tbl_input.DataBodyRange(1, tbl_input.ListColumns("Delete").Index)
    
    Call CREATE_DELETE_BUTTON(SELECTED_CELL, FIRST_ROW)

    
    tbl_input.DataBodyRange(1, 1).Select
    Application.ScreenUpdating = True
End Sub


Sub LOG_INPUT()

    Application.ScreenUpdating = False
    
    '' log the input data into the source data table
    ' object definition
    Dim tbl_source As ListObject
    Dim tbl_source_row As Long
    Dim tbl_input As ListObject
    Dim tbl_input_row As Long
    Dim vt_headers As Variant
    Dim vt_row_index As Integer
    
    Set tbl_source = Sheets("Time Sheet").ListObjects("TABLE_SOURCE")
        tbl_source_row = tbl_source.Range.Rows.Count - 1
    Set tbl_input = Sheets("Input Page").ListObjects("TABLE_INPUT")
        vt_headers = Array("Date", "Start", "End", "Task", "Comment")
    
    
    Sheets("Input Page").Select
    tbl_input.ListColumns("Index").DataBodyRange.Select
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    '' sorting the input data
    'Range("TABLE_INPUT[[#Headers],[Index]]").Select
    'ActiveWorkbook.Worksheets("Input Page").ListObjects("TABLE_INPUT").Sort. _
    '    SortFields.Clear
    'ActiveWorkbook.Worksheets("Input Page").ListObjects("TABLE_INPUT").Sort. _
    '    SortFields.Add Key:=Range("TABLE_INPUT[[#Headers],[Index]]"), SortOn:= _
    '    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'With ActiveWorkbook.Worksheets("Input Page").ListObjects("TABLE_INPUT").Sort
    '    .Header = xlYes
    '    .MatchCase = False
    '   .Orientation = xlTopToBottom
    '    .SortMethod = xlPinYin
    '    .Apply
    'End With
    
    ActiveWorkbook.Worksheets("Input Page").ListObjects("TABLE_INPUT").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Input Page").ListObjects("TABLE_INPUT").Sort. _
        SortFields.Add Key:=Range("TABLE_INPUT[Date]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Input Page").ListObjects("TABLE_INPUT").Sort. _
        SortFields.Add Key:=Range("TABLE_INPUT[Start]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Input Page").ListObjects("TABLE_INPUT").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

          
    Sheets("Time Sheet").Select
    ' situation where the time sheet is first used
    If (tbl_source_row) = 1 And IsEmpty(tbl_source.DataBodyRange(tbl_source_row, 1).Value) Then
    
        vt_row_index = 1
                
    Else
    'situation where time sheet has been used
        tbl_source.ListRows.Add AlwaysInsert:=True
        tbl_source_row = tbl_source.Range.Rows.Count - 1
        vt_row_index = tbl_source_row
        
        
        tbl_source.ListRows(vt_row_index - 1).Range.Select
        Selection.Copy
        tbl_source.ListRows(vt_row_index).Range.Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        
    End If
    
    ' copy and paste the input data table into the source data table
    For i = 0 To UBound(vt_headers)
        Sheets("Input Page").Select
        tbl_input.ListColumns(vt_headers(i)).DataBodyRange.Select
        Selection.Copy
            
        Sheets("Time Sheet").Select
        tbl_source.DataBodyRange(tbl_source_row, tbl_source.ListColumns(vt_headers(i)).Index).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
    Next i
    
    '' refresh the input log
    Sheets("Input Page").Select
    
    ' create entry 0
    ADD_ROW (True)
    
    tbl_input.DataBodyRange(1, tbl_input.ListColumns("Index").Index).Select
    ActiveCell.FormulaR1C1 = "1"
       
    ' delete the old entries
    tbl_input_row = tbl_input.Range.Rows.Count - 1
    Do While tbl_input_row > 1
        tbl_input.ListRows(2).Delete
        tbl_input_row = tbl_input.Range.Rows.Count - 1
    Loop
 
    Application.ScreenUpdating = True
    
    MsgBox "New entries have been logged!"
    
    
End Sub
