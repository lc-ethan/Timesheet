Attribute VB_Name = "Module2"
Private Sub SWITCH_ON()
    Application.ScreenUpdating = False
    Application.Cursor = xlNorthwestArrow
    Dim tbl_input As ListObject
    Set tbl_input = Sheets("Input Page").ListObjects("TABLE_INPUT")
    
    'MsgBox Date
    
    If IsEmpty(tbl_input.DataBodyRange(1, tbl_input.ListColumns("Start").Index).Value) Then
        tbl_input.DataBodyRange(1, tbl_input.ListColumns("Start").Index).Value = Format(Time, "hh:mm AM/PM")
        tbl_input.DataBodyRange(1, tbl_input.ListColumns("Date").Index).Value = Date
        tbl_input.DataBodyRange(1, tbl_input.ListColumns("Task").Index).Value = Sheets("Input Page").Range("TRACKER_TASK").Value
    Else
        ADD_ROW
        tbl_input.DataBodyRange(1, tbl_input.ListColumns("Start").Index).Value = Format(Time, "hh:mm AM/PM")
        tbl_input.DataBodyRange(1, tbl_input.ListColumns("Date").Index).Value = Date
        tbl_input.DataBodyRange(1, tbl_input.ListColumns("Task").Index).Value = Sheets("Input Page").Range("TRACKER_TASK").Value
    End If
    
    Call START_TRACKER
End Sub

Private Sub SWITCH_OFF()
    Dim tbl_input As ListObject
    Set tbl_input = Sheets("Input Page").ListObjects("TABLE_INPUT")
    
    tbl_input.DataBodyRange(1, tbl_input.ListColumns("End").Index).Value = Format(Time, "hh:mm AM/PM")
    tbl_input.DataBodyRange(1, tbl_input.ListColumns("Date").Index).Value = Date
    
    Call STOP_TRACKER
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    
End Sub




Private Sub START_TRACKER()
    Application.OnTime Now + TimeValue("00:00:01"), "NEXT_TICK"
End Sub



Private Sub NEXT_TICK()
    Dim WB As Workbook
    Set WB = ThisWorkbook
    
    WB.Sheets("Task List").Range("TRACKED_TIMER").Value = WB.Sheets("Task List").Range("TRACKED_TIMER").Value + TimeValue("00:00:01")
    
    Call START_TRACKER
    
End Sub



Private Sub STOP_TRACKER()
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "NEXT_TICK", , False

End Sub

Private Sub RESET_TRACKER()
    
    Call STOP_TRACKER
    
    Sheets("Task List").Range("TRACKED_TIMER").Value = Sheets("Task List").Range("INITIAL_TIMER").Value
    Sheets("Input Page").Range("TRACKER_TASK").MergeArea.ClearContents

End Sub



