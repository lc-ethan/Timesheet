Attribute VB_Name = "Module3"
Sub CREATE_DELETE_BUTTON(SELECTED_CELL As Range, Optional ByVal FIRST_ROW As Boolean = False)
'
' CREATE_DELETE_BUTTON Macro
'

'
Dim clLeft As Double
Dim clTop As Double
Dim clWidth As Double
Dim clHeight As Double

clLeft = SELECTED_CELL.Left
clTop = SELECTED_CELL.Top
clHeight = SELECTED_CELL.Height
clWidth = SELECTED_CELL.Width

    ActiveSheet.Shapes.AddShape(msoShapeMathMultiply, clLeft + 21, clTop + 1.5, 13.0434645669, 13.0434645669).Select
    Selection.ShapeRange.IncrementLeft 5.2173228346
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.5
        .Transparency = 0
        .Solid
    End With
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.5
        .Transparency = 0
    End With
    
    If FIRST_ROW Then
        Selection.OnAction = "CLEAR_LAST_ROW"
    Else
        Selection.OnAction = "DELETE_ROW"
    End If
    
End Sub


Sub DELETE_ROW()

    Application.ScreenUpdating = False
    Dim b As Object
    Dim SelectedCell As Range

    Dim lFirstCol As Long
    Dim lLastCol As Long
    Dim lActiveCol As Long

    
    Set b = ActiveSheet.Shapes(Application.Caller)
    With b
        .TopLeftCell.Select
    End With
    
    Set SelectedCell = ActiveCell
    
    lActiveCol = SelectedCell.Column
    
    With SelectedCell.ListObject
        lFirstCol = .ListColumns(1).Range.Column
        lLastCol = .ListColumns(.ListColumns.Count).Range.Column
    End With
    
    SelectedCell.Offset(, -(lActiveCol - lFirstCol)) _
        .Resize(, lLastCol - lFirstCol + 1).Delete
    
    'MsgBox SelectedCell.ListObject.DataBodyRange.Row
    
    'With b.TopLeftCell
    '    RowNum = .Row
    'End With
    'Rows(RowNum).Select
    'Selection.Delete Shift:=xlUp
    'Application.ScreenUpdating = False
End Sub




Sub CLEAR_LAST_ROW()

    Application.ScreenUpdating = False
    Dim b As Object
    Dim SelectedCell As Range

    Dim lFirstCol As Long
    Dim lLastCol As Long
    Dim lActiveCol As Long
    Dim lRow As Long

    
    Set b = ActiveSheet.Shapes(Application.Caller)
    With b
        .TopLeftCell.Select
    End With
    
    Set SelectedCell = ActiveCell
    
    lActiveCol = SelectedCell.Column
    
    With SelectedCell.ListObject
        lFirstCol = .ListColumns(1).Range.Column
        lLastCol = .ListColumns(.ListColumns.Count).Range.Column
    End With
    
    SelectedCell.Offset(, -(lActiveCol - lFirstCol)) _
        .Resize(, lLastCol - lFirstCol + 1).ClearContents
    
    lRow = SelectedCell.ListObject.Range.Rows.Count - 1
    
    SelectedCell.ListObject.DataBodyRange(lRow, SelectedCell.ListObject.ListColumns("Index").Index).Value = 1
    
    
    'MsgBox SelectedCell.ListObject.DataBodyRange.Row
    
    'With b.TopLeftCell
    '    RowNum = .Row
    'End With
    'Rows(RowNum).Select
    'Selection.Delete Shift:=xlUp
    'Application.ScreenUpdating = False
End Sub


Sub CLEAR_ALL()

    Application.ScreenUpdating = False

    Dim tbl_input As ListObject
    Set tbl_input = Sheets("Input Page").ListObjects("TABLE_INPUT")


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
End Sub

