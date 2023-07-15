Attribute VB_Name = "Module1"

Sub 醫療口罩排序()
Attribute 醫療口罩排序.VB_ProcData.VB_Invoke_Func = "G\n14"
'
' 醫療口罩排序由小到大 巨集
'
' 快速鍵: Ctrl+Shift+G
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B415") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B415")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 巨集4由大到小()
Attribute 巨集4由大到小.VB_ProcData.VB_Invoke_Func = " \n14"
'
'醫療口罩由大到小 巨集
'

'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B416") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B416")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("L9").Select
End Sub
