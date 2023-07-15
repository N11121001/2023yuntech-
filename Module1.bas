Attribute VB_Name = "Module1"
Sub Demostep()
If (Cells(3, 4).Value = "新北市") Then
Cells(3, 4).Font.ColorIndex = 3      '字體顏色
Cells(3, 4).Interior.ColorIndex = 6   '儲存格底色
Else
Cells(3, 4).Font.ColorIndex = 1
Cells(3, 4).Interior.ColorIndex = 0
End If
End Sub

Sub Demostep2()
Dim odRidx As Long

'實務動態互動

Dim TargetValue As String
TargetValue = InputBox("請輸入您要區域別篩選值")
'end of 實務動態互動

For odRidx = 3 To 24

    If (Cells(odRidx, 4).Value = TargetValue) Then
    
    Cells(odRidx, 4).Font.ColorIndex = 3      '字體顏色，異常標註顏色
    
    Cells(odRidx, 4).Interior.ColorIndex = 6    '儲存格底色，異常標註顏色
Else
    Cells(3, 4).Font.ColorIndex = 1
    Cells(3, 4).Interior.ColorIndex = 0
    
End If
Next
End Sub

Sub Demostep3()
Dim odRidx As Long

'實務動態互動
Dim TargetValue As Variant   '改用萬用型別
TargetValue = InputBox("請輸入您要篩選值")

'關鍵步驟，數值處理
If IsNumeric(TargetValue) Then
TargetValue = CDbl(TargetValue)
End If
'end of關鍵步驟，數值處理

Dim TargetCIdx As Integer
TargetCIdx = InputBox("請輸入您要篩選的 欄索引")

'end of 實務動態互動
'最後優化
Dim odStartRIdx As Integer
odStartRIdx = InputBox("請輸入您要的起始處理_列索引_欄索引")

Dim dtRange As Range
Set dtRange = ActiveSheet.UsedRange
Dim rowNum As Long
rowNum = dtRange.Rows.Count
MsgBox rowNum
'end of 最終優化!並把迴圈得3和24改成odStartRIdx和rowNum

For odRidx = odStartRIdx To rowNum

    If (Cells(odRidx, TargetCIdx).Value = TargetValue) Then
    
    Cells(odRidx, TargetCIdx).Font.ColorIndex = 5     '字體顏色
    
    Cells(odRidx, TargetCIdx).Interior.ColorIndex = 37   '儲存格底色
Else
    Cells(odRidx, TargetCIdx).Font.ColorIndex = 1
    Cells(odRidx, TargetCIdx).Interior.ColorIndex = 0
    
End If
Next
End Sub
