Attribute VB_Name = "Module1"
Sub Demostep()
If (Cells(3, 4).Value = "�s�_��") Then
Cells(3, 4).Font.ColorIndex = 3      '�r���C��
Cells(3, 4).Interior.ColorIndex = 6   '�x�s�橳��
Else
Cells(3, 4).Font.ColorIndex = 1
Cells(3, 4).Interior.ColorIndex = 0
End If
End Sub

Sub Demostep2()
Dim odRidx As Long

'��ȰʺA����

Dim TargetValue As String
TargetValue = InputBox("�п�J�z�n�ϰ�O�z���")
'end of ��ȰʺA����

For odRidx = 3 To 24

    If (Cells(odRidx, 4).Value = TargetValue) Then
    
    Cells(odRidx, 4).Font.ColorIndex = 3      '�r���C��A���`�е��C��
    
    Cells(odRidx, 4).Interior.ColorIndex = 6    '�x�s�橳��A���`�е��C��
Else
    Cells(3, 4).Font.ColorIndex = 1
    Cells(3, 4).Interior.ColorIndex = 0
    
End If
Next
End Sub

Sub Demostep3()
Dim odRidx As Long

'��ȰʺA����
Dim TargetValue As Variant   '��θU�Ϋ��O
TargetValue = InputBox("�п�J�z�n�z���")

'����B�J�A�ƭȳB�z
If IsNumeric(TargetValue) Then
TargetValue = CDbl(TargetValue)
End If
'end of����B�J�A�ƭȳB�z

Dim TargetCIdx As Integer
TargetCIdx = InputBox("�п�J�z�n�z�諸 �����")

'end of ��ȰʺA����
'�̫��u��
Dim odStartRIdx As Integer
odStartRIdx = InputBox("�п�J�z�n���_�l�B�z_�C����_�����")

Dim dtRange As Range
Set dtRange = ActiveSheet.UsedRange
Dim rowNum As Long
rowNum = dtRange.Rows.Count
MsgBox rowNum
'end of �̲��u��!�ç�j��o3�M24�令odStartRIdx�MrowNum

For odRidx = odStartRIdx To rowNum

    If (Cells(odRidx, TargetCIdx).Value = TargetValue) Then
    
    Cells(odRidx, TargetCIdx).Font.ColorIndex = 5     '�r���C��
    
    Cells(odRidx, TargetCIdx).Interior.ColorIndex = 37   '�x�s�橳��
Else
    Cells(odRidx, TargetCIdx).Font.ColorIndex = 1
    Cells(odRidx, TargetCIdx).Interior.ColorIndex = 0
    
End If
Next
End Sub
