Attribute VB_Name = "Module1"
Sub temperature()

Dim i As Long
Dim Sum As Single

Sum = 0
For i = 1 To 14
    Randomize                         '�����n�񏉊���
    temp = Round(1.1 * Rnd + 35.7, 1) '�̉�
    Sum = Sum + temp
    ave = Round(Sum / i, 1)
    Cells(i, 1).Value = temp & " ��"
    Cells(i + 1, 1).ClearContents
    Cells(i + 2, 1).Value = "���ϑ̉��F" & ave & " ��"
    
Next i

End Sub
