Attribute VB_Name = "Scratch2"
Dim ANum As Long, BTxt As String

Sub Test001Selection()
    myArr1 = Array()
    myArr2 = Array(2, 3, 4)
    myArr3 = Array("eo", "o", "z")
    newArr = A_CombineArray(myArr1, myArr2, myArr3)
    A_printArr (newArr)
End Sub

Sub Test002Input()
    Dim currSelection As range
    Set currSelect = Selection
    
    txt1 = currSelect.Areas(1).value
    txt2 = currSelect.Areas(2).value
    n1 = Len(txt1)
    n2 = Len(txt2)
    Dim word_list, sentence_list As range
    
    
    If n1 > n2 Then
        Set sentence_list = currSelect.Areas(1)
        Set word_list = currSelect.Areas(2)
    Else
        Set word_list = currSelect.Areas(1)
        Set sentence_list = currSelect.Areas(2)
    End If
    
    sentence_list.Interior.color = vbBlue
    word_list.Interior.color = vbRed
    
    
End Sub



