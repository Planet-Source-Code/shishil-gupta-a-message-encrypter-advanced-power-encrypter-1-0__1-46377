Attribute VB_Name = "Module1"
Function GenRnd(DigNum As String) As String
If DigNum = "" Then DigNum = "15"
If Form1.List1.ListIndex = -1 Then Form1.List1.ListIndex = 0
GenRnd = "."
For b = 1 To Val(DigNum)
    If Form1.List1.ListIndex + 900 > 1000 Then
        a = Form1.List1.ListIndex + 899 - 1000
    Else
        a = Form1.List1.ListIndex + 900
    End If
    Form1.List1.ListIndex = a
    GenRnd = GenRnd & Form1.List1.Text
Next b

End Function
