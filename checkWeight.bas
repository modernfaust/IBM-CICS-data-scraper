Attribute VB_Name = "Module5"
Sub checkWeight(matWeight As String, currentSheet As Worksheet)
    Dim matWeightLNG As Long
    If IsNumeric(matWeight) Then
        If IsEmpty(matWeight) Then
            matWeightLNG = 0
        Else
            matWeightLNG = CLng(matWeight)
        End If
    Else
        matWeightLNG = 0
    End If
    If matWeightLNG >= 0 And matWeightLNG <= 30 Then
        currentSheet.Range("H6") = "0-30g"
    ElseIf matWeightLNG > 30 And matWeightLNG <= 50 Then
        currentSheet.Range("H6") = "30-50g"
    ElseIf matWeightLNG > 50 And matWeightLNG <= 100 Then
        currentSheet.Range("H6") = "50 - 100g"
    ElseIf matWeightLNG > 100 And matWeightLNG <= 200 Then
        currentSheet.Range("H6") = "100-200g"
    ElseIf matWeightLNG > 200 And matWeightLNG <= 300 Then
        currentSheet.Range("H6") = "200-300g"
    ElseIf matWeightLNG > 300 And matWeightLNG <= 400 Then
        currentSheet.Range("H6") = "300-400g"
    ElseIf matWeightLNG > 400 And matWeightLNG <= 500 Then
        currentSheet.Range("H6") = "400-500g"
    Else
        currentSheet.Range("H6") = "Over Weight"
    End If
    
    
End Sub

