Attribute VB_Name = "Module6"
Sub check1_stampTime()
    ActiveSheet.Range("AUDITD1") = "Date: " & Date
    ActiveSheet.Range("AUDITT1") = "Time: " & Time
End Sub
Sub check2_stampTime()
    ActiveSheet.Range("AUDITD2") = "Date: " & Date
    ActiveSheet.Range("AUDITT2") = "Time: " & Time
End Sub

Sub check3_stampTime()
    ActiveSheet.Range("AUDITD3") = "Date: " & Date
    ActiveSheet.Range("AUDITT3") = "Time: " & Time
End Sub

Sub prcCtrl_stampTime()
    ActiveSheet.Range("AUDITDPC") = "Date: " & Date
    ActiveSheet.Range("AUDITTPC") = "Time: " & Time
End Sub

