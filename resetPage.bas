Attribute VB_Name = "Module2"
Sub RESET_CANWO()

'Delete Notice sheet
For Each Sheet In ActiveWorkbook.Worksheets
    If Sheet.Name = "Notice" Then
        Application.DisplayAlerts = False
        Worksheets("Notice").Delete
        Application.DisplayAlerts = True
    End If
Next Sheet
'Change fields to default
Sheet1.Name = "Work Order"
Sheet1.Range("H4:J4").ClearContents
Sheet1.Range("H6:J6").ClearContents
Sheet1.Range("K4:M4").ClearContents
Sheet1.Range("THRESHOLD1") = ""

Sheet3.Range("A2:F1048576").ClearContents
Sheet3.Range("M2:M1048576").ClearContents
Sheet3.Range("O69:O74").ClearContents

Sheet1.Range("G31:G37").ClearContents

Sheet1.Range("OBONOBOCON") = ""
Sheet1.Range("MULTIPRCON") = ""
Sheet1.Range("SAMPLEAMT") = "SAMPLE # _______ OF _______"
Sheet1.Range("D4") = "ENTER JOB# HERE"
Sheet1.Range("THRESHOLD1") = ""
Sheet1.Range("D5").ClearContents
Sheet1.Range("D6").ClearContents
Sheet1.Range("D7").ClearContents
Sheet1.Range("E4").ClearContents
Sheet1.Range("E5").ClearContents
Sheet1.Range("E6").ClearContents
Sheet1.Range("E7").ClearContents
Sheet1.Range("ISSUERNAME") = ""
Sheet1.Range("D11").ClearContents
Sheet1.Range("D12").ClearContents
Sheet1.Range("D13").ClearContents
Sheet1.Range("D14").ClearContents
Sheet1.Range("H28").ClearContents
Sheet1.Range("G12").ClearContents
Sheet1.Range("H12").ClearContents
Sheet1.Range("G13").ClearContents
Sheet1.Range("H13").ClearContents
Sheet1.Range("G14").ClearContents
Sheet1.Range("H14").ClearContents
Sheet1.Range("G15").ClearContents
Sheet1.Range("H15").ClearContents
Sheet1.Range("G16").ClearContents
Sheet1.Range("H16").ClearContents
Sheet1.Range("G17").ClearContents
Sheet1.Range("H17").ClearContents
Sheet1.Range("M10") = "NO"
Sheet1.Range("H24").ClearContents
Sheet1.Range("H25").ClearContents

Sheet1.Range("DESENC1") = ""
Sheet1.Range("DESLNG1") = ""
Sheet1.Range("DESENC2") = ""
Sheet1.Range("DESLNG2") = ""
Sheet1.Range("DESENC3") = ""
Sheet1.Range("DESLNG3") = ""
Sheet1.Range("DESENC4") = ""
Sheet1.Range("DESLNG4") = ""
Sheet1.Range("DESENC5") = ""
Sheet1.Range("DESLNG5") = ""
Sheet1.Range("DESENC6") = ""
Sheet1.Range("DESLNG6") = ""
Sheet1.Range("DESENC7") = ""
Sheet1.Range("DESLNG7") = ""
Sheet1.Range("DESENC8") = ""
Sheet1.Range("DESLNG8") = ""
Sheet1.Range("DESENC9") = ""
Sheet1.Range("DESLNG9") = ""

Sheet1.Range("AUDITI1") = "Initial:"
Sheet1.Range("AUDITI2") = "Initial:"
Sheet1.Range("AUDITI3") = "Initial:"
Sheet1.Range("AUDITIPC") = "Initial:"
Sheet1.Range("AUDITD1") = "Date:"
Sheet1.Range("AUDITD2") = "Date:"
Sheet1.Range("AUDITD3") = "Date:"
Sheet1.Range("AUDITDPC") = "Date:"
Sheet1.Range("AUDITT1") = "Time:"
Sheet1.Range("AUDITT2") = "Time:"
Sheet1.Range("AUDITT3") = "Time:"
Sheet1.Range("AUDITTPC") = "Time:"

Sheet1.Range("SPECINST") = ""
Sheet1.Range("OTHERSRVFEE") = ""

End Sub

Sub MULTI_JOBS()

Dim session As Object
Set session = ConnectPP
'AddReferences

dd = 3


Sheet3.Range("A2") = Sheet1.Range("D4")

job = Sheet3.Range("A2")
'''''REMOVE FOR PHASE 2
GoTo PULLWJOB
'''''REMOVE FOR PHASE 2
    Do Until session.GetDisplayText(1, 2, 12) = "CMENUM CMENU"
    session.TransmitTerminalKey rcIBMPf1Key
    session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    Loop
    With session
        .WaitForEvent rcEnterPos, "30", "0", 24, 76
        .WaitForDisplayString "NO.-->", "30", 24, 69
        .TransmitANSI "195"
        .TransmitTerminalKey rcIBMEnterKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        .WaitForEvent rcEnterPos, "30", "0", 7, 16
        .WaitForDisplayString ":", "30", 7, 14
        .TransmitANSI job
        .TransmitTerminalKey rcIBMEnterKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    End With
    
session.TransmitTerminalKey rcIBMPf15Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
session.TransmitTerminalKey rcIBMPf5Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
session.TransmitTerminalKey rcIBMPf8Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1

aa = 9
Do Until aa = 21
bb = 6
    Do Until bb = 69
        If session.GetDisplayText(aa, bb, 1) = "_" Then
        Else
           Sheet3.Range("A" & dd) = session.GetDisplayText(aa, bb, 6)
           dd = dd + 1
        End If
    bb = bb + 9
    Loop
aa = aa + 1
Loop
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

PULLWJOB:
dd = 2
Do Until Sheet3.Range("A" & dd) = ""
    job = Sheet3.Range("A" & dd)
    Do Until session.GetDisplayText(1, 2, 12) = "CMENUM CMENU"
    session.TransmitTerminalKey rcIBMPf1Key
    session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    Loop
    With session
        .WaitForEvent rcEnterPos, "30", "0", 24, 76
        .WaitForDisplayString "NO.-->", "30", 24, 69
        .TransmitANSI "195"
        .TransmitTerminalKey rcIBMEnterKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        .WaitForEvent rcEnterPos, "30", "0", 7, 16
        .WaitForDisplayString ":", "30", 7, 14
        .TransmitANSI job
        .TransmitTerminalKey rcIBMEnterKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    End With
    
    Sheet3.Range("C" & dd) = session.GetDisplayText(11, 47, 2)
dd = dd + 1
Loop



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''pull W job----need to complete

dd = 2
ff = 0
Do Until Sheet3.Range("A" & dd) = ""
ff = ff + 1
If Sheet3.Range("C" & dd) = "BC" Then
    GoTo BCData
Else
    GoTo USCANData
End If
USCANData:
    Sheet3.Range("B" & dd) = "NO W JOB - US/BN or RC"
GoTo NEXTITEM
BCData:
'''''''pull W job----need to complete
    job = Sheet3.Range("A" & dd)
    Do Until session.GetDisplayText(1, 2, 12) = "CMENUM CMENU"
    session.TransmitTerminalKey rcIBMPf1Key
    session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    Loop
    With session
        .WaitForEvent rcEnterPos, "30", "0", 24, 76
        .WaitForDisplayString "NO.-->", "30", 24, 69
        .TransmitANSI "11"
        .TransmitTerminalKey rcIBMEnterKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        .WaitForEvent rcEnterPos, "30", "0", 24, 74
        .WaitForDisplayString ">", "30", 24, 72
        .TransmitANSI "2"
        .TransmitTerminalKey rcIBMEnterKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        .WaitForEvent rcEnterPos, "30", "0", 24, 74
        .WaitForDisplayString ">", "30", 24, 72
        .TransmitANSI "1"
        .TransmitTerminalKey rcIBMEnterKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        .WaitForEvent rcEnterPos, "30", "0", 12, 44
        .WaitForDisplayString ":", "30", 12, 42
        .TransmitANSI job
        .TransmitTerminalKey rcIBMEnterKey
        .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    End With
    
    Sheet3.Range("B" & dd) = session.GetDisplayText(5, 13, 6)
    
    OBONOBO = session.GetDisplayText(3, 78, 1)
    
    If OBONOBO = "P" Then
        Sheet3.Range("D" & dd) = "OBO ONLY (P-PAYING)"
        GoTo NEXTITEM
    End If
    If OBONOBO = "U" Then
        Sheet3.Range("D" & dd) = "OBO ONLY (U-NOT PAYING)"
        GoTo NEXTITEM
    End If
    If OBONOBO = "F" Then
        Sheet3.Range("D" & dd) = "SINGLE PRINT (F)"
        GoTo NEXTITEM
    End If
    If OBONOBO = "S" Then
        Sheet3.Range("D" & dd) = "NOBO/OBO SPLIT (S)"
        GoTo NEXTITEM
    End If
    
    Sheet3.Range("D" & dd) = "NOT CODED: ENTER MANUALLY"

NEXTITEM:
Sheet3.Range("F" & dd) = ff
dd = dd + 1
Loop


If dd = 3 Then
dd = dd - 1
Sheet3.Range("F" & dd) = "1"
    If Sheet3.Range("B" & dd) = "NO W JOB - US/BN or RC" Then
    Else
        Sheet1.Range("D5") = Sheet3.Range("B" & dd)
    End If
'Sheet1.Range("SAMPLEAMT") = "SAMPLE #   1   OF   1"
'Sheet1.Range("OBONOBOCON") = Sheet3.Range("D" & dd)
Else
'add code for multi sheets-----need to complete



End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Do Until session.GetDisplayText(1, 2, 12) = "CMENUM CMENU"
session.TransmitTerminalKey rcIBMPf1Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
Loop
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub


