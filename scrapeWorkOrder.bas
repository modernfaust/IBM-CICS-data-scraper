Attribute VB_Name = "Module1"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Outstanding Issues:
'1. Attachmate References removed when running this script on certain devices
'    Work around: create seperate copies of this workbook
'    Potential Solution: AddReferences() and RemoveReferences() subroutines should be run each time the script is run
'2. No custom exceptions made for:
'    A. Running the script while a job hasn't been received
'    B. Running the script while PP+ is in maintenance
'3. Unable to pick up more than 6 enclosures listed on the AJIS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub CANADA_WORK_ORDER()

Dim session As Object
Set session = ConnectPP
Dim ws
Dim weight(2) As String
Dim numEncl As Variant
ws = Worksheet
'AddReferences

JOBDIGIT = Len(Sheet1.Range("D4"))

If Sheet1.Range("D4") = "ENTER JOB# HERE" Or Sheet1.Range("D4") = "" Or JOBDIGIT <> 6 Then
    MsgBox ("PLEASE ENTER 6 DIGIT JOB# IN CELL D4")
GoTo ENDPROCESS
End If

If Sheet1.Range("E4") <> "" Or Sheet1.Range("DESENC1") <> "" Then
    MsgBox ("PLEASE RESET THE SHEET (RESET SHEET BUTTON")
GoTo ENDPROCESS
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''Get receipt, record, mailing and meeting dates
job = Sheet1.Range("D4")
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

RVDTYR = session.GetDisplayText(15, 21, 2)
RVDTMT = session.GetDisplayText(15, 23, 2)
RVDTDY = session.GetDisplayText(15, 25, 2)
RVDTFULL = RVDTMT & "/" & RVDTDY & "/" & RVDTYR

Sheet1.Range("D11") = RVDTFULL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''AN ERROR HERE MEANS THAT THE JOB HAS NOT BEEN RECEIVED ON PROXYPLUS
If Trim(RVDTFULL) = "  /  /  " Then
    Sheet1.Range("D12") = "JOB NOT RECEIVED"
Else
    Sheet1.Range("D12") = Application.WorksheetFunction.WorkDay(Sheet1.Range("D11"), 3)
End If
Sheet1.Range("H4:J4") = "Traditional Mailing"
Sheet1.Range("H6:J6") = "0-30g"
Sheet1.Range("K4:M4") = "Incentive Lettermail"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MULTI_JOBS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''Get Issuer name and classes
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
If CLng(Trim(session.GetDisplayText(18, 71, 9))) < CLng(Trim(session.GetDisplayText(19, 71, 9))) Then
    Sheet1.Range("E4") = Trim(session.GetDisplayText(18, 71, 9))
Else
    Sheet1.Range("E4") = Trim(session.GetDisplayText(19, 71, 9))
End If
Sheet1.Range("G12") = session.GetDisplayText(8, 16, 6)
Sheet1.Range("ISSUERNAME") = Trim(session.GetDisplayText(9, 16, 40))
session.TransmitTerminalKey rcIBMPf7Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    aa = 8
    zz = 12
    counter = 0
'''''Populate classes column
    Do While session.GetDisplayText(aa, 10, 1) <> " "
        If counter = 6 Then
            Sheet1.Range("I17") = "Add. Classes see below"
            GoTo RECORDDATE
        End If
        Sheet1.Range("H" & zz) = session.GetDisplayText(aa, 10, 3)
        zz = zz + 1
        aa = aa + 1
        counter = counter + 1
    Loop

RECORDDATE:
        session.TransmitTerminalKey rcIBMPf1Key
        session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1

'''''RECORD DATE
    RDYR = session.GetDisplayText(11, 21, 2)
    RDMT = session.GetDisplayText(11, 23, 2)
    RDDY = session.GetDisplayText(11, 25, 2)
    RCDDT = RDMT & "/" & RDDY & "/" & RDYR
    Sheet1.Range("D13") = RCDDT
'''''MEETING DATE
    MDYR = session.GetDisplayText(13, 21, 2)
    MDMT = session.GetDisplayText(13, 23, 2)
    MDDY = session.GetDisplayText(13, 25, 2)
    MCDDT = MDMT & "/" & MDDY & "/" & MDYR
    Sheet1.Range("D14") = MCDDT
'''''INTERNET DELIVERY
    Sheet1.Range("H25") = session.GetDisplayText(17, 48, 1)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET NI5 INDICATOR PRINTING TYPE
session.TransmitTerminalKey rcIBMPf11Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1

Dim PrintType As String
PrintType = "Manually Enter: Printing Type"
done = False
maxPages = 12

Do Until done = True Or maxPages = 0
    For Row = 9 To 21
        If session.GetDisplayText(Row, 11, 3) = "NI5" Then
            If session.GetDisplayText(Row, 3, 1) = "F" Then
                PrintType = "SINGLE PRINT (F)"
            ElseIf session.GetDisplayText(Row, 3, 1) = "U" Then
                PrintType = "OBO ONLY (U-NOT PAYING)"
            ElseIf session.GetDisplayText(Row, 3, 1) = "P" Then
                PrintType = "OBO ONLY (P-PAYING)"
            ElseIf session.GetDisplayText(Row, 3, 1) = "S" Then
                PrintType = "NOBO/OBO SPLIT (S)"
            done = True
            End If
        End If
    Next Row
    session.TransmitTerminalKey rcIBMPf2Key
    session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    maxPages = maxPages - 1
Loop

PRINTINSTRUCTION:
    Sheet1.Range("OBONOBOCON") = PrintType
    PrintType = ""

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
session.TransmitTerminalKey rcIBMPf15Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1

session.TransmitTerminalKey rcIBMPf5Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
INSCNT = 0
aa = 2
Do Until Sheet3.Range("N" & aa) = ""
    If Sheet3.Range("M" & aa) = "" Then
    Else
        INSCNT = INSCNT + 1
    End If
aa = aa + 1
Loop

If INSCNT > 9 Then
    MsgBox ("THERE ARE MORE THAN 9 INSERTS FOR THIS JOB, PLEASE REVIEW THE INSERT DATA AND APPLY REMAINING INSERTS MANUALLY TO THE WORK ORDER FORM")
INSCNT = 9
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
session.TransmitTerminalKey rcIBMPf1Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
session.TransmitTerminalKey rcIBMPf1Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''Get AIQ
session.TransmitTerminalKey rcIBMPf11Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    aa = 9
    Do Until session.GetDisplayText(aa, 11, 3) = "AIQ" Or aa = 22
        aa = aa + 1
    Loop
    If aa = 22 Then
    Else
        Sheet1.Range("H24") = Trim(session.GetDisplayText(aa, 18, 3))
    End If
session.TransmitTerminalKey rcIBMPf1Key
session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GET CLASSES FOR W JOB
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Trim(Sheet1.Range("D5")) = "" Then
    GoTo GETENCLOSURES
Else
    job = Sheet1.Range("D5")
End If

If job = "" Then
Else
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
    Sheet1.Range("E5") = Trim(session.GetDisplayText(18, 71, 9))
  
    session.TransmitTerminalKey rcIBMPf7Key
    session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    zz = 12
    If Sheet1.Range("I17") <> "Add. Classes please see below" Then
        Do Until Sheet1.Range("H" & zz) = ""
            zz = zz + 1
        Loop
    End If
    aa = 8
'''''Populate classes column
    For i = 8 To 14
        If session.GetDisplayText(i, 10, 1) <> " " Then
            If i = 14 Then
                Sheet1.Range("I17") = "Add. Classes please see below"
                GoTo RECORDDATE
            End If
            Sheet1.Range("H" & zz) = session.GetDisplayText(i, 10, 3)
            zz = zz + 1
        End If
    Next i
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GETENCLOSURES:
    Do Until session.GetDisplayText(1, 2, 12) = "CMENUM CMENU"
    session.TransmitTerminalKey rcIBMPf1Key
    session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    Loop
    job = Sheet1.Range("D4")
    
'''''Get enclosures
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
'''''Get Material weights (for both Notice and Full Package)
'''''If this is an N&A job, then weight(0) will be the Notice Weight
'''''Otherwise, weight(0) represents the total material weight
'''''weight(1) stores the Full Package weight for N&A jobs
    x = 15
    For i = 0 To 1
        If session.GetDisplayText(x, 74, 5) <> "00000" Then
            weight(i) = Trim(session.GetDisplayText(x, 74, 5))
            x = x - 1
        End If
    Next i
    
    session.TransmitTerminalKey rcIBMPf15Key
    session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
    
'''''Divide sheet into Full Package and Notice sheets for N&A type orders
    If session.GetDisplayText(8, 80, 1) = "Y" Then
        session.TransmitTerminalKey rcIBMPf15Key
        session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        
        Sheet1.Range("H4") = "N&A Full Package"
        Sheet1.Range("SAMPLEAMT") = "SAMPLE # 2 OF 2"
        
        ThisWorkbook.Sheets("Work Order").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ActiveSheet.Name = "Notice"
        Sheet1.Name = "Full Package"
        
        Sheets("Notice").Range("H4") = "N&A Notice Package"
        Sheets("Notice").Range("SAMPLEAMT") = "SAMPLE # 1 OF 2"
        
        session.TransmitTerminalKey rcIBMPf3Key
        session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1

'''''Fill enclosures table for Notice sheet
        Row = 12
        col = 7
        lang = 38
        For num = 1 To 9
            If Sheets("Notice").Range("DESENC" & num) = "" Then
                If Row = 17 Then
                    Row = 12
                    col = 46
                    lang = 77
                End If
                If session.GetDisplayText(Row, lang, 1) <> " " And session.GetDisplayText(Row, lang, 1) <> "_" Then
                    Sheets("Notice").Range("DESENC" & num) = session.GetDisplayText(Row, col, 29)
                    Sheets("Notice").Range("DESLNG" & num) = session.GetDisplayText(Row, lang, 1)
                End If
                Row = Row + 1
            End If
        Next num
'''''Set Material weight and Mailing type for Notice
        Call checkWeight(weight(0), Sheets("Notice"))
        Sheets("Notice").Range("K4") = typeMail(CLng(Sheets("Notice").Range("E4")), Sheets("Notice").Range("H6"))
        '''''''''''''''''''''''''''''''''''''
        session.TransmitTerminalKey rcIBMPf2Key
        session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        
'''''Fill enclosures table for Full Package sheet
        Row = 12
        col = 7
        lang = 38
        For num = 1 To 9
            If Sheets("Full Package").Range("DESENC" & num) = "" Then
                If Row = 17 Then
                    Row = 12
                    col = 46
                    lang = 77
                End If
                If session.GetDisplayText(Row, lang, 1) <> " " And session.GetDisplayText(Row, lang, 1) <> "_" Then
                    Sheets("Full Package").Range("DESENC" & num) = session.GetDisplayText(Row, col, 29)
                    Sheets("Full Package").Range("DESLNG" & num) = session.GetDisplayText(Row, lang, 1)
                End If
                Row = Row + 1
            End If
        Next num

'''''Set Material weight and mailing type for Full Package
        Call checkWeight(weight(1), Sheets("Full Package"))
        Sheets("Full Package").Range("K4") = typeMail(CLng(Sheets("Full Package").Range("E4")), Sheets("Full Package").Range("H6"))
    Else
        
        Sheet1.Range("SAMPLEAMT") = "SAMPLE # 1 OF 1"
        session.TransmitTerminalKey rcIBMPf15Key
        session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
            
'''''Fill enclosures table for Work Order sheet
        Row = 12
        col = 7
        lang = 38
        For num = 1 To 9
            If Sheet1.Range("DESENC" & num) = "" Then
                If Row = 17 Then
                    Row = 12
                    col = 46
                    lang = 77
                End If
                If session.GetDisplayText(Row, lang, 1) <> " " And session.GetDisplayText(Row, lang, 1) <> "_" Then
                    Sheet1.Range("DESENC" & num) = session.GetDisplayText(Row, col, 29)
                    Sheet1.Range("DESLNG" & num) = session.GetDisplayText(Row, lang, 1)
                End If
                Row = Row + 1
            End If
        Next num

'''''Set Material weight and mailing type for Work Order
        Call checkWeight(weight(0), Sheet1)
        Sheet1.Range("K4") = typeMail(CLng(Sheet1.Range("E4")), Sheet1.Range("H6"))
    End If
    
Do Until session.GetDisplayText(1, 2, 12) = "CMENUM CMENU"
    session.TransmitTerminalKey rcIBMPf1Key
    session.WaitForEvent rcKbdEnabled, "30", "0", 1, 1
Loop

ENDPROCESS:
End Sub


