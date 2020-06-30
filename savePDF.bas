Attribute VB_Name = "Module7"
Sub savePDF()
'''''TODO: CREATE FOLDER FOR MONTH IF IT DOESN'T EXIST
    Dim dateArray() As Variant
    
    job = ActiveSheet.Range("D4")
    issuer = ActiveSheet.Range("ISSUERNAME")
    mailDate = VBA.format(ActiveSheet.Range("D12"), "mmmm.dd")
    
    pdfName = "\\mkipvwdrsf01.bsg.ad.adp.com\common\LimitedAccess\work orders\1 PDF Work Orders\" & mailDate & " " & job & " " & issuer
    If ActiveSheet.Name = "Full Package" Then
        pdfName = pdfName & " (FULL PACKAGE)"
    ElseIf ActiveSheet.Name = "Notice" Then
        pdfName = pdfName & " (NOTICE)"
    End If
        pdfName = pdfName & ".pdf"

    ActiveSheet.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=pdfName, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=False

End Sub

Sub saveXLSM()
    job = ActiveSheet.Range("D4")
    issuer = ActiveSheet.Range("ISSUERNAME")
    mailDate = VBA.format(ActiveSheet.Range("D12"), "mmmm.dd")
    mailFolder = VBA.format(ActiveSheet.Range("D12"), "mmmm") & " EXCEL"
    
    xlsmName = "\\mkipvwdrsf01.bsg.ad.adp.com\common\LimitedAccess\work orders\" & mailFolder & "\" & mailDate & " " & job & " " & issuer
    If ActiveSheet.Name = "Full Package" Or ActiveSheet.Name = "Notice" Then
        xlsmName = xlsmName & " (N&A)"
    End If
        xlsmName = xlsmName & ".xlsm"
    ThisWorkbook.SaveCopyAs _
    Filename:=xlsmName
    

End Sub


