Attribute VB_Name = "Module4"
Sub addVIF()
num = 1
    For Each cell In ActiveSheet.Range("DESENC1", "DESENC9")
        If cell = "VIF" Then
            GoTo ENDSUB:
        End If
        If cell = "" Then
            ActiveSheet.Range("DESENC" & num) = "VIF"
            ActiveSheet.Range("DESLNG" & num) = "E/F"
            GoTo ENDSUB:
        End If
        num = num + 1
    Next cell
ENDSUB:
End Sub


