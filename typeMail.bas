Attribute VB_Name = "Module8"
Function typeMail(holders As Long, weight As String)
    Dim mail As String
    If holders >= 1000 And weight <> "Over Weight" Then
        mail = "Incentive Lettermail"
    ElseIf holders < 1000 And weight <> "Over Weight" Then
        mail = "Lettermail"
    Else
        mail = "Parcel Post"
    End If
    typeMail = mail
        
End Function

