Attribute VB_Name = "OOPGlobalPPFunctions"
Option Explicit
'GlobalPPFunction Methods:
'AddReferences
'DetermineReferences
'ConnectPP() As Object
'RemoveReferences

Public PPOldNew As Boolean  'False for old ProxyPlus - True for new ProxyPlus
Public pp As Object ' This object will be instantiated as either an OldReflections object, or a NewReflections object.  Which type is determined at run-time

Public Sub AddReferences()
Dim x As Integer
Dim i As Integer
Dim theRef As Variant
On Error GoTo ErrorHandling

'Add an "AddFromGuid" statement for each reference that is displayed in the "DetermineReferences" procedure.
'ex/
'    ActiveWorkbook.VBProject.References.AddFromGuid "--Replace with correct GUID numbers--", 0, 0

 'Remove any missing references
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.isbroken = True Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i




PPOldNew = False

    On Error Resume Next
    'Micro Focus Reflections GUID
    ActiveWorkbook.VBProject.References.AddFromGuid "{ECF246D9-E871-11D2-8CC2-00C04F72C0ED}", 0, 0
    'Attachmate_Reflection_Objects
    ActiveWorkbook.VBProject.References.AddFromGuid "{6857A7F4-4CDE-43F2-A7B1-CB18BA8AA35F}", 0, 0
    'Attachmate_Reflection_Objects_Emulation_IbmHosts
    ActiveWorkbook.VBProject.References.AddFromGuid "{0D5D17DF-B511-4BE5-9CD0-10DE1385229D}", 0, 0
    'Attachmate_Reflection_Objects_Framework
    ActiveWorkbook.VBProject.References.AddFromGuid "{88EC0C50-0C86-4679-B27D-63B2FCF1C6F4}", 0, 0

'Check if one of the Attachmate references have been selected. If so, this means the user has the New version of ProxyPlus
For x = 1 To ActiveWorkbook.VBProject.References.Count
    If ActiveWorkbook.VBProject.References(x).Name = "Attachmate_Reflection_Objects" Then PPOldNew = True
Next x
    
If PPOldNew = False Then
'GUID for Reflections 32 bit GUID
     ActiveWorkbook.VBProject.References.AddFromGuid "{13298D80-5585-101C-9596-040224007802}", 0, 0
    'Reflection Reference
End If


Application.DisplayAlerts = False

Exit Sub
ErrorHandling:
Resume Next
End Sub

Public Sub DetermineReferences()
Dim Reference As Object
'Run this procedure one time in order to determine what references will need to be added to the "AddReferences" procedure.
'For each reference needed for the macro, add an "AddFromGuid" statement in the "AddReference" procedure.
'Then call the "AddReference" in the Workbook Open event.  The ClosePP procedure will remove the references
'After running this procedure, the reference attributes will be displayed in the Immediate window.



On Error GoTo ErrorHandling


    For Each Reference In ThisWorkbook.VBProject.References
        Debug.Print Reference.Name
        Debug.Print Reference.GUID
    Next Reference



Set Reference = Nothing
Exit Sub

ErrorHandling:
Resume Next

Set Reference = Nothing



End Sub
Public Function ConnectPP() As Object

Dim x As Integer


PPOldNew = False
For x = 1 To ActiveWorkbook.VBProject.References.Count
    If ActiveWorkbook.VBProject.References(x).Name = "Attachmate_Reflection_Objects" Then PPOldNew = True
Next x


'since patch was installed, the NewOldReflections is not needed for 32 bit reflections.

        'If PPOldNew = True Then
        '64 bit reflections
           ' Set pp = New NewReflections
            
        'ElseIf Len(Environ("ProgramW6432")) > 0 Then
        '64 bit version OS and  32 bit reflections
            'Set pp = New NewOldReflections
        'Else
        '32 bit OS with 32 bit reflections
            'Set pp = New OldReflections
        'End If
        
        'Set ConnectPP = pp

    If PPOldNew = True Then
    Set ConnectPP = GetObject("RIBM")
       GoTo leave
     Else
     Set ConnectPP = GetObject(, "ReflectionIBM.Session")
     End If

leave:

End Function
Sub RemoveReferences()
Dim Reference As Object

For Each Reference In ThisWorkbook.VBProject.References
    If Reference.Name <> "VBA" And Reference.Name <> "Excel" And Reference.Name <> "Stdole" And Reference.Name <> "MSForms" And Reference.Name <> "DAO" And Reference.Name <> "Scripting" And Reference.Name <> "Office" Then
        ThisWorkbook.VBProject.References.Remove Reference
    End If
    Next Reference

Application.DisplayAlerts = True

Set Reference = Nothing
End Sub



