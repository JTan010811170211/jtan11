VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TRANSFER 
   Caption         =   "UserForm1"
   ClientHeight    =   6348
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4668
   OleObjectBlob   =   "TRANSFER.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TRANSFER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnTransfer_Click()
     Dim ws As Worksheet
    Dim wsLog As Worksheet
    Dim wsVacant As Worksheet
    Dim nameToFind As String
    Dim plateToFind As String
    Dim newEmployee As String
    Dim foundCell As Range
    Dim currentOwner As String
    Dim transferDate As Date
    Dim newDepartment As String
    Dim nameExists As Boolean
    Dim cell As Range
    Dim lastRow As Long
    
    ' Set the worksheets
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("UserDetails")
    Set wsLog = ThisWorkbook.Sheets("TransferLog")
    Set wsVacant = ThisWorkbook.Sheets("VACANTS")
    On Error GoTo 0
    
    If ws Is Nothing Or wsLog Is Nothing Or wsVacant Is Nothing Then
        MsgBox "One or more specified sheets do not exist.", vbCritical
        Exit Sub
    End If
    
    ws.Unprotect password:="CORPLAN"
    wsLog.Unprotect password:="CORPLAN"
    wsVacant.Unprotect password:="CORPLAN"
   
    ' Get the name and plate number to find, trimming any extra spaces
    nameToFind = Trim(Me.txtName.Value)
    newEmployee = Trim(Me.txtNewEmployee.Value)
    newDepartment = Trim(Me.txtDepartment.Value) ' Added .Value to ensure it's read correctly
    
    ' Debugging: Print the values to the Immediate Window (Ctrl + G to view)
    Debug.Print "Searching for Name: '" & nameToFind & "'"
    Debug.Print "Searching for Plate: '" & plateToFind & "'"
    Debug.Print "New Employee: '" & newEmployee & "'"
    Debug.Print "Department: '" & newDepartment & "'"
    
    ' Check if both fields are empty
    If nameToFind = "" And plateToFind = "" Then
        MsgBox "Please enter a name or a plate number to transfer ownership."
        Exit Sub
    End If
    
    ' Search for the name first
    If nameToFind <> "" Then
        Set foundCell = ws.Range("B:B").Find(What:=nameToFind, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    End If
    
    ' If name is not found, search for the plate number
    If foundCell Is Nothing And plateToFind <> "" Then
        Set foundCell = ws.Range("F:F").Find(What:=plateToFind, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) ' Assuming plate number is in column E
    End If
    
    ' Check if the entry was found
    If Not foundCell Is Nothing Then
        ' Get the current owner's name
        currentOwner = foundCell.Offset(0, 0).Value
         
        If newDepartment <> "" Then
            foundCell.Offset(0, 3).Value = newDepartment
        End If
        
        ' Check if newEmployee is empty or "VACANT"
        If newEmployee = "" Then
            newEmployee = "VACANT"
        End If
        
        ' Prompt for confirmation
        If MsgBox("Do you really want to transfer ownership of '" & currentOwner & "' to '" & newEmployee & "'?", vbQuestion + vbYesNo) = vbYes Then
            ' Update the ownership to the new employee
            foundCell.Offset(0, 0).Value = newEmployee
                    
            ' If newEmployee is "VACANT", move the entire row to the VACANTS sheet
            If newEmployee = "VACANT" Then
                lastRow = wsVacant.Cells(wsVacant.Rows.Count, 1).End(xlUp).Row + 1
                ws.Rows(foundCell.Row).Copy Destination:=wsVacant.Rows(lastRow)
                ws.Rows(foundCell.Row).Delete
                                ' Log the transfer if necessary
                transferDate = Now
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = transferDate
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 1).Value = currentOwner
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 2).Value = newEmployee
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 4).Value = newDepartment
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 5).Value = Me.txtEffect.Value ' Year
                
                ' Display a confirmation message
                MsgBox "Ownership transferred to " & newEmployee & ".", vbInformation
            Else
                ' If the new employee is not "VACANT", just log the transfer
                transferDate = Now
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = transferDate
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 1).Value = currentOwner
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 2).Value = newEmployee
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 4).Value = newDepartment
                wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 5).Value = Me.txtEffect.Value ' Year
                
                ' Display a confirmation message
                MsgBox "Ownership transferred to " & newEmployee & ".", vbInformation
            End If
        Else
            MsgBox "Transfer cancelled.", vbInformation
        End If
    Else
        MsgBox "No entry found for the specified name or plate number.", vbExclamation
    End If
    
    ' Protect the worksheets again
    ws.Protect password:="CORPLAN", UserInterfaceOnly:=True
    wsLog.Protect password:="CORPLAN", UserInterfaceOnly:=True
    wsVacant.Protect password:="CORPLAN", UserInterfaceOnly:=True
End Sub
