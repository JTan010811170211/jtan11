VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VACANTTRANS 
   Caption         =   "TRANSFER  "
   ClientHeight    =   6324
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "VACANTTRANS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VACANTTRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnTransfer_Click()
    Dim ws As Worksheet
    Dim wsLog As Worksheet
    Dim wsUserDetails As Worksheet
    Dim plateToFind As String
    Dim newEmployee As String
    Dim foundCell As Range
    Dim currentOwner As String
    Dim newID As String
    Dim transferDate As Date
    Dim newDepartment As String
    Dim newEffect As String

    On Error GoTo ErrHandler

    ' Set worksheets
    Set ws = ThisWorkbook.Sheets("VACANTS")
    Set wsLog = ThisWorkbook.Sheets("TransferLog")
    Set wsUserDetails = ThisWorkbook.Sheets("UserDetails")
    ws.Unprotect password:="CORPLAN"
    wsLog.Unprotect password:="CORPLAN"
    wsUserDetails.Unprotect password:="CORPLAN"

    ' Get input values
    plateToFind = Trim(Me.txtPlate.Value)
    newEmployee = Trim(Me.txtName.Value)
    newDepartment = Trim(Me.txtDepartment.Value)
    newID = Trim(Me.txtID.Value)
    newEffect = Trim(Me.txtEffect.Value)

    ' Find the row with the specified plate number
    Set foundCell = ws.Range("F:F").Find(What:=plateToFind, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

    If Not foundCell Is Nothing Then
        ' Check if the name already exists in UserDetails
        Dim lastRow As Long
        lastRow = wsUserDetails.Cells(wsUserDetails.Rows.Count, 1).End(xlUp).Row
        Dim checkRange As Range
        Set checkRange = wsUserDetails.Range("B2:B" & lastRow)

        If Not Application.WorksheetFunction.CountIf(checkRange, foundCell.Offset(0, 0).Value) = 0 Then
            MsgBox "The name '" & foundCell.Offset(0, 0).Value & "' already exists in UserDetails.", vbExclamation
            Exit Sub
        End If
        If newID <> "" Then
            foundCell.Offset(0, -5).Value = newID
        End If
        ' Update the company and department if changed
        If newDepartment <> "" Then
            foundCell.Offset(0, -2).Value = newDepartment
        End If
        If newEffect <> "" Then
            foundCell.Offset(0, 20).Value = newEffect
        End If

        ' Prompt for confirmation
        If MsgBox("Do you really want to transfer ownership of '" & currentOwner & "' to '" & newEmployee & "'?", vbQuestion + vbYesNo) = vbYes Then
            ' Update the ownership to the new employee
            foundCell.Offset(0, -4).Value = newEmployee

            ' Record the transfer log
            transferDate = Now
            wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = transferDate
            wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 1).Value = newID
            wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 2).Value = currentOwner
            wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 4).Value = newEmployee
            wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Offset(0, 5).Value = newDepartment

            ' Move the entire row to UserDetails
            foundCell.EntireRow.Copy wsUserDetails.Cells(wsUserDetails.Rows.Count, 1).End(xlUp).Offset(1, 0)
            foundCell.EntireRow.Delete

            MsgBox "Ownership transferred to " & newEmployee & ".", vbInformation
        Else
            MsgBox "Transfer cancelled.", vbInformation
        End If
    Else
        MsgBox "No entry found for the specified plate number.", vbExclamation
    End If
    

    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    ws.Protect password:="CORPLAN", UserInterfaceOnly:=True
    wsLog.Protect password:="CORPLAN", UserInterfaceOnly:=True
    wsUserDetails.Protect password:="CORPLAN", UserInterfaceOnly:=True
End Sub
