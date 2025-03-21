VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USERDETAILS 
   Caption         =   "USER DETAILS FORM"
   ClientHeight    =   9492.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7692
   OleObjectBlob   =   "USERDETAILS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USERDETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub txtMV_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEngine_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtChasis_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtLan_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtORCR_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDepartment_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtBrand_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtCondition_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDepot_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPlate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtInsurance_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtCTPLD_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(Me.txtCTPLD.Value) Then
        MsgBox "Please enter a valid date for Registration Date.", vbExclamation
        Cancel = True
    Else
        Me.txtCTPLD.Value = Format(CDate(Me.txtCTPLD.Value), "MM/DD/YYYY")
    End If
End Sub
Private Sub txtCompreD_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(Me.txtCompreD.Value) Then
        MsgBox "Please enter a valid date for Registration Date.", vbExclamation
        Cancel = True
    Else
        Me.txtCompreD.Value = Format(CDate(Me.txtCompreD.Value), "MM/DD/YYYY")
    End If
End Sub
Private Sub txtINSD_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(Me.txtINSD.Value) Then
        MsgBox "Please enter a valid date for Registration Date.", vbExclamation
        Cancel = True
    Else
        Me.txtINSD.Value = Format(CDate(Me.txtINSD.Value), "MM/DD/YYYY")
    End If
End Sub
Private Sub txtEffect_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(Me.txtEffect.Value) Then
        MsgBox "Please enter a valid date for Registration Date.", vbExclamation
        Cancel = True
    Else
        Me.txtEffect.Value = Format(CDate(Me.txtEffect.Value), "MM/DD/YYYY")
    End If
End Sub

Private Sub txtExpD_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(Me.txtExpD.Value) Then
        MsgBox "Please enter a valid date for Expiration Date.", vbExclamation
        Cancel = True
    Else
        Me.txtExpD.Value = Format(CDate(Me.txtExpD.Value), "MM/DD/YYYY")
    End If
End Sub

Private Sub txtDop_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(Me.txtDop.Value) Then
        MsgBox "Please enter a valid date for DOP.", vbExclamation
        Cancel = True
    Else
        Me.txtDop.Value = Format(CDate(Me.txtDop.Value), "MM/DD/YYYY")
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Populate the ComboBox with company names
    With Me.cmbCompany
        .AddItem "LDI"
        .AddItem "FEI"
        .AddItem "LCPI"
    End With
    With Me.cmbType
        .AddItem "FOUR WHEELS"
        .AddItem "TWO WHEELS"
    End With
End Sub

Private Sub btnSubmit_Click()
    Dim ws As Worksheet
    Dim logWs As Worksheet
    Dim vacantWs As Worksheet
    Dim nextRow As Long
    Dim vacantNextRow As Long
    Dim plateExists As Boolean
    Dim cell As Range
    Dim logRow As Long
    Dim actionType As String
    Const password As String = "CORPLAN"
    
    On Error GoTo ErrorHandler
    
    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets("UserDetails")
    Set logWs = ThisWorkbook.Sheets("ActionLog")
    Set vacantWs = ThisWorkbook.Sheets("VACANTS")
    ws.Unprotect password:=password
    logWs.Unprotect password:=password
    vacantWs.Unprotect password:=password
    
    ' Validate that a company and vehicle type are selected
    If Me.cmbCompany.Value = "" Then
        MsgBox "Please select a company."
        Exit Sub
    End If
    If Me.cmbType.Value = "" Then
        MsgBox "Please select Vehicle Type."
        Exit Sub
    End If
    
    ' Set txtName to "VACANT" if it is empty
    If Trim(Me.txtName.Value) = "" Then
        Me.txtName.Value = "VACANT"
    End If
    
    ' Check if the plate number already exists in the Vacant sheet
    plateExists = False
    For Each cell In vacantWs.Range("F1:F" & vacantWs.Cells(vacantWs.Rows.Count, 5).End(xlUp).Row) ' Assuming Plate is in column E
        If StrComp(cell.Value, Me.txtPlate.Value, vbTextCompare) = 0 Then
            plateExists = True
            Exit For
        End If
    Next cell
    
    ' If txtName is "VACANT", copy the details to the Vacant sheet
    If Me.txtName.Value = "VACANT" Then
        If Not plateExists Then
            vacantNextRow = vacantWs.Cells(vacantWs.Rows.Count, 1).End(xlUp).Row + 1
            
            With vacantWs
                .Cells(vacantNextRow, 1).Value = Me.txtID.Value
                .Cells(vacantNextRow, 2).Value = Me.txtName.Value
                .Cells(vacantNextRow, 3).Value = Me.cmbCompany.Value
                .Cells(vacantNextRow, 4).Value = Me.txtDepartment.Value
                .Cells(vacantNextRow, 5).Value = Me.txtDepot.Value
                .Cells(vacantNextRow, 6).Value = Me.txtPlate.Value
                .Cells(vacantNextRow, 7).Value = Me.cmbType.Value
                .Cells(vacantNextRow, 8).Value = Me.txtBrand.Value
                .Cells(vacantNextRow, 9).Value = Me.txtYear.Value
                .Cells(vacantNextRow, 10).Value = Me.txtMV.Value
                .Cells(vacantNextRow, 11).Value = Me.txtRegSched.Value
                .Cells(vacantNextRow, 13).Value = Me.txtEngine.Value
                .Cells(vacantNextRow, 14).Value = Me.txtChasis.Value
                .Cells(vacantNextRow, 15).Value = Me.txtAmount.Value
                .Cells(vacantNextRow, 16).Value = Me.txtCondition.Value
                .Cells(vacantNextRow, 17).Value = Me.txtDop.Value
                .Cells(vacantNextRow, 18).Value = Me.txtORCR.Value
                .Cells(vacantNextRow, 19).Value = Me.txtLan.Value
                .Cells(vacantNextRow, 20).Value = Me.txtInsurance.Value
                .Cells(vacantNextRow, 21).Value = Me.txtINSD.Value
                .Cells(vacantNextRow, 22).Value = Me.txtCTPL.Value
                .Cells(vacantNextRow, 23).Value = Me.txtCTPLD.Value
                .Cells(vacantNextRow, 24).Value = Me.txtCompre.Value
                .Cells(vacantNextRow, 25).Value = Me.txtCompreD.Value
                .Cells(vacantNextRow, 26).Value = Me.txtEffect.Value
            End With
            
            MsgBox "Details copied to Vacant sheet successfully!"
        Else
            MsgBox "The plate number '" & Me.txtPlate.Value & "' already exists in the Vacant sheet. No details copied."
            Exit Sub
        End If
    Else
        ' If txtName is not "VACANT", you can choose to handle it differently
        ' For example, you can insert the details into the User Details sheet
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        
                With ws
            .Cells(nextRow, 1).Value = Me.txtID.Value
            .Cells(nextRow, 2).Value = Me.txtName.Value
            .Cells(nextRow, 3).Value = Me.cmbCompany.Value
            .Cells(nextRow, 4).Value = Me.txtDepartment.Value
            .Cells(nextRow, 5).Value = Me.txtDepot.Value
            .Cells(nextRow, 6).Value = Me.txtPlate.Value
            .Cells(nextRow, 7).Value = Me.cmbType.Value
            .Cells(nextRow, 8).Value = Me.txtBrand.Value
            .Cells(nextRow, 9).Value = Me.txtYear.Value
            .Cells(nextRow, 10).Value = Me.txtMV.Value
            .Cells(nextRow, 11).Value = Me.txtRegSched.Value
            .Cells(nextRow, 13).Value = Me.txtEngine.Value
            .Cells(nextRow, 14).Value = Me.txtChasis.Value
            .Cells(nextRow, 15).Value = Me.txtAmount.Value
            .Cells(nextRow, 16).Value = Me.txtCondition.Value
            .Cells(nextRow, 17).Value = Me.txtDop.Value
            .Cells(nextRow, 18).Value = Me.txtORCR.Value
            .Cells(nextRow, 19).Value = Me.txtLan.Value
            .Cells(nextRow, 20).Value = Me.txtInsurance.Value
            .Cells(nextRow, 21).Value = Me.txtINSD.Value
            .Cells(nextRow, 22).Value = Me.txtCTPL.Value
            .Cells(nextRow, 23).Value = Me.txtCTPLD.Value
            .Cells(nextRow, 24).Value = Me.txtCompre.Value
            .Cells(nextRow, 25).Value = Me.txtCompreD.Value
            .Cells(nextRow, 26).Value = Me.txtEffect.Value
        End With
        
        MsgBox "Details added to User Details sheet successfully!"
    End If

    ' Log the action in the ActionLog sheet
    logRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    With logWs
        .Cells(logRow, 1).Value = Now ' Timestamp
        .Cells(logRow, 2).Value = "ADD" ' Action (Add)
        .Cells(logRow, 3).Value = Me.txtID.Value ' Name
        .Cells(logRow, 4).Value = Me.txtName.Value ' Name
        .Cells(logRow, 5).Value = Me.cmbCompany.Value ' Company
        .Cells(logRow, 6).Value = Me.txtDepartment.Value ' Department
        .Cells(logRow, 7).Value = Me.txtPlate.Value ' Plate
        .Cells(logRow, 8).Value = Me.txtBrand.Value ' Brand
        .Cells(logRow, 9).Value = Me.txtEffect.Value ' Year
    End With

    ' Optionally, protect the worksheets to prevent editing
    ws.Protect password:=password, UserInterfaceOnly:=True
    logWs.Protect password:=password, UserInterfaceOnly:=True
    vacantWs.Protect password:=password, UserInterfaceOnly:=True

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
