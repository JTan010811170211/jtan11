VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EDITUSER 
   Caption         =   "EDIT USER FORM"
   ClientHeight    =   7908
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5712
   OleObjectBlob   =   "EDITUSER.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EDITUSER"
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

Private Sub UserForm_Initialize()
    ' Populate the ComboBox with company names
  
 
End Sub

Private Sub btnLoadData_Click()
    Dim ws As Worksheet
    Dim selectedRow As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("UserDetails")

    ' Check if a cell is selected
    If Not Application.ActiveCell Is Nothing Then
        selectedRow = Application.ActiveCell.Row

        ' Populate the text boxes with the data from the selected row
        Me.txtID.Value = ws.Cells(selectedRow, 1).Value    ' Name
        Me.txtName.Value = ws.Cells(selectedRow, 2).Value    ' Name
        'Me.cmbCompany.Value = ws.Cells(selectedRow, 3).Value ' Company
        'Me.txtDepartment.Value = ws.Cells(selectedRow, 4).Value ' Department
        'Me.txtDepot.Value = ws.Cells(selectedRow, 5).Value    ' Depot
        Me.txtPlate.Value = ws.Cells(selectedRow, 6).Value ' Company
        'Me.cmbType.Value = ws.Cells(selectedRow, 7).Value ' Department
        Me.txtBrand.Value = ws.Cells(selectedRow, 8).Value    ' Brand
       ' Me.txtYear.Value = ws.Cells(selectedRow,9).Value ' Company
       ' Me.txtMV.Value = ws.Cells(selectedRow, 10).Value ' Department
        Me.txtRegSched.Value = ws.Cells(selectedRow, 11).Value    ' Name
       ' Me.txtEngine.Value = ws.Cells(selectedRow, 13).Value ' Company
        'Me.txtChasis.Value = ws.Cells(selectedRow, 14).Value ' Department
        'Me.txtAmount.Value = ws.Cells(selectedRow, 15).Value    ' Name
        'Me.txtCondition.Value = ws.Cells(selectedRow, 16).Value ' Company
       ' Me.txtDop.Value = ws.Cells(selectedRow, 17).Value ' Department
        'Me.txtORCR.Value = ws.Cells(selectedRow, 18).Value    ' Name
        'Me.txtLan.Value = ws.Cells(selectedRow, 19).Value ' Company
        Me.txtInsurance.Value = ws.Cells(selectedRow, 20).Value ' Department
        Me.txtINSD.Value = ws.Cells(selectedRow, 21).Value ' Company
        Me.txtCTPL.Value = ws.Cells(selectedRow, 22).Value    ' Name
        Me.txtCTPLD.Value = ws.Cells(selectedRow, 23).Value    ' Name
        Me.txtCompre.Value = ws.Cells(selectedRow, 24).Value ' Company
        Me.txtCompreD.Value = ws.Cells(selectedRow, 25).Value ' Company
        'Me.txtEffect.Value = ws.Cells(selectedRow, 2).Value ' Company
            
    Else
        MsgBox "Please select a cell in the row you want to edit."
    End If
End Sub

Private Sub btnSubmit_Click()
    Dim ws As Worksheet
    Dim selectedRow As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("UserDetails")
    Set logWs = ThisWorkbook.Sheets("ActionLog")
    ws.Unprotect password:="CORPLAN"
    logWs.Unprotect password:="CORPLAN"

    ' Check if a cell is selected
    If Not Application.ActiveCell Is Nothing Then
        selectedRow = Application.ActiveCell.Row

        ' Log the changes to the ActionLog worksheet
        Dim logRow As Long
        logRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1

        ' Log the action
        logWs.Cells(logRow, 1).Value = Now() ' Timestamp
        logWs.Cells(logRow, 2).Value = "EDIT"
        ' Log each changed field individually
        logWs.Cells(logRow, 3).Value = Me.txtID.Value ' Name
       logWs.Cells(logRow, 4).Value = Me.txtName.Value ' Name
    'logWs.Cells(logRow, 5).Value = Me.cmbCompany.Value ' Company
   'logWs.Cells(logRow, 6).Value = Me.txtDepartment.Value ' Department
    'logWs.Cells(logRow, 6).Value = Me.txtDepot.Value ' Depot
    logWs.Cells(logRow, 7).Value = Me.txtPlate.Value ' Plate
    'logWs.Cells(logRow, 8).Value = Me.cmbType.Value ' Type
 
    ' Optionally, you can log more fields if needed
     logWs.Cells(logRow, 8).Value = Me.txtBrand.Value ' Brand
      'logWs.Cells(logRow, 9).Value = Me.txtEffect.Value ' ExpDate
    ' logWs.Cells(logRow, 9).Value = Me.txtYear.Value ' Year
    ' logWs.Cells(logRow, 10).Value = Me.txtMV.Value ' MV
    ' logWs.Cells(logRow, 11).Value = Me.txtRegSched.Value ' RegSched
    ' logWs.Cells(logRow, 12).Value = Me.txtEngine.Value ' Engine
    ' logWs.Cells(logRow, 13).Value = Me.txtChasis.Value ' Chasis
    ' logWs.Cells(logRow, 14).Value = Me.txtAmount.Value ' Vatable
    ' logWs.Cells(logRow, 15).Value = Me.txtCondition.Value ' Condition
    ' logWs.Cells(logRow, 16).Value = Me.txtDop.Value ' DOP
    ' logWs.Cells(logRow, 17).Value = Me.txtORCR.Value ' ORCR
    ' logWs.Cells(logRow, 18).Value = Me.txtLan.Value ' Lan
    logWs.Cells(logRow, 19).Value = Me.txtInsurance.Value ' Insurance
    logWs.Cells(logRow, 19).Value = Me.txtINSD.Value ' Insurance
    logWs.Cells(logRow, 19).Value = Me.txtCTPL.Value ' Insurance
    logWs.Cells(logRow, 19).Value = Me.txtCTPLD.Value ' Insurance
    logWs.Cells(logRow, 19).Value = Me.txtCompre.Value ' Insurance
    logWs.Cells(logRow, 19).Value = Me.txtCompreD.Value ' Insurance
    ' logWs.Cells(logRow, 20).Value = Me.txtRegD.Value ' RegDate
    ' logWs.Cells(logRow, 21).Value = Me.txtExpD.Value ' ExpDate


        ' ... (similarly for other fields)

        ' Update the data in the selected row
        ws.Cells(selectedRow, 1).Value = Me.txtID.Value    ' Name
        ws.Cells(selectedRow, 2).Value = Me.txtName.Value    ' Name
        'ws.Cells(selectedRow, 3).Value = Me.cmbCompany.Value ' Company
       ' ws.Cells(selectedRow, 4).Value = Me.txtDepartment.Value ' Department
       ' ws.Cells(selectedRow, 5).Value = Me.txtDepot.Value    ' Depot
        'ws.Cells(selectedRow, 6).Value = Me.txtPlate.Value ' Company
        'ws.Cells(selectedRow, 7).Value = Me.cmbType.Value ' Department
        'ws.Cells(selectedRow, 8).Value = Me.txtBrand.Value    ' Brand
        'ws.Cells(selectedRow, 9).Value = Me.txtYear.Value ' Company
        'ws.Cells(selectedRow, 10).Value = Me.txtMV.Value ' Department
        ws.Cells(selectedRow, 11).Value = Me.txtRegSched.Value    ' Name
        'ws.Cells(selectedRow, 13).Value = Me.txtEngine.Value ' Company
        'ws.Cells(selectedRow, 14).Value = Me.txtChasis.Value ' Department
        'ws.Cells(selectedRow, 15).Value = Me.txtAmount.Value    ' Name
        'ws.Cells(selectedRow, 16).Value = Me.txtCondition.Value ' Company
        'ws.Cells(selectedRow, 17).Value = Me.txtDop.Value ' Department
        'ws.Cells(selectedRow, 18).Value = Me.txtORCR.Value    ' ORCR
       ' ws.Cells(selectedRow, 19).Value = Me.txtLan.Value ' Lan
        'ws.Cells(selectedRow, 20).Value = Me.txtInsurance.Value ' Insurance
        ws.Cells(selectedRow, 21).Value = Me.txtINSD.Value ' ExpDate
        ws.Cells(selectedRow, 22).Value = Me.txtCTPL.Value    ' RegDate
        ws.Cells(selectedRow, 23).Value = Me.txtCTPLD.Value    ' RegDate
        ws.Cells(selectedRow, 24).Value = Me.txtCompre.Value ' ExpDate
        ws.Cells(selectedRow, 25).Value = Me.txtCompreD.Value ' ExpDate
       'ws.Cells(selectedRow, 26).Value = Me.txtEffect.Value ' ExpDate

MsgBox "Record updated successfully!"

' Clear the text boxes after submission
Me.txtID.Value = ""
Me.txtName.Value = ""
'Me.cmbCompany.Value = ""
'Me.txtDepartment.Value = ""
'Me.txtDepot.Value = ""
'Me.txtPlate.Value = ""
'Me.cmbType.Value = ""
'Me.txtBrand.Value = ""
'Me.txtYear.Value = ""
'Me.txtMV.Value = ""
Me.txtRegSched.Value = ""
'Me.txtEngine.Value = ""
'Me.txtChasis.Value = ""
'Me.txtAmount.Value = ""
'Me.txtCondition.Value = ""
'Me.txtDop.Value = ""
'Me.txtORCR.Value = ""
'Me.txtLan.Value = ""
Me.txtInsurance.Value = ""
Me.txtINSD.Value = ""
Me.txtCTPLD.Value = ""
Me.txtCompreD.Value = ""
Me.txtCTPL.Value = ""
Me.txtCompre.Value = ""
'Me.txtRegD.Value = ""
'Me.txtEffect.Value = ""

Else
    MsgBox "Please select a cell in the row you want to edit."
End If

logWs.Protect password:="CORPLAN"
ws.Protect password:="CORPLAN"
End Sub
