VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "Welcome User to Report Form"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddWorksheet_Click()
    Dim tryAgain As Integer
    
    On Error GoTo errHandler
    
    Worksheets().Add before:=Worksheets(1)
    
    ActiveSheet.Name = InputBox("Please enter a new worksheet name", "Add name")
    
    Exit Sub
errHandler:
    
    tryAgain = MsgBox("Invalid Name", vbYesNo)
    
    If tryAgain = 6 Then
        Application.DisplayAlerts = False
        ActiveSheet.Delete
        btnAddWorksheet_Click
    Else
        ' turn off the alert of deleting the worksheet if the user selects no
        Application.DisplayAlerts = False
        ' delete the invalid worksheet
        ActiveSheet.Delete
    End If
End Sub

Private Sub btnRunReport_Click()
    LoopYearlyReport
End Sub

Private Sub cboWhichSheet_Change()
    Worksheets(Me.cboWhichSheet.Value).Select
End Sub

Private Sub UserForm_Click()
    MsgBox ("hello world!")
End Sub

Private Sub UserForm_Initialize()

    Dim i As Integer
    i = 1
    Do While i <= Worksheets.Count()
        Me.cboWhichSheet.AddItem Worksheets(i).Name
        i = i + 1
    Loop
    
End Sub
