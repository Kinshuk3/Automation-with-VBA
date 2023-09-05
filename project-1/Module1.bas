Attribute VB_Name = "Module1"

' sorts the list by division column
Public Sub DivisionSort()
    
    Columns("A:F").Sort key1:=Range("A2"), order1:=xlDescending, Header:=xlYes
     
End Sub


' sorts the list by category column
Public Sub CategorySort()
    
    Columns("A:F").Sort key1:=Range("B2"), order1:=xlDescending, Header:=xlYes
     
End Sub


' sorts the list by total column
Public Sub TotalSort()
    
    Columns("A:F").Sort key1:=Range("F2"), order1:=xlDescending, Header:=xlYes
     
End Sub


'get the input from user on what to sort
Public Sub UserSortInput()
    
    Dim sortOrder As Integer
    Dim promtMSG As String
    Dim ErrorInputNum As Integer
    
    On Error GoTo errHandler
    
    promtMSG = "How would you want to sort the list" & vbCrLf & _
    "1 - Sort by Division" & vbCrLf & _
    "2 - Sort by Category" & vbCrLf & _
    "3 - Sort by Total"
    
    sortOrder = InputBox(promtMSG, "User Input")
    
    If sortOrder = 1 Then
        DivisionSort
    ElseIf sortOrder = 2 Then
        CategorySort
    ElseIf sortOrder = 3 Then
        TotalSort
    Else
errHandler:
        ErrorInputNum = MsgBox("Invalid input. Please try again!", vbYesNo)
        
        If ErrorInputNum = 6 Then
            UserSortInput
        End If
    End If
End Sub
