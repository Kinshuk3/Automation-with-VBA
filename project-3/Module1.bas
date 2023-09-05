Attribute VB_Name = "Module1"

Public Sub AutomateTotalSum()
    Dim lastCell As String
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        Worksheets(ws.Name).Select
        
        Range("F2").Select
        Selection.End(xlDown).Select
        
        lastCell = ActiveCell.Address(False, False)
        
        ActiveCell.Offset(1, 0).Select
        
        ActiveCell.Value = "=sum(F2:" & lastCell & ")"
    Next ws
End Sub
