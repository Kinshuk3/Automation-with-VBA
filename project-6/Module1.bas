Attribute VB_Name = "Module1"

Public Sub ImportTextFile()
    Dim openFiles() As Variant
    Dim TextFile As Workbook
    Dim i As Integer
    
    openFiles = GetFiles
    Application.ScreenUpdating = False
    
    For i = 1 To Application.CountA(openFiles)
        Set TextFile = Workbooks.Open(openFiles(i))
        
        TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
        
        Workbooks(1).Activate
        Workbooks(1).Worksheets.Add
        ActiveSheet.Paste
        ActiveSheet.Name = TextFile.Name ' rename the worksheet
        
        Application.CutCopyMode = False ' clears the clipboard
        TextFile.Close
    Next i
    
    Application.ScreenUpdating = True
End Sub

Public Function GetFiles() As Variant
    GetFiles = Application.GetOpenFilename(Title:="Select File(s) to Import", MultiSelect:=True)
End Function

