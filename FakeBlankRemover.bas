Attribute VB_Name = "FakeBlankRemover"
Option Explicit

Sub ClearFakeBlanksInActiveWorkbook()
    
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Sheets
        
        ClearFakeBlanksInWorksheet ws
        
    Next ws
    
End Sub

Sub ClearFakeBlanksInActiveSheet()
    
    ClearFakeBlanksInWorksheet ActiveSheet
    
End Sub

Sub ClearFakeBlanksInWorksheet(ws As Worksheet)
    
    ClearFakeBlanksInRange ws.UsedRange
    
End Sub

Sub ClearFakeBlanksInSelection()
    
    ClearFakeBlanksInRange Intersect(Selection, Selection.Parent.UsedRange)
    
End Sub
Sub ClearFakeBlanksInRange(tr As Range)
    
    Dim s As String
    Dim c
    Dim i As Long
    
    i = 0
    
    For Each c In tr.Cells
        
        i = i + 1
        
        Application.StatusBar = "Processing - wb: " & c.Parent.Parent.Name & ", sheet: " & c.Parent.Index & "/" & c.Parent.Parent.Sheets.Count & ", cell: " & i & "/" & tr.Cells.Count
        
        s = Trim(c.Value2)
        
        If Len(s) = 0 Then c.ClearContents
        
    Next c
    
    Application.StatusBar = False
    
End Sub
