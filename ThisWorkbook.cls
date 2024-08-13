VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub SplitSalesList()

    Application.ScreenUpdating = False
    
    Dim r As Range
    Dim CellRange As Range
    Dim mDate As String
    Dim ws As Worksheet
    
    'set a reference to the cells to loop over
    Set CellRange = Sheet1.Range("A2", Sheet1.Range("A2").End(xlDown))
    
    
    'loop over the range of dates
    For Each r In CellRange
        
        'capture the current dino's mDate
        mDate = Format(r.Offset(0, 4).Value, "MM")
        
        'copy the current row
        Range(r, r.End(xlToRight)).Copy
        
        'go to the correct worksheet
        CreateSheet (mDate)
        Worksheets(mDate).Select
        
        'find the next blank row
        Range("A1048576").End(xlUp).Offset(1, 0).Select
        
        'paste the copied data
        ActiveCell.PasteSpecial
        
    Next r
    
    'copy the column headings from Dino List
    Sheet1.Range("A1").EntireRow.Copy
    
    'loop over the worksheets collection
    For Each ws In Worksheets
    
        'check that we don't include the Dino List sheet
        If Not ws Is Sheet1 Then
        
            'paste the copied headers in
            ws.Range("A1").PasteSpecial
            
            'change the column widths to fit the data
            ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
            
        End If
        
    Next ws
    
    
End Sub

Function CreateSheet(sheetName As String)

    If CheckSheetExists(sheetName) = False Then
        Worksheets.Add.Name = sheetName
    End If

End Function

Function CheckSheetExists(sheetName As String) As Boolean
    Dim sht As Worksheet

    On Error Resume Next
    Set sht = Worksheets(sheetName)
    On Error GoTo 0
    CheckSheetExists = Not sht Is Nothing
End Function


