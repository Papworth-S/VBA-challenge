Attribute VB_Name = "Module2"
Sub Reset()

    For Each ws In Worksheets
    
        ws.Range("I1:L100").Value = ""
        ws.Range("I1:L100").Interior.ColorIndex = 0
        ws.Range("I1:L100").NumberFormat = "General"
        ws.Range("O1:Q4").Value = ""
        ws.Range("O1:Q4").NumberFormat = "General"
        ws.Columns("I:Q").ColumnWidth = 8
    
    Next ws
    
End Sub
