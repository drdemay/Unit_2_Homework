Sub Ticker_Loop()
' declare variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim totalstockvolume As Double
    
    Dim row_counter As Long
    Dim column_counter As Long
    
    Dim summary_table_row As Long
    
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
    ' Hold total per credit card brand
    totalstockvolume = 0
    
    ' Add column headers for Ticker and Total Stock Volume
    Range("I1") = "Ticker"
    Range("J1") = "Total Stock Volume"
          
    ' Set starting locations for summary table
    summary_table_row = 2
        
    For row_counter = 2 To Cells(Rows.Count, "A").End(xlUp).Row

        ' loop through and find unique values for rows
        If Cells(row_counter + 1, 1).Value <> Cells(row_counter, 1) Then
        
        ' Set the ticker value
        ticker = Cells(row_counter, 1).Value
        
        ' Print unique values for rows
        Range("I" & summary_table_row).Value = ticker
        
        ' Print totalstockvolume
        Range("J" & summary_table_row).Value = totalstockvolume + Cells(row_counter, 7).Value
        
        ' Go to the next row
        summary_table_row = summary_table_row + 1
        
        ' Reset the total
        totalstockvolume = 0
        
        Else
          
        ' loop through stock volumns and add up total - once we have iterated through each column, print total stock volumn in column j
        totalstockvolume = totalstockvolume + Cells(row_counter, 7).Value
            
        End If
        
    Next row_counter

Next ws
    
End Sub