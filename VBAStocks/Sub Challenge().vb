Sub Challenge()
        
        'CHALLENGE : our solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:
    For Each ws In Worksheets
    
        'Setting the variables for the challenge
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Tot_Vol As Double
        Dim Ticker As String
        Dim a, b, c As Integer
        
        
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_Tot_Vol = 0
        
        'Create the new table
        ws.Cells(1, 18) = "Ticker"
        ws.Cells(1, 19) = "Value"
        ws.Cells(2, 17) = "Greatest % Increase"
        ws.Cells(3, 17) = "Greatest % Decrease"
        ws.Cells(4, 17) = "Greatest total volume"
        ws.Cells(2, 19) = "0.00%"
        ws.Cells(3, 19) = "0.00%"
              
        'Define Greatest % Increase
        ws.Cells(2, 19) = WorksheetFunction.Max(ws.Range("M:M"))
        Greatest_Increase = ws.Cells(2, 19)
        
        'Match Greatest % Increase Ticker
        a = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M:M")), ws.Range("M:M"), 0)
        ws.Cells(2, 18) = ws.Cells(a, 11)
        
        'Define Greatest % Decrease
        ws.Cells(3, 19) = WorksheetFunction.Min(ws.Range("M:M"))
        Greatest_Decrease = ws.Cells(3, 19)
        
        'Match Greatest % Decrease Ticker
        b = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("M:M")), ws.Range("M:M"), 0)
        ws.Cells(3, 18) = ws.Cells(b, 11)
        
        'Define Greatest Total Volume
        ws.Cells(4, 19) = WorksheetFunction.Max(ws.Range("N:N"))
        Greatest_Tot_Vol = ws.Cells(4, 19)
        
        'Match Greatest & Decrease Ticker
        c = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("N:N")), ws.Range("N:N"), 0)
        ws.Cells(4, 18) = ws.Cells(c, 11)
    Next ws
    
End Sub
