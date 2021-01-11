Attribute VB_Name = "Module1"
Sub VBA_Challenge()
    
   'Loop through all sheets
    Dim WS_Count As Integer
    Dim WS_Name As String
    
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For W = 1 To WS_Count
         WS_Name = (ActiveWorkbook.Worksheets(W).Name)
        Worksheets(ActiveWorkbook.Worksheets(W).Name).Activate
    
    
    
    
    
        'Placeholders
        Range("K1").Value = "Ticker"
        Range("L1").Value = "Yearly Change"
        Range("M1").Value = "Pecent Change"
        Range("N1").Value = "Total Stock Volume"
        'Set variables
        Dim Ticker_Symbol As String
        Dim Last_Row As Long
        
        Last_Row = Cells(Rows.Count, 2).End(xlUp).Row
         
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Opening_Price = 0
        Closing_Price = 0
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim TSV As Double
        Yearly_Change = 0
        Percent_Change = 0
        TSV = 0
         
         'Add summary table
         Dim Summary_Table_Row As Long
         Summary_Table_Row = 2
         
        
        Opening_Price = Cells(2, 3).Value
        
        
        'Loop through all stock information
        For i = 2 To Last_Row
        
            
            'Check for new stock ticker
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'find closing price
                Closing_Price = Cells(i, 6).Value
                
                 'Set Ticker Symbol
                Ticker_Symbol = Cells(i, 1).Value
                 
                'Set TSV to table
                 
                TSV = TSV + Cells(i, 7).Value
                 
                'Print Ticker to summary table
                Range("K" & Summary_Table_Row).Value = Ticker_Symbol
                 
                'Print TSV to summary table
                Range("N" & Summary_Table_Row).Value = TSV
                 
               
                
                'Calculate yearly change
                Yearly_Change = Closing_Price - Opening_Price
                
                'Calculate Percent Change
                If Opening_Price = 0 Then
                    Percent_Change = 0
                Else
                    
                    Percent_Change = (Closing_Price - Opening_Price) / Opening_Price * 100
                End If
                
                'Add yearly change to summary table
                Range("L" & Summary_Table_Row).Value = Yearly_Change
                
                'Format Color to show positive/negative
                If Yearly_Change >= 0 Then
                    Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                Else
                    Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                End If
                
                
                'Add Percent Change to summart table
                Range("M" & Summary_Table_Row).Value = Percent_Change
                
                'find opening price in ticker
                Opening_Price = Cells(i + 1, 3).Value
                
                 'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                 
                'Reset the TSV
                TSV = 0
            
            Else
                TSV = TSV + Cells(i, 7).Value
            
            End If
        
        Next i
    
    Next W
    
End Sub
