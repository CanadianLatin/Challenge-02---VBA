

Sub Stock_info()

'Loop thought all Sheets
For Each ws In Worksheets
    
    'Set Variable to  store the Ticker
    Dim Ticker As String
    
    'set variable to store the Total_Stock_Volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    'Variable that keeps track of the Ticker in the result table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'set the open value and close value variables
    Dim Open_Value As Double
    Dim Close_Value As Double
    
    Open_Value = Close_Value = 0

    'Set the summary table headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Toatl Stock Volume"
    
    ' Set numrows = number of rows of data.
    'for the following 3 lines of code I use help from google and I found it here: https://www.extendoffice.com/documents/excel/4438-excel-loop-until-blank.html
    Dim x As Integer
    NumRows = ws.Range("A1", ws.Range("A1").End(xlDown)).Rows.Count
    
    
    'Hard code the first Open Value as it is skiped in the for
    Open_Value = ws.Cells(2, 3).Value
    
    'Loop though all the Tickers until is empty
    For I = 2 To NumRows
        
        'look for the same ticker until gets to a new one. If the next cell is a DIFFERENT Ticker
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
                 'PRINT the ticker name
                 Ticker = ws.Cells(I, 1).Value
                 
                 'Assigns the Close Value
                 Close_Value = ws.Cells(I, 6).Value
             
                 'Adds value to the "Total Stock Volume"
                 Total_Stock_Volume = Total_Stock_Volume + ws.Cells(I, 7).Value
                 
                 'Add ticker to the Summary_Table
                 ws.Range("I" & Summary_Table_Row).Value = Ticker
                 
                 'PRINT "Yearly Change"
                 ws.Range("J" & Summary_Table_Row).Value = Close_Value - Open_Value
                 
                 'Condicional Formating "Yearly Change"
                 If (Close_Value - Open_Value) > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4 'Green
                 Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3 'Red
                 End If
                 
                'PRINT "Porcentage Change"
                ws.Range("K" & Summary_Table_Row).Value = (Close_Value / Open_Value) - 1
                
                'Condicional Formating "Porcentage Change"
                 If (Close_Value - Open_Value) > 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4 'Green
                 Else
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3 'Red
                 End If
                 
                 'PRINT "Total Stock Volume"
                 ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                'Gets ready for the next value summary table
                 Summary_Table_Row = Summary_Table_Row + 1
                 
                 'Reset Variables
                 Total_Stock_Volume = 0
                 Open_Value = ws.Cells(I + 1, 3).Value
         
            'if the next row is the SAME Ticker
            Else
                    'PRINT value to the "Total Stock Volume"
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(I, 7).Value

            End If
        
        Next I
        
        
        'Variablesto count the minimum and maximum in an array
        Dim Min As Double
        Dim Max As Double
        Dim CountMin As Integer
        Dim CountMax As Integer
        
        'Set the Bounus table. Headers and First Column
          ws.Cells(1, 17).Value = "Ticker"
          ws.Cells(1, 18).Value = "Value"
          ws.Cells(2, 16).Value = "Greatest % Increase: "
          ws.Cells(3, 16).Value = "Greatest % Decrease: "
          ws.Cells(4, 16).Value = "Greatest Total Volume: "
            
            
          'Greatest %Increase:
          Max = WorksheetFunction.Max(ws.Range("K:K"))
          CountMax = WorksheetFunction.Match(Max, ws.Range("K:K"), 0)
          ws.Cells(2, 18).Value = Max
          ws.Cells(2, 17).Value = ws.Range("I" & CountMax).Value

          'Greatest %Decrease:
          Min = WorksheetFunction.Min(ws.Range("K:K"))
          CountMin = WorksheetFunction.Match(Min, ws.Range("K:K"), 0)
          ws.Cells(3, 18).Value = Min
          ws.Range("Q3").Value = ws.Range("I" & CountMin).Value
          
          'Greatest Total Volume:
          Max = WorksheetFunction.Max(ws.Range("L:L"))
          CountMax = WorksheetFunction.Match(Max, ws.Range("L:L"), 0)
          ws.Cells(4, 18).Value = Max
          ws.Range("Q4").Value = ws.Range("I" & CountMax).Value
          
          
          'Extra Bonous :)
          ws.Cells(5, 16).Value = "Number of records (rows) Analyzed: "
          ws.Cells(5, 18).Value = NumRows - 1
          ws.Range("R4:R5").NumberFormat = "#,##0"
          ws.Range("R2:R3").NumberFormat = "0.00%           "
          ws.Range("K:K").NumberFormat = "0.00%"
          ws.Range("L:L").NumberFormat = "#,##0"
          ws.Range("A:S").EntireColumn.AutoFit
          
 
Next ws

End Sub




