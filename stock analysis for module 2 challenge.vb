Sub stock_analysis()

'-----------------------------------------------------------------------------------------------
'CREATE LOOP THAT WILL RUN THE SUBMODULE FOR ALL WORKSHEETS IN WORKBOOK
'-----------------------------------------------------------------------------------------------

Dim i As Integer

'Select the worksheet, starting with the first sheet
For i = 1 To ActiveWorkbook.Worksheets.Count
   
   Worksheets(i).Select
   
    '-----------------------------------------------------------------------------------------------
    'IDENTIFY VARIABLES AND BUILD A SUMMARY TABLE
    '-----------------------------------------------------------------------------------------------
    
    'Create column headers for summary table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("I1:L1").EntireColumn.AutoFit
    
    'Identify the last row in the sheet and set aside memory for stock at first open and last close,
    'yearly change, percent change, row, volumetotal, ticker name, and summary table row tracker
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).row
    Dim firstopen As Currency
    Dim lastclose As Currency
    Dim percentchange As Currency
    Dim row As Long
    Dim volumetotal As LongLong
    Dim ticker As String
    Dim summarytablerow As Long
    summarytablerow = 2
    
    'Create variables to track the greatest % increase, decrease, and total volume
    Dim greatestincreaseticker As String
    Dim greatestincrease As Variant
    greatestincrease = 0
    Dim greatestdecreaseticker As String
    Dim greatestdecrease As Variant
    greatestdecrease = 0
    Dim greatestvolumeticker As String
    Dim greatestvolume As LongLong
    greatestvolume = 0
    
    '-----------------------------------------------------------------------------------------------
    'SEARCH THROUGH DATA AND COMPILE / COMPUTE THE DESIRED OUTCOMES
    '-----------------------------------------------------------------------------------------------
    
    'Loop through all entries
    For row = 2 To lastrow
    
        'Check if this is the first row of a new ticker
        If Cells(row - 1, 1).Value <> Cells(row, 1).Value Then
        
            'If it is a new ticker, set the ticker name, stock value at first open, and volume
            ticker = Cells(row, 1).Value
            firstopen = Cells(row, 3).Value
            volumetotal = Cells(row, 7).Value
            
            Else
        
            'If it is not a new ticker, then add to the existing ticker's volume total
            volumetotal = volumetotal + Cells(row, 7).Value
            
            End If
        
        'Check if this is the last row of the current ticker
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
            
            'If yes, then set the stock at last close and add to the volume total
            lastclose = Cells(row, 6).Value
            volumetotal = volumetotal + Cells(row, 7).Value
            
            'Calculate the yearly change and percent change
            yearlychange = (lastclose - firstopen)
            percentchange = ((lastclose - firstopen) / firstopen)
            
            'Print ticker, yearly change, percent change, and total stock volume to summary table
            Range("I" & summarytablerow).Value = ticker
            Range("J" & summarytablerow).Value = yearlychange
            Range("K" & summarytablerow).Value = FormatPercent(percentchange, 2)
            Range("L" & summarytablerow).Value = volumetotal
            
            'Make yearly change cells red or green, for yearly decreases or increases, respectively
            If yearlychange < 0 Then
            Range("J" & summarytablerow).Interior.ColorIndex = 3
            
                Else
                Range("J" & summarytablerow).Interior.ColorIndex = 4
                
                End If
            
            'Move to next row on the summary table, to prepare for the next ticker summary
            summarytablerow = summarytablerow + 1
            
            End If
    
    '-----------------------------------------------------------------------------------------------
    'CHECK IF CURRENT TICKER IS THE GREATEST INCREASE, DECREASE, OR TOTAL VOLUME
    '-----------------------------------------------------------------------------------------------
            
            'Is the percent change more than greatest increase on record?
            If percentchange > greatestincrease Then
            
                'If yes, then record the current ticker and percent change
                greatestincrease = percentchange
                greatestincreaseticker = ticker
                
                End If
            
            'Is the percent change negative and greater than the greatest decrease on record?
            If percentchange < greatestdecrease Then
            
                'If yes, then record the current ticker and percent change
                greatestdecrease = percentchange
                greatestdecreaseticker = ticker
                
                End If
            
            'Is the total volume more than the greatest total volume on record?
            If volumetotal > greatestvolume Then
            
                'If yes, then record the current ticker and total volume
                greatestvolume = volumetotal
                greatestvolumeticker = ticker
            
                End If
               
    Next row
            
            
    '-----------------------------------------------------------------------------------------------
    'END OF SEARCHING THROUGH DATA ... TIME TO PRINT THE GREATEST % CHANGES AND VOLUME...
    '-----------------------------------------------------------------------------------------------
            
    'After summary table is finished, print the table for greatest % changes and total stock volume
    
        'Create row and column headers
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("N2:P4").EntireColumn.AutoFit
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
        'Print greatest % increase, in percent format with 2 decimals
        Range("O2").Value = greatestincreaseticker
        Range("P2").Value = FormatPercent(greatestincrease, 2)
        
        'Print greatest % decrease, in percent format with 2 decimals
        Range("O3").Value = greatestdecreaseticker
        Range("P3").Value = FormatPercent(greatestdecrease, 2)
        
        'Print greatest total volume
        Range("O4").Value = greatestvolumeticker
        Range("P4").Value = greatestvolume

'Go to the next worksheet until you reach the last sheet
Next i

End Sub

