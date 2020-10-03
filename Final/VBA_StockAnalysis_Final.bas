Attribute VB_Name = "Module1"

Sub stockanalysis():
Application.ScreenUpdating = False

    Dim lastrowold As Long
    Dim lastrowenw As Integer
    Dim newcounter As Integer
    Dim oldcounter As Long
    Dim annualopen As Double
    Dim annualclose As Double
    Dim totalstock As LongLong
    Dim totalworksheets As Long
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolume As LongLong
    Dim winner As String
    Dim loser As String
    Dim volumewinnder As String
    Dim zerocounter As Long
    
    
 
    'counts the number of worksheets in the workbook and sets that value equal to totalworksheets
    
    totalworksheets = ActiveWorkbook.Worksheets.Count

    'Loops through each Worksheet
    
    For j = 1 To totalworksheets
    
        'activates each worksheet in the workbook
        
        Worksheets(j).Activate

        'sorting base on ticker name then date
                  
        ActiveWorkbook.Sheets(j).Sort.SortFields.Clear
        ActiveWorkbook.Sheets(j).Sort.SortFields.Add2 Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Sheets(j).Sort.SortFields.Add2 Key:=Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Sheets(j).Sort
            .SetRange Range("A:G")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            
        End With
        
        'copying tickers to new table and creating headers
        
        ActiveSheet.Range("A:A").Copy Range("J:J")
        ActiveSheet.Range("J:J").RemoveDuplicates Columns:=1, Header:=xlYes
        ActiveSheet.Range("K1").Value = "Yearly Change"
        ActiveSheet.Range("L1").Value = "Percent Change"
        ActiveSheet.Range("M1").Value = "Total Stock Volume"
        ActiveSheet.Columns("J:M").AutoFit
                            

        'Defines the last row we're copying from(lastrowold) and copying to(lastrownew)and counters for the new and old table
        
        lastrowold = Cells(Rows.Count, 1).End(xlUp).Row
        newcounter = 2
        oldcounter = 2
        
        'sets greatest increase , greatest decrease, and greatest volume as zero
        
        greatestincrease = 0
        greatestdecrease = 0
        greatestvolume = 0
                    
                    
        'loops through the entire new table while the yearly change column is blank
                
        While Cells(newcounter, 10).Value <> ""

            'defines the first open stock price in the sorted data set as the annual open
            
            annualopen = Cells(oldcounter, 3).Value
            
             'pulls in the next open that is not zero if the current open is zero
            zerocounter = 0
                    
            While annualopen = 0 And Cells(oldcounter + zerocounter, 1).Value = Cells(newcounter, 10).Value
                annualopen = Cells(oldcounter + zerocounter, 3).Value
                zerocounter = zerocounter + 1
                        
            Wend
                                               
                                               
                        
            
            'adds the volume column to the totalsock and defines the last close price as the annualclose
            totalstock = 0
                   
            For i = oldcounter To lastrowold
                If Cells(newcounter, 10).Value = Cells(oldcounter, 1).Value Then
                    totalstock = totalstock + Cells(oldcounter, 7).Value
                    oldcounter = oldcounter + 1
                ElseIf Cells(oldcounter, 1).Value <> Cells(oldcounter - 1, 1).Value Then
                    annualclose = Cells(oldcounter - 1, 6).Value
                    Exit For
                End If
            Next i
            
            'prints the annualopen,yearly change, annualclose and total stock on the new table
            'if statement deals with possiblity of annualclose being zero
            If annualclose = 0 Then
                
                yearlychange = annualclose - annualopen
                Cells(newcounter, 11).Value = yearlychange
                Cells(newcounter, 13).Value = totalstock
                yearlychange = annualclose - annualopen
                percentchange = 0
                
                
            Else:
                yearlychange = annualclose - annualopen
                percentchange = yearlychange / annualopen
                Cells(newcounter, 11).Value = yearlychange
                Cells(newcounter, 12).Value = percentchange
                Cells(newcounter, 13).Value = totalstock
                
            End If
            
            
            'checks to see if stock's increase is better/worse than the greatestincrease/greatestdecrease and redefines them if they are
            
            If percentchange > greatestincrease Then
                greatestincrease = percentchange
                winner = Cells(newcounter, 10).Value
                
            ElseIf percentchange < greatestdecrease Then
                greatestdecrease = percentchange
                loser = Cells(newcounter, 10).Value
                
            Else
            End If
                  
            'checks to see if stocks's volume is greater than greatest volume and redefines it
            
            If totalstock > greatestvolume Then
                greatestvolume = totalstock
                volumewinner = Cells(newcounter, 10).Value
            Else
            End If
    

            'changes the color of the yearly change cell based on its value
            If Cells(newcounter, 11).Value > 0 Then
                Cells(newcounter, 11).Interior.ColorIndex = 4
            Else
                Cells(newcounter, 11).Interior.ColorIndex = 3
            End If
            
            'resets the totalstock number and increases the counter for the new table by 1
            totalstock = 0
            newcounter = newcounter + 1
            
            

        Wend
               
        
        'prints header for challenge table
        Cells(1, 18).Value = "Ticker"
        Cells(1, 19).Value = "Value"
        
        
        'prints stats for greatest increase on challenge table
        Cells(2, 17).Value = "Greatest % Increase"
        Cells(2, 18).Value = winner
        Cells(2, 19).Value = greatestincrease
        
        
        'prints stats for greatest decrease on challenge table
        Cells(3, 17).Value = "Greatest % Decrease"
        Cells(3, 18).Value = loser
        Cells(3, 19).Value = greatestdecrease
        
        'prints stats for greatest volumen on challenge table
        Cells(4, 17).Value = "Greatest Total Volume"
        Cells(4, 18).Value = volumewinner
        Cells(4, 19).Value = greatestvolume
        
                
        'formats data on challenge table
        Columns("Q:S").AutoFit
        Range("S2:S3").NumberFormat = "0.00%"
        Cells(4, 19).NumberFormat = "#,##0.00"
                
        'formats data on new table
        Range("L:L").NumberFormat = "0.00%"
        Range("K:K").NumberFormat = "$#,##0.00"
        Range("M:M").NumberFormat = "#,##0"
        


    Next j
    
Application.ScreenUpdating = True

End Sub


