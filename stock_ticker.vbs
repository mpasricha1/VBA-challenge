Attribute VB_Name = "Module1"
Sub tickerLoop()
    'Declare Variables
    Dim wSheet As Worksheet
    Dim volTotal As Double
    Dim ticker As String
    Dim tickerCount As Integer
    Dim yearlyOpen As Double
    Dim yearlyClose As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim summaryTableLoc As Integer
    Dim greatestPerIncrease As Double
    Dim greatestPerDecrease As Double
    Dim greatestTicker As String
    Dim worstTicker As String
    Dim gVolTicker As String
    Dim greatestVol As Double
    Dim rowCount As Long
    
      
    For Each wSheet In Worksheets
        'Set headers on the Page
        wSheet.Cells(1, 9).Value = "Ticker"
        wSheet.Cells(1, 10).Value = "Yearly Change"
        wSheet.Cells(1, 11).Value = "Percentage Change"
        wSheet.Cells(1, 12).Value = "Total Volume"
        wSheet.Cells(1, 15).Value = "Ticker"
        wSheet.Cells(1, 16).Value = "Value"
        wSheet.Cells(2, 14).Value = "Greatest % Increase"
        wSheet.Cells(3, 14).Value = "Greatest % Decrease"
        wSheet.Cells(4, 14).Value = "Greatest Volume"
        
        'Initialize Variables
        summaryTableLoc = 2
        volTotal = 0
        yearlyOpen = Cells(2, 3).Value
        ticker = Cells(2, 1).Value
        greatestPerIncrease = 0
        greatestPerDecrease = 0
        greatestVol = 0
        rowCount = wSheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through all rows
        For i = 2 To rowCount
            'Determine where Ticker switches to new ticker
            If wSheet.Cells(i + 1, 1).Value <> wSheet.Cells(i, 1) Then
    
                'Set Values to be Printed
                yearlyClose = wSheet.Cells(i, 6).Value
                yearlyChange = (yearlyClose - yearlyOpen)
                volTotal = volTotal + Cells(i, 7).Value
                'ticker = Cells(i, 1).Value
                
                'Prevent divide by Zero
                If yearlyOpen = 0 Then
                    percentageChange = 0
                Else
                    percentageChange = ((yearlyClose - yearlyOpen) / yearlyOpen) * 100
                End If
                
                'Print the results to the Summary Table
                wSheet.Cells(summaryTableLoc, 9).Value = ticker
                wSheet.Cells(summaryTableLoc, 10).Value = yearlyChange
                wSheet.Cells(summaryTableLoc, 10).Value = yearlyChange
                wSheet.Cells(summaryTableLoc, 11).Value = percentageChange
                wSheet.Cells(summaryTableLoc, 12).Value = volTotal
                           
                'Conditional Formatting neg/red(3) pos/green(4)
                If Cells(summaryTableLoc, 10).Value > 0 Then
                    Cells(summaryTableLoc, 10).Interior.ColorIndex = 4
                ElseIf Cells(summaryTableLoc, 10) < 0 Then
                    Cells(summaryTableLoc, 10).Interior.ColorIndex = 3
                End If
                
                'Running Total of Challenge Section
                If percentageChange > greatestPerIncrease Then
                    greatestPerIncrease = percentageChange
                    greatestTicker = ticker
                End If
                If percentageChange < greatestPerDecrease Then
                    greatestPerDecrease = percentageChange
                    worstTicker = ticker
                End If
                If volTotal > greatestVol Then
                    greatestVol = volTotal
                    gVolTicker = ticker
                End If

                'Clean up and Increment values
                summaryTableLoc = summaryTableLoc + 1
                ticker = wSheet.Cells(i + 1, 1)
                yearlyOpen = Cells(i + 1, 4)
                volTotal = 0
                
            Else
                volTotal = volTotal + Cells(i, 7).Value
            End If
        Next i
        'Print Values for Challenge
        wSheet.Cells(2, 15).Value = greatestTicker
        wSheet.Cells(3, 15).Value = worstTicker
        wSheet.Cells(4, 15).Value = gVolTicker
        wSheet.Cells(2, 16).Value = greatestPerIncrease
        wSheet.Cells(3, 16).Value = greatestPerDecrease
        wSheet.Cells(4, 16).Value = greatestVol
    Next wSheet
End Sub

