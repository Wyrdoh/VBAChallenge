Attribute VB_Name = "Module1"
Sub StockAnalysis():

'I was struggling with this challenge due to the pace of the class and my work load in my job. I feel so behind on Python already that instead of asking for an extension I got some help from a friend that took
'this course in the past. Their project was a little different but he helped me with a good chunk of the last part with the Totals and helped me clean up my code. He explained everything pretty well and I feel I have a good
'understanding now but still might schedule a tutoring session on 4/1 to help me understand a few of the parts I'm still fuzzy on. I feel bad I needed so much help but really wanted to focus on Python as the pace is
'already pretty fast.

    Dim startTime As Single
    Dim endTime As Single
    
    startTime = Timer
    
For Each ws In Worksheets

    Dim WorksheetName As String
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecease As Double
    Dim GreatestTotalVolume As Double
    
WorksheetName = ws.Name
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
        TickCount = 2
        j = 2
        
LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRowA
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            If ws.Cells(TickCount, 10).Value < 0 Then
                ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                'I tried to use vbGreen and vbRed like I rememberd from class but it threw an error each time, once I found the index number it worked. Why is that?
            Else
                ws.Cells(TickCount, 10).Interior.ColorIndex = 4
            End If
            If ws.Cells(j, 3).Value <> 0 Then
                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                ws.Cells(TickCount, 11).Value = Format(PercentChange, "Percent")
            Else
                ws.Cells(TickCount, 11).Value = Format(0, "Percent")
            End If
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                TickCount = TickCount + 1
                j = i + 1
            End If
        Next i
            LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
            GreatestIncrease = ws.Cells(2, 11).Value
            GreatestDecrease = ws.Cells(2, 11).Value
            GreatestTotalVolume = ws.Cells(2, 12).Value
              
        For i = 2 To LastRowI
            If ws.Cells(i, 12).Value > GreatestTotalVolume Then
            GreatestTotalVolume = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        Else
            GreatestTotalVolume = GreatestTotalVolume
        End If
            If ws.Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        Else
            GreatestIncrease = GreatestIncrease
        End If
            If ws.Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecrease = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        Else
            GreatestDecrease = GreatestDecrease
        End If
            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatestTotalVolume, "Scientific")
            Next i
            
     endTime = Timer
     MsgBox "This script ran in " & (endTime - startTime) & " seconds"
     'Pressed for time, I could not for the life of me get all 3 sheets to add the times together and display at the end instead of counting each sheet and adding the times as it went
    
    Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
    Next ws
    
End Sub
