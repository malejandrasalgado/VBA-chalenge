Attribute VB_Name = "Module1"
Option Explicit
Sub Summary_Wall_Street()

'1. Define the variables

Dim CurrentRow As Range
Dim Ticker As String
Dim OpenPrice As Single
Dim ClosePrice As Single
Dim Volume_Traded As Double
Dim TickerCounter As Long
Dim RowCount As Long
Dim Sheet As Integer
Dim Current As Worksheet
Dim IncPer As Single '%Increase
Dim DecPer As Single '%Decrese



'Variables to track the greatest increase,decreas and volumen associated ticker codes

Dim GreatestPercentageIncrease As Single
Dim GreatestIncreaseTicker As String

Dim GreatestPercentageDecrease As Single
Dim GreatestDecreaseTicker As String

Dim GreatestVolume_Traded As Double
Dim GreatestVolume_Traded_Ticker As String

'Macro to Loop trough all Worksheet in a Workbook

For Each Current In Worksheets

    Current.Select
    
    
    '2 Set values to the Variables Formats & Location
    
    'Initial the variables before starting processing
    Ticker = ""
    RowCount = 2
    TickerCounter = 0
    OpenPrice = 0
    ClosePrice = 0
    Volume_Traded = 0
    'Initial the variables for the summary before starting processing
    
    GreatestPercentageIncrease = 0
    GreatestIncreaseTicker = ""
    
    GreatestPercentageDecrease = 0
    GreatestDecreaseTicker = ""
    
    GreatestVolume_Traded = 0
    GreatestVolume_Traded_Ticker = ""
    
    
    'Formats
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("J:J").NumberFormat = "#.00"
    Range("K:K").NumberFormat = "0.00%"
    Range("P2").NumberFormat = "0.00%"
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0.00#"
    
    
    
    Set CurrentRow = Current.Application.Range("A" & RowCount)
    
    '3. Find the Ticker using the While Loop
    
    Do While CurrentRow.Value <> ""
        
        'Find next ticker
        If CurrentRow.Value <> Ticker Then
            TickerCounter = TickerCounter + 1
        
            If Ticker <> "" Then
            'Create a Summary "ListTheTicker"
                ListTheTicker TickerCounter, Ticker, OpenPrice, ClosePrice, Volume_Traded
            
            If Volume_Traded > GreatestVolume_Traded Then
                GreatestVolume_Traded = Volume_Traded
                GreatestVolume_Traded_Ticker = Ticker
            End If
            
            If OpenPrice <> 0 Then
                IncPer = (ClosePrice - OpenPrice) / OpenPrice
            Else
                IncPer = 0
            End If
            
            If IncPer > GreatestPercentageIncrease Then
                GreatestPercentageIncrease = IncPer
                GreatestIncreaseTicker = Ticker
            End If
                
            If OpenPrice <> 0 Then
                DecPer = (ClosePrice - OpenPrice) / OpenPrice
            Else
                DecPer = 0
            End If
            
            If DecPer < GreatestPercentageDecrease Then
                GreatestPercentageDecrease = DecPer
                GreatestDecreaseTicker = Ticker
            End If
            
                      
            'Reset the sum of the Volume_Traded as process another ticket
                 Volume_Traded = 0
            
            End If
            'Sets the ticker to be the value of the row in Column A
                Ticker = CurrentRow.Value
             
            'Sets the Opening Price
                OpenPrice = CurrentRow.Offset(0, 2).Value
                ClosePrice = CurrentRow.Offset(0, 5).Value
            'Sets the daily Volume traded
                Volume_Traded = Volume_Traded + CurrentRow.Offset(0, 6).Value
           
        Else
            'Move onto the next row
                    RowCount = RowCount + 1
            'Select the next row
                    Set CurrentRow = Application.Range("A" & RowCount)
            'Add the daily Volume traded
                    If CurrentRow.Value = Ticker Then
                        Volume_Traded = Volume_Traded + CurrentRow.Offset(0, 6).Value
                        ClosePrice = CurrentRow.Offset(0, 5).Value
                    End If
        End If
    Loop
    
    'Reach the end of the file so output the Last Ticker
    'Adds the ListTheTicker Summury
        TickerCounter = TickerCounter + 1
        ListTheTicker TickerCounter, Ticker, OpenPrice, ClosePrice, Volume_Traded
        
        If Volume_Traded > GreatestVolume_Traded Then
            GreatestVolume_Traded = Volume_Traded
            GreatestVolume_Traded_Ticker = Ticker
        End If
        
        If OpenPrice <> 0 Then
            IncPer = (ClosePrice - OpenPrice) / OpenPrice
        Else
            IncPer = 0
        End If
        
        If IncPer > GreatestPercentageIncrease Then
                GreatestPercentageIncrease = IncPer
                GreatestIncreaseTicker = Ticker
        End If
        If OpenPrice <> 0 Then
            DecPer = (ClosePrice - OpenPrice) / OpenPrice
        Else
            DecPer = 0
        End If
        
        If DecPer < GreatestPercentageDecrease Then
                GreatestPercentageDecrease = DecPer
                GreatestDecreaseTicker = Ticker
        End If
        
        Range("P2").Value = GreatestIncreaseTicker
        Range("P3").Value = GreatestDecreaseTicker
        Range("P4").Value = GreatestVolume_Traded_Ticker
        Range("Q2").Value = GreatestPercentageIncrease
        Range("Q3").Value = GreatestPercentageDecrease
        Range("Q4").Value = GreatestVolume_Traded
           
Next

End Sub

Sub ListTheTicker(TickerCounter As Long, Ticker As String, OpenPrice As Single, ClosePrice As Single, Volume_Traded As Double)

'Summary of the data to cells starting in Column I
    '1. Define the variables
        Dim OutTicker As Range
    
    '2 Set values to the Variables & Location
        Set OutTicker = Application.Range("I" & TickerCounter)
        OutTicker.Value = Ticker
    
    '3 Calculate the difference in value between the opening and closing price & Location
        OutTicker.Offset(0, 1).Value = ClosePrice - OpenPrice
    
    '4 Calculate the difference in value expressed as a % between the opening and closing price & Location
        If OpenPrice <> 0 Then
            OutTicker.Offset(0, 2).Value = (ClosePrice - OpenPrice) / OpenPrice
        Else
            OutTicker.Offset(0, 2).Value = 0
        End If
        
    '5 Put the total Volume_Traded in the column netx to  the difference in value expressed as a %
        OutTicker.Offset(0, 3).Value = Volume_Traded
    
    If (ClosePrice < OpenPrice) Then
        OutTicker.Offset(0, 1).Interior.Color = RGB(255, 0, 0)
    Else
        OutTicker.Offset(0, 1).Interior.Color = RGB(0, 255, 0)
    End If
    Application.Range("I:Q").Columns.AutoFit

End Sub

