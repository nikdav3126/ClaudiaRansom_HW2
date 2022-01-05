Attribute VB_Name = "Module1"
Sub StockInformationCalculator()

    'Place designations here
    'Dim RunThrough As Integer
    Dim YearlyChange As Double
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    TotalVolume = 0
    Dim TickerType As String
    Dim Final As Long
    'Set variable to store different types into
    Dim ChangeVary As Long
    Dim ChangeVary2 As Long
    Dim ChangeVary3 As Long
    Dim ChangeVary4 As Long
    ChangeVary = 2
    ChangeVary2 = 2
    ChangeVary3 = 2
    ChangeVary4 = 2
    YearlyChange = 0
    YearlyOpen = 0
    YearlyClose = 0
    'Run code through all worksheets listed
    'RunThrough = Application.Worksheet.Count
    
    'Create For loop to run through all worksheets
    'For w = 1 To RunThrough
    
    'Next w
    
    'Create headers for variables we need to return
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("Q1").Value = "Ticker"
    Range("R1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    
    
    'Make workbook fluid by finding last row and using that in for loop
    Final = Cells(Rows.Count, 1).End(xlUp).Row
 
    'Create for loop to return ticker type
    For c = 2 To Final

        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then
            TickerType = Cells(c, 1).Value '
            Cells(ChangeVary, 9).Value = TickerType
            ChangeVary = ChangeVary + 1
        End If
    
    Next c
    
    'Create for loop to find total stock volume
    For c = 2 To Final
    
        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then
        TotalVolume = TotalVolume + Cells(c, 7).Value
        Cells(ChangeVary2, 12).Value = TotalVolume
        ChangeVary2 = ChangeVary2 + 1
        
        TotalVolume = 0
        
        Else
        
        TotalVolume = TotalVolume + Cells(c, 7).Value
        
        End If
        
        Next c
        
    'Create For loop to store variables of yearly open and close
    YearlyOpen = Cells(2, 3).Value
    
    
    For c = 2 To Final
        
        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then
        YearlyClose = Cells(c, 6).Value
        YearlyChange = YearlyClose - YearlyOpen
        Cells(ChangeVary3, 10).Value = YearlyChange
        ChangeVary3 = ChangeVary3 + 1
        
        YearlyOpen = Cells(c + 1, 3).Value
        End If
        
        YearlyClose = 0
        YearlyChange = 0
    
    Next c
    
    For c = 2 To Final
    
        If Cells(c, 10).Value <= 0 Then
        Cells(c, 10).Interior.ColorIndex = 3
        
        ElseIf Cells(c, 10).Value > 0 Then
        Cells(c, 10).Interior.ColorIndex = 4
        
        End If
    
    Next c
    
    'Create For loop to find calculated percentage
    YearlyOpen = Cells(2, 3).Value
    
    
    For c = 2 To Final
        
        If Cells(c + 1, 1).Value <> Cells(c, 1).Value Then
        YearlyClose = Cells(c, 6).Value
        YearlyChange = YearlyClose - YearlyOpen
        PercentChange = (YearlyChange / YearlyOpen) * 100
        Cells(ChangeVary4, 11).Value = Round(PercentChange, 2) & "%"
        ChangeVary4 = ChangeVary4 + 1
        
        YearlyOpen = Cells(c + 1, 3).Value
        
        End If
        
        YearlyClose = 0
        YearlyChange = 0
        PercentChange = 0
    
    Next c
        
    
End Sub

