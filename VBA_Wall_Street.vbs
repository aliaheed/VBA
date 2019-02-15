Attribute VB_Name = "Module1"
Sub Stock_Market()

Dim lRow As Long ' calculates the rows in the document
Dim lCol As Long ' calculates the columns in the document

Dim iOpenPrice As Single
Dim iClosePrice As Single
Dim dYearlyChange As Double
Dim bOpenPriceFlag As Byte
Dim dPercentChange As Double
Dim dTotalStockVolume As Double
Dim k As Byte ' Flag k is used to handle the first row of the ticker in the data
Dim iOutputPringingCounter As Integer
Dim ws As Worksheet

Dim dGreatestIncrease As Double
Dim dGreatestDecrease As Double
Dim dGreatestTotVolume As Double

Dim bGreatestIncreaseRowIndex As Integer
Dim bGreatestDecreaseRowIndex As Integer
Dim bGreatestTotVolumeRowIndex As Integer


For Each ws In ActiveWorkbook.Worksheets

    ws.Activate
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Columns("I").AutoFit
    Columns("J").AutoFit
    Columns("K").AutoFit
    Columns("L").AutoFit
    
    
    bOpenPriceFlag = 0
    dTotalStockVolume = 0
    
    k = 0
    iOutputPringingCounter = 2

    'Find the count of non-blank cells in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
 
    For i = 2 To lRow

        'storing the Open price for the Stock
        If bOpenPriceFlag = 0 Then

            Do While Cells(i, 3).Value = 0
                i = i + 1

            Loop

            iOpenPrice = Cells(i, 3).Value
            bOpenPriceFlag = 1

        End If
        
        ' Below condition stores the information in the spreadsheet when the loop
        ' counter finds a non-matching Ticker in the next row
        If Trim(Cells(i, 1).Value) <> Trim(Cells(i + 1, 1).Value) Then

                k = 0

                iClosePrice = Cells(i, 6).Value
                dYearlyChange = iClosePrice - iOpenPrice
                dPercentChange = (dYearlyChange / iOpenPrice)
                bOpenPriceFlag = 0 ' Reseting the open price flag to 0

                Cells(iOutputPringingCounter, 9).Value = Cells(i, 1) ' Ticker
                Cells(iOutputPringingCounter, 10).Value = Round(dYearlyChange, 2) ' Yearly Change
                Cells(iOutputPringingCounter, 11).Value = dPercentChange  ' Percent Change
                Cells(iOutputPringingCounter, 12).Value = Format(dTotalStockVolume, "#,###") ' Total Stock Volume

                dTotalStockVolume = 0
    
                If Cells(iOutputPringingCounter, 10).Value < 0 Then
                    Cells(iOutputPringingCounter, 10).Interior.ColorIndex = 3
                ElseIf Cells(iOutputPringingCounter, 10).Value > 0 Then
                    Cells(iOutputPringingCounter, 10).Interior.ColorIndex = 4
                End If
                
                Cells(iOutputPringingCounter, 11).Style = "Percent"
                Cells(iOutputPringingCounter, 11).NumberFormat = "0.00%"
                
                
                iOutputPringingCounter = iOutputPringingCounter + 1
    
        Else
            
            'Go in the 1st k = 0 condition if this is the 1st row of the ticker in the data
            If k = 0 Then
                dTotalStockVolume = Cells(i, 7).Value + Cells(i + 1, 7).Value
                k = 1
            
            Else
                dTotalStockVolume = dTotalStockVolume + Cells(i + 1, 7).Value
            
            End If
            
        End If
    
    Next i
    
Next ws

' Following loop calculates the maximum and minimum % change
' as well as the maximum total volume in each sheet
For Each ws In ActiveWorkbook.Worksheets

    ws.Activate

    lRow = Cells(Rows.Count, 9).End(xlUp).Row

    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    Columns("O").AutoFit
    Columns("P").AutoFit
    Columns("Q").AutoFit

    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"

    dGreatestIncrease = 0
    dGreatestDecrease = 0
    dGreatestTotVolume = 0

    bGreatestIncreaseRowIndex = 0
    bGreatestDecreaseRowIndex = 0
    bGreatestTotVolumeRowIndex = 0

        For i = 2 To lRow
            If Cells(i, 11).Value > dGreatestIncrease Then
                dGreatestIncrease = Cells(i, 11).Value
                bGreatestIncreaseRowIndex = i
            End If

            If Cells(i, 11).Value < dGreatestDecrease Then
                dGreatestDecrease = Cells(i, 11).Value
                bGreatestDecreaseRowIndex = i
            End If
            
            If Cells(i, 12).Value > dGreatestTotVolume Then
                dGreatestTotVolume = Cells(i, 12).Value
                bGreatestTotVolumeRowIndex = i
            End If

        Next i

 Cells(2, 16).Value = Cells(bGreatestIncreaseRowIndex, 9).Value
 Cells(2, 17).Value = dGreatestIncrease
 Cells(2, 17).Style = "Percent"
 Cells(2, 17).NumberFormat = "0.00%"
 
 Cells(3, 16).Value = Cells(bGreatestDecreaseRowIndex, 9).Value
 Cells(3, 17).Value = dGreatestDecrease
 Cells(3, 17).Style = "Percent"
 Cells(3, 17).NumberFormat = "0.00%"
 
 Cells(4, 16).Value = Cells(bGreatestTotVolumeRowIndex, 9).Value
 Cells(4, 17).Value = Format(dGreatestTotVolume, "#,###")
 
Next ws


End Sub



