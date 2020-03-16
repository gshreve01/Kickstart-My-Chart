Attribute VB_Name = "modStockData"
Option Explicit ' Force explicit variable declaration.

' Factory Function method to create a Stock Object
Public Function CreateStock(Year As Integer) As clsStock

    Dim stockObject As clsStock
    
    Set stockObject = New clsStock
    stockObject.Year = Year
    
    Set CreateStock = stockObject
End Function

Sub TickerSymbolAnalyze()
    ' Declare variables
    Dim sheet As Object
    Dim stock As clsStock
    Dim Year As Integer
    Dim currentRow As Long
    Dim lastRow As Long
    Dim stocks() As clsStock
    Dim stockIdx As Integer
    Dim stockUpperBound As Integer
    
    Dim percentComplete As Double
    
    ' Hard Items
    Dim greatestIncreaseStock As clsStock
    Dim greatestDecreaseStock As clsStock
    Dim largestVolumeStock As clsStock
    
    
    ' Get the active sheet and set the year
    Set sheet = ActiveSheet
    Year = CInt(sheet.Name)
    
    ' Set the last Row
    lastRow = Cells(Rows.Count, 1).End(xlUp).row
    
    ' Clear any previous data set
    Call ClearSheetOutput(sheet, lastRow)
    
    ' Set the current row to where the data begins
    currentRow = 2
    
    ' Set the idx of stock to -1 so that the first one added will be 0
    stockIdx = -1
    
    ' Size the upper bound limit of stocks
    stockUpperBound = 100
    ReDim stocks(stockUpperBound)
    
    ' Populate clsStock based on every unique ticker symbol
    While currentRow <= lastRow
        stockIdx = stockIdx + 1
        
        ' Check if array size growth is needed
        If stockIdx > stockUpperBound Then
            stockUpperBound = stockUpperBound + 50
            ReDim Preserve stocks(stockUpperBound)
        End If
        
        ' Create the stock and add it to the array
        Set stock = CreateStock(Year)
        Set stocks(stockIdx) = stock
        
        ' Populate the stock with its values.  Current Row
        ' will return with the next Ticker symbol
        Call stock.PopulateFromSheet(sheet, currentRow)
        
        If currentRow Mod 100 = 0 Then
            Call OutputPercentComplete(sheet, currentRow / lastRow)
        End If
        
    Wend
    
    
    ' Free up unused space
    ReDim Preserve stocks(stockIdx)
   
    ' Ensure some data was loaded by checking stockIdx.
    ' If it is greater than -1 at least 1 stock should have been
    ' created
    If stockIdx = -1 Then
        Exit Sub
    End If
    
    ' Generate Headers using the first stock entry
    Call stocks(0).GenerateOutputHeaders(sheet)

    ' Output data on each stock
    currentRow = 2
    For stockIdx = LBound(stocks) To UBound(stocks)
        Set stock = stocks(stockIdx)
        
        ' Is this the greatest increase stock
        If greatestIncreaseStock Is Nothing Then
            Set greatestIncreaseStock = stock
        ElseIf greatestIncreaseStock.PercentChange < stock.PercentChange Then
            Set greatestIncreaseStock = stock
        End If
        If greatestDecreaseStock Is Nothing Then
            Set greatestDecreaseStock = stock
        ElseIf greatestDecreaseStock.PercentChange > stock.PercentChange Then
            Set greatestDecreaseStock = stock
        End If
        If largestVolumeStock Is Nothing Then
            Set largestVolumeStock = stock
        ElseIf largestVolumeStock.TotalStockVolume < stock.TotalStockVolume Then
            Set largestVolumeStock = stock
        End If
        Call stock.GenerateOutputData(sheet, currentRow)
        currentRow = currentRow + 1
    Next stockIdx
    
    Call OutputHardData(greatestIncreaseStock, greatestDecreaseStock, largestVolumeStock)
    
    ' Clear PercentComplete
    Call OutputPercentComplete(sheet, -1)
    
    ' AutoFit output data
    sheet.Columns("I:Q").AutoFit
        
End Sub

Sub OutputHardData(greatestIncreaseStock As clsStock, greatestDecreaseStock As clsStock, greatestVolumeStock As clsStock)
    'Output constants
    Const OutLabels As Integer = 15
    Const OutTicker As Integer = 16
    Const OutValue As Integer = 17

    ' Output hard data labels
    Cells(1, OutTicker).Value = "Ticker"
    Cells(1, OutValue).Value = "Value"
    Cells(2, OutLabels).Value = "Greatest % Increase"
    Cells(3, OutLabels).Value = "Greatest % Decrease"
    Cells(4, OutLabels).Value = "Greatest Total Volume"
    
    ' Output greatest increase
    Cells(2, OutTicker).Value = greatestIncreaseStock.TickerSymbol
    Cells(2, OutValue) = Format(greatestIncreaseStock.PercentChange, "Percent")
    
    ' Output greatest decrease
    Cells(3, OutTicker).Value = greatestDecreaseStock.TickerSymbol
    Cells(3, OutValue) = Format(greatestDecreaseStock.PercentChange, "Percent")
    
    ' Output Largest Volume
    Cells(4, OutTicker).Value = greatestVolumeStock.TickerSymbol
    Cells(4, OutValue).Value = greatestVolumeStock.TotalStockVolume
    
End Sub

Sub OutputPercentComplete(sheet As Object, percentComplete As Double)
    If (percentComplete > 0) Then
        sheet.Range("Q18") = Format(percentComplete, "Percent")
    Else
        sheet.Range("Q18").Value = ""
    End If
End Sub



Sub ClearSheetOutput(sheet As Object, lastRow As Long)
    Dim rangeStr As String
    
    rangeStr = "I1:Q" + CStr(lastRow)
    sheet.Range(rangeStr).Clear
End Sub
