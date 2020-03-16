VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit ' Force explicit variable declaration.


Private mTickerSymbol As String
Private mOpeningPrice As Double
Private mClosingPrice As Double
Private mStockVolume As Integer
Private mYear As Integer
Private mDailyInfo() As clsStockDailyInfo

Private mYearlyChange As Double
Private mPercentChange As Double
Private mTotalStockVolume As Double


' Constant information for the columns
Const TickerColumn As Integer = 1
Const OutTickerColumn As Integer = 9
Const OutYearlyChange As Integer = 10
Const OutPercentChange As Integer = 11
Const OutTotalStockVolume As Integer = 12


Public Property Let Year(pYear As Integer)
    mYear = pYear
End Property

Public Property Get Year() As Integer
    Year = mYear
End Property

Private Property Let TickerSymbol(pTickerSymbol As String)
    mTickerSymbol = pTickerSymbol
End Property
Public Property Get TickerSymbol() As String
    TickerSymbol = mTickerSymbol
End Property

Private Property Let OpeningPrice(pOpeningPrice As Double)
    mOpeningPrice = pOpeningPrice
End Property
Public Property Get OpeningPrice() As Double
    OpeningPrice = mOpeningPrice
End Property

Private Property Let ClosingPrice(pClosingPrice As Double)
    mClosingPrice = pClosingPrice
End Property
Public Property Get ClosingPrice() As Double
    ClosingPrice = mClosingPrice
End Property

Private Property Let YearlyChange(pYearlyChange As Double)
    mYearlyChange = pYearlyChange
End Property
Private Property Get YearlyChange() As Double
    YearlyChange = mYearlyChange
End Property

Private Property Let PercentChange(pPercentChange As Double)
    mPercentChange = pPercentChange
End Property
Public Property Get PercentChange() As Double
    PercentChange = mPercentChange
End Property

Private Property Let TotalStockVolume(pTotalStockVolume As Double)
    mTotalStockVolume = pTotalStockVolume
End Property
Public Property Get TotalStockVolume() As Double
    TotalStockVolume = mTotalStockVolume
End Property


Public Sub PopulateFromSheet(sheet As Object, ByRef currentRow As Long)
    Dim dailyInfo As clsStockDailyInfo
    Dim dailyInfoUpperBound As Integer
    Dim localTickerName As String
    Dim dailyInfoIdx As Integer
    Dim changePct As Double
    Dim runningVolume As Double
    
    Dim initialRow As Long
    
    ' Capture currentRow provided at start - may help with debugging
    initialRow = currentRow
    
    ' Capture the ticker symbol name
    TickerSymbol = sheet.Cells(currentRow, TickerColumn)
    Debug.Print ("Processing " & TickerSymbol)
    localTickerName = TickerSymbol
    
    ' Set daily info idx to -1 to allow for first entry to be 0
    dailyInfoIdx = -1
    
    ' Set an original size for the dimension, but resize if needed and shrink at end
    dailyInfoUpperBound = 100
    ReDim mDailyInfo(dailyInfoUpperBound)
    While localTickerName = TickerSymbol
        ' Create dailyInfo instance
        Set dailyInfo = New clsStockDailyInfo
        
        ' Populate it from the sheet
        Call dailyInfo.PopulateFromSheet(sheet, currentRow)
        
        ' Add it to the array
        dailyInfoIdx = dailyInfoIdx + 1
        If dailyInfoIdx > dailyInfoUpperBound Then
            dailyInfoUpperBound = dailyInfoUpperBound + 50
            ReDim Preserve mDailyInfo(dailyInfoUpperBound)
        End If

        Set mDailyInfo(dailyInfoIdx) = dailyInfo
        
        ' Go to next row and set comparison value to see if we are done with this ticker
        currentRow = currentRow + 1
        localTickerName = sheet.Cells(currentRow, TickerColumn)
    Wend
    
    ' Free up unneeded space
    ReDim Preserve mDailyInfo(dailyInfoIdx)
    
    ' Get Opening Price as the first entry in daily Info
    OpeningPrice = mDailyInfo(LBound(mDailyInfo)).OpenPrice
    
    ' Get Closing Price as the last entry in daily info
    ClosingPrice = mDailyInfo(UBound(mDailyInfo)).ClosePrice
    
    ' Yearly Change is the start price - end price
    YearlyChange = ClosingPrice - OpeningPrice
    
    ' percent change over year is the yearly change divided by starting price
    If OpeningPrice <> 0 Then
        PercentChange = YearlyChange / OpeningPrice
    Else
        PercentChange = 0
    End If
    
    ' Sum each Days volume to get yearly volume
    runningVolume = 0
    For dailyInfoIdx = LBound(mDailyInfo) To UBound(mDailyInfo)
        runningVolume = runningVolume + mDailyInfo(dailyInfoIdx).Volume
    Next dailyInfoIdx
    TotalStockVolume = runningVolume
 
    ' Clean up space
    ReDim mDailyInfo(0)
    
End Sub

' Generates the expected headers for the output information for
' the analysis
Public Sub GenerateOutputHeaders(sheet As Object)
    sheet.Cells(1, OutTickerColumn).Value = "Ticker"
    sheet.Cells(1, OutYearlyChange).Value = "Yearly Change"
    sheet.Cells(1, OutPercentChange).Value = "Percent Change"
    sheet.Cells(1, OutTotalStockVolume).Value = "Total Stock Volume"
End Sub


' Output the yearly values
Public Sub GenerateOutputData(sheet As Object, row As Long)
    sheet.Cells(row, OutTickerColumn).Value = TickerSymbol
    sheet.Cells(row, OutYearlyChange).Value = YearlyChange
    
    ' Set the color to green if change is positive, and red if change is negative
    If YearlyChange > 0 Then
        sheet.Cells(row, OutYearlyChange).Interior.ColorIndex = 4
    ElseIf YearlyChange < 0 Then
        sheet.Cells(row, OutYearlyChange).Interior.ColorIndex = 3
    End If
        
    sheet.Cells(row, OutPercentChange) = Format(PercentChange, "Percent") ' will round to 2 decimal places
    sheet.Cells(row, OutTotalStockVolume).Value = TotalStockVolume
End Sub










