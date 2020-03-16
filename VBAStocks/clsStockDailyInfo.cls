VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStockDailyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit ' Force explicit variable declaration.

Private mTradeDate As Date
Private mOpenPrice As Double
Private mHighPrice As Double
Private mLowPrice As Double
Private mClosePrice As Double
Private mVolume As Long

' Constant definitions for the columns
Const DateColumn As Integer = 2
Const OpenColumn As Integer = 3
Const HighColumn As Integer = 4
Const LowColumn As Integer = 5
Const CloseColumn As Integer = 6
Const VolumeColumn As Integer = 7

Public Property Get TradeDate() As Date
    TradeDate = mTradeDate
End Property

Public Property Get OpenPrice() As Double
    OpenPrice = mOpenPrice
End Property

Public Property Get HighPrice() As Double
    HighPrice = mHighPrice
End Property

Public Property Get LowPrice() As Double
    LowPrice = mLowPrice
End Property

Public Property Get ClosePrice() As Double
    ClosePrice = mClosePrice
End Property

Public Property Get Volume() As Long
    Volume = mVolume
End Property


Public Sub PopulateFromSheet(sheet As Object, currentRow As Long)
    ' Capture values from sheet
    Dim columnDate As String
    columnDate = CStr(sheet.Cells(currentRow, DateColumn).Value)
    
    mTradeDate = CDate(Left(columnDate, 4) + "/" + Mid(columnDate, 5, 2) + "/" + Right(columnDate, 2))
    
    mOpenPrice = sheet.Cells(currentRow, OpenColumn).Value
    mHighPrice = sheet.Cells(currentRow, HighColumn).Value
    mLowPrice = sheet.Cells(currentRow, LowColumn).Value
    mClosePrice = sheet.Cells(currentRow, CloseColumn).Value
    mVolume = sheet.Cells(currentRow, VolumeColumn).Value
    
End Sub

