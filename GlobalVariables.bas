Attribute VB_Name = "GlobalVariables"
Option Explicit

Public Type ChartItem
    Date As String
    Name As String
    WKN As String
    Value As Double
    SD As Double
    Distance As Double      ' Distance to moving average
    Account As Double
    Trend As String
    
End Type

Public Type ShareItem
    Date As String
    Time As String
    Name As String
    WKN As String
    ISIN As String
    Index As String
End Type

Public Type MousePos
    X As Double
    Y As Double
End Type



'=== Visualisieung ===
Public GlbScaleX As Double
Public GlbScaleY As Double
Public GlbOffX As Double
Public GlbOffY As Double

Public GlbPicOffX As Double
Public GlbPicOffY As Double


Public ChartFilename As String

Public ChartArray() As ChartItem



Public MouseDnPos As MousePos
Public MouseUpPos As MousePos
Public MouseMove As MousePos
Public OffsetCurrent As MousePos
Public OffsetLast As MousePos
Public ScaleCurrent As MousePos
Public ScaleLast As MousePos

'=== Analyse ===
Public SdLength As Long
'Public SharePrice As Double
'Public Rise As Boolean
Public StartSharePrice As Double
Public StartEuro As Double
Public StartAccount As Double


Public WKS_Download_idx As Long
Public DownloadPause As Long
Public WknReadCount As Long
Public DelayTime As Long
Public AccessErrorCnt As Long
Public AccessCnt As Long

'=== today's share price ===
Public CompanyListArray() As ShareItem
Public CompPartialLstArr() As ShareItem

Public AccountArray() As ChartItem



