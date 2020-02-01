Attribute VB_Name = "Chart"
Option Explicit

Dim ColorCoord As Long


Public Sub DispCoordinateSystem()
    Dim mx As Double
    Dim my As Double
    Dim ex As Double
    Dim ey As Double

    ColorCoord = vbRed

'    mx = GlbCx * GlbScaleX
'    my = GlbCY * -GlbScaleY
    
'    GlbOffX = 100
'    GlbOffY = 100
    'Draw Axis
    Form1.PicChart.Line (GlbOffX, Form1.PicChart.Height - GlbOffY)-(GlbOffX + 27000, Form1.PicChart.Height - GlbOffY), ColorCoord
    Form1.PicChart.Line (GlbOffX, Form1.PicChart.Height - GlbOffY)-(GlbOffX, Form1.PicChart.Height - (GlbOffY + 2000)), ColorCoord
    
End Sub


Public Sub ReadHistoryFile(HistoryFileName As String, CompanyName As String)
    Dim ChartFile As Integer
    Dim Zeile As String
    Dim ChartEntities() As String
    Dim idx As Long
    
    ReDim ChartArray(0 To 0)
    
    On Error GoTo ReadHistoryFileErr
    
    ChartFile = FreeFile
    Open HistoryFileName For Input As ChartFile
        
    While Not EOF(ChartFile)
        Line Input #ChartFile, Zeile
        SepariereString Zeile, ChartEntities, ";"
        idx = UBound(ChartArray)
        ChartArray(idx).Date = ChartEntities(0)
        ChartArray(idx).Value = Zahl(ChartEntities(4))
        ChartArray(idx).WKN = DateiName$(NameOhneKennung$(HistoryFileName))
        ChartArray(idx).Name = CompanyName
                
        ReDim Preserve ChartArray(0 To UBound(ChartArray) + 1)
    Wend
    ReDim Preserve ChartArray(0 To UBound(ChartArray) - 1)
    Close ChartFile
      
    Exit Sub
ReadHistoryFileErr:
    MsgBox HistoryFileName & vbCr & Err.Description, , "xxxxx"
End Sub






Public Sub DisplayChart()

    Dim idx As Long
DisplayTestPattern
    'idx 0 is the head line
    idx = 1

    If (0 / 1) + (Not Not ChartArray) = 0 Then
      ' Array ist nicht nicht dimensioniert
      Exit Sub
    End If

    DrawStart CDbl(idx), ChartArray(idx).Value, ColorCoord
    For idx = 1 To UBound(ChartArray)
        DrawLine CDbl(idx), ChartArray(idx).Value, ColorCoord
    Next idx


    'idx 0 is the head line
    idx = 1
    Dim DistanceColor As Long
    DistanceColor = vbWhite
    DrawStart CDbl(idx), ChartArray(idx).SD, DistanceColor
    For idx = 1 To UBound(ChartArray)
        If ChartArray(idx).Distance > 0 Then
            DistanceColor = vbGreen
        Else
             DistanceColor = vbBlue
       End If
        DrawLine CDbl(idx), ChartArray(idx).SD, DistanceColor
    Next idx
    

    idx = 1
    DrawStart CDbl(idx), ChartArray(idx).Value, vbBlack
    For idx = 1 To UBound(ChartArray)
        DrawLine CDbl(idx), ChartArray(idx).Account, vbBlack
    Next idx
    


End Sub


Public Sub MovingAverage(Length As Long)
    Dim idx As Long
    Dim i As Long
    Dim Sum As Double
    Dim Average As Double
    Dim Distance As Double
    Static LastDistance As Double
    
    If (0 / 1) + (Not Not ChartArray) = 0 Then
      ' Array ist nicht nicht dimensioniert
      Exit Sub
    End If
   
    
    If UBound(ChartArray) <= Length Then
        Exit Sub
    End If
    
    
    If Length = 0 Then
        Length = 1
    End If
    
    idx = 1
    Sum = 0
    While idx < Length
        Sum = Sum + ChartArray(idx).Value
        Average = Sum / idx
        ChartArray(idx).SD = Average
        Distance = (ChartArray(idx).Value - Average) / ChartArray(idx).Value
        ChartArray(idx).Distance = Distance
        idx = idx + 1
    Wend

    For idx = Length To UBound(ChartArray)
        Sum = 0
        For i = idx - Length + 1 To idx
            Sum = Sum + ChartArray(i).Value
        Next i
        Average = Sum / Length
        ChartArray(idx).SD = Average
        
        ' Share prises in history files sometimes are zero
        ' Just avoid division by zero
        If ChartArray(idx).Value = 0 Then
            Distance = LastDistance
        Else
            Distance = (ChartArray(idx).Value - Average) / ChartArray(idx).Value
        End If
        ChartArray(idx).Distance = Distance
        LastDistance = Distance
    Next idx
End Sub

'=====================================================================
'                       Analyse_01
' Einstieg: wenn der Kurs über dem GD liegt, wird gekauft.
' Wenn der Kurs von unten durch den GD sticht, wird gekauft.
' Wenn der Kurs von oben unter den GD fällt, wird verkauft.
'=====================================================================
Public Sub Analyse_01()
    Dim idx As Long
    Dim BuyNow As Boolean
    Dim SharePrice As Double
    Dim Rise As Boolean


    
    SharePrice = 1
'    ChartArray(idx).Account = SharePrice
'    For idx = 0 To 1
'        ChartArray(idx).Account = SharePrice
'    Next idx
    
    idx = CLng(Zahl(Form1.T_InvestmentStart))
    
'    SharePrice = 10
    SharePrice = ChartArray(idx).Value
    
    ChartArray(idx).Account = SharePrice
    ChartArray(idx + 1).Account = SharePrice
    ChartArray(idx - 1).Account = SharePrice
    Rise = False
     
    While idx <= UBound(ChartArray)
        If ChartArray(idx).Distance > 0 Then
            ChartArray(idx).Trend = "Rise"
            If Rise = False Then
                StartSharePrice = ChartArray(idx).Value
                StartAccount = ChartArray(idx - 1).Account
                BuyNow = True
            Else
                BuyNow = False
            End If
            
            Rise = True
        Else
             ChartArray(idx).Trend = "Hold"
            If Rise = True Then
                ChartArray(idx - 1).Account = ChartArray(idx - 1).Account * 0.985
            End If
            Rise = False
        End If
        
        If Rise Then
            ChartArray(idx).Account = ChartArray(idx).Value / StartSharePrice * StartAccount
'    If ChartArray(idx).Account > 1000 Then
'        ChartArray(idx).Account = 1000
'    End If
        Else
            ChartArray(idx).Account = ChartArray(idx - 1).Account
        End If
        
      
    idx = idx + 1
    Wend
    
'    For idx = Length To UBound(ChartArray)
'
'    Next idx
    
End Sub


'=====================================================================
'                       Analyse_02
' Einstieg: Im Gegelsatz zu Analye_01 wird zuerst gewartet, bis
'           bis der Kurs von unten durch den GS sticht.
' Wenn der Kurs von unten durch den GD sticht, wird gekauft.
' Wenn der Kurs von oben unter den GD fällt, wird verkauft.
' InvestmentStart: Der Index im History file
' StartAccount: Die Investitionssumme
'=====================================================================
Public Sub Analyse_02(InvestmentStart As Long, StartAccount As Double)
    Dim idx As Long
    Dim ReadyForFirstRise As Boolean
    Dim Step As Long
    Dim CostFactor As Double
    
    CostFactor = 0.9926
'    CostFactor = 0.991
    
    idx = 0
    Step = 0
    
    If (0 / 1) + (Not Not ChartArray) = 0 Then
      ' Array ist nicht nicht dimensioniert
      Exit Sub
    End If
    
    While idx <= UBound(ChartArray)
        Select Case Step
            Case 0:
                ' no investment before InvestmentStart
                If idx >= InvestmentStart Then
                    ChartArray(idx).Account = 0
                    ChartArray(idx).Trend = "0"
                    Step = 5
                Else
                    ChartArray(idx).Account = 0
                    ChartArray(idx).Trend = "0"
                End If
        
            Case 5:
                ' share price under GD now
                If ChartArray(idx).Distance <= 0 Then
                    ChartArray(idx).Account = 0
                    ChartArray(idx).Trend = "5:wait"
                    Step = 10
                ' wait until share price under GD
                Else
                    ChartArray(idx).Account = 0
                    ChartArray(idx).Trend = "5: wait"
                End If
            Case 10:
                ' wait until share price is over GD again the first time
                '*** buy now
                If ChartArray(idx).Distance > 0 Then
                    StartSharePrice = ChartArray(idx).Value
                    ChartArray(idx).Account = (ChartArray(idx).Value / StartSharePrice) * StartAccount
                    ' Remove transfer costs
                    ChartArray(idx).Account = ChartArray(idx).Account * CostFactor
                    ChartArray(idx).Trend = "10: Rise"
                    Step = 20
                Else
                    ChartArray(idx).Account = 0
                    ChartArray(idx).Trend = "10: wait"
                End If
            Case 20:
                ' if share price falls under GD again
                If ChartArray(idx).Distance <= 0 Then
                    ChartArray(idx).Trend = "20: Hold"
                    ChartArray(idx).Account = ChartArray(idx - 1).Account * CostFactor
                    Step = 30
                ' share price stays over GD
                Else
                    ChartArray(idx).Trend = "20: Rise"
                    ChartArray(idx).Account = (ChartArray(idx).Value / StartSharePrice) * StartAccount
                End If
            Case 30:
                ' if share price raised over GD again
                If ChartArray(idx).Distance > 0 Then
                    ChartArray(idx).Trend = "30: Rise"
'                    ChartArray(idx).Account = ChartArray(idx).Value / StartSharePrice * StartAccount
                    StartSharePrice = ChartArray(idx).Value
                    ChartArray(idx).Account = ChartArray(idx - 1).Account * CostFactor
                    StartAccount = ChartArray(idx).Account
                    Step = 20
                Else
                     ' share price stays under GD
                    ChartArray(idx).Trend = "30: Hold"
                    ChartArray(idx).Account = ChartArray(idx - 1).Account
                End If
            
        End Select
      
    idx = idx + 1
    Wend
    
End Sub


Public Sub DrawStart(X As Double, Y As Double, LclColor As Long)
    Dim PicX As Double
    Dim PicY As Double
    
    PicX = (X * GlbScaleX) + GlbOffX
    PicY = Form1.PicChart.Height - (Y * GlbScaleY) - GlbOffY
    
    If PicX > Form1.PicChart.Width + 100 Then
        PicX = Form1.PicChart.Width + 100
    End If
    
    If PicX < -100 Then
        PicX = -100
    End If
    
    If PicY > Form1.PicChart.Height + 100 Then
        PicY = Form1.PicChart.Height + 100
    End If
    
    If PicY < -100 Then
        PicY = -100
    End If

    Form1.PicChart.PSet (PicX, PicY), ColorCoord


End Sub

Public Sub DrawLine(X As Double, Y As Double, LclColor As Long)
    Dim PicX As Double
    Dim PicY As Double
    
    PicX = (X * GlbScaleX) + GlbOffX
    PicY = Form1.PicChart.Height - (Y * GlbScaleY) - GlbOffY
    
    If PicX > Form1.PicChart.Width + 100 Then
        PicX = Form1.PicChart.Width + 100
    End If
    
    If PicX < -100 Then
        PicX = -100
    End If
    
    If PicY > Form1.PicChart.Height + 100 Then
        PicY = Form1.PicChart.Height + 100
    End If
    
    If PicY < -100 Then
        PicY = -100
    End If
    
    Form1.PicChart.Line (Form1.PicChart.CurrentX, Form1.PicChart.CurrentY)-(PicX, PicY), LclColor
'    Form1.PicChart.Line (Form1.PicChart.CurrentX, Form1.PicChart.CurrentY)-((idx * GlbScaleX) + GlbOffX, Form1.PicChart.Height - (ChartArray(idx).SD * GlbScaleY) - GlbOffY), DistanceColor

End Sub


Public Sub DisplayTestPattern()
    Dim DistanceColor As Long
    
'''    DistanceColor = vbBlue
'''    DrawStart 0, 0, DistanceColor
'''    DrawLine 0, 1000, DistanceColor
'''    DrawLine 2000, 1000, DistanceColor
'''    DrawLine 2000, 0, DistanceColor
'''    DrawLine 0, 0, DistanceColor
'''
'''    DrawStart 0, 0, DistanceColor
'''    DrawLine 2000, 1000, DistanceColor
'''
'''    DrawStart 0, 1000, DistanceColor
'''    DrawLine 2000, 0, DistanceColor

End Sub



