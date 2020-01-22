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
    Form1.PicChart.Line (GlbOffX, Form1.PicChart.Height - GlbOffY)-(GlbOffX + 7000, Form1.PicChart.Height - GlbOffY), ColorCoord
    Form1.PicChart.Line (GlbOffX, Form1.PicChart.Height - GlbOffY)-(GlbOffX, Form1.PicChart.Height - (GlbOffY + 2000)), ColorCoord
    
End Sub


Public Sub ReadChartFile()
        Dim ChartFile As Integer
        Dim Zeile As String
        Dim ChartEntities() As String
        Dim idx As Long
        
        ReDim ChartArray(0 To 0)

        On Error GoTo ReadChartFileErr
        
'        ChartFilename =
        ChartFile = FreeFile
'        Open ReadChartFileErr For Binary Access Read As Chartfile
        Open ChartFilename For Input As ChartFile
        
    While Not EOF(ChartFile)
        Line Input #ChartFile, Zeile
        SepariereString Zeile, ChartEntities, ";"
        idx = UBound(ChartArray)
        ChartArray(idx).Date = ChartEntities(0)
        ChartArray(idx).Value = Zahl(ChartEntities(4))
                
        ReDim Preserve ChartArray(0 To UBound(ChartArray) + 1)
    Wend
    ReDim Preserve ChartArray(0 To UBound(ChartArray) - 1)
    Close ChartFile
      
     Exit Sub
ReadChartFileErr:
    MsgBox ChartFilename & vbCr & Err.Description, , "xxxxx"
End Sub


Public Sub DisplayChart()
'''    Dim idx As Long
'''
'''    'idx 0 is the head line
'''    idx = 1
'''    Form1.PicChart.PSet ((GlbOffX + idx) * GlbScaleX, Form1.PicChart.Height - (GlbOffY + ChartArray(idx).Value * 10)), ColorCoord
'''
'''
'''    For idx = 1 To UBound(ChartArray)
'''        Form1.PicChart.Line (Form1.PicChart.CurrentX, Form1.PicChart.CurrentY)-((GlbOffX + idx) * GlbScaleX, Form1.PicChart.Height - (GlbOffY + ChartArray(idx).Value * GlbScaleY)), ColorCoord
'''    Next idx
'''
'''
'''    'idx 0 is the head line
'''    idx = 1
'''    Form1.PicChart.PSet ((GlbOffX + idx) * GlbScaleX, Form1.PicChart.Height - (GlbOffY + ChartArray(idx).Value * 10)), ColorCoord
'''
'''
'''    Dim DistanceColor As Long
'''    For idx = 1 To UBound(ChartArray)
'''        If ChartArray(idx).Distance > 0 Then
'''            DistanceColor = vbGreen
'''        Else
'''             DistanceColor = vbBlue
'''       End If
'''        Form1.PicChart.Line (Form1.PicChart.CurrentX, Form1.PicChart.CurrentY)-((GlbOffX + idx) * GlbScaleX, Form1.PicChart.Height - (GlbOffY + ChartArray(idx).SD * GlbScaleY)), DistanceColor
'''    Next idx
    Dim idx As Long

    'idx 0 is the head line
    idx = 1

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
    
    
'''    Form1.PicChart.PSet ((idx * GlbScaleX) + GlbOffX, Form1.PicChart.Height - (ChartArray(idx).Value * GlbScaleY) - GlbOffY), ColorCoord
'''
'''
'''    Dim DistanceColor As Long
'''    For idx = 1 To UBound(ChartArray)
'''        If ChartArray(idx).Distance > 0 Then
'''            DistanceColor = vbGreen
'''        Else
'''             DistanceColor = vbBlue
'''       End If
'''       Form1.PicChart.Line (Form1.PicChart.CurrentX, Form1.PicChart.CurrentY)-((idx * GlbScaleX) + GlbOffX, Form1.PicChart.Height - (ChartArray(idx).SD * GlbScaleY) - GlbOffY), DistanceColor
'''    Next idx

    idx = 1
    DrawStart CDbl(idx), ChartArray(idx).Value, vbBlack
    For idx = 1 To UBound(ChartArray)
        DrawLine CDbl(idx), ChartArray(idx).Account, vbBlack
    Next idx
    
    
'    For idx = 1 To UBound(ChartArray)
'        Form1.PicChart.Line (Form1.PicChart.CurrentX, Form1.PicChart.CurrentY)-((idx * GlbScaleX) + GlbOffX, Form1.PicChart.Height - (ChartArray(idx).Account * Form1.VS_ScaleY.Value / 5) - GlbOffY), vbBlack
'    Next idx


End Sub


Public Sub MovingAverage(Length As Long)
    Dim idx As Long
    Dim i As Long
    Dim Sum As Double
    Dim Average As Double
    Dim Distance As Double
    
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
        Distance = (ChartArray(idx).Value - Average) / ChartArray(idx).Value
        ChartArray(idx).Distance = Distance
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
'=====================================================================
Public Sub Analyse_02()
    Dim idx As Long
'    Dim BuyNow As Boolean
    Dim SharePrice As Double
    Dim Rise As Boolean
    Dim ReadyForFirstRise As Boolean
    

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
            If ReadyForFirstRise Then
                ChartArray(idx).Trend = "Rise"
                If Rise = False Then
                    StartSharePrice = ChartArray(idx).Value
                    StartAccount = ChartArray(idx - 1).Account
    '                BuyNow = True
                Else
    '                BuyNow = False
                End If
                
                Rise = True
            End If
        Else
             ChartArray(idx).Trend = "Hold"
            If Rise = True Then
                ChartArray(idx - 1).Account = ChartArray(idx - 1).Account * 0.985
            End If
            ReadyForFirstRise = True
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

