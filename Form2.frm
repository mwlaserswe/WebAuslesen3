VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Web"
   ClientHeight    =   8025
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17550
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   17550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton C_LastYear 
      Height          =   375
      Left            =   14640
      TabIndex        =   37
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton C_LastMonth 
      Height          =   375
      Left            =   14640
      TabIndex        =   36
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton C_LaserWeek 
      Height          =   375
      Left            =   14640
      TabIndex        =   35
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox T_SD 
      Height          =   285
      Left            =   7560
      TabIndex        =   33
      Text            =   "--"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox T_Account 
      Height          =   285
      Left            =   5520
      TabIndex        =   31
      Text            =   "--"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox T_Value 
      Height          =   285
      Left            =   3600
      TabIndex        =   29
      Text            =   "--"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox T_CursorDate 
      Height          =   285
      Left            =   1680
      TabIndex        =   28
      Text            =   "--"
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton C_HomeView 
      Height          =   375
      Left            =   14640
      TabIndex        =   26
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox T_Current_Sc_Off 
      Height          =   285
      Left            =   7560
      TabIndex        =   25
      Text            =   "--"
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox T_X_Sc_Off 
      Height          =   285
      Left            =   7560
      TabIndex        =   24
      Text            =   "--"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox T_MouseXY 
      Height          =   285
      Left            =   720
      TabIndex        =   23
      Text            =   "--"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox T_MouseCenter 
      Height          =   285
      Left            =   7560
      TabIndex        =   22
      Text            =   "--"
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox T_StartSharePrice 
      Height          =   285
      Left            =   10440
      TabIndex        =   20
      Text            =   "100"
      Top             =   7680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   13440
      TabIndex        =   19
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   8880
      TabIndex        =   18
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton C_Investhopping 
      Caption         =   "Invest Hopping"
      Height          =   555
      Left            =   7080
      TabIndex        =   17
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox T_HistoryFileName 
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3480
      Width           =   6255
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   30
      Left            =   5160
      TabIndex        =   15
      Top             =   2880
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
   End
   Begin VB.CommandButton C_RefreshFlexGrid 
      Caption         =   "Refresh FlexGrid"
      Height          =   495
      Left            =   7080
      TabIndex        =   14
      Top             =   0
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   30
      Left            =   3360
      TabIndex        =   13
      Top             =   1320
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      Cols            =   4
   End
   Begin MSFlexGridLib.MSFlexGrid FG_CompPartial 
      Height          =   3375
      Left            =   720
      TabIndex        =   12
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   4
   End
   Begin VB.CommandButton C_WriteChart 
      Caption         =   "Write Chart to file"
      Height          =   495
      Left            =   11160
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.HScrollBar HS_SD 
      Height          =   375
      LargeChange     =   10
      Left            =   600
      Max             =   300
      TabIndex        =   8
      Top             =   7560
      Width           =   3735
   End
   Begin VB.Timer Timer_DisplayUpdate 
      Interval        =   100
      Left            =   13080
      Top             =   3120
   End
   Begin VB.VScrollBar VS_ScaleY 
      Height          =   3135
      LargeChange     =   10
      Left            =   16920
      Max             =   100
      TabIndex        =   2
      Top             =   3960
      Width           =   255
   End
   Begin MSComDlg.CommonDialog DispChartDialog 
      Left            =   16800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicChart 
      Height          =   3135
      Left            =   720
      ScaleHeight     =   3075
      ScaleWidth      =   15795
      TabIndex        =   1
      Top             =   3960
      Width           =   15855
   End
   Begin VB.TextBox T_InvestmentStart 
      Height          =   285
      Left            =   7560
      TabIndex        =   0
      Text            =   "200"
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "Last Year"
      Height          =   255
      Left            =   15240
      TabIndex        =   40
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Last Month"
      Height          =   255
      Left            =   15240
      TabIndex        =   39
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Last Week"
      Height          =   255
      Left            =   15240
      TabIndex        =   38
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "SD:"
      Height          =   255
      Left            =   6960
      TabIndex        =   34
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Account:"
      Height          =   255
      Left            =   4800
      TabIndex        =   32
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Value:"
      Height          =   255
      Left            =   3000
      TabIndex        =   30
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Home View"
      Height          =   255
      Left            =   15240
      TabIndex        =   27
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Start Share Price [€]"
      Height          =   255
      Left            =   8760
      TabIndex        =   21
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "Inverment Start"
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "GD"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15480
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15360
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   14520
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   13800
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Menu M_DisplayChart 
      Caption         =   "Display Chart"
   End
   Begin VB.Menu M_Chartlist 
      Caption         =   "Chart List"
   End
   Begin VB.Menu M_ReadTodaysSharePrice 
      Caption         =   "Read today's share price"
   End
   Begin VB.Menu M_Web 
      Caption         =   "Web"
      Begin VB.Menu M_ScanWebForWKN 
         Caption         =   "Scan Web for WKN"
      End
      Begin VB.Menu M_SaveWebpageAsHTML 
         Caption         =   "Save Web page as HTML"
      End
      Begin VB.Menu M_ReadSingleShareValue 
         Caption         =   "Read single share value"
      End
   End
   Begin VB.Menu M_Service 
      Caption         =   "Service"
      Begin VB.Menu M_CheckWeekDay 
         Caption         =   "Check Weekday"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnt As Long




Private Sub C_Investhopping_Click()
    Dim idx As Long
    Dim Fullpath As String
    Dim CompPartialIdx As Long
    Dim ChartArrayIdx As Long
    Dim InvestmentStart As Long
    Dim Zeile As String
    Dim EarliestInvestStart As Long
    Dim EarliestWKN As String
    Dim EarliestCompany As String
    Dim InvestmentHold As Long
    Dim StartPriceRisePeriode As Double
Dim Cnt As Long

    InvestmentStart = 200
    StartPriceRisePeriode = T_StartSharePrice
    
    Do

            EarliestInvestStart = 99999999
        
            '*** walk all companies in CompPartialLstArr
            For CompPartialIdx = LBound(CompPartialLstArr) To UBound(CompPartialLstArr)
                Zeile = ""
            
                Fullpath = App.Path & "\History\" & CompPartialLstArr(CompPartialIdx).WKN & ".txt"
        '        Zeile = CompPartialLstArr(CompPartialIdx).Name
                
                ReadHistoryFile Fullpath, CompPartialLstArr(CompPartialIdx).Name
                MovingAverage (SdLength)
                Analyse_02 InvestmentStart, 0
                '*** find earlest investment start point in all companies
                For ChartArrayIdx = InvestmentStart To UBound(ChartArray)
                    If ChartArray(ChartArrayIdx).Trend = "10: Rise" Then
                        Exit For
                    End If
                Next ChartArrayIdx
        '        Zeile = Zeile & " " & ChartArrayIdx
                
                If ChartArrayIdx < EarliestInvestStart Then
                    EarliestInvestStart = ChartArrayIdx
                    EarliestWKN = CompPartialLstArr(CompPartialIdx).WKN
                    EarliestCompany = CompPartialLstArr(CompPartialIdx).Name
                End If
                
                '*** earlieset company found
                Zeile = EarliestCompany & " Start: " & EarliestWKN & " " & EarliestInvestStart
                
                
                
                
            Next CompPartialIdx
            
            Fullpath = App.Path & "\History\" & EarliestWKN & ".txt"
            ReadHistoryFile Fullpath, EarliestCompany
            MovingAverage (SdLength)
            Analyse_02 InvestmentStart, StartPriceRisePeriode
            
            '*** find next HOLD
            For ChartArrayIdx = EarliestInvestStart To UBound(ChartArray)
                If ChartArray(ChartArrayIdx).Trend = "20: Hold" Then
                    InvestmentHold = ChartArrayIdx
                    Exit For
                End If
            Next ChartArrayIdx
        
            ReDim Preserve AccountArray(0 To UBound(ChartArray))
            
            ' No-Invest periode
            For idx = InvestmentStart To EarliestInvestStart - 1
                AccountArray(idx).Name = "No Inv."
                AccountArray(idx).Account = -1
            Next idx

            ' Invest periode
            For idx = EarliestInvestStart To InvestmentHold
                AccountArray(idx) = ChartArray(idx)
            Next idx
        
            '*** prepare next investment start
            StartPriceRisePeriode = ChartArray(InvestmentHold).Account
            InvestmentStart = ChartArrayIdx + 1
            
            Zeile = EarliestCompany & " Start: " & EarliestWKN & " " & EarliestInvestStart & ";  Stop: " & ChartArrayIdx
        
        
            List1.AddItem Zeile
            
            Cnt = Cnt + 1
            T_HistoryFileName.Text = Cnt
            
            DoEvents
            
    Loop While InvestmentStart < 1000


    WriteAccountFile App.Path & "\Account.txt"




End Sub

Private Sub C_LaserWeek_Click()
     ScaleChart UBound(ChartArray), 5
End Sub

Private Sub C_LastMonth_Click()
     ScaleChart UBound(ChartArray), 22
End Sub

Private Sub C_LastYear_Click()
    ScaleChart UBound(ChartArray), 261
End Sub


Private Sub ScaleChart(idx_End As Double, idx_Length As Double)
    ' Gleichungssystem:
    '     0     = m * idx-Start + t
    ' Pic.Width = m * idx-End   + t
    '
    ' m * idx-Start + t -    0      = 0
    ' m * idx-End   + t - Pic.Width = 0
    '
    ' Gleichungssystem:
    ' A(1, 1) = idx_Start: A(1, 2) = 1: A(1, 3) = 0
    ' A(2, 1) = idx_End:   A(2, 2) = 1: A(2, 3) = Pic.Width
    Dim idx_Start As Double
    
    idx_Start = idx_End - idx_Length
    
    A(1, 1) = idx_Start: A(1, 2) = 1: A(1, 3) = 0
    A(2, 1) = idx_End:   A(2, 2) = 1: A(2, 3) = PicChart.Width

    GaussPivot A, X, 2
    
    GlbScaleX = X(1)
    GlbOffX = X(2)

    RefreshChart
End Sub








Private Sub C_WriteChart_Click()
    Dim ChartFilename As String
    Dim ChartFile As Integer
    Dim idx As Long
    Dim Zeile As String
    
    On Error GoTo OpenError
    
    ChartFilename = App.Path & "\Chart.txt"
    ChartFile = FreeFile
    Open ChartFilename For Output As ChartFile
    
    For idx = 0 To UBound(ChartArray)
        Zeile = idx _
                & vbTab & ChartArray(idx).Date _
                & vbTab & ChartArray(idx).Value _
                & vbTab & ChartArray(idx).SD _
                & vbTab & ChartArray(idx).Distance _
                & vbTab & ChartArray(idx).Account _
                & vbTab & ChartArray(idx).Trend
         Print #ChartFile, Zeile
    Next idx
           
    Close ChartFile
    
    Exit Sub

OpenError:
    MsgBox ChartFilename, , "Write error"

End Sub




Private Sub C_RefreshFlexGrid_Click()
     ArrayToFlexFrid CompPartialLstArr
End Sub


Private Sub Command1_Click()
            ReDim Preserve AccountArray(0 To 2)
            
AccountArray(0).Account = 77.8
AccountArray(0).Date = "2020-01-22"
AccountArray(0).Distance = 0.4
AccountArray(0).Name = "123456789"
AccountArray(0).SD = 34
AccountArray(0).Trend = "30: Rise"
AccountArray(0).Value = 75
AccountArray(0).WKN = "123456"

AccountArray(1).Account = 77.8
AccountArray(1).Date = "2020-01-22"
AccountArray(1).Distance = 0.4
AccountArray(1).Name = "ABCDEFGHI"
AccountArray(1).SD = 34
AccountArray(1).Trend = "30: Rise"
AccountArray(1).Value = 75
AccountArray(1).WKN = "123456"

          
            


    WriteAccountFile App.Path & "\Account.txt"

End Sub

Private Sub Command4_Click()

End Sub

'''Private Sub FG_CompPartial_Click()
'''    Dim Fullpath As String
'''
'''    FG_CompPartial.Col = 0
'''    Form1.Caption = FG_CompPartial.Text
'''
'''    ' FG_CompPartial.Row is cursor
'''    FG_CompPartial.Col = 1  ' Point to WKN columnn
'''    Fullpath = App.Path & "\History\" & FG_CompPartial.Text & ".txt"
'''    T_HistoryFileName.Text = Fullpath
'''
'''    FG_CompPartial.Col = 0  ' Point to company name columnn
'''    ReadHistoryFile Fullpath, FG_CompPartial.Text
'''
'''    RefreshChart
'''End Sub


Private Sub FG_CompPartial_SelChange()
    Dim Fullpath As String
    
    FG_CompPartial.Col = 0
    Form1.Caption = FG_CompPartial.Text
    
    ' FG_CompPartial.Row is cursor
    FG_CompPartial.Col = 1  ' Point to WKN columnn
    Fullpath = App.Path & "\History\" & FG_CompPartial.Text & ".txt"
    T_HistoryFileName.Text = Fullpath
    
    FG_CompPartial.Col = 0  ' Point to company name columnn
    ReadHistoryFile Fullpath, FG_CompPartial.Text

    RefreshChart
End Sub


Private Sub Form_Load()

    Dim rLeadAngleDeg
    Dim rLeadAngleRad
    Dim rFlipAngleDeg
    Dim rFlipAngleRad
    Dim rRoboRotAngleDeg
    Dim rRoboRotAngleRad
    Dim rLineAngleDeg
    Dim rLineAngleRad
    Dim OffsetCamX
    Dim OffsetCamY
    Dim PosAfterLeadX
    Dim PosAfterLeadY
    Dim PosAfterFlipX
    Dim PosAfterFlipY
    Dim PosAfterRotationX
    Dim PosAfterRotationY
    Dim POSXY_X
    Dim POSXY_Y
    
    Dim Pi
    


    ' OMO's Rechnug:
    ' Inits
    rLeadAngleDeg = -56
    rFlipAngleDeg = -90
    rRoboRotAngleDeg = -21.1 + 180
    rLineAngleDeg = -21.1
    
    ' Camera Offset:
    OffsetCamX = 5
    OffsetCamY = 3
    
    ' Degree To Radian
    Pi = 3.14159265359
    rLeadAngleRad = rLeadAngleDeg * Pi / 180
    rFlipAngleRad = rFlipAngleDeg * Pi / 180
    rRoboRotAngleRad = rRoboRotAngleDeg * Pi / 180
    rLineAngleRad = rLineAngleDeg * Pi / 180
    
    ' Lead Rotation
    PosAfterLeadX = OffsetCamX * Cos(rLeadAngleRad) - OffsetCamY * Sin(rLeadAngleRad)
    PosAfterLeadY = OffsetCamX * Sin(rLeadAngleRad) + OffsetCamY * Cos(rLeadAngleRad)
    
    ' Flip
    PosAfterFlipX = PosAfterLeadX * Cos(2 * rFlipAngleRad) + PosAfterLeadY * Sin(2 * rFlipAngleRad)
    PosAfterFlipY = PosAfterLeadX * Sin(2 * rFlipAngleRad) - PosAfterLeadY * Cos(2 * rFlipAngleRad)

    ' Roboter Rotation
    PosAfterRotationX = PosAfterFlipX * Cos(rRoboRotAngleRad) - PosAfterFlipY * Sin(rRoboRotAngleRad)
    PosAfterRotationY = PosAfterFlipX * Sin(rRoboRotAngleRad) + PosAfterFlipY * Cos(rRoboRotAngleRad)
    
    ' POSXY Command To ISEL Roboter (!! passive Rotation Matrix !!)
    POSXY_X = -(PosAfterRotationX * Cos(rLineAngleRad) + PosAfterRotationY * Sin(rLineAngleRad))
    POSXY_Y = -(-PosAfterRotationX * Sin(rLineAngleRad) + PosAfterRotationY * Cos(rLineAngleRad))
    

    ' Init FG_CompPartial FlexGrid
            FG_CompPartial.Cols = 5
        
            FG_CompPartial.ColWidth(0) = 1600
            FG_CompPartial.ColWidth(1) = 1000
            FG_CompPartial.ColWidth(2) = 1500
            FG_CompPartial.ColWidth(3) = 600
            FG_CompPartial.Rows = 5
            FG_CompPartial.FixedCols = 1      '1. Column fix
            'FG_CompPartial.FixedRows = 1      '1. Row fix (not used here)
            FG_CompPartial.Row = 0
            FG_CompPartial.Col = 0: FG_CompPartial.Text = "Company"
            FG_CompPartial.Col = 1: FG_CompPartial.Text = "WKN"
            FG_CompPartial.Col = 2: FG_CompPartial.Text = "ISIN"
            FG_CompPartial.Col = 3: FG_CompPartial.Text = "Index"
            FG_CompPartial.Col = 4: FG_CompPartial.Text = "Status"

    ' "\ISIN-WKN.txt" -> CompanyListArray()
    CompanyFileToArray App.Path & "\ISIN-WKN.txt", CompanyListArray
    
    
    ArrayToFlexFrid CompanyListArray
    FG_CompPartial.Rows = UBound(CompanyListArray) + 2
            'Dim idx As Long
            'For idx = 0 To UBound(CompanyListArray)
            '     FG_CompPartial.Row = idx + 1
            '    FG_CompPartial.Col = 0: FG_CompPartial.Text = CompanyListArray(idx).Name
            '    FG_CompPartial.Col = 1: FG_CompPartial.Text = CompanyListArray(idx).WKN
            '    FG_CompPartial.Col = 2: FG_CompPartial.Text = "--": FG_CompPartial.CellForeColor = RGB(0, 255, 0)
            'Next idx




    SdLength = 200
    HS_SD.Value = SdLength
    GlbScaleX = 3
    GlbScaleY = 20
    ScaleLast.X = GlbScaleX
    ScaleLast.Y = GlbScaleY


    GlbOffX = 100
    GlbOffY = 100
    OffsetLast.X = GlbOffX
    OffsetLast.Y = GlbOffY


End Sub

Private Sub HS_SD_Change()
    Label11 = HS_SD.Value
    SdLength = HS_SD.Value

    If (0 / 1) + (Not Not ChartArray) = 0 Then
      ' Array ist nicht nicht dimensioniert
      Exit Sub
    End If


    PicChart.Cls
    MovingAverage (SdLength)
    Analyse_02 Form1.T_InvestmentStart, T_StartSharePrice
    DispCoordinateSystem
    DisplayChart

End Sub

Private Sub M_Chartlist_Click()
    ChartList.Show
End Sub

Private Sub M_CheckWeekDay_Click()
    ServiceCheckWeekDay.Show
End Sub

Private Sub M_DisplayChart_Click()
    Dim HistoryFileName As String
    
    On Error GoTo errhandler
        
    DispChartDialog.CancelError = True
    DispChartDialog.InitDir = App.Path
    'DispChartDialog.Filter = "CSV-Datei (*.csv)|*.csv|Text-Datei (*.txt)|*.txt"
    DispChartDialog.Filter = "Share Files |*.csv; *.txt|"
    DispChartDialog.Filename = ""
    DispChartDialog.ShowOpen
    HistoryFileName = DispChartDialog.Filename
    
    T_HistoryFileName = HistoryFileName
    
    ReadHistoryFile HistoryFileName, ""
    
    RefreshChart

    Exit Sub
  
errhandler:
' Benutzer hat auf Abbrechen-Schaltfläche geklickt.
  Exit Sub


End Sub


Private Sub M_ReadSingleShareValue_Click()
    ReadSingleShareValue.Show
End Sub

Private Sub M_ReadTodaysSharePrice_Click()
    ReadTodaysSharePrice.Show
End Sub

Private Sub M_SaveWebpageAsHTML_Click()
    SaveWebpageAsHTML.Show
End Sub

Private Sub M_ScanWebForWKN_Click()
    ScanWebForWKN.Show
End Sub


Private Sub C_HomeView_Click()
    GlbScaleX = 3
    GlbScaleY = 20
    ScaleLast.X = GlbScaleX
    ScaleLast.Y = GlbScaleY


    GlbOffX = 100
    GlbOffY = 100
    OffsetLast.X = GlbOffX
    OffsetLast.Y = GlbOffY

    RefreshChart
End Sub

Private Sub PicChart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseDnPos.X = X
    MouseDnPos.Y = Y
    OffsetCurrent = OffsetLast
    ScaleCurrent = ScaleLast
    MouseCenterPos = MouseXY

End Sub

Private Sub PicChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim CursorDate As Date

'    Dim MouseX As Single
'    Dim MouseY As Single
    
    ' change offset
    If Button = 1 Then
        MouseMove.X = X - MouseDnPos.X
        MouseMove.Y = -(Y - MouseDnPos.Y)
        
    
        GlbOffX = OffsetCurrent.X + MouseMove.X
        GlbOffY = OffsetCurrent.Y + MouseMove.Y
        OffsetLast.X = GlbOffX
        OffsetLast.Y = GlbOffY

        RefreshChart
        
    ' change scaling
    ElseIf Button = 2 Then
        MouseMove.X = X - MouseDnPos.X
        MouseMove.Y = -(Y - MouseDnPos.Y)
        
        ' Scaling eigher X or Y
        If Abs(MouseMove.X) > Abs(MouseMove.Y) Then
            'GlbScaleX = ScaleCurrent.X + MouseMove.X / 5000
            GlbScaleX = ScaleCurrent.X + (MouseMove.X / 500) * (ScaleCurrent.X)
            If GlbScaleX < 1 Then
                GlbScaleX = 1
            End If
            ScaleLast.X = GlbScaleX
        Else
            'GlbScaleY = ScaleCurrent.Y + MouseMove.Y / 100
            GlbScaleY = ScaleCurrent.Y + MouseMove.Y / 100
            If GlbScaleY < 0.1 Then
                GlbScaleY = 0.1
            End If
            ScaleLast.Y = GlbScaleY
        End If

        'Intersection of 2 lines: t2 = x (m1 - m2) + t1
        GlbOffX = MouseCenterPos.X * (ScaleCurrent.X - GlbScaleX) + OffsetCurrent.X
'        GlbOffY = MouseCenterPos.Y * (ScaleCurrent.Y - GlbScaleY) + OffsetCurrent.Y
        OffsetLast.X = GlbOffX

T_Current_Sc_Off = "X current sc: " & ScaleCurrent.X & "   X current off: " & OffsetCurrent.X
T_MouseCenter = "X MouseCenterPos: " & MouseCenterPos.X
DoEvents


        RefreshChart
    Else
        If GlbScaleX <> 0 And GlbScaleY <> 0 Then
            MouseXY.X = (X - GlbOffX) / GlbScaleX
            MouseXY.Y = (Y - (PicChart.Height - GlbOffY)) / -GlbScaleY
            T_MouseXY.Text = Format(MouseXY.X, "0") & " " & Format(MouseXY.Y, "0")
            CursorDate = DateSerial(2000, 1, 1) + CInt(MouseXY.X) - 1
            ' T_CursorDate = CursorDate
            
            ' Array ist nicht nicht dimensioniert
            If (0 / 1) + (Not Not ChartArray) <> 0 Then
                If MouseXY.X <= UBound(ChartArray) And MouseXY.X >= LBound(ChartArray) Then
                  T_CursorDate = ChartArray(MouseXY.X).Date
                  T_Value = ChartArray(MouseXY.X).Value
                  T_Account = Format(ChartArray(MouseXY.X).Account, "0.00")
                  T_SD = Format(ChartArray(MouseXY.X).SD, "0.00")
                End If
            
            End If
            
        End If
        T_X_Sc_Off.Text = "X-Scale: " & GlbScaleX & "   X-Offset: " & GlbOffX
'        T_Y_Sc_Off.Text = "Y-Scale: " & GlbScaleY & "   Y-Offset: " & GlbOffY
    End If

End Sub


Private Sub VS_ScaleY_Change()
'    GlbScaleY = VS_ScaleY.Value
'    PicChart.Cls
'    C_DrawChart_Click
End Sub


Private Sub CompanyFileToArray(CompanyListFilename As String, CompanyListArray() As ShareItem)
    Dim CompanyListFile As Integer
    Dim Zeile As String
    Dim CompanyListEntities() As String
    Dim idx As Long
    
    ReDim CompanyListArray(0 To 0)
'    MyList.Clear
'    List2.Clear
    
    On Error GoTo ReadCompanyListFileErr
    
    'CompanyListFilename = App.Path & "\ISIN-WKN.txt"
    CompanyListFile = FreeFile
    Open CompanyListFilename For Input As CompanyListFile
        
    While Not EOF(CompanyListFile)
        Line Input #CompanyListFile, Zeile
        If Zeile <> "" Then
            'MyList.AddItem Zeile
            SepariereString Zeile, CompanyListEntities, vbTab
            idx = UBound(CompanyListArray)
            CompanyListArray(idx).Name = CompanyListEntities(0)
            CompanyListArray(idx).WKN = CompanyListEntities(1)
            CompanyListArray(idx).ISIN = CompanyListEntities(2)
            If UBound(CompanyListEntities) >= 3 Then
                CompanyListArray(idx).Index = CompanyListEntities(3)
            End If

            
'''            'Search doubbles
'''            Dim i As Long
'''            For i = 0 To UBound(CompanyListArray) - 1
'''                If CompanyListArray(i).WKN = CompanyListArray(idx).WKN Then
'''                    List2.AddItem Zeile
'''                End If
'''            Next i
            
            ReDim Preserve CompanyListArray(0 To UBound(CompanyListArray) + 1)
        End If
                
    Wend
    ReDim Preserve CompanyListArray(0 To UBound(CompanyListArray) - 1)
    Close CompanyListFile
      
     Exit Sub
ReadCompanyListFileErr:
    MsgBox CompanyListFilename & vbCr & Err.Description, , "xxxxx"
End Sub

Private Sub ArrayToFlexFrid(CompanyListArray() As ShareItem)
    Dim idx As Long
    
    If (0 / 1) + (Not Not CompanyListArray) = 0 Then
      ' Array ist nicht nicht dimensioniert
      Exit Sub
    End If


    FG_CompPartial.Rows = UBound(CompanyListArray) + 2
    For idx = 0 To UBound(CompanyListArray)
        FG_CompPartial.Row = idx + 1
        FG_CompPartial.Col = 0: FG_CompPartial.Text = CompanyListArray(idx).Name
        FG_CompPartial.Col = 1: FG_CompPartial.Text = CompanyListArray(idx).WKN
        FG_CompPartial.Col = 2: FG_CompPartial.Text = CompanyListArray(idx).ISIN
        FG_CompPartial.Col = 3: FG_CompPartial.Text = CompanyListArray(idx).Index
        FG_CompPartial.Col = 4: FG_CompPartial.Text = "--": FG_CompPartial.CellForeColor = RGB(0, 255, 0)
    Next idx
End Sub



Private Sub RefreshChart()
    PicChart.Cls
    MovingAverage (SdLength)
    Analyse_02 Form1.T_InvestmentStart, T_StartSharePrice
    DispCoordinateSystem
    DisplayChart
End Sub

