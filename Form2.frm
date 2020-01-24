VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   Begin VB.CommandButton C_WriteChart 
      Caption         =   "Write Chart to file"
      Height          =   495
      Left            =   11160
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.HScrollBar HS_SD 
      Height          =   375
      LargeChange     =   10
      Left            =   720
      Max             =   300
      TabIndex        =   10
      Top             =   7200
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
      TabIndex        =   4
      Top             =   3960
      Width           =   255
   End
   Begin MSComDlg.CommonDialog DispChartDialog 
      Left            =   15960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton C_DrawChart 
      Caption         =   "Draw Chart 1"
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
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
      Left            =   8280
      TabIndex        =   0
      Text            =   "200"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Inverment Start"
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "GD"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15480
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15360
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   14520
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   13800
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label L_ChartFile 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3480
      Width           =   6975
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnt As Long

Private Sub C_DrawChart_Click()
    PicChart.Cls
    DispCoordinateSystem
    ReadChartFile
    MovingAverage (SdLength)
    Analyse_02
    DisplayChart
End Sub

'''Private Sub C_ReadWKN_Click()
'''
'''Dim WKN_Start As Long
'''Dim WKN_String As String
'''Dim HtmlCode As String
'''Dim URL_String As String
'''Dim Zeile1 As String
'''Dim Zeile2 As String
'''
'''
'''
'''
'''
'''Dim idx As Long
'''
'''WKN_Start = CLng(Zahl(T_WKNStart))
'''WriteWknFile ("WKN-Start:" & WKN_Start)
'''
'''    Dim SearchItem As String
'''    Dim EndString As String
'''
'''
'''
'''
'''For idx = WKN_Start To WKN_Start + 1000000
'''    WKN_String = Format(idx, "000000")
'''    Label4 = WKN_String
'''    URL_String = "https://peketec.de/portal/suche/?s=" & WKN_String
'''
'''    HtmlCode = GetHTMLCode(URL_String)
'''    'SaveQuelltext HtmlCode, App.Path & "\HTML-Code.HTML"
'''
'''    'Bezeichnung ISIN WPKN extrahieren
'''    SearchItem = "Bezeichnung"
'''    EndString = "</table>"
'''    Zeile1 = ExtraxtValue(HtmlCode, SearchItem, EndString)
'''    Text1 = Zeile1
'''
'''        SearchItem = "portal"
'''        EndString = "</tr>"
'''        Zeile2 = ExtraxtValue(Zeile1, SearchItem, EndString)
'''        Label6 = Zeile2
'''
'''    WriteWknFile (WKN_String & ":" & vbTab & Zeile2)
'''
'''
'''Next idx





'''End Sub


'''Private Sub GetLastWkn()
'''        Dim WknFileName As String
'''        Dim WknFile As Integer
'''        Dim Zeile As String
'''        Dim WknEntities() As String
'''        Dim idx As Long
'''        Dim Number As Long
'''
'''        Dim LastWkn As Long
'''
'''
'''        On Error GoTo ReadWknFileErr
'''
'''        WknFileName = App.Path & "\WKN.txt"
'''        WknFile = FreeFile
''''        Open ReadWknFileErr For Binary Access Read As Wknfile
'''        Open WknFileName For Input As WknFile
'''
'''
'''    While Not EOF(WknFile)
'''        Line Input #WknFile, Zeile
'''        Number = CLng(Zahl(Zeile))
'''        If Number > 0 Then
'''            LastWkn = Number
'''        End If
'''    Wend
'''
'''    T_WKNStart = LastWkn + 1
'''
'''    Close WknFile
'''
'''     Exit Sub
'''ReadWknFileErr:
'''    MsgBox WknFileName & vbCr & Err.Description, , "xxxxx"
'''
'''End Sub


'''Private Sub WriteWknFile(Zeile As String)
'''    Dim WknFileName As String
'''    Dim WknFile As Integer
'''    Dim i As Integer
''''    Dim Prexit As String
'''
'''    WknFileName = App.Path & "\WKN.txt"
'''
'''    WknFile = FreeFile                'Nächste freie DateiNr.
'''    On Error GoTo OpenError
'''    Open WknFileName For Append As WknFile
'''
'''    Dim idx As Long
'''
'''    Print #WknFile, Zeile
'''
'''
'''    Close WknFile
'''
'''
'''    Exit Sub
'''
'''OpenError:
'''  MsgBox WknFileName, , "Write error"
'''
'''End Sub








Private Sub C_WriteChart_Click()
    Dim ChartFilename As String
    Dim ChartFile As Integer
    Dim i As Integer
'    Dim Prexit As String
    
    ChartFilename = App.Path & "\Chart.txt"
    
    ChartFile = FreeFile                'Nächste freie DateiNr.
    On Error GoTo OpenError
    Open ChartFilename For Output As ChartFile
    
    Dim idx As Long
    Dim Zeile As String
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
        
'        Date As String
'    Value As Double
'    SD As Double
'    Distance As Double      ' Distance to moving average
'    Account As Double
'    Trend As String
    
    Close ChartFile
    
    
    Exit Sub

OpenError:
  MsgBox ChartFilename, , "Write error"

End Sub







Private Sub Command2_Click()
    Read_Peketec (T_WKN)
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
    



    


    SdLength = 20
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

End Sub

Private Sub M_Chartlist_Click()
    ChartList.Show
End Sub

Private Sub M_DisplayChart_Click()
'    Dim Befehlszeile As String
'    Dim EditFileName As String
    
    ' CancelError ist auf True gesetzt.
    On Error GoTo errhandler
    
    
    DispChartDialog.CancelError = True
    DispChartDialog.InitDir = App.Path
    'DispChartDialog.Filter = "CSV-Datei (*.csv)|*.csv|Text-Datei (*.txt)|*.txt"
    DispChartDialog.Filter = "Share Files |*.csv; *.txt|"
    DispChartDialog.Filename = ""
    DispChartDialog.ShowOpen
    ChartFilename = DispChartDialog.Filename
    
    L_ChartFile = ChartFilename
'    DefaultPath = Pfad$(DefaultStarKatalog)
'    DefaultStarKatalog = EditFileName
    '  LastEditFileName = EditFileName
    
'    INISetValue IniFileName, "Basics", "DefaultPath", DefaultPath
'    INISetValue IniFileName, "Basics", "DefaultStarKatalog", DefaultStarKatalog
'    LoadAlignmetStarFile
    

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

Private Sub PicChart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Label4 = Shift
'    Label5 = X
'    Label6 = Y
'    Label7 = Button
'    If Button = 1 Then
'        MouseDnPos.X = X
'        MouseDnPos.Y = Y
'    End If
'
'    If Button = 2 Then
'    End If
    
    MouseDnPos.X = X
    MouseDnPos.Y = Y
    OffsetCurrent = OffsetLast
    ScaleCurrent = ScaleLast

End Sub

Private Sub PicChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Label4 = Shift
'    Label5 = X
'    Label6 = Y
'    Label7 = Button
    
    If Button = 1 Then
        MouseMove.X = X - MouseDnPos.X
        MouseMove.Y = -(Y - MouseDnPos.Y)
        
    
        GlbOffX = OffsetCurrent.X + MouseMove.X
        GlbOffY = OffsetCurrent.Y + MouseMove.Y
        OffsetLast.X = GlbOffX
        OffsetLast.Y = GlbOffY

'        PicChart.Cls
        C_DrawChart_Click
    End If

    If Button = 2 Then
        MouseMove.X = X - MouseDnPos.X
        MouseMove.Y = -(Y - MouseDnPos.Y)
        
    
        GlbScaleX = ScaleCurrent.X + MouseMove.X / 5000
        GlbScaleY = ScaleCurrent.Y + MouseMove.Y / 100
        ScaleLast.X = GlbScaleX
        ScaleLast.Y = GlbScaleY

        'Intersection of 2 lines: t2 = x (m1 - m2) + t1
        GlbOffX = (MouseDnPos.X - OffsetCurrent.X) * (ScaleCurrent.X - GlbScaleX) + OffsetCurrent.X
        OffsetLast.X = GlbOffX
''        Timer1_Timer
        
'        PicChart.Cls
        C_DrawChart_Click
    End If

End Sub

'''Private Sub Timer1_Timer()
'''    ' reconnect to webpage
'''    If Not Timer2.Enabled Then
'''        'C_ReadWKN1_Click
'''    End If
'''End Sub


'''Private Sub C_ReadWKN1_Click()
'''
'''    GetLastWkn
'''    WknReadCount = CDbl(Zahl(T_WknReadCount)) - 1
'''    DelayTime = CDbl(Zahl(T_DelayTime))
'''
'''    WKS_Download_idx = CLng(Zahl(T_WKNStart))
'''    WriteWknFile ("WKN-Start:" & WKS_Download_idx)
'''    Timer2.Enabled = Not Timer2.Enabled
'''
'''    DownloadPause = 0
'''End Sub




'''Private Sub Timer_DisplayUpdate_Timer()
'''    If Timer2.Enabled Then
'''        C_ReadWKN1.BackColor = vbGreen
'''    Else
'''        C_ReadWKN1.BackColor = vbWhite
'''    End If
'''
'''End Sub



Private Sub VS_ScaleY_Change()
'    GlbScaleY = VS_ScaleY.Value
'    PicChart.Cls
'    C_DrawChart_Click
End Sub
