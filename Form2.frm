VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17550
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   17550
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   16920
      Top             =   1680
   End
   Begin VB.TextBox T_WKN 
      Height          =   285
      Left            =   120
      TabIndex        =   41
      Text            =   "578560"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Share Value"
      Height          =   615
      Left            =   1440
      TabIndex        =   40
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox T_WknReadCount 
      Height          =   285
      Left            =   13200
      TabIndex        =   34
      Text            =   "100"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox T_DelayTime 
      Height          =   285
      Left            =   11280
      TabIndex        =   32
      Text            =   "1"
      Top             =   240
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   16920
      Top             =   0
   End
   Begin VB.CommandButton C_ReadWKN1 
      Caption         =   "Read WKN"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1125
      Left            =   8520
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "Form2.frx":0000
      Top             =   1080
      Width           =   5655
   End
   Begin VB.TextBox T_WKNStart 
      Height          =   285
      Left            =   9960
      TabIndex        =   29
      Text            =   "500000"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton C_ReadWKN 
      Caption         =   "Read WKN   do not use"
      Height          =   495
      Left            =   6960
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton C_WriteChart 
      Caption         =   "Write Chart to file"
      Height          =   495
      Left            =   11160
      TabIndex        =   21
      Top             =   3240
      Width           =   1215
   End
   Begin VB.HScrollBar HS_SD 
      Height          =   375
      LargeChange     =   10
      Left            =   720
      Max             =   300
      TabIndex        =   20
      Top             =   7200
      Width           =   3735
   End
   Begin VB.Timer Timer_DisplayUpdate 
      Interval        =   100
      Left            =   14280
      Top             =   2880
   End
   Begin VB.VScrollBar VS_ScaleY 
      Height          =   3135
      LargeChange     =   10
      Left            =   16920
      Max             =   100
      TabIndex        =   12
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
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox PicChart 
      Height          =   3135
      Left            =   720
      ScaleHeight     =   3075
      ScaleWidth      =   15795
      TabIndex        =   9
      Top             =   3960
      Width           =   15855
   End
   Begin VB.CommandButton C_SavePageToFile 
      Caption         =   "Save Page To File"
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox T_URL 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "https://www.google.de/"
      Top             =   240
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6960
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6960
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox T_InvestmentStart 
      Height          =   285
      Left            =   8280
      TabIndex        =   1
      Text            =   "200"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "WKN"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   "Access Cnt:"
      Height          =   255
      Left            =   14400
      TabIndex        =   39
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "Error Cnt:"
      Height          =   255
      Left            =   14640
      TabIndex        =   38
      Top             =   840
      Width           =   735
   End
   Begin VB.Label L_AccessCnt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15600
      TabIndex        =   37
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label L_ErrorCnt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15600
      TabIndex        =   36
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "count"
      Height          =   255
      Left            =   14160
      TabIndex        =   35
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Delay [sec]"
      Height          =   255
      Left            =   11760
      TabIndex        =   33
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Inverment Start"
      Height          =   255
      Left            =   6840
      TabIndex        =   27
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "GD"
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label L_WebPage 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label L_ISIN 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label L_WKN 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label L_SharePrice 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   5280
      TabIndex        =   19
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15480
      TabIndex        =   18
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15360
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15360
      TabIndex        =   16
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   15360
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label L_ChartFile 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   3480
      Width           =   6975
   End
   Begin VB.Label Label3 
      Caption         =   "ISIN"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "WKN"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Aktueller Kurs:"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Menu M_DisplayChart 
      Caption         =   "Display Chart"
   End
   Begin VB.Menu M_ReadTodaysSharePrice 
      Caption         =   "Read today's share price"
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

Private Sub C_ReadWKN_Click()

Dim WKN_Start As Long
Dim WKN_String As String
Dim HtmlCode As String
Dim URL_String As String
Dim Zeile1 As String
Dim Zeile2 As String





Dim idx As Long

WKN_Start = CLng(Zahl(T_WKNStart))
WriteWknFile ("WKN-Start:" & WKN_Start)

    Dim SearchItem As String
    Dim EndString As String




For idx = WKN_Start To WKN_Start + 1000000
    WKN_String = Format(idx, "000000")
    Label4 = WKN_String
    URL_String = "https://peketec.de/portal/suche/?s=" & WKN_String
    
    HtmlCode = GetHTMLCode(URL_String)
    'SaveQuelltext HtmlCode, App.Path & "\HTML-Code.HTML"

    'Bezeichnung ISIN WPKN extrahieren
    SearchItem = "Bezeichnung"
    EndString = "</table>"
    Zeile1 = ExtraxtValue(HtmlCode, SearchItem, EndString)
    Text1 = Zeile1
    
        SearchItem = "portal"
        EndString = "</tr>"
        Zeile2 = ExtraxtValue(Zeile1, SearchItem, EndString)
        Label6 = Zeile2
    
    WriteWknFile (WKN_String & ":" & vbTab & Zeile2)


Next idx





End Sub


Private Sub GetLastWkn()
        Dim WknFileName As String
        Dim WknFile As Integer
        Dim Zeile As String
        Dim WknEntities() As String
        Dim idx As Long
        Dim Number As Long
        
        Dim LastWkn As Long
        

        On Error GoTo ReadWknFileErr
        
        WknFileName = App.Path & "\WKN.txt"
        WknFile = FreeFile
'        Open ReadWknFileErr For Binary Access Read As Wknfile
        Open WknFileName For Input As WknFile
        
        
    While Not EOF(WknFile)
        Line Input #WknFile, Zeile
        Number = CLng(Zahl(Zeile))
        If Number > 0 Then
            LastWkn = Number
        End If
    Wend
    
    T_WKNStart = LastWkn + 1
    
    Close WknFile
      
     Exit Sub
ReadWknFileErr:
    MsgBox WknFileName & vbCr & Err.Description, , "xxxxx"

End Sub


Private Sub WriteWknFile(Zeile As String)
    Dim WknFileName As String
    Dim WknFile As Integer
    Dim i As Integer
'    Dim Prexit As String
    
    WknFileName = App.Path & "\WKN.txt"
    
    WknFile = FreeFile                'Nächste freie DateiNr.
    On Error GoTo OpenError
    Open WknFileName For Append As WknFile
    
    Dim idx As Long
    
    Print #WknFile, Zeile

    
    Close WknFile
    
    
    Exit Sub

OpenError:
  MsgBox WknFileName, , "Write error"

End Sub






Private Sub C_SavePageToFile_Click()
    Dim HtmlCode As String
    HtmlCode = GetHTMLCode(T_URL)
    SaveQuelltext HtmlCode, App.Path & "\HTML-Code.HTML"
End Sub

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



Private Sub Command1_Click()
    Dim GoogleText As String
'    Dim Von As String
'    Dim Nach As String
'    Dim Zeit As String
'    Dim Stunden As String
'    Dim Minuten As String
'    Dim KM As String

    Dim WebPage As String
    Dim SearchItem As String
    Dim AktuellerKurs As String
    Dim InstrumentId As String
    Dim WKN As String
    Dim ISIN As String
'    Dim DmyString As String
    
    
    
    Dim PosStart As Long
    Dim PosEnd As Long
    Dim EndString As String
    
    WebPage = "https://www.finanzen.net/aktien/bmw-aktie"
    GoogleText = GetHTMLCode(WebPage)
    'oder https://tradingdesk.finanzen.net/aktie/DE0005190003'
  
    SaveQuelltext GoogleText, App.Path & "\GoogleText.HTML"
 
    'AktuellerKurs extrahieren
    SearchItem = "Aktueller Kurs"
    EndString = Chr$(34)
    AktuellerKurs = ExtraxtValue(GoogleText, SearchItem, EndString)
    
    'InstrumentId extrahieren
    SearchItem = "instrument-id"
    EndString = Chr$(34)
    InstrumentId = ExtraxtValue(GoogleText, SearchItem, EndString)

    
       'WSK extrahieren
       SearchItem = "WKN:"
       EndString = "/"
       WKN = ExtraxtValue(InstrumentId, SearchItem, EndString)
       'ISIN extrahieren
       SearchItem = "ISIN:"
       EndString = "</span>"
       ISIN = ExtraxtValue(InstrumentId, SearchItem, EndString)
    L_WebPage = WebPage
    L_SharePrice = AktuellerKurs
    L_WKN = WKN
    L_ISIN = ISIN
 

    WebPage = "https://www.finanzen.net/aktien/Basler-aktie"
    GoogleText = GetHTMLCode(WebPage)
    'oder https://tradingdesk.finanzen.net/aktie/DE0005190003'
  
    SaveQuelltext GoogleText, App.Path & "\GoogleText.HTML"
 
    
 
    'AktuellerKurs extrahieren
    SearchItem = "Aktueller Kurs"
    EndString = Chr$(34)
    AktuellerKurs = ExtraxtValue(GoogleText, SearchItem, EndString)
    
    'InstrumentId extrahieren
    SearchItem = "instrument-id"
    EndString = Chr$(34)
    InstrumentId = ExtraxtValue(GoogleText, SearchItem, EndString)

    
       'WSK extrahieren
       SearchItem = "WKN:"
       EndString = "/"
       WKN = ExtraxtValue(InstrumentId, SearchItem, EndString)
       'ISIN extrahieren
       SearchItem = "ISIN:"
       EndString = "</span>"
       ISIN = ExtraxtValue(InstrumentId, SearchItem, EndString)
    L_WebPage = WebPage
    L_SharePrice = AktuellerKurs
    L_WKN = WKN
    L_ISIN = ISIN



    GoogleText = GetHTMLCode("https://www.finanzen.net/aktien/deutschland-aktien-realtimekurse")
    'oder https://tradingdesk.finanzen.net/aktie/DE0005190003'
  
    SaveQuelltext GoogleText, App.Path & "\GoogleText.txt"


End Sub
 




Private Sub Command2_Click()
    Read_Peketec (T_WKN)
End Sub

Private Sub Form_Load()

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


Private Sub M_ReadTodaysSharePrice_Click()
    ReadTodaysSharePrice.Show
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
        Timer1_Timer
        
'        PicChart.Cls
        C_DrawChart_Click
    End If

End Sub

Private Sub Timer1_Timer()
    ' reconnect to webpage
    If Not Timer2.Enabled Then
        'C_ReadWKN1_Click
    End If
End Sub


Private Sub C_ReadWKN1_Click()

    GetLastWkn
    WknReadCount = CDbl(Zahl(T_WknReadCount)) - 1
    DelayTime = CDbl(Zahl(T_DelayTime))
    
    WKS_Download_idx = CLng(Zahl(T_WKNStart))
    WriteWknFile ("WKN-Start:" & WKS_Download_idx)
    Timer2.Enabled = Not Timer2.Enabled
    
    DownloadPause = 0
End Sub




Private Sub Timer_DisplayUpdate_Timer()
    If Timer2.Enabled Then
        C_ReadWKN1.BackColor = vbGreen
    Else
        C_ReadWKN1.BackColor = vbWhite
    End If

End Sub

Private Sub Timer2_Timer()

Dim WKN_Start As Long
Dim WKN_String As String
Dim HtmlCode As String
Dim URL_String As String
Dim Zeile1 As String
Dim Zeile2 As String








    Dim SearchItem As String
    Dim EndString As String


    Timer2.Interval = DelayTime * 1000

    WKN_String = Format(WKS_Download_idx, "000000")
    Label4 = WKN_String
    URL_String = "https://peketec.de/portal/suche/?s=" & WKN_String
    
    HtmlCode = GetHTMLCode(URL_String)
    
    If HtmlCode = ">>>ERROR<<<" Then
        Exit Sub
    End If
    'SaveQuelltext HtmlCode, App.Path & "\HTML-Code.HTML"

    'Bezeichnung ISIN WPKN extrahieren
    SearchItem = "Bezeichnung"
    EndString = "</table>"
    Zeile1 = ExtraxtValue(HtmlCode, SearchItem, EndString)
    Text1 = Zeile1
    
        SearchItem = "portal"
        EndString = "</tr>"
        Zeile2 = ExtraxtValue(Zeile1, SearchItem, EndString)
        Label6 = Zeile2
    
    WriteWknFile (WKN_String & ":" & vbTab & Zeile2)
    Form1.L_AccessCnt = WKS_Download_idx - CLng(Zahl(T_WKNStart))
    If DownloadPause >= WknReadCount Then
        Timer2.Enabled = False
        DownloadPause = 0
    End If


    WKS_Download_idx = WKS_Download_idx + 1
    DownloadPause = DownloadPause + 1






End Sub

Private Sub VS_ScaleY_Change()
'    GlbScaleY = VS_ScaleY.Value
'    PicChart.Cls
'    C_DrawChart_Click
End Sub
