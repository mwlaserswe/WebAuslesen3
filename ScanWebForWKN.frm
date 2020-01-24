VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ScanWebForWKN 
   Caption         =   "Form2"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11205
   LinkTopic       =   "Form2"
   ScaleHeight     =   3465
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_DisplayUpdate 
      Interval        =   100
      Left            =   10560
      Top             =   1920
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox T_WKNStart 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Text            =   "500000"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1125
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "ScanWebForWKN.frx":0000
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton C_ReadWKN1 
      Caption         =   "Read WKN"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10560
      Top             =   240
   End
   Begin VB.TextBox T_DelayTime 
      Height          =   285
      Left            =   4320
      TabIndex        =   1
      Text            =   "1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox T_WknReadCount 
      Height          =   285
      Left            =   6240
      TabIndex        =   0
      Text            =   "100"
      Top             =   600
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   10560
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog DispChartDialog 
      Left            =   9000
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Label Label5 
      Caption         =   "Delay [sec]"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "count"
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.Label L_ErrorCnt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   8640
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label L_AccessCnt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   8640
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "Error Cnt:"
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "Access Cnt:"
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "ScanWebForWKN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub C_ReadWKN1_Click()

    GetLastWkn
    WknReadCount = CDbl(Zahl(T_WknReadCount)) - 1
    DelayTime = CDbl(Zahl(T_DelayTime))
    
    WKS_Download_idx = CLng(Zahl(T_WKNStart))
    WriteWknFile ("WKN-Start:" & WKS_Download_idx)
    Timer2.Enabled = Not Timer2.Enabled
    
    DownloadPause = 0
End Sub


Private Sub Timer1_Timer()
    ' reconnect to webpage
    If Not Timer2.Enabled Then
        'C_ReadWKN1_Click
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
    L_AccessCnt = WKS_Download_idx - CLng(Zahl(T_WKNStart))
    If DownloadPause >= WknReadCount Then
        Timer2.Enabled = False
        DownloadPause = 0
    End If


    WKS_Download_idx = WKS_Download_idx + 1
    DownloadPause = DownloadPause + 1

End Sub


Private Sub Timer_DisplayUpdate_Timer()
    If Timer2.Enabled Then
        C_ReadWKN1.BackColor = vbGreen
    Else
        C_ReadWKN1.BackColor = vbWhite
    End If

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
