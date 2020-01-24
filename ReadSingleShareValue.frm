VERSION 5.00
Begin VB.Form ReadSingleShareValue 
   Caption         =   "Read Single Share Value"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Text            =   "Darstelleung der Werte funktioniert nicht"
      Top             =   240
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Share Value"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox T_WKN 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "578560"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Aktueller Kurs:"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "WKN"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "ISIN"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label L_SharePrice 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label L_WKN 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label L_ISIN 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "WKN"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "ReadSingleShareValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
'''    L_WebPage = WebPage
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
'''    L_WebPage = WebPage
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

