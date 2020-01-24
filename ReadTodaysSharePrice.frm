VERSION 5.00
Begin VB.Form ReadTodaysSharePrice 
   Caption         =   "Form2"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   LinkTopic       =   "Form2"
   ScaleHeight     =   7800
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   600
      TabIndex        =   13
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   480
   End
   Begin VB.TextBox T_Search 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "Name, ISISN, WKN"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton C_ReadSingleShare 
      Caption         =   "     Read      single share"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   7275
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
   Begin VB.CommandButton C_ReadAllShares 
      Caption         =   "Read all shares"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label L_WebPage 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label L_Name 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Aktueller Kurs:"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "WKN"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "ISIN"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label L_SharePrice 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label L_WKN 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label L_ISIN 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   3960
      Width           =   1815
   End
End
Attribute VB_Name = "ReadTodaysSharePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Share_Download_idx As Long

Private Sub C_ReadAllShares_Click()

    Share_Download_idx = LBound(CompanyListArray)
    Timer1.Enabled = True
    List1.Clear


'''    List1.Clear
'''
'''    Dim i As Long
'''    For i = LBound(CompanyListArray) To UBound(CompanyListArray)
'''        Read_Peketec (CompanyListArray(i).WKN)
'''        L_Name = CompanyListArray(i).Name
'''        L_WKN = CompanyListArray(i).WKN
'''        L_ISIN = CompanyListArray(i).ISIN
'''        List1.AddItem CompanyListArray(i).Name & "  " & CompanyListArray(i).WKN & "  " & L_SharePrice
'''        WriteHistoryFile CompanyListArray(i).WKN, Date, L_SharePrice
'''    Next i
    
End Sub

Private Sub C_ReadSingleShare_Click()
    Dim Result As Double
    List1.Clear
    Dim i As Long
   
    For i = LBound(CompanyListArray) To UBound(CompanyListArray)
        If InStr(1, CompanyListArray(i).Name, T_Search, vbTextCompare) > 0 _
            Or InStr(1, CompanyListArray(i).WKN, T_Search, vbTextCompare) > 0 _
            Or InStr(1, CompanyListArray(i).ISIN, T_Search, vbTextCompare) _
        Then
            Result = Read_Peketec(CompanyListArray(i).WKN)
            If Result = 0 Then
                L_Name = CompanyListArray(i).Name
                L_WKN = CompanyListArray(i).WKN
                L_ISIN = CompanyListArray(i).ISIN
                List1.AddItem CompanyListArray(i).Name & "  " & CompanyListArray(i).WKN & "  " & L_SharePrice
            Else
                L_Name = CompanyListArray(i).Name
                L_WKN = CompanyListArray(i).WKN
                L_ISIN = CompanyListArray(i).ISIN
                List1.AddItem CompanyListArray(i).Name & "  " & CompanyListArray(i).WKN & "  " & L_SharePrice
            End If
        End If
    Next i
    
End Sub


Public Sub ReadCompanyListFile()
    Dim CompanyListFilename As String
    Dim CompanyListFile As Integer
    Dim Zeile As String
    Dim CompanyListEntities() As String
    Dim idx As Long
    
    ReDim CompanyListArray(0 To 0)
    List1.Clear
    List2.Clear
    
    On Error GoTo ReadCompanyListFileErr
    
    CompanyListFilename = App.Path & "\ISIN-WKN.txt"
    CompanyListFile = FreeFile
    Open CompanyListFilename For Input As CompanyListFile
        
    While Not EOF(CompanyListFile)
        Line Input #CompanyListFile, Zeile
        If Zeile <> "" Then
            List1.AddItem Zeile
            SepariereString Zeile, CompanyListEntities, vbTab
            idx = UBound(CompanyListArray)
            CompanyListArray(idx).Name = CompanyListEntities(0)
            CompanyListArray(idx).WKN = CompanyListEntities(1)
            CompanyListArray(idx).ISIN = CompanyListEntities(2)
            
            'Search doubbles
            Dim i As Long
            For i = 0 To UBound(CompanyListArray) - 1
                If CompanyListArray(i).WKN = CompanyListArray(idx).WKN Then
                    List2.AddItem Zeile
                End If
            Next i
            
            ReDim Preserve CompanyListArray(0 To UBound(CompanyListArray) + 1)
        End If
                
    Wend
    ReDim Preserve CompanyListArray(0 To UBound(CompanyListArray) - 1)
    Close CompanyListFile
      
     Exit Sub
ReadCompanyListFileErr:
    MsgBox CompanyListFilename & vbCr & Err.Description, , "xxxxx"
End Sub

Private Sub WriteHistoryFile(WKN As String, CurrentDate As String, SharePrice As String)
    Dim HistoryFileName As String
    Dim HistoryFile As Integer
    Dim i As Integer
    Dim Zeile As String
    
    HistoryFileName = App.Path & "\History\" & WKN & ".txt"
    
    HistoryFile = FreeFile                'Nächste freie DateiNr.
    On Error GoTo OpenError
    Open HistoryFileName For Append As HistoryFile
    
    Dim idx As Long
    
    Zeile = CurrentDate & "; ; ; ;" & SharePrice
    Print #HistoryFile, Zeile

    
    Close HistoryFile
    
    
    Exit Sub

OpenError:
  MsgBox HistoryFileName, , "Write error"

End Sub



Private Sub Form_Load()
    ReadCompanyListFile
End Sub

Private Sub Timer1_Timer()
    Dim i As Long
    Dim Result As Long
    
    i = Share_Download_idx
    
    Result = Read_Peketec(CompanyListArray(i).WKN)
    If Result = 0 Then
        L_Name = CompanyListArray(i).Name
        L_WKN = CompanyListArray(i).WKN
        L_ISIN = CompanyListArray(i).ISIN
        List1.AddItem CompanyListArray(i).Name & "  " & CompanyListArray(i).WKN & "  " & L_SharePrice
        WriteHistoryFile CompanyListArray(i).WKN, Date, L_SharePrice
    
        Share_Download_idx = Share_Download_idx + 1
    
        If Share_Download_idx > UBound(CompanyListArray) Then
            Timer1.Enabled = False
        End If
    Else
        List2.AddItem CompanyListArray(i).Name & " " & CompanyListArray(i).WKN
    End If
End Sub






