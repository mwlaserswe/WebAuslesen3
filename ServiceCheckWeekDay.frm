VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ServiceCheckWeekDay 
   Caption         =   "Check Weekday"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5700
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Text            =   "1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Text            =   "1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Text            =   "1"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Demos"
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   4095
      Begin VB.CommandButton C_TageszahlInDatum 
         Caption         =   "Tageszahl ab 01.01.2000 in Datum"
         Height          =   855
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Text            =   "375"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Text            =   "--"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton C_WeekdayDemo 
         Caption         =   "Weekday Demo"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   "27.1.2020"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Text            =   "4"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Wochentag-Nr."
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.ListBox List2 
      Height          =   5130
      Left            =   9120
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "1"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton C_ServiceHistory 
      Caption         =   "Service History"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13440
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "- falsches Datumsformat korrigieren"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "- fehlende Einträge ergänzen"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "ServiceCheckWeekDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FinalArray() As HistoryItem

Private Type HistoryItem
    Datum As String
    Erster As String
    Hoch As String
    Tief As String
    Schlusskurs As String
    Stuecke As String
    Volumen As String
    Bemerkung As String
    
End Type


Private Sub C_WeekdayDemo_Click()
    Dim Datum1 As String

    Datum1 = Text1.Text
    ' ermittelt den Wochentag zu einem Datum. 1 = Sonntag, 2 = Montag,...
    Text2 = Weekday(Datum1)
End Sub


Private Sub C_TageszahlInDatum_Click()
    Dim intTag As Integer
    Dim TageszahlInDatum  As Date
    
    intTag = Text4.Text
    TageszahlInDatum = DateSerial(2000, 1, 1) + intTag - 1
    Text5.Text = TageszahlInDatum
End Sub

Private Sub C_ServiceHistory_Click()
    Dim idx As Long
    
    If (0 / 1) + (Not Not CompPartialLstArr) = 0 Then
      ' Array ist nicht nicht dimensioniert
      Exit Sub
    End If

    For idx = 0 To UBound(CompPartialLstArr)
        ServiceHistory CompPartialLstArr(idx).WKN
    Next idx
    
'    ServiceHistory "tst"
End Sub


Private Sub ServiceHistory(WKN As String)
    Dim Fullpath As String
    Dim HistoryFileName As String
    Dim HistoryFile As Integer
    Dim Zeile As String
    Dim HistoryEntities() As String
    Dim HistoryIdx As Long
    Dim TageszahlInDatum As String
    Dim HistoryLine As HistoryItem
    Dim DateIdx As Long
    Dim FinalIdx As Long
    Dim Today As String
    Dim Lastvalue As String
    
    ReDim FinalArray(0 To 0)

    List2.Clear
    HistoryFileName = App.Path & "\History\" & WKN & ".txt"
    List2.AddItem "**** " & WKN & " ****"
    Today = TodayFunction


    On Error GoTo ReadHistoryFileErr
    HistoryFile = FreeFile
    Open HistoryFileName For Input As HistoryFile

    Lastvalue = 0

    ' Read Headline
    HistoryIdx = 0
    DateIdx = 0
    FinalIdx = 0

    Line Input #HistoryFile, Zeile
    
    
'    List1.AddItem Zeile
    SepariereString Zeile, HistoryEntities, ";"
        HistoryLine.Datum = HistoryEntities(0)
        HistoryLine.Erster = HistoryEntities(1)
        HistoryLine.Hoch = HistoryEntities(2)
        HistoryLine.Tief = HistoryEntities(3)
        HistoryLine.Schlusskurs = HistoryEntities(4)
        If UBound(HistoryEntities) > 5 Then
            HistoryLine.Stuecke = HistoryEntities(5)
            HistoryLine.Volumen = HistoryEntities(6)
            DoEvents
        End If
    FinalArray(FinalIdx) = HistoryLine
     
    ReDim Preserve FinalArray(0 To UBound(FinalArray) + 1)
    HistoryIdx = 1
    DateIdx = 1
    FinalIdx = 1
    
    While Not EOF(HistoryFile) And (TageszahlInDatum <> Today)

        HistoryLine.Datum = ""
        HistoryLine.Erster = ""
        HistoryLine.Hoch = ""
        HistoryLine.Tief = ""
        HistoryLine.Schlusskurs = ""
        HistoryLine.Stuecke = ""
        HistoryLine.Volumen = ""
        HistoryLine.Bemerkung = ""


        Line Input #HistoryFile, Zeile
        
        TageszahlInDatum = FormatDate(DateSerial(2000, 1, 1) + DateIdx - 1)
        
        SepariereString Zeile, HistoryEntities, ";"
            HistoryLine.Datum = FormatDate(HistoryEntities(0))
            HistoryLine.Erster = HistoryEntities(1)
            HistoryLine.Hoch = HistoryEntities(2)
            HistoryLine.Tief = HistoryEntities(3)
            HistoryLine.Schlusskurs = HistoryEntities(4)
            If UBound(HistoryEntities) > 5 Then
                HistoryLine.Stuecke = HistoryEntities(5)
                HistoryLine.Volumen = HistoryEntities(6)
                DoEvents
            End If
        
        ' Check missing days
        While (TageszahlInDatum <> HistoryLine.Datum) And (TageszahlInDatum <> Today)
            Dim ds As String
            
            ds = Weekday(TageszahlInDatum)
            ' No entry for Saturday and Sunday
            If (Weekday(TageszahlInDatum) = 7) Or (Weekday(TageszahlInDatum) = 1) Then

            ' Normal working day: If line is missing, insert line and take previous value
            Else
                List2.AddItem TageszahlInDatum & "   Inserted "
                ReDim Preserve FinalArray(0 To UBound(FinalArray) + 1)
                FinalArray(FinalIdx).Datum = TageszahlInDatum
                FinalArray(FinalIdx).Schlusskurs = Lastvalue
                FinalArray(FinalIdx).Bemerkung = "Inserted"
                FinalIdx = FinalIdx + 1
            End If
            
            DateIdx = DateIdx + 1
            TageszahlInDatum = FormatDate(DateSerial(2000, 1, 1) + DateIdx - 1)
            DoEvents
        Wend
        
        
'        List1.AddItem Zeile
        
        
        ' No entry for Saturday and Sunday
        If (Weekday(HistoryLine.Datum) = 7) Or (Weekday(HistoryLine.Datum) = 1) Then
                List2.AddItem TageszahlInDatum & " --> Wochenende"
        
        ' Normal working day
        Else
            ' If there is a value = 0, replace with last value
            If HistoryLine.Schlusskurs = 0 Then
                HistoryLine.Schlusskurs = Lastvalue
                HistoryLine.Bemerkung = "Zero replaced"
                List2.AddItem TageszahlInDatum & " --> Zero replaced"
            End If
            FinalArray(FinalIdx) = HistoryLine
            Lastvalue = HistoryLine.Schlusskurs
            FinalIdx = FinalIdx + 1
            ReDim Preserve FinalArray(0 To UBound(FinalArray) + 1)
        End If
        
'        Text2.Text = Zeile
        
        HistoryIdx = HistoryIdx + 1
        DateIdx = DateIdx + 1
        
        DoEvents
                
        
    Wend
    ReDim Preserve FinalArray(0 To UBound(FinalArray) - 1)
    Close HistoryFile


    WriteFinalFile App.Path & "\HistoryNew\" & WKN & ".txt"
    
    Exit Sub
ReadHistoryFileErr:
    MsgBox HistoryFileName & vbCr & Err.Description, , "xxxxx"
    
End Sub
'''Private Sub C_ServiceHistory_Click()
'''    Dim Fullpath As String
'''    Dim HistoryFileName As String
'''    Dim HistoryFile As Integer
'''    Dim Zeile As String
'''    Dim HistoryEntities() As String
'''    Dim HistoryIdx As Long
'''    Dim TageszahlInDatum As String
'''    Dim HistoryLine As HistoryItem
'''    Dim DateIdx As Long
'''    Dim Today As String
'''
'''    ReDim FinalArray(0 To 0)
'''
'''
'''    HistoryFileName = App.Path & "\History\" & Text3 & ".txt"
'''    Today = TodayFunction
'''
'''
'''' On Error GoTo ReadHistoryFileErr
'''    HistoryFile = FreeFile
'''    Open HistoryFileName For Input As HistoryFile
'''
'''    ' Read Headline
'''    HistoryIdx = 0
'''    Line Input #HistoryFile, Zeile
''''    List1.AddItem Zeile
'''    SepariereString Zeile, HistoryEntities, ";"
'''        HistoryLine.Datum = HistoryEntities(0)
'''        HistoryLine.Erster = HistoryEntities(1)
'''        HistoryLine.Hoch = HistoryEntities(2)
'''        HistoryLine.Tief = HistoryEntities(3)
'''        HistoryLine.Schlusskurs = HistoryEntities(4)
'''        HistoryLine.Stuecke = HistoryEntities(5)
'''        HistoryLine.Volumen = HistoryEntities(6)
''''    List2.AddItem Zeile
'''    FinalArray(DateIdx) = HistoryLine
'''
'''    ReDim Preserve FinalArray(0 To UBound(FinalArray) + 1)
'''    HistoryIdx = 1
'''    DateIdx = 1
'''
'''    While Not EOF(HistoryFile) And (TageszahlInDatum <> Today)
'''
'''        HistoryLine.Datum = ""
'''        HistoryLine.Erster = ""
'''        HistoryLine.Hoch = ""
'''        HistoryLine.Tief = ""
'''        HistoryLine.Schlusskurs = ""
'''        HistoryLine.Stuecke = ""
'''        HistoryLine.Volumen = ""
'''
'''
'''        Line Input #HistoryFile, Zeile
'''
'''        TageszahlInDatum = FormatDate(DateSerial(2000, 1, 1) + DateIdx - 1)
'''
'''        SepariereString Zeile, HistoryEntities, ";"
'''            HistoryLine.Datum = FormatDate(HistoryEntities(0))
'''            HistoryLine.Erster = HistoryEntities(1)
'''            HistoryLine.Hoch = HistoryEntities(2)
'''            HistoryLine.Tief = HistoryEntities(3)
'''            HistoryLine.Schlusskurs = HistoryEntities(4)
'''            If UBound(HistoryEntities) > 5 Then
'''                HistoryLine.Stuecke = HistoryEntities(5)
'''                HistoryLine.Volumen = HistoryEntities(6)
'''                DoEvents
'''            End If
'''
'''        While (TageszahlInDatum <> HistoryLine.Datum) And (TageszahlInDatum <> Today)
'''            List2.AddItem TageszahlInDatum & "   Inserted "
'''            ReDim Preserve FinalArray(0 To UBound(FinalArray) + 1)
'''            FinalArray(DateIdx).Datum = TageszahlInDatum
'''            FinalArray(DateIdx).Bemerkung = "Inserted"
'''
'''            If (Weekday(FinalArray(DateIdx).Datum) = 7) Or (Weekday(FinalArray(DateIdx).Datum) = 1) Then
'''                FinalArray(DateIdx).Bemerkung = "Weekend"
'''            End If
'''
'''            DateIdx = DateIdx + 1
'''            TageszahlInDatum = FormatDate(DateSerial(2000, 1, 1) + DateIdx - 1)
'''            DoEvents
'''        Wend
'''
'''
''''        List1.AddItem Zeile
'''
'''        List2.AddItem Zeile
'''        FinalArray(DateIdx) = HistoryLine
'''
'''        If (Weekday(FinalArray(DateIdx).Datum) = 7) Or (Weekday(FinalArray(DateIdx).Datum) = 1) Then
'''            FinalArray(DateIdx).Bemerkung = "Weekend"
'''        End If
'''
''''        Text2.Text = Zeile
'''
'''        HistoryIdx = HistoryIdx + 1
'''        DateIdx = DateIdx + 1
'''
'''        DoEvents
'''
'''        ReDim Preserve FinalArray(0 To UBound(FinalArray) + 1)
'''    Wend
'''    ReDim Preserve FinalArray(0 To UBound(FinalArray) - 1)
'''    Close HistoryFile
'''
'''
'''    WriteFinalFile App.Path & "\final.txt"
'''
'''    Exit Sub
'''ReadHistoryFileErr:
'''    MsgBox HistoryFileName & vbCr & Err.Description, , "xxxxx"
'''
'''End Sub


Public Sub WriteFinalFile(FinalFilename As String)
'    Dim FinalFilename As String
    Dim FinalFile As Integer
    Dim idx As Long
    Dim Zeile As String
    
'    On Error GoTo OpenError
    
'    FinalFilename = App.Path & "\Final.txt"
    FinalFile = FreeFile
    Open FinalFilename For Output As FinalFile
    
    For idx = 0 To UBound(FinalArray)
    
        If idx > 5230 Then
            idx = idx
        End If
        
        Zeile = FinalArray(idx).Datum _
                & ";" & FinalArray(idx).Erster _
                & ";" & FinalArray(idx).Hoch _
                & ";" & FinalArray(idx).Tief _
                & ";" & FinalArray(idx).Schlusskurs _
                & ";" & FinalArray(idx).Stuecke _
                & ";" & FinalArray(idx).Volumen _
                & ";" & FinalArray(idx).Bemerkung
         Print #FinalFile, Zeile
    Next idx
           
    Close FinalFile
    
    Exit Sub

OpenError:
    MsgBox FinalFilename, , "Write error"

End Sub



Private Sub Command1_Click()
    Dim a1 As Double
    Dim a2 As Double
    Dim a3 As Double
    
    Init
    GaussPivot A, X, 2
    
    a1 = X(1)
    a2 = X(2)
  
End Sub



















