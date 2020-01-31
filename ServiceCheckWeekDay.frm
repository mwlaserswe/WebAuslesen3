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
   Begin VB.ListBox List2 
      Height          =   5130
      Left            =   9120
      TabIndex        =   8
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   4200
      TabIndex        =   5
      Top             =   240
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "tst"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "4"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "27.1.2020"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7920
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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


Private Sub Command1_Click()
    Dim Datum1 As String
    Dim Tag1 As Integer
'    Datum1 = "24.12.2002"     ' Datum zuweisen
'    Tag1 = Weekday(Datum1)    ' ergibt 3
'    Text2 = Weekday(Datum1)
'    MsgBox "Die Variable Tag1" & vbCrLf & "enthält " & Weekday(Datum1) & " - für Dienstag."


    Datum1 = Text1.Text
    Tag1 = Weekday(Datum1)     ' ergibt 3
    Text2 = Tag1



End Sub

Private Sub Command2_Click()
    Dim Fullpath As String
    Dim HistoryFileName As String
    Dim HistoryFile As Integer
    Dim Zeile As String
    Dim HistoryEntities() As String
    Dim HistoryIdx As Long
    Dim TageszahlInDatum As String
    Dim HistoryLine As HistoryItem
    Dim DateIdx As Long
    Dim Today As String
    
    ReDim FinalArray(0 To 0)

    
    HistoryFileName = App.Path & "\History\" & Text3 & ".txt"
    Today = TodayFunction


' On Error GoTo ReadHistoryFileErr
    HistoryFile = FreeFile
    Open HistoryFileName For Input As HistoryFile

    ' Read Headline
    HistoryIdx = 0
    Line Input #HistoryFile, Zeile
'    List1.AddItem Zeile
    SepariereString Zeile, HistoryEntities, ";"
        HistoryLine.Datum = HistoryEntities(0)
        HistoryLine.Erster = HistoryEntities(1)
        HistoryLine.Hoch = HistoryEntities(2)
        HistoryLine.Tief = HistoryEntities(3)
        HistoryLine.Schlusskurs = HistoryEntities(4)
        HistoryLine.Stuecke = HistoryEntities(5)
        HistoryLine.Volumen = HistoryEntities(6)
'    List2.AddItem Zeile
    FinalArray(DateIdx) = HistoryLine
     
    ReDim Preserve FinalArray(0 To UBound(FinalArray) + 1)
    HistoryIdx = 1
    DateIdx = 1
    
    While Not EOF(HistoryFile) And (TageszahlInDatum <> Today)

        HistoryLine.Datum = ""
        HistoryLine.Erster = ""
        HistoryLine.Hoch = ""
        HistoryLine.Tief = ""
        HistoryLine.Schlusskurs = ""
        HistoryLine.Stuecke = ""
        HistoryLine.Volumen = ""


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
        
        While (TageszahlInDatum <> HistoryLine.Datum) And (TageszahlInDatum <> Today)
            List2.AddItem TageszahlInDatum & "   Inserted "
            ReDim Preserve FinalArray(0 To UBound(FinalArray) + 1)
            FinalArray(DateIdx).Datum = TageszahlInDatum
            FinalArray(DateIdx).Bemerkung = "Inserted"
            
            If (Weekday(FinalArray(DateIdx).Datum) = 7) Or (Weekday(FinalArray(DateIdx).Datum) = 1) Then
                FinalArray(DateIdx).Bemerkung = "Weekend"
            End If
            
            DateIdx = DateIdx + 1
            TageszahlInDatum = FormatDate(DateSerial(2000, 1, 1) + DateIdx - 1)
            DoEvents
        Wend
        
        
'        List1.AddItem Zeile
        
        List2.AddItem Zeile
        FinalArray(DateIdx) = HistoryLine
        
        If (Weekday(FinalArray(DateIdx).Datum) = 7) Or (Weekday(FinalArray(DateIdx).Datum) = 1) Then
            FinalArray(DateIdx).Bemerkung = "Weekend"
        End If
        
'        Text2.Text = Zeile
        
        HistoryIdx = HistoryIdx + 1
        DateIdx = DateIdx + 1
        
        DoEvents
                
        ReDim Preserve FinalArray(0 To UBound(FinalArray) + 1)
    Wend
    ReDim Preserve FinalArray(0 To UBound(FinalArray) - 1)
    Close HistoryFile




    WriteFinalFile App.Path & "\final.txt"




    Exit Sub
ReadHistoryFileErr:
    MsgBox HistoryFileName & vbCr & Err.Description, , "xxxxx"
    
End Sub

Public Function TodayFunction() As String
    Dim DateTimeString As String
    Dim DateString As String
    Dim SepArray() As String
    
    DateTimeString = Now
    SepariereString DateTimeString, SepArray, " "
    DateString = SepArray(0)
    SepariereString DateString, SepArray, "."
    TodayFunction = SepArray(2) & "-" & SepArray(1) & "-" & SepArray(0)
End Function

Private Sub Command3_Click()
    Dim HistoryEntities() As String
    Dim HistoryIdx As Long
    Dim Zeile As String
    Dim HistoryLine As HistoryItem
    
    Zeile = "2020-01-16;316,40;317,45;313,40;313,90;348.285;109.660.544"
    SepariereString Zeile, HistoryEntities, ";"
        HistoryLine.Datum = HistoryEntities(0)
        HistoryLine.Erster = HistoryEntities(1)
        HistoryLine.Hoch = HistoryEntities(2)
        HistoryLine.Tief = HistoryEntities(3)
        HistoryLine.Schlusskurs = HistoryEntities(4)
        HistoryLine.Stuecke = HistoryEntities(5)

    Zeile = "2020-01-16;;;;313,90;;"
    SepariereString Zeile, HistoryEntities, ";"
        HistoryLine.Datum = HistoryEntities(0)
        HistoryLine.Erster = HistoryEntities(1)
        HistoryLine.Hoch = HistoryEntities(2)
        HistoryLine.Tief = HistoryEntities(3)
        HistoryLine.Schlusskurs = HistoryEntities(4)
        HistoryLine.Stuecke = HistoryEntities(5)




End Sub

''Private Sub Command3_Click()
''Dim idx As Long
''Dim day As Long
''Dim month As Long
''Dim year As Long
''
''idx = 1
''
''day = 1
''month = 1
''year = 2000
''
''While idx < 7
''  dt = day &
''
''
''
''Wend
''
''
''
''
''
''
''End Sub

Private Sub Command4_Click()
'Public Function TageszahlInDatum(intTag As Integer) As Date
   intTag = Text2.Text
'   TageszahlInDatum = DateSerial(year(Now), 1, 1) + intTag - 1
   TageszahlInDatum = DateSerial(2000, 1, 1) + intTag - 1
   Text1.Text = TageszahlInDatum
End Sub

Public Function FormatDate(DString As String) As String
    Dim DateEntities() As String

    If InStr(DString, "-") Then
        FormatDate = DString
    ElseIf InStr(DString, ".") Then
        SepariereString DString, DateEntities, "."
        FormatDate = DateEntities(2) & "-" & DateEntities(1) & "-" & DateEntities(0)
    Else
        FormatDate = "0000-00-00"
    End If
End Function





Public Sub WriteFinalFile(FinalFilename As String)
'    Dim FinalFilename As String
    Dim FinalFile As Integer
    Dim idx As Long
    Dim Zeile As String
    
    On Error GoTo OpenError
    
'    FinalFilename = App.Path & "\Final.txt"
    FinalFile = FreeFile
    Open FinalFilename For Output As FinalFile
    
    For idx = 0 To UBound(FinalArray)
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



