VERSION 5.00
Begin VB.Form ChartList 
   Caption         =   "Form2"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   LinkTopic       =   "Form2"
   ScaleHeight     =   7485
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton C_MoveToPartialList 
      Caption         =   ">>>"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ListBox ListPartial 
      Height          =   6885
      Left            =   8640
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
   Begin VB.ListBox ListCompelete 
      Height          =   6885
      Left            =   840
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "ChartList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim CompanyEntry As String
    Dim CompanyIndex As Long

Private Sub C_MoveToPartialList_Click()
    Dim k As Long
    Dim CompanyEntry As String
    
    For k = 0 To ListCompelete.ListCount - 1
        If ListCompelete.Selected(k) Then
            CompanyEntry = ListCompelete.List(k)
            ' ... do something
            ListPartial.AddItem CompanyEntry
            ListCompelete.Selected(k) = False
        End If
    Next
End Sub


Private Sub Form_Load()
    ReadCompanyListFile ListCompelete
    CompanyIndex = -1       ' -1 means no list item selected
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim Zeile As String
    Dim CompanyListEntities() As String
    Dim idx As Long
    
    ReDim CompPartialLstArr(0 To 0)
    
    For idx = 0 To ListPartial.ListCount - 1
        Zeile = ListPartial.List(idx)
        SepariereString Zeile, CompanyListEntities, vbTab
        CompPartialLstArr(idx).Name = CompanyListEntities(0)
        CompPartialLstArr(idx).WKN = CompanyListEntities(1)
        CompPartialLstArr(idx).ISIN = CompanyListEntities(2)
          
        ReDim Preserve CompPartialLstArr(0 To UBound(CompPartialLstArr) + 1)
    Next idx
                
    ReDim Preserve CompPartialLstArr(0 To UBound(CompPartialLstArr) - 1)

End Sub










Private Sub ListCompelete_Click()
'    CompanyIndex = ListCompelete.ListIndex
'    CompanyEntry = ListCompelete.List(CompanyIndex)
End Sub

Private Sub Text1_Change()
    Dim k As Long
    Dim CompanyEntry As String

    ListPartial.Clear
    For k = 0 To ListCompelete.ListCount - 1
        CompanyEntry = ListCompelete.List(k)

        If InStr(1, CompanyEntry, Text1, vbTextCompare) <> 0 Then
            ListPartial.AddItem CompanyEntry
        End If
    Next
End Sub
