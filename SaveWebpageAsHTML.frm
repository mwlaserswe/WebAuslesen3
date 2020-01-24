VERSION 5.00
Begin VB.Form SaveWebpageAsHTML 
   Caption         =   "Save Web page as HTML"
   ClientHeight    =   2100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   LinkTopic       =   "Form2"
   ScaleHeight     =   2100
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "SaveWebpageAsHTML.frx":0000
      Top             =   480
      Width           =   6735
   End
   Begin VB.TextBox T_URL 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "https://www.google.de/"
      Top             =   1440
      Width           =   5175
   End
   Begin VB.CommandButton C_SavePageToFile 
      Caption         =   "Save Page To File"
      Height          =   735
      Left            =   5760
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "SaveWebpageAsHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub C_SavePageToFile_Click()
    Dim HtmlCode As String
    HtmlCode = GetHTMLCode(T_URL)
    SaveQuelltext HtmlCode, App.Path & "\HTML-Code.HTML"
End Sub
