Attribute VB_Name = "peketec_de"
Option Explicit

Public Function Read_Peketec(WKN As String) As Integer
    Dim HtmlCode As String

    Dim WebPage As String
    Dim SearchItem As String
    Dim AktuellerKurs As String
    Dim InstrumentId As String
    Dim ISIN As String
    
    
    
    Dim PosStart As Long
    Dim PosEnd As Long
    Dim EndString As String
    
    WebPage = "https://peketec.de/portal/charts/show/key/" & WKN
    HtmlCode = GetHTMLCode(WebPage)
    If HtmlCode = ">>>ERROR<<<" Then
        Read_Peketec = 1
        ReadTodaysSharePrice.L_SharePrice = "Error"
        Exit Function
    End If
  
'    SaveQuelltext HtmlCode, App.Path & "\GoogleText.HTML"
 
    'AktuellerKurs extrahieren
    SearchItem = "<div class=""rtlast"">"
    EndString = "</div>"
    AktuellerKurs = ExtraxtValue(HtmlCode, SearchItem, EndString)
    Dim dmy As Double
    dmy = Zahl(AktuellerKurs)
    
    ReadTodaysSharePrice.L_WebPage = WebPage
     ReadTodaysSharePrice.L_SharePrice = dmy
'    L_WKN = WKN
'    L_ISIN = ISIN
 


     Read_Peketec = 0

End Function

