Attribute VB_Name = "Utilil"
Option Explicit

Public Function ExtraxtValue(HtmlString As String, SearchItem As String, EndString As String) As String
    Dim GoogleText As String
    
'    Dim SearchItem As String
'    Dim AktuellerKurs As String
'    Dim InstrumentId As String
    Dim DmyString As String
    
    Dim PosStart As Long
    Dim PosEnd As Long
    
'    SearchItem = "Aktueller Kurs"
    PosStart = InStr(HtmlString, SearchItem)
    DmyString = Mid$(HtmlString, PosStart + Len(SearchItem), 20)
    If PosStart > 0 Then
        PosStart = PosStart + Len(SearchItem) + 1
        PosEnd = InStr(PosStart, HtmlString, EndString)
        If PosStart > 0 Then
            ExtraxtValue = Mid$(HtmlString, PosStart, PosEnd - PosStart)
        End If
    End If

End Function


Public Function GetHTMLCode(strURL) As String
    Dim strReturn                   ' As String
    Dim objHTTP                     ' As MSXML.XMLHTTPRequest
    
    On Error GoTo AccessError
    
    If Len(strURL) = 0 Then Exit Function
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", strURL, False
    objHTTP.Send                    'Get it.
    strReturn = objHTTP.responseText
    Set objHTTP = Nothing
    GetHTMLCode = strReturn
    
    Exit Function

AccessError:
    GetHTMLCode = ">>>ERROR<<<"
    AccessErrorCnt = AccessErrorCnt + 1
'    L_ErrorCnt = AccessErrorCnt
'    Form1.Timer2.Enabled = False
End Function


' Bsp.: Seitenquelltext in Datei speichern
Public Sub SaveQuelltext(DateiString As String, ByVal sFilename As String)
  Dim f As Integer
 
    f = FreeFile
    Open sFilename For Output As #f
    Print #f, DateiString
    Close #f
End Sub



Public Function AddVector(v1 As MousePos, v2 As MousePos) As MousePos
    AddVector.X = v1.X + v2.X
    AddVector.Y = v1.Y + v2.Y
End Function

Public Function SubVector(v1 As MousePos, v2 As MousePos) As MousePos
    SubVector.X = v1.X - v2.X
    SubVector.Y = v1.Y - v2.Y
End Function


Sub ReadCompanyListFile(MyList As Variant)
    Dim CompanyListFilename As String
    Dim CompanyListFile As Integer
    Dim Zeile As String
    Dim CompanyListEntities() As String
    Dim idx As Long
    
    ReDim CompanyListArray(0 To 0)
    MyList.Clear
'    List2.Clear
    
    On Error GoTo ReadCompanyListFileErr
    
    CompanyListFilename = App.Path & "\ISIN-WKN.txt"
    CompanyListFile = FreeFile
    Open CompanyListFilename For Input As CompanyListFile
        
    While Not EOF(CompanyListFile)
        Line Input #CompanyListFile, Zeile
        If Zeile <> "" Then
            MyList.AddItem Zeile
            SepariereString Zeile, CompanyListEntities, vbTab
            idx = UBound(CompanyListArray)
            CompanyListArray(idx).Name = CompanyListEntities(0)
            CompanyListArray(idx).WKN = CompanyListEntities(1)
            CompanyListArray(idx).ISIN = CompanyListEntities(2)
            
'''            'Search doubbles
'''            Dim i As Long
'''            For i = 0 To UBound(CompanyListArray) - 1
'''                If CompanyListArray(i).WKN = CompanyListArray(idx).WKN Then
'''                    List2.AddItem Zeile
'''                End If
'''            Next i
            
            ReDim Preserve CompanyListArray(0 To UBound(CompanyListArray) + 1)
        End If
                
    Wend
    ReDim Preserve CompanyListArray(0 To UBound(CompanyListArray) - 1)
    Close CompanyListFile
      
     Exit Sub
ReadCompanyListFileErr:
    MsgBox CompanyListFilename & vbCr & Err.Description, , "xxxxx"
End Sub








