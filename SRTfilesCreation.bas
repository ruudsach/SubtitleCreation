Dim textData As Variant
Sub SelectFile()

    fileNo = FreeFile 'Get first free file number
    Set spath = Application.FileDialog(msoFileDialogFilePicker)
    spath.Show

    Open spath.SelectedItems(1) For Input As #fileNo
    textData = Input$(LOF(fileNo), fileNo)
    Close #fileNo
    
End Sub


Sub PlayWithHTMLObjects()

Dim html As HTMLDocument, Element As HTMLDivElement

Set Coursefolder = Application.FileDialog(msoFileDialogFolderPicker)
Coursefolder.Show
cfPath = Coursefolder.SelectedItems(1)

Set selFolder = New FileSystemObject
Set selFolder = selFolder.GetFolder(cfPath)

Dim RegEx: Set RegEx = New RegExp
RegEx.Pattern = "[a-zA-Z_]+" '"(a-zA-Z)+"
RegEx.Global = True

Dim mFldrDict As Dictionary
Set mFldrDict = New Dictionary
mFldrDict.CompareMode = TextCompare

Dim cFlDict As Dictionary
Set cFlDict = New Dictionary
cFlDict.CompareMode = TextCompare

For Each mFolder In selFolder.SubFolders

  Set Words = RegEx.Execute(mFolder.Name)
  If Words.Count <> 0 Then
      c_mtch = vbNullString
    For Each mtch In Words
      c_mtch = c_mtch & mtch.Value
    Next mtch
    c_mtch = UCase(c_mtch)
    mFldrDict.Add c_mtch, mFolder.Name
  End If
  
  Set Words = Nothing
  
  For Each cFile In mFolder.Files
  
  If InStr(1, cFile.Type, "mp4", vbTextCompare) Then
    fnameOnly = Mid(cFile.Name, 1, InStrRev(cFile.Name, ".") - 1)
  Set Words = RegEx.Execute(fnameOnly)
  
  If Words.Count <> 0 Then
      f_mtch = vbNullString
    For Each mtch In Words
      f_mtch = f_mtch & mtch.Value
    Next mtch
    f_mtch = UCase(f_mtch)
    cFlDict.Add c_mtch & f_mtch, fnameOnly
  
  End If
    Set Words = Nothing
  End If
  Next cFile
  
Next mFolder

Set html = New HTMLDocument

SelectFile

html.body.innerHTML = textData

Set Element = html.getElementsByClassName("course-transcript").Item(0)
Set mNameList = Element.getElementsByClassName("course-transcript__module")
Debug.Print "Total of 8 Modules:  " & mNameList.Length

For mCount = 0 To mNameList.Length - 1
  
  Set mName = mNameList.Item(mCount)
  Debug.Print "MODULE " & mCount + 1 & ":  " & mName.ParentNode.getElementsByTagName("H2").Item(mCount).innerText
  Set cNameList = mName.getElementsByTagName("H3")
  Set pList = mName.getElementsByTagName("P")
  Debug.Print "Total of " & cNameList.Length & " clips in Module " & mCount + 1

  Set Words = RegEx.Execute(mName.ParentNode.getElementsByTagName("H2").Item(mCount).innerText)
  If Words.Count <> 0 Then
      c_mtch = vbNullString
    For Each mtch In Words
      c_mtch = c_mtch & mtch.Value
    Next mtch
    c_mtch = UCase(c_mtch)
  End If
    Set Words = Nothing
    If mFldrDict.Exists(c_mtch) Then ChDir selFolder & "\" & mFldrDict.Item(c_mtch)
    Debug.Print CurDir$()
  
  For cCount = 0 To (cNameList.Length - 1)
    
    Set cName = cNameList.Item(cCount)
    Debug.Print "Clip " & cCount + 1 & ": " & cName.innerText
    Set ccList = pList.Item(cCount).Children
    
      For lCount = 0 To (ccList.Length - 1)
        
        tm1 = 0
        tm2 = 0
        tLine = ccList.Item(lCount).innerText
        innerStr = ccList.Item(lCount).innerHTML
        
        a = InStr(1, innerStr, "start=", vbTextCompare) + 6
        b = InStr(a, innerStr, Chr(34), vbTextCompare)
        tm1 = Mid(innerStr, a, b - a)
        StartTime = Format(TimeSerial(0, 0, Int(tm1)), "hh:mm:ss") & IIf(InStr(1, tm1, ".", vbTextCompare) > 0, "," & Mid(CStr(tm1), InStr(1, tm1, ".", vbTextCompare) + 1, 3), ",000")
        If tm1 = 0 Then StartTime = "00:00:00,599"
        
        If lCount <> ccList.Length - 1 Then
        
          nInnerStr = ccList.Item(lCount + 1).innerHTML
          a = InStr(1, nInnerStr, "start=", vbTextCompare) + 6
          b = InStr(a, nInnerStr, Chr(34), vbTextCompare)
          tm2 = Mid(nInnerStr, a, b - a) - 0.1
          EndTime = Format(TimeSerial(0, 0, Int(tm2)), "hh:mm:ss") & "," & Mid(CStr(tm2), InStr(1, tm2, ".", vbTextCompare) + 1, 3)
        
        ElseIf lCount = ccList.Length - 1 Then
          
          EndTime = Format(TimeSerial(0, 0, tm1 + 3), "hh:mm:ss") & IIf(InStr(1, tm1, ".", vbTextCompare) > 0, "," & Mid(CStr(tm1), InStr(1, tm1, ".", vbTextCompare) + 1, 3), ",000")
        
        End If
        
        ConcatenatedString = ConcatenatedString & CInt(lCount + 1) & vbCrLf & StartTime & " --> " & EndTime & vbCrLf & tLine & vbCrLf & vbCrLf
        
      Next lCount
      
      Set Words = RegEx.Execute(cName.innerText)
      
      If Words.Count <> 0 Then
          f_mtch = vbNullString
        For Each mtch In Words
          f_mtch = f_mtch & mtch.Value
        Next mtch
        f_mtch = UCase(f_mtch)
      End If
      Set Words = Nothing
      
      If cFlDict.Exists(c_mtch & f_mtch) Then
      
        srtName = CurDir$ & "\" & cFlDict.Item(c_mtch & f_mtch) & ".srt"
        srtFile = FreeFile
        Open srtName For Output As #srtFile
        Print #srtFile, ConcatenatedString
        Close #srtFile
      
      End If
      
      ConcatenatedString = vbNullString
    
  Next cCount
  
Next mCount

End Sub

