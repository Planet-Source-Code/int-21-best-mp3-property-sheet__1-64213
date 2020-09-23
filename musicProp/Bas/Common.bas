Attribute VB_Name = "Common"
Option Explicit

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal PIDL As Long, ByVal pszPath As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Type SHELLITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHELLITEMID
End Type
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, PIDL As ITEMIDLIST) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'window pos
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_BOTTOM = 1

'GetSystemMetrics
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

Private Const gintMAX_SIZE& = 255

'uptime
Private Declare Function GetTickCount Lib "kernel32" () As Long



Public Function FileExists2(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer

    On Error Resume Next

    ' If the string is quoted, remove the quotes.
    strPathName = strUnQuoteString(strPathName)
    
    'Remove any trailing directory separator character
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists2 = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function

Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
'
    strQuotedString = Trim(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = """" And Right$(strQuotedString, 1) = """" Then
        '
        ' It's quoted.  Get rid of the quotes.
        '
        strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
    End If
    strUnQuoteString = strQuotedString
End Function

Public Function FormatPath(sPath As String) As String
  FormatPath = IIf(Right$(sPath, 1) <> "\", sPath & "\", sPath)
End Function

Public Sub StayOnTop(Frm As Form)

  SetWindowPos Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags

End Sub

Public Sub CenterForm(Frm As Form)
  Dim LX As Long, lY As Long
  LX = GetSystemMetrics(SM_CXSCREEN)
  lY = GetSystemMetrics(SM_CYSCREEN)
  
  If Frm.WindowState <> 0 Then Exit Sub
  
  With Frm
    .Left = ((LX * Screen.TwipsPerPixelX) - Frm.Width) / 2
    .Top = ((lY * Screen.TwipsPerPixelY) - Frm.Height) / 2
  End With
End Sub

Public Function Round(ByVal Expression As Variant, Optional ByVal NumDigitsAfterDecimal As Long) As Variant

  Dim dFactor As Double

  dFactor = CDbl("1" & String$(NumDigitsAfterDecimal, "0"))
  Round = Int(Expression * dFactor + 0.5) / dFactor

End Function

Public Function ComputerUptimeMinutes() As Double

  Dim dwSeconds As Double, dwMinutes As Double, Seconds As Double
  
  Seconds = Round(GetTickCount / 1000, 0) 'GetTickCount tells us how many milleseconds since the computer was started, so divide this by 1000 To get seconds.
  dwMinutes = (Seconds - (Seconds Mod 60)) / 60 '60 is the number of seconds in a minute
  
  ComputerUptimeMinutes = dwMinutes

End Function

Public Function FileExtension(ByVal vInput As String) As String
  Dim L1 As String, L2 As Long
  
  L1 = ParseForFileName(vInput)
  L2 = InStrRev(L1, ".", -1, vbTextCompare)
  
  If L2 = 0 Then
    FileExtension = ""
  Else
    FileExtension = Right(vInput, Len(L1) - L2)
  End If
  
End Function

Public Function InStrRev(ByVal StringCheck As String, ByVal StringMatch As String, Optional ByVal Start As Long = -1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long

  Dim lPosition As Long
  Dim strCheck As String
  Dim lCheckLength As Long

  If Start = 0 Then
    Err.Raise 5
  Else
    lCheckLength = Len(StringCheck)
    If Start < 1 Then Start = (lCheckLength + 1) - Len(StringMatch)
    If Start <= lCheckLength Then
      Do While (Start > 1) And (lPosition = 0)
        lPosition = InStr(Start, StringCheck, StringMatch, Compare)
        If lPosition > Start Then lPosition = 0
        Start = Start - 1
      Loop
      InStrRev = lPosition
    End If
  End If

End Function

Public Function ParseForFileName(ByVal vInput As String) As String

  Dim L1
  
  L1 = InStrRev(vInput, "\", -1, vbTextCompare)
  
  If L1 = 0 Then
    ParseForFileName = vInput
    Exit Function
  End If
  
  ParseForFileName = Right(vInput, Len(vInput) - L1)
  
End Function

Public Function StripNulls(ByVal sString As String) As String
  Dim ZX As Long, sTemp As String, sTest As String
  For ZX = 1 To Len(sString)
    sTest = Mid(sString, ZX, 1)
    If sTest <> Chr(0) Then
      sTemp = sTemp & sTest
    End If
  Next
  StripNulls = sTemp
End Function

Public Function ParseForDir2(ByVal vInput As String) As String

  'if it is a dir leave it be
  If DirExists2(vInput) Then
    ParseForDir2 = vInput
    Exit Function
  End If
  
  Dim L1&
  
  L1 = InStrRev(vInput, "\", -1, vbTextCompare)
  
  If L1 = 0 Then
    ParseForDir2 = vInput
    Exit Function
  End If
  
  ParseForDir2 = Left(vInput, L1)
  
End Function

Public Function MakeLongFilename(ByVal sFileName As String) As String

  On Error GoTo ErrHand
  
  MakeLongFilename = sFileName
  
  If (FileExists2(sFileName)) And _
     (InStr(1, sFileName, "~", vbBinaryCompare) <> 0) Then
     
      Dim sTemp As String
      sTemp = GetLongFileName3(sFileName)
      MakeLongFilename = sTemp
      
  End If
  
  Exit Function
  
ErrHand:

End Function

Public Function GetLongFileName3(ByVal short_name As String) As String
'http://www.vb-helper.com/howto_long_short_file_names.html
Dim pos As Integer
Dim Result As String
Dim long_name As String

    ' Start after the drive letter if any.
    If Mid$(short_name, 2, 1) = ":" Then
        Result = Left$(short_name, 2)
        pos = 3
    Else
        Result = ""
        pos = 1
    End If

    ' Consider each section in the file name.
    Do While pos > 0
        ' Find the next \.
        pos = InStr(pos + 1, short_name, "\")

        ' Get the next piece of the path.
        If pos = 0 Then
            long_name = Dir$(short_name, vbNormal + _
                vbHidden + vbSystem + vbDirectory)
        Else
            long_name = Dir$(Left$(short_name, pos - 1), _
                vbNormal + vbHidden + vbSystem + _
                vbDirectory)
        End If
        Result = Result & "\" & long_name
    Loop

    GetLongFileName3 = Result
End Function

Public Function FormatFilesize(nValue As Variant, Optional bUseFullText As Boolean = False) As String
On Local Error Resume Next
    If nValue <= 1024 Then
        '// Upto 1K
        If Not bUseFullText Then
          FormatFilesize = Format$(nValue, "#,##0 B")
        Else
          FormatFilesize = Format$(nValue, "#,##0 Bytes")
        End If
    ElseIf nValue > 1024 And nValue < 1048576 Then
        '// From 1K+1 to 1MB-1
        FormatFilesize = Format$(nValue / 1024, "###,###,##0.00 KB")
    ElseIf nValue >= 1048576 And nValue < 1073741824 Then
        '// From 1MB +1 to 1GB -1
        FormatFilesize = Format$(nValue / 1048576, "###,###,##0.00 MB")
    Else
        '// Greater than 1GB
        FormatFilesize = Format$(nValue / 1073741824, "###,###,##0.00 GB")
    End If
End Function

Public Function WinSysDir() As String
    Dim strBuf As String
    strBuf = Space$(gintMAX_SIZE)
    If GetSystemDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator(strBuf)
        'AddDirSep strBuf
        WinSysDir = strBuf
    Else
        WinSysDir = vbNullString
    End If
End Function

Public Function WinDir() As String
  Dim Buffer As String * 512, Length As Integer
    Length = GetWindowsDirectory(Buffer, Len(Buffer))
    WinDir = Left$(Buffer, Length)
End Function

Public Function WinDirROOT() As String
  Dim Buffer As String * 512, Length As Integer, WinDirgy
    Length = GetWindowsDirectory(Buffer, Len(Buffer))
    WinDirgy = Left$(Buffer, Length)
    WinDirROOT = Left(WinDirgy, 3)
End Function

Public Function WinTemp() As String
  Dim Buffer As String * 512, Length As Integer
    Length = GetTempPath(Len(Buffer), Buffer)
    WinTemp = Left$(Buffer, Length)
End Function

Public Function WinDesktop() As String
On Error GoTo ErrHand
  Dim strPath As String
  Dim IDL As ITEMIDLIST
  If SHGetSpecialFolderLocation(0&, 16&, IDL) = 0& Then
  strPath = Space(255)
    If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath) Then
      WinDesktop = Left$(strPath, InStr(strPath, vbNullChar) - 1&)
    End If
  End If
  Exit Function
ErrHand:
  Oops "Public Function WinDesktop"
End Function

Public Sub Oops(Optional Location As String)
  MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "LOCATION:" & vbCrLf & Location, vbExclamation, "Oops!"
End Sub

Public Function DirExists2(DirPath As String) As Boolean
  On Error GoTo ErrHand
  
  If DirPath = "" Then Exit Function
  
  If (GetAttr(DirPath) And vbDirectory) = vbDirectory Then
    DirExists2 = True
  End If
  
  Exit Function
  
ErrHand:

  DirExists2 = False
  
End Function

Public Function StripTerminator(ByVal strString As String) As String
  Dim intZeroPos As Integer
  intZeroPos = InStr(strString, Chr$(0))
  If intZeroPos > 0 Then
      StripTerminator = Left$(strString, intZeroPos - 1)
  Else
      StripTerminator = strString
  End If
End Function

Public Function GetShortFileName(ByVal FileName As String) As String
  Dim rc As Long
  Dim ShortPath As String
  'get the short filename
  ShortPath = String$(Len(FileName) + 1, 0)
  rc = GetShortPathName(FileName, ShortPath, Len(FileName) + 1)
  GetShortFileName = Left$(ShortPath, rc)
End Function

Public Sub DebugA(sText As Variant, Optional bAppend As Boolean = True)

  If bAppend Then
    Dim sGet As String
    sGet = Clipboard.GetText
    
    Clipboard.Clear
    Clipboard.SetText sGet & vbCrLf & sText '& vbCrLf
  Else
    Clipboard.Clear
    Clipboard.SetText sText
  End If
    
End Sub

Public Function IsFolder(ByVal sFile As String) As Boolean
  
  If (GetAttr(sFile) And vbDirectory) = vbDirectory Then
    IsFolder = True
  End If
  
End Function

Public Function ContainsFolders(ByRef sArr() As String) As Boolean

  Dim ZX As Long
  
  For ZX = LBound(sArr) To UBound(sArr)
    If IsFolder(sArr(ZX)) Then
      ContainsFolders = True
      Exit Function
    End If
  Next

End Function

Public Function LOWORD(ByVal lVal As Long) As Integer
    LOWORD = lVal And &HFFFF&
End Function
 
Public Function HIWORD(ByVal lVal As Long) As Integer
    HIWORD = 0
    If lVal Then
        HIWORD = lVal \ &H10000 And &HFFFF&
    End If
End Function

Public Function ParseForDirMinusOne(ByVal vInput As String) As String

  Dim L1 As Long
  
  If Len(vInput) = 3 And InStr(1, vInput, ":\") = 2 Then
    ParseForDirMinusOne = vInput
    Exit Function
  End If
  
  If Right(vInput, 1) = "\" Then vInput = Mid(vInput, 1, Len(vInput) - 1)
  L1 = InStrRev(vInput, "\", -1, vbTextCompare)
  
  If L1 = 0 Then
    ParseForDirMinusOne = vInput
    Exit Function
  End If
  
  If L1 = 3 And Right(Left(vInput, 3), 2) = ":\" Then
    ParseForDirMinusOne = Mid(vInput, 1, 3)
    Exit Function
  End If
    
  ParseForDirMinusOne = Mid(vInput, 1, (L1 - 1))
    
End Function
