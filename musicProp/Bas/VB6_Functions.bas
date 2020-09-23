Attribute VB_Name = "VB6_Functions"
Option Explicit

' Author: Lewis Miller (aka Deth)
'   Date: 11/08/03
'Purpose: Complete vb6 for vb5 replacement functions


Public Function Replace(ByVal Expression As String, _
                        ByVal Find As String, _
                        ByVal Replacement As String, _
               Optional ByVal Start As Long = 1, _
               Optional ByVal Count As Long = -1, _
               Optional ByVal Compare As VbCompareMethod) As String

  Dim lExpLength As Long
  Dim lPosition As Long
  Dim lFindLen As Long
  Dim lRepLen As Long
  Dim lCount As Long

    If Find <> Replacement Then
        lExpLength = Len(Expression)
        lFindLen = Len(Find)
        If lExpLength > 0 And lFindLen > 0 Then
            lRepLen = Len(Replacement)
            If Start <= lExpLength Then
                lPosition = InStr(Start, Expression, Find, Compare)
                Do While lPosition > 0
                    Expression = Left$(Expression, lPosition - 1) & Replacement & Mid$(Expression, lPosition + lFindLen)
                    lPosition = InStr(lPosition + lRepLen, Expression, Find, Compare)
                    lCount = lCount + 1
                    If lCount = Count Then
                        Exit Do
                    End If
                Loop
            End If
        End If
    End If
    Replace = Expression

End Function

Public Function Split(ByVal Expression As String, _
             Optional ByVal Delimiter As String, _
             Optional ByVal Limit As Long = -1, _
             Optional ByVal Compare As VbCompareMethod) As Variant

  Dim lPosition  As Long
  Dim lDelimLen  As Long
  Dim strArr()   As String
  Dim lExpLen    As Long
  Dim lCount     As Long

    lExpLen = Len(Expression)
    lDelimLen = Len(Delimiter)
    If (lExpLen > 0) And (lDelimLen > 0) Then
        lPosition = InStr(1, Expression, Delimiter, Compare)
        If lPosition > 0 Then
            Do While lPosition > 0
                lCount = lCount + 1
                If lCount = Limit Then
                    Exit Do
                End If
                lPosition = InStr(lPosition + 1, Expression, Delimiter, Compare)
            Loop
            ReDim strArr(lCount) As String
            lCount = 0
            
            lPosition = InStr(1, Expression, Delimiter, Compare)
            lExpLen = 1
            Do While lPosition > 0
                strArr(lCount) = Left$(Expression, lPosition - lExpLen)
                lExpLen = lPosition + 1
                lPosition = InStr(lPosition + lDelimLen, Expression, Delimiter, Compare)
                lCount = lCount + 1
                If lCount = Limit - 1 Then
                    Exit Do
                End If
            Loop
            strArr(lCount) = Mid$(Expression, lExpLen)
            GoTo Done
        End If
    End If

    ReDim strArr(0) As String
    strArr(0) = Expression

Done:
    Split = strArr

End Function

Public Function Join(SourceArray() As String, _
      Optional ByVal Delimiter As String, _
      Optional ByVal Count As Long = -1) As String

  Dim lTotal As Long
  Dim lUpperBound As Long
  Dim lLowerBound As Long
  Dim lPosition As Long
  Dim lDelimLen As Long
    
    Err.Clear
    On Error Resume Next 'just in case array is not initialized
        lUpperBound = UBound(SourceArray)
        If Err.Number = 0 Then
            lLowerBound = LBound(SourceArray)
            If (Count <> -1) Then
                If (Count <= lUpperBound + 1) And (Count > lLowerBound) Then
                    lUpperBound = Count - 1
                End If
            End If
            lPosition = lLowerBound
            lDelimLen = Len(Delimiter)
            Do
                lTotal = lTotal + Len(SourceArray(lPosition)) + lDelimLen
                lPosition = lPosition + 1
            Loop While lPosition < lUpperBound + 1
            Join = Space$(lTotal - lDelimLen)
            lPosition = 1
            If lLowerBound < lUpperBound Then
                Do While lLowerBound < lUpperBound + 1
                    lTotal = Len(SourceArray(lLowerBound))
                    Mid$(Join, lPosition, lTotal) = SourceArray(lLowerBound)
                    lPosition = lPosition + lTotal
                    Mid$(Join, lPosition, lDelimLen) = Delimiter
                    lPosition = lPosition + lDelimLen
                    lLowerBound = lLowerBound + 1
                Loop
            End If
            Mid$(Join, lPosition, Len(SourceArray(lUpperBound))) = SourceArray(lUpperBound)
        End If

End Function

Function StrReverse(ByVal Expression As String) As String

  Dim lngLength As Long
  Dim X As Long

    lngLength = Len(Expression)
    If lngLength > 0 Then
        StrReverse = Space$(lngLength)
        For X = lngLength To 1 Step -1
            Mid$(StrReverse, X, 1) = Mid$(Expression, (lngLength + 1) - X, 1)
        Next X
    End If

End Function

Public Function Round(ByVal Expression As Variant, _
             Optional ByVal NumDigitsAfterDecimal As Long) As Variant

  Dim dFactor As Double

    dFactor = CDbl("1" & String$(NumDigitsAfterDecimal, "0"))
    Round = Int(Expression * dFactor + 0.5) / dFactor

End Function

Public Function InStrRev(ByVal StringCheck As String, _
                         ByVal StringMatch As String, _
                Optional ByVal Start As Long = -1, _
                Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long

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

Public Function Filter(InputStrings As Variant, _
                 ByVal Value As String, _
        Optional ByVal Include As Boolean, _
        Optional ByVal Compare As VbCompareMethod) As Variant

  Dim strFinal()  As String
  Dim lLowerBound As Long
  Dim lUpperBound As Long
  Dim lInclude    As Long
  Dim lCount      As Long
  Dim X           As Long

    ReDim strFinal(0) As String

    On Error Resume Next 'just in case array is not initialized
        lUpperBound = UBound(InputStrings)
        If Err.Number = 0 Then
            lLowerBound = LBound(InputStrings)
            lInclude = Abs(Include)
            'count dimensions needed
            For X = lLowerBound To lUpperBound
                If (StrComp(InputStrings(X), Value, Compare) Xor lInclude) Then
                    lCount = lCount + 1
                End If
            Next X
            ReDim strFinal(lCount - 1)
            lCount = 0
            'fill new array
            For X = lLowerBound To lUpperBound
                If (StrComp(InputStrings(X), Value, Compare) Xor lInclude) Then
                    strFinal(lCount) = InputStrings(X)
                    lCount = lCount + 1
                End If
            Next X
        End If
        Filter = strFinal

End Function

Public Function FormatNumber(ByVal dwNum As Currency, Optional NumPlacesAfterDec As Long, _
  Optional lNothing As Long, Optional lNothing2 As Long, Optional iUnknown As Boolean) As Currency
  Dim sTemp As String
  Dim sNew As String
  Dim lDec As Long
  sTemp = dwNum
  lDec = InStr(1, sTemp, ".", vbTextCompare)
  If lDec = 0 Then
    FormatNumber = dwNum
  Else
    sNew = Left(sTemp, lDec + NumPlacesAfterDec)
    FormatNumber = Val(sNew)
  End If
End Function

Public Function FormatPercent(ByVal dwNum As Currency, Optional NumPlacesAfterDec As Long) As String
  Dim sTemp As String
  Dim sNew As String
  Dim lDec As Long
  
  If dwNum > 1 Then
    dwNum = 0 'ErrHandling?
  End If
  
  dwNum = dwNum * 100
  
  sTemp = dwNum
  lDec = InStr(1, sTemp, ".", vbTextCompare)
  If lDec = 0 Then
    FormatPercent = dwNum & " %"
  Else
    sNew = Left(sTemp, lDec + NumPlacesAfterDec)
    FormatPercent = Val(sNew) & " %"
  End If
End Function

Public Sub SplitB(Expression$, ResultSplit$(), Optional Delimiter$ = " ")
' By Chris Lucas, cdl1051@earthlink.net, 20011208
'    example
'    Sub FindPaths(sString As String)
'      Dim sPaths() As String, ZX As Long
'      SplitB sString, sPaths()
'      For ZX = LBound(sPaths) To UBound(sPaths)
'        MsgBox sPaths(ZX)
'      Next
'    End Sub
    Dim c&, SLen&, DelLen&, tmp&, Results&()

    SLen = LenB(Expression) \ 2
    DelLen = LenB(Delimiter) \ 2

    ' Bail if we were passed an empty delimiter or an empty expression
    If SLen = 0 Or DelLen = 0 Then
        ReDim Preserve ResultSplit(0 To 0)
        ResultSplit(0) = Expression
        Exit Sub
    End If

    ' Count delimiters and remember their positions
    ReDim Preserve Results(0 To SLen)
    tmp = InStr(Expression, Delimiter)

    Do While tmp
        Results(c) = tmp
        c = c + 1
        tmp = InStr(Results(c - 1) + 1, Expression, Delimiter)
    Loop

    ' Size our return array
    ReDim Preserve ResultSplit(0 To c)

    ' Populate the array
    If c = 0 Then
        ' lazy man's call
        ResultSplit(0) = Expression
    Else
        ' typical call
        ResultSplit(0) = Left$(Expression, Results(0) - 1)
        For c = 0 To c - 2
            Dim p
            p = ResultSplit(c)
            ResultSplit(c + 1) = Mid$(Expression, _
                Results(c) + DelLen, _
                Results(c + 1) - Results(c) - DelLen)
        Next c
        ResultSplit(c + 1) = Right$(Expression, SLen - Results(c) - DelLen + 1)
    End If

End Sub

Public Sub SplitC(Expression$, ResultSplit$(), Optional Delimiter$ = " ")
On Error Resume Next
    Dim c&, SLen&, DelLen&, tmp&, Results&()

    SLen = LenB(Expression) \ 2
    DelLen = LenB(Delimiter) \ 2

    ' Bail if we were passed an empty delimiter or an empty expression
    If SLen = 0 Or DelLen = 0 Then
        ReDim Preserve ResultSplit(0 To 0)
        ResultSplit(0) = Expression
        Exit Sub
    End If

    ' Count delimiters and remember their positions
    ReDim Preserve Results(0 To SLen)
    'tmp = InStr(Expression, Delimiter)
    Dim sTemp As String
    Dim ZX As Long
    
    For ZX = 1 To Len(Expression)
      sTemp = Mid(Expression, ZX, Len(Delimiter))
      If sTemp = Delimiter Then
        Results(c) = ZX
        c = c + 1
      End If
    Next

'    Do While tmp
'        Results(c) = tmp
'        c = c + 1
'        tmp = InStr(Results(c - 1) + 1, Expression, Delimiter)
'    Loop

    ' Size our return array
    ReDim ResultSplit(0)
    ' Populate the array
    If c = 0 Then
        ' lazy man's call
        ResultSplit(0) = Expression
    Else
        ' typical call
        ResultSplit(0) = Left$(Expression, Results(0) - 1)
        For c = 0 To c - 2
            Dim p, r
            p = ResultSplit(c)
            r = Mid$(Expression, _
                Results(c) + DelLen, _
                Results(c + 1) - Results(c) - DelLen)
            If (r <> "") And (r <> ResultSplit(UBound(ResultSplit))) Then
              ReDim Preserve ResultSplit(UBound(ResultSplit) + 1)
              ResultSplit(UBound(ResultSplit)) = r
            End If
        Next c
        'ResultSplit(c + 1) = Right$(Expression, SLen - Results(c) - DelLen + 1)
        
'        For ZX = LBound(ResultSplit) To UBound(ResultSplit)
'          Debug.Print ResultSplit(ZX)
'        Next
'        Debug.Print vbCrLf & UBound(ResultSplit)
    End If

End Sub

