Attribute VB_Name = "cRegistry_Addons"
Public Function Wordpad() As String
  
  Dim cReg As New cRegistry
  Dim sKey() As String, lCount As Long, ZX As Long
  Dim sVal As String
  
  cReg.ClassKey = HKEY_CLASSES_ROOT
  cReg.SectionKey = "CLSID"
  
  cReg.EnumerateSections sKey(), lCount
  
  For ZX = 1 To lCount
    sVal = GetString(&H80000000, "CLSID\" & sKey(ZX), "")
    If sVal = "WordPad Document" Then
      sVal = GetString(&H80000000, "CLSID\" & sKey(ZX) & "\LocalServer32", "")
      Wordpad = sVal
      Set cReg = Nothing
      Exit Function
    End If
  Next
  
  Set cReg = Nothing
  
End Function

Public Function CLSIDFromPath(ByVal sFileName As String) As String
  
  Dim cReg As New cRegistry
  Dim sKeySection() As String, lCount As Long, ZX As Long
  Dim sVal As String
  
  Dim lFindKey As HKEYS_Constants
  lFindKey = HKEY_CLS_ROOT
  
  Dim sTemp As String
  
  cReg.ClassKey = HKEY_CLASSES_ROOT
  cReg.SectionKey = "CLSID"
  
  cReg.EnumerateSections sKeySection(), lCount
  
  For ZX = 1 To lCount

    sTemp = "CLSID\" & sKeySection(ZX) & "\InprocServer32"
    sVal = GetString(lFindKey, sTemp, "")
    
    'Debug.Print GetString(HKEY_CLS_ROOT, "CLSID\" & sKeySection(ZX), "")
    
    'i have seen short paths stored in the registry
    If LCase(MakeLongFilename(sVal)) = LCase(sFileName) Then
    
      'Debug.Print sKeySection(ZX)
      CLSIDFromPath = sKeySection(ZX)
      
      Set cReg = Nothing
      Exit Function
      
    End If
    
  Next
  
  Set cReg = Nothing
  
End Function
