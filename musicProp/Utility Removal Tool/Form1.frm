VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Helper Removal Tool"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   7830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnregisterTypeLibrary 
      Caption         =   "Unregister Type Library"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   7575
   End
   Begin VB.CommandButton cmdRemoveReferences 
      Caption         =   "Remove All References To DLL"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0442
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This will actually remove all references to any dll / ocx"""
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ExKey As HKEYS_Constants

Private Sub cmdRemoveReferences_Click()

  Dim cr As New cRegistry
  Dim sKey() As String, lCount As Long, ZX As Long
  Dim sVal As String, lFound As Long, lAll As Long
  
  Dim sTemp As String
  
  Dim sPath As String
  sPath = Text1
  
  If Not FileExists2(sPath) Then
    MsgBox "Click the browse button and select a valid dll / ocx."
    Exit Sub
  End If

  cr.ClassKey = HKEY_CLASSES_ROOT
  cr.SectionKey = "CLSID"
  
  cr.EnumerateSections sKey(), lCount
  
  For ZX = 1 To lCount

    sTemp = "CLSID\" & sKey(ZX) & "\InprocServer32"
    sVal = GetString(ExKey, sTemp, "")
    If LCase(sVal) = LCase(sPath) Then
    
      Debug.Print "CLSID\" & sKey(ZX) & "\InprocServer32\" & sVal
      DeleteKey &H80000000, "CLSID\" & sKey(ZX)
      lFound = lFound + 1
      
    End If
    
  Next
  
  Set cr = New cRegistry
  
  cr.ClassKey = HKEY_CLASSES_ROOT
  cr.SectionKey = "TypeLib"
  
  cr.EnumerateSections sKey(), lCount
  
  For ZX = 1 To lCount

    Dim cReg1 As New cRegistry
    Dim sKeySection1() As String, lCount1 As Long
    
    cReg1.ClassKey = HKEY_CLASSES_ROOT
    cReg1.SectionKey = "TypeLib\" & sKey(ZX)
    
    cReg1.EnumerateSections sKeySection1(), lCount1
    
    If lCount1 > 0 Then
    
      sTemp = "TypeLib\" & sKey(ZX) & "\" & sKeySection1(1) & "\0\win32"
      sVal = GetString(HKEY_CLS_ROOT, sTemp, "")
      
      If LCase(sVal) = LCase(sPath) Then
      
        Debug.Print sTemp & "\" & sVal
        DeleteKey &H80000000, "TypeLib\" & sKey(ZX) & "\" & sKeySection1(1)
        lFound = lFound + 1
        
      End If
    
    End If
    
  Next
  
  Set cr = Nothing
  
  MsgBox "Deleted " & lFound & " references."
  
End Sub

Private Sub cmdBrowse_Click()

  Dim sTemp As String
  sTemp = CommonDialogShowOpen(App.Path, "Choose File", "DLL OCX TLB Files" & Chr(0) & "*.dll;*.ocx;*.tlb", Me, True)
  If Not CommonDialogShowOpenERROR Then
    Text1 = sTemp
  End If
  
End Sub

Private Sub cmdUnregisterTypeLibrary_Click()

  Dim sPath As String
  sPath = MakeLongFilename(Text1)
  
  If Not FileExists2(sPath) Then
    MsgBox "Click the browse button and select a valid type library."
    Exit Sub
  End If
  
  RegisterTypeLib_Confirm sPath, False, True
  
End Sub

Private Sub Form_Load()
  Text1 = ""
  
  ExKey = HKEY_CLS_ROOT
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '
End Sub

