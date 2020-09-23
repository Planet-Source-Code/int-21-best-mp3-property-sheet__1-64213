VERSION 5.00
Begin VB.Form frmEnumHandlers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enum Handlers"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   Icon            =   "frmEnumHandlers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Enum Context Menu Handlers"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "This outputs to the ide immediate window so compiling this isn't going to be very helpful."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmEnumHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

  Dim cr As New cRegistry
  Dim sKey() As String, lCount As Long, ZX As Long
  Dim sVal As String, lFound As Long
  
  Dim sTemp As String
  
  Debug.Print "Key:       Registered Name:"
  Debug.Print "----------------------------------"
  Debug.Print
  
  cr.ClassKey = HKEY_CLASSES_ROOT
  cr.SectionKey = ""
  
  cr.EnumerateSections sKey(), lCount
  
  For ZX = 1 To lCount

    Dim cr2 As New cRegistry
    Dim iKey() As String, iKeyCount As Long, YX As Long
    
    cr2.ClassKey = HKEY_CLASSES_ROOT
    cr2.SectionKey = sKey(ZX) & "\shellex\ContextMenuHandlers"
    cr2.EnumerateSections iKey, iKeyCount
    
    For YX = 1 To iKeyCount
      lFound = lFound + 1
      Debug.Print sKey(ZX) & " - " & iKey(YX)
    Next YX
    
    Set cr2 = Nothing
    
  Next ZX
  
  Set cr = Nothing
  
  MsgBox "Found " & lFound & " handlers."
  
  
End Sub


