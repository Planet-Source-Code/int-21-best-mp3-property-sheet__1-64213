VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register Helper Menu"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unregister"
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ExKey As HKEYS_Constants
Private ExPathFiles As String
Private ExPathFolders As String
Private ExPathShortCuts As String


Private Sub Command1_Click()

  Dim cr As New cRegistry
  Dim sKey() As String, lCount As Long, ZX As Long
  Dim sVal As String
  
  Dim sTemp As String
  
  cr.ClassKey = HKEY_CLASSES_ROOT
  cr.SectionKey = "CLSID"
  
  cr.EnumerateSections sKey(), lCount
  Shell "regsvr32 " & App.Path & "\HMp3Info.dll", vbHide
  For ZX = 1 To lCount

    sTemp = "CLSID\" & sKey(ZX)
    sVal = GetString(ExKey, sTemp, "")
    
    If sVal = "Michael's Helper Menu Handler" Then
    
      Debug.Print sKey(ZX)
      SaveString ExKey, ExPathFiles, "", sKey(ZX)
      SaveString ExKey, ExPathFolders, "", sKey(ZX)
      SaveString ExKey, ExPathShortCuts, "", sKey(ZX)
      MsgBox "Registered."
      Unload Me
      Exit Sub
      
    End If
    
  Next
  MsgBox "ID not found - register it."
  Text1 = "ID not found - register it."
  
  Set cr = Nothing
  
End Sub

Private Sub Command2_Click()
  
  DeleteKey ExKey, ExPathFiles
  DeleteKey ExKey, ExPathFolders
  DeleteKey ExKey, ExPathShortCuts
  Shell "regsvr32 /u " & App.Path & "\HMp3Info.dll"
  Unload Me
End Sub

Private Sub Form_Load()
  Text1 = ""
  
  ExKey = HKEY_CLS_ROOT
  'we're adding to 3 different places
  'shortcuts, folders, and all files
  ExPathFiles = "*\shellex\ContextMenuHandlers\Michael's Helper"
  ExPathFolders = "Folder\shellex\ContextMenuHandlers\Michael's Helper"
  ExPathShortCuts = "lnkfile\shellex\ContextMenuHandlers\Michael's Helper"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '
End Sub
