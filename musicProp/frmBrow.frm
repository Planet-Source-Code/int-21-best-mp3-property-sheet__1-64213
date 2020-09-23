VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browsing..."
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   4875
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   5925
      ExtentX         =   10451
      ExtentY         =   8599
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmBrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type tAlbums
    sArtist As String
    sAlbum As String
    sScript As String
End Type
Dim xAlbums() As tAlbums

Private Sub Form_Load()
'
End Sub

Function BuildAlbums(memo As Long, lCount As Long)
Dim I&
    ReDim xAlbums(1 To lCount)
    'read albums data from memory
    CopyMemory ByVal VarPtrArray(xAlbums), memo, 4
    'create a html file to build all the albums
    Open App.Path & "\xalbums.html" For Output As #50
        Print #50, "<html><style>a:link{color:blue}a:visited{color:blue}</style><body bgColor=silver><h1 style='align=center'>Albums founds</h1><br>"
        For I = 1 To lCount
            Print #50, "<a href='javascript:" & I & "' id='album_" & I; "'>Set cover</a>"
            Print #50, "<IMG style='WIDTH: 120px; HEIGHT: 120px' src=""" & xAlbums(I).sScript & """ align=left border=1 >"
            Print #50, "<TABLE cellSpacing=0 cellPadding=0 border=0 style='FONT-SIZE: 12px; WIDTH: 196px; FONT-FAMILY: tahoma; HEIGHT: 57px' ><TR vAlign=top>"
            Print #50, "<TD style='FONT-WEIGHT: bold'>Artist</TD><TD style='COLOR: dimgray'>" & xAlbums(I).sArtist & "</TD></TR>"
            Print #50, "<TR bgColor=silver><TD><IMG style='WIDTH: 1px; HEIGHT: 1px'></TD><TD></TD></TR>"
            Print #50, "<TR vAlign=top ><TD style='FONT-WEIGHT: bold'>Albums</TD><TD style='COLOR: dimgray'>" & xAlbums(I).sAlbum & "</TD></TR>"
            Print #50, "</TABLE><br><br><br><br><hr>"
        Next
        Print #50, "</body></html><script>document.body.style.scrollbarBaseColor='#000080';" & _
        "document.body.style.scrollbarArrowColor='#000080';" & _
        "document.body.style.scrollbarDarkShadowColor='#808080';" & _
         "document.body.style.scrollbarFaceColor='#808080';" & _
         "document.body.style.scrollbarHighlightColor='#FFFFFF';" & _
         "document.body.style.scrollbarShadowColor='#000000';" & _
         "document.body.style.scrollbar3dlightColor='#000000';" & _
        "</script>"
    Close #50
    
    Web.Navigate App.Path & "\xalbums.html"
    Me.Show vbModal, frmpage

End Function


Private Sub Form_Unload(Cancel As Integer)
    CopyMemory ByVal VarPtrArray(xAlbums), 0&, 4
    Erase xAlbums
End Sub

Private Sub Web_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'capture here the click events over the link and get the albums selected
    If InStr(1, URL, "javascript") Then
        Dim xUrl
        xUrl = Split(URL, ":")
        frmpage.lbWait.Enabled = False: frmpage.lbWait = "Downloading..."
        frmpage.Inet.Execute xAlbums(CLng(xUrl(1))).sScript, "GET"
        Unload Me
    End If
End Sub
