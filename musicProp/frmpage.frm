VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmpage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmpage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   361
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   2190
      Top             =   3030
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Mp3File|*.mp3"
   End
   Begin VB.CommandButton cmdLoad 
      Cancel          =   -1  'True
      Caption         =   "Load MP3File"
      Height          =   315
      Left            =   1560
      TabIndex        =   18
      Top             =   5010
      Width           =   1335
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   2310
      Top             =   2430
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   15
      Top             =   5010
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   5010
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   90
      TabIndex        =   12
      Top             =   2460
      Width           =   4995
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Left            =   3090
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1800
      End
      Begin VB.Label lbMpegInfo 
         Caption         =   "Mpeg_Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   90
         TabIndex        =   13
         Top             =   180
         Width           =   2670
      End
      Begin VB.Image ImgClock 
         Height          =   225
         Left            =   3840
         Picture         =   "frmpage.frx":058A
         Top             =   420
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label lbWait 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   825
         Left            =   3150
         MouseIcon       =   "frmpage.frx":05F0
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   810
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2325
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4995
      Begin VB.ComboBox cboGenre 
         Height          =   315
         Left            =   2700
         TabIndex        =   16
         Top             =   1380
         Width           =   2025
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1000
         TabIndex        =   5
         Top             =   1860
         Width           =   3705
      End
      Begin VB.TextBox txtYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1000
         TabIndex        =   4
         Top             =   1380
         Width           =   795
      End
      Begin VB.TextBox txtAlbum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1000
         TabIndex        =   3
         Top             =   990
         Width           =   3705
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1000
         TabIndex        =   2
         Top             =   600
         Width           =   3705
      End
      Begin VB.TextBox txtArtist 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1000
         TabIndex        =   1
         Top             =   210
         Width           =   3705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Comment:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   11
         Top             =   1890
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Genre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2100
         TabIndex        =   10
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   1410
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Album:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Artist:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Dim arrGenre()
Dim sCurrentFile$
Private Type ID3v1Data           'This type is standard for ID3v1 tags
  Title       As String * 30    '30 bytes Title
  Artist      As String * 30    '30 bytes Artist
  Album       As String * 30    '30 bytes Album
  Year        As String * 4     '4 bytes Year
  Comments    As String * 28    '28 bytes Comments
  IsTrack     As Byte           '1 byte Istrack / +1 byte comments
  Tracknumber As Byte           '1 byte Tracknumber / +1 byte comments
  Genre       As Byte           '1 byte Genre
End Type

Private Type tAlbums
    sArtist As String
    sAlbum As String
    sScript As String
End Type
Dim aAlbums() As tAlbums
Dim memoArray&

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
    CD.ShowOpen
    If CD.FileName <> "" Then LoadData CD.FileName
End Sub
'Save the changes to the mp3 file
Private Sub cmdUpdate_Click()
If sCurrentFile <> "" Then
    Dim idTAG As ID3v1Data, sTAG As String * 3
    Dim lenFile&
        idTAG.Album = txtAlbum
        idTAG.Artist = txtArtist
        idTAG.Comments = txtComment
        idTAG.Genre = CByte(cboGenre.ListIndex)
        idTAG.Title = txtTitle
        idTAG.Year = txtYear
    lenFile = FileLen(sCurrentFile)
    Open sCurrentFile For Binary As #45
        Get #45, lenFile - 127, sTAG
        If UCase(sTAG) = "TAG" Then
            Put #45, lenFile - 124, idTAG
        Else
            Put #45, lenFile - 127, "TAG"
            Put #45, lenFile - 124, idTAG
        End If
    Close #45
    Unload Me
End If
End Sub

Private Sub Form_Load()

Dim I&
    
    Inet.Cancel
    'Load a few genres
     arrGenre() = Array("Blues", "Classic Rock", "Country", "Dance", "Disco", "Funk", _
    "Grunge", "Hip-Hop", "Jazz", "Metal", "New Age", "Oldies", "Other", "Pop", _
    "R&B", "Rap", "Reggae", "Rock", "Techno", "Industrial", "Alternative", "Ska", _
    "Death Metal", "Pranks", "Soundtrack", "Euro-Techno", "Ambient", "Trip-Hop", "Vocal", _
    "Jazz+Funk", "Fusion", "Trance", "Classical", "Instrumental", "Acid", "House", _
    "Game", "Sound Clip", "Gospel", "Noise", "AlternRock", "Bass", "Soul", "Punk", _
    "Space", "Meditative", "Instrum. Pop", "Instrum. Rock", "Ethnic", "Gothic", "Darkwave", _
    "Techno-Industrial", "Electronic", "Pop-Folk", "Eurodance", "Dream", "Southern Rock", _
    "Comedy", "Cult", "Gangsta", "Top", "Christian Rap", "Pop/Funk", "Jungle", "Native American", _
    "Cabaret", "New Wave", "Psychadelic", "Rave", "Showtunes", "Trailer", "Lo-Fi", "Tribal", _
    "Acid Punk", "Acid Jazz", "Polka", "Retro", "Musical", "Rock & Roll", "Hard Rock", _
    "Folk", "Folk-Rock", "National Folk", "Swing", "Fast Fusion", "Bebob", "Latin", _
    "Revival", "Celtic", "Bluegrass", "Avantgarde", "Gothic Rock", "Prog. Rock", "Psychedel. Rock", _
    "Symph. Rock", "Slow Rock", "Big Band", "Chorus", "Easy Listening", "Acoustic", "Humour", _
    "Speech", "Chanson", "Opera", "Chamber Music", "Sonata", "Symphony", "Booty Bass", "Primus", _
    "Porn Groove", "Satire", "Slow Jam", "Club", "Tango", "Samba", "Folklore", "Ballad", _
    "Power Ballad", "Rhythmic Soul", "Freestyle", "Duet", "Punk Rock", "Drum Solo", "Acapella", _
    "Euro-House", "Dance Hall")
    
    
    For I = 0 To UBound(arrGenre)
        cboGenre.AddItem arrGenre(I)
    Next I
    'check for command line example: myapp.exe mymp3.mp3
    If Command$ <> "" Then LoadData Command$
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'free all memory
    Inet.Cancel
    Erase aAlbums
    If memoArray > 0 Then
        CopyMemory ByVal VarPtrArray(aAlbums), 0&, 4
        Erase aAlbums
    End If
    memoArray = 0
    Unload frmBrow
    Do While Inet.StillExecuting
        End
    Loop
End Sub

Private Sub Inet_StateChanged(ByVal State As Integer)
Dim TypeDoc$, sInfo, I&
Dim PosIni&, PosFin&, sItem$
Dim Art_Alb
Dim bBuffer() As Byte, sData$, lenData&

'Check for state 12=Response completed
If State = 12 Then
    sInfo = Split(Inet.GetHeader(), vbCrLf, , vbBinaryCompare) 'read all http headers
    Select Case sInfo(0)
        Case "HTTP/1.1 404 Not Found"
            lbWait = "Artist/Album Not Found!"
            Inet.Cancel
            Exit Sub
        Case "HTTP/1.1 200 OK"
            TypeDoc = Inet.GetHeader("Content-type")
            If (TypeDoc = "image/jpeg") Then 'Save image
                Open App.Path & "\covers\" & txtArtist & "_" & txtAlbum & ".jpg" For Binary As #33
            End If
            Do
                sData = sData & StrConv(bBuffer, vbUnicode)
                bBuffer = Inet.GetChunk(1024, icByteArray)
                If (TypeDoc = "image/jpeg") Then Put #33, , bBuffer
            Loop Until UBound(bBuffer) = -1
        
        If (TypeDoc = "image/jpeg") Then
            Image1.Picture = LoadPicture(App.Path & "\covers\" & txtArtist & "_" & txtAlbum & ".jpg")
            lbWait.ZOrder 1
            Close #33
            Exit Sub
        End If
        'we foud mor than single albums, parse the page result and read all albums founds
        PosIni = InStr(1, sData, "class=""texto9"">&nbsp;")
        If InStr(TypeDoc, "text/html") And PosIni > 0 Then
            PosFin = InStr(PosIni, sData, "</table>")
            PosFin = InStr(PosFin + 1, sData, "</table>")
            I = 0
            sData = Mid(sData, PosIni, PosFin - PosIni)
            PosIni = 1
            Do
                PosIni = InStr(PosIni, sData, "class=""texto9"">&nbsp;")
                If PosIni > 0 Then
                    PosFin = InStr(PosIni, sData, "</td>")
                    I = I + 1
                    ReDim Preserve aAlbums(1 To I)
                    sItem = Mid(sData, PosIni + 21, PosFin - (PosIni + 21))
                    Art_Alb = Split(sItem, "-")
                    aAlbums(I).sArtist = RTrim(Art_Alb(0))
                    aAlbums(I).sAlbum = LTrim(Art_Alb(1))
                    'set format to http://www.coveralia.com/audio/[first letter artist in lowercase]/[artist name with initial capital and each word joined with "_"]-[album name build like artist name]-frontal.jpg
                    'example http://www.coveralia.com/audio/l/Linkin_Park-Meteora-Frontal.jpg
                    aAlbums(I).sScript = "http://www.coveralia.com/audio/" & LCase(Mid(Art_Alb(0), 1, 1)) & "/" & BuildItem(CStr(Art_Alb(0))) & "-" & BuildItem(LTrim(Art_Alb(1)) & Space(1)) & "-Frontal.jpg"
                    PosIni = PosIni + 21
                End If
            Loop Until PosIni <= 0
            lbWait = "Albums found " & I & vbCrLf & "Click here to pick one"
            lbWait.Enabled = True
            lbWait.ZOrder 0
        Else
            lbWait = "Nothing found!"
        End If
    End Select
End If
End Sub
'Buil the string with initial capital and replace any space with "_"

Private Function BuildItem(xItem As String) As String
Dim sTmp$, PosIni&, PosFin&
PosIni = 0: PosFin = 0
sTmp = ""
    Do
        sTmp = sTmp & UCase(Mid(xItem, PosIni + 1, 1))
        PosFin = InStr(PosIni + 1, xItem, Space(1))
        sTmp = sTmp & LCase(Mid(xItem, PosIni + 2, (PosFin - PosIni) - 2)) & "_"
        PosIni = PosFin
    Loop Until PosFin >= Len(xItem)
    BuildItem = Mid(sTmp, 1, Len(sTmp) - 1)
End Function

Private Sub lbWait_Click()
Dim lLen&
    'pass albums data to memory to use it in frmBrow form
    CopyMemory memoArray, ByVal VarPtrArray(aAlbums()), 4
    
    lLen = UBound(aAlbums())
    Load frmBrow
    frmBrow.BuildAlbums memoArray, lLen
    
End Sub

Private Function LoadData(Mp3File As String)
Dim mp3Data As cls_IdTag
Dim sData$, sGenres$
Dim I&
On Local Error GoTo Fix_Bug
    sCurrentFile = Mp3File
    Image1.Picture = LoadPicture("") 'clear image
   Set mp3Data = New cls_IdTag
    mp3Data.Mp3File = Mp3File 'read mpeg data and build info
    sData = "Size: " & mp3Data.Mp3Len & " bytes" & vbCrLf
    sData = sData & "Header found at " & mp3Data.Header & vbCrLf
    sData = sData & "Seconds: " & mp3Data.Segun2 & vbCrLf
    sData = sData & mp3Data.VersionLayer & vbCrLf
    sData = sData & mp3Data.BitRate & "kbit, " & mp3Data.Frames & " frames" & vbCrLf
    sData = sData & mp3Data.Frequency & "Hz " & mp3Data.Mode & vbCrLf
    lbMpegInfo = sData
    'extract mp3 TAG
    txtArtist = mp3Data.Artist
    txtTitle = mp3Data.Title
    txtAlbum = mp3Data.Album
    txtYear = mp3Data.Year
    txtComment = mp3Data.Comment
    cboGenre.ListIndex = Val(mp3Data.Genre)
    lbWait = "Searching Cover..."
    ImgClock.Visible = True
    '1st look in app cover folder, to check if was downloaded
If Dir(App.Path & "\covers\" & mp3Data.Artist & "_" & mp3Data.Album & ".jpg") = "" Then
    
    Dim jName$ 'Just get the main artist remove "feat", "&","ft","and","with"
    Dim Pos&, sWords()
    'some artist make a "feat" with another artist, to get exact data , we need to use just one artist and remove the rest
    sWords = Array("feat", "&", "ft", "and", "with") 'this list can be upgrade
    For I = 0 To UBound(sWords())
        Pos = InStr(1, LCase(mp3Data.Artist), sWords(I))
        If Pos > 0 Then
            jName = Mid(mp3Data.Artist, 1, (Pos - 2))
            Exit For
        End If
    Next
    If jName = "" Then jName = mp3Data.Artist
    'if we dont' have the albums, we need to look in all albums for this artist , and show to later pick one
    If mp3Data.Album = "" Then
        jName = Replace(jName, Space(1), "+")
        Inet.Execute "http://www.coveralia.com/mostrar.php?bus=" & jName & "&amb=Todo&bust=2", "GET"
    Else
        'if we know the artist and his album, get directly the cover image
        Inet.Execute "http://www.coveralia.com/audio/" & LCase(Mid(jName, 1, 1)) & "/" & Replace(jName, Space(1), "_") & "-" & Replace(mp3Data.Album, Space(1), "_") & "-Frontal.jpg", "GET"
    End If
Else
    'if the image was download previusly, read from local data
    Image1.Picture = LoadPicture(App.Path & "\covers\" & mp3Data.Artist & "_" & mp3Data.Album & ".jpg")
End If
    Set mp3Data = Nothing
Exit Function
Fix_Bug:
    '380 problems with Genre
    If Err.Number <> 380 Then MsgBox Err.Description
    Resume Next
End Function
