Attribute VB_Name = "CommonDialog"
'Paul Mather
'http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
'Modified by Michael W.

Option Explicit

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHOOK As Long) As Long

Private Const GWL_HINSTANCE = (-6)
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 256

Public Const LF_FACESIZE = 32

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOVERTFONTS = &H1000000

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2

Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000

Public Const DN_DEFAULTPRN = &H1

Public Const LOGPIXELSY = 90

'font weights
Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_ULTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_REGULAR = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_DEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_ULTRABOLD = 800
Const FW_HEAVY = 900
Const FW_BLACK = 900
Const ANSI_CHARSET = 0
Const ARABIC_CHARSET = 178
Const BALTIC_CHARSET = 186
Const CHINESEBIG5_CHARSET = 136
Const DEFAULT_CHARSET = 1
Const EASTEUROPE_CHARSET = 238
Const GB2312_CHARSET = 134
Const GREEK_CHARSET = 161
Const HANGEUL_CHARSET = 129
Const HEBREW_CHARSET = 177
Const JOHAB_CHARSET = 130
Const MAC_CHARSET = 77
Const OEM_CHARSET = 255
Const RUSSIAN_CHARSET = 204
Const SHIFTJIS_CHARSET = 128
Const SYMBOL_CHARSET = 2
Const THAI_CHARSET = 222
Const TURKISH_CHARSET = 162
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
Const OUT_OUTLINE_PRECIS = 8
Const OUT_RASTER_PRECIS = 6
Const OUT_STRING_PRECIS = 1
Const OUT_STROKE_PRECIS = 3
Const OUT_TT_ONLY_PRECIS = 7
Const OUT_TT_PRECIS = 4
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_EMBEDDED = 128
Const CLIP_LH_ANGLES = 16
Const CLIP_STROKE_PRECIS = 2
Const ANTIALIASED_QUALITY = 4
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const NONANTIALIASED_QUALITY = 3
Const PROOF_QUALITY = 2
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2
Const FF_DECORATIVE = 80
Const FF_DONTCARE = 0
Const FF_MODERN = 48
Const FF_ROMAN = 16
Const FF_SCRIPT = 64
Const FF_SWISS = 32

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type OPENFILENAME
    nStructSize As Long
    hWndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String        '  May be NULL
End Type

Type CHOOSECOLORS
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
   'lpCustColors As String
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type CHOOSEFONTS
    lStructSize As Long
    hWndOwner As Long          '  caller's window handle
    hDC As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    Flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
    lpszStyle As String          '  return the style field here
    nFontType As Integer          '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
End Type

Type PRINTDLGS
        lStructSize As Long
        hWndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hDC As Long
        Flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Type DEVNAMES
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
End Type

Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type

Public Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled As Boolean
End Type

Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type

'Public FileDialog As OPENFILENAME
'Public ColorDialog As CHOOSECOLORS
'Public FontDialog As CHOOSEFONTS
'Public PrintDialog As PRINTDLGS
Private ParenthWnd As Long
Private hHOOK As Long
Public CommonDialogShowOpen_MultipleCOUNT As Long
Public CommonDialogShowOpen_MultipleARRAY() As String
Public CDLExtensions() As String
Public CommonDialogShowColorERROR As Boolean
Public CommonDialogShowOpenERROR As Boolean
Public CommonDialogShowSaveERROR As Boolean
Public CustomColors(0 To 15) As Long
Global LastCommonDialogDirectory As String

Public Function ShowOpen(ByVal hwnd As Long, vFileDialog As OPENFILENAME, Optional ByVal CenterFormA As Boolean = False) As SelectedFile
Dim Ret As Long
Dim Count As Integer
Dim fileNameHolder As String
Dim LastCharacter As Integer
Dim NewCharacter As Integer
Dim tempFiles(1 To 200) As String
Dim hInst As Long
Dim Thread As Long
CommonDialogShowOpenERROR = False
    ParenthWnd = hwnd
    With vFileDialog
      .nStructSize = Len(vFileDialog)
      .hWndOwner = hwnd
      .sFileTitle = Space$(2048)
      .nTitleSize = Len(.sFileTitle)
      .sFile = .sFile & Space$(2047) & Chr$(0)
      .nFileSize = Len(.sFile)
    
    'If FileDialog.flags = 0 Then
        .Flags = OFS_FILE_OPEN_FLAGS
    'End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterFormA = True Then
        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    Ret = GetOpenFileName(vFileDialog)

    If Ret Then
        If Trim$(.sFileTitle) = "" Then
            LastCharacter = 0
            Count = 0
            While ShowOpen.nFilesSelected = 0
                NewCharacter = InStr(LastCharacter + 1, .sFile, Chr$(0), vbTextCompare)
                If Count > 0 Then
                    tempFiles(Count) = Mid(.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                Else
                    ShowOpen.sLastDirectory = Mid(.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                End If
                Count = Count + 1
                If InStr(NewCharacter + 1, .sFile, Chr$(0), vbTextCompare) = InStr(NewCharacter + 1, .sFile, Chr$(0) & Chr$(0), vbTextCompare) Then
                    tempFiles(Count) = Mid(.sFile, NewCharacter + 1, InStr(NewCharacter + 1, .sFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = Count
                End If
                LastCharacter = NewCharacter
            Wend
            ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)
            For Count = 1 To ShowOpen.nFilesSelected
                ShowOpen.sFiles(Count) = tempFiles(Count)
            Next
        Else
            ReDim ShowOpen.sFiles(1 To 1)
            ShowOpen.sLastDirectory = Left$(.sFile, .nFileOffset)
            ShowOpen.nFilesSelected = 1
            ShowOpen.sFiles(1) = Mid(.sFile, .nFileOffset + 1, InStr(1, .sFile, Chr$(0), vbTextCompare) - .nFileOffset - 1)
        End If
        ShowOpen.bCanceled = False
        Exit Function
    Else
        CommonDialogShowOpenERROR = True
        ShowOpen.sLastDirectory = ""
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        Erase ShowOpen.sFiles
        Exit Function
    End If
End With
End Function
Public Function ShowSave(ByVal hwnd As Long, vFileDialog As OPENFILENAME, ByRef vFileExtensions() As String, vFileExtensionsCount As Long, Optional vInitialFileName As String, Optional ByVal CenterFormA As Boolean = False) As SelectedFile
Dim Ret As Long
Dim hInst As Long
Dim Thread As Long
CommonDialogShowSaveERROR = False
    ParenthWnd = hwnd
    With vFileDialog
      .hWndOwner = hwnd
    If vInitialFileName <> "" Then
      .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
      .sFile = vInitialFileName & Space$(1024) & vbNullChar & vbNullChar
    Else
      .sFileTitle = Space$(2048)
      .sFile = Space$(2047) & Chr$(0)
    End If
    
    .nStructSize = Len(vFileDialog)
    .nTitleSize = Len(.sFileTitle)
    .nFileSize = Len(.sFile)
    
    If .Flags = 0 Then
        .Flags = OFS_FILE_SAVE_FLAGS
    End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterFormA = True Then
        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    Ret = GetSaveFileName(vFileDialog)
    ReDim ShowSave.sFiles(1)

    If Ret Then
        ShowSave.sLastDirectory = Left$(.sFile, .nFileOffset)
        ShowSave.nFilesSelected = 1
        ShowSave.sFiles(1) = Mid(.sFile, .nFileOffset + 1, InStr(1, .sFile, Chr$(0), vbTextCompare) - .nFileOffset - 1)
        ShowSave.bCanceled = False
        Exit Function
    Else
        CommonDialogShowSaveERROR = True
        ShowSave.sLastDirectory = ""
        ShowSave.nFilesSelected = 0
        ShowSave.bCanceled = True
        Erase ShowSave.sFiles
        Exit Function
    End If
    End With
End Function
Public Function ShowColor(ByVal hwnd As Long, vColorDialog As CHOOSECOLORS, Optional ByVal CenterFormA As Boolean = False) As SelectedColor
Dim CustomColors() As Byte  ' dynamic (resizable) array
Dim i As Integer
Dim Ret As Long
Dim hInst As Long
Dim Thread As Long

    ParenthWnd = hwnd
' commented this out because of string/long conflict
'
'    If vColorDialog.lpCustColors = "" Then
'        ReDim CustomColors(0 To 16 * 4 - 1) As Byte  'resize the array
'
'        For i = LBound(CustomColors) To UBound(CustomColors)
'          CustomColors(i) = 254 ' sets all custom colors to white
'        Next i
'
'        vColorDialog.lpCustColors = StrConv(CustomColors, vbUnicode)  ' convert array
'    End If
    
    vColorDialog.hWndOwner = hwnd
    vColorDialog.lStructSize = Len(vColorDialog)
    vColorDialog.Flags = COLOR_FLAGS
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterFormA = True Then
        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    Ret = ChooseColor(vColorDialog)
    If Ret Then
        ShowColor.bCanceled = False
        ShowColor.oSelectedColor = vColorDialog.rgbResult
        Exit Function
    Else
        ShowColor.bCanceled = True
        ShowColor.oSelectedColor = &H0&
        Exit Function
    End If
End Function
Public Function ShowFont_Original(ByVal hwnd As Long, ByVal startingFontName As String, Optional ByVal CenterFormA As Boolean = False) As SelectedFont
'Dim ret As Long
'Dim lfLogFont As LOGFONT
'Dim hInst As Long
'Dim Thread As Long
'Dim i As Integer
'
'    ParenthWnd = hWnd
'    FontDialog.nSizeMax = 0
'    FontDialog.nSizeMin = 0
'    FontDialog.nFontType = Screen.FontCount
'    FontDialog.hWndOwner = hWnd
'    FontDialog.hDC = 0
'    FontDialog.lpfnHook = 0
'    FontDialog.lCustData = 0
'    FontDialog.lpLogFont = VarPtr(lfLogFont)
'    If FontDialog.iPointSize = 0 Then
'        FontDialog.iPointSize = 10 * 10
'    End If
'    FontDialog.lpTemplateName = Space$(2048)
'    FontDialog.rgbColors = RGB(0, 255, 255)
'    FontDialog.lStructSize = Len(FontDialog)
'
'    If FontDialog.flags = 0 Then
'        FontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT 'Or CF_EFFECTS
'    End If
'
'    For i = 0 To Len(startingFontName) - 1
'        lfLogFont.lfFaceName(i) = Asc(Mid(startingFontName, i + 1, 1))
'    Next
'
'    'Set up the CBT hook
'    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
'    Thread = GetCurrentThreadId()
'    If CenterFormA = True Then
'        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
'    Else
'        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
'    End If
'
'    ret = ChooseFont(FontDialog)
'
'    If ret Then
'        ShowFont.bCanceled = False
'        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
'        ShowFont.bItalic = lfLogFont.lfItalic
'        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
'        ShowFont.bUnderline = lfLogFont.lfUnderline
'        ShowFont.lColor = FontDialog.rgbColors
'        ShowFont.nSize = FontDialog.iPointSize / 10
'        For i = 0 To 31
'            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(i))
'        Next
'
'        ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
'        Exit Function
'    Else
'        ShowFont.bCanceled = True
'        Exit Function
'    End If
End Function
Public Function ShowPrinter(ByVal hwnd As Long, Optional ByVal CenterFormA As Boolean = False) As Long
'Dim hInst As Long
'Dim Thread As Long
'
'    ParenthWnd = hWnd
'    PrintDialog.hWndOwner = hWnd
'    PrintDialog.lStructSize = Len(PrintDialog)
'
'    'Set up the CBT hook
'    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
'    Thread = GetCurrentThreadId()
'    If CenterFormA = True Then
'        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
'    Else
'        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
'    End If
'
'    ShowPrinter = PrintDlg(PrintDialog)
End Function
Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    If lMsg = HCBT_ACTIVATE Then
        'Show the MsgBox at a fixed location (0,0)
        GetWindowRect wParam, rectMsg
        X = ((Screen.Width / Screen.TwipsPerPixelX) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
        Y = ((Screen.Height / Screen.TwipsPerPixelY) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
        Debug.Print "Screen " & Screen.Height / 2
        Debug.Print "MsgBox " & (rectMsg.Right - rectMsg.Left) / 2
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHOOK
    End If
    WinProcCenterScreen = False
End Function
Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim X As Long, Y As Long
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect ParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        X = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
        Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
        ' Michael changed this because of
        ' problems if the form is not visible
        'X = ((Screen.Width / Screen.TwipsPerPixelX) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
        'Y = ((Screen.Height / Screen.TwipsPerPixelY) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
        'Position the msgbox
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHOOK
     End If
     WinProcCenterForm = False
End Function
Public Function CommonDialogShowOpen(ByVal vInitDir As String, ByVal vDialogTitle As String, ByVal vFilter As Variant, Frm As Form, Optional UseLastDirectory As Boolean = False) As String
Dim sOpen As SelectedFile
Dim CDL1 As OPENFILENAME

On Error GoTo ErrHand

If Not UseLastDirectory Then
  CDL1.sInitDir = vInitDir
ElseIf UseLastDirectory Then
  If LastCommonDialogDirectory = "" Then
    CDL1.sInitDir = vInitDir
  Else
    CDL1.sInitDir = LastCommonDialogDirectory
  End If
End If
CDL1.sDlgTitle = vDialogTitle
CDL1.sFilter = vFilter
CDL1.Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
sOpen = ShowOpen(Frm.hwnd, CDL1)
If Err.Number <> 32755 And sOpen.bCanceled = False Then
  CommonDialogShowOpen = sOpen.sLastDirectory & sOpen.sFiles(1)
  LastCommonDialogDirectory = sOpen.sLastDirectory
End If
Exit Function
ErrHand:
  Oops "Function CommonDialogShowOpen"
End Function
Public Function CommonDialogShowOpen_Multiple(ByVal vInitDir As String, ByVal vDialogTitle As String, ByVal vFilter As Variant, Frm As Form, Optional UseLastDirectory As Boolean = False) As Boolean
Dim sOpen As SelectedFile
Dim sCount As Long
Dim CDL1 As OPENFILENAME

On Error GoTo ErrHand

CDL1.Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT

If Not UseLastDirectory Then
  CDL1.sInitDir = vInitDir
ElseIf UseLastDirectory Then
  If LastCommonDialogDirectory = "" Then
    CDL1.sInitDir = vInitDir
  Else
    CDL1.sInitDir = LastCommonDialogDirectory
  End If
End If

CDL1.sDlgTitle = vDialogTitle
CDL1.sFilter = vFilter
sOpen = ShowOpen(Frm.hwnd, CDL1)
If Err.Number <> 32755 And sOpen.bCanceled = False Then
  ReDim CommonDialogShowOpen_MultipleARRAY(1 To sOpen.nFilesSelected)
  For sCount = 1 To sOpen.nFilesSelected
    CommonDialogShowOpen_MultipleARRAY(sCount) = sOpen.sLastDirectory & sOpen.sFiles(sCount)
    LastCommonDialogDirectory = sOpen.sLastDirectory
  Next sCount
End If
CommonDialogShowOpen_MultipleCOUNT = sCount
CommonDialogShowOpen_Multiple = True
Exit Function
ErrHand:
  CommonDialogShowOpen_Multiple = False
  Oops "Function CommonDialogShowOpen_Multiple"
End Function
Public Function CommonDialogShowSave(ByVal vInitDir As String, ByVal vDialogTitle As String, ByVal vFilter As Variant, Frm As Form, ByRef vFileExtensions() As String, vFileExtensionsCount As Long, Optional vInitialFileName As String, Optional vDefaultExtension As String, Optional UseLastDirectory As Boolean = False) As String
Dim sSave As SelectedFile
Dim CDL1 As OPENFILENAME

On Error GoTo ErrHand

If Not UseLastDirectory Then
  CDL1.sInitDir = vInitDir
ElseIf UseLastDirectory Then
  If LastCommonDialogDirectory = "" Then
    CDL1.sInitDir = vInitDir
  Else
    CDL1.sInitDir = LastCommonDialogDirectory
  End If
End If
CDL1.sDlgTitle = vDialogTitle
CDL1.sFilter = vFilter

If vDefaultExtension <> "" Then CDL1.sDefFileExt = vDefaultExtension
'CDL1.flags = OFN_EXPLORER 'OFS_FILE_SAVE_FLAGS 'Or OFS_MAXPATHNAME
'CDL1.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
sSave = ShowSave(Frm.hwnd, CDL1, vFileExtensions(), vFileExtensionsCount, vInitialFileName)
If Err.Number <> 32755 And sSave.bCanceled = False Then
Dim z1, Sat
z1 = CDL1.nFilterIndex
If vFileExtensionsCount = 0 Then
  CommonDialogShowSave = sSave.sLastDirectory & sSave.sFiles(1)
  LastCommonDialogDirectory = sSave.sLastDirectory
  Exit Function
End If

For Sat = 1 To vFileExtensionsCount
  If InStr(1, sSave.sFiles(1), "." & vFileExtensions(Sat), vbTextCompare) <> 0 Then
    CommonDialogShowSave = sSave.sLastDirectory & sSave.sFiles(1)
    LastCommonDialogDirectory = sSave.sLastDirectory
    Exit Function
  End If
Next Sat

  CommonDialogShowSave = sSave.sLastDirectory & sSave.sFiles(1) & "." & vFileExtensions(z1)
  LastCommonDialogDirectory = sSave.sLastDirectory
'  If InStrRev(CommonDialogShowSave, ".", 1, vbTextCompare) <> 0 Then
'    Exit Function
'  Else
'    CommonDialogShowSave = CommonDialogShowSave & CDL1.sFilter
'  End If
End If
Exit Function
ErrHand:
  'Oops "Function CommonDialogShowSave"
End Function
Public Function CommonDialogShowColor(Frm As Form, Optional initColor As Long) As Long
Dim vColorDialog As CHOOSECOLORS
Dim sColor As SelectedColor
Dim Mine2 As String, z, MyK As Long
Dim FR As String, fs As String
CommonDialogShowColorERROR = False
On Error GoTo ErrHand
  MyK = &H80000001
  Mine2 = "Software\Helper\Custom Colors"
  For z = 0 To 15
    FR = z
    CustomColors(z) = Val(GetString(MyK, Mine2, FR))
    If CustomColors(z) = 0 Then CustomColors(z) = 16777215
  Next z
  vColorDialog.lpCustColors = VarPtr(CustomColors(0))
  
  If initColor <> 0 Then
    vColorDialog.rgbResult = initColor
  End If
  vColorDialog.lStructSize = Len(vColorDialog)
  sColor = ShowColor(Frm.hwnd, vColorDialog)
  If sColor.bCanceled = True Then
    CommonDialogShowColor = -1
    CommonDialogShowColorERROR = True
    Exit Function
  Else
    CommonDialogShowColor = sColor.oSelectedColor
  End If
  For z = 0 To 15
    FR = z
    fs = CustomColors(z)
    SaveString MyK, Mine2, FR, fs
  Next z
Exit Function
ErrHand:
  Oops "Function CommonDialogShowColor"
End Function
Public Function CommonDialogShowFont(Frm As Form, vObject As Control) As SelectedFont
Dim sFont As SelectedFont
  With sFont
    .bBold = vObject.FontBold
    .bItalic = vObject.FontItalic
    .sFaceName = vObject.FontName
    .nSize = vObject.FontSize
    .bStrikeOut = vObject.FontStrikethru
    .bUnderline = vObject.FontUnderline
    .lColor = vObject.ForeColor
  End With
    
  CommonDialogShowFont = ShowFont(Frm.hwnd, sFont, vObject)
End Function
Public Function ShowFont(ByVal hwnd As Long, ByRef vFormInfo As SelectedFont, Optional vObject As Control, Optional ByVal CenterFormA As Boolean = False, Optional bBypassControl As Boolean = False) As SelectedFont
Dim Ret As Long
Dim lfLogFont As LOGFONT
Dim hInst As Long
Dim Thread As Long
Dim i As Integer
Dim FontDialog As CHOOSEFONTS
    ParenthWnd = hwnd
    lfLogFont.lfItalic = vFormInfo.bItalic * -1
    lfLogFont.lfStrikeOut = vFormInfo.bStrikeOut * -1
    lfLogFont.lfUnderline = vFormInfo.bUnderline * -1
    lfLogFont.lfHeight = -MulDiv(CLng(vFormInfo.nSize), GetDeviceCaps(GetDC(hwnd), LOGPIXELSY), 72)
    If vFormInfo.bBold = True Then
      lfLogFont.lfWeight = FW_BOLD
    Else
      lfLogFont.lfWeight = FW_NORMAL
    End If
    
    For i = 0 To Len(vFormInfo.sFaceName) - 1
        lfLogFont.lfFaceName(i) = Asc(Mid(vFormInfo.sFaceName, i + 1, 1))
    Next
    
    FontDialog.lpLogFont = VarPtr(lfLogFont)
    FontDialog.iPointSize = vFormInfo.nSize * 10
    FontDialog.rgbColors = vFormInfo.lColor
    FontDialog.Flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    FontDialog.lStructSize = Len(FontDialog)
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If CenterFormA = True Then
        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHOOK = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    Ret = ChooseFont(FontDialog)
        
    If Ret Then
        ShowFont.bCanceled = False
        ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
        ShowFont.bItalic = lfLogFont.lfItalic
        ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
        ShowFont.bUnderline = lfLogFont.lfUnderline
        ShowFont.lColor = FontDialog.rgbColors
        ShowFont.nSize = FontDialog.iPointSize / 10
        For i = 0 To 31
            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(i))
        Next
    
        ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
        
        If Not bBypassControl Then
            With vObject
              .FontBold = ShowFont.bBold
              .FontItalic = ShowFont.bItalic
              .FontName = ShowFont.sSelectedFont
              .FontSize = ShowFont.nSize
              .FontStrikethru = ShowFont.bStrikeOut
              .FontUnderline = ShowFont.bUnderline
              .ForeColor = ShowFont.lColor
            End With
        End If
        Exit Function
    Else
        ShowFont.bCanceled = True
        Exit Function
    End If
End Function
Private Function MulDiv(In1 As Long, In2 As Long, In3 As Long) As Long
Dim lngTemp As Long
  On Error GoTo MulDiv_err
  If In3 <> 0 Then
    lngTemp = In1 * In2
    lngTemp = lngTemp / In3
  Else
    lngTemp = -1
  End If
MulDiv_end:
  MulDiv = lngTemp
  Exit Function
MulDiv_err:
  lngTemp = -1
  Resume MulDiv_err
End Function
Private Sub Oops(Optional Location As String)
  MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "LOCATION:" & vbCrLf & Location, vbExclamation, "Oops!"
End Sub
Public Sub DeleteCustomColors()
Dim Mine2 As String, z, MyK As Long

On Error GoTo ErrHand
  MyK = &H80000001
  Mine2 = "Software\Helper\Custom Colors"
  For z = 0 To 15
    DeleteValue MyK, Mine2, z
  Next z
ErrHand:
End Sub
