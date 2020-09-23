VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Michael's Helper Menu Handler"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long

Private Const PAGE_EXECUTE_READWRITE    As Long = &H40

Implements IContextMenu
Implements IContextMenu2
Implements IContextMenu3
Implements IShellExtInit

Private m_sFiles()            As String
Private m_lFileCount          As Long
Private m_pOldQueryCtxMenu    As Long

Private mnuIDs                As Collection
Private sTag()                As String
Private lTagCount             As Long
Private lngTimesCalled        As Long
Private bDisableOwnerDraw     As Boolean
Private bContainsFolders      As Boolean
Private sWhereApp$

Private Sub Class_Initialize()

  Dim pVtable As IContextMenu
  sWhereApp = App.Path & "\musicProp.exe "

  Set pVtable = Me
  m_pOldQueryCtxMenu = SwapVtableEntry(ObjPtr(pVtable), 4, AddressOf QueryContextMenuVB)

End Sub

Private Sub Class_Terminate()

  Dim pVtable As IContextMenu

  Set pVtable = Me
  m_pOldQueryCtxMenu = SwapVtableEntry(ObjPtr(pVtable), 4, m_pOldQueryCtxMenu)

End Sub

Public Function SwapVtableEntry(pObj As Long, ByVal EntryNumber As Integer, ByVal lpfn As Long) As Long
  
  'A variable of type IContextMenu is set to Me.
  'This gives us a pointer to the IContextMenu
  'portion of the vtable. This memory location is copied
  'into pVtable, effectively giving us a pointer to the
  'IContextMenu portion of our object's vtable.
  'Then, SwapVtableEntry is called with the address of
  'the first method of IContextMenu
  '(this is the portion of the vtable where IContextMenu begins),
  'the relative position in the vtable of the method we want
  'to replace, and the address of the new function.
  'One thing of interest in SwapVtableEntry is the call to
  'VirtualProtect. VB has marked the object memory as protected.
  'This call changes the access permissions,
  'allowing us to swap the addresses.
  
  'this explanation is staright from here
  'http://www.oreilly.com/catalog/vbshell/chapter/ch04.html#41689
  
  Dim lOldAddr     As Long
  Dim lpVtableHead As Long
  Dim lpfnAddr     As Long
  Dim lOldProtect  As Long
  
  CopyMemory lpVtableHead, ByVal pObj, 4
  lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
  CopyMemory lOldAddr, ByVal lpfnAddr, 4
  VirtualProtect lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect
  CopyMemory ByVal lpfnAddr, lpfn, 4
  VirtualProtect lpfnAddr, 4, lOldProtect, lOldProtect
  SwapVtableEntry = lOldAddr

End Function

Private Sub IContextMenu_QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As shlctxmnu.QueryContextMenuFlags)
'
End Sub

Private Sub IContextMenu2_GetCommandString(ByVal idCmd As Long, ByVal uType As shlctxmnu.GETCOMMANDSTRINGFLAGS, pwReserved As Long, ByVal pszName As Long, ByVal cchMax As Long)
'
End Sub

Private Sub IContextMenu2_InvokeCommand(lpici As shlctxmnu.CMINVOKECOMMANDINFO)
'
End Sub

Private Sub IContextMenu2_QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As shlctxmnu.QueryContextMenuFlags)
'
End Sub

Private Sub IContextMenu3_GetCommandString(ByVal idCmd As Long, ByVal uType As shlctxmnu.GETCOMMANDSTRINGFLAGS, pwReserved As Long, ByVal pszName As Long, ByVal cchMax As Long)
'
End Sub

Private Sub IContextMenu3_HandleMenuMsg(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'
End Sub

Private Sub IContextMenu3_InvokeCommand(lpici As shlctxmnu.CMINVOKECOMMANDINFO)
'
End Sub

Private Sub IContextMenu3_QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As shlctxmnu.QueryContextMenuFlags)
'
End Sub

' This procedure is called before any other extension interface to pass the selected
' files to the shell extension
Private Sub IShellExtInit_Initialize(ByVal pidlFolder As Long, ByVal lpIDataObject As shlctxmnu.IDataObject, ByVal hkeyProgID As Long)

  'for an explanation of this look here
  'http://www.oreilly.com/catalog/vbshell/chapter/ch04.html#41689
  
  'msdn says:
  'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/programmersguide/shell_int/shell_int_extending/extensionhandlers/shell_ext.asp
  'pIDFolder holds a folder's pointer to an item identifier list (PIDL).
  '  For property sheet extensions, it is NULL. For shortcut menu extensions,
  '  it is the PIDL of the folder that contains the item whose shortcut menu
  '  is being displayed. For nondefault drag-and-drop handlers, it is
  '  the PIDL of the target folder.
  'pDataObject holds a pointer to a data object's IDataObject interface.
  '  The data object holds one or more file names in CF_HDROP format.
  'hRegKey holds a registry key for the file object or folder type.
  
  'this is standard code for the IShellExtInit_Initialize method -
  'the only method it has.

  'Most shell extension handlers must
  'implement either an IPersistFile or IShellExtInit interface.
  'The Shell uses these interfaces to initialize the handler.

  Dim fmtEtc  As FORMATETC
  Dim pmedium As STGMEDIUM
  Dim i       As Long
  Dim lResult As Long
  Dim sTemp   As String

  With fmtEtc
    .cfFormat = CF_HDROP
    .dwAspect = DVASPECT_CONTENT
    .lIndex = -1
    .TYMED = TYMED_HGLOBAL
  End With
  
  lpIDataObject.GetData fmtEtc, pmedium
  m_lFileCount = DragQueryFile(pmedium.Data, &HFFFFFFFF, vbNullString, 0)
  
  ReDim m_sFiles(m_lFileCount - 1)
  
  For i = 0 To (m_lFileCount - 1)
    sTemp = String$(1024, 0)
    lResult = DragQueryFile(pmedium.Data, i, sTemp, Len(sTemp))
    If lResult > 0 Then
        m_sFiles(i) = Left$(sTemp, lResult)
    End If
  Next i
  
  ReleaseStgMedium pmedium

End Sub

Private Sub IContextMenu_GetCommandString(ByVal idCmd As Long, ByVal uType As shlctxmnu.GETCOMMANDSTRINGFLAGS, pwReserved As Long, ByVal pszName As Long, ByVal cchMax As Long)
  
  'GetCommandString allows the handler to specify
  'the text that will be displayed in the status
  'bar of Explorer. This occurs when a particular
  'context menu item is selected.
  
  'idCmd - The ordinal position of the selected menu item.
  'uFlags - A flag specifying the information to return.
  'pwReserved - Unused; handlers must ignore this parameter, which should be set to NULL.
  'pszName - A pointer to the string buffer that holds the null-terminated string to be displayed.
  'cchMax - Size of the buffer defined by pszName.

  Dim szName As String
  Dim bszName() As Byte

  Dim sMenuHelp As String

  'again we'll use our string array to simplify things
  
    If sTag(idCmd) = "Full mp3 view" Then szName = "View full mp3 data."
    szName = Left$(szName, cchMax) & vbNullChar

  Select Case uType
    Case GCS_VERBA, GCS_HELPTEXTA, GCS_VALIDATEA
      If LenB(szName) Then
        bszName = StrConv(szName, vbFromUnicode)
        CopyMemory ByVal pszName, bszName(0), UBound(bszName) + 1
      End If
    Case GCS_VERBW, GCS_HELPTEXTW, GCS_VALIDATEW
      If LenB(szName) Then
        bszName = szName
        CopyMemory ByVal pszName, bszName(0), UBound(bszName) + 1
      End If
  End Select

End Sub

Private Sub IContextMenu2_HandleMenuMsg(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  'DebugA "IContextMenu2_HandleMenuMsg"
  Dim lDummy&
  HandleMenuMsgVB uMsg, wParam, lParam, lDummy
End Sub

Private Sub IContextMenu3_HandleMenuMsg2(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, plResult As Long)
  
  'straight from msdn:
  
  'IContextMenu3::HandleMenuMsg2 generally supplants
  'IContextMenu2::HandleMenuMsg, and is called when IContextMenu
  'determines that IContextMenu3 is supported and one of the
  'supported messages (see uMsg) has been received. However, in some cases,
  'IContextMenu2::HandleMenuMsg is still called. Context menu hosts may
  'dispatch menu messages through either or both methods. Consequently,
  'if a shell extension implements both IContextMenu2::HandleMenuMsg and
  'IContextMenu3::HandleMenuMsg2, it must be prepared for menu messages
  'to arrive through either method.

  'Note:  If IContextMenu3 is not implemented, there is no guarantee
  'that IContextMenu2 will be called in its place. In some cases, the absence
  'of IContextMenu3 is determined and then the process is halted.
  
  'DebugA "IContextMenu3_HandleMenuMsg2"
  HandleMenuMsgVB uMsg, wParam, lParam, plResult
  
End Sub

Private Sub IContextMenu_InvokeCommand(lpici As shlctxmnu.CMINVOKECOMMANDINFO)
  
  Dim sTemp  As String
  Dim ZX     As Long
  Dim sExt   As String
  Dim bRes   As Boolean

  On Error GoTo ErrHand
  
  'this is where we handle what item was clicked on
  'instead of trying to remember menu ids i chose to use a string array
  'to remember the caption we added.
  'that way if we change the order of the menus - it doesn't matter
  
  sExt = LCase(FileExtension(m_sFiles(0)))
  
  'LOWORD(cmi.lpVerb) = 0  vbshell
  '(lpici.lpVerb \ &H10000) <> 0 edanmo
  
    If sTag(lpici.lpVerb) = "Full mp3 view" Then
        Shell sWhereApp & GetShortFileName(m_sFiles(0)), vbNormalFocus
    End If
  Exit Sub

ErrHand:

  MsgBox "Error : " & Err.Description

End Sub

Friend Function QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As Long) As Long

  Dim idCmd        As Long
  Dim hMenuPop1    As Long
  Dim hMenuPop2    As Long
  
  'this is where we actually add menu items
  'some filtering is used depending on file extension
  'whether or not its a folder and if multiple files
  'were selected - all handled in addmenuitem
  
  Dim sExt As String, lFileCount As Long
  
  lFileCount = m_lFileCount
  sExt = LCase(FileExtension(m_sFiles(0)))
  'DebugA "QueryContextMenu Extension: " & sExt

  On Error Resume Next
  
  bDisableOwnerDraw = False
  
  If (uFlags And CMF_VERBSONLY) Then
    'This is a context menu for a shortcut item.
    
    'we've registered as a shortcut menu handler -
    'this should never get called
    'but lets be safe
    'shortcuts can cause major problems
    
    'handlemenumsg never gets called
    'in this case so disable ownerdraw
    'to insure our menuitems aren't blank
    bDisableOwnerDraw = True

  End If
  
  bContainsFolders = ContainsFolders(m_sFiles())
    
  Set mnuIDs = New Collection

  lTagCount = 0

  idCmd = idCmdFirst
  'create submenus - note that they remain nothing
  'unless you add something to it
  hMenuPop1 = CreatePopupMenu()
  hMenuPop2 = CreatePopupMenu()
  
  'this is the main menu with a separator on either side of it
  AddSeparator hMenu, idCmd, indexMenu, lFileCount
  If AddMenuItem(hMenu, idCmd, indexMenu, lFileCount, sExt, "Full mp3 view", 2, , , , False, "mp3", False) Then
    idCmd = idCmd + 1
    indexMenu = indexMenu + 1
  End If
  AddSeparator hMenu, idCmd, indexMenu, lFileCount

  
ExitLabel:

  QueryContextMenu = indexMenu
    
End Function

Public Function AddMenuItem(ByVal hMenu As Long, ByVal idCmd As Long, ByVal indexMenu As Long, ByVal lFileCount As Long, ByVal sExtension As String, ByVal sCaption As String, Optional ByVal lPictureIndex As Long = -1, Optional ByVal lSubMenu As Long = 0, Optional bEnabled As Boolean = True, Optional bChecked As Boolean = False, Optional bSupportsMultiFiles As Boolean = True, Optional sFileExtensionSpecific As String = "*.*", Optional bSupportsFolders As Boolean = True) As Boolean

  'ok this looks horribly confusing, but this was the method i chose
  
  Dim MenuIIMain As MENUITEMINFO
  
  'if our item doesn't work on multiple files
  'but multiple files are selected - don't add the menu item
  If (Not bSupportsMultiFiles) And (lFileCount > 1) Then
    Exit Function
  End If
  
  'if our item doesn't work on folders,
  'but at least one item selected is a folder
  'don't add the menu item
  If (bContainsFolders) And (Not bSupportsFolders) Then
    Exit Function
  End If
  
  'if out item is file extension specific
  'we passed the file extension of the first file to
  'this procedure. if it doesn't the file extension we want
  'then don't add the menu item
  If sFileExtensionSpecific <> "*.*" Then
    Dim sArr() As String, ZX As Long
    SplitB sFileExtensionSpecific, sArr, ";"
    For ZX = LBound(sArr) To UBound(sArr)
      If sExtension = sArr(ZX) Then
        GoTo ExtensionFound
      End If
    Next
    Exit Function
  End If

ExtensionFound:

  With MenuIIMain
    .cbSize = Len(MenuIIMain)
    .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
    .fType = MF_STRING
    'i don't thinnk i'll ever add a disabled item
    'i think i'd rather just not add the menu
    'but here is the option for doing so
    If bEnabled Then
      .fState = MF_ENABLED
    Else
      .fState = MF_DISABLED Or MF_GRAYED
    End If
    'if we want our item checked or not
    If bChecked Then
      .fState = .fState Or MF_CHECKED
    End If
    .wid = idCmd
    .dwTypeData = sCaption
    .cch = Len(sCaption)
    .hSubMenu = lSubMenu
  End With
  
  'if this item is to have a submenu then assign it
  'this will create the right arrow on the menu
  If lSubMenu <> 0 Then
    MenuIIMain.fMask = MenuIIMain.fMask Or MIIM_SUBMENU
  End If
  
  If lPictureIndex <> -1 Then
    'we only declare the menu item as owner drawn if
    'we've assigned an icon to it (an index in the imagelist of frmicon)
    'this is our shortcut failsafe
    'if our boolean flag got tripped
    'we don't make it ownerdrawn even if it does
    'have an icon assigned to it
    If Not bDisableOwnerDraw Then
      MenuIIMain.fType = MF_OWNERDRAW
    End If
    
    'mnuinfo is a class module we'll use for holding various info
    'we'll use in the owner drawing in handlemenumsgvb
    Dim myID As New mnuInfo
    With myID
      .Caption = sCaption
      .ItemID = idCmd
      .PictureIndex = lPictureIndex
    End With
    mnuIDs.Add myID
    Set myID = Nothing
  End If
  
  InsertMenuItem hMenu, 0, indexMenu, MenuIIMain
  
  'here we keep an array of the captions we use in our menuitems
  'the point of this is to make handling clicks
  'in IContextMenu_InvokeCommand SO much easier.
  ReDim Preserve sTag(lTagCount)
  sTag(lTagCount) = sCaption
  lTagCount = lTagCount + 1
  
  AddMenuItem = True
    
End Function

Public Sub AddSeparator(ByVal hMenu As Long, ByVal idCmd As Long, ByVal indexMenu As Long, ByVal lFileCount As Long, Optional bSupportsMultiFiles As Boolean = True, Optional lPositionOffset As Long = 0)

  'add a menu separator
  Dim MenuIISep As MENUITEMINFO
  
  'this isn't likely to be used, but its still here
  If (Not bSupportsMultiFiles) And (lFileCount > 1) Then
    Exit Sub
  End If
  
  With MenuIISep
    .cbSize = Len(MenuIISep)
    .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
    .fType = MF_SEPARATOR
    .fState = MF_ENABLED
    .wid = idCmd
    .dwTypeData = ""
    .cch = Len("")
    .hSubMenu = 0
  End With

  'in various instances i want to place the separator immediately before
  'the last menu item i inserted (and that being dependent on whether
  'it actually got inserted or not - hence the boolean return of addmenuitem)
  'when this is the case lPositionOffset is -1 effectively
  'placing the separator in the desired position
  If lPositionOffset <> 0 Then
    'custom insert
    InsertMenuItem hMenu, idCmd + lPositionOffset, False, MenuIISep
  Else
    InsertMenuItem hMenu, 0, indexMenu, MenuIISep
  End If
  
End Sub

Private Function HandleMenuMsgVB(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, plResult As Long)

  Dim MeasureInfo As MEASUREITEMSTRUCT
  Dim DrawInfo    As DRAWITEMSTRUCT
  
  Dim myID        As mnuInfo
  Dim r           As RECT
  Dim chkR        As RECT
  Dim IsSelected  As Boolean
  Dim IsChecked   As Boolean
  Dim IsDisabled  As Boolean
  
  'this is where we receive messages and do our drawing
  'if i'm not mistaken, this procedure only receives 3 messages
  'WM_MEASUREITEM, WM_DRAWITEM, and WM_CHAR - this being specific
  'to IContextMenu2 and IContextMenu3
  'either or both can be called which is why we have this sub
  'rather than code directly in the item event
  
  'for the specifics of the actual drawing
  'Maybe you ought to ask MrBobo - see credits.txt
  'the basics of this entire function were derived from MrBobo here
  'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=44758&lngWId=1
  
  'the numbers he used did not work for this particular menu, though.
  'using his numbers the text was offset about 2 characters to
  'the right - so i conclude that where text starts must be different
  'between a regular vb menu and a context menu - or perhaps just
  'the shell menu. If the text or icon don't line up right you may
  'have to adjust the numbers yourself.
  
  'DebugA "HandleMenuMsgVB Msg: " & uMsg
  
  Select Case uMsg
    
    Case WM_MEASUREITEM '43

      Call CopyMemory(MeasureInfo, ByVal lParam, Len(MeasureInfo)) 'get the UDT for the menuitem's dimensions
      
      If ValidId(myID, MeasureInfo.ItemID) Then
        MeasureInfo.ItemHeight = 20 'this is constant images will always be 16x16 with a 2 padding on either side
        MeasureInfo.ItemWidth = (frmIcon.TextWidth(myID.Caption) / Screen.TwipsPerPixelY) + 20
        Call CopyMemory(ByVal lParam, MeasureInfo, Len(MeasureInfo))
      End If
            
    Case WM_DRAWITEM '44
        
      Call CopyMemory(DrawInfo, ByVal lParam, LenB(DrawInfo)) 'get the UDT for the menuitem's appearance
      
      If ValidId(myID, DrawInfo.ItemID) Then
        IsSelected = ((DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED) 'selected ?
        IsDisabled = ((DrawInfo.itemState And ODS_DISABLED) = ODS_DISABLED) 'disabled ?
        IsChecked = ((DrawInfo.itemState And ODS_CHECKED) = ODS_CHECKED) 'checked ?
        r = DrawInfo.rcItem
              
        'Set background and forecolor appropriately
        If IsSelected And Not IsDisabled Then
          FillRect DrawInfo.hDC, r, GetSysColorBrush(COLOR_HIGHLIGHT) ' paint blue background for selection
          SetTextColor DrawInfo.hDC, GetSysColor(COLOR_HIGHLIGHTTEXT) ' write in a color that will be readable through blue background
        ElseIf IsDisabled Then
          FillRect DrawInfo.hDC, r, GetSysColorBrush(COLOR_MENU) ' paint gray background
          SetTextColor DrawInfo.hDC, vbWhite ' text white for disabled (we'll write this text again as gray, offset by one pixel to look 'disabled')
        Else
          FillRect DrawInfo.hDC, r, GetSysColorBrush(COLOR_MENU) ' paint gray background
          SetTextColor DrawInfo.hDC, GetSysColor(COLOR_MENUTEXT) ' normal text color
        End If
        
        SetBkMode DrawInfo.hDC, 1  ' write text transparent
  
        'Do the caption next
        OffsetRect r, 20, 0
        SetTextColor DrawInfo.hDC, GetSysColor(COLOR_MENUTEXT)
        If IsSelected Then SetTextColor DrawInfo.hDC, GetSysColor(COLOR_HIGHLIGHTTEXT)
        If IsDisabled Then SetTextColor DrawInfo.hDC, vbWhite
        DrawText DrawInfo.hDC, myID.Caption, Len(myID.Caption), r, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
        
        If IsDisabled Then
          'Do it again with gray and offset by one pixel
          SetTextColor DrawInfo.hDC, GetSysColor(COLOR_GRAYTEXT)
          OffsetRect r, -1, -1
          DrawText DrawInfo.hDC, myID.Caption, Len(myID.Caption), r, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
        End If
  
        If IsChecked Then
          SetRect chkR, 0, r.Top, 17, r.Top + 20
          DrawEdge DrawInfo.hDC, chkR, BDR_SUNKENOUTER, BF_RECT
          SetRect chkR, 1, r.Top + 1, 16, r.Top + 18
          FillRect DrawInfo.hDC, chkR, IIf(IsSelected, GetSysColorBrush(COLOR_HIGHLIGHT), GetSysColorBrush(COLOR_BTNHIGHLIGHT))
        End If
        
        If IsDisabled Then
          DrawState DrawInfo.hDC, 0, 0, frmIcon.IL1.ListImages(myID.PictureIndex).Picture.Handle, 0, 4, (r.Top + 2), 16, 16, DST_ICON Or DSS_DISABLED
        Else
          frmIcon.IL1.ListImages(myID.PictureIndex).Draw DrawInfo.hDC, 4, (r.Top + 2) * Screen.TwipsPerPixelY, 1
        End If
        
      End If
      
    End Select

End Function

Private Function ValidId(myID As mnuInfo, ID As Long) As Boolean
  
  'this may not be necessary, but i'm using it anyway
  'loop through our collection and find the id that matches
  'the id passed by the lparam of the handlemenumsgvb procedure
  'set myID to the saved mnuInfo that matches that ID and pass
  'it back to the caller for use in measuring and drawing
  
  'I have run into problems in passing items back to the caller
  'inside the parentheses when using 'Byval'. Especially with
  'boolean variables, it would return the wrong value.
  
  Dim ZX As Long, mTemp As mnuInfo
  
  For ZX = 1 To mnuIDs.Count
    Set mTemp = mnuIDs.Item(ZX)
    If mTemp.ItemID = ID Then
      Set myID = mnuIDs.Item(ZX)
      ValidId = True
      Exit Function
    End If
  Next
  
End Function

'don't know how well this works...
Private Function IsWindowNT() As Boolean
   IsWindowNT = Len(Environ$("OS"))
End Function
