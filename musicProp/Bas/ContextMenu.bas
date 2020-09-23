Attribute VB_Name = "ContextMenu"
Option Explicit


Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long

Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hbr As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Type MENUITEMINFO
  cbSize          As Long
  fMask           As Long
  fType           As Long
  fState          As Long
  wid             As Long
  hSubMenu        As Long
  hbmpChecked     As Long
  hbmpUnchecked   As Long
  dwItemData      As Long
  dwTypeData      As String
  cch             As Long
End Type

Public Type RECT
  Left            As Long
  Top             As Long
  Right           As Long
  Bottom          As Long
End Type

Public Type DRAWITEMSTRUCT
  CtlType         As Long
  CtlID           As Long
  ItemID          As Long
  itemAction      As Long
  itemState       As Long
  hWndItem        As Long
  hDC             As Long
  rcItem          As RECT
  ItemData        As Long
End Type

Public Type MEASUREITEMSTRUCT
  CtlType         As Long
  CtlID           As Long
  ItemID          As Long
  ItemWidth       As Long
  ItemHeight      As Long
  ItemData        As Long
End Type


Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_SUNKENOUTER = &H2

Public Const DT_LEFT = &H0
Public Const DT_RIGHT = &H2
Public Const DT_NOCLIP = &H100
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER = &H4

Public Const ODT_MENU = 1

Public Const ODS_SELECTED = &H1
Public Const ODS_DISABLED = &H4
Public Const ODS_CHECKED = &H8

Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_MENU = 4
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNHIGHLIGHT = 20

Public Const DSS_DISABLED = &H20
Public Const DST_ICON = &H3

Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20


Public Const MF_INSERT = &H0
Public Const MF_CHANGE = &H80
Public Const MF_APPEND = &H100
Public Const MF_DELETE = &H200
Public Const MF_REMOVE = &H1000
Public Const MF_BYCOMMAND = &H0
Public Const MF_BYPOSITION = &H400
Public Const MF_SEPARATOR = &H800
Public Const MF_ENABLED = &H0
Public Const MF_GRAYED = &H1
Public Const MF_DISABLED = &H2
Public Const MF_UNCHECKED = &H0
Public Const MF_CHECKED = &H8
Public Const MF_USECHECKBITMAPS = &H200
Public Const MF_STRING = &H0
Public Const MF_BITMAP = &H4
Public Const MF_OWNERDRAW = &H100
Public Const MF_POPUP = &H10
Public Const MF_MENUBARBREAK = &H20
Public Const MF_MENUBREAK = &H40
Public Const MF_UNHILITE = &H0
Public Const MF_HILITE = &H80
Public Const MF_DEFAULT = &H1000
Public Const MF_SYSMENU = &H2000
Public Const MF_HELP = &H4000
Public Const MF_RIGHTJUSTIFY = &H4000
Public Const MF_MOUSESELECT = &H8000
Public Const MF_END = &H80

Public Const WM_DRAWITEM = &H2B    '43
Public Const WM_MEASUREITEM = &H2C '44


Public Function QueryContextMenuVB(ByVal This As Object, ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As Long) As Long
  
  Dim ctxMenu As clsContextMenu
  Set ctxMenu = This
  
  QueryContextMenuVB = ctxMenu.QueryContextMenu(hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags)

  Set ctxMenu = Nothing

End Function



'If (uFlags And &HF) = CMF_NORMAL Then
'
'  'Implement this for Drag-and-Drop handler.
'
'ElseIf (uFlags And CMF_VERBSONLY) Then
'
'  'This is a context menu for a shortcut item.
'
'ElseIf (uFlags And CMF_EXPLORE) Then
'
'  'Right-click on file in Explorer.
'  'This is what we are interested in for our context
'    'menu.
'
'ElseIf (uFlags And CMF_DEFAULTONLY) Then
'
'    'Indicates a default action is being performed (typically a
'    'user is double-clicking on the file).
'
'End If
