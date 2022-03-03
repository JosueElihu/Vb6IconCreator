Attribute VB_Name = "mMnuImage"
Option Explicit

Private Type MENUINFO
  cbSize          As Long
  fMask           As Long
  dwStyle         As Long
  cyMax           As Long
  RhbrBack        As Long
  dwContextHelpID As Long
  dwMenuData      As Long
End Type

Private Type MENUITEMINFO
  cbSize          As Long
  fMask           As Long
  fType           As Long
  fState          As Long
  wID             As Long
  hSubMenu        As Long
  hbmpChecked     As Long
  hbmpUnchecked   As Long
  dwItemData      As Long
  dwTypeData      As Long
  cch             As Long
  hbmpItem        As Long
End Type

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, ByRef lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuInfo Lib "user32" (ByVal hMenu As Long, ByRef LPMENUINFO As MENUINFO) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, ByRef LPCMENUINFO As MENUINFO) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Private Const MIIM_STATE            As Long = &H1
Private Const MIIM_ID               As Long = &H2
Private Const MIIM_SUBMENU          As Long = &H4
Private Const MIIM_CHECKMARKS       As Long = &H8
Private Const MIIM_TYPE             As Long = &H10
Private Const MIIM_DATA             As Long = &H20
Private Const MIIM_BITMAP           As Long = &H80
Private Const MIM_STYLE             As Long = &H10

Private Const ODT_MENU              As Long = 1
Private Const ODS_GRAYED            As Long = &H2
Private Const ODS_CHECKED           As Long = &H8
Private Const MNS_NOCHECK           As Long = &H80000000

Public Function PutIconToVBMenu(hWnd As Long, hBitmap As Long, ByVal MenuPos As Long, ParamArray vSubMenuPos() As Variant) As Boolean
Dim hMenu    As Long
Dim hSubMenu As Long
Dim MII      As MENUITEMINFO
Dim Elmnt    As Variant

    
    hMenu = GetMenu(hWnd)
    hSubMenu = hMenu
    
    For Each Elmnt In vSubMenuPos
        hSubMenu = GetSubMenu(hSubMenu, Elmnt)
    Next
    
    With MII
        .cbSize = Len(MII)
        .fMask = MIIM_BITMAP
        .hbmpItem = hBitmap
    End With
    
    PutIconToVBMenu = SetMenuItemInfo(hSubMenu, MenuPos, True, MII)
    If hSubMenu = hMenu Then DrawMenuBar hWnd
    
End Function


Public Sub RemoveMenuCheckVB(hWnd As Long, ParamArray vSubMenuPos() As Variant)
Dim MI      As MENUINFO
Dim hMenu   As Long
Dim hSubMenu As Long
Dim Elmnt   As Variant

    hMenu = GetMenu(hWnd)
    hSubMenu = hMenu
    
    For Each Elmnt In vSubMenuPos
        hSubMenu = GetSubMenu(hSubMenu, Elmnt)
    Next
    
    With MI
        .cbSize = Len(MI)
        .fMask = MIM_STYLE
        .dwStyle = MNS_NOCHECK
    End With

    SetMenuInfo hSubMenu, MI
End Sub

Public Function PutImageToApiMenu(hBitmap As Long, ByVal hMenu As Long, ByVal MenuPos As Long, Optional ByVal ItemData As Long) As Boolean
Dim MII As MENUITEMINFO

    With MII
        .fMask = MIIM_BITMAP Or MIIM_DATA
        .hbmpItem = hBitmap
        .dwItemData = ItemData
    End With
    PutImageToApiMenu = SetMenuItemInfo(hMenu, MenuPos, True, MII)

End Function

