VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNSTEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INameSpaceTreeControlEvents

Public Sub INameSpaceTreeControlEvents_OnItemClick(ByVal psi As IShellItem, ByVal nstceHitTest As NSTCEHITTEST, ByVal nstceClickType As NSTCECLICKTYPE)
'MSDN does not make it entirely clear; not returning a non-zero
'results in not being able to click the arrows to expand a node
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_NORMALDISPLAY, lpsz

frmNST.List1.AddItem "OnItemClick " & LPWSTRtoStr(lpsz), 0
Err.ReturnHResult = 1
End Sub
Public Sub INameSpaceTreeControlEvents_OnPropertyItemCommit(ByVal psi As IShellItem)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_NORMALDISPLAY, lpsz

frmNST.List1.AddItem "OnPropertyItemCommit " & LPWSTRtoStr(lpsz), 0

End Sub
Public Sub INameSpaceTreeControlEvents_OnItemStateChanging(ByVal psi As IShellItem, ByVal nstcisMask As NSTCITEMSTATE, ByVal nstcisState As NSTCITEMSTATE)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_NORMALDISPLAY, lpsz

frmNST.List1.AddItem "OnItemStateChanging " & LPWSTRtoStr(lpsz), 0

End Sub
Public Sub INameSpaceTreeControlEvents_OnItemStateChanged(ByVal psi As IShellItem, ByVal nstcisMask As NSTCITEMSTATE, ByVal nstcisState As NSTCITEMSTATE)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_NORMALDISPLAY, lpsz

frmNST.List1.AddItem "OnItemStateChanged " & LPWSTRtoStr(lpsz), 0

End Sub
Public Sub INameSpaceTreeControlEvents_OnSelectionChanged(ByVal psiaSelection As IShellItemArray)


Dim lpName As LongPtr, sName As String
Dim iesi As IEnumShellItems
Dim psi As IShellItem
psiaSelection.EnumItems iesi

Do While (iesi.Next(1, psi, 0) = NOERROR)

    psi.GetDisplayName SIGDN_FILESYSPATH, lpName
    sName = sName & LPWSTRtoStr(lpName) & ", "
    Set psi = Nothing
Loop

frmNST.List1.AddItem "OnSelectionChanged " & sName, 0
Set iesi = Nothing

End Sub
Public Sub INameSpaceTreeControlEvents_OnKeyboardInput(ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr)
frmNST.List1.AddItem "OnKeyboardInput"
End Sub
Public Sub INameSpaceTreeControlEvents_OnBeforeExpand(ByVal psi As IShellItem)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_FILESYSPATH, lpsz

frmNST.List1.AddItem "OnBeforeExpand " & LPWSTRtoStr(lpsz), 0

End Sub
Public Sub INameSpaceTreeControlEvents_OnAfterExpand(ByVal psi As IShellItem)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_FILESYSPATH, lpsz
frmNST.List1.AddItem "OnAfterExpand " & LPWSTRtoStr(lpsz), 0
End Sub
Public Sub INameSpaceTreeControlEvents_OnBeginLabelEdit(ByVal psi As IShellItem)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_FILESYSPATH, lpsz
frmNST.List1.AddItem "OnBeginLabelEdit " & LPWSTRtoStr(lpsz), 0

End Sub
Public Sub INameSpaceTreeControlEvents_OnEndLabelEdit(ByVal psi As IShellItem)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_FILESYSPATH, lpsz
frmNST.List1.AddItem "OnEndLabelEdit " & LPWSTRtoStr(lpsz), 0

End Sub
Public Sub INameSpaceTreeControlEvents_OnGetToolTip(ByVal psi As IShellItem, pszTip As LongPtr, ByVal cchTip As Long)
Err.ReturnHResult = E_NOTIMPL
End Sub
Public Sub INameSpaceTreeControlEvents_OnBeforeItemDelete(ByVal psi As IShellItem)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_FILESYSPATH, lpsz
frmNST.List1.AddItem "OnBeforeItemDelete " & LPWSTRtoStr(lpsz), 0

End Sub
Public Sub INameSpaceTreeControlEvents_OnItemAdded(ByVal psi As IShellItem, ByVal fIsRoot As BOOL)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_FILESYSPATH, lpsz
frmNST.List1.AddItem "OnItemAdded " & LPWSTRtoStr(lpsz), 0

End Sub
Public Sub INameSpaceTreeControlEvents_OnItemDeleted(ByVal psi As IShellItem, ByVal fIsRoot As BOOL)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_FILESYSPATH, lpsz
frmNST.List1.AddItem "OnItemDeleted " & LPWSTRtoStr(lpsz), 0

End Sub
Public Sub INameSpaceTreeControlEvents_OnBeforeContextMenu(ByVal psi As IShellItem, riid As UUID, ppv As LongPtr)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_FILESYSPATH, lpsz
frmNST.List1.AddItem "OnBeforeContextMenu " & LPWSTRtoStr(lpsz), 0
Err.ReturnHResult = E_NOTIMPL

End Sub
Public Sub INameSpaceTreeControlEvents_OnAfterContextMenu(ByVal psi As IShellItem, ByVal pcmIn As IContextMenu, riid As UUID, ppv As LongPtr)
Dim lpsz As LongPtr
psi.GetDisplayName SIGDN_FILESYSPATH, lpsz
frmNST.List1.AddItem "OnAfterContextMenu " & LPWSTRtoStr(lpsz), 0
Err.ReturnHResult = E_NOTIMPL
End Sub
Public Sub INameSpaceTreeControlEvents_OnBeforeStateImageChange(ByVal psi As IShellItem)

End Sub
Public Sub INameSpaceTreeControlEvents_OnGetDefaultIconIndex(ByVal psi As IShellItem, piDefaultIcon As Long, piOpenIcon As Long)
Err.ReturnHResult = E_NOTIMPL
'NOTE: This function is completely undocumented, there's not even an entry on MSDN
'This function must return -1, or you must look up the icon index yourself.
'Returning -1 is impossible without a v-table swap, which this project shows
'If you did not want to do that, delete the SwapVTable entry for this function
'and use the code below to get the correct icon index (also do a proper lookup for open)

'Dim ppidl As Long
'Dim pUnk As oleexp.IUnknown
'Set pUnk = psi
'Call SHGetIDListFromObject(ObjPtr(pUnk), ppidl)
'If ppidl Then
'    piDefaultIcon = GetFileIconIndexPIDL(ppidl, SHGFI_SMALLICON)
'    piOpenIcon = piDefaultIcon
'End If
'Call CoTaskMemFree(ppidl)
End Sub
