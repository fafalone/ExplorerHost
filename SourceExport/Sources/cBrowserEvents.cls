VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowserEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IExplorerBrowserEvents

Public Sub IExplorerBrowserEvents_OnNavigationPending(ByVal pidlFolder As LongPtr)
frmBasic.List1.AddItem "OnNavigationPending", 0
End Sub
Public Sub IExplorerBrowserEvents_OnViewCreated(ByVal psv As IShellView)
frmBasic.List1.AddItem "OnViewCreated", 0
End Sub
Public Sub IExplorerBrowserEvents_OnNavigationComplete(ByVal pidlFolder As LongPtr)
frmBasic.List1.AddItem "OnNavigationComplete " & GetPathFromPIDLW(pidlFolder), 0
End Sub
Public Sub IExplorerBrowserEvents_OnNavigationFailed(ByVal pidlFolder As LongPtr)
frmBasic.List1.AddItem "OnNavigationFailed" & GetPathFromPIDLW(pidlFolder), 0
End Sub
Private Function GetPathFromPIDLW(pidl As LongPtr) As String
  Dim pszPath As String
  pszPath = String(MAX_PATH, 0)
  If SHGetPathFromIDListW(pidl, StrPtr(pszPath)) Then
    If InStr(pszPath, vbNullChar) Then
        GetPathFromPIDLW = Left$(pszPath, InStr(pszPath, vbNullChar) - 1)
    End If
  End If
End Function