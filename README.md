# ExplorerHost
Host an instance of Explorer itself with INamespaceTreeControl and IExplorerBrowser

This is a [twinBASIC](https://github.com/twinbasic/twinbasic) version of the following original VB6 project:

[[VB6, Vista+] Host Windows Explorer on your form: navigation tree and/or folder](https://www.vbforums.com/showthread.php?798633-VB6-Vista-Host-Windows-Explorer-on-your-form-navigation-tree-and-or-folder)

**Update (19 Dec 2023):** .twinproj updated to reference WinDevLib (formerly tbShellLib) 7.0-- this eliminates package errors that tB did not raise at the time this project was initially released.

**Update:** On 08 May 2023 the source was updated just to change the dependency version: tbShellLib had Optional UDTs, which are not allowed. Previously, nothing would happen unless you actually tried to use one without supplying it, but in more recent versions, this is flagged as an error before compiling, preventing that from happening. If you get this error in other projects using tbShellLib, or don't want to re-DL the source here, just update tbShellLib to the newest version. 

IExplorerBrowser is an easy to use, more complete version of IShellView (in fact, it has an IShellView at its core that you can access) that lets you have a complete Explorer frame on your form, with very little code. You can either have just a plain file view, or with a navigation tree and toolbar. It uses all the same settings and does all the same things as Explorer, and your program can interact with those actions to do things like browse for files, or be the basis of a namespace extension.
The only complication is that there's no event notifying of an individual file selection within the view, and getting a list of selected files is fairly complex- however there is a function to do it in the demo project.

If all you want is the navigation tree, you have the INamespaceTreeControl. It's got a decent amount of options for however you want to display things, including checkboxes. There is a wide range of events that you're notified of via the event sink, and most of these use IShellItem- the demo project does show to to convert that into a path, but it's a very useful interface to learn if you're going to be doing shell programming. The selection is reported through IShellItemArray, which is slightly easier than IDataObject.

It's got one little quirk though... you have the option to set the folder icons yourself, but if you don't want to do that and just use the default icon that you see in Explorer, you have to return -1. The demo project shows how to go both ways, no thanks to MSDN and their complete lack of documentation of this.
But this is by far the easiest to create way of having a full-featured Explorer-like navigation- I've made a regular TreeView into this, and it took hundreds of lines and heavy subclassing. This is a simple object. (Note that it does support some advanced features through related interfaces, like custom draw, drop handling, and accessibility... these interfaces are included in oleexp, but have not been brought to the sample project here, perhaps in the future I'll do a more in-depth one if there's any interest)

---

twinBASIC makes things a bit easier on us. The direct control over the hresult in implemented interfaces with `Err.ReturnHResult` means we don't need to worry about v-table swapping (although the method still works). In addition to removing those, [tbShellLib](https://github.com/fafalone/tbShellLib) replaces oleexp, and has enough API coverage now we didn't need any local declares. The code has also been updated to fully support 64bit mode, and the INamespaceTreeControl is now resizable, which demonstrates how to query it's hWnd through `IOleWindow`. 

![image](https://user-images.githubusercontent.com/7834493/226088561-45767132-0abd-4763-a632-b2f7cb9c1d19.png)

![image](https://user-images.githubusercontent.com/7834493/226088596-fafda860-6834-4f30-b718-2d4f7ac413fc.png)

The whole thing is very simple to set up; creating an ExplorerBrowser goes like this:

```vb6
    Private pEBrowse As ExplorerBrowser
    Private cSink As New cBrowserEvents
    
      Dim prc As RECT
    Dim pfs As FOLDERSETTINGS
    Dim sPath As String
    Dim lFlag As EXPLORER_BROWSER_OPTIONS
    If Form1.Check1.Value = vbChecked Then
        lFlag = EBO_SHOWFRAMES
    Else
        lFlag = EBO_NONE
    End If
    sPath = "C:\"
    pidlt = ILCreateFromPathW(StrPtr(sPath))
    
    pfs.fFlags = FWF_ALIGNLEFT
    pfs.ViewMode = FVM_DETAILS
    prc.Top = 0
    prc.Bottom = (Me.Height / 15) - 97
    prc.Left = 0
    prc.Right = (Me.Width / 15) - 16
    
    Set pEBrowse = New ExplorerBrowser
    pEBrowse.Initialize Me.hWnd, prc, pfs
    pEBrowse.SetOptions lFlag
    pEBrowse.Advise cSink, lpck
    pEBrowse.BrowseToIDList pidlt, SBSP_ABSOLUTE 'if you want to use the desktop, for its own sake, or want a quick substitute
                                                 'for not having C:\ , instead of pidlt use VarPtr(0) -NOT just 0.
```


While you can do a bit of customization with these controls, it's a bit limited. If you want a high level of customization and additional features, you can check out [my ShellControls project](https://github.com/fafalone/ShellControls), the twinBASIC port of my ucShellTree and ucShellBrowse controls.
    
