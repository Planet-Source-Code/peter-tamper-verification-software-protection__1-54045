Attribute VB_Name = "modFile"
Option Explicit

Public Type BROWSEINFOTYPE
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BROWSEINFOTYPE) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal PV As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                            (ByVal hWnd As Long, ByVal lpOperation As String, _
                            ByVal lpFile As String, ByVal lpParameters As String, _
                            ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Const WM_USER = &H400

Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Public Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Const LPTR = (&H0 Or &H40)

Public Function BrowseForFolder(selectedPath As String) As String
Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim TmpPath As String * 256
With Browse_for_folder
    .hOwner = Form1.hWnd
    .lpszTitle = "Browse for folders"
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr)
    selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1)
    CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1
    .lParam = selectedPathPointer
End With
itemID = SHBrowseForFolder(Browse_for_folder)
If itemID Then
    If SHGetPathFromIDList(itemID, TmpPath) Then
        BrowseForFolder = Left$(TmpPath, InStr(TmpPath, vbNullChar) - 1)
    End If
    Call CoTaskMemFree(itemID)
End If
Call LocalFree(selectedPathPointer)
End Function

Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
If uMsg = 1 Then
    Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
End If
End Function

Public Function FunctionPointer(FunctionAddress As Long) As Long
FunctionPointer = FunctionAddress
End Function

'--------------------------------------------------------
' Check to see if a file exists
'--------------------------------------------------------
Public Function FileExists( _
              ByVal FileName As String) As Boolean
Dim Value As Boolean

On Error Resume Next
Value = CBool(Len(Dir$(FileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly)) > 0)

If Err.Number <> 0 Then
   Err.Clear
   Value = False
End If

FileExists = Value

On Error GoTo 0
End Function
