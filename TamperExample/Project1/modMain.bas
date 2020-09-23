Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function MapFileAndCheckSum Lib "imagehlp.dll" Alias "MapFileAndCheckSumA" (ByVal FileName As String, ByRef HeaderSum As Long, ByRef CheckSum As Long) As Long

Public Sub Main()
'--------------------------------------------------------
' Remove this If/End If from the project when ready to compile
'--------------------------------------------------------
If InIDE Then
   MsgBox "Your in the IDE, Tamper Verification wont work on NON-COMPILED Projects.", vbInformation
   frmMain.Show
   Exit Sub
End If

'--------------------------------------------------------
' Check to see if program data has been modified
'--------------------------------------------------------
If VerifyExeFile(App.Path & Chr(92) & App.EXEName & Chr(46) & Chr(101) & Chr(120) & Chr(101)) = False Then
'--------------------------------------------------------
' Terminate Program Here.
' Don't Notify cracker of result (MsgBox etc) (((MSGBOX FOR TEST ONLY)))
' Just terminate program...
' If the user isn't tampering with your app then this won't happen.
'--------------------------------------------------------
   MsgBox "Program tampered", vbCritical
   End
 Else
'--------------------------------------------------------
' Start Program Here.
'--------------------------------------------------------
   frmMain.Show
End If

End Sub

Public Function InIDE() As Boolean
'------------------------------------------------------
' This function determines whether or not you're in development mode.
'------------------------------------------------------
On Error Resume Next
Debug.Print (1 / 0)

InIDE = CBool(Err.Number <> 0)
If Err.Number <> 0 Then
   Err.Clear
End If
On Error GoTo 0

End Function

'------------------------------------------------------
' Returns True if file is not changed since compilation.
' False if file is changed or if an error occured.
'------------------------------------------------------
Private Function VerifyExeFile(FileName As String) As Boolean
Dim HeaderSum As Long
Dim CheckSum As Long

If MapFileAndCheckSum(FileName, HeaderSum, CheckSum) = 0 Then
   VerifyExeFile = HeaderSum = CheckSum
End If
End Function

