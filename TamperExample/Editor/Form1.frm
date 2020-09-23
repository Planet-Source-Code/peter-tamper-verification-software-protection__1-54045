VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Project Editor Utility"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Execute Project1.exe"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   310
      Left            =   6720
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   6495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Execute Project2.exe"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modify Project1.exe"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "3. Run the original Project1.exe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "1. Load the Path to Project1.exe..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "4. Run the modified Project2.exe and see result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "2. Modify the Project1.exe, add some extra data and rename it Project2.exe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Integer
Dim sText As String
Dim sPath1 As String
Dim sPath2 As String

sPath1 = txtPath.Text & "\Project1.exe"
sPath2 = txtPath.Text & "\Project2.exe"

'If Project2 already exists then delete it for re-creation check
If FileExists(sPath2) Then
   On Error Resume Next
   Kill sPath2
End If

On Error GoTo ErrCtrl

i = FreeFile

If Not FileExists(sPath1) Then
   MsgBox sPath1 & " not found," & vbCrLf & vbCrLf & _
          "Make sure you have compiled the project to the above directory!", _
          vbInformation
   Exit Sub
End If

'Read the app data
Open sPath1 For Binary Access Read As #i
     sText = Space$(LOF(i))
     Get i, , sText
Close #i

i = FreeFile

'Write the app data, add some extra data (Tamper with it a little)
Open sPath2 For Binary Access Write As #i
     Put #i, , sText
     Put #1, , "Additional data"
Close #i

If FileExists(sPath2) Then
   MsgBox sPath2 & " created," & vbCrLf & vbCrLf & _
          "Run Project2.exe, a modified version of Project1.exe and see the result!", _
          vbInformation
          Command2.Enabled = True
          Command4.Enabled = True
   Exit Sub
End If

Exit Sub
ErrCtrl:
   MsgBox Err.Description & " " & Err.Number
End Sub

Private Sub Command2_Click()
ShellExecute Me.hWnd, "Open", txtPath.Text & "\Project2.exe", vbNullString, App.path, vbNormalFocus
End Sub

Private Sub Command3_Click()
Dim TmpPath As String
Dim path As String

Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False

TmpPath = Label3
If Len(TmpPath) > 0 Then
    If Not Right$(TmpPath, 1) <> "\" Then TmpPath = Left$(TmpPath, Len(TmpPath) - 1) ' Remove "\" if the user added
End If
TmpPath = BrowseForFolder(TmpPath) ' Browse for folder
If TmpPath = "" Then
    path = "" ' If the user pressed cancel
Else
    path = TmpPath ' If the user selected a folder
    Command1.Enabled = True
End If
txtPath = path
End Sub

Private Sub Command4_Click()
ShellExecute Me.hWnd, "Open", txtPath.Text & "\Project1.exe", vbNullString, App.path, vbNormalFocus
End Sub
