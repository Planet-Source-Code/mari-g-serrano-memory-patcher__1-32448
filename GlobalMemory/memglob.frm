VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mario´s VB Memory Patcher"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   Icon            =   "memglob.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAddress 
      Height          =   345
      Left            =   1290
      TabIndex        =   10
      Text            =   "&H60000"
      Top             =   2970
      Width           =   1530
   End
   Begin VB.TextBox txt2Find 
      Height          =   345
      Left            =   1290
      TabIndex        =   6
      Text            =   "I AM A STRING"
      Top             =   2235
      Width           =   1530
   End
   Begin VB.ListBox lstApp 
      Height          =   2010
      Left            =   0
      TabIndex        =   4
      Top             =   195
      Width           =   5730
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   4215
      TabIndex        =   3
      Top             =   2865
      Width           =   1455
   End
   Begin VB.TextBox txt2Change 
      Height          =   345
      Left            =   1290
      TabIndex        =   2
      Text            =   "PATCHING TEST"
      Top             =   2610
      Width           =   1530
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Start Search"
      Height          =   495
      Left            =   4470
      TabIndex        =   0
      Top             =   2295
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Memory Locations:"
      Height          =   495
      Left            =   3435
      TabIndex        =   12
      Top             =   2880
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "Start at Address:"
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   11
      Top             =   3090
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Address:"
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   3585
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "Replace With..."
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   2730
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Text to Find:"
      Height          =   210
      Index           =   0
      Left            =   30
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Select an App..."
      Height          =   210
      Left            =   15
      TabIndex        =   5
      Top             =   0
      Width           =   1620
   End
   Begin VB.Label lblAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1650
      TabIndex        =   1
      Top             =   3615
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' MaRiØ Glez. Serrrano. 07-Marzo-2002
' code to acccess to the memory of any proccess and modify values
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteValue Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteString Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

'EXAMPLE OF USE:
'Open a notepad
' write "I AM A STRING" in notepad
' minimize the notepad
' run this proyect
' select "Untiled - Notepad" in the listbox
' insert "I AM A STRING" in the text to find textbox
' insert "PATCHING TEST" in the text to change textbox
' push Start the Search Button
' when the program terminates (im my PC at memory Address 493024),
' restore the notepad and ... tachan!!

Private Sub Form_Load()
Dim lhW() As Long, i As Long
Dim lCap() As String

lCap = GetWindows(lhW)
For i = 0 To UBound(lhW)
    lstApp.AddItem lCap(i)
    lstApp.ItemData(lstApp.NewIndex) = lhW(i)
Next
lstApp.ListIndex = 0
End Sub

Private Sub cmdFind_Click()
'proc to patch the memory..
    ' Declare some variables we need
    Static bStop As Boolean
    If cmdFind.Caption = "Stop" Then
        cmdFind.Caption = "Start Search"
        bStop = True
        Exit Sub
    Else
       cmdFind.Caption = "Stop"
       bStop = False
    End If
    cmdFind.Caption = "Stop"
    Dim hWnd As Long       ' Holds the handle returned by FindWindow
    Dim pid As Long        ' Used to hold the Process Id
    Dim pHandle As Long    ' Holds the Process Handle
    Dim Money As Long      ' Holds the value to write to memory
    Dim str As String      ' Holds the string for the Trainer Spy user

    ' get the handle to a window by its Caption
    ' hWnd = FindWindow(vbNullString, "Untiled - Notepad")
    hWnd = lstApp.ItemData(lstApp.ListIndex)
    If (hWnd = 0) Then
         MsgBox "Window not found!"
         Exit Sub
    End If

    ' Get the ProcId of the Window
    GetWindowThreadProcessId hWnd, pid

    ' use the pId to get a handle
    pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
    If (pHandle = 0) Then
         MsgBox "Unable to open process!"
         Exit Sub
    End If
    '..if its a long-> len=4
    Dim n As Long, h As Long, value As Long, i As Long, j As Long
    Dim Text As String
    If Trim(txtAddress) = "" Then txtAddress = "&H60000"
    h = CLng(txtAddress)
    If h = 0 Then h = &H60000
    On Error Resume Next
    Do
    
    lblAddress = h
       h = h + 1
       value = 0
       ReadProcessMemory pHandle, h, value, 1, 0&
       If value = Asc(Left(txt2Find, 1)) Then
          'first char found!!
           j = h
           For i = 2 To Len(txt2Find)
               j = j + 2 'jumps chr(0)
               value = 0
               ReadProcessMemory pHandle, j, value, 1, 0&
               If value <> Asc(Mid(txt2Find, i, 1)) Then Exit For
               If i = Len(txt2Find) Then Exit Do
               If bStop Then Exit Sub
           Next
       End If
       DoEvents
       If bStop Then Exit Sub
    Loop
    
    Text = Space$(Len(txt2Find))
    Text = txt2Change
    List1.AddItem h
    lblAddress = "MODIFIED OK"
    cmdFind.Caption = "Start Search"

    'to write numeric values you can ..(Must) use WriteValue API
    WriteString pHandle, h, StrPtr(Text), LenB(Text), 0&
    CloseHandle pHandle
    
End Sub


Private Sub List1_Click()
    txtAddress.Text = List1.Text
End Sub
