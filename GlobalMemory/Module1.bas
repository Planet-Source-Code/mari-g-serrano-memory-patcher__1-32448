Attribute VB_Name = "modFindWindows"
Option Explicit
'Find Windows..
  
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Const GWL_HWNDPARENT = (-8)
Private Const GW_HWNDNEXT = 2
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80

Private Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Function IsToolWindow(ByVal hWnd As Long) As Boolean
   ' is toolwindow??
   IsToolWindow = ((GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW)
End Function


Public Function GetWindows(ByRef hWnds() As Long) As String()
Dim hWnd As Long, hWndOwner As Long, i As Long
Dim txt As String * 255
Dim t2 As String
Dim arWnd() As String
ReDim arWnd(0) As String
hWnd = GetTopWindow(0)
Do While hWnd
   If IsWindowVisible(hWnd) Then 'Is Visible??
      If GetParent(hWnd) = 0 Then ' It not Child?
         hWndOwner = GetWindowLong(hWnd, GWL_HWNDPARENT)
         If (hWndOwner = 0) Or IsToolWindow(hWndOwner) Then
            If Not IsToolWindow(hWnd) Then
               ' Isn´t a ToolWindow
               ' we have a window!
               ' find the next!
               ReDim Preserve arWnd(i)
               'stores windows handles
               ReDim Preserve hWnds(i)
               hWnds(i) = hWnd
               'stores windows  captions
               GetWindowText hWnd, txt, 255
               arWnd(i) = txt
               i = i + 1
            End If
         End If
      End If
   End If
   hWnd = GetNextWindow(hWnd, GW_HWNDNEXT)
Loop
GetWindows = arWnd
End Function


