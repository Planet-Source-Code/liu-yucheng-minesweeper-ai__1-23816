VERSION 5.00
Begin VB.Form main2 
   Caption         =   "Main"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Allow use of Best Probabilty Approach (not guaranteed correct"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox Status 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   2715
   End
   Begin VB.Label Label2 
      Caption         =   $"main.frx":0000
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "main2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private DesktopDC As Long
Private ResultYCoors(10) As Long
Private Terminate As Boolean

Private Sub Command2_Click()
Terminate = False
Timer1.Interval = 100
Status.Text = "Awaiting Transfer to game window..."
End Sub


Private Sub Form_Load()
InitColours
End Sub


Private Sub Form_Terminate()
DeleteDC DesktopDC
End Sub


Private Sub Timer1_Timer()
'Do
'DoEvents
'Loop Until GetActiveWindow <> Form1.hwnd
Dim Wnd As Long
'Wnd = GetActiveWindow
Wnd = GetForegroundWindow
'Do
'wnd2 = GetParent(Wnd)
'If wnd2 = GetDesktopWindow Or wnd2 = 0 Then
'    Exit Do
'End If
'Wnd = wnd2
'Loop
s$ = GetText(Wnd)
'List1.AddItem s$
If InStr(1, s$, "Minesweeper") > 0 Then
Timer1.Interval = 0
ret& = GetWindowRect(Wnd, WInRECT)
Status.Text = "Window Captured"
Sleep 500
StartMainLoop
'End
End If
End Sub

Public Sub StartMainLoop()
Dim x As Long, y As Long
Dim TempStr As String
CurWindow = GetForegroundWindow
DesktopDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
Do
    UnClickCount = GetDisplay
    If Terminate = True Then
        Status.Text = "Sorry."
        Exit Do
    End If
    If UnClickCount = 0 Then
        Status.Text = "Game Solved"
        Exit Do
    End If
    DoEvents
    CalculateAI Check1.Value
  '  Sleep 500
 '   List1.Clear
Loop
DeleteDC DesktopDC
End Sub

Public Sub LeftClickTranspose(x As Long, y As Long)
LeftClick x + WInRECT.Left, y + WInRECT.Top

End Sub

Private Function GetDisplay() As Integer
Dim R As Long
Dim UnClickCount As Integer
For x = 1 To 30
    For y = 1 To 16
        LocalGrid(x, y) = GetPixelTranspose2(DesktopDC, CoorXbegin + (x - 1) * 16 + 2, CoorYBegin + (y - 1) * 16 + 2, 10)
        If LocalGrid(x, y) = 0 Then
            If GetPixelTranspose(DesktopDC, CoorXbegin + (x) * 16 - 1, CoorYBegin + (y - 1) * 16 + 2) = RGB(192, 192, 192) Then
                LocalGrid(x, y) = Blank
            End If
        End If
        If LocalGrid(x, y) = UnClicked Then UnClickCount = UnClickCount + 1
    Next y
Next x
GetDisplay = UnClickCount
End Function
Public Function GetPixelTranspose(DC As Long, x As Long, y As Long) As Long
'DeleteDC DesktopDC
'DesktopDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
'GetPixelTranspose = GetPixel(GetDC(GetDesktopWindow), X + WInRECT.Left, Y + WInRECT.Top)
GetPixelTranspose = GetPixel(DesktopDC, x + WInRECT.Left, y + WInRECT.Top)
'GetPixelTranspose = GetPixel(DC, X, Y)
End Function

Public Function GetPixelTranspose2(DC As Long, XBegin As Long, YBegin As Long, XYRange As Long) As Long
Dim A As Long, Possible As Long
Dim x As Long, y As Long, Z As Long
For x = XBegin To XBegin + XYRange
    For y = YBegin To YBegin + XYRange
        A = GetPixel(DesktopDC, x + WInRECT.Left, y + WInRECT.Top)
        If x = XBegin And y = YBegin And A = RGB(255, 0, 0) Then
            Terminate = True
        End If
        If A = MineColours(10) Then
            GetPixelTranspose2 = 10
            Exit Function
        End If
        If Possible = False Then
            For Z = 1 To 8
                If A = MineColours(Z) Then
                    Possible = Z
                    Exit For
                End If
            Next Z
            If Not (Possible = 0 Or Possible = 3 Or Possible = 5 Or Possible = 8) Then
                GetPixelTranspose2 = Possible
                Exit Function
            End If
        End If
    Next y
Next x
GetPixelTranspose2 = Possible
End Function

