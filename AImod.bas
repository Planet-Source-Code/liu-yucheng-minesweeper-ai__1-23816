Attribute VB_Name = "AImod"
Public Const Correct = 1
Public Const WrongPos = 2
Public Enum EnumColor
    Red = 0
    White = 1
    Yellow = 2
    Brown = 3
    Purple = 4
    Green = 5
End Enum

Public Type DEVMODE
        dmDeviceName As String * 32
        dmSpecVersion As Long
        dmDriverVersion As Long
        dmSize As Long
        dmDriverExtra As Long
        dmFields As Long
        dmOrientation As Long
        dmPaperSize As Long
        dmPaperLength As Long
        dmPaperWidth As Long
        dmScale As Long
        dmCopies As Long
        dmDefaultSource As Long
        dmPrintQuality As Long
        dmColor As Long
        dmDuplex As Long
        dmYResolution As Long
        dmTTOption As Long
        dmCollate As Long
        dmFormName As String * 32
        dmUnusedPadding As Long
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const mouse_eventC = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 ' right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10 ' right button up

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal _
       lParam As Long)

Public Const WM_SETTEXT = &HC

Public Declare Function SendMessageByString Lib "user32" Alias _
       "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
       wParam As Long, ByVal lParam As String) As Long


Public Declare Function GetCursorPos Lib "user32" (lpPoint As Where) As Long


Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, _
       ByVal yPoint As Long) As Long


Public Type Where
       Pointa As Long
       Pointb As Long
       End Type
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3


Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6

Public Const UnClicked = 0
Public Const Mine = 10
Public Const Blank = 9
Public LocalGrid(31, 17) As Long
Public Const CoorXbegin = 15
Public Const CoorYBegin = 96
Public Const Size = 16
Public MineColours(10) As Long
Public WInRECT As RECT
Public Type Coor
    x As Integer
    y As Integer
End Type
Sub LeftClick(xP As Long, yP As Long)
Dim junk As Long
'GetCurPoint
    junk = SetCursorPos(xP, yP)
    Sleep 7
    mouse_event MOUSEEVENTF_LEFTDOWN, xP, yP, 0, 0
    Sleep 7
    mouse_event MOUSEEVENTF_LEFTUP, xP, yP, 0, 0
'SetCurPoint
End Sub


   Function GetText(hnd)


       'this will get the text of the "hnd" window
       GetTrim = SendMessageByNum(hnd, 14, 0&, 0&)
       TrimSpace$ = Space$(GetTrim)
       GetString = SendMessageByString(hnd, 13, GetTrim + 1, TrimSpace$)
       GetText = TrimSpace$
   End Function





Public Sub InitColours()
MineColours(1) = RGB(0, 0, 255)
MineColours(2) = RGB(0, 128, 0)
MineColours(3) = RGB(255, 0, 0)
MineColours(4) = RGB(0, 0, 128)
MineColours(5) = RGB(128, 0, 0)
MineColours(6) = RGB(0, 128, 128)
MineColours(7) = RGB(128, 0, 128)
MineColours(8) = RGB(0, 0, 0)
MineColours(10) = RGB(128, 128, 128)
End Sub

Public Sub RightClick(xP As Long, yP As Long)
Dim junk As Long
'GetCurPoint
    junk = SetCursorPos(xP, yP)
    Sleep 7
    mouse_event MOUSEEVENTF_RIGHTDOWN, xP, yP, 0, 0
    Sleep 7
    mouse_event MOUSEEVENTF_RIGHTUP, xP, yP, 0, 0
'SetCurPoint
End Sub

Public Sub BothClick(xP As Long, yP As Long)
Dim junk As Long
'GetCurPoint
    junk = SetCursorPos(xP, yP)
    Sleep 7
    mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_LEFTDOWN, xP, yP, 0, 0
    Sleep 7
    mouse_event MOUSEEVENTF_RIGHTUP Or MOUSEEVENTF_LEFTUP, xP, yP, 0, 0
'SetCurPoint

End Sub

Public Sub CalculateAI(AllowBest As Integer)
'TYPE 1:
Dim x As Long, y As Long, TypeExecuted As Boolean
Dim XX As Long, YY As Long
Dim TempCount As Integer
main2.Status.Text = "Using Approach Type 1"
DoEvents
'number of mines full. bothclick button
For x = 1 To 30
    For y = 1 To 16
        If LocalGrid(x, y) >= 1 And LocalGrid(x, y) <= 8 Then
            If CountSurrounding(x, y, Mine) = LocalGrid(x, y) Then
                If CountSurrounding(x, y, UnClicked) > 0 Then
                    XX = x
                    YY = y
                    TransposeXY XX, YY
                    BothClick XX, YY
                    TypeExecuted = True
                End If
            End If
        End If
    Next y
Next x
If TypeExecuted = True Then Exit Sub
'TYPE 2:
'number of blanks=number of mines
'rightclick all
main2.Status.Text = "Using Approach Type 2"
DoEvents
For x = 1 To 30
    For y = 1 To 16
        If LocalGrid(x, y) >= 1 And LocalGrid(x, y) <= 8 Then
            TempCount = CountSurrounding(x, y, UnClicked)
            If CountSurrounding(x, y, Mine) + TempCount = LocalGrid(x, y) And TempCount > 0 Then
                RightClickSurrounding x, y, UnClicked
                TypeExecuted = True
            End If
        End If
    Next y
Next x
If TypeExecuted = True Then Exit Sub
'Type 3:
'Probabilistic approach
'Create new probability Array
Dim FinalProbCount(31, 17) As Long
Dim FinalProbUsed(31, 17) As Long
main2.Status.Text = "Accurate Probabilistic Approach"
DoEvents
Dim MinestoAssign As Long
Dim BlankstoUse As Long
Dim fail As Boolean
Dim SurroundTable(8) As Coor
Dim ProbCount(31, 17) As Long
Dim ProbUsed(31, 17) As Long
Dim Z As Long, count As Long, BitCount As Integer
For x = 1 To 30
    For y = 1 To 16
    'ResetProbability Tables
    For XX = 1 To 31
        For YY = 1 To 17
        ProbCount(XX, YY) = 0
        ProbUsed(XX, YY) = 0
        Next YY
    Next XX
1
        If LocalGrid(x, y) >= 1 And LocalGrid(x, y) <= 8 Then
            BlankstoUse = CountSurrounding2(x, y, UnClicked, SurroundTable)
            MinestoAssign = LocalGrid(x, y) - CountSurrounding(x, y, Mine)
            If BlankstoUse > 0 And MinestoAssign > 0 Then
            If MinestoAssign < 0 Then
                MsgBox "Error has occured"
            End If
            Z = 2 ^ BlankstoUse - 1
            For count = 1 To Z
                BitCount = 0
                For bits = 0 To BlankstoUse - 1
                    If count And (2 ^ bits) Then BitCount = BitCount + 1
                Next bits
                If BitCount = MinestoAssign Then
                    For bits = 0 To BlankstoUse - 1
                        If count And (2 ^ bits) Then
                            With SurroundTable(bits + 1)
                                LocalGrid(.x, .y) = Mine
                            End With
                        End If
                    Next bits
                    fail = False
                    If MineCount > 99 Then fail = True
                    For XX = x - 2 To x + 2
                        For YY = y - 2 To y + 2
                        If XX >= 1 And XX <= 30 And YY >= 1 And YY <= 16 Then
                        If LocalGrid(XX, YY) >= 1 And LocalGrid(XX, YY) <= 8 Then
                        If CountSurrounding(XX, YY, Mine) > LocalGrid(XX, YY) Then
                            fail = True
                        End If
                        End If
                        End If
                        Next YY
                    Next XX
                    For bits = 0 To BlankstoUse - 1
                        If count And (2 ^ bits) Then
                            With SurroundTable(bits + 1)
                                LocalGrid(.x, .y) = UnClicked
                            End With
                        End If
                    Next bits
                    If fail = False Then
                    ProbUsed(x, y) = ProbUsed(x, y) + 1
                    FinalProbUsed(x, y) = FinalProbUsed(x, y) + 1
                    
                    For bits = 0 To BlankstoUse - 1
                        If count And (2 ^ bits) Then
                            With SurroundTable(bits + 1)
                                ProbCount(.x, .y) = ProbCount(.x, .y) + 1
                                FinalProbCount(.x, .y) = FinalProbCount(.x, .y) + 1
                            End With
                        End If
                    Next bits
                    End If
                End If
            Next count
        'Type 3.1 :
        'A particular square has all the possibilities. without failing
        For Z = 1 To BlankstoUse
            With SurroundTable(Z)
            If ProbCount(.x, .y) = ProbUsed(x, y) Then
            If ProbCount(.x, .y) <> 0 Then
                XX = .x
                YY = .y
                TransposeXY XX, YY
                RightClick XX, YY
                TypeExecuted = True
            'critical modification : refresh data
                Exit Sub
            End If
            End If
            End With
        Next Z
        End If
    End If
    Next y
Next x
If AllowBest > 0 Then
main2.Status.Text = "Best Case Probabilistic Approach"
DoEvents
Dim MaxX As Integer, MaxY As Integer
MaxX = 1
MaxY = 1
For x = 1 To 30
    For y = 1 To 16
       If FinalProbCount(x, y) > FinalProbCount(MaxX, MaxY) Then
        MaxX = x
        MaxY = y
        End If
    Next y
Next x
XX = MaxX
YY = MaxY
TransposeXY XX, YY
RightClick XX, YY
End If
End Sub
Public Sub TransposeXY(x As Long, y As Long) 'X,Y is gridcoor
    x = CoorXbegin + (x - 1) * 16 + WInRECT.Left + 7
    y = CoorYBegin + (y - 1) * 16 + WInRECT.Top + 7
End Sub

Public Function CountSurrounding2(x As Long, y As Long, CountType As Long, CountTable() As Coor) As Long
Dim count As Long
For XX = x - 1 To x + 1
    For YY = y - 1 To y + 1
    If XX >= 1 And XX <= 30 Then
        If YY >= 1 And YY <= 16 Then
            If Not (XX = x And YY = y) Then
                If LocalGrid(XX, YY) = CountType Then
                    count = count + 1
                    CountTable(count).x = XX
                    CountTable(count).y = YY
                End If
            End If
        End If
    End If
    Next YY
Next XX
CountSurrounding2 = count
End Function

Public Function CountSurrounding(x As Long, y As Long, CountType As Long) As Long
Dim count As Long
If x >= 1 And x <= 30 Then
    If y >= 1 And y <= 16 Then
For XX = x - 1 To x + 1
    For YY = y - 1 To y + 1
    If XX >= 1 And XX <= 30 Then
        If YY >= 1 And YY <= 16 Then
            If Not (XX = x And YY = y) Then
                If LocalGrid(XX, YY) = CountType Then
                    count = count + 1
                End If
            End If
        End If
    End If
    Next YY
Next XX
    End If
End If
CountSurrounding = count
End Function


Public Sub RightClickSurrounding(x As Long, y As Long, ClickType As Long)
Debug.Print x, y
Dim tempx As Long, tempy As Long
For XX = x - 1 To x + 1
    For YY = y - 1 To y + 1
    If XX >= 1 And XX <= 30 Then
        If YY >= 1 And YY <= 16 Then
            If Not (XX = x And YY = y) Then
                If LocalGrid(XX, YY) = ClickType Then
                    tempx = XX
                    tempy = YY
                    TransposeXY tempx, tempy
                    RightClick tempx, tempy
                    LocalGrid(XX, YY) = Mine
                End If
            End If
        End If
    End If
    Next YY
Next XX
End Sub

Public Function MineCount()
Dim x As Long, y As Long, mc As Long
For x = 1 To 30
    For y = 1 To 16
        If LocalGrid(x, y) = Mine Then
            mc = mc + 1
        End If
    Next y
Next x
MineCount = mc
End Function
