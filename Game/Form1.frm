VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit
Public DirectX As New DirectX7
Public Primary As DirectDrawSurface7, DD As DirectDraw7, BackBuffer As DirectDrawSurface7
Public ball As DirectDrawSurface7, jPalette As DirectDrawSurface7
Dim Brick(4) As DirectDrawSurface7
Dim mExit As Boolean, GameStarted As Boolean
Dim Score As Integer
Public PaletteX As Single, PaletteY As Single
Public BallCount As Integer
Public Level As Integer
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Public gLang As String
Dim AllBalls() As tBall
Dim Bricks(200) As tBrick
Dim BrickNo As Integer
Private Type tBall
    X As Single
    Y As Single
    Active As Boolean
    SpeedX As Single
    SpeedY As Single
    bRect As RECT
End Type

Private Type tBrick
    X As Single
    Y As Single
    State As Integer
End Type

Sub InitGame()
    BallCount = BallCount - 1
    ShowCursor 0
    Set DD = DirectX.DirectDrawCreate("")
    DD.SetCooperativeLevel Form1.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWREBOOT
    DD.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
    
    Dim DDsD As DDSURFACEDESC2
        
    DDsD.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    DDsD.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
    DDsD.lBackBufferCount = 1
    Set Primary = DD.CreateSurface(DDsD)
    
    'Get the Backbuffer
    Dim DDSCap As DDSCAPS2
    DDSCap.lCaps = DDSCAPS_BACKBUFFER
    Set BackBuffer = Primary.GetAttachedSurface(DDSCap)
    '
    'Load the bitmaps
    Dim DDSBitMap(1) As DDSURFACEDESC2, DDSBtColorkey(1) As DDCOLORKEY
    Set ball = DD.CreateSurfaceFromFile(App.Path & "\ball.bmp", DDSBitMap(0))
    DDSBtColorkey(0).high = RGB(0, 0, 0)
    DDSBtColorkey(0).low = RGB(0, 0, 0)
    DDSBtColorkey(1).high = RGB(0, 0, 0)
    DDSBtColorkey(1).low = RGB(0, 0, 0)
    Set jPalette = DD.CreateSurfaceFromFile(App.Path & "\palette.bmp", DDSBitMap(1))
    ball.SetColorKey DDCKEY_SRCBLT, DDSBtColorkey(0)
    jPalette.SetColorKey DDCKEY_SRCBLT, DDSBtColorkey(1)
    Dim DDSBtBrk(3) As DDSURFACEDESC2, DDCOlorKEyBrk(3) As DDCOLORKEY
    For i = 1 To 4
        Set Brick(i) = DD.CreateSurfaceFromFile(App.Path & "\brick" & CStr(i) & ".bmp", DDSBtBrk(i - 1))
        DDCOlorKEyBrk(i - 1).high = RGB(0, 0, 0)
        DDCOlorKEyBrk(i - 1).low = RGB(0, 0, 0)
        Brick(i).SetColorKey DDCKEY_SRCBLT, DDCOlorKEyBrk(i - 1)
    Next
    Randomize Timer
     ReDim AllBalls(BallCount) As tBall
    For i = 0 To BallCount
        AllBalls(i).Active = True
        AllBalls(i).SpeedX = 4
        AllBalls(i).SpeedY = -3
        AllBalls(i).X = Rnd * 630 + 10
        AllBalls(i).Y = Rnd * 400 + 10
    Next
    SetLevel (1)
End Sub
Sub Game()
    Dim mColorKey As DDCOLORKEY
    mColorKey.high = RGB(0, 0, 0)
    mColorKey.low = RGB(0, 0, 0)
    ball.SetColorKey DDCKEY_SRCBLT, mColorKey
    BackBuffer.SetFontTransparency True
    BackBuffer.SetForeColor 3000
    Dim T
    T = Timer
    Do
        If GameStarted = True Then MoveAll
        DrawAll
        DoEvents
    Loop Until mExit = True
    Set DD = Nothing
    Set Primary = Nothing
    Set BackBuffer = Nothing
    Set ball = Nothing
    For i = 0 To UBound(Brick)
        Set Brick(i) = Nothing
    Next
    Set jPalette = Nothing
    Set DirectX = Nothing
    ShowCursor 1
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 39 '  —«” 
            If PaletteX < 560 Then PaletteX = PaletteX + 4 Else PaletteX = 560
        Case 37 'çÅ
            If PaletteX > 0 Then PaletteX = PaletteX - 4 Else PaletteX = 0
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then mExit = True: Unload Me
    If KeyAscii = 13 Then GameStarted = Not GameStarted
    'GameStarted = True
End Sub

Private Sub Form_Load()
    SetCurrentDirectory App.Path
    InitGame
    Game
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X <= 640 - 80 Then PaletteX = X Else PaletteX = 640 - 80
    PaletteY = 420
    For i = 0 To BallCount
        'If GameStarted = False Then AllBalls(i).X = PaletteX + 40 - 5: AllBalls(i).Y = 410
    Next
End Sub

Public Function jRect(X, Y, x2, y2) As RECT
    With jRect
        .Left = X
        .Right = x2
        .Top = Y
        .Bottom = y2
    End With
End Function

Sub DrawAll()
    Randomize Timer
    BackBuffer.BltColorFill jRect(0, 0, 640, 480), 100
    Dim i As Integer
    For i = 1 To BrickNo
        If Bricks(i).State > 0 Then
            BackBuffer.BltFast Bricks(i).X, Bricks(i).Y, Brick(Bricks(i).State), jRect(0, 0, 0, 0), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
    Next
        Dim mFont As IFont
        Set mFont = Form1.Font
        mFont.Name = "Arial"
        For i = 0 To BallCount
            BackBuffer.BltFast AllBalls(i).X, AllBalls(i).Y, ball, jRect(0, 0, 0, 0), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        Next
    mFont.Size = 12
    BackBuffer.SetForeColor vbGreen
    BackBuffer.SetFont mFont
    BackBuffer.DrawText 500, 440, "Point : " + CStr(Score), False
    BackBuffer.BltFast PaletteX, PaletteY, jPalette, jRect(0, 0, 0, 0), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    For i = 0 To BallCount
        If AllBalls(i).Active = True Then
            sbcnt = sbcnt + 1
        End If
    Next
        If sbcnt = 0 Then
            mFont.Size = 30
            BackBuffer.SetFont mFont
            BackBuffer.SetFontTransparency True
            BackBuffer.SetForeColor 700
            BackBuffer.BltColorFill jRect(0, 0, 640, 480), 4000
            BackBuffer.DrawText 220, 150, "You Lose", False
        End If
    For i = 1 To BrickNo
        If Bricks(i).State > 0 And Bricks(i).State <> 4 Then br = br + 1
    Next
    If br = 0 Then
        mFont.Size = 30
        With BackBuffer
            .SetFont mFont
            .SetForeColor 700
            .BltColorFill jRect(0, 0, 640, 480), 3500
            .DrawText 200, 200, "Bravo... You won", False
            SetLevel Level + 1
        End With
    End If
    Primary.Flip Nothing, DDFLIP_WAIT
End Sub

Sub MoveAll()
    CheckBall (1)
    Dim bBall As tBall
    For i = 0 To BallCount
        bBall = AllBalls(i)
        bBall.X = bBall.X + bBall.SpeedX
        bBall.Y = bBall.Y + bBall.SpeedY
        If bBall.X <= 0 Or bBall.X >= 630 Then bBall.SpeedX = -bBall.SpeedX
        If bBall.Y <= 0 Then bBall.SpeedY = -bBall.SpeedY
        If bBall.Y >= 480 Then bBall.Active = False
        AllBalls(i) = bBall
    Next
End Sub

Function CheckCollision(r1 As RECT, r2 As RECT, IntX As Single, IntY As Single) As Long
    Dim mTempRect As RECT
    CheckCollision = IntersectRect(mTempRect, r1, r2)
    If CheckCollision > 0 Then
        IntX = (mTempRect.Right + mTempRect.Left) / 2 - r1.Left
        IntY = (mTempRect.Bottom + mTempRect.Top) / 2 - r1.Top
    End If
End Function

Sub CheckBall(ball)
    Dim IntX As Single, IntY As Single, Tr As RECT
    Dim mCol As Long
    Randomize Timer
    Dim bBall  As tBall
    For j = 0 To BallCount
            AllBalls(j).bRect = jRect(AllBalls(j).X, AllBalls(j).Y, AllBalls(j).X + 10, AllBalls(j).Y + 10)
            mCol = IntersectRect(Tr, AllBalls(j).bRect, jRect(PaletteX, PaletteY, PaletteX + 80, PaletteY + 20))
            If mCol > 0 Then AllBalls(j).SpeedY = -AllBalls(j).SpeedY: AllBalls(j).SpeedX = Rnd * 4 * Sgn(AllBalls(j).SpeedX)
        For i = 1 To BrickNo
            mCol = CheckCollision(AllBalls(j).bRect, jRect(Bricks(i).X, Bricks(i).Y, Bricks(i).X + 40, Bricks(i).Y + 10), IntX, IntY)
            If mCol > 0 And Bricks(i).State > 0 Then
                'Beep 1000, 25
                'Bricks(i).State = 0
                If Bricks(i).State <> 4 Then
                    Beep 1000, 25: Bricks(i).State = Bricks(i).State - 1: Score = Score + 100
                End If
                    If IntY < Abs(AllBalls(j).SpeedY) Or IntY > 10 - Abs(AllBalls(j).SpeedY) Then AllBalls(j).SpeedY = -AllBalls(j).SpeedY: Exit For
                    If IntX < Abs(AllBalls(j).SpeedX) Or IntX > 40 - Abs(AllBalls(j).SpeedX) Then AllBalls(j).SpeedX = -AllBalls(j).SpeedX:  Exit For
            End If
        Next
         
        For k = j To BallCount
            If k <> j Then
                AllBalls(k).bRect = jRect(AllBalls(k).X, AllBalls(k).Y, AllBalls(k).X + 10, AllBalls(k).Y + 10)
                d = IntersectRect(Tr, AllBalls(j).bRect, AllBalls(k).bRect)
                If d > 0 Then
                    With AllBalls(k)
                        .SpeedX = -.SpeedX
                        .SpeedY = -.SpeedY
                    End With
                    With AllBalls(j)
                        .SpeedX = -.SpeedX
                        .SpeedY = -.SpeedY
                    End With
                End If
            End If
        Next
    Next
End Sub

Sub SetLevel(jLevel As Integer)
    Dim X As Integer, Y As Integer
    GameStarted = False
    Level = jLevel
    BrickNo = 0
    Select Case Level
    Case 1
            For X = 80 To 520 Step 40 '80,520
                For Y = 60 To 140 Step 10
                    BrickNo = BrickNo + 1
                    Bricks(BrickNo).X = X
                    Bricks(BrickNo).Y = Y
                    Bricks(BrickNo).State = 1
                Next
            Next
    Case 2
            For X = 80 To 520 Step 40 '80,520
                For Y = 60 To 140 Step 10
                    BrickNo = BrickNo + 1
                    Bricks(BrickNo).X = X
                    Bricks(BrickNo).Y = Y
                    If (X = 80 Or X = 520 Or Y = 60 Or Y = 140) Then Bricks(BrickNo).State = 3 Else Bricks(BrickNo).State = 1
                Next
            Next
            'BrickNo = BrickNo + 1
            'Bricks(BrickNo).State = 4
            'Bricks(BrickNo).X = 10
            'Bricks(BrickNo).Y = 40
            'BrickNo = BrickNo + 1
            'Bricks(BrickNo).State = 4
            'Bricks(BrickNo).X = 610
            'Bricks(BrickNo).Y = 40
    End Select
End Sub
