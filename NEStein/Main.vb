Option Explicit On
Imports System.IO
Module Main
    'Binary Constants
    Public Const BIT_0 As Long = &H1     'Decimal: 1
    Public Const BIT_1 As Long = &H2     'Decimal: 2
    Public Const BIT_2 As Long = &H4     'Decimal: 4
    Public Const BIT_3 As Long = &H8     'Decimal: 8
    Public Const BIT_4 As Long = &H10    'Decimal: 16
    Public Const BIT_5 As Long = &H20    'Decimal: 32
    Public Const BIT_6 As Long = &H40    'Decimal: 64
    Public Const BIT_7 As Long = &H80    'Decimal: 128
    Public Const BIT_8 As Long = &H100    'Decimal: 256
    Public Const BIT_9 As Long = &H200    'Decimal: 512
    Public Const BIT_10 As Long = &H400   'Decimal: 1024
    Public Const BIT_11 As Long = &H800   'Decimal: 2048
    Public Const BIT_12 As Long = &H1000  'Decimal: 4096
    Public Const BIT_13 As Long = &H2000  'Decimal: 8192
    Public Const BIT_14 As Long = &H4000  'Decimal: 16384
    Public Const BIT_15 As Long = &H8000  'Decimal: 32768

    'FPS and Timing
    Private Declare Function QueryPerformanceCounter Lib "Kernel32" (ByRef X As Long) As Integer
    Private Declare Function QueryPerformanceFrequency Lib "Kernel32" (ByRef X As Long) As Integer

    Private Ticks_Per_Second As Long
    Private Start_Time As Long

    Private Milliseconds As Integer
    Private Get_Frames_Per_Second As Integer
    Private Frame_Count As Integer

    'ROM Arrays
    Public PROM() As Byte
    Public VROM() As Byte

    'Header info
    Public Structure Header
        Dim PrgSize As Byte
        Dim ChrSize As Byte
        Dim Mapper As Byte
        Dim Mirroring As Byte
        Dim Trainer As Byte
        Dim FourScreen As Byte
        Dim SRAM As Boolean
    End Structure
    Public ROMHeader As Header

    'Input
    Public KeyCodes(255) As Byte 'Used to store Pressed keys
    Public Controller_1(7) As Byte
    Public Controller_2(7) As Byte
    Public InputKeys(1, 7) As Byte

    'Config
    Public FrameSkip As Integer = 1
    Public Slot As Integer
    Public ShowBG As Boolean = True
    Public ShowSP As Boolean = True
    Public ShowFPS As Boolean = True
    Public LimitFPS As Boolean = True
    Public DipSwitch As Boolean
    Public FullScreen As Boolean
    Public NES_On, NES_Pause As Boolean
    Public Sub LoadROM(ByVal FileName As String)
        DestroyNes()

        ' Load ROM File
        '---------------------------------------
        Dim Input As New FileStream(FileName, FileMode.Open)
        Dim ROM(CInt(Input.Length - 1)) As Byte
        Input.Read(ROM, 0, CInt(Input.Length))
        Input.Close()

        ' Parse Header
        '---------------------------------------
        Dim PrgMark As Integer
        Dim ROMCtrl As Byte
        Dim ROMCtrl2 As Byte

        ROMHeader.PrgSize = ROM(4)
        ROMHeader.ChrSize = ROM(5)
        ROMCtrl = ROM(6)
        ROMCtrl2 = ROM(7)

        ROMHeader.Mapper = (ROMCtrl And &HF0) \ 16
        ROMHeader.Mapper += ROMCtrl2

        ROMHeader.Mirroring = ROMCtrl And &H1
        ROMHeader.Trainer = ROMCtrl And &H4
        ROMHeader.FourScreen = ROMCtrl And &H8
        If ROMCtrl And &H2 Then ROMHeader.SRAM = True
        PrgMark = ROMHeader.PrgSize * &H4000

        ' Fill PRG-ROM and Video ROM
        '---------------------------------------
        ReDim PROM(PrgMark)
        Dim StartAt As Integer
        If ROMHeader.Trainer Then StartAt = 528 Else StartAt = 16
        Buffer.BlockCopy(ROM, StartAt, PROM, 0, UBound(PROM))

        ReDim VROM(ROMHeader.ChrSize * &H2000)
        PrgMark = &H4000 * ROMHeader.PrgSize + StartAt
        If ROMHeader.ChrSize Then Buffer.BlockCopy(ROM, PrgMark, VROM, 0, UBound(VROM))

        ' Reset NES
        '---------------------------------------
        Mapper_Reset()
        Reset_CPU()
        For i As Integer = 0 To 7 'Erase Input
            Controller_1(i) = &H40
            Controller_2(i) = &H40
        Next

        ' Setup Mirroring
        '---------------------------------------
        If ROMHeader.FourScreen Then ROMHeader.Mirroring = MIRROR_FOURSCREEN
        Do_Mirroring()

        Play_Sound_Loop(Sound_Buffer) 'Start Sound Buffer

        Run() 'Turn on the NES!
    End Sub
    Public Sub Run()
        NES_On = True
        Do While NES_On
            Do While NES_Pause
                Application.DoEvents()
            Loop
            If LimitFPS Then Lock_Framerate(60)
            Execute_CPU()
            FrmMain.Text = Get_FPS()
            Application.DoEvents()
        Loop
    End Sub
    Public Sub ReadInput()
        If JoystickPlugged() = 0 Then
            'Keyboard only
            For i As Integer = 0 To 7
                Controller_1(i) = KeyCodes(InputKeys(0, i))
                Controller_2(i) = KeyCodes(InputKeys(1, i))
            Next i
        ElseIf JoystickPlugged() = 1 Then
            'Controller 1 - Gamepad
            For i As Integer = 0 To 3
                If GetJoy_Btn(0, i) Then Controller_1(i) = &H41 Else Controller_1(i) = &H40
            Next
            If GetJoy_Y(0) = 0 Then Controller_1(4) = &H41 Else Controller_1(4) = &H40
            If GetJoy_Y(0) = &HFFFF Then Controller_1(5) = &H41 Else Controller_1(5) = &H40
            If GetJoy_X(0) = 0 Then Controller_1(6) = &H41 Else Controller_1(6) = &H40
            If GetJoy_X(0) = &HFFFF Then Controller_1(7) = &H41 Else Controller_1(7) = &H40

            'Controller 2 - Keyboard
            For i As Integer = 0 To 7
                Controller_2(i) = KeyCodes(InputKeys(1, i))
            Next i
        ElseIf JoystickPlugged() = 2 Then
            'Controller 1 - Gamepad
            For i As Integer = 0 To 3
                If GetJoy_Btn(0, i) Then Controller_1(i) = &H41 Else Controller_1(i) = &H40
            Next
            If GetJoy_Y(0) = 0 Then Controller_1(4) = &H41 Else Controller_1(4) = &H40
            If GetJoy_Y(0) = &HFFFF Then Controller_1(5) = &H41 Else Controller_1(5) = &H40
            If GetJoy_X(0) = 0 Then Controller_1(6) = &H41 Else Controller_1(6) = &H40
            If GetJoy_X(0) = &HFFFF Then Controller_1(7) = &H41 Else Controller_1(7) = &H40

            'Controller 2 - Gamepad
            For i As Integer = 0 To 3
                If GetJoy_Btn(1, i) Then Controller_2(i) = &H41 Else Controller_2(i) = &H40
            Next
            If GetJoy_Y(1) = 0 Then Controller_2(4) = &H41 Else Controller_2(4) = &H40
            If GetJoy_Y(1) = &HFFFF Then Controller_2(5) = &H41 Else Controller_2(5) = &H40
            If GetJoy_X(1) = 0 Then Controller_2(6) = &H41 Else Controller_2(6) = &H40
            If GetJoy_X(1) = &HFFFF Then Controller_2(7) = &H41 Else Controller_2(7) = &H40
        End If
    End Sub
    Public Sub DestroyNes()
        'Turn off the NES
        NES_On = False

        'Stop Sound
        Pause_Sound(Sound_Buffer)

        'Erase Video Buffer
        Array.Clear(Video_Buffer, 0, Video_Buffer.Length)

        'Erase ROM Data
        Erase PROM
        Erase VROM

        'Erase RAM Stuff
        Array.Clear(VRAM, 0, VRAM.Length)
        Array.Clear(SpriteRAM, 0, SpriteRAM.Length)
        Array.Clear(Bank0, 0, Bank0.Length)
        Array.Clear(Bank6, 0, Bank6.Length)
        Array.Clear(Bank8, 0, Bank8.Length)
        Array.Clear(BankA, 0, BankA.Length)
        Array.Clear(BankC, 0, BankC.Length)
        Array.Clear(BankE, 0, BankE.Length)

        'Erase Input
        Controller1_Count = 0
        Controller2_Count = 0
    End Sub
    Public Sub Fill_Array(a() As Integer, ParamArray B() As Object)
        For i As Integer = 0 To UBound(a)
            a(i) = B(i)
        Next i
    End Sub
    Public Sub Fill_Array_Byte(a() As Byte, ParamArray B() As Object)
        For i As Integer = 0 To UBound(a)
            a(i) = B(i)
        Next i
    End Sub

    '-------------------------------------------------------------------------

    Public Function Hi_Res_Timer_Initialize() As Boolean
        If QueryPerformanceFrequency(Ticks_Per_Second) = 0 Then
            Hi_Res_Timer_Initialize = False
        Else
            QueryPerformanceCounter(Start_Time)
            Hi_Res_Timer_Initialize = True
        End If
    End Function
    Private Function Get_Elapsed_Time() As Single
        Dim Last_Time As Long
        Dim Current_Time As Long

        QueryPerformanceCounter(Current_Time)
        Get_Elapsed_Time = Convert.ToSingle((Current_Time - Last_Time) / Ticks_Per_Second)
        QueryPerformanceCounter(Last_Time)
    End Function
    Private Function Get_Elapsed_Time_Per_Frame() As Single
        Static Last_Time As Long
        Static Current_Time As Long

        QueryPerformanceCounter(Current_Time)
        Get_Elapsed_Time_Per_Frame = Convert.ToSingle((Current_Time - Last_Time) / Ticks_Per_Second)
        QueryPerformanceCounter(Last_Time)
    End Function
    Private Sub Lock_Framerate(ByVal Target_FPS As Long)
        Static Last_Time As Long
        Dim Current_Time As Long
        Dim FPS As Single

        Do
            QueryPerformanceCounter(Current_Time)
            FPS = Convert.ToSingle(Ticks_Per_Second / (Current_Time - Last_Time))
        Loop While (FPS > Target_FPS)

        QueryPerformanceCounter(Last_Time)
    End Sub
    Public Function Get_FPS() As String
        Frame_Count = Frame_Count + 1

        If Get_Elapsed_Time() - Milliseconds >= 1 Then
            Get_Frames_Per_Second = Frame_Count
            Frame_Count = 0
            Milliseconds = Convert.ToInt32(Get_Elapsed_Time)
        End If

        Get_FPS = Get_Frames_Per_Second & " fps"
    End Function

#Region "Palette and Color stuff"
    Public Sub LoadDefaultPalette()
        'Row 1
        Palette(0) = NES_RGB(104, 104, 104)
        Palette(1) = NES_RGB(0, 42, 136)
        Palette(2) = NES_RGB(20, 18, 167)
        Palette(3) = NES_RGB(59, 0, 164)
        Palette(4) = NES_RGB(92, 0, 126)
        Palette(5) = NES_RGB(110, 0, 64)
        Palette(6) = NES_RGB(108, 6, 0)
        Palette(7) = NES_RGB(86, 29, 0)
        Palette(8) = NES_RGB(51, 53, 0)
        Palette(9) = NES_RGB(11, 72, 0)
        Palette(10) = NES_RGB(0, 82, 0)
        Palette(11) = NES_RGB(0, 79, 8)
        Palette(12) = NES_RGB(0, 64, 77)
        Palette(13) = NES_RGB(0, 0, 0)
        Palette(14) = NES_RGB(0, 0, 0)
        Palette(15) = NES_RGB(0, 0, 0)

        'Row 2
        Palette(16) = NES_RGB(173, 173, 173)
        Palette(17) = NES_RGB(21, 95, 217)
        Palette(18) = NES_RGB(66, 64, 255)
        Palette(19) = NES_RGB(117, 39, 254)
        Palette(20) = NES_RGB(160, 26, 204)
        Palette(21) = NES_RGB(183, 30, 123)
        Palette(22) = NES_RGB(141, 49, 12)
        Palette(23) = NES_RGB(113, 63, 0)
        Palette(24) = NES_RGB(100, 99, 46)
        Palette(25) = NES_RGB(56, 135, 0)
        Palette(26) = NES_RGB(0, 110, 0)
        Palette(27) = NES_RGB(0, 143, 50)
        Palette(28) = NES_RGB(0, 104, 101)
        Palette(29) = NES_RGB(0, 0, 0)
        Palette(30) = NES_RGB(0, 0, 0)
        Palette(31) = NES_RGB(0, 0, 0)

        'Row 3
        Palette(32) = NES_RGB(255, 254, 255)
        Palette(33) = NES_RGB(100, 176, 255)
        Palette(34) = NES_RGB(146, 154, 255)
        Palette(35) = NES_RGB(198, 118, 255)
        Palette(36) = NES_RGB(243, 106, 255)
        Palette(37) = NES_RGB(254, 110, 204)
        Palette(38) = NES_RGB(254, 129, 112)
        Palette(39) = NES_RGB(225, 184, 114)
        Palette(40) = NES_RGB(188, 190, 0)
        Palette(41) = NES_RGB(175, 196, 0)
        Palette(42) = NES_RGB(92, 228, 48)
        Palette(43) = NES_RGB(69, 224, 130)
        Palette(44) = NES_RGB(72, 205, 222)
        Palette(45) = NES_RGB(79, 79, 79)
        Palette(46) = NES_RGB(0, 0, 0)
        Palette(47) = NES_RGB(0, 0, 0)

        'Row 4
        Palette(48) = NES_RGB(255, 254, 255)
        Palette(49) = NES_RGB(192, 223, 255)
        Palette(50) = NES_RGB(211, 210, 255)
        Palette(51) = NES_RGB(232, 200, 255)
        Palette(52) = NES_RGB(251, 194, 255)
        Palette(53) = NES_RGB(254, 196, 234)
        Palette(54) = NES_RGB(254, 204, 197)
        Palette(55) = NES_RGB(247, 236, 225)
        Palette(56) = NES_RGB(228, 229, 148)
        Palette(57) = NES_RGB(207, 239, 150)
        Palette(58) = NES_RGB(189, 244, 171)
        Palette(59) = NES_RGB(179, 243, 204)
        Palette(60) = NES_RGB(181, 235, 242)
        Palette(61) = NES_RGB(184, 184, 184)
        Palette(62) = NES_RGB(0, 0, 0)
        Palette(63) = NES_RGB(0, 0, 0)

        For Current_Color As Integer = 0 To 63
            Palette(Current_Color + 64) = Palette(Current_Color)
            Palette(Current_Color + 128) = Palette(Current_Color)
            Palette(Current_Color + 192) = Palette(Current_Color)
        Next Current_Color
    End Sub
    Public Sub Fill_Color_LookUp()
        Dim c As Integer

        For b1 As Integer = 0 To 255
            For b2 As Integer = 0 To 255
                For X As Integer = 0 To 7
                    If b1 And (1 << X) Then c = 1 Else c = 0
                    If b2 And (1 << X) Then c = c + 2
                    Color_LookUp(b1 * 2048 + b2 * 8 + X) = c
                Next X
        Next b2, b1
    End Sub
    Public Function NES_RGB(B As Byte, G As Byte, R As Byte)
        Return B * 65536 + G * 256 + R
    End Function
#End Region

End Module
