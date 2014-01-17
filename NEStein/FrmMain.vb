Public Class FrmMain
    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        EmuMenu.Renderer = New clsMenuRenderer

        Fill_Color_LookUp() 'Color LookUp for faster Background Rendering

        'Init All the Stuff
        Direct3D_Initialize()
        pAPU_Initialize()
        LoadDefaultPalette()
        Hi_Res_Timer_Initialize()
        Joystick_Initialize()

        LoadConfig()

        Show()
    End Sub
    Private Sub FrmMain_FormClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.FormClosing
        DirectSound_Shutdown()
        SaveConfig()
        End
    End Sub
    Private Sub FrmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        KeyCodes(e.KeyCode) = &H41
    End Sub
    Private Sub FrmMain_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        KeyCodes(e.KeyCode) = &H40
    End Sub

    '-------------------------------------------------------------------------

    Private Sub AbrirROMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuOpen.Click
        AbreRom.ShowDialog()
        If AbreRom.FileName <> Nothing Then LoadROM(AbreRom.FileName)
    End Sub
    Private Sub FecharROMToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuClose.Click
        DestroyNes()
        NesScreen.Image = Nothing
    End Sub
    Private Sub BGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BGToolStripMenuItem.Click
        ShowBG = Not ShowBG
        BGToolStripMenuItem.Checked = ShowBG
    End Sub
    Private Sub SpriteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpriteToolStripMenuItem.Click
        ShowSP = Not ShowSP
        SpriteToolStripMenuItem.Checked = ShowSP
    End Sub
    Private Sub FPSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FPSToolStripMenuItem.Click
        ShowFPS = Not ShowFPS
        FPSToolStripMenuItem.Checked = ShowFPS
    End Sub
    Private Sub AtivadoToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AtivadoToolStripMenuItem1.Click
        DipSwitch = Not DipSwitch
        AtivadoToolStripMenuItem1.Checked = DipSwitch
    End Sub
    Private Sub LimitarFPSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LimitarFPSToolStripMenuItem.Click
        LimitFPS = Not LimitFPS
        LimitarFPSToolStripMenuItem.Checked = LimitFPS
    End Sub
    'APU and Channels On/Off
    Private Sub AtivadoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AtivadoToolStripMenuItem.Click
        APU.Enabled = Not APU.Enabled
        AtivadoToolStripMenuItem.Checked = APU.Enabled
    End Sub
    Private Sub Square1ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Square1ToolStripMenuItem.Click
        Square_Channel_1.TurnOn = Not Square_Channel_1.TurnOn
        Square1ToolStripMenuItem.Checked = Square_Channel_1.TurnOn
    End Sub
    Private Sub Square2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Square2ToolStripMenuItem.Click
        Square_Channel_2.TurnOn = Not Square_Channel_2.TurnOn
        Square2ToolStripMenuItem.Checked = Square_Channel_2.TurnOn
    End Sub
    Private Sub TriangleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TriangleToolStripMenuItem.Click
        Triangle_Channel.TurnOn = Not Triangle_Channel.TurnOn
        TriangleToolStripMenuItem.Checked = Triangle_Channel.TurnOn
    End Sub
    Private Sub NoiseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NoiseToolStripMenuItem.Click
        Noise_Channel.TurnOn = Not Noise_Channel.TurnOn
        NoiseToolStripMenuItem.Checked = Noise_Channel.TurnOn
    End Sub
    Private Sub DMCToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DMCToolStripMenuItem.Click
        DMC_Channel.TurnOn = Not DMC_Channel.TurnOn
        DMCToolStripMenuItem.Checked = DMC_Channel.TurnOn
    End Sub

    '-------------------------------------------------------------------------

    Private Sub PauseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PauseToolStripMenuItem.Click
        NES_Pause = Not NES_Pause
        PauseToolStripMenuItem.Checked = NES_Pause
        If NES_Pause Then
            Pause_Sound(Sound_Buffer)
        Else
            Play_Sound_Loop(Sound_Buffer)
        End If
    End Sub
    Private Sub ResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetToolStripMenuItem.Click
        Mapper_Reset()
        CPU_Reset()
    End Sub

    '-------------------------------------------------------------------------

    'Private Sub SlotSel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '
    '   Uncheck_Menus(SlotToolStripMenuItem)
    '   sender.Checked = True
    '   Slot = sender.Tag
    'End Sub
    Private Sub FrameSkipSel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuFS0.Click, _
        MnuFS1.Click, _
        MnuFS2.Click, _
        MnuFS3.Click, _
        MnuFS4.Click

        Uncheck_Menus(FrameSkipToolStripMenuItem)
        sender.Checked = True
        FrameSkip = sender.Tag + 1
    End Sub
    Private Sub ZoomControls_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XToolStripMenuItem.Click, _
        XToolStripMenuItem1.Click

        Zoom = sender.Tag + 1
        Window_Setup()
    End Sub
    Private Sub InterpolationControls_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NNToolStripMenuItem.Click, _
        LowToolStripMenuItem.Click, _
        HighToolStripMenuItem.Click, _
        BilinearToolStripMenuItem.Click, _
        BicubicToolStripMenuItem.Click

        InterpolMode = sender.Tag
    End Sub
    Private Sub TelaCheiaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TelaCheiaToolStripMenuItem.Click
        FullScreen = Not FullScreen
        If FullScreen Then
            FrmRender.Show()
            Me.Hide()
        End If
    End Sub
    Private Sub ConfigurarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfigurarToolStripMenuItem.Click
        FrmConfigKeys.Show()
    End Sub

    '-------------------------------------------------------------------------

    Private Sub SaveConfig()
        '*** Input ***
        'Controller 1
        My.Settings.Input_K1A = InputKeys(0, 0)
        My.Settings.Input_K1B = InputKeys(0, 1)
        My.Settings.Input_K1Select = InputKeys(0, 2)
        My.Settings.Input_K1Start = InputKeys(0, 3)
        My.Settings.Input_K1Up = InputKeys(0, 4)
        My.Settings.Input_K1Down = InputKeys(0, 5)
        My.Settings.Input_K1Left = InputKeys(0, 6)
        My.Settings.Input_K1Right = InputKeys(0, 7)
        'Controller 2
        My.Settings.Input_K2A = InputKeys(1, 0)
        My.Settings.Input_K2B = InputKeys(1, 1)
        My.Settings.Input_K2Select = InputKeys(1, 2)
        My.Settings.Input_K2Start = InputKeys(1, 3)
        My.Settings.Input_K2Up = InputKeys(1, 4)
        My.Settings.Input_K2Down = InputKeys(1, 5)
        My.Settings.Input_K2Left = InputKeys(1, 6)
        My.Settings.Input_K2Right = InputKeys(1, 7)

        My.Settings.Save()
    End Sub
    Private Sub LoadConfig()
        '*** Input ***
        'Controller 1
        InputKeys(0, 0) = My.Settings.Input_K1A
        InputKeys(0, 1) = My.Settings.Input_K1B
        InputKeys(0, 2) = My.Settings.Input_K1Select
        InputKeys(0, 3) = My.Settings.Input_K1Start
        InputKeys(0, 4) = My.Settings.Input_K1Up
        InputKeys(0, 5) = My.Settings.Input_K1Down
        InputKeys(0, 6) = My.Settings.Input_K1Left
        InputKeys(0, 7) = My.Settings.Input_K1Right
        'Controller 2
        InputKeys(1, 0) = My.Settings.Input_K2A
        InputKeys(1, 1) = My.Settings.Input_K2B
        InputKeys(1, 2) = My.Settings.Input_K2Select
        InputKeys(1, 3) = My.Settings.Input_K2Start
        InputKeys(1, 4) = My.Settings.Input_K2Up
        InputKeys(1, 5) = My.Settings.Input_K2Down
        InputKeys(1, 6) = My.Settings.Input_K2Left
        InputKeys(1, 7) = My.Settings.Input_K2Right
    End Sub
    Private Sub Window_Setup()
        With NesScreen
            .Width = 256 * Zoom
            .Height = 240 * Zoom
        End With

        With Me
            .Width = NesScreen.Width + (Me.Width - Me.ClientSize.Width)
            .Height = NesScreen.Height + ((Me.Height - Me.ClientSize.Height) + EmuMenu.Height)
            .Left = (Screen.PrimaryScreen.Bounds.Width / 2) - (.Width / 2)
            .Top = (Screen.PrimaryScreen.Bounds.Height / 2) - (.Height / 2)
        End With
    End Sub
    Private Sub Uncheck_Menus(ByVal Menu As ToolStripMenuItem)
        Dim mnu As ToolStripMenuItem
        For Each mnu In Menu.DropDownItems
            mnu.Checked = False
        Next
    End Sub
End Class
