Option Explicit On
Imports Microsoft.DirectX
Imports Microsoft.DirectX.DirectSound
Module APU_DirectSound
    Private Sound_Device As New Device
    Public Sound_Buffer As SecondaryBuffer
    Private Sound_Buffer_Description As New BufferDescription()
    Public Sound_Buffer_Wave_Format As New WaveFormat()

    Public Const MAX_VOLUME As Integer = 100
    Public Buffer_Length As Integer, Half_Buffer_Length As Integer
    Public Buffer_Size As Integer
    Public Sound_Data() As Byte
    Public Sub DirectSound_Initialize()
        With Sound_Buffer_Wave_Format
            .FormatTag = WaveFormatTag.Pcm
            .Channels = 1 '1 = Mono   2 = Stereo. The NES is Mono
            .BitsPerSample = 16
            .SamplesPerSecond = 44100
            .BlockAlign = (.BitsPerSample * .Channels) / 8
            .AverageBytesPerSecond = ((.BitsPerSample / 8) * .Channels) * .SamplesPerSecond

            Buffer_Size = .AverageBytesPerSecond '* 5 Cause Force Close :(
            ReDim Sound_Data(Buffer_Size)

            Buffer_Length = .AverageBytesPerSecond
            Half_Buffer_Length = .AverageBytesPerSecond / 2
            Half_Buffer_Length = Half_Buffer_Length + (Half_Buffer_Length Mod .BlockAlign)
        End With

        With Sound_Buffer_Description
            .Format = Sound_Buffer_Wave_Format
            .BufferBytes = Buffer_Length
            .Flags = BufferDescriptionFlags.ControlPositionNotify Or _
                      BufferDescriptionFlags.StickyFocus Or _
                      BufferDescriptionFlags.ControlFrequency Or _
                      BufferDescriptionFlags.ControlPan Or _
                      BufferDescriptionFlags.ControlVolume
        End With

        Sound_Device.SetCooperativeLevel(FrmMain.Handle, CooperativeLevel.Priority)
        Sound_Buffer = New SecondaryBuffer(Sound_Buffer_Description, Sound_Device)
        Sound_Channel_Clear_Buffer()
        Set_Sound_Volume(Sound_Buffer, 100)
    End Sub
    Public Sub Play_Sound_Loop(ByVal Sound_Buffer As SecondaryBuffer)
        Sound_Buffer.Play(0, BufferPlayFlags.Looping)
    End Sub
    Public Sub Pause_Sound(ByVal Sound_Buffer As SecondaryBuffer)
        Sound_Buffer.Stop()
    End Sub
    Public Sub Set_Sound_Volume(ByVal Buffer As SecondaryBuffer, ByVal Volume As Long)
        If Volume >= MAX_VOLUME Then Volume = MAX_VOLUME
        If Volume <= 0 Then Volume = 0
        Buffer.Volume = ((Volume / MAX_VOLUME) * 10000) + -10000
    End Sub
    Public Sub DirectSound_Shutdown()
        Sound_Buffer = Nothing
    End Sub
End Module
