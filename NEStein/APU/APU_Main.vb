Option Explicit On
Imports Microsoft.DirectX
Imports Microsoft.DirectX.DirectSound
Module APU_Main
    'Definitions:

    'Envelope - The way a sound's parameter changes over time.

    'Divider - outputs a clock every n input clocks, where n is the divider's period. It contains a counter which is decremented
    '          on the arrival of each clock. When the counter reaches 0, it is reloaded with the period and an output clock is generated.
    '          A divider can also be forced to reload its counter immediately, but this does not output a clock. When a divider's period
    '          is changed, the current count is not affected.

    '          A divider may be implemented as a down counter (5, 4, 3, ...) or as a linear feedback shift register (LFSR). The dividers
    '          in the pulse and triangle channels are linear down-counters. The dividers for noise, DMC, and the APU Frame Counter are
    '          implemented as LFSRs to save gates compared to the equivalent down counter.

    'Length Counter - The pulse, triangle, and noise channels each have their own length counter unit. It is clocked twice per sequence,
    '                 and counts down to zero if enabled. When the length counter reaches zero the channel is silenced. It is reloaded by
    '                 writing a 5-bit value to the appropriate channel's length counter register, which will load a value from a lookup table.

    '                 The triangle channel has an additional linear counter unit which is clocked four times per sequence (like the envelope
    '                 of the other channels). It functions independently of the length counter, and will also silence the triangle channel
    '                 when it reaches zero.

    'Gate - Takes the input on the left and outputs it on the right, unless the control input on top tells the gate to ignore the input and
    '       always output 0.

    'Sequencer - Continuously loops over a sequence of values or events. When clocked, the next item in the sequence is generated.

    'Timer - Used in each of the five channels to control the sound frequency. It contains a divider which is clocked by the CPU clock. The
    '        triangle channel's timer is clocked on every CPU cycle, but the pulse, noise, and DMC timers are clocked only on every second
    '        CPU cycle and thus produce only even periods.

    'Low Pass Filter - an electronic filter that passes low-frequency signals and attenuates (reduces the amplitude of) signals with frequencies
    '                  higher than the cutoff frequency.

    Public Structure NES_Channel_Type
        Dim Volume As Integer
        Dim Length_Counter As Integer
        Dim Length_Counter_Disable As Boolean
        Dim Envelope As Integer
        Dim Envelope_Counter As Integer
        Dim Envelope_Decay_Disable As Boolean
        Dim Envelope_Start_Flag As Boolean
        Dim Duty As Integer
        Dim DutyPercentage As Double
        Dim Length_Values() As Integer
        Dim Wavelength As Integer
        Dim WaveStatus As Boolean
        Dim Frequency As Double
        Dim RenderedWavelength As Double
        Dim SamplingRate As Double
        Dim Enabled As Boolean 'Enabled / Disabled by the CPU
        Dim TurnOn As Boolean 'Enabled / Disabled by the user
        Dim SampleCount As Double

        'Square 1 & 2
        Dim Sweep_Enabled As Boolean
        Dim Sweep_Right_Shift As Integer
        Dim Sweep_Negate As Boolean
        Dim Sweep_Refresh_Rate As Integer
        Dim Sweep_Counter As Integer
        Dim Sweep_Reload As Boolean
        Dim Sweep_Position As Integer
        Dim Sweep_Silence As Boolean

        'Triangle
        Dim LinearCounter As Integer
        Dim LinearCounterLoad As Integer
        Dim Triangle_Sequence_Data() As Byte
        Dim Triangle_Sequence As Integer
        Dim Linear_Counter_Disable As Boolean

        'Noise
        Dim Noise_Random_Generator_Mode As Boolean
        Dim NoiseShiftData As Integer
        Dim NoiseWavelengths() As Integer
        Dim Noise_Feedback As Integer

        'DMC
        Dim SampleAddress As Integer
        Dim SampleLength As Integer
        Dim InitialAddress As Integer
        Dim InitialLength As Integer
        Dim Shift As Integer
        Dim DAC As Integer
        Dim DAC_Counter As Integer
        Dim DMCWavelengths() As Integer
        Dim DMC_Loop As Byte
        Dim DMC_IrqOn As Byte
    End Structure

    Public Structure NES_Sound_Type
        Dim Length_Values() As Integer
        Dim DataPosition As Integer
        Dim LastPosition As Integer
        Dim FirstRender As Boolean
        Dim Enabled As Boolean
        Dim Frame_Counter As Integer
        Dim Timer As Double
        Dim Filter As Integer
    End Structure

    Public Const APU_FILTER_NONE As Long = 0
    Public Const APU_FILTER_LOW_PASS As Long = 1

    Public Square_Channel_1 As NES_Channel_Type
    Public Square_Channel_2 As NES_Channel_Type
    Public Triangle_Channel As NES_Channel_Type
    Public Noise_Channel As NES_Channel_Type
    Public DMC_Channel As NES_Channel_Type

    Public APU As NES_Sound_Type

    Public Frame_Bits As Long
    Public Frame_Counter_Mode As Long
    Public IRQ_Disable As Boolean
    Public Initial_Time As Double
    Public Mixed_Sound As Long

    Public DoSound As Boolean
    Public Sub Sound_Write(ByVal Address As Integer, ByVal Value As Byte)
        Select Case Address
            Case &H4000 'pAPU Pulse 1 Control Register
                '$4000 / $4004   DDLC VVVV   Duty (D), envelope loop / length counter disable (L), constant volume (C), volume/envelope (V)
                Square_Channel_1_Write_Register_1(Value)
            Case &H4001 'pAPU Pulse 1 Ramp Control Register
                '$4001 / $4005   EPPP NSSS   Sweep unit: enabled (E), period (P), negate (N), shift (S)
                Square_Channel_1_Write_Register_2(Value)
            Case &H4002 'pAPU Pulse 1 Fine Tune (FT) Register
                '$4002 / $4006   TTTT TTTT   Timer low (T)
                Square_Channel_1_Write_Register_3(Value)
            Case &H4003 'pAPU Pulse 1 Coarse Tune (CT) Register
                '$4003 / $4007   LLLL LTTT   Length counter load (L), timer high (T)
                Square_Channel_1_Write_Register_4(Value)
            Case &H4004
                Square_Channel_2_Write_Register_1(Value)
            Case &H4005
                Square_Channel_2_Write_Register_2(Value)
            Case &H4006
                Square_Channel_2_Write_Register_3(Value)
            Case &H4007
                Square_Channel_2_Write_Register_4(Value)
            Case &H4008
                Triangle_Channel_Write_Register_1(Value)
            Case &H4009
                Triangle_Channel_Write_Register_2(Value)
            Case &H400A
                Triangle_Channel_Write_Register_3(Value)
            Case &H400B
                Triangle_Channel_Write_Register_4(Value)
            Case &H400C
                Noise_Channel_Write_Register_1(Value)
            Case &H400D
                'Not used at all. Noise channel only uses $400C, $400E, and $400F
            Case &H400E
                Noise_Channel_Write_Register_2(Value)
            Case &H400F
                Noise_Channel_Write_Register_3(Value)
            Case &H4010
                DMC_Channel_Write_Register_1(Value)
            Case &H4011
                DMC_Channel_Write_Register_2(Value)
            Case &H4012
                DMC_Channel_Write_Register_3(Value)
            Case &H4013
                DMC_Channel_Write_Register_4(Value)
        End Select
    End Sub
    Public Sub pAPU_Initialize()
        'Init Direct Sound
        DirectSound_Initialize()

        'Load Length Tables
        ReDim APU.Length_Values(31)
        Fill_Array(APU.Length_Values, &H5, &H7F, _
                   &HA, &H1, _
                   &H14, &H2, _
                   &H28, &H3, _
                   &H50, &H4, _
                   &H1E, &H5, _
                   &H7, &H6, _
                   &HD, &H7, _
                   &H6, &H8, _
                   &HC, &H9, _
                   &H18, &HA, _
                   &H30, &HB, _
                   &H60, &HC, _
                   &H24, &HD, _
                   &H8, &HE, _
                   &H10, &HF)
        ReDim Triangle_Channel.Triangle_Sequence_Data(31)
        Fill_Array_Byte(Triangle_Channel.Triangle_Sequence_Data, &H0, &H1, _
                        &H2, &H3, _
                        &H4, &H5, _
                        &H6, &H7, _
                        &H8, &H9, _
                        &HA, &HB, _
                        &HC, &HD, _
                        &HE, &HF, _
                        &HF, &HE, _
                        &HD, &HC, _
                        &HB, &HA, _
                        &H9, &H8, _
                        &H7, &H6, _
                        &H5, &H4, _
                        &H3, &H2, _
                        &H1, &H0)
        ReDim Noise_Channel.NoiseWavelengths(15)
        Fill_Array(Noise_Channel.NoiseWavelengths, &H4, &H8, _
                   &H10, &H20, _
                   &H40, &H60, _
                   &H80, &HA0, _
                   &HCA, &HFE, _
                   &H17C, &H1FC, _
                   &H2FA, &H3FB, _
                   &H7F2, &HF34)
        ReDim DMC_Channel.DMCWavelengths(15)
        Fill_Array(DMC_Channel.DMCWavelengths, &HD60, &HBE0, _
                   &HAA0, &HA00, _
                   &H8F0, &H7F0, _
                   &H710, &H6B0, _
                   &H5F0, &H500, _
                   &H470, &H400, _
                   &H350, &H2A8, _
                   &H240, &H1B0)

        APU.Enabled = True
        APU.FirstRender = True
        APU.Filter = APU_FILTER_LOW_PASS

        'You can turn on and off the individual channels manually if you would like using this On property. However don't be confused.
        'The Enabled property is NOT manual. Instead Enabled is done by the CPU according to the read/write of register $4015.
        Square_Channel_1.TurnOn = True
        Square_Channel_2.TurnOn = True
        Triangle_Channel.TurnOn = True
        Noise_Channel.TurnOn = True
        DMC_Channel.TurnOn = True

        'Square 1
        Square_Channel_1.Enabled = True
        Square_Channel_1.Volume = 6
        Square_Channel_1.SamplingRate = Sound_Buffer_Wave_Format.SamplesPerSecond

        'Square 2
        Square_Channel_2.Enabled = True
        Square_Channel_2.Volume = 6
        Square_Channel_2.SamplingRate = Sound_Buffer_Wave_Format.SamplesPerSecond

        'Triangle
        Triangle_Channel.Enabled = True
        Triangle_Channel.Volume = 6
        Triangle_Channel.SamplingRate = Sound_Buffer_Wave_Format.SamplesPerSecond

        'Noise
        Noise_Channel.Enabled = True
        Noise_Channel.Volume = 6
        Noise_Channel.SamplingRate = Sound_Buffer_Wave_Format.SamplesPerSecond
        Noise_Channel.NoiseShiftData = 1

        'DMC
        DMC_Channel.Enabled = True
        DMC_Channel.Volume = 6
        DMC_Channel.SamplingRate = Sound_Buffer_Wave_Format.SamplesPerSecond
    End Sub
    Public Sub Sound_Channel_Clear_Buffer()
        For i As Integer = 0 To Buffer_Size - 1
            Sound_Data(i) = 0
        Next i
        Sound_Buffer.Write(0, Sound_Data, DirectSound.LockFlag.EntireBuffer)
    End Sub
    Public Sub Sound_Process_Through_Mixer(ByVal Cycles As Integer)
        Dim Write_Position As Integer
        Dim Buffer_Position As Integer
        Dim Prev_Sample As Integer
        Dim Next_Sample As Integer

        Sound_Channel_Update_Registers(Cycles)
        Write_Position = Sound_Buffer.WritePosition
        With APU
            If .FirstRender = True Then
                .FirstRender = False
                .DataPosition = Sound_Buffer.WritePosition + &H1000
                .LastPosition = Sound_Buffer.WritePosition
            End If
            Buffer_Position = Write_Position - .LastPosition
            If Buffer_Position < 0 Then Buffer_Position = (Buffer_Size - .LastPosition) + Write_Position
            If Buffer_Position <> 0 Then
                For i As Integer = 0 To Buffer_Position - 2 Step 2
                    Mixed_Sound = 0

                    If Square_Channel_1.Enabled And Square_Channel_1.TurnOn Then Mixed_Sound += Square_Channel_1_Render_Sample()
                    If Square_Channel_2.Enabled And Square_Channel_2.TurnOn Then Mixed_Sound += Square_Channel_2_Render_Sample()
                    If Triangle_Channel.Enabled And Triangle_Channel.TurnOn Then Mixed_Sound += Triangle_Channel_Render_Sample()
                    If Noise_Channel.Enabled And Noise_Channel.TurnOn Then Mixed_Sound += Noise_Channel_Render_Sample()
                    If DMC_Channel.Enabled And DMC_Channel.TurnOn Then Mixed_Sound += DMC_Channel_Render_Sample()

                    Mixed_Sound *= 7 'Increase the volume of the mixed sound 7 notches

                    If APU.Filter <> APU_FILTER_NONE Then
                        Next_Sample = Mixed_Sound
                        If APU.Filter = APU_FILTER_LOW_PASS Then
                            Mixed_Sound += Prev_Sample
                            Mixed_Sound \= 2
                        End If
                        Prev_Sample = Next_Sample
                    End If

                    Sound_Data(.DataPosition + 1) = (Mixed_Sound And &HFF00) >> 8 'Sometimes return an error. Why?
                    Sound_Data(.DataPosition) = Mixed_Sound And &HFF
                    .DataPosition = .DataPosition + 2
                    .DataPosition = .DataPosition Mod Buffer_Size
                Next i
                Sound_Buffer.Write(0, Sound_Data, DirectSound.LockFlag.EntireBuffer)
                .LastPosition = Write_Position
            End If
        End With
    End Sub
    Public Sub Sound_Channel_Update_Registers(ByVal Cycles As Long)
        Dim Hz As Long 'Hertz
        Dim Sequence As Long

        Hz = 60

        'Must execute 4 times per frame in NTSC mode to be a total of 240 hz
        While (Hz <= 240)
            APU.Frame_Counter = APU.Frame_Counter + 1
            If Frame_Counter_Mode = 0 Then
                Sequence = APU.Frame_Counter Mod 4 'Keep values in between 0 and 3 so it's sequenced 0, 1, 2, 3, 0, 1, 2, 3,....
            ElseIf Frame_Counter_Mode = 1 Then
                Sequence = APU.Frame_Counter Mod 5
            End If

            'mode 0:    mode 1:       function
            '---------  -----------  -----------------------------
            ' - - - f    - - - - -    IRQ (if bit 6 is clear)
            ' - l - l    l - l - -    Length counter and sweep
            ' e e e e    e e e e -    Envelope and linear counter
            If Frame_Counter_Mode = 0 Then 'NTSC
                Select Case Sequence
                    Case 0 '60 Hz
                        'Update Envelopes
                        Square_Channel_1_Update_Envelope()
                        Square_Channel_2_Update_Envelope()
                        Noise_Channel_Update_Envelope()

                        'Update Linear Counter
                        Triangle_Channel_Update_Linear_Counter()
                    Case 1 '120 Hz
                        'Update Sweeps
                        Square_Channel_1_Update_Sweep()
                        Square_Channel_2_Update_Sweep()

                        'Update Length Counters
                        Square_Channel_2_Decrement_Length_Counter()
                        Square_Channel_1_Decrement_Length_Counter()
                        Triangle_Channel_Update_Length_Counter()
                        Noise_Channel_Update_Length_Counter()

                        'Update Envelopes
                        Square_Channel_1_Update_Envelope()
                        Square_Channel_2_Update_Envelope()
                        Noise_Channel_Update_Envelope()

                        'Update Linear Counter
                        Triangle_Channel_Update_Linear_Counter()
                    Case 2 '180 Hz
                        'Update Envelopes
                        Square_Channel_1_Update_Envelope()
                        Square_Channel_2_Update_Envelope()
                        Noise_Channel_Update_Envelope()

                        'Update Linear Counter
                        Triangle_Channel_Update_Linear_Counter()
                    Case 3 '240 Hz
                        'Update Sweeps
                        Square_Channel_1_Update_Sweep()
                        Square_Channel_2_Update_Sweep()

                        'Update Length Counters
                        Square_Channel_2_Decrement_Length_Counter()
                        Square_Channel_1_Decrement_Length_Counter()
                        Triangle_Channel_Update_Length_Counter()
                        Noise_Channel_Update_Length_Counter()

                        'Update Envelopes
                        Square_Channel_1_Update_Envelope()
                        Square_Channel_2_Update_Envelope()
                        Noise_Channel_Update_Envelope()

                        'Update Linear Counter
                        Triangle_Channel_Update_Linear_Counter()

                        'IRQ??? Incomplete. Fill this in
                End Select
            ElseIf Frame_Counter_Mode = 1 Then 'PAL
                Select Case Sequence
                    Case 0
                        'Update Sweeps
                        Square_Channel_1_Update_Sweep()
                        Square_Channel_2_Update_Sweep()

                        'Update Length Counters
                        Square_Channel_2_Decrement_Length_Counter()
                        Square_Channel_1_Decrement_Length_Counter()
                        Triangle_Channel_Update_Length_Counter()
                        Noise_Channel_Update_Length_Counter()

                        'Update Envelopes
                        Square_Channel_1_Update_Envelope()
                        Square_Channel_2_Update_Envelope()
                        Noise_Channel_Update_Envelope()

                        'Update Linear Counter
                        Triangle_Channel_Update_Linear_Counter()
                    Case 1
                        'Update Envelopes
                        Square_Channel_1_Update_Envelope()
                        Square_Channel_2_Update_Envelope()
                        Noise_Channel_Update_Envelope()

                        'Update Linear Counter
                        Triangle_Channel_Update_Linear_Counter()
                    Case 2
                        'Update Sweeps
                        Square_Channel_1_Update_Sweep()
                        Square_Channel_2_Update_Sweep()

                        'Update Length Counters
                        Square_Channel_2_Decrement_Length_Counter()
                        Square_Channel_1_Decrement_Length_Counter()
                        Triangle_Channel_Update_Length_Counter()
                        Noise_Channel_Update_Length_Counter()

                        'Update Envelopes
                        Square_Channel_1_Update_Envelope()
                        Square_Channel_2_Update_Envelope()
                        Noise_Channel_Update_Envelope()

                        'Update Linear Counter
                        Triangle_Channel_Update_Linear_Counter()
                    Case 3
                        'Update Envelopes
                        Square_Channel_1_Update_Envelope()
                        Square_Channel_2_Update_Envelope()
                        Noise_Channel_Update_Envelope()

                        'Update Linear Counter
                        Triangle_Channel_Update_Linear_Counter()
                    Case 4
                        'Dummy. Nothing happens
                End Select
            End If
            Hz = Hz + 60
        End While
    End Sub
    Public Function Sound_Channel_Read_Status_Register() As Byte
        Dim Return_Value As Long

        If Square_Channel_1.Enabled Then Return_Value = Return_Value Or BIT_0
        If Square_Channel_2.Enabled Then Return_Value = Return_Value Or BIT_1
        If Triangle_Channel.Enabled Then Return_Value = Return_Value Or BIT_2
        If Noise_Channel.Enabled Then Return_Value = Return_Value Or BIT_3
        If DMC_Channel.Enabled Then Return_Value = Return_Value Or BIT_4

        Sound_Channel_Read_Status_Register = CByte(Return_Value)
    End Function
    Public Sub Sound_Channel_Write_Status_Register(ByVal Value As Byte)
        '$4015

        'Square 1
        If (Value And BIT_0) <> 0 Then
            Square_Channel_1.Enabled = True
        ElseIf (Value And BIT_0) = 0 Then
            Square_Channel_1.Enabled = False
            Square_Channel_1.Length_Counter = 0
        End If

        'Square 2
        If (Value And BIT_1) <> 0 Then
            Square_Channel_2.Enabled = True
        ElseIf (Value And BIT_1) = 0 Then
            Square_Channel_2.Enabled = False
            Square_Channel_2.Length_Counter = 0
        End If

        'Triangle
        If (Value And BIT_2) <> 0 Then
            Triangle_Channel.Enabled = True
        ElseIf (Value And BIT_2) = 0 Then
            Triangle_Channel.Enabled = False
            Triangle_Channel.Length_Counter = 0
            Triangle_Channel.LinearCounter = 0
        End If

        'Noise
        If (Value And BIT_3) <> 0 Then
            Noise_Channel.Enabled = True
        ElseIf (Value And BIT_3) = 0 Then
            Noise_Channel.Enabled = False
            Noise_Channel.Length_Counter = 0
        End If

        'DMC
        If (Value And BIT_4) <> 0 Then
            DMC_Channel.Enabled = True
        ElseIf (Value And BIT_4) = 0 Then
            DMC_Channel.Enabled = False
            DMC_Channel.Length_Counter = 0
        End If
    End Sub
    Public Sub Sound_Frame_Counter_Control_Write(ByVal Value As Long)
        Frame_Bits = Value And &HC0
        If (Value And BIT_7) <> 0 Then Frame_Counter_Mode = 0 Else Frame_Counter_Mode = 1
        If (Frame_Bits And BIT_6) <> 0 Then IRQ_Disable = True Else IRQ_Disable = False

        If Frame_Counter_Mode = 0 Then 'NTSC
            APU.Frame_Counter = 4 '4 for just the first value so it's sequenced 4, 0, 1, 2, 3, 0, 1, 2, 3, 0,....etc
        ElseIf Frame_Counter_Mode = 1 Then 'PAL
            APU.Frame_Counter = 0
        End If
    End Sub
End Module
