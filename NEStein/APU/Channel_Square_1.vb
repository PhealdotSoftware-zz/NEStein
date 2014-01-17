Option Explicit On
Module Channel_Square_1


    Public Function Square_Channel_1_Render_Sample() As Long
        With Square_Channel_1
            If .Length_Counter > 0 Then
                '.SampleCount = Within_Range(.SampleCount + 1, 0, LONG_RANGE) 'causes minor slowdown
                .SampleCount = .SampleCount + 1
                If .WaveStatus = True And (.SampleCount > (.RenderedWavelength * .DutyPercentage)) Then
                    .SampleCount = .SampleCount - .RenderedWavelength * .DutyPercentage
                    .WaveStatus = Not .WaveStatus
                ElseIf .WaveStatus = False And (.SampleCount > (.RenderedWavelength * (1 - .DutyPercentage))) Then
                    .SampleCount = .SampleCount - .RenderedWavelength * (1 - .DutyPercentage)
                    .WaveStatus = Not .WaveStatus
                End If

                If .WaveStatus = True Then
                    Square_Channel_1_Render_Sample = 0
                    Exit Function
                End If

                If .Envelope_Decay_Disable = False Then
                    Square_Channel_1_Render_Sample = &H40 * .Envelope
                    Exit Function
                End If

                Square_Channel_1_Render_Sample = &H40 * .Volume
                Exit Function
            Else
                '.SampleCount = 0
            End If
            Square_Channel_1_Render_Sample = 0
            Exit Function
        End With
    End Function

    Public Sub Square_Channel_1_Update_Frequency()
        With Square_Channel_1
            'From apu_ref.txt
            '   F = 1.79 MHz / (N + 1) / 16 for Rectangle channels
            'However over in http://wiki.nesdev.com/w/index.php/APU_Misc
            '   f = CPU / (16 * (t + 1))
            .Frequency = NES_NTSC_CPU_SPEED_MHZ / ((.Wavelength + 1) * 16)
            '.Frequency = TEST_SPEED / ((.Wavelength + 1) * &H10)
            .RenderedWavelength = .SamplingRate / .Frequency
        End With
    End Sub

    Public Sub Square_Channel_1_Update_Sweep()
        Dim New_Wave_Length As Long
        Dim Current_Wavelength As Long

        With Square_Channel_1
            If .Sweep_Enabled Then .Sweep_Counter = .Sweep_Counter - 1

            If .Sweep_Counter <= 0 Then
                .Sweep_Counter = .Sweep_Refresh_Rate
                If .Wavelength >= 8 Then
                    Current_Wavelength = .Wavelength
                    If .Sweep_Negate Then 'If Sweep Directions Decrese only
                        '(For Channel 1 Decrese only: Minus an additional 1)
                        New_Wave_Length = Current_Wavelength - (Current_Wavelength >> .Sweep_Right_Shift) - 1
                    Else 'If Sweep Directions Increase only
                        New_Wave_Length = Current_Wavelength + (Current_Wavelength >> .Sweep_Right_Shift)
                    End If
                    'Sweep End check
                    If (Current_Wavelength < &H8&) And (New_Wave_Length > &H7FF&) Then 'The channel is silenced and the sweep clock is halted
                        .Volume = 0
                        .Sweep_Enabled = False
                    Else
                        If .Sweep_Enabled And .Sweep_Right_Shift <> 0 And .Length_Counter > 0 Then  'If all 3 conditions are met, update the wavelength register
                            'Wavelength register will be updated only if all 3 conditions are met:
                            '- Bit 7 is set (Sweeping Enabled)
                            '- The shift value (which is S in the formula) does not equal 0
                            '- The channels length counter contains a non zero value
                            .Wavelength = New_Wave_Length
                            Square_Channel_1_Update_Frequency()
                        End If
                    End If
                End If
            End If
        End With
    End Sub

    Public Sub Square_Channel_1_Update_Envelope()
        With Square_Channel_1
            .Envelope_Counter = .Envelope_Counter - 1
            If .Envelope_Counter < 0 Then .Envelope_Counter = 0
            If Not .Envelope_Decay_Disable And .Envelope > 0 And .Envelope_Counter = 0 Then
                .Envelope = .Envelope - 1
                .Envelope_Counter = .Volume
            End If
            If .Length_Counter_Disable And .Envelope = 0 Then .Envelope = 15
        End With
    End Sub

    Public Sub Square_Channel_1_Decrement_Length_Counter()
        With Square_Channel_1
            If Not .Length_Counter_Disable Then .Length_Counter = .Length_Counter - 1
            If (.Length_Counter <= 0) Then .Length_Counter = 0
        End With
    End Sub

    Public Sub Square_Channel_1_Write_Register_1(ByVal Value As Byte)
        '$4000
        '          Bits: 7654 3210
        '$4000 / $4004   DDLC VVVV   Duty (D), envelope loop / length counter disable (L), constant volume (C), volume/envelope (V)
        With Square_Channel_1
            .Volume = Value And &HF& 'Where (BIT_0 + BIT_1 + BIT_2 + BIT_3) = Hex &HF&, Dec 15
            .Envelope_Counter = Value And &HF& 'Where (BIT_0 + BIT_1 + BIT_2 + BIT_3) = Hex &HF&, Dec 15
            .Envelope_Counter = .Envelope_Counter + 1
            .Envelope_Decay_Disable = (Value And BIT_4) <> 0 'False = Envelope Decay / True = Constant Volume
            .Length_Counter_Disable = (Value And BIT_5) <> 0 'False = Envelope loop / True = Length Counter Disable
            .Duty = (Value And &HC0&) \ 64 'Where (BIT_6 + BIT_7) = Hex &C0&, Dec 192
            Select Case .Duty 'This is the width of the rectangular pulses
                Case 0 : .DutyPercentage = 0.125 '12.5%
                Case 1 : .DutyPercentage = 0.25  '25%
                Case 2 : .DutyPercentage = 0.5   '50%
                Case 3 : .DutyPercentage = 0.75 '25% negated (according to the apu_ref.txt)
                    '-25% fails miserably on double dragon. 75% sounds perfect
            End Select
        End With
    End Sub

    Public Sub Square_Channel_1_Write_Register_2(ByVal Value As Byte)
        '$4001

        '          Bits: 7654 3210
        '$4001 / $4005   EPPP NSSS   E = Sweep Enabled
        '                            P = Sweep Period
        '                            N = Sweep Negate
        '                            S = Sweep Shift
        With Square_Channel_1
            .Sweep_Right_Shift = Value And &H7& ''where (BIT_0 + BIT_1 + BIT_2) = Hex &H7&, Dec 7
            .Sweep_Negate = (Value And BIT_3) <> 0 'Sweep Direction (Negate -/+)
            .Sweep_Refresh_Rate = ((Value And &H70&) \ 16) + 1 'Sweep Speed (Period). The refresh rate frequency is 120Hz/(N+1),
            'where N is the value written, between 0 and 7.
            'where (BIT_4 + BIT_5 + BIT_6) = Hex &H70&, Dec 112
            .Sweep_Counter = .Sweep_Refresh_Rate
            .Sweep_Enabled = (Value And BIT_7) <> 0
        End With
    End Sub

    Public Sub Square_Channel_1_Write_Register_3(ByVal Value As Byte)
        '$4002

        '          Bits: 7654 3210
        '$4002 / $4006   TTTT TTTT   Timer low (T)
        With Square_Channel_1
            .Wavelength = (.Wavelength And &H700) Or Value
            Square_Channel_1_Update_Frequency()
        End With
    End Sub

    Public Sub Square_Channel_1_Write_Register_4(ByVal Value As Byte)
        '$4003

        '          Bits: 7654 3210
        '$4003 / $4007   LLLL LTTT   Length counter load (L), timer high (T)

        'Writing to $4003/4007 reloads the length counter, restarts the envelope, and resets the phase of the pulse generator.
        'Because it resets phase, vibrato should only write the low timer register to avoid a phase reset click. At some pitches,
        'particularly near A440, wide vibrato should normally be avoided (e.g. this flaw is heard throughout the Mega Man 2 ending).
        With Square_Channel_1
            .Wavelength = (.Wavelength And &HFF) Or ((Value And &H7&) * 256) 'where (BIT_0 + BIT_1 + BIT_2) = Hex &H7&, Dec 7
            Square_Channel_1_Update_Frequency()
            If .Enabled = True Then .Length_Counter = APU.Length_Values((Value And &HF8&) \ 8) 'where (BIT_3 + BIT_4 + BIT_5 + BIT_6 + BIT_7) = Hex &HF8&, Dec 248
            If .Envelope_Decay_Disable = False Then .Envelope = 15
        End With
    End Sub



End Module
