Option Explicit On
Module Channel_Noise
    Public Function Noise_Channel_Render_Sample() As Long
        Dim Bit0 As Long
        Dim Bit1 As Long
        Dim Bit6 As Long

        With Noise_Channel
            If .Length_Counter > 0 Then
                .SampleCount += 1
                If .SampleCount > .RenderedWavelength Then
                    .SampleCount = .SampleCount - .RenderedWavelength
                    If .Noise_Random_Generator_Mode = True Then '93 bits long
                        Bit0 = .NoiseShiftData >> 14
                        Bit6 = .NoiseShiftData >> 8
                        .Noise_Feedback = Bit0 Xor Bit6
                    Else
                        Bit0 = .NoiseShiftData >> 14
                        Bit1 = .NoiseShiftData >> 13
                        .Noise_Feedback = Bit0 Xor Bit1
                    End If

                    'Step 2) The shift register is shifted right by one bit.
                    .NoiseShiftData = .NoiseShiftData << 1
                    'Step 3) Bit 14, the leftmost bit, is "set" to the feedback calculated earlier.
                    '        Note: You set it using the Or operator, not the And operator like you would see in most emulators' source code
                    .Noise_Feedback = .Noise_Feedback Or BIT_14

                    'Then the Noise shift gets updated and prepared to be outputed for sound.
                    .NoiseShiftData = .NoiseShiftData Or (.Noise_Feedback And &H1&)
                End If

                If .Envelope_Decay_Disable Then
                    Return ((.NoiseShiftData And 1) * &H20) * .Volume
                Else
                    Return ((.NoiseShiftData And 1) * &H20) * .Envelope_Counter
                End If
            Else
                .Volume = 0
            End If
            Return 0
        End With
    End Function
    Public Sub Noise_Channel_Update_Envelope()
        With Noise_Channel
            If .Envelope_Start_Flag = True Then
                .Envelope_Start_Flag = False
                .Envelope = .Volume + 1
                .Envelope_Counter = 15
            Else
                .Envelope = .Envelope - 1
            End If

            If .Envelope <= 0 Then
                .Envelope = .Volume + 1
                If .Envelope_Counter > 0 Then
                    .Envelope_Counter = .Envelope_Counter - 1
                ElseIf .Length_Counter_Disable And .Envelope_Counter <= 0 Then
                    .Envelope_Counter = 15
                End If
            End If
        End With
    End Sub
    Public Sub Noise_Channel_Update_Frequency()
        With Noise_Channel
            .Frequency = NES_NTSC_CPU_SPEED_MHZ / ((.Wavelength + 1))
            .RenderedWavelength = .SamplingRate / .Frequency
        End With
    End Sub
    Public Sub Noise_Channel_Update_Length_Counter()
        With Noise_Channel
            If .Length_Counter_Disable = False Then .Length_Counter = .Length_Counter - 1
            If .Length_Counter <= 0 Then .Length_Counter = 0
        End With
    End Sub
    Public Sub Noise_Channel_Write_Register_1(ByVal Value As Byte)
        With Noise_Channel
            .Volume = Value And &HF
            .Envelope_Counter = Value And &HF
            .Envelope_Decay_Disable = (Value And BIT_4) <> 0
            .Length_Counter_Disable = (Value And BIT_5) <> 0 
        End With
    End Sub
    Public Sub Noise_Channel_Write_Register_2(ByVal Value As Byte)
        With Noise_Channel
            .Wavelength = .NoiseWavelengths(Value And &HF)
            Noise_Channel_Update_Frequency()
            .Noise_Random_Generator_Mode = (Value And BIT_7) <> 0
        End With
    End Sub
    Public Sub Noise_Channel_Write_Register_3(ByVal Value As Byte)
        With Noise_Channel
            If .Enabled = True Then .Length_Counter = APU.Length_Values((Value And &HF8&) \ 8)
        End With
    End Sub
End Module
