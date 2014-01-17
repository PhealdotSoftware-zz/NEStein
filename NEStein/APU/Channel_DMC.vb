Option Explicit On
Module Channel_DMC
    Public Function DMC_Channel_Render_Sample() As Long
        With DMC_Channel
            .SampleCount += 1
            If .SampleCount > .RenderedWavelength Then
                .SampleCount -= .RenderedWavelength
                If .SampleLength > 0 And .Shift > 0 Then
                    Dim Temp As Integer = .SampleAddress

                    If .SampleAddress = &HFFFF Then
                        .SampleAddress = &H8000
                    Else
                        .SampleAddress += 1
                    End If
                    .DAC = ReadMemory(Temp)
                    .SampleLength -= 1
                    .Shift = 8

                    If .DMC_Loop And (.SampleLength <= 0) Then
                        .SampleLength = .InitialLength
                        .SampleAddress = .InitialAddress
                    End If
                End If

                If .SampleLength > 0 Then
                    If .DAC <> 0 Then
                        If (.DAC And 1) = 0 And .DAC_Counter > 1 Then
                            .DAC_Counter -= 2
                        ElseIf (.DAC And 1) <> 0 And .DAC_Counter < &H7E Then
                            .DAC_Counter += 2
                        End If
                    End If

                    .DAC_Counter -= 1

                    If .DAC_Counter <= 0 Then
                        .DAC_Counter = 8
                    End If

                    .DAC = .DAC >> 1
                    .Shift -= 1
                End If
            End If
            Return .DAC_Counter * &H30
        End With
    End Function
    Public Sub DMC_Channel_Update_Frequency()
        With DMC_Channel
            .Frequency = NES_NTSC_CPU_SPEED_MHZ / (.Wavelength + 1)
            .RenderedWavelength = .SamplingRate / .Frequency
        End With
    End Sub
    Public Sub DMC_Channel_Write_Register_1(ByVal Value As Byte)
        With DMC_Channel
            .DMC_IrqOn = (Value And &H80)
            .DMC_Loop = (Value And &H40)
            .Wavelength = .DMCWavelengths(Value And &HF)
        End With
        DMC_Channel_Update_Frequency()
    End Sub
    Public Sub DMC_Channel_Write_Register_2(ByVal Value As Byte)
        With DMC_Channel
            .DAC = Value And &H7F
            .Shift = 8
        End With
    End Sub
    Public Sub DMC_Channel_Write_Register_3(ByVal Value As Byte)
        With DMC_Channel
            .SampleAddress = (Value * &H40) + &HC000
            .InitialAddress = .SampleAddress
        End With
    End Sub
    Public Sub DMC_Channel_Write_Register_4(ByVal Value As Byte)
        With DMC_Channel
            .SampleLength = (Value * &H10) + 1
            .InitialLength = .SampleLength
        End With
    End Sub
End Module
