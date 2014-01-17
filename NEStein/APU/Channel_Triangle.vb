Option Explicit On
Module Channel_Triangle


    Public Function Triangle_Channel_Render_Sample() As Long
        With Triangle_Channel
            If (.Length_Counter > 0 And .LinearCounter > 0) And .Wavelength > 0 Then
                '.SampleCount = Within_Range(.SampleCount + 1, 0, LONG_RANGE) 'causes minor slowdown
                .SampleCount = .SampleCount + 1
                If .SampleCount > .RenderedWavelength Then
                    .SampleCount = .SampleCount - .RenderedWavelength
                    .Triangle_Sequence = .Triangle_Sequence + 1
                End If

                Triangle_Channel_Render_Sample = .Triangle_Sequence_Data(.Triangle_Sequence And 31) * &H40
                Exit Function
            End If
            Triangle_Channel_Render_Sample = 0
            Exit Function
        End With
    End Function

    Public Sub Triangle_Channel_Update_Frequency()
        With Triangle_Channel
            .Frequency = NES_NTSC_CPU_SPEED_MHZ / ((.Wavelength + 1))
            '.Frequency = TEST_SPEED / (.Wavelength + 1)
            .RenderedWavelength = .SamplingRate / .Frequency
        End With
    End Sub

    Public Sub Triangle_Channel_Update_Linear_Counter()
        With Triangle_Channel
            If .Linear_Counter_Disable = True Then
                .LinearCounter = .LinearCounterLoad
            ElseIf .LinearCounter > 0 Then
                .LinearCounter = .LinearCounter - 1
            End If

            If .Length_Counter_Disable = False Then .Linear_Counter_Disable = False
            If .LinearCounter <= 0 Then .LinearCounter = 0
        End With
    End Sub

    Public Sub Triangle_Channel_Update_Length_Counter()
        With Triangle_Channel
            If .Length_Counter > 0 And .Length_Counter_Disable = False Then .Length_Counter = .Length_Counter - 1

            'The length counter and soundoutput are stopped after reaching 0
            If (.Length_Counter <= 0) Then
                .Length_Counter = 0
                .Wavelength = 0
            End If
        End With
    End Sub

    Public Sub Triangle_Channel_Write_Register_1(ByVal Value As Byte)
        '$4008
        With Triangle_Channel
            .Length_Counter_Disable = (Value And BIT_7) <> 0 'length counter clock disable / linear counter start
            'If .Length_Counter > 0 Then
            .LinearCounterLoad = Value And &H7F '(BIT_0 + BIT_1 + BIT_2 + BIT_3 + BIT_4 + BIT_5 + BIT_6) 'linear counter load register
            'End If
        End With
    End Sub

    Public Sub Triangle_Channel_Write_Register_2(ByVal Value As Byte)
        '$4009
        With Triangle_Channel
            'Unused
        End With
    End Sub

    Public Sub Triangle_Channel_Write_Register_3(ByVal Value As Byte)
        '$400A
        With Triangle_Channel
            .Wavelength = (.Wavelength And &H700)
            .Wavelength = .Wavelength Or Value
            Triangle_Channel_Update_Frequency()
        End With
    End Sub

    Public Sub Triangle_Channel_Write_Register_4(ByVal Value As Byte)
        '$400B
        With Triangle_Channel
            .Wavelength = .Wavelength And &HFF
            .Wavelength = .Wavelength Or ((Value And (BIT_0 + BIT_1 + BIT_2)) * 256)
            Triangle_Channel_Update_Frequency()
            If .Enabled = True Then .Length_Counter = APU.Length_Values((Value And (BIT_3 + BIT_4 + BIT_5 + BIT_6 + BIT_7)) \ 8)
            .LinearCounter = .LinearCounterLoad
            .Linear_Counter_Disable = True
        End With
    End Sub


End Module
