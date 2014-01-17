Module Mapper_1
    '+==========+
    '| Mapper 1 |
    '+==========+

    'MMC1
    Private MMC1_Accumulator, MMC1_Sequence As Integer
    Private MMC1_Register(4) As Byte
    Public Sub Mapper1_Reset()
        Array.Clear(MMC1_Register, 0, MMC1_Register.Length)
        MMC1_Register(0) = &H1F
        MMC1_Sequence = 0
        MMC1_Accumulator = 0
        DefaultReset()
    End Sub
    Public Sub Mapper1_Write(ByVal Address As Integer, ByVal Value As Byte)
        Dim Bank_Select As Integer

        If Value And &H80 Then
            MMC1_Register(0) = MMC1_Register(0) Or &HC
            MMC1_Accumulator = MMC1_Register((Address \ &H2000) And 3)
            MMC1_Sequence = 5
        Else
            If Value And 1 Then
                MMC1_Accumulator = MMC1_Accumulator Or (1 << MMC1_Sequence)
            End If
            MMC1_Sequence += 1
        End If

        If MMC1_Sequence = 5 Then
            MMC1_Register(Address \ &H2000 And 3) = MMC1_Accumulator
            MMC1_Sequence = 0
            MMC1_Accumulator = 0

            If (ROMHeader.PrgSize = &H20) Then '/* 512k cart */'
                Bank_Select = (MMC1_Register(1) And &H10) * 2
            Else '/* other carts */'
                Bank_Select = 0
            End If

            If MMC1_Register(0) And 2 Then 'enable panning
                ROMHeader.Mirroring = (MMC1_Register(0) And 1) Xor 1
            Else 'disable panning
                ROMHeader.Mirroring = 2
            End If
            Do_Mirroring()

            If (MMC1_Register(0) And 8) = 0 Then
                Select32KPROM(4 * (MMC1_Register(3) And 15) + Bank_Select)
            ElseIf (MMC1_Register(0) And 4) Then '16k
                Select16KPROM(&HFE, 6)
                Select16KPROM(((MMC1_Register(3) And 15) * 2) + Bank_Select, 4)
            Else '32k
                Select16KPROM(0, 4)
                Select16KPROM(((MMC1_Register(3) And 15) * 2) + Bank_Select, 6)
            End If

            If (MMC1_Register(0) And &H10) Then '4k
                Select4KVROM(MMC1_Register(1), 0)
                Select4KVROM(MMC1_Register(2), 1)
            Else '8k
                Select8KVROM(MMC1_Register(1) \ 2)
            End If
        End If
    End Sub
End Module
