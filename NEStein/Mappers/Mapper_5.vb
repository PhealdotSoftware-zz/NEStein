Module Mapper_5
    '+==========+
    '| Mapper 5 |
    '+==========+

    'MMC5
    'Has some weird Mirroring and Name Tables that I don't understand yet.

    'PRG/CHR Sizes
    Private MMC5_PrgSize As Byte
    Private MMC5_ChrSize As Byte
    Private MMC5_GfxMode As Byte

    'CHR Pages
    Private MMC5_ChrPageSP(7) As Byte
    Private MMC5_ChrPageBG(3) As Byte

    'IRQ
    Private MMC5_IrqClear As Byte
    Private MMC5_IrqScanline As Integer
    Private MMC5_IrqLine As Integer
    Private MMC5_IrqStatus As Byte
    Private MMC5_IrqEnable As Byte
    Public Sub Mapper5_Reset()
        MMC5_PrgSize = 3
        MMC5_ChrSize = 3
        MMC5_GfxMode = 0
        Array.Clear(MMC5_ChrPageSP, 0, MMC5_ChrPageSP.Length)
        Array.Clear(MMC5_ChrPageBG, 0, MMC5_ChrPageBG.Length)
        MMC5_IrqClear = 0
        MMC5_IrqScanline = 0
        MMC5_IrqLine = 0
        MMC5_IrqStatus = 0
        MMC5_IrqEnable = 0
        Select32KPROM(&HFC)
        Select8KVROM(0)
    End Sub
    Public Sub Mapper5_WriteLow(ByVal Address As Integer, ByVal Value As Byte)
        Select Case Address
            Case &H5100 : MMC5_PrgSize = Value And &H3
            Case &H5101 : MMC5_ChrSize = Value And &H3
            Case &H5104 : MMC5_GfxMode = Value And &H3
            Case &H5105
                'Here I try to implement the weird MMC5 Mirroring
                Select Case Value
                    Case 0 : ROMHeader.Mirroring = MIRROR_ONESCREEN_LOW
                    Case &H44 : ROMHeader.Mirroring = MIRROR_VERTICAL
                    Case &H50 : ROMHeader.Mirroring = MIRROR_HORIZONTAL
                    Case &H55 : ROMHeader.Mirroring = MIRROR_ONESCREEN_HIGH
                    Case &HD8 '3120 DBCA
                End Select
                Do_Mirroring()
            Case &H5114 To &H5117
                If Value And &H80 Then
                    Select Case Address And &H7
                        Case 4
                            If MMC5_PrgSize = 3 Then
                                Select8KPROM(Value And &H7F, 4)
                            End If
                        Case 5
                            If MMC5_PrgSize = 1 Or MMC5_PrgSize = 2 Then
                                Select16KPROM(Value And &H7F, 4)
                            ElseIf MMC5_PrgSize = 3 Then
                                Select8KPROM(Value And &H7F, 5)
                            End If
                        Case 6
                            If MMC5_PrgSize = 2 Or MMC5_PrgSize = 3 Then
                                Select8KPROM(Value And &H7F, 6)
                            End If
                        Case 7
                            If MMC5_PrgSize = 0 Then
                                Select32KPROM((Value And &H7F) >> 2)
                            ElseIf MMC5_PrgSize = 1 Then
                                Select16KPROM((Value And &H7F) >> 1, 6)
                            ElseIf MMC5_PrgSize = 2 Or MMC5_PrgSize = 3 Then
                                Select8KPROM(Value And &H7F, 7)
                            End If
                    End Select
                Else
                    'WRAM Bank
                End If
            Case &H5120 To &H5127
                'SP
                MMC5_ChrPageSP(Address And &H7) = Value

                Select Case MMC5_ChrSize
                    Case 0
                        Select8KVROM(MMC5_ChrPageSP(7))
                    Case 1
                        Select4KVROM(MMC5_ChrPageSP(3), 0)
                        Select4KVROM(MMC5_ChrPageSP(7), 4)
                    Case 2
                        Select2KVROM(MMC5_ChrPageSP(1), 0)
                        Select2KVROM(MMC5_ChrPageSP(3), 2)
                        Select2KVROM(MMC5_ChrPageSP(5), 4)
                        Select2KVROM(MMC5_ChrPageSP(7), 6)
                    Case 3
                        Select1KVROM(MMC5_ChrPageSP(0), 0)
                        Select1KVROM(MMC5_ChrPageSP(1), 1)
                        Select1KVROM(MMC5_ChrPageSP(2), 2)
                        Select1KVROM(MMC5_ChrPageSP(3), 3)
                        Select1KVROM(MMC5_ChrPageSP(4), 4)
                        Select1KVROM(MMC5_ChrPageSP(5), 5)
                        Select1KVROM(MMC5_ChrPageSP(7), 7)
                End Select
            Case &H5128 To &H512B
                'BG
                MMC5_ChrPageBG(Address And &H3) = Value

                Select Case MMC5_ChrSize
                    Case 1
                        Select8KVROM(MMC5_ChrPageBG(3))
                    Case 3
                        Select1KVROM(MMC5_ChrPageBG(0), 4)
                        Select1KVROM(MMC5_ChrPageBG(1), 5)
                        Select1KVROM(MMC5_ChrPageBG(2), 6)
                        Select1KVROM(MMC5_ChrPageBG(3), 7)
                End Select
            Case &H5203
                MMC5_IrqLine = Value
            Case &H5204
                MMC5_IrqEnable = Value
            Case Else
                If Address >= &H5000 And Address <= &H5015 Then
                    'APU ExWrite
                ElseIf Address >= &H5C00 And Address <= &H5FFF Then
                    If MMC5_GfxMode = 2 Then 'ExRAM
                        VRAM(&H800 + (Address And &H3FF)) = Value
                    ElseIf MMC5_GfxMode <> 3 Then 'Split,ExGraphic
                        If MMC5_IrqStatus And &H40 Then
                            VRAM(&H800 + (Address And &H3FF)) = Value
                        Else
                            VRAM(&H800 + (Address And &H3FF)) = 0
                        End If
                    End If
                End If
        End Select
    End Sub
    Public Function Mapper5_ReadLow(ByVal Address As Integer)
        Dim Value As Byte = Address >> 8

        Select Case Address
            Case &H5204
                Value = MMC5_IrqStatus
                MMC5_IrqStatus = 0
                MMC5_IrqStatus = MMC5_IrqStatus And &H80
        End Select
        If Address >= &H5C00 And Address <= &H5FFF Then
            Value = VRAM(&H800 + (Address And &H3FF))
        End If

        Return Value
    End Function
    Public Sub Mapper5_HBlank(ByVal ScanLine As Integer)
        If ScanLine < 240 Then
            MMC5_IrqScanline += 1
            MMC5_IrqStatus = MMC5_IrqStatus Or &H40
            MMC5_IrqClear = 0
        End If

        If MMC5_IrqScanline = MMC5_IrqLine Then
            MMC5_IrqStatus = MMC5_IrqStatus Or &H80
        End If
        MMC5_IrqClear += 1
        If MMC5_IrqClear > 2 Then
            MMC5_IrqScanline = 0
            MMC5_IrqStatus = MMC5_IrqStatus And Not &H80
            MMC5_IrqStatus = MMC5_IrqStatus And Not &H40
        End If

        If (MMC5_IrqEnable And &H80) And (MMC5_IrqStatus And &H80) Then
            CPU_IRQ()
        End If
    End Sub
End Module
