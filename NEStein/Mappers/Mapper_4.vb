Module Mapper_4
    '+==========+
    '| Mapper 4 |
    '+==========+

    'MMC3
    Private MMC3_Command As Byte
    Private MMC3_PrgAddr As Byte
    Public MMC3_ChrAddr As Integer
    Private MMC3_IrqVal As Byte
    Private MMC3_TmpVal As Byte
    Private MMC3_IrqOn As Boolean
    Private MMC3_Swap As Boolean
    Public PrgSwitch1 As Byte
    Public PrgSwitch2 As Byte
    Public Sub Mapper4_Reset()
        MMC3_Command = 0
        MMC3_PrgAddr = 0
        MMC3_ChrAddr = 0
        MMC3_IrqVal = 0
        MMC3_TmpVal = 0
        MMC3_IrqOn = False
        MMC3_Swap = False
        PrgSwitch1 = 0
        PrgSwitch2 = 0
        DefaultReset()
        Mapper4_Sync()
    End Sub
    Public Sub Mapper4_Write(ByVal Address As Integer, ByVal Value As Byte)
        Select Case Address
            Case &H8000
                MMC3_Command = Value And &H7
                If Value And &H80 Then MMC3_ChrAddr = &H1000& Else MMC3_ChrAddr = 0
                If Value And &H40 Then MMC3_Swap = True Else MMC3_Swap = False
            Case &H8001
                Select Case MMC3_Command
                    Case 0 : Select1KVROM(Value, 0) : Select1KVROM(Value + 1, 1)
                    Case 1 : Select1KVROM(Value, 2) : Select1KVROM(Value + 1, 3)
                    Case 2 : Select1KVROM(Value, 4)
                    Case 3 : Select1KVROM(Value, 5)
                    Case 4 : Select1KVROM(Value, 6)
                    Case 5 : Select1KVROM(Value, 7)
                    Case 6 : PrgSwitch1 = Value : Mapper4_Sync()
                    Case 7 : PrgSwitch2 = Value : Mapper4_Sync()
                End Select
            Case &HA000
                If Not ROMHeader.FourScreen Then
                    If Value And &H1 Then
                        ROMHeader.Mirroring = 0
                    Else
                        ROMHeader.Mirroring = 1
                    End If
                    Do_Mirroring()
                End If
            Case &HA001 : If Value Then ROMHeader.SRAM = True Else ROMHeader.SRAM = False
            Case &HC000 : MMC3_IrqVal = Value
            Case &HC001 : MMC3_TmpVal = Value
            Case &HE000 : MMC3_IrqOn = False : MMC3_IrqVal = MMC3_TmpVal
            Case &HE001 : MMC3_IrqOn = True
        End Select
    End Sub
    Public Sub Mapper4_HBlank(ByVal ScanLine As Integer, ByVal Two As Byte)
        If ScanLine = 0 Then
            MMC3_IrqVal = MMC3_TmpVal
        ElseIf ScanLine > 239 Then
            Exit Sub
        ElseIf MMC3_IrqOn And (Two And &H18) Then
            MMC3_IrqVal = (MMC3_IrqVal - 1) And &HFF
            If (MMC3_IrqVal = 0) Then
                CPU_IRQ()
                MMC3_IrqVal = MMC3_TmpVal
            End If
        End If
    End Sub
    Public Sub Mapper4_Sync()
        If MMC3_Swap Then
            Select32KPROM(&HFE, PrgSwitch2, PrgSwitch1, &HFF)
        Else
            Select32KPROM(PrgSwitch1, PrgSwitch2, &HFE, &HFF)
        End If
    End Sub
End Module
