Option Explicit On
Module MMU
    'Registers
    Private Register_8000 As Byte
    Private Register_A000 As Byte
    Private Register_C000 As Byte
    Private Register_E000 As Byte

    'Banks
    Public Bank0(&H7FF) As Byte 'RAM
    Public Bank6(&H1FFF) As Byte 'SaveRAM
    Public Bank8(&H1FFF) As Byte '*PRG-ROM (8-E)
    Public BankA(&H1FFF) As Byte
    Public BankC(&H1FFF) As Byte
    Public BankE(&H1FFF) As Byte

    'Input
    Public Controller1_Count, Controller2_Count As Integer

#Region "Read Memory"
    Public Function ReadMemory(Address As Integer) As Byte
        Select Case Address
            Case &H0 To &H1FFF : Return Bank0(Address And &H7FF)
            Case &H2000 To &H2007 : Return PPU_Read(Address)
            Case &H4016
                Dim Temp As Byte = Controller_1(Controller1_Count)
                Controller1_Count = (Controller1_Count + 1) And 7
                Return Temp
            Case &H4017
                Dim Temp As Byte = Controller_2(Controller2_Count)
                Controller2_Count = (Controller2_Count + 1) And 7
                Return Temp
            Case &H5000 To &H5FFF : Return Mapper_ReadLow(Address)
            Case &H6000 To &H7FFF : Return Bank6(Address And &H1FFF)
            Case &H8000 To &H9FFF : Return Bank8(Address And &H1FFF)
            Case &HA000 To &HBFFF : Return BankA(Address And &H1FFF)
            Case &HC000 To &HDFFF : Return BankC(Address And &H1FFF)
            Case &HE000 To &HFFFF : Return BankE(Address And &H1FFF)
        End Select

        Return Nothing
    End Function
    Public Function ReadMemory16(Address As Integer) As Integer
        Return ReadMemory(Address) + (ReadMemory(Address + 1) * &H100)
    End Function
    Public Function ReadMemory16_ZP(Address As Integer) As Integer
        Return ReadMemory(Address And &HFF) + (ReadMemory((Address + 1) And &HFF) * &H100)
    End Function
#End Region

#Region "Write Memory"
    Public Sub WriteMemory(Address As Integer, Value As Byte)
        Select Case Address
            Case &H0 To &H1FFF : Bank0(Address And &H7FF) = Value
            Case &H2000 To &H2007 : PPU_Write(Address, Value)
            Case &H4000 To &H4013 : Sound_Write(Address, Value)
            Case &H4014 : Buffer.BlockCopy(Bank0, (Value * &H100) And &H7FF, SpriteRAM, 0, &H100)
            Case &H4015 : Sound_Channel_Write_Status_Register(Value)
            Case &H5000 To &H5FFF : Mapper_WriteLow(Address, Value)
            Case &H6000 To &H7FFF : Bank6(Address And &H1FFF) = Value
            Case &H8000 To &HFFFF : Mapper_Write(Address, Value)
        End Select
    End Sub
#End Region

#Region "Copy Banks (PRG/CHR)"
    Public Sub SetupBanks()
        Register_8000 = MaskBankAddress(Register_8000)
        Register_A000 = MaskBankAddress(Register_A000)
        Register_C000 = MaskBankAddress(Register_C000)
        Register_E000 = MaskBankAddress(Register_E000)

        Buffer.BlockCopy(PROM, Register_8000 * &H2000, Bank8, 0, &H2000)
        Buffer.BlockCopy(PROM, Register_A000 * &H2000, BankA, 0, &H2000)
        Buffer.BlockCopy(PROM, Register_C000 * &H2000, BankC, 0, &H2000)
        Buffer.BlockCopy(PROM, Register_E000 * &H2000, BankE, 0, &H2000)
    End Sub
    Private Sub CopyBanks(Dest, Src, Count)
        If ROMHeader.ChrSize = 0 Then Exit Sub 'No VROM cart
        If ROMHeader.Mapper = 4 Then 'MMC3
            For i As Integer = 0 To Count - 1
                Buffer.BlockCopy(VROM, (Src + i) * &H400, VRAM, MMC3_ChrAddr Xor (Dest + i) * &H400, &H400)
            Next i
        Else
            Buffer.BlockCopy(VROM, Src * &H400, VRAM, Dest * &H400, Count * &H400)
        End If
    End Sub
#End Region

#Region "Mask Banks (PRG/CHR)"
    Public Function MaskBankAddress(Bank As Byte)
        If Bank >= ROMHeader.PrgSize * 2 Then
            Dim i As Byte = &HFF
            Do While (Bank And i) >= ROMHeader.PrgSize * 2
                i = i \ 2
            Loop
            MaskBankAddress = (Bank And i)
        Else
            MaskBankAddress = Bank
        End If
    End Function
    Public Function MaskVROM(Page As Byte, Mask As Long) As Byte
        Dim i As Integer
        If Mask = 0 Then Mask = 256
        If Mask And Mask - 1 Then
            i = 1
            Do While i < Mask
                i += 1
            Loop
        Else
            i = Mask
        End If
        i = Page And (i - 1)
        If i >= Mask Then i = Mask - 1

        Return i
    End Function
#End Region

#Region "Select PROM"
    Public Sub Select8KPROM(Bank As Byte, Page As Byte)
        Select Case Page
            Case 4 : Register_8000 = Bank
            Case 5 : Register_A000 = Bank
            Case 6 : Register_C000 = Bank
            Case 7 : Register_E000 = Bank
        End Select

        SetupBanks()
    End Sub
    Public Sub Select16KPROM(Bank As Byte, Page As Byte)
        Select8KPROM(Bank, Page)
        Select8KPROM(Bank + 1, Page + 1)
    End Sub
    Public Sub Select32KPROM(Bank As Byte)
        Select8KPROM(Bank, 4)
        Select8KPROM(Bank + 1, 5)
        Select8KPROM(Bank + 2, 6)
        Select8KPROM(Bank + 3, 7)
    End Sub
    Public Sub Select32KPROM(Bank0 As Byte, _
                             Bank1 As Byte, _
                             Bank2 As Byte, _
                             Bank3 As Byte)
        Select8KPROM(Bank0, 4)
        Select8KPROM(Bank1, 5)
        Select8KPROM(Bank2, 6)
        Select8KPROM(Bank3, 7)
    End Sub
#End Region

#Region "Select VROM"
    Public Sub Select1KVROM(Bank As Byte, Page As Byte)
        Bank = MaskVROM(Bank, ROMHeader.ChrSize * 8)
        CopyBanks(Page, Bank, 1)
    End Sub
    Public Sub Select2KVROM(Bank As Byte, Page As Byte)
        Bank = MaskVROM(Bank, ROMHeader.ChrSize * 4)
        CopyBanks(Page * 2, Bank * 2, 2)
    End Sub
    Public Sub Select4KVROM(Bank As Byte, Page As Byte)
        Bank = MaskVROM(Bank, ROMHeader.ChrSize * 2)
        CopyBanks(Page * 4, Bank * 4, 4)
    End Sub
    Public Sub Select8KVROM(Bank As Byte)
        Bank = MaskVROM(Bank, ROMHeader.ChrSize)
        CopyBanks(0, Bank * 8, 8)
    End Sub
#End Region

End Module
