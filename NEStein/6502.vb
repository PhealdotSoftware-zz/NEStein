Option Explicit On
Module _6502

#Region "Opcode Constants"
    Const INS_ADC As Byte = 0
    Const INS_AND As Byte = 1
    Const INS_ASL As Byte = 2
    Const INS_BCC As Byte = 3
    Const INS_BCS As Byte = 4
    Const INS_BEQ As Byte = 5
    Const INS_BIT As Byte = 6
    Const INS_BMI As Byte = 7
    Const INS_BNE As Byte = 8
    Const INS_BPL As Byte = 9
    Const INS_BRK As Byte = 10
    Const INS_BVC As Byte = 11
    Const INS_BVS As Byte = 12
    Const INS_CLC As Byte = 13
    Const INS_CLD As Byte = 14
    Const INS_CLI As Byte = 15
    Const INS_CLV As Byte = 16
    Const INS_CMP As Byte = 17
    Const INS_CPX As Byte = 18
    Const INS_CPY As Byte = 19
    Const INS_DEC As Byte = 20
    Const INS_DEX As Byte = 21
    Const INS_DEY As Byte = 22
    Const INS_EOR As Byte = 23
    Const INS_INC As Byte = 24
    Const INS_INX As Byte = 25
    Const INS_INY As Byte = 26
    Const INS_JMP As Byte = 27
    Const INS_JSR As Byte = 28
    Const INS_LDA As Byte = 29
    Const INS_LDX As Byte = 30
    Const INS_LDY As Byte = 31
    Const INS_LSR As Byte = 32
    Const INS_NOP As Byte = 33
    Const INS_ORA As Byte = 34
    Const INS_PHA As Byte = 35
    Const INS_PHP As Byte = 36
    Const INS_PLA As Byte = 37
    Const INS_PLP As Byte = 38
    Const INS_ROL As Byte = 39
    Const INS_ROR As Byte = 40
    Const INS_RTI As Byte = 41
    Const INS_RTS As Byte = 42
    Const INS_SBC As Byte = 43
    Const INS_SEC As Byte = 44
    Const INS_SED As Byte = 45
    Const INS_SEI As Byte = 46
    Const INS_STA As Byte = 47
    Const INS_STX As Byte = 48
    Const INS_STY As Byte = 49
    Const INS_TAX As Byte = 50
    Const INS_TAY As Byte = 51
    Const INS_TSX As Byte = 52
    Const INS_TXA As Byte = 53
    Const INS_TXS As Byte = 54
    Const INS_TYA As Byte = 55
    'Unofficial
    Const INS_LAX As Byte = 56
    Const INS_SAX As Byte = 57
    Const INS_DCP As Byte = 58
    Const INS_ISB As Byte = 59
    Const INS_RLA As Byte = 60
    Const INS_RRA As Byte = 61
    Const INS_SLO As Byte = 62
    Const INS_SRE As Byte = 63
#End Region

#Region "Addressing Mode Constants"
    Const ADDR_IMP As Byte = 0
    Const ADDR_ACC As Byte = 1
    Const ADDR_IMM As Byte = 2
    Const ADDR_ZP As Byte = 3
    Const ADDR_ZPX As Byte = 4
    Const ADDR_ZPY As Byte = 5
    Const ADDR_REL As Byte = 6
    Const ADDR_ABSO As Byte = 7
    Const ADDR_ABSX As Byte = 8
    Const ADDR_ABSY As Byte = 9
    Const ADDR_IND As Byte = 10
    Const ADDR_INDX As Byte = 11
    Const ADDR_INDY As Byte = 12
#End Region

#Region "Instruction Table"
    Private Instruction() As Byte = {
        INS_BRK, INS_ORA, INS_NOP, INS_SLO, INS_NOP, INS_ORA, INS_ASL, INS_SLO, INS_PHP, INS_ORA, INS_ASL, INS_NOP, INS_NOP, INS_ORA, INS_ASL, INS_SLO,
        INS_BPL, INS_ORA, INS_NOP, INS_SLO, INS_NOP, INS_ORA, INS_ASL, INS_SLO, INS_CLC, INS_ORA, INS_NOP, INS_SLO, INS_NOP, INS_ORA, INS_ASL, INS_SLO,
        INS_JSR, INS_AND, INS_NOP, INS_RLA, INS_BIT, INS_AND, INS_ROL, INS_RLA, INS_PLP, INS_AND, INS_ROL, INS_NOP, INS_BIT, INS_AND, INS_ROL, INS_RLA,
        INS_BMI, INS_AND, INS_NOP, INS_RLA, INS_NOP, INS_AND, INS_ROL, INS_RLA, INS_SEC, INS_AND, INS_NOP, INS_RLA, INS_NOP, INS_AND, INS_ROL, INS_RLA,
        INS_RTI, INS_EOR, INS_NOP, INS_SRE, INS_NOP, INS_EOR, INS_LSR, INS_SRE, INS_PHA, INS_EOR, INS_LSR, INS_NOP, INS_JMP, INS_EOR, INS_LSR, INS_SRE,
        INS_BVC, INS_EOR, INS_NOP, INS_SRE, INS_NOP, INS_EOR, INS_LSR, INS_SRE, INS_CLI, INS_EOR, INS_NOP, INS_SRE, INS_NOP, INS_EOR, INS_LSR, INS_SRE,
        INS_RTS, INS_ADC, INS_NOP, INS_RRA, INS_NOP, INS_ADC, INS_ROR, INS_RRA, INS_PLA, INS_ADC, INS_ROR, INS_NOP, INS_JMP, INS_ADC, INS_ROR, INS_RRA,
        INS_BVS, INS_ADC, INS_NOP, INS_RRA, INS_NOP, INS_ADC, INS_ROR, INS_RRA, INS_SEI, INS_ADC, INS_NOP, INS_RRA, INS_NOP, INS_ADC, INS_ROR, INS_RRA,
        INS_NOP, INS_STA, INS_NOP, INS_SAX, INS_STY, INS_STA, INS_STX, INS_SAX, INS_DEY, INS_NOP, INS_TXA, INS_NOP, INS_STY, INS_STA, INS_STX, INS_SAX,
        INS_BCC, INS_STA, INS_NOP, INS_NOP, INS_STY, INS_STA, INS_STX, INS_SAX, INS_TYA, INS_STA, INS_TXS, INS_NOP, INS_NOP, INS_STA, INS_NOP, INS_NOP,
        INS_LDY, INS_LDA, INS_LDX, INS_LAX, INS_LDY, INS_LDA, INS_LDX, INS_LAX, INS_TAY, INS_LDA, INS_TAX, INS_NOP, INS_LDY, INS_LDA, INS_LDX, INS_LAX,
        INS_BCS, INS_LDA, INS_NOP, INS_LAX, INS_LDY, INS_LDA, INS_LDX, INS_LAX, INS_CLV, INS_LDA, INS_TSX, INS_LAX, INS_LDY, INS_LDA, INS_LDX, INS_LAX,
        INS_CPY, INS_CMP, INS_NOP, INS_DCP, INS_CPY, INS_CMP, INS_DEC, INS_DCP, INS_INY, INS_CMP, INS_DEX, INS_NOP, INS_CPY, INS_CMP, INS_DEC, INS_DCP,
        INS_BNE, INS_CMP, INS_NOP, INS_DCP, INS_NOP, INS_CMP, INS_DEC, INS_DCP, INS_CLD, INS_CMP, INS_NOP, INS_DCP, INS_NOP, INS_CMP, INS_DEC, INS_DCP,
        INS_CPX, INS_SBC, INS_NOP, INS_ISB, INS_CPX, INS_SBC, INS_INC, INS_ISB, INS_INX, INS_SBC, INS_NOP, INS_SBC, INS_CPX, INS_SBC, INS_INC, INS_ISB,
        INS_BEQ, INS_SBC, INS_NOP, INS_ISB, INS_NOP, INS_SBC, INS_INC, INS_ISB, INS_SED, INS_SBC, INS_NOP, INS_ISB, INS_NOP, INS_SBC, INS_INC, INS_ISB
    }
#End Region

#Region "Addressing Mode Table"
    Private AddrMode() As Byte = {
        ADDR_IMP, ADDR_INDX, ADDR_IMP, ADDR_INDX, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_IMP, ADDR_IMM, ADDR_ACC, ADDR_IMM, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO,
        ADDR_REL, ADDR_INDY, ADDR_IMP, ADDR_INDY, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_IMP, ADDR_ABSY, ADDR_IMP, ADDR_ABSY, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX,
        ADDR_ABSO, ADDR_INDX, ADDR_IMP, ADDR_INDX, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_IMP, ADDR_IMM, ADDR_ACC, ADDR_IMM, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO,
        ADDR_REL, ADDR_INDY, ADDR_IMP, ADDR_INDY, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_IMP, ADDR_ABSY, ADDR_IMP, ADDR_ABSY, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX,
        ADDR_IMP, ADDR_INDX, ADDR_IMP, ADDR_INDX, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_IMP, ADDR_IMM, ADDR_ACC, ADDR_IMM, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO,
        ADDR_REL, ADDR_INDY, ADDR_IMP, ADDR_INDY, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_IMP, ADDR_ABSY, ADDR_IMP, ADDR_ABSY, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX,
        ADDR_IMP, ADDR_INDX, ADDR_IMP, ADDR_INDX, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_IMP, ADDR_IMM, ADDR_ACC, ADDR_IMM, ADDR_IND, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO,
        ADDR_REL, ADDR_INDY, ADDR_IMP, ADDR_INDY, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_IMP, ADDR_ABSY, ADDR_IMP, ADDR_ABSY, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX,
        ADDR_IMM, ADDR_INDX, ADDR_IMM, ADDR_INDX, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_IMP, ADDR_IMM, ADDR_IMP, ADDR_IMM, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO,
        ADDR_REL, ADDR_INDY, ADDR_IMP, ADDR_INDY, ADDR_ZPX, ADDR_ZPX, ADDR_ZPY, ADDR_ZPY, ADDR_IMP, ADDR_ABSY, ADDR_IMP, ADDR_ABSY, ADDR_ABSX, ADDR_ABSX, ADDR_ABSY, ADDR_ABSY,
        ADDR_IMM, ADDR_INDX, ADDR_IMM, ADDR_INDX, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_IMP, ADDR_IMM, ADDR_IMP, ADDR_IMM, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO,
        ADDR_REL, ADDR_INDY, ADDR_IMP, ADDR_INDY, ADDR_ZPX, ADDR_ZPX, ADDR_ZPY, ADDR_ZPY, ADDR_IMP, ADDR_ABSY, ADDR_IMP, ADDR_ABSY, ADDR_ABSX, ADDR_ABSX, ADDR_ABSY, ADDR_ABSY,
        ADDR_IMM, ADDR_INDX, ADDR_IMM, ADDR_INDX, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_IMP, ADDR_IMM, ADDR_IMP, ADDR_IMM, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO,
        ADDR_REL, ADDR_INDY, ADDR_IMP, ADDR_INDY, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_IMP, ADDR_ABSY, ADDR_IMP, ADDR_ABSY, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX,
        ADDR_IMM, ADDR_INDX, ADDR_IMM, ADDR_INDX, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_ZP, ADDR_IMP, ADDR_IMM, ADDR_IMP, ADDR_IMM, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO, ADDR_ABSO,
        ADDR_REL, ADDR_INDY, ADDR_IMP, ADDR_INDY, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_ZPX, ADDR_IMP, ADDR_ABSY, ADDR_IMP, ADDR_ABSY, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX, ADDR_ABSX
    }
#End Region

#Region "Tick Table"
    Private Ticks() As Byte = {
        7, 6, 2, 8, 3, 3, 5, 5, 3, 2, 2, 2, 4, 4, 6, 6,
        2, 5, 2, 8, 4, 4, 6, 6, 2, 4, 2, 7, 4, 4, 7, 7,
        6, 6, 2, 8, 3, 3, 5, 5, 4, 2, 2, 2, 4, 4, 6, 6,
        2, 5, 2, 8, 4, 4, 6, 6, 2, 4, 2, 7, 4, 4, 7, 7,
        6, 6, 2, 8, 3, 3, 5, 5, 3, 2, 2, 2, 3, 4, 6, 6,
        2, 5, 2, 8, 4, 4, 6, 6, 2, 4, 2, 7, 4, 4, 7, 7,
        6, 6, 2, 8, 3, 3, 5, 5, 4, 2, 2, 2, 5, 4, 6, 6,
        2, 5, 2, 8, 4, 4, 6, 6, 2, 4, 2, 7, 4, 4, 7, 7,
        2, 6, 2, 6, 3, 3, 3, 3, 2, 2, 2, 2, 4, 4, 4, 4,
        2, 6, 2, 6, 4, 4, 4, 4, 2, 5, 2, 5, 5, 5, 5, 5,
        2, 6, 2, 6, 3, 3, 3, 3, 2, 2, 2, 2, 4, 4, 4, 4,
        2, 5, 2, 5, 4, 4, 4, 4, 2, 4, 2, 4, 4, 4, 4, 4,
        2, 6, 2, 8, 3, 3, 5, 5, 2, 2, 2, 2, 4, 4, 6, 6,
        2, 5, 2, 8, 4, 4, 6, 6, 2, 4, 2, 7, 4, 4, 7, 7,
        2, 6, 2, 8, 3, 3, 5, 5, 2, 2, 2, 2, 4, 4, 6, 6,
        2, 5, 2, 8, 4, 4, 6, 6, 2, 4, 2, 7, 4, 4, 7, 7
    }
#End Region

    Const C_Flag As Byte = &H1
    Const Z_Flag As Byte = &H2
    Const I_Flag As Byte = &H4
    Const D_Flag As Byte = &H8
    Const B_Flag As Byte = &H10
    Const R_Flag As Byte = &H20
    Const V_Flag As Byte = &H40
    Const N_Flag As Byte = &H80

    Const NMI_Vector As Integer = &HFFFA
    Const Reset_Vector As Integer = &HFFFC
    Const IRQ_Vector As Integer = &HFFFE

    ' Registers
    '---------------------------------------
    Private Structure CPU_Registers
        Dim PC As Integer 'Program Counter (16 bits)
        Dim AC As Byte 'Accumulator
        Dim X As Byte 'X Register
        Dim Y As Byte 'Y Register
        Dim SR As Byte 'Status Register
        Dim SP As Byte 'Stack Pointer
    End Structure
    Private Regs As CPU_Registers

    Private Effective_Address As Integer
    Private Tick_Count As Integer
    Private Opcode As Byte
    Public Executed_Cycles As Long
    Public CurrentLine As Integer
    Public Frames As Long

    Public Const NES_NTSC_CPU_SPEED_MHZ As Single = 1789773 'MHz (System 21.47727Mhz / 12)
    Public Const NES_PAL_CPU_SPEED_MHZ As Single = 1773447 'MHz (System 26.601171Mhz / 15)

    Const Cycles_Per_Scanline As Single = NES_NTSC_CPU_SPEED_MHZ / 60 / 262
    Const Base_Stack = &H100
    Public Sub CPU_Reset()
        Regs.AC = 0
        Regs.X = 0
        Regs.Y = 0
        Regs.SP = &HFF
        Regs.SR = &H20
        Regs.PC = ReadMemory16(Reset_Vector)
    End Sub
    Public Sub CPU_Execute()
        While Tick_Count <= Cycles_Per_Scanline
            Opcode = Fetch()
            Tick_Count += Ticks(Opcode)

            Select Case AddrMode(Opcode)
                Case ADDR_ABSO : Effective_Address = Fetch16()
                Case ADDR_ABSX
                    Dim Data As Integer = Fetch16()
                    Effective_Address = Data + Regs.X
                    If (Data And &HFF00) <> (Effective_Address And &HFF00) Then Tick_Count += 1 'Page Crossed
                Case ADDR_ABSY
                    Dim Data As Integer = Fetch16()
                    Effective_Address = Data + Regs.Y
                    If (Data And &HFF00) <> (Effective_Address And &HFF00) Then Tick_Count += 1 'Page Crossed
                Case ADDR_IMP
                Case ADDR_IMM
                    Effective_Address = Regs.PC
                    Regs.PC += 1
                Case ADDR_IND
                    Dim Temp As Integer = Fetch16()
                    Dim Data As Integer = (Temp And &HFF00) Or ((Temp + 1) And &HFF) 'Zero Page Wrap
                    Effective_Address = ReadMemory(Temp) + (ReadMemory(Data) * &H100)
                Case ADDR_INDX : Effective_Address = ReadMemory16_ZP(Fetch() + Regs.X)
                Case ADDR_INDY
                    Dim Temp As Integer = Fetch()
                    Dim Data As Integer = (Temp And &HFF00) Or ((Temp + 1) And &HFF) 'Zero Page Wrap
                    Data = ReadMemory(Temp) + (ReadMemory(Data) * &H100)
                    Effective_Address = Data + Regs.Y
                    If (Data And &HFF00) <> (Effective_Address And &HFF00) Then Tick_Count += 1 'Page Crossed
                Case ADDR_REL
                    Effective_Address = Fetch()
                    If Effective_Address >= &H80 Then Effective_Address -= &H100
                Case ADDR_ZP : Effective_Address = Fetch()
                Case ADDR_ZPX : Effective_Address = (Fetch() + Regs.X) And &HFF
                Case ADDR_ZPY : Effective_Address = (Fetch() + Regs.Y) And &HFF
            End Select

            Select Case Instruction(Opcode)
                Case INS_ADC : Add_With_Carry()
                Case INS_AND : And_With_Accumulator()
                Case INS_ASL : Arithmetic_Shift_Left()
                Case INS_BCC : Branch_On_Carry_Clear()
                Case INS_BCS : Branch_On_Carry_Set()
                Case INS_BEQ : Branch_On_Equal()
                Case INS_BIT : Bit_Test()
                Case INS_BMI : Branch_On_Minus()
                Case INS_BNE : Branch_On_Not_Equal()
                Case INS_BPL : Branch_On_Plus()
                Case INS_BRK : Break()
                Case INS_BVC : Branch_On_Overflow_Clear()
                Case INS_BVS : Branch_On_Overflow_Set()
                Case INS_CLC : Clear_Carry()
                Case INS_CLD : Clear_Decimal()
                Case INS_CLI : Clear_Interrupt_Disable()
                Case INS_CLV : Clear_Overflow()
                Case INS_CMP : Compare()
                Case INS_CPX : Compare_With_X()
                Case INS_CPY : Compare_With_Y()
                Case INS_DEC : Decrement()
                Case INS_DEX : Decrement_X()
                Case INS_DEY : Decrement_Y()
                Case INS_EOR : Exclusive_Or()
                Case INS_INC : Increment()
                Case INS_INX : Increment_X()
                Case INS_INY : Increment_Y()
                Case INS_JMP : Jump()
                Case INS_JSR : Jump_To_Subroutine()
                Case INS_LDA : Load_Accumulator()
                Case INS_LDX : Load_X()
                Case INS_LDY : Load_Y()
                Case INS_LSR : Logical_Shift_Right()
                Case INS_NOP : No_Operation()
                Case INS_ORA : Or_With_Accumulator()
                Case INS_PHA : Push_Accumulator()
                Case INS_PHP : Push_Processor_Status()
                Case INS_PLA : Pull_Accumulator()
                Case INS_PLP : Pull_Processor_Status()
                Case INS_ROL : Rotate_Left()
                Case INS_ROR : Rotate_Right()
                Case INS_RTI : Return_From_Interrupt()
                Case INS_RTS : Return_From_Subroutine()
                Case INS_SBC : Subtract_With_Carry()
                Case INS_SEC : Set_Carry()
                Case INS_SED : Set_Decimal()
                Case INS_SEI : Set_Interrupt_Disable()
                Case INS_STA : Store_Accumulator()
                Case INS_STX : Store_X()
                Case INS_STY : Store_Y()
                Case INS_TAX : Transfer_Accumulator_To_X()
                Case INS_TAY : Transfer_Accumulator_To_Y()
                Case INS_TSX : Transfer_Stack_Pointer_To_X()
                Case INS_TXA : Transfer_X_To_Accumulator()
                Case INS_TXS : Transfer_X_To_Stack_Pointer()
                Case INS_TYA : Transfer_Y_To_Accumulator()
                Case INS_LAX : Load_Accumulator() : Load_X()
                Case INS_SAX
                    Store_Accumulator() : Store_X()
                    WriteMemory(Effective_Address, Regs.AC And Regs.X)
                Case INS_DCP : Decrement() : Compare()
                Case INS_ISB : Increment() : Subtract_With_Carry()
                Case INS_SLO : Arithmetic_Shift_Left() : Or_With_Accumulator()
                Case INS_RLA : Rotate_Left() : And_With_Accumulator()
                Case INS_SRE : Logical_Shift_Right() : Exclusive_Or()
                Case INS_RRA : Rotate_Right() : Add_With_Carry()
            End Select
        End While

        Executed_Cycles += Tick_Count
        Tick_Count -= Cycles_Per_Scanline
    End Sub
    Private Sub Add_With_Carry() 'ADC
        Dim Data As Integer = ReadMemory(Effective_Address)
        Dim Temp As Integer = Regs.AC + Data + (Regs.SR And C_Flag)
        Test_Flag(Temp > &HFF, C_Flag)
        Test_Flag(((Not (Regs.AC Xor Data)) And (Regs.AC Xor Temp) And &H80), V_Flag)
        Regs.AC = Temp And &HFF
        Set_ZN_Flag(Regs.AC)
    End Sub
    Private Sub And_With_Accumulator() 'AND
        Dim Data As Integer = ReadMemory(Effective_Address)
        Regs.AC = Regs.AC And Data
        Set_ZN_Flag(Regs.AC)
    End Sub
    Private Sub Arithmetic_Shift_Left() 'ASL
        If AddrMode(Opcode) = ADDR_ACC Then
            Test_Flag(Regs.AC And &H80, C_Flag)
            Regs.AC <<= 1
            Set_ZN_Flag(Regs.AC)
        Else
            Dim Data As Byte = ReadMemory(Effective_Address)
            Test_Flag(Data And &H80, C_Flag)
            Data <<= 1
            WriteMemory(Effective_Address, Data)
            Set_ZN_Flag(Data)
        End If
    End Sub
    Private Sub Branch_On_Carry_Clear() 'BCC
        If (Regs.SR And C_Flag) = 0 Then
            If (Regs.PC And &HFF00) <> (Regs.PC + Effective_Address And &HFF00) Then
                Tick_Count += 2
            Else
                Tick_Count += 1
            End If
            Regs.PC += Effective_Address
        End If
    End Sub
    Private Sub Branch_On_Carry_Set() 'BCS
        If Regs.SR And C_Flag Then
            If (Regs.PC And &HFF00) <> (Regs.PC + Effective_Address And &HFF00) Then
                Tick_Count += 2
            Else
                Tick_Count += 1
            End If
            Regs.PC += Effective_Address
        End If
    End Sub
    Private Sub Branch_On_Equal() 'BEQ
        If Regs.SR And Z_Flag Then
            If (Regs.PC And &HFF00) <> (Regs.PC + Effective_Address And &HFF00) Then
                Tick_Count += 2
            Else
                Tick_Count += 1
            End If
            Regs.PC += Effective_Address
        End If
    End Sub
    Private Sub Bit_Test() 'BIT
        Dim Data As Integer = ReadMemory(Effective_Address)
        Test_Flag((Data And Regs.AC) = 0, Z_Flag)
        Test_Flag(Data And &H80, N_Flag)
        Test_Flag(Data And &H40, V_Flag)
    End Sub
    Private Sub Branch_On_Minus() 'BMI
        If Regs.SR And N_Flag Then
            If (Regs.PC And &HFF00) <> (Regs.PC + Effective_Address And &HFF00) Then
                Tick_Count += 2
            Else
                Tick_Count += 1
            End If
            Regs.PC += Effective_Address
        End If
    End Sub
    Private Sub Branch_On_Not_Equal() 'BNE
        If (Regs.SR And Z_Flag) = 0 Then
            If (Regs.PC And &HFF00) <> (Regs.PC + Effective_Address And &HFF00) Then
                Tick_Count += 2
            Else
                Tick_Count += 1
            End If
            Regs.PC += Effective_Address
        End If
    End Sub
    Private Sub Branch_On_Plus() 'BPL
        If (Regs.SR And N_Flag) = 0 Then
            If (Regs.PC And &HFF00) <> (Regs.PC + Effective_Address And &HFF00) Then
                Tick_Count += 2
            Else
                Tick_Count += 1
            End If
            Regs.PC += Effective_Address
        End If
    End Sub
    Private Sub Break() 'BRK
        Regs.PC += 1
        Push(Regs.PC >> 8)
        Push(Regs.PC And &HFF)
        Set_Flag(B_Flag)
        Push(Regs.SR)
        Set_Flag(I_Flag)
        Regs.PC = ReadMemory16(IRQ_Vector)
    End Sub
    Private Sub Branch_On_Overflow_Clear() 'BVC
        If (Regs.SR And V_Flag) = 0 Then
            If (Regs.PC And &HFF00) <> (Regs.PC + Effective_Address And &HFF00) Then
                Tick_Count += 2
            Else
                Tick_Count += 1
            End If
            Regs.PC += Effective_Address
        End If
    End Sub
    Private Sub Branch_On_Overflow_Set() 'BVS
        If Regs.SR And V_Flag Then
            If (Regs.PC And &HFF00) <> (Regs.PC + Effective_Address And &HFF00) Then
                Tick_Count += 2
            Else
                Tick_Count += 1
            End If
            Regs.PC += Effective_Address
        End If
    End Sub
    Private Sub Clear_Carry() 'CLC
        Clear_Flag(C_Flag)
    End Sub
    Private Sub Clear_Decimal() 'CLD
        Clear_Flag(D_Flag)
    End Sub
    Private Sub Clear_Interrupt_Disable() 'CLI
        Clear_Flag(I_Flag)
    End Sub
    Private Sub Clear_Overflow() 'CLV
        Clear_Flag(V_Flag)
    End Sub
    Private Sub Compare() 'CMP
        Dim Data As Integer = ReadMemory(Effective_Address)
        Dim Temp As Integer = Regs.AC - Data
        Test_Flag((Temp And &H8000) = 0, C_Flag)
        Set_ZN_Flag(Temp And &HFF)
    End Sub
    Private Sub Compare_With_X() 'CPX
        Dim Data As Integer = ReadMemory(Effective_Address)
        Dim Temp As Integer = Regs.X - Data
        Test_Flag((Temp And &H8000) = 0, C_Flag)
        Set_ZN_Flag(Temp And &HFF)
    End Sub
    Private Sub Compare_With_Y() 'CPY
        Dim Data As Integer = ReadMemory(Effective_Address)
        Dim Temp As Integer = Regs.Y - Data
        Test_Flag((Temp And &H8000) = 0, C_Flag)
        Set_ZN_Flag(Temp And &HFF)
    End Sub
    Private Sub Decrement() 'DEC
        WriteMemory(Effective_Address, (ReadMemory(Effective_Address) - 1) And &HFF)
        Dim Data As Integer = ReadMemory(Effective_Address)
        Set_ZN_Flag(Data)
    End Sub
    Private Sub Decrement_X() 'DEX
        Regs.X = (Regs.X - 1) And &HFF
        Set_ZN_Flag(Regs.X)
    End Sub
    Private Sub Decrement_Y() 'DEY
        Regs.Y = (Regs.Y - 1) And &HFF
        Set_ZN_Flag(Regs.Y)
    End Sub
    Private Sub Exclusive_Or() 'EOR
        Dim Data As Integer = ReadMemory(Effective_Address)
        Regs.AC = Regs.AC Xor Data
        Set_ZN_Flag(Regs.AC)
    End Sub
    Private Sub Increment() 'INC
        WriteMemory(Effective_Address, (ReadMemory(Effective_Address) + 1) And &HFF)
        Dim Data As Integer = ReadMemory(Effective_Address)
        Set_ZN_Flag(Data)
    End Sub
    Private Sub Increment_X() 'INX
        Regs.X = (Regs.X + 1) And &HFF
        Set_ZN_Flag(Regs.X)
    End Sub
    Private Sub Increment_Y() 'INY
        Regs.Y = (Regs.Y + 1) And &HFF
        Set_ZN_Flag(Regs.Y)
    End Sub
    Private Sub Jump() 'JMP
        Regs.PC = Effective_Address
    End Sub
    Private Sub Jump_To_Subroutine() 'JSR
        Push((Regs.PC - 1) >> 8)
        Push((Regs.PC - 1) And &HFF)
        Regs.PC = Effective_Address
    End Sub
    Private Sub Load_Accumulator() 'LDA
        Regs.AC = ReadMemory(Effective_Address)
        Set_ZN_Flag(Regs.AC)
    End Sub
    Private Sub Load_X() 'LDX
        Regs.X = ReadMemory(Effective_Address)
        Set_ZN_Flag(Regs.X)
    End Sub
    Private Sub Load_Y() 'LDY
        Regs.Y = ReadMemory(Effective_Address)
        Set_ZN_Flag(Regs.Y)
    End Sub
    Private Sub Logical_Shift_Right() 'LSR
        If AddrMode(Opcode) = ADDR_ACC Then
            Test_Flag(Regs.AC And &H1, C_Flag)
            Regs.AC >>= 1
            Set_ZN_Flag(Regs.AC)
        Else
            Dim Data As Byte = ReadMemory(Effective_Address)
            Test_Flag(Data And &H1, C_Flag)
            Data >>= 1
            WriteMemory(Effective_Address, Data)
            Set_ZN_Flag(Data)
        End If
    End Sub
    Private Sub No_Operation() 'NOP
        'Nothing to do here!
    End Sub
    Private Sub Or_With_Accumulator() 'ORA
        Dim Data As Integer = ReadMemory(Effective_Address)
        Regs.AC = Regs.AC Or Data
        Set_ZN_Flag(Regs.AC)
    End Sub
    Private Sub Push_Accumulator() 'PHA
        Push(Regs.AC)
    End Sub
    Private Sub Push_Processor_Status() 'PHP
        Push(Regs.SR Or B_Flag)
    End Sub
    Private Sub Pull_Accumulator() 'PLA
        Regs.AC = Pull()
        Set_ZN_Flag(Regs.AC)
    End Sub
    Private Sub Pull_Processor_Status() 'PLP
        Regs.SR = Pull() Or R_Flag
    End Sub
    Private Sub Rotate_Left() 'ROL
        If AddrMode(Opcode) = ADDR_ACC Then
            If (Regs.SR And C_Flag) Then
                Test_Flag(Regs.AC And &H80, C_Flag)
                Regs.AC = (Regs.AC << 1) Or &H1
            Else
                Test_Flag(Regs.AC And &H80, C_Flag)
                Regs.AC <<= 1
            End If
            Set_ZN_Flag(Regs.AC)
        Else
            Dim Data As Byte = ReadMemory(Effective_Address)
            If (Regs.SR And C_Flag) Then
                Test_Flag(Data And &H80, C_Flag)
                Data = (Data << 1) Or &H1
            Else
                Test_Flag(Data And &H80, C_Flag)
                Data <<= 1
            End If
            WriteMemory(Effective_Address, Data)
            Set_ZN_Flag(Data)
        End If
    End Sub
    Private Sub Rotate_Right() 'ROR
        If AddrMode(Opcode) = ADDR_ACC Then
            If (Regs.SR And C_Flag) Then
                Test_Flag(Regs.AC And &H1, C_Flag)
                Regs.AC = (Regs.AC >> 1) Or &H80
            Else
                Test_Flag(Regs.AC And &H1, C_Flag)
                Regs.AC >>= 1
            End If
            Set_ZN_Flag(Regs.AC)
        Else
            Dim Data As Byte = ReadMemory(Effective_Address)
            If (Regs.SR And C_Flag) Then
                Test_Flag(Data And &H1, C_Flag)
                Data = (Data >> 1) Or &H80
            Else
                Test_Flag(Data And &H1, C_Flag)
                Data >>= 1
            End If
            WriteMemory(Effective_Address, Data)
            Set_ZN_Flag(Data)
        End If
    End Sub
    Private Sub Return_From_Interrupt() 'RTI
        Regs.SR = Pull() Or R_Flag
        Regs.PC = Pull()
        Regs.PC = Regs.PC Or Pull() * &H100
    End Sub
    Private Sub Return_From_Subroutine() 'RTS
        Regs.PC = Pull()
        Regs.PC = Regs.PC Or Pull() * &H100
        Regs.PC += 1
    End Sub
    Private Sub Subtract_With_Carry() 'SBC
        Dim Data As Integer = ReadMemory(Effective_Address) Xor &HFF
        Dim Temp As Integer = Regs.AC + Data + (Regs.SR And C_Flag)
        Test_Flag(Temp > &HFF, C_Flag)
        Test_Flag(((Not (Regs.AC Xor Data)) And (Regs.AC Xor Temp) And &H80), V_Flag)
        Regs.AC = Temp And &HFF
        Set_ZN_Flag(Regs.AC)
    End Sub
    Private Sub Set_Carry() 'SEC
        Set_Flag(C_Flag)
    End Sub
    Private Sub Set_Decimal() 'SED
        Set_Flag(D_Flag)
    End Sub
    Private Sub Set_Interrupt_Disable() 'SEI
        Set_Flag(I_Flag)
    End Sub
    Private Sub Store_Accumulator() 'STA
        WriteMemory(Effective_Address, Regs.AC)
    End Sub
    Private Sub Store_X() 'STX
        WriteMemory(Effective_Address, Regs.X)
    End Sub
    Private Sub Store_Y() 'STY
        WriteMemory(Effective_Address, Regs.Y)
    End Sub
    Private Sub Transfer_Accumulator_To_X() 'TAX
        Regs.X = Regs.AC
        Set_ZN_Flag(Regs.X)
    End Sub
    Private Sub Transfer_Accumulator_To_Y() 'TAY
        Regs.Y = Regs.AC
        Set_ZN_Flag(Regs.Y)
    End Sub
    Private Sub Transfer_Stack_Pointer_To_X() 'TSX
        Regs.X = Regs.SP
        Set_ZN_Flag(Regs.X)
    End Sub
    Private Sub Transfer_X_To_Accumulator() 'TXA
        Regs.AC = Regs.X
        Set_ZN_Flag(Regs.AC)
    End Sub
    Private Sub Transfer_X_To_Stack_Pointer() 'TXS
        Regs.SP = Regs.X
    End Sub
    Private Sub Transfer_Y_To_Accumulator() 'TYA
        Regs.AC = Regs.Y
        Set_ZN_Flag(Regs.AC)
    End Sub

    'Flags
    Private Sub Set_Flag(Value As Byte)
        Regs.SR = Regs.SR Or Value
    End Sub
    Private Sub Clear_Flag(Value As Byte)
        Regs.SR = Regs.SR And Not Value
    End Sub
    Private Sub Set_ZN_Flag(Value As Byte)
        If Value Then
            Clear_Flag(Z_Flag)
        Else
            Set_Flag(Z_Flag)
        End If

        If Value And N_Flag Then
            Set_Flag(N_Flag)
        Else
            Clear_Flag(N_Flag)
        End If
    End Sub
    Private Sub Test_Flag(SetFlag As Boolean, Value As Byte)
        If SetFlag Then
            Set_Flag(Value)
        Else
            Clear_Flag(Value)
        End If
    End Sub

    'Misc
    Private Function Fetch() As Integer
        Dim RetVal As Integer = ReadMemory(Regs.PC) : Regs.PC += 1
        Return RetVal
    End Function
    Private Function Fetch16() As Integer
        Dim RetVal As Integer = ReadMemory16(Regs.PC) : Regs.PC += 2
        Return RetVal
    End Function
    Private Sub Push(Value As Byte)
        WriteMemory(Base_Stack + Regs.SP, Value)
        Regs.SP = (Regs.SP - 1) And &HFF
    End Sub
    Private Function Pull() As Byte
        Regs.SP = (Regs.SP + 1) And &HFF
        Return ReadMemory(Base_Stack + Regs.SP)
    End Function

    'Interrupts (Note: Break is on Instructions)
    Public Sub CPU_NMI()
        Push(Regs.PC >> 8)
        Push(Regs.PC And &HFF)
        Clear_Flag(B_Flag)
        Push(Regs.SR)
        Set_Flag(I_Flag)
        Regs.PC = ReadMemory16(NMI_Vector)
        Tick_Count += 7
    End Sub
    Public Sub CPU_IRQ()
        Push(Regs.PC >> 8)
        Push(Regs.PC And &HFF)
        Clear_Flag(B_Flag)
        Push(Regs.SR)
        Set_Flag(I_Flag)
        Regs.PC = ReadMemory16(IRQ_Vector)
        Tick_Count += 7
    End Sub
End Module
