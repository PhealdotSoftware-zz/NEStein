Option Explicit On
Module Mapper
    Public Sub Mapper_Reset()
        Select Case ROMHeader.Mapper
            Case 0 : DefaultReset()
            Case 1 : Mapper1_Reset()
            Case 2 : DefaultReset()
            Case 3 : DefaultReset()
            Case 4 : Mapper4_Reset()
            Case 5 : Mapper5_Reset()
            Case 7 : DefaultReset()
            Case 228 : Mapper228_Reset()
            Case Else
                MsgBox("Mapper " & ROMHeader.Mapper & " is not supported yet!", vbCritical, "Error!")
        End Select
    End Sub
    Public Sub Mapper_Write(ByVal Address As Integer, ByVal Value As Byte)
        Select Case ROMHeader.Mapper
            Case 1 : Mapper1_Write(Address, Value)
            Case 2 : Mapper2_Write(Address, Value)
            Case 3 : Mapper3_Write(Address, Value)
            Case 4 : Mapper4_Write(Address, Value)
            Case 7 : Mapper7_Write(Address, Value)
            Case 228 : Mapper228_Write(Address, Value)
        End Select
    End Sub
    Public Sub Mapper_WriteLow(ByVal Address As Integer, ByVal Value As Byte)
        Select Case ROMHeader.Mapper
            Case 5 : Mapper5_WriteLow(Address, Value)
        End Select
    End Sub
    Public Function Mapper_ReadLow(ByVal Address As Integer)
        Select Case ROMHeader.Mapper
            Case 5 : Return Mapper5_ReadLow(Address)
            Case Else : Return Nothing
        End Select
    End Function
    Public Sub Mapper_HBlank(ByVal ScanLine As Integer, ByVal Two As Byte)
        Select Case ROMHeader.Mapper
            Case 4 : Mapper4_HBlank(ScanLine, Two)
            Case 5 : Mapper5_HBlank(ScanLine)
        End Select
    End Sub
    Public Sub DefaultReset()
        Select32KPROM(0, 1, &HFE, &HFF)
        If ROMHeader.ChrSize Then Select8KVROM(0)
    End Sub
    Private Sub Mapper228_Reset()
        Select32KPROM(0)
        Select8KVROM(0)
    End Sub
    Private Sub Mapper228_Write(ByVal Address As Integer, ByVal Value As Byte)
        Dim Prg As Byte = (Address And &H780) >> 7

        Select Case (Address And &H1800) >> 11
            Case 1 : Prg = Prg Or &H10
            Case 3 : Prg = Prg Or &H20
        End Select

        If Address And &H20 Then
            Prg = Prg << 1
            If Address And &H40 Then Prg += 1

            Select16KPROM(Prg * 4, 4)
            Select16KPROM(Prg * 4, 6)
        Else
            Select32KPROM(Prg * 4)
        End If

        Select8KVROM(((Address And &HF) << 2) Or Value And &H3)

        If Address And &H2000 Then
            ROMHeader.Mirroring = MIRROR_HORIZONTAL
        Else
            ROMHeader.Mirroring = MIRROR_VERTICAL
        End If
        Do_Mirroring()
    End Sub
End Module
