Module Mapper_7
    '+==========+
    '| Mapper 7 |
    '+==========+

    Public Sub Mapper7_Write(ByVal Address As Integer, ByVal Value As Byte)
        Select32KPROM((Value And &HF) * 4)

        If Value And &H10 Then
            ROMHeader.Mirroring = MIRROR_ONESCREEN_LOW
        Else
            ROMHeader.Mirroring = MIRROR_ONESCREEN_HIGH
        End If

        Do_Mirroring()
    End Sub
End Module
