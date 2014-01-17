Module Mapper_2
    '+==========+
    '| Mapper 2 |
    '+==========+

    Public Sub Mapper2_Write(ByVal Address As Integer, ByVal Value As Byte)
        Select16KPROM(Value * 2, 4)
    End Sub
End Module
