Module Mapper_3
    '+==========+
    '| Mapper 3 |
    '+==========+

    Public Sub Mapper3_Write(ByVal Address As Integer, ByVal Value As Byte)
        Select8KVROM(Value)
    End Sub
End Module
