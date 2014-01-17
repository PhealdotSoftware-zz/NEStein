Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D
Module PPU_Direct3D
    Dim Device As Device
    Dim Buffer As VertexBuffer
    Dim Texture As Texture
    Dim pData As GraphicsStream

    Dim Texture_Video_Buffer(256 * 240 * 4) As Byte
    Public Sub Direct3D_Initialize()
        Dim Present As New PresentParameters
        With Present
            .BackBufferCount = 1
            .BackBufferFormat = Manager.Adapters(0).CurrentDisplayMode.Format
            .BackBufferWidth = 256
            .BackBufferHeight = 240
            .Windowed = True
            .SwapEffect = SwapEffect.Discard
            .PresentationInterval = PresentInterval.Immediate 'FPS ILIMITADO
        End With

        Device = New Device(0, DeviceType.Hardware, FrmMain.NesScreen.Handle, CreateFlags.HardwareVertexProcessing, Present)
        Device.RenderState.CullMode = Cull.None

        Create_Polygon()
    End Sub
    Private Sub Create_Polygon()
        Buffer = New VertexBuffer(GetType(CustomVertex.TransformedTextured), 4, Device, Usage.None, CustomVertex.TransformedTextured.Format, Pool.Managed)
        Dim Verts(3) As CustomVertex.TransformedTextured
        Verts(0) = New CustomVertex.TransformedTextured(0, 0, 0, 1, 0, 0)
        Verts(1) = New CustomVertex.TransformedTextured(256, 0, 0, 1, 1, 0)
        Verts(2) = New CustomVertex.TransformedTextured(0, 256, 0, 1, 0, 1)
        Verts(3) = New CustomVertex.TransformedTextured(256, 256, 0, 1, 1, 1)
        Buffer.SetData(Verts, 0, LockFlags.None)
    End Sub
    Public Sub Draw_Screen()
        Device.Clear(ClearFlags.Target, Color.Black, 0, 0)
        Device.BeginScene()

        '---------------------------------------

        Texture = New Texture(Device, 256, 256, 1, Usage.None, Format.A8R8G8B8, Pool.Managed)
        pData = Texture.LockRectangle(0, LockFlags.None)
        Dim j As Integer
        For i As Integer = 0 To UBound(Texture_Video_Buffer) - 4 Step 4
            Texture_Video_Buffer(i) = Video_Buffer(j) And &HFF 'Red
            Texture_Video_Buffer(i + 1) = (Video_Buffer(j) And &HFF00) \ &H100  'Green
            Texture_Video_Buffer(i + 2) = (Video_Buffer(j) And &HFF0000) \ &H10000  'Blue
            Texture_Video_Buffer(i + 3) = &HFF 'Alpha
            j += 1
        Next
        pData.Write(Texture_Video_Buffer, 0, UBound(Texture_Video_Buffer))
        Texture.UnlockRectangle(0)
        Device.SetTexture(0, Texture)

        '---------------------------------------
        
        Device.VertexFormat = CustomVertex.TransformedTextured.Format
        Device.SetStreamSource(0, Buffer, 0)
        Device.DrawPrimitives(PrimitiveType.TriangleStrip, 0, 2)

        '---------------------------------------

        Device.EndScene()
        Device.Present()

        Texture.Dispose()
        pData.Dispose()
    End Sub
End Module
