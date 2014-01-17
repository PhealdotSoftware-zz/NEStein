Option Explicit On
Imports System.Drawing.Drawing2D
Module PPU_Main
    Public Render As Boolean = True
    Public Zoom As Integer = 1
    Public InterpolMode As Integer = 5

    ' Mirror Constants
    '---------------------------------------
    Public Const MIRROR_HORIZONTAL = 0
    Public Const MIRROR_VERTICAL = 1
    Public Const MIRROR_ONESCREEN_LOW = 2
    Public Const MIRROR_ONESCREEN_HIGH = 3
    Public Const MIRROR_FOURSCREEN = 4

    ' PPU Registers
    '---------------------------------------
    Public PPU_Control1 As Byte '$2000
    Public PPU_Control2 As Byte '$2001
    Public PPU_Status As Byte '$2002
    Private Loopy_V As Integer '$2006
    Private Loopy_T As Integer
    Private Loopy_X As Integer 'Scrolling
    Private HScroll As Byte, VScroll As Integer

    Private PPU_Latch As Byte
    Private SpriteAddress As Integer '$2003
    Private PPU_AddressIsHi As Boolean '$2005/$2006 Hi Toggle
    Private Buffer_Address_2007 As Byte

    ' Arrays
    '---------------------------------------
    Public VRAM(&H3FFF) As Byte
    Public SpriteRAM(&HFF) As Byte
    Public Video_Buffer(256 * 240) As Integer
    Public Palette(&HFF) As Integer
    Private Palette_32(31) As Integer
    Public Color_LookUp(65536 * 8 - 1) As Byte

    Public Name_Table(3, &H3FF) As Byte 'Mirroring
    Public Mirror(3) As Byte
    Private MirrorXor As Integer
    Public Function PPU_Read(Address As Integer) As Integer
        Select Case Address
            Case &H2000, &H2001, &H2005, &H2006 : Return PPU_Latch
            Case &H2002
                Dim Temp As Byte = PPU_Status
                PPU_Status = PPU_Status And &H60
                PPU_AddressIsHi = True
                Return Temp
            Case &H2004
                Return SpriteRAM((SpriteAddress + 1) And &HFF)
            Case &H2007
                Dim V_Address As Integer = (Loopy_V And &H3FFF)
                If (PPU_Control1 And &H4) Then Loopy_V += 32 Else Loopy_V += 1

                If V_Address <= &H3EFF Then
                    If V_Address >= &H3000 Then V_Address = V_Address And &HEFFF
                    Dim Temp As Byte = Buffer_Address_2007
                    If V_Address >= &H2000 Then Buffer_Address_2007 = Name_Table(Mirror((V_Address And &HC00) \ &H400), V_Address And &H3FF) Else _
                        Buffer_Address_2007 = VRAM(V_Address)
                    Return Temp
                ElseIf V_Address >= &H3F00 Then
                    Return VRAM(V_Address)
                End If
        End Select

        Return Nothing
    End Function
    Public Sub PPU_Write(ByVal Address As Integer, ByVal Value As Integer)
        Select Case Address
            Case &H2000
                Loopy_T = (Loopy_T And &HF3FF) Or (Value And 3) * &H400
                PPU_Control1 = Value
            Case &H2001 : PPU_Control2 = Value
            Case &H2002 : PPU_Latch = Value
            Case &H2003 : SpriteAddress = Value
            Case &H2004
                SpriteAddress = (SpriteAddress + 1) And &HFF
                SpriteRAM(SpriteAddress) = Value
            Case &H2005
                If PPU_AddressIsHi Then 'Horizontal Scrolling
                    Loopy_T = (Loopy_T And &HFFE0) Or Value \ 8
                    Loopy_X = Value And 7
                Else 'Vertical Scrolling
                    Loopy_T = (Loopy_T And &HFC1F) Or ((Value And &HF8) * 4)
                    Loopy_T = (Loopy_T And &H8FFF) Or ((Value And 7) * &H1000)
                End If
                PPU_AddressIsHi = Not PPU_AddressIsHi
            Case &H2006
                If PPU_AddressIsHi Then
                    Loopy_T = (Loopy_T And &HFF) Or ((Value And &H3F) * &H100&)
                Else
                    Loopy_T = (Loopy_T And &HFF00) Or Value
                    Loopy_V = Loopy_T
                End If
                PPU_AddressIsHi = Not PPU_AddressIsHi
            Case &H2007
                Dim V_Address As Integer = (Loopy_V And &H3FFF)
                If (PPU_Control1 And &H4) Then Loopy_V += 32 Else Loopy_V += 1

                If V_Address >= &H2000 And V_Address <= &H3EFF Then 'Name Table
                    If V_Address >= &H3000 Then V_Address = V_Address And &HEFFF
                    Name_Table(Mirror((V_Address And &HC00) \ &H400), V_Address And &H3FF) = Value
                ElseIf V_Address >= &H3F00 Then 'Palette
                    If (V_Address And &HF) = 0 Then
                        VRAM(&H3F00) = Value 'BG 0 Pal
                        VRAM(&H3F10) = Value 'SPR 0 Pal
                    ElseIf (V_Address And &H10) = 0 Then
                        VRAM(&H3F00 + (V_Address And &HF)) = Value 'BG Pal
                    Else
                        VRAM(&H3F10 + (V_Address And &HF)) = Value 'SPR Pal
                    End If
                Else
                    VRAM(V_Address) = Value
                End If
        End Select
    End Sub
    Public Sub Loopy_Next_Frame()
        Loopy_V = Loopy_T
    End Sub
    Public Sub Loopy_Scanline_Start()
        Loopy_V = (Loopy_V And &HFBE0) Or (Loopy_T And &H41F)
    End Sub
    Public Sub Loopy_Next_Line()
        If (Loopy_V And &H7000) = &H7000 Then
            Loopy_V = Loopy_V And &H8FFF
            If (Loopy_V And &H3E0) = &H3A0 Then
                Loopy_V = Loopy_V Xor &H800
                Loopy_V = Loopy_V And &HFC1F
            Else
                If (Loopy_V And &H3E0) = &H3E0 Then
                    Loopy_V = Loopy_V And &HFC1F
                Else
                    Loopy_V = Loopy_V + &H20
                End If
            End If
        Else
            Loopy_V = Loopy_V + &H1000
        End If
    End Sub
    Public Sub Render_ScanLine(ByVal ScanLine As Long)
        If ScanLine >= 240 Then Exit Sub

        For i As Integer = 0 To 31 Step 4
            Palette_32(i) = Palette(VRAM(i + &H3F00))
            Palette_32(i + 1) = Palette(VRAM(i + 1 + &H3F00))
            Palette_32(i + 2) = Palette(VRAM(i + 2 + &H3F00))
            Palette_32(i + 3) = Palette(VRAM(i + 3 + &H3F00))
        Next i

        If ScanLine < 240 Then 'Clear current Scanline
            For i As Integer = 256 * ScanLine To (256 * ScanLine) + 255 Step 4
                Video_Buffer(i) = Palette_32(16)
                Video_Buffer(i + 1) = Palette_32(16)
                Video_Buffer(i + 2) = Palette_32(16)
                Video_Buffer(i + 3) = Palette_32(16)
            Next i
        End If

        '-------------------------------------------------------------------------

        If (PPU_Control2 And BIT_3) = 0 And (PPU_Control2 And BIT_4) = 0 Or Not Render Then
            If ScanLine > SpriteRAM(0) + 8 Then PPU_Status = PPU_Status Or &H40
            Exit Sub
        End If
        If ScanLine = 239 Then PPU_Status = PPU_Status Or BIT_7

        If ScanLine = 0 Then
            Loopy_Next_Frame()
        Else
            Loopy_Scanline_Start()
        End If

        If PPU_Control2 And BIT_4 Then Render_Sprites(ScanLine, True)
        If PPU_Control2 And BIT_3 Then Render_Background(ScanLine)
        If PPU_Control2 And BIT_4 Then Render_Sprites(ScanLine, False)

        Loopy_Next_Line()
    End Sub
    Public Sub Render_Background(ByVal ScanLine As Integer)
        Dim Background_Pattern_Table_Address As Integer
        Dim Name_Table_Address As Integer = &H2000 + (Loopy_V And &HC00)
        Dim Name_Table_Number As Integer = (Name_Table_Address And &HC00) \ &H400
        Dim Tile_Row As Byte, Tile_Y_Offset As Integer
        Dim Tile_Counter As Integer
        Dim Pixel_Color As Byte
        Dim Background_Tile_Offset As Integer
        Dim Tile_Index As Byte, Low_Byte As Byte, High_Byte As Byte
        Dim LookUp As Byte
        Dim TileX, TileY, StartX As Integer
        Dim Color_LookUp_Offset As Integer
        Dim Offset As Integer
        Dim Background_Palette As Integer

        'Scrolling X/Y
        HScroll = (Loopy_V And 31) * 8 + Loopy_X
        VScroll = (Loopy_V \ 32 And 31) * 8 Or ((Loopy_V \ &H1000) And 7)

        Tile_Row = (VScroll \ 8) Mod 30
        Tile_Y_Offset = VScroll And 7

        If (PPU_Control2 And BIT_1) = 0 Then StartX = 8 'Clip Left 8 pixels?
        If PPU_Control1 And BIT_4 Then Background_Pattern_Table_Address = &H1000 'Bg Table

        For Tile_Counter = (HScroll \ 8) To 31
            TileX = Tile_Counter * 8 - HScroll + 7
            If TileX >= StartX Then
                Tile_Index = Name_Table(Mirror(Name_Table_Number), Tile_Counter + Tile_Row * 32)
                If TileX < 7 Then Offset = TileX Else Offset = 7
                TileX = TileX + ScanLine * 256
                LookUp = Name_Table(Mirror(Name_Table_Number), (&H3C0 + (Tile_Counter \ 4) + (Tile_Row \ 4) * 8))
                Select Case (Tile_Counter And 2) Or (Tile_Row And 2) * 2
                    Case 0
                        Background_Palette = LookUp * 4 And 12
                    Case 2
                        Background_Palette = LookUp And 12
                    Case 4
                        Background_Palette = LookUp \ 4 And 12
                    Case 6
                        Background_Palette = LookUp \ 16 And 12
                End Select
                Background_Tile_Offset = Background_Pattern_Table_Address + (Tile_Index * 16)
                If Tile_Y_Offset = 0 Then
                    For TileY = 0 To 7
                        Low_Byte = VRAM(Background_Tile_Offset + TileY)
                        High_Byte = VRAM(Background_Tile_Offset + TileY + 8)
                        Color_LookUp_Offset = Low_Byte * &H800 + High_Byte * &H8
                        For Current_Pixel = 0 To Offset
                            Pixel_Color = Color_LookUp(Color_LookUp_Offset + Current_Pixel)
                            If Pixel_Color Mod 4 <> 0 Then Video_Buffer(TileX - Current_Pixel) = Palette_32(Pixel_Color Or Background_Palette)
                        Next Current_Pixel
                        If TileX <= 61183 Then TileX = TileX + 256
                    Next TileY
                Else
                    Low_Byte = VRAM(Background_Tile_Offset + Tile_Y_Offset)
                    High_Byte = VRAM(Background_Tile_Offset + Tile_Y_Offset + 8)
                    Color_LookUp_Offset = Low_Byte * &H800 + High_Byte * &H8
                    For Current_Pixel = 0 To Offset
                        Pixel_Color = Color_LookUp(Color_LookUp_Offset + Current_Pixel)
                        If Pixel_Color Mod 4 <> 0 Then Video_Buffer(TileX - Current_Pixel) = Palette_32(Pixel_Color Or Background_Palette)
                    Next Current_Pixel
                End If
            End If
        Next Tile_Counter

        Name_Table_Address = Name_Table_Address Xor &H400
        Name_Table_Number = (Name_Table_Address And &HC00) \ &H400

        For Tile_Counter = 0 To (HScroll \ 8)
            TileX = Tile_Counter * 8 + 256 - HScroll + 7
            If TileX >= StartX Then
                Tile_Index = Name_Table(Mirror(Name_Table_Number), Tile_Counter + Tile_Row * 32)
                If TileX > 255 Then Offset = TileX - 255 Else Offset = 0
                TileX = TileX + ScanLine * 256
                LookUp = Name_Table(Mirror(Name_Table_Number), (&H3C0 + (Tile_Counter \ 4) + (Tile_Row \ 4) * &H8))

                Select Case (Tile_Counter And 2) Or (Tile_Row And 2) * 2
                    Case 0
                        Background_Palette = LookUp * 4 And 12
                    Case 2
                        Background_Palette = LookUp And 12
                    Case 4
                        Background_Palette = LookUp \ 4 And 12
                    Case 6
                        Background_Palette = LookUp \ 16 And 12
                End Select
                Background_Tile_Offset = Background_Pattern_Table_Address + (Tile_Index * 16)
                If Tile_Y_Offset = 0 Then
                    For TileY = 0 To 7
                        Low_Byte = VRAM(Background_Tile_Offset + TileY)
                        High_Byte = VRAM(Background_Tile_Offset + TileY + 8)
                        Color_LookUp_Offset = Low_Byte * 2048 + High_Byte * 8
                        For Current_Pixel = Offset To 7
                            Pixel_Color = Color_LookUp(Color_LookUp_Offset + Current_Pixel)
                            If Pixel_Color Mod 4 <> 0 Then Video_Buffer(TileX - Current_Pixel) = Palette_32(Pixel_Color Or Background_Palette)
                        Next Current_Pixel
                        If TileX <= 61183 Then TileX = TileX + 256
                    Next TileY
                Else
                    Low_Byte = VRAM(Background_Tile_Offset + Tile_Y_Offset)
                    High_Byte = VRAM(Background_Tile_Offset + Tile_Y_Offset + 8)
                    Color_LookUp_Offset = Low_Byte * 2048 + High_Byte * 8
                    For Current_Pixel = Offset To 7
                        Pixel_Color = Color_LookUp(Color_LookUp_Offset + Current_Pixel)
                        If Pixel_Color Mod 4 <> 0 Then Video_Buffer(TileX - Current_Pixel) = Palette_32(Pixel_Color Or Background_Palette)
                    Next Current_Pixel
                End If
            End If
        Next Tile_Counter
    End Sub
    Public Sub Render_Sprites(ByVal Scanline As Long, ByVal Drawn_In_Foreground As Boolean)
        Dim Current_Sprite_Address As Integer
        Dim Pixel_Color As Byte
        Dim Sprite_Offset As Integer
        Dim Index As Integer

        Dim SpriteX, SpriteY, SprPal As Integer
        Dim TileHeight As Integer
        Dim TileIndex, Attr As Byte
        Dim Pattern_Table_Address As Integer
        Dim Drawn_In_Background, Vertical_Flip, Horizontal_Flip As Boolean

        Dim Low_Byte As Byte, High_Byte As Byte
        Dim Current_X_Pixel As Integer, Sprite_Scanline_To_Draw As Integer

        For Current_Sprite_Address = 252 To 0 Step -4
            Index = Current_Sprite_Address / 4

            SpriteY = SpriteRAM(Current_Sprite_Address) + 1
            TileIndex = SpriteRAM(Current_Sprite_Address + 1)
            Attr = SpriteRAM(Current_Sprite_Address + 2)
            SpriteX = SpriteRAM(Current_Sprite_Address + 3)

            If PPU_Control1 And BIT_5 Then TileHeight = 16 Else TileHeight = 8
            If PPU_Control1 And BIT_3 Then Pattern_Table_Address = &H1000

            Drawn_In_Background = Attr And 32
            SprPal = 16 + (Attr And 3) * 4
            Vertical_Flip = Attr And 128
            Horizontal_Flip = Attr And 64

            If (Drawn_In_Background = Drawn_In_Foreground) And (SpriteY <= Scanline) And ((SpriteY + TileHeight) > Scanline) Then
                If TileHeight = 8 Then
                    If Vertical_Flip = False Then
                        Sprite_Scanline_To_Draw = Scanline - SpriteY
                    Else 'Vertical Flip
                        Sprite_Scanline_To_Draw = SpriteY + 7 - Scanline
                    End If

                    Sprite_Offset = Pattern_Table_Address + (TileIndex * 16)
                    Low_Byte = VRAM(Sprite_Offset + Sprite_Scanline_To_Draw)
                    High_Byte = VRAM(Sprite_Offset + Sprite_Scanline_To_Draw + 8)

                    For Current_X_Pixel = 0 To 7
                        If Horizontal_Flip = True Then  'Horizontal Flip
                            Pixel_Color = SprPal + (((High_Byte And (1 << (Current_X_Pixel))) >> (Current_X_Pixel)) << 1) + ((Low_Byte And (1 << (Current_X_Pixel))) >> (Current_X_Pixel))
                        Else
                            Pixel_Color = SprPal + (((High_Byte And (1 << (7 - Current_X_Pixel))) >> (7 - Current_X_Pixel)) << 1) + ((Low_Byte And (1 << (7 - Current_X_Pixel))) >> (7 - Current_X_Pixel))
                        End If

                        If Pixel_Color Mod 4 <> 0 Then
                            If (SpriteX + Current_X_Pixel) < 256 Then
                                Video_Buffer(Scanline * 256 + SpriteX + Current_X_Pixel) = Palette_32(SprPal Or Pixel_Color)
                                If Current_Sprite_Address = 0 Then PPU_Status = PPU_Status Or &H40
                            End If
                        End If
                    Next Current_X_Pixel
                ElseIf TileHeight = 16 Then
                    If Vertical_Flip = False Then
                        Sprite_Scanline_To_Draw = Scanline - SpriteY
                    Else
                        Sprite_Scanline_To_Draw = SpriteY + 15 - Scanline
                    End If

                    If Sprite_Scanline_To_Draw < 8 Then
                        If TileIndex Mod 2 = 0 Then
                            Sprite_Offset = &H0 + TileIndex * 16
                        Else
                            Sprite_Offset = &H1000 + (TileIndex - 1) * 16
                        End If
                    Else
                        Sprite_Scanline_To_Draw = Sprite_Scanline_To_Draw - 8
                        If TileIndex Mod 2 = 0 Then
                            Sprite_Offset = &H0 + (TileIndex + 1) * 16
                        Else
                            Sprite_Offset = &H1000 + TileIndex * 16
                        End If
                    End If

                    Low_Byte = VRAM(Sprite_Offset + Sprite_Scanline_To_Draw)
                    High_Byte = VRAM(Sprite_Offset + Sprite_Scanline_To_Draw + 8)

                    For Current_X_Pixel = 0 To 7
                        If Horizontal_Flip = True Then  'Horizontal Flip
                            Pixel_Color = SprPal + (((High_Byte And (1 << (Current_X_Pixel))) >> (Current_X_Pixel)) << 1) + ((Low_Byte And (1 << (Current_X_Pixel))) >> (Current_X_Pixel))
                        Else
                            Pixel_Color = SprPal + (((High_Byte And (1 << (7 - Current_X_Pixel))) >> (7 - Current_X_Pixel)) << 1) + ((Low_Byte And (1 << (7 - Current_X_Pixel))) >> (7 - Current_X_Pixel))
                        End If

                        If Pixel_Color Mod 4 <> 0 Then
                            If (SpriteX + Current_X_Pixel) < 256 Then
                                Video_Buffer(Scanline * 256 + SpriteX + Current_X_Pixel) = Palette_32(SprPal Or Pixel_Color)
                                If Current_Sprite_Address = 0 Then PPU_Status = PPU_Status Or &H40
                            End If
                        End If
                    Next Current_X_Pixel
                End If
            End If
        Next Current_Sprite_Address
    End Sub
    Public Sub Do_Mirroring()
        MirrorXor = (((ROMHeader.Mirroring + 1) Mod 3) * &H400)
        If ROMHeader.Mirroring = MIRROR_HORIZONTAL Then 'H
            Mirror(0) = 0 : Mirror(1) = 0 : Mirror(2) = 1 : Mirror(3) = 1
        ElseIf ROMHeader.Mirroring = MIRROR_VERTICAL Then 'V
            Mirror(0) = 0 : Mirror(1) = 1 : Mirror(2) = 0 : Mirror(3) = 1
        ElseIf ROMHeader.Mirroring = MIRROR_ONESCREEN_LOW Then '4L
            Mirror(0) = 0 : Mirror(1) = 0 : Mirror(2) = 0 : Mirror(3) = 0
        ElseIf ROMHeader.Mirroring = MIRROR_ONESCREEN_HIGH Then '4H
            Mirror(0) = 1 : Mirror(1) = 1 : Mirror(2) = 1 : Mirror(3) = 1
        ElseIf ROMHeader.Mirroring = MIRROR_FOURSCREEN Then '4
            Mirror(0) = 0 : Mirror(1) = 1 : Mirror(2) = 2 : Mirror(3) = 3
        End If
    End Sub
End Module
