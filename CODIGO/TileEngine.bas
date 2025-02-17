Attribute VB_Name = "Mod_TileEngine"
'Nexus AO mod Argentum Online 0.13
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Nexus AO mod Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'Quad Draw
Public indexList(0 To 5)     As Integer

Public ibQuad                As DxVBLibA.Direct3DIndexBuffer8

Public vbQuadIdx             As DxVBLibA.Direct3DVertexBuffer8

Dim temp_verts(3)            As TLVERTEX

Public OffsetCounterX        As Single

Public OffsetCounterY        As Single
    
Public WeatherFogX1          As Single

Public WeatherFogY1          As Single

Public WeatherFogX2          As Single

Public WeatherFogY2          As Single

Public WeatherFogCount       As Byte

Public ParticleOffsetX       As Long

Public ParticleOffsetY       As Long

Public LastOffsetX           As Integer

Public LastOffsetY           As Integer

'Map sizes in tiles
Public Const XMaxMapSize     As Byte = 100

Public Const XMinMapSize     As Byte = 1

Public Const YMaxMapSize     As Byte = 100

Public Const YMinMapSize     As Byte = 1

Private Const GrhFogata      As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Encabezado bmp
Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

'Posicion en un mapa
Public Type Position

    X As Long
    Y As Long

End Type

'Posicion en el Mundo
Public Type WorldPos

    Map As Integer
    X As Integer
    Y As Integer

End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData

    SX As Integer
    SY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single

End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh

    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer

End Type

'Lista de cuerpos
Public Type BodyData

    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

    Head(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Apariencia del personaje
Public Type Char

    Movement As Boolean
    active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    FX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    Atacable As Byte
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
    ParticleIndex As Integer

End Type

'Info de un objeto
Public Type Obj

    OBJIndex As Integer
    Amount As Integer

End Type

'Tipo de las celdas del mapa
Public Type MapBlock

    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    Engine_Light(0 To 3) As Long 'Standelf, Light Engine.

End Type

'Info de cada mapa
Public Type MapInfo

    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer

End Type

'Bordes del mapa
Public MinXBorder              As Byte

Public MaxXBorder              As Byte

Public MinYBorder              As Byte

Public MaxYBorder              As Byte

'Status del user
Public CurMap                  As Integer 'Mapa actual

Public UserIndex               As Integer

Public UserMoving              As Byte

Public UserBody                As Integer

Public UserHead                As Integer

Public UserPos                 As Position 'Posicion

Public AddtoUserPos            As Position 'Si se mueve

Public UserCharIndex           As Integer

Public EngineRun               As Boolean

Public FPS                     As Long

Public FramesPerSecCounter     As Long

Public FPSLastCheck            As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth        As Integer

Private WindowTileHeight       As Integer

Public HalfWindowTileWidth     As Integer

Public HalfWindowTileHeight    As Integer

'Offset del desde 0,0 del main view
Private MainViewTop            As Integer

Private MainViewLeft           As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize          As Integer

Private TileBufferPixelOffsetX As Integer

Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight         As Integer

Public TilePixelWidth          As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX   As Integer

Public ScrollPixelsPerFrameY   As Integer

Dim timerElapsedTime           As Single

Dim timerTicksPerFrame         As Single

Dim engineBaseSpeed            As Single

Public NumBodies               As Integer

Public Numheads                As Integer

Public NumFxs                  As Integer

Public NumChars                As Integer

Public LastChar                As Integer

Public NumWeaponAnims          As Integer

Public NumShieldAnims          As Integer

Private MainDestRect           As RECT

Private MainViewRect           As RECT

Private BackBufferRect         As RECT

Private MainViewWidth          As Integer

Private MainViewHeight         As Integer

Private MouseTileX             As Byte

Private MouseTileY             As Byte

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData()               As GrhData 'Guarda todos los grh

Public BodyData()              As BodyData

Public HeadData()              As HeadData

Public FxData()                As tIndiceFx

Public WeaponAnimData()        As WeaponAnimData

Public ShieldAnimData()        As ShieldAnimData

Public CascoAnimData()         As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData()               As MapBlock ' Mapa

Public MapInfo                 As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public Normal_RGBList(0 To 3)  As Long

'   Control de Lluvia
Public bRain                   As Boolean

Public bTecho                  As Boolean 'hay techo?

Public brstTick                As Long

Public bFogata                 As Boolean

Private iFrameIndex            As Byte  'Frame actual de la LL

Private llTick                 As Long  'Contador

Public charlist(1 To 10000)    As Char

' Used by GetTextExtentPoint32
Private Type Size

    cx As Long
    cy As Long

End Type

'[CODE 001]:MatuX
Public Enum PlayLoop

    plNone = 0
    plLluviain = 1
    plLluviaout = 2

End Enum

'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 _
                Lib "gdi32" _
                Alias "GetTextExtentPoint32A" (ByVal hDC As Long, _
                                               ByVal lpsz As String, _
                                               ByVal cbString As Long, _
                                               lpSize As Size) As Long

Private Declare Function SetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, _
                  ByVal viewPortY As Integer, _
                  ByRef tX As Byte, _
                  ByRef tY As Byte)
    '******************************************
    'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
    '******************************************
    tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2

End Sub

Sub MakeChar(ByVal CharIndex As Integer, _
             ByVal Body As Integer, _
             ByVal Head As Integer, _
             ByVal Heading As Byte, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal Arma As Integer, _
             ByVal Escudo As Integer, _
             ByVal Casco As Integer)

    On Error Resume Next

    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)

        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .active = 1

    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex

End Sub

Public Sub InitGrh(ByRef Grh As Grh, _
                   ByVal GrhIndex As Integer, _
                   Optional ByVal Started As Byte = 2)
    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0

        End If

    Else

        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started

    End If
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0

    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed

End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)

    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************
    Dim addx As Integer

    Dim addy As Integer

    Dim X    As Integer

    Dim Y    As Integer

    Dim nX   As Integer

    Dim nY   As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading

            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1

        End Select
        
        nX = X + addx
        nY = Y + addy
        
        If nX <= 0 Then nX = 1
        If nY <= 0 Then nY = 1
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        
        If (X Or Y) = 0 Then Exit Sub
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy

    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call Char_Erase(CharIndex)

        End If

    End If

End Sub

Public Sub DoFogataFx()

    Dim Location As Position
    
    If bFogata Then
        bFogata = HayFogata(Location)

        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0

        End If

    Else
        bFogata = HayFogata(Location)

        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", Location.X, Location.Y, LoopStyle.Enabled)

    End If

End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 09/21/2010
    ' 09/21/2010: C4b3z0n - Changed from Private Funtion tu Public Function.
    '***************************************************
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder

    End With

End Function

Sub DoPasosFx(ByVal CharIndex As Integer)

    If Not UserNavegando Then

        With charlist(CharIndex)

            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)

                End If

            End If

        End With

    Else
        ' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)

    End If

End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

    On Error Resume Next

    Dim X        As Integer

    Dim Y        As Integer

    Dim addx     As Integer

    Dim addy     As Integer

    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH

        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0

        End If

    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call Char_Erase(CharIndex)

    End If

End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)

    '******************************************
    'Starts the screen moving in a direction
    '******************************************
    Dim X  As Integer

    Dim Y  As Integer

    Dim tX As Integer

    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading

        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1

    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)

    End If

End Sub

Private Function HayFogata(ByRef Location As Position) As Boolean

    Dim j As Long

    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6

            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    Location.X = j
                    Location.Y = k
                    
                    HayFogata = True
                    Exit Function

                End If

            End If

        Next k
    Next j

End Function

Function NextOpenChar() As Integer

    '*****************************************************************
    'Finds next open char slot in CharList
    '*****************************************************************
    Dim LoopC As Long

    Dim Dale  As Boolean
    
    LoopC = 1

    Do While charlist(LoopC).active And Dale
        LoopC = LoopC + 1
        Dale = (LoopC <= UBound(charlist))
    Loop
    
    NextOpenChar = LoopC

End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is legal
    '*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function

    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function

    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function

    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function

    End If
    
    LegalPos = True

End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 01/08/2009
    'Checks to see if a tile position is legal, including if there is a casper in the tile
    '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
    '01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
    '*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function

    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function

    End If
    
    CharIndex = MapData(X, Y).CharIndex

    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function

        End If
        
        With charlist(CharIndex)

            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else

                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else

                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function

                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function

                End If

            End If

        End With

    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function

    End If
    
    MoveToLegalPos = True

End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function

    End If
    
    InMapBounds = True

End Function

Function GetBitmapDimensions(ByVal BmpFile As String, _
                             ByRef bmWidth As Long, _
                             ByRef bmHeight As Long)

    '*****************************************************************
    'Gets the dimensions of a bmp
    '*****************************************************************
    Dim BMHeader    As BITMAPFILEHEADER

    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight

End Function

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, _
                                   ByVal srchdc As Long, _
                                   ByRef SourceRect As RECT, _
                                   ByRef destRect As RECT, _
                                   ByVal TransparentColor As Long)

    '**************************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 27/07/2012 - ^[GS]^
    '*************************************************************
    Dim Color As Long

    Dim X     As Long

    Dim Y     As Long
    
    For X = SourceRect.Left To SourceRect.Right
        For Y = SourceRect.Top To SourceRect.Bottom
            Color = GetPixel(srchdc, X, Y)
            
            If Color <> TransparentColor Then
                Call SetPixel(dsthdc, destRect.Left + (X - SourceRect.Left), destRect.Top + (Y - SourceRect.Top), Color)

            End If

        Next Y
    Next X

End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, _
                              ByRef Picture As StdPicture, _
                              ByVal X1 As Single, _
                              ByVal Y1 As Single, _
                              Optional Width1, _
                              Optional Height1, _
                              Optional X2, _
                              Optional Y2, _
                              Optional Width2, _
                              Optional Height2)
    '**************************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 12/28/2009
    'Draw Picture in the PictureBox
    '*************************************************************

    Call PictureBox.PaintPicture(Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2)

End Sub

Sub RenderScreen(ByVal tilex As Integer, _
                 ByVal tiley As Integer, _
                 ByVal PixelOffsetX As Integer, _
                 ByVal PixelOffsetY As Integer)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************
    Dim Y                As Long     'Keeps track of where on map we are

    Dim X                As Long     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim maxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim maxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim ColorTechos(3)   As Long

    Dim ElapsedTime      As Single
    
    ElapsedTime = Engine_ElapsedTime()
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - Engine_Get_TileBuffer
    maxY = screenmaxY + Engine_Get_TileBuffer
    minX = screenminX - Engine_Get_TileBuffer
    maxX = screenmaxX + Engine_Get_TileBuffer
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    ParticleOffsetX = (Engine_PixelPosX(screenminX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(screenminY) - PixelOffsetY)

    '<----- Layer 1, 2 ----->
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
        
            If Map_InBounds(X, Y) Then
                'Layer 1
                Call DDrawGrhtoSurface(MapData(X, Y).Graphic(1), (ScreenX - 1) * TilePixelWidth + PixelOffsetX, (ScreenY - 1) * TilePixelHeight + PixelOffsetY, 0, 1, X, Y)
                    
                'Layer 2
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), (ScreenX - 1) * TilePixelWidth + PixelOffsetX, (ScreenY - 1) * TilePixelHeight + PixelOffsetY, 0, MapData(X, Y).Engine_Light(), 0, X, Y)

                End If

            End If
            
            ScreenX = ScreenX + 1
        Next X

        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - Engine_Get_TileBuffer

    For Y = minY To maxY
        ScreenX = minXOffset - Engine_Get_TileBuffer

        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            
            If Map_InBounds(X, Y) Then

                With MapData(X, Y)

                    'Object Layer
                    If .ObjGrh.GrhIndex <> 0 Then
                        Call DDrawTransGrhtoSurface(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1, X, Y)

                    End If
                    
                    'Char layer
                    If .CharIndex <> 0 Then
                        Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)

                    End If
                    
                    'Layer 3
                    If .Graphic(3).GrhIndex <> 0 Then
                    
                        If .Graphic(3).GrhIndex = 735 Or .Graphic(3).GrhIndex >= 6994 And .Graphic(3).GrhIndex <= 7002 Then
                            If Abs(UserPos.X - X) < 4 And (Abs(UserPos.Y - Y)) < 4 Then
                                Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1, X, Y, True)
                            Else 'NORMAL
                                Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1, X, Y)
    
                            End If

                        Else 'NORMAL
                            Call DDrawTransGrhtoSurface(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1, X, Y)

                        End If

                    End If

                End With

            End If
            
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y
    
    '<----- Layer 4 ----->
    ScreenY = minYOffset - Engine_Get_TileBuffer

    For Y = minY To maxY
        ScreenX = minXOffset - Engine_Get_TileBuffer

        For X = minX To maxX

            If Map_InBounds(X, Y) Then
                'If Abs(MouseTileX - X) < 1 And (Abs(MouseTileY - Y)) < 1 And Settings.NombreItems And MapData(X, Y).OBJInfo.Name <> "" Then
                'Engine_Draw_Box ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, Fonts_Render_String_Width(MapData(X, Y).OBJInfo.Name, Settings.Engine_Font) + 1, Fuentes(Settings.Engine_Font).CharactersHeight, D3DColorARGB(100, 0, 0, 0)
                'Fonts_Render_String MapData(X, Y).OBJInfo.Name, ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, D3DColorARGB(100, 255, 255, 255), Settings.Engine_Font
                        
                'End If
                    
                'Layer 4
                If Not bTecho Then
                    If MapData(X, Y).Graphic(4).GrhIndex Then
                        Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, 1, MapData(X, Y).Engine_Light(), 1, X, Y)

                    End If
                        
                Else

                    If MapData(X, Y).Graphic(4).GrhIndex Then
                        Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, 1, MapData(X, Y).Engine_Light(), 1, X, Y, True)

                    End If

                End If

            End If

            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    'Weather Update & Render
    Call Engine_Weather_Update
    
    'Effects Update
    Call Effect_UpdateAll
    
    If ClientSetup.ProyectileEngine = True Then

        Dim j As Integer
        
        If LastProjectile > 0 Then

            For j = 1 To LastProjectile

                If ProjectileList(j).Grh.GrhIndex Then

                    Dim Angle As Single

                    'Update the position
                    Angle = DegreeToRadian * Engine_GetAngle(ProjectileList(j).X, ProjectileList(j).Y, ProjectileList(j).tX, ProjectileList(j).tY)
                    ProjectileList(j).X = ProjectileList(j).X + (Sin(Angle) * ElapsedTime * 0.63)
                    ProjectileList(j).Y = ProjectileList(j).Y - (Cos(Angle) * ElapsedTime * 0.63)
                    
                    'Update the rotation
                    If ProjectileList(j).RotateSpeed > 0 Then
                        ProjectileList(j).Rotate = ProjectileList(j).Rotate + (ProjectileList(j).RotateSpeed * ElapsedTime * 0.01)

                        Do While ProjectileList(j).Rotate > 360
                            ProjectileList(j).Rotate = ProjectileList(j).Rotate - 360
                        Loop

                    End If
    
                    'Draw if within range
                    X = ((-minX - 1) * 32) + ProjectileList(j).X + PixelOffsetX + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(j).OffsetX
                    Y = ((-minY - 1) * 32) + ProjectileList(j).Y + PixelOffsetY + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(j).OffsetY

                    If Y >= -32 Then
                        If Y <= (ScreenHeight + 32) Then
                            If X >= -32 Then
                                If X <= (ScreenWidth + 32) Then
                                    If ProjectileList(j).Rotate = 0 Then
                                        DDrawTransGrhtoSurface ProjectileList(j).Grh, X, Y, 0, MapData(50, 50).Engine_Light(), 0, 50, 50, True, 0
                                    Else
                                        DDrawTransGrhtoSurface ProjectileList(j).Grh, X, Y, 0, MapData(50, 50).Engine_Light(), 0, 50, 50, True, ProjectileList(j).Rotate

                                    End If

                                End If

                            End If

                        End If

                    End If
                    
                End If

            Next j
            
            'Check if it is close enough to the target to remove
            For j = 1 To LastProjectile

                If ProjectileList(j).Grh.GrhIndex Then
                    If Abs(ProjectileList(j).X - ProjectileList(j).tX) < 20 Then
                        If Abs(ProjectileList(j).Y - ProjectileList(j).tY) < 20 Then
                            Engine_Projectile_Erase j

                        End If

                    End If

                End If

            Next j
            
        End If

    End If
    
    '   Set Offsets
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    
    If ClientSetup.PartyMembers Then Call Draw_Party_Members
    Call RenderCount

End Sub

Public Function RenderSounds()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 3/30/2008
    'Actualiza todos los sonidos del mapa.
    '**************************************************************

    Dim Location As Position

    If bFogata Then
        bFogata = Map_CheckBonfire(Location)

        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0

        End If

    Else
        bFogata = Map_CheckBonfire(Location)

        If bFogata And FogataBufferIndex = 0 Then
            FogataBufferIndex = Audio.PlayWave("fuego.wav", Location.X, Location.Y, LoopStyle.Enabled)

        End If

    End If

End Function

Function HayUserAbajo(ByVal X As Integer, _
                      ByVal Y As Integer, _
                      ByVal GrhIndex As Long) As Boolean

    If GrhIndex > 0 Then
        HayUserAbajo = charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) And charlist(UserCharIndex).Pos.Y <= Y

    End If

End Function

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, _
                               ByVal setTilePixelHeight As Integer, _
                               ByVal setTilePixelWidth As Integer, _
                               ByVal pixelsToScrollPerFrameX As Integer, _
                               pixelsToScrollPerFrameY As Integer) As Boolean
    '***************************************************
    'Author: Aaron Perkins
    'Last Modification: 08/14/07
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Configures the engine to start running.
    '***************************************************
    
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = Round(frmMain.MainViewPic.Height / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.Width / 32, 0)
    
    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2

    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

    On Error GoTo 0

    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call LoadGraphics
    
    'Index Buffer. Dunkan
    indexList(0) = 0: indexList(1) = 1: indexList(2) = 2
    indexList(3) = 3: indexList(4) = 4: indexList(5) = 5
    
    Set ibQuad = DirectDevice.CreateIndexBuffer(Len(indexList(0)) * 4, 0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
    D3DIndexBuffer8SetData ibQuad, 0, Len(indexList(0)) * 4, 0, indexList(0)
    
    Set vbQuadIdx = DirectDevice.CreateVertexBuffer(Len(temp_verts(0)) * 4, 0, D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1, D3DPOOL_MANAGED)
    
    InitTileEngine = True

End Function

Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, DirGraficos, ClientSetup.byMemory)

End Sub

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, _
                  ByVal DisplayFormLeft As Integer, _
                  ByVal MouseViewX As Integer, _
                  ByVal MouseViewY As Integer)

    If EngineRun Then
        Engine_BeginScene
        
        If UserMoving Then

            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame

                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False

                End If

            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame

                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False

                End If

            End If

        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
        '****** Update screen ******
        If UserCiego Then
            DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
            
        Else
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX - ZoomOffset(1), OffsetCounterY - ZoomOffset(0))

        End If
        
        Call Dialogos.Render
        Call DibujarCartel
        
        '     Calculamos los FPS y los mostramos
        Call Engine_Update_FPS
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_Get_BaseSpeed
        
        Engine_EndScene MainScreenRect, 0

    End If
    
    '//Banco
    If frmBancoObj.PicBancoInv.Visible Then Call InvBanco(0).DrawInv
         
    If frmBancoObj.PicInv.Visible Then Call InvBanco(1).DrawInv
    
    '//Comercio
    If frmComerciar.picInvNpc.Visible Then Call InvComNpc.DrawInv
        
    If frmComerciar.picInvUser.Visible Then Call InvComUsu.DrawInv
    
    '//Comercio entre usuarios
    If frmComerciarUsu.picInvComercio.Visible Then InvComUsu.DrawInv (1)
            
    If frmComerciarUsu.picInvOfertaProp.Visible Then InvOfferComUsu(0).DrawInv (1)
            
    If frmComerciarUsu.picInvOfertaOtro Then InvOfferComUsu(1).DrawInv (1)
            
    If frmComerciarUsu.picInvOroProp.Visible Then InvOroComUsu(0).DrawInv (1)
            
    If frmComerciarUsu.picInvOroOfertaProp.Visible Then InvOroComUsu(1).DrawInv (1)
                
    If frmComerciarUsu.picInvOroOfertaOtro.Visible Then InvOroComUsu(2).DrawInv (1)
        
    '//Herrero
    If frmHerrero.Visible Then

        With frmHerrero

            If .picLingotes0.Visible Or .picMejorar0.Visible Then InvLingosHerreria(1).DrawInv (1)
            
            If .picLingotes1.Visible Or .picMejorar1.Visible Then InvLingosHerreria(2).DrawInv (1)
            
            If .picLingotes2.Visible Or .picMejorar2.Visible Then InvLingosHerreria(3).DrawInv (1)
            
            If .picLingotes3.Visible Or .picMejorar3.Visible Then InvLingosHerreria(4).DrawInv (1)

        End With

    End If
        
    '//FIN HERRERO
    
    '//Carpintero
    If frmCarp.Visible Then

        With frmCarp

            If .picMaderas0.Visible Or .imgMejorar0.Visible Then InvMaderasCarpinteria(1).DrawInv (1)
                
            If .picMaderas1.Visible Or .imgMejorar1.Visible Then InvMaderasCarpinteria(2).DrawInv (1)
            
            If .picMaderas2.Visible Or .imgMejorar2.Visible Then InvMaderasCarpinteria(3).DrawInv (1)
            
            If .picMaderas3.Visible Or .imgMejorar3.Visible Then InvMaderasCarpinteria(4).DrawInv (1)

        End With

    End If

    '//Inventario
    If frmMain.Visible Then Call Inventario.DrawInv
    
End Sub

Public Sub RenderText(ByVal lngXPos As Integer, _
                      ByVal lngYPos As Integer, _
                      ByRef strText As String, _
                      ByVal lngColor As Long, _
                      ByRef Font As StdFont)

    If strText <> "" Then
        'Call BackBufferSurface.SetForeColor(vbBlack)
        'Call BackBufferSurface.SetFont(Font)
        ' Call BackBufferSurface.DrawText(lngXPos - 2, lngYPos - 1, strText, False)
        
        'Call BackBufferSurface.SetForeColor(lngColor)
        'Call BackBufferSurface.DrawText(lngXPos, lngYPos, strText, False)
    End If

End Sub

Private Function GetElapsedTime() As Single

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim Start_Time    As Currency

    Static end_time   As Currency

    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq

    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

End Function

Private Sub CharRender(ByVal CharIndex As Long, _
                       ByVal PixelOffsetX As Integer, _
                       ByVal PixelOffsetY As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 16/09/2010 (Zama)
    'Draw char's to screen without offcentering them
    '16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
    '***************************************************
    Dim moved As Boolean

    Dim Pos   As Integer

    Dim line  As String

    Dim Color As Long
    
    With charlist(CharIndex)

        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            
            '//Evito runtime
            If Not .Heading <> 0 Then .Heading = EAST
            
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            '//Movimiento del arma y el escudo
            If Not .Movement Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
                
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 1

            End If
            
            .Moving = False

        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        Dim ColorFinal(0 To 3) As Long

        Dim RenderSpell        As Boolean

        Dim OffSetName         As Integer
        
        If Not .muerto Then
            If Abs(MouseTileX - .Pos.X) < 1 And (Abs(MouseTileY - .Pos.Y)) < 1 And CharIndex <> UserCharIndex And ClientSetup.TonalidadPJ Then
                If .Nombre <> "" Then
                    Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorXRGB(0, 255, 0))
                Else
                    ColorFinal(0) = MapData(.Pos.X, .Pos.Y).Engine_Light(0)
                    ColorFinal(1) = MapData(.Pos.X, .Pos.Y).Engine_Light(1)
                    ColorFinal(2) = MapData(.Pos.X, .Pos.Y).Engine_Light(2)
                    ColorFinal(3) = MapData(.Pos.X, .Pos.Y).Engine_Light(3)

                End If

                RenderSpell = True
            Else
                ColorFinal(0) = MapData(.Pos.X, .Pos.Y).Engine_Light(0)
                ColorFinal(1) = MapData(.Pos.X, .Pos.Y).Engine_Light(1)
                ColorFinal(2) = MapData(.Pos.X, .Pos.Y).Engine_Light(2)
                ColorFinal(3) = MapData(.Pos.X, .Pos.Y).Engine_Light(3)

            End If

        Else

            If esGM(Val(CharIndex)) Then
                Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(150, 200, 200, 0))
            Else

                If .Criminal Then
                    Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(100, 255, 100, 100))
                Else
                    Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(100, 128, 255, 255))

                End If

            End If

        End If
        
        If Not .invisible Then
            Movement_Speed = 0.5

            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.X, .Pos.Y, False, 0)
            
            'Draw Head
            If .Head.Head(.Heading).GrhIndex Then
                Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, ColorFinal(), 0, .Pos.X, .Pos.Y)
                
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, ColorFinal(), 0, .Pos.X, .Pos.Y)
                
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.X, .Pos.Y)

                End If
                
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then
                    Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.X, .Pos.Y)

                End If
            
                'Draw name over head
                If LenB(.Nombre) > 0 Then
                    If Nombres Then
                        Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)

                    End If

                End If
            
            Else 'Usuario invisible
        
                If CharIndex = UserCharIndex Or mid$(charlist(CharIndex).Nombre, getTagPosition(.Nombre)) = mid$(charlist(UserCharIndex).Nombre, getTagPosition(charlist(UserCharIndex).Nombre)) And Len(mid$(charlist(CharIndex).Nombre, getTagPosition(.Nombre))) > 0 Then
                
                    Movement_Speed = 0.5
                
                    'Draw Body
                    If .Body.Walk(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.X, .Pos.Y, True, 0)
                
                    'Draw Head
                    If .Head.Head(.Heading).GrhIndex Then
                        Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, ColorFinal(), 0, .Pos.X, .Pos.Y, True)
                    
                        'Draw Helmet
                        If .Casco.Head(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, ColorFinal(), 0, .Pos.X, .Pos.Y, True)
                    
                        'Draw Weapon
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.X, .Pos.Y, True)
                    
                        'Draw Shield
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.X, .Pos.Y, True)
                
                        'Draw name over head
                        If LenB(.Nombre) > 0 Then
                            If Nombres Then
                                Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY, True)

                            End If

                        End If
                    
                        'OffSetName = 35

                    End If

                End If

            End If

        End If
        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex) '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        
        Movement_Speed = 1

        'Draw FX
        If .FxIndex <> 0 Then
            Call DDrawTransGrhtoSurface(.FX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, MapData(.Pos.X, .Pos.Y).Engine_Light(), 1, .Pos.X, .Pos.Y, False)
            
            'Check if animation is over
            If .FX.Started = 0 Then .FxIndex = 0

        End If
        
    End With

End Sub

Private Sub RenderName(ByVal CharIndex As Long, _
                       ByVal X As Integer, _
                       ByVal Y As Integer, _
                       Optional ByVal Invi As Boolean = False)

    Dim Pos   As Integer

    Dim line  As String

    Dim Color As Long
   
    With charlist(CharIndex)
        Pos = getTagPosition(.Nombre)
    
        If .priv = 0 Then
            If .muerto Then
                Color = D3DColorARGB(255, 220, 220, 255)
            Else

                If .Criminal Then
                    Color = ColoresPJ(50)
                Else
                    Color = ColoresPJ(49)

                End If

            End If

        Else
            Color = ColoresPJ(.priv)

        End If
    
        If Invi = True Then
            Color = D3DColorARGB(180, 150, 180, 220)

        End If
            
        'Nick
        line = Left$(.Nombre, Pos - 2)
        'Fonts_Render_String line, (X + 16) - Fonts_Render_String_Width(line, Settings.Engine_Name_Font) / 2, Y + 30, color, Settings.Engine_Name_Font
        Call DrawText(X - (Len(line) * 6 / 2) + 14, Y + 30, line, Color)
            
        'Clan
        line = mid$(.Nombre, Pos)
        'Fonts_Render_String line, (X + 16) - Fonts_Render_String_Width(line, Settings.Engine_Name_Font) / 2, Y + 30 + Fuentes(Settings.Engine_Font).CharactersHeight, D3DColorXRGB(255, 230, 130), Settings.Engine_Name_Font
        Call DrawText(X - (Len(line) * 6 / 2) + 14, Y + 45, line, Color)

    End With

End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, _
                          ByVal FX As Integer, _
                          ByVal Loops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    With charlist(CharIndex)
        .FxIndex = FX
        
        If .FxIndex > 0 Then
            Call InitGrh(.FX, FxData(FX).Animacion)
        
            .FX.Loops = Loops

        End If

    End With

End Sub

Public Sub Device_Textured_Render(ByVal X As Integer, _
                                  ByVal Y As Integer, _
                                  ByVal Texture As Direct3DTexture8, _
                                  ByRef src_rect As RECT, _
                                  ByRef Color_List() As Long, _
                                  Optional Alpha As Boolean = False, _
                                  Optional ByVal Angle As Single = 0, _
                                  Optional ByVal Shadow As Boolean = False)

    If Shadow And ClientSetup.UsarSombras = False Then Exit Sub
    
    Dim dest_rect     As RECT

    Dim temp_verts(3) As TLVERTEX

    Dim srdesc        As D3DSURFACE_DESC

    With dest_rect
        .Bottom = Y + (src_rect.Bottom - src_rect.Top)
        .Left = X
        .Right = X + (src_rect.Right - src_rect.Left)
        .Top = Y

    End With
    
    Dim texwidth As Long, texheight As Long

    Texture.GetLevelDesc 0, srdesc

    texwidth = srdesc.Width
    texheight = srdesc.Height
    
    If Shadow Then

        Dim Color_Shadow(3) As Long

        Engine_Long_To_RGB_List Color_Shadow(), D3DColorARGB(50, 0, 0, 0)
        Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color_Shadow(), texwidth, texheight, Angle
    Else
        Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color_List(), texwidth, texheight, Angle

    End If
    
    DirectDevice.SetTexture 0, Texture

    If Shadow Then
        temp_verts(1).X = temp_verts(1).X + (src_rect.Bottom - src_rect.Top) * 0.5
        temp_verts(1).Y = temp_verts(1).Y - (src_rect.Right - src_rect.Left) * 0.5
       
        temp_verts(3).X = temp_verts(3).X + (src_rect.Right - src_rect.Left)
        temp_verts(3).Y = temp_verts(3).Y - (src_rect.Right - src_rect.Left) * 0.5

    End If
    
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    End If
    
    ' Medium load.
    'DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
        
    ' Faster load.
    DirectDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, indexList(0), D3DFMT_INDEX16, temp_verts(0), Len(temp_verts(0))
                
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If

End Sub

Public Sub Device_Textured_Render_Scale(ByVal X As Integer, _
                                        ByVal Y As Integer, _
                                        ByVal Texture As Direct3DTexture8, _
                                        ByRef src_rect As RECT, _
                                        ByRef Color_List() As Long, _
                                        Optional Alpha As Boolean = False, _
                                        Optional ByVal Angle As Single = 0, _
                                        Optional ByVal Shadow As Boolean = False)

    If Shadow And ClientSetup.UsarSombras = False Then Exit Sub
    
    Dim dest_rect     As RECT

    Dim temp_verts(3) As TLVERTEX

    Dim srdesc        As D3DSURFACE_DESC

    With dest_rect
        .Bottom = Y + 2 '(src_rect.bottom - src_rect.Top)
        .Left = X
        .Right = X + 2 '(src_rect.Right - src_rect.Left)
        .Top = Y

    End With
    
    Dim texwidth As Long, texheight As Long

    Texture.GetLevelDesc 0, srdesc

    texwidth = srdesc.Width
    texheight = srdesc.Height
    
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color_List(), texwidth, texheight, Angle
    
    DirectDevice.SetTexture 0, Texture
    
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    End If
        
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
        
    If Alpha Then
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If

End Sub

Public Sub RenderItem(ByVal hWndDest As Long, ByVal GrhIndex As Long)

    Dim DR As RECT
    
    With DR
        .Left = 0
        .Top = 0
        .Right = 32
        .Bottom = 32

    End With
    
    Engine_BeginScene

    Call DDrawTransGrhIndextoSurface(GrhIndex, 0, 0, 0, Normal_RGBList(), 0, False)
        
    Engine_EndScene DR, hWndDest
    
End Sub

Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, _
                              ByVal X As Integer, _
                              ByVal Y As Integer, _
                              ByVal Center As Byte, _
                              ByVal Animate As Byte, _
                              ByVal posX As Byte, _
                              ByVal posY As Byte)

    Dim CurrentGrhIndex As Integer

    Dim SourceRect      As RECT
    
    If Grh.GrhIndex = 0 Then Exit Sub

    On Error GoTo error
        
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0

                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, MapData(posX, posY).Engine_Light(), False, 0, False)

    End With

    Exit Sub

error:

    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in DDrawGrhtoSurface, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Gráfico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call CloseClient

    End If

End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, _
                                ByVal X As Integer, _
                                ByVal Y As Integer, _
                                ByVal Center As Byte, _
                                ByRef Color_List() As Long, _
                                Optional ByVal Angle As Single = 0, _
                                Optional ByVal Alpha As Boolean = False)

    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, Color_List(), Alpha, Angle, False)

    End With

End Sub

Public Sub DDrawGrhtoSurfaceScale(ByRef Grh As Grh, _
                                  ByVal X As Integer, _
                                  ByVal Y As Integer, _
                                  ByVal Center As Byte, _
                                  ByVal Animate As Byte, _
                                  ByVal posX As Byte, _
                                  ByVal posY As Byte)

    Dim CurrentGrhIndex As Integer

    Dim SourceRect      As RECT
    
    If Grh.GrhIndex = 0 Then Exit Sub

    On Error GoTo error
        
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0

                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render_Scale(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, MapData(posX, posY).Engine_Light(), False, 0, False)

    End With

    Exit Sub

error:

    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in DDrawGrhtoSurface, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Gráfico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call CloseClient

    End If

End Sub

Sub DDrawTransGrhIndextoSurfaceScale(ByVal GrhIndex As Integer, _
                                     ByVal X As Integer, _
                                     ByVal Y As Integer, _
                                     ByVal Center As Byte, _
                                     ByRef Color_List() As Long, _
                                     Optional ByVal Angle As Single = 0, _
                                     Optional ByVal Alpha As Boolean = False)

    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
        
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render_Scale(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, Color_List(), Alpha, Angle, False)

    End With

End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, _
                           ByVal X As Integer, _
                           ByVal Y As Integer, _
                           ByVal Center As Byte, _
                           ByRef Color_List() As Long, _
                           ByVal Animate As Byte, _
                           ByVal posX As Byte, _
                           ByVal posY As Byte, _
                           Optional ByVal Alpha As Boolean = False, _
                           Optional ByVal Angle As Single = 0, _
                           Optional ByVal Shadow As Boolean = False)

    '*****************************************************************
    'Draws a GRH transparently to a X and Y position
    '*****************************************************************
    Dim CurrentGrhIndex As Integer

    Dim SourceRect      As RECT
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
    On Error GoTo error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Movement_Speed
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0

                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
                
        SourceRect.Left = .SX
        SourceRect.Top = .SY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.Bottom = SourceRect.Top + .pixelHeight
        
        Call Device_Textured_Render(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, Color_List(), Alpha, Angle, False)

    End With

    Exit Sub

error:

    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in DDrawTransGrhtoSurface, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Gráfico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call CloseClient

    End If

End Sub

Public Function GrhCheck(ByVal GrhIndex As Long) As Boolean
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check grh_index

    If GrhIndex > 0 And GrhIndex <= UBound(GrhData()) Then
        GrhCheck = GrhData(GrhIndex).NumFrames

    End If

End Function

Public Sub GrhUninitialize(Grh As Grh)
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Resets a Grh
    '*****************************************************************

    With Grh
        
        'Copy of parameters
        .GrhIndex = 0
        .Started = False
        .Loops = 0
        
        'Set frame counters
        .FrameCounter = 0
        .Speed = 0
                
    End With

End Sub
