Attribute VB_Name = "mDx8_Engine"
Option Explicit

' No matter what you do with DirectX8, you will need to start with
' the DirectX8 object. You will need to create a new instance of
' the object, using the New keyword, rather than just getting a
' pointer to it, since there's nowhere to get a pointer from yet (duh!).
Public DirectX              As New DirectX8

' The D3DX8 object contains lots of helper functions, mostly math
' to make Direct3D alot easier to use. Notice we create a new
' instance of the object using the New keyword.
Public DirectD3D8           As D3DX8

Public DirectD3D            As Direct3D8

' The Direct3DDevice8 represents our rendering device, which could
' be a hardware or a software device. The great thing is we still
' use the same object no matter what it is
Public DirectDevice         As Direct3DDevice8

' The D3DDISPLAYMODE type structure that holds
' the information about your current display adapter.
Public DispMode             As D3DDISPLAYMODE

' The D3DPRESENT_PARAMETERS type holds a description of the way
' in which DirectX will display it's rendering.
Public D3DWindow As D3DPRESENT_PARAMETERS

Public SurfaceDB            As New clsSurfaceManager

Public Engine_BaseSpeed     As Single

Public TileBufferSize       As Integer

Public ScreenWidth          As Long

Public ScreenHeight         As Long

Public MainScreenRect       As RECT

'
Public Type TLVERTEX

    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single

End Type

Private EndTime As Long

Public Sub Engine_DirectX8_Init()

    On Error GoTo EngineHandler:

    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8

    If ClientSetup.OverrideVertexProcess > 0 Then
        
        Select Case ClientSetup.OverrideVertexProcess
            
            Case 1:

                If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then _
                    Call MsgBox("No se pudo inicializar el motor grafico. Por favor, verifique si tiene sus librerias y sus controladores actualizados.")
            
            Case 2:

                If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then _
                    Call MsgBox("No se pudo inicializar el motor grafico. Por favor, verifique si tiene sus librerias y sus controladores actualizados.")
            
            Case 3:

                If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then _
                    Call MsgBox("No se pudo inicializar el motor grafico. Por favor, verifique si tiene sus librerias y sus controladores actualizados.")

        End Select
        
    Else

        'Detectamos el modo de renderizado mas compatible con tu PC.
        If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
            
                    Call MsgBox("No se pudo inicializar el motor grafico. Por favor, verifique si tiene sus librerias y sus controladores actualizados.")
                
                    End
                
                End If

            End If

        End If

    End If

    Engine_Init_FontTextures
    Engine_Init_FontSettings
    
    ' Set rendering options
    Call Engine_Init_RenderStates
    
    EndTime = GetTickCount
    
    Exit Sub
EngineHandler:
    
    Call LogError(Err.number, Err.Description, "mDx8_Engine.Engine_DirectX8")
    
    Call CloseClient

End Sub

Private Function Engine_Init_DirectDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
On Error GoTo ErrorDevice:

    'Establecemos cual va a ser el tamano del render.
    ScreenWidth = frmMain.MainViewPic.ScaleWidth
    ScreenHeight = frmMain.MainViewPic.ScaleHeight

    ' Retrieve the information about your current display adapter.
    Call DirectD3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    
    ' Fill the D3DPRESENT_PARAMETERS type, describing how DirectX should
    ' display it's renders.
    With D3DWindow
        .Windowed = True
        
        ' The swap effect determines how the graphics get from the backbuffer to the screen.
        ' D3DSWAPEFFECT_DISCARD:
        '   Means that every time the render is presented, the backbuffer
        '   image is destroyed, so everything must be rendered again.
        .SwapEffect = D3DSWAPEFFECT_DISCARD
        
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = ScreenWidth
        .BackBufferHeight = ScreenHeight
        .hDeviceWindow = frmMain.MainViewPic.hwnd

    End With
    
    If Not DirectDevice Is Nothing Then
        Set DirectDevice = Nothing
    End If
    
    ' Create the rendering device.
    ' Here we request a Hardware or Mixed rasterization.
    ' If your computer does not have this, the request may fail, so use
    ' D3DDEVTYPE_REF instead of D3DDEVTYPE_HAL if this happens. A real
    ' program would be able to detect an error and automatically switch device.
    ' We also request software vertex processing, which means the CPU has to
    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, D3DCREATEFLAGS, D3DWindow)
    
    'Lo pongo xq es bueno saberlo...
    Select Case D3DCREATEFLAGS
    
        Case D3DCREATE_MIXED_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: MIXED"
        
        Case D3DCREATE_HARDWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: HARDWARE"
            
        Case D3DCREATE_SOFTWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: SOFTWARE"
            
    End Select
    
    'Everything was successful
    Engine_Init_DirectDevice = True
    
    Exit Function
    
ErrorDevice:
    
    'Destroy the D3DDevice so it can be remade
    Set DirectDevice = Nothing

    'Return a failure
    Engine_Init_DirectDevice = False
    
End Function

Private Sub Engine_Init_RenderStates()

    'Set the render states
    With DirectDevice
    
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
        
    End With
    
End Sub

Public Sub Engine_DirectX8_End()

    '***************************************************
    'Author: Standelf
    'Last Modification: 26/05/2010
    'Destroys all DX objects
    '***************************************************
    On Error Resume Next

    Dim i As Byte
    
    ' DeInit Lights
    Call DeInit_LightEngine
    
    ' Clean Particles
    For i = 1 To UBound(ParticleTexture)
        If Not ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing
    Next i
    
    ' Clean Texture
    DirectDevice.SetTexture 0, Nothing

    ' Erase Data
    Erase MapData()
    Erase charlist()
    
    Set DirectD3D8 = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set DirectDevice = Nothing

End Sub

Public Sub Engine_DirectX8_Aditional_Init()
    '**************************************************************
    'Author: Standelf
    'Last Modify Date: 30/12/2010
    '**************************************************************

    FPS = 101
    FramesPerSecCounter = 101

    Call Engine_Set_TileBuffer(9)
    
    Engine_Set_BaseSpeed 0.018
    
    With MainScreenRect
        .Bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth

    End With

    Call Engine_Long_To_RGB_List(Normal_RGBList(), -1)

    Init_MeteoEngine
    Engine_Init_ParticleEngine
    
End Sub

Public Sub Engine_Draw_Line(X1 As Single, _
                            Y1 As Single, _
                            X2 As Single, _
                            Y2 As Single, _
                            Optional Color As Long = -1, _
                            Optional Color2 As Long = -1)

    On Error GoTo error

    Dim Vertex(1) As TLVERTEX

    Vertex(0) = Geometry_Create_TLVertex(X1, Y1, 0, 1, Color, 0, 0, 0)
    Vertex(1) = Geometry_Create_TLVertex(X2, Y2, 0, 1, Color2, 0, 0, 0)

    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, Vertex(0), Len(Vertex(0))
    Exit Sub

error:

    'Call Log_Engine("Error in Engine_Draw_Line, " & Err.Description & " (" & Err.number & ")")
End Sub

Public Sub Engine_Draw_Point(X1 As Single, Y1 As Single, Optional Color As Long = -1)

    On Error GoTo error

    Dim Vertex(0) As TLVERTEX

    Vertex(0) = Geometry_Create_TLVertex(X1, Y1, 0, 1, Color, 0, 0, 0)

    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_POINTLIST, 1, Vertex(0), Len(Vertex(0))
    Exit Sub

error:

    'Call Log_Engine("Error in Engine_Draw_Point, " & Err.Description & " (" & Err.number & ")")
End Sub

Public Function Engine_ElapsedTime() As Long

    '**************************************************************
    'Gets the time that past since the last call
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ElapsedTime
    '**************************************************************
    Dim Start_Time As Long

    'Get current time
    Start_Time = timeGetTime

    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - EndTime

    'Get next end time
    EndTime = Start_Time

End Function

Public Function Engine_PixelPosX(ByVal X As Integer) As Integer
    '*****************************************************************
    'Converts a tile position to a screen position
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosX
    '*****************************************************************

    Engine_PixelPosX = (X - 1) * 32
    
End Function

Public Function Engine_PixelPosY(ByVal Y As Integer) As Integer
    '*****************************************************************
    'Converts a tile position to a screen position
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosY
    '*****************************************************************

    Engine_PixelPosY = (Y - 1) * 32
    
End Function

Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
    '************************************************************
    'Tile Position to Screen Position
    'Takes the tile position and returns the pixel location on the screen
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPX
    '************************************************************

    Engine_TPtoSPX = Engine_PixelPosX(X - ((UserPos.X - HalfWindowTileWidth) - Engine_Get_TileBuffer)) + OffsetCounterX - 272 + ((10 - TileBufferSize) * 32)
    
End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
    '************************************************************
    'Tile Position to Screen Position
    'Takes the tile position and returns the pixel location on the screen
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPY
    '************************************************************

    Engine_TPtoSPY = Engine_PixelPosY(Y - ((UserPos.Y - HalfWindowTileHeight) - Engine_Get_TileBuffer)) + OffsetCounterY - 272 + ((10 - TileBufferSize) * 32)
    
End Function

Public Sub Engine_Draw_Box(ByVal X As Integer, _
                           ByVal Y As Integer, _
                           ByVal Width As Integer, _
                           ByVal Height As Integer, _
                           Color As Long)

    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 29/12/10
    'Blisse-AO | Render Box
    '***************************************************
    Dim b_Rect           As RECT

    Dim b_Color(0 To 3)  As Long

    Dim b_Vertex(0 To 3) As TLVERTEX
    
    Engine_Long_To_RGB_List b_Color(), Color

    With b_Rect
        .Bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y

    End With

    Geometry_Create_Box b_Vertex(), b_Rect, b_Rect, b_Color(), 0, 0
    
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, b_Vertex(0), Len(b_Vertex(0))

End Sub

Public Sub Engine_D3DColor_To_RGB_List(RGB_List() As Long, Color As D3DCOLORVALUE)
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 14/05/10
    'Blisse-AO | Set a D3DColorValue to a RGB List
    '***************************************************
    RGB_List(0) = D3DColorARGB(Color.a, Color.r, Color.g, Color.b)
    RGB_List(1) = RGB_List(0)
    RGB_List(2) = RGB_List(0)
    RGB_List(3) = RGB_List(0)

End Sub

Public Sub Engine_Long_To_RGB_List(RGB_List() As Long, long_color As Long)
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 16/05/10
    'Blisse-AO | Set a Long Color to a RGB List
    '***************************************************
    RGB_List(0) = long_color
    RGB_List(1) = RGB_List(0)
    RGB_List(2) = RGB_List(0)
    RGB_List(3) = RGB_List(0)

End Sub

Private Function Engine_Collision_Between(ByVal Value As Single, _
                                          ByVal Bound1 As Single, _
                                          ByVal Bound2 As Single) As Byte
    '*****************************************************************
    'Find if a value is between two other values (used for line collision)
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Between
    '*****************************************************************

    'Checks if a value lies between two bounds
    If Bound1 > Bound2 Then
        If Value >= Bound2 Then
            If Value <= Bound1 Then Engine_Collision_Between = 1

        End If

    Else

        If Value >= Bound1 Then
            If Value <= Bound2 Then Engine_Collision_Between = 1

        End If

    End If
    
End Function

Public Function Engine_Collision_Line(ByVal L1X1 As Long, _
                                      ByVal L1Y1 As Long, _
                                      ByVal L1X2 As Long, _
                                      ByVal L1Y2 As Long, _
                                      ByVal L2X1 As Long, _
                                      ByVal L2Y1 As Long, _
                                      ByVal L2X2 As Long, _
                                      ByVal L2Y2 As Long) As Byte

    '*****************************************************************
    'Check if two lines intersect (return 1 if true)
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Line
    '*****************************************************************
    Dim m1 As Single

    Dim M2 As Single

    Dim b1 As Single

    Dim b2 As Single

    Dim IX As Single

    'This will fix problems with vertical lines
    If L1X1 = L1X2 Then L1X1 = L1X1 + 1
    If L2X1 = L2X2 Then L2X1 = L2X1 + 1

    'Find the first slope
    m1 = (L1Y2 - L1Y1) / (L1X2 - L1X1)
    b1 = L1Y2 - m1 * L1X2

    'Find the second slope
    M2 = (L2Y2 - L2Y1) / (L2X2 - L2X1)
    b2 = L2Y2 - M2 * L2X2
    
    'Check if the slopes are the same
    If M2 - m1 = 0 Then
    
        If b2 = b1 Then
            'The lines are the same
            Engine_Collision_Line = 1
        Else
            'The lines are parallel (can never intersect)
            Engine_Collision_Line = 0

        End If
        
    Else
        
        'An intersection is a point that lies on both lines. To find this, we set the Y equations equal and solve for X.
        'M1X+B1 = M2X+B2 -> M1X-M2X = -B1+B2 -> X = B1+B2/(M1-M2)
        IX = ((b2 - b1) / (m1 - M2))
        
        'Check for the collision
        If Engine_Collision_Between(IX, L1X1, L1X2) Then
            If Engine_Collision_Between(IX, L2X1, L2X2) Then Engine_Collision_Line = 1

        End If
        
    End If
    
End Function

Public Function Engine_Collision_LineRect(ByVal SX As Long, _
                                          ByVal SY As Long, _
                                          ByVal SW As Long, _
                                          ByVal SH As Long, _
                                          ByVal X1 As Long, _
                                          ByVal Y1 As Long, _
                                          ByVal X2 As Long, _
                                          ByVal Y2 As Long) As Byte
    '*****************************************************************
    'Check if a line intersects with a rectangle (returns 1 if true)
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_LineRect
    '*****************************************************************

    'Top line
    If Engine_Collision_Line(SX, SY, SX + SW, SY, X1, Y1, X2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function

    End If
    
    'Right line
    If Engine_Collision_Line(SX + SW, SY, SX + SW, SY + SH, X1, Y1, X2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function

    End If

    'Bottom line
    If Engine_Collision_Line(SX, SY + SH, SX + SW, SY + SH, X1, Y1, X2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function

    End If

    'Left line
    If Engine_Collision_Line(SX, SY, SX, SY + SW, X1, Y1, X2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function

    End If

End Function

Function Engine_Collision_Rect(ByVal X1 As Integer, _
                               ByVal Y1 As Integer, _
                               ByVal Width1 As Integer, _
                               ByVal Height1 As Integer, _
                               ByVal X2 As Integer, _
                               ByVal Y2 As Integer, _
                               ByVal Width2 As Integer, _
                               ByVal Height2 As Integer) As Boolean
    '*****************************************************************
    'Check for collision between two rectangles
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Rect
    '*****************************************************************

    If X1 + Width1 >= X2 Then
        If X1 <= X2 + Width2 Then
            If Y1 + Height1 >= Y2 Then
                If Y1 <= Y2 + Height2 Then
                    Engine_Collision_Rect = True

                End If

            End If

        End If

    End If

End Function

Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 29/12/10
    'Blisse-AO | DD Clear & BeginScene
    '***************************************************

    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0
    DirectDevice.BeginScene

End Sub

Public Sub Engine_EndScene(ByRef destRect As RECT, Optional ByVal hWndDest As Long = 0)
    '***************************************************
    'Author: Ezequiel Juárez (Standelf)
    'Last Modification: 29/12/10
    'Blisse-AO | DD EndScene & Present
    '***************************************************
    
    If hWndDest = 0 Then
        DirectDevice.EndScene
        DirectDevice.Present destRect, ByVal 0&, ByVal 0&, ByVal 0&
    Else
        DirectDevice.EndScene
        DirectDevice.Present destRect, ByVal 0, hWndDest, ByVal 0

    End If

End Sub

Public Sub Geometry_Create_Box(ByRef Verts() As TLVERTEX, _
                               ByRef dest As RECT, _
                               ByRef src As RECT, _
                               ByRef RGB_List() As Long, _
                               Optional ByRef Textures_Width As Long, _
                               Optional ByRef Textures_Height As Long, _
                               Optional ByVal Angle As Single)
    '**************************************************************
    'Author: Aaron Perkins
    'Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 11/17/2002
    '**************************************************************

    Dim x_center    As Single

    Dim y_center    As Single

    Dim radius      As Single

    Dim x_Cor       As Single

    Dim y_Cor       As Single

    Dim left_point  As Single

    Dim right_point As Single

    Dim temp        As Single
    
    If Angle > 0 Then
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.Bottom - dest.Top) / 2
        
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Bottom - y_center) ^ 2)
        
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = 3.1459 - right_point

    End If
    
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius

    End If

    If Textures_Width And Textures_Height Then
        Verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(0), 0, src.Left / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        Verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(0), 0, 0, 0)

    End If

    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius

    End If
    
    If Textures_Width And Textures_Height Then
        Verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        Verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(1), 0, 0, 1)

    End If

    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius

    End If

    If Textures_Width And Textures_Height Then
        Verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(2), 0, (src.Right + 1) / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        Verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(2), 0, 1, 0)

    End If

    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius

    End If

    If Textures_Width And Textures_Height Then
        Verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        Verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, RGB_List(3), 0, 1, 1)

    End If

End Sub

Public Function Geometry_Create_TLVertex(ByVal X As Single, _
                                         ByVal Y As Single, _
                                         ByVal Z As Single, _
                                         ByVal rhw As Single, _
                                         ByVal Color As Long, _
                                         ByVal Specular As Long, _
                                         tu As Single, _
                                         ByVal tv As Single) As TLVERTEX
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv

End Function

Public Sub Engine_ZoomIn()
    '**************************************************************
    'Author: Standelf
    'Last Modify Date: 29/12/2010
    '**************************************************************

    With MainScreenRect
        .Top = 0
        .Left = 0
        .Bottom = IIf(.Bottom - 1 <= 367, .Bottom, .Bottom - 1)
        .Right = IIf(.Right - 1 <= 491, .Right, .Right - 1)

    End With
    
End Sub

Public Sub Engine_ZoomOut()
    '**************************************************************
    'Author: Standelf
    'Last Modify Date: 29/12/2010
    '**************************************************************

    With MainScreenRect
        .Top = 0
        .Left = 0
        .Bottom = IIf(.Bottom + 1 >= 459, .Bottom, .Bottom + 1)
        .Right = IIf(.Right + 1 >= 583, .Right, .Right + 1)

    End With
    
End Sub

Public Sub Engine_ZoomNormal()
    '**************************************************************
    'Author: Standelf
    'Last Modify Date: 29/12/2010
    '**************************************************************

    With MainScreenRect
        .Top = 0
        .Left = 0
        .Bottom = ScreenHeight
        .Right = ScreenWidth

    End With
    
End Sub

Public Function ZoomOffset(ByVal offset As Byte) As Single
    '**************************************************************
    'Author: Standelf
    'Last Modify Date: 30/01/2011
    '**************************************************************

    ZoomOffset = IIf((offset = 1), (ScreenHeight - MainScreenRect.Bottom) / 2, (ScreenWidth - MainScreenRect.Right) / 2)
    
End Function

Public Sub Engine_Set_BaseSpeed(ByVal BaseSpeed As Single)
    '**************************************************************
    'Author: Standelf
    'Last Modify Date: 29/12/2010
    '**************************************************************

    Engine_BaseSpeed = BaseSpeed
    
End Sub

Public Function Engine_Get_BaseSpeed() As Single
    '**************************************************************
    'Author: Standelf
    'Last Modify Date: 29/12/2010
    '**************************************************************

    Engine_Get_BaseSpeed = Engine_BaseSpeed
    
End Function

Public Sub Engine_Set_TileBuffer(ByVal setTileBufferSize As Single)
    '**************************************************************
    'Author: Standelf
    'Last Modify Date: 30/12/2010
    '**************************************************************

    TileBufferSize = setTileBufferSize
    
End Sub

Public Function Engine_Get_TileBuffer() As Single
    '**************************************************************
    'Author: Standelf
    'Last Modify Date: 30/12/2010
    '**************************************************************

    Engine_Get_TileBuffer = TileBufferSize
    
End Function

Function Engine_Distance(ByVal X1 As Integer, _
                         ByVal Y1 As Integer, _
                         ByVal X2 As Integer, _
                         ByVal Y2 As Integer) As Long
    '***************************************************
    'Author: Standelf
    'Last Modification: -
    '***************************************************

    Engine_Distance = Abs(X1 - X2) + Abs(Y1 - Y2)
    
End Function

Public Sub Engine_Update_FPS()
    '***************************************************
    'Author: Standelf
    'Last Modification: 10/01/2011
    'Limit FPS & Calculate later
    '***************************************************

    If ClientSetup.LimiteFPS And Not ClientSetup.vSync Then

        While (GetTickCount - FPSLastCheck) \ 10 < FramesPerSecCounter

            Sleep 5
        Wend

    End If
        
    If FPSLastCheck + 1000 < GetTickCount Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        FPSLastCheck = GetTickCount
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1

    End If

    If FPSFLAG Then DrawText 685, 2, FPS, -1

End Sub

