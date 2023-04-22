Attribute VB_Name = "modCarga"
Option Explicit

Public Type tSetupMods

    ' VIDEO
    bDinamic    As Boolean
    byMemory    As Integer
    bFullScreen As Boolean
    ProyectileEngine As Boolean
    PartyMembers As Boolean
    TonalidadPJ As Boolean
    UsarSombras As Boolean
    ParticleEngine As Boolean
    vSync As Boolean
    OverrideVertexProcess As Byte
    LimiteFPS As Boolean

    
    ' AUDIO
    bMusic      As Boolean
    bSound      As Boolean
    MusicVolume As Byte
    SoundVolume As Byte
    bSoundEffects As Boolean
    
    ' GUILD
    bGuildNews  As Boolean ' 11/19/09
    
    ' OTROS
    MostrarTips As Boolean

End Type

Public ClientSetup   As tSetupMods

Public Type tCabecera 'Cabecera de los con

    Desc As String * 255
    CRC As Long
    MagicWord As Long

End Type

Public MiCabecera As tCabecera

Private Lector As ClsIniReader
Private Const CLIENT_FILE As String = "Config.ini"

Public Sub IniciarCabecera()

    With MiCabecera
    
        .Desc = "Nexus AO by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
        .CRC = Rnd * 100
        .MagicWord = Rnd * 10
    
    End With

End Sub

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Public Sub LoadGrhData()

On Error GoTo ErrorHandler:

    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    Open DirIni & "Graficos.ind" For Binary Access Read As handle
    
        Get handle, , fileVersion
        
        Get handle, , grhCount
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Not EOF(handle)
            Get handle, , Grh
            
            With GrhData(Grh)
            
                '.active = True
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    
                    Get handle, , .Speed
                    If .Speed <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                    
                Else
                    
                    Get handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , GrhData(Grh).SX
                    If .SX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .SY
                    If .SY < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend
    
    Close handle
    
Exit Sub

ErrorHandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
        
    End If

End Sub

Sub CargarCabezas()

    Dim N            As Integer

    Dim i            As Long

    Dim Numheads     As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open DirIni & "Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If

    Next i
    
    Close #N

End Sub

Sub CargarCascos()

    Dim N            As Integer

    Dim i            As Long

    Dim NumCascos    As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open DirIni & "Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)

        End If

    Next i
    
    Close #N

End Sub

Sub CargarCuerpos()

    Dim N            As Integer

    Dim i            As Long

    Dim NumCuerpos   As Integer

    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open DirIni & "Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY

        End If

    Next i
    
    Close #N

End Sub

Sub CargarFxs()

    Dim N      As Integer

    Dim i      As Long

    Dim NumFxs As Integer
    
    N = FreeFile()
    Open DirIni & "Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(0 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
        'MsgBox FxData(i).Animacion & FxData(i).OffsetX
    Next i
    
    Close #N

End Sub

Sub CargarTips()

    Dim N       As Integer

    Dim i       As Long

    Dim NumTips As Integer
    
    N = FreeFile
    Open DirIni & "Tips.ayu" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For i = 1 To NumTips
        Get #N, , Tips(i)
    Next i
    
    Close #N

End Sub

Sub CargarArrayLluvia()

    Dim N  As Integer

    Dim i  As Long

    Dim Nu As Integer
    
    N = FreeFile()
    Open DirIni & "fk.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For i = 1 To Nu
        Get #N, , bLluvia(i)
    Next i
    
    Close #N

End Sub

Sub CargarAnimArmas()

    On Error Resume Next

    Dim LoopC As Long

    Dim arch  As String
    
    arch = DirIni & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For LoopC = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(LoopC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & LoopC, "Dir1")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & LoopC, "Dir2")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & LoopC, "Dir3")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & LoopC, "Dir4")), 0
    Next LoopC

End Sub

Sub CargarAnimEscudos()

    On Error Resume Next

    Dim LoopC As Long

    Dim arch  As String
    
    arch = DirIni & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For LoopC = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(LoopC).ShieldWalk(1), Val(GetVar(arch, "ESC" & LoopC, "Dir1")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(2), Val(GetVar(arch, "ESC" & LoopC, "Dir2")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(3), Val(GetVar(arch, "ESC" & LoopC, "Dir3")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(4), Val(GetVar(arch, "ESC" & LoopC, "Dir4")), 0
    Next LoopC

End Sub

Public Sub CargarColores()

    On Error Resume Next

    Dim archivoC As String

    archivoC = DirIni & "colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub

    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i) = D3DColorXRGB(GetVar(archivoC, CStr(i), "R"), GetVar(archivoC, CStr(i), "G"), GetVar(archivoC, CStr(i), "B"))
    Next i
    
    '   Crimi
    ColoresPJ(50) = D3DColorXRGB(GetVar(archivoC, "CR", "R"), GetVar(archivoC, "CR", "G"), GetVar(archivoC, "CR", "B"))

    '   Ciuda
    ColoresPJ(49) = D3DColorXRGB(GetVar(archivoC, "CI", "R"), GetVar(archivoC, "CI", "G"), GetVar(archivoC, "CI", "B"))
    
    '   Atacable
    ColoresPJ(50) = D3DColorXRGB(GetVar(archivoC, "AT", "R"), GetVar(archivoC, "AT", "G"), GetVar(archivoC, "AT", "B"))

End Sub

Public Sub CargarServidores()

    '********************************
    'Author: Unknown
    'Last Modification: 07/26/07
    'Last Modified by: Rapsodius
    'Added Instruction "CloseClient" before End so the mutex is cleared
    '********************************
    On Error GoTo errorH

    Dim F As String

    Dim c As Integer

    Dim i As Long
    
    F = DirIni & "sinfo.dat"
    c = Val(GetVar(F, "INIT", "Cant"))
    
    ReDim ServersLst(1 To c) As tServerInfo

    For i = 1 To c
        ServersLst(i).Desc = GetVar(F, "S" & i, "Desc")
        ServersLst(i).Ip = Trim$(GetVar(F, "S" & i, "Ip"))
        ServersLst(i).PassRecPort = CInt(GetVar(F, "S" & i, "P2"))
        ServersLst(i).Puerto = CInt(GetVar(F, "S" & i, "PJ"))
    Next i

    CurServer = 1
    Exit Sub

errorH:
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Nexus AO")
    
    Call CloseClient

End Sub

Public Sub SwitchMap(ByVal Map As Integer)

    '**************************************************************
    'Formato de mapas optimizado para reducir el espacio que ocupan.
    'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
    '**************************************************************
    Dim Y         As Long

    Dim X         As Long

    Dim tempint   As Integer

    Dim ByFlags   As Byte

    Dim handle    As Integer

    Dim CharIndex As Integer

    Dim Obj       As Integer
    
    handle = FreeFile()
    
    Call Char_CleanAll
    
    Open DirMapas & "Mapa" & Map & ".map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            Get handle, , ByFlags
            
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0

            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0

            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0

            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0

            End If
            
            'Erase NPCs
            CharIndex = Char_MapPosExits(X, Y)
 
            If (CharIndex > 0) Then
                Call Char_Erase(CharIndex)

            End If

            'Erase OBJs
            Obj = Map_PosExitsObject(X, Y)

            If (Obj > 0) Then
                Call Map_DestroyObject(X, Y)

            End If
            
            'Erase Lights
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual) 'Standelf, Light & Meteo Engine
        Next X
    Next Y
    
    Close handle
    
    Call LightRemoveAll
    
    '   Erase particle effects
    ReDim Effect(1 To NumEffects)
    
    MapInfo.name = ""
    MapInfo.Music = ""
    
    CurMap = Map
    
    Init_Ambient Map
    
    'If UserMap = 120 Then Effect_Waterfall_Begin Engine_TPtoSPX(8), Engine_TPtoSPY(3), 1, 800
End Sub

Public Sub LeerConfiguracion()
    On Local Error GoTo fileErr:
    
        Call IniciarCabecera
    
        Set Lector = New ClsIniReader
        Call Lector.Initialize(DirIni & CLIENT_FILE)
        
        With ClientSetup
        
            ' VIDEO
            .byMemory = Lector.GetValue("VIDEO", "DynamicMemory")
            .bFullScreen = CBool(Lector.GetValue("VIDEO", "FullScreen"))
            .ProyectileEngine = CBool(Lector.GetValue("VIDEO", "ProjectileEngine"))
            .PartyMembers = CBool(Lector.GetValue("VIDEO", "PartyMembers"))
            .TonalidadPJ = CBool(Lector.GetValue("VIDEO", "TonalidadPJ"))
            .UsarSombras = CBool(Lector.GetValue("VIDEO", "Sombras"))
            .ParticleEngine = CBool(Lector.GetValue("VIDEO", "ParticleEngine"))
            .LimiteFPS = CBool(Lector.GetValue("VIDEO", "LimitarFPS"))
            .OverrideVertexProcess = CByte(Lector.GetValue("VIDEO", "VertexProcessingOverride"))
        
            ' AUDIO
            .bMusic = CBool(Lector.GetValue("AUDIO", "Music"))
            .bSound = CBool(Lector.GetValue("AUDIO", "Sound"))
            .bSoundEffects = CBool(Lector.GetValue("AUDIO", "SoundEffects"))
            .MusicVolume = CByte(Lector.GetValue("AUDIO", "MusicVolume"))
            .SoundVolume = CByte(Lector.GetValue("AUDIO", "SoundVolume"))

            ' GUILD
            .bGuildNews = CBool(Lector.GetValue("GUILD", "News"))
            
            ' OTROS
            .MostrarTips = CBool(Lector.GetValue("OTHER", "MOSTRAR_TIPS"))

        End With
        
        Exit Sub
    
fileErr:

    If Err.number <> 0 Then
      MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
      End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If
    
End Sub
