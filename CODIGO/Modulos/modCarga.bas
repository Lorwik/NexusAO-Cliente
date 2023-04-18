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
    Aceleracion As Byte
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
            .Aceleracion = CByte(Lector.GetValue("VIDEO", "Aceleracion"))
        
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
