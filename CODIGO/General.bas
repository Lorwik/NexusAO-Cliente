Attribute VB_Name = "Mod_General"
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

Private m_Jpeg      As clsJpeg
Private m_FileName  As String

Public iplst        As String

Public bFogata      As Boolean

Public bLluvia()    As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameTimer As Long

Public Function DirGraficos() As String
    DirGraficos = App.path & "\Recursos\Graficos\"

End Function

Public Function DirInterfaces() As String
    DirInterfaces = App.path & "\Recursos\Interfaces\"
    
End Function

Public Function DirIni() As String
    DirIni = App.path & "\Recursos\Init\"
    
End Function

Public Function DirSound() As String
    DirSound = App.path & "\Recursos\Wav\"

End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\Recursos\Midi\"

End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\Recursos\Mapas\"

End Function

Public Function DirExtras() As String
    DirExtras = App.path & "\Recursos\EXTRAS\"

End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

Public Function GetRawName(ByRef sName As String) As String
    '***************************************************
    'Author: ZaMa
    'Last Modify Date: 13/01/2010
    'Last Modified By: -
    'Returns the char name without the clan name (if it has it).
    '***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim(Left(sName, Pos - 1))
    Else
        GetRawName = sName

    End If

End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                     ByVal Text As String, _
                     Optional ByVal Red As Integer = -1, _
                     Optional ByVal Green As Integer, _
                     Optional ByVal Blue As Integer, _
                     Optional ByVal bold As Boolean = False, _
                     Optional ByVal italic As Boolean = False, _
                     Optional ByVal bCrLf As Boolean = True)

    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    '******************************************r
    With RichTextBox

        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        RichTextBox.Refresh

    End With

End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()

    '*****************************************************************
    'Goes through the charlist and replots all the characters on the map
    'Used to make sure everyone is visible
    '*****************************************************************
    Dim LoopC As Long
    
    For LoopC = 1 To LastChar

        If charlist(LoopC).active = 1 Then
            MapData(charlist(LoopC).Pos.X, charlist(LoopC).Pos.Y).CharIndex = LoopC

        End If

    Next LoopC

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean

    Dim car As Byte

    Dim i   As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function

        End If

    Next i
    
    AsciiValidos = True

End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean

    'Validamos los datos del user
    Dim LoopC     As Long

    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function

    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function

    End If
    
    For LoopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, LoopC, 1))

        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function

        End If

    Next LoopC
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function

    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function

    End If
    
    For LoopC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, LoopC, 1))

        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function

        End If

    Next LoopC
    
    CheckUserData = True

End Function

Sub UnloadAllForms()

    On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next

End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean

    '*****************************************************************
    'Only allow characters that are Win 95 filename compatible
    '*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function

    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function

    End If
    
    If KeyAscii > 126 Then
        Exit Function

    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function

    End If
    
    'else everything is cool
    LegalCharacter = True

End Function

Sub SetConnected()
    '*****************************************************************
    'Sets the client to "Connect" mode
    '*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    
    frmMain.lblName.Caption = UserName
    'Load main form
    frmMain.Visible = True

End Sub

Sub CargarTip()

    Dim N As Integer

    N = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(N)

End Sub

Sub MoveTo(ByVal Direccion As E_Heading)

    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/28/2008
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
    ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
    ' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
    '***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion

        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)

        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)

        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)

        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)

    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)

        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion

        End If

    Else

        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)

        End If

    End If
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)

End Sub

Sub RandomMove()
    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modify Date: 06/03/2006
    ' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
    '***************************************************
    Call Map_MoveTo(RandomNumber(NORTH, WEST))

End Sub

Private Sub CheckKeys()

    '*****************************************************************
    'Checks keys and respond
    '*****************************************************************
    Static LastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub

    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If GetTickCount - LastMovement > 56 Then
        LastMovement = GetTickCount
    Else
        Exit Sub

    End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then

            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call Map_MoveTo(NORTH)
                Call Char_UserPos
                Exit Sub

            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call Map_MoveTo(EAST)
                'frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
                Call Char_UserPos
                Exit Sub

            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call Map_MoveTo(SOUTH)
                Call Char_UserPos
                Exit Sub

            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call Map_MoveTo(WEST)
                Call Char_UserPos
                Exit Sub

            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else

            Dim kp As Boolean

            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)

            End If
            
            Call Char_UserPos

        End If

    End If

End Sub

Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String

    '*****************************************************************
    'Gets a field from a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/15/2004
    '*****************************************************************
    Dim i          As Long

    Dim lastPos    As Long

    Dim CurrentPos As Long

    Dim delimiter  As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)

    End If

End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long

    '*****************************************************************
    'Gets the number of fields in a delimited string
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 07/29/2007
    '*****************************************************************
    Dim Count     As Long

    Dim curPos    As Long

    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count

End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")

End Function

Public Function IsIp(ByVal Ip As String) As Boolean

    Dim i As Long
    
    For i = 1 To UBound(ServersLst)

        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function

        End If

    Next i

End Function

Public Sub InitServersList()

    On Error Resume Next

    Dim NumServers As Integer

    Dim i          As Integer

    Dim Cont       As Integer
    
    i = 1
    
    Do While (ReadField(i, RawServersList, Asc(";")) <> "")
        i = i + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For i = 1 To Cont

        Dim cur$

        cur$ = ReadField(i, RawServersList, Asc(";"))
        ServersLst(i).Ip = ReadField(1, cur$, Asc(":"))
        ServersLst(i).Puerto = ReadField(2, cur$, Asc(":"))
        ServersLst(i).Desc = ReadField(4, cur$, Asc(":"))
        ServersLst(i).PassRecPort = ReadField(3, cur$, Asc(":"))
    Next i
    
    CurServer = 1

End Sub

Public Function CurServerPasRecPort() As Integer

    If CurServer <> 0 Then
        CurServerPasRecPort = 7667
    Else
        CurServerPasRecPort = CInt(frmConnect.PortTxt)

    End If

End Function

Public Function CurServerIp() As String

    If CurServer <> 0 Then
        CurServerIp = ServersLst(CurServer).Ip
    Else
        CurServerIp = frmConnect.IPTxt

    End If

End Function

Public Function CurServerPort() As Integer

    If CurServer <> 0 Then
        CurServerPort = ServersLst(CurServer).Puerto
    Else
        CurServerPort = Val(frmConnect.PortTxt)

    End If

End Function

Sub Main()
    
    Call modCarga.LeerConfiguracion
    
    Call modCompression.GenerateContra(vbNullString, 0) ' 0 = Graficos.AO
    
    CargarHechizos
    
    If ClientSetup.bDinamic Then
        Set SurfaceDB = New clsSurfaceManDyn
    Else
        Set SurfaceDB = New clsSurfaceManStatic

    End If
 
    ' Map Sounds
    Set Sonidos = New clsSoundMapas
    Call Sonidos.LoadSoundMapInfo
       
    #If Testeo = 0 Then

        If FindPreviousInstance Then
            Call MsgBox("Nexus AO ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
            End

        End If

    #End If

    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    ChDrive App.path
    ChDir App.path

    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
    
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    Call Resolution.SetResolution(1024, 768)
    
    ' Load constants, classes, flags, graphics..
    LoadInitialConfig

    #If UsarWrench = 1 Then
        frmMain.Socket1.Startup
    #End If

    frmConnect.Visible = True
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    pausa = False
    
    ' Intervals
    LoadTimerIntervals
        
    'Set the dialog's font
    Dialogos.Font = frmMain.Font
    
    lFrameTimer = GetTickCount
        
    Do While prgRun

        'Sólo dibujamos si la ventana no está minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            Call RenderSounds
            
            Call CheckKeys

        End If

        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then _
            lFrameTimer = GetTickCount
        
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
    Loop
    
    Call CloseClient

End Sub

Private Sub LoadInitialConfig()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/03/2011
    '15/03/2011: ZaMa - Initialize classes lazy way.
    '***************************************************

    Dim i As Long

    frmCargando.Show
    frmCargando.Refresh

    frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    
    '###########
    ' SERVIDORES
    'TODO : esto de ServerRecibidos no se podría sacar???
    Call AddtoRichTextBox(frmCargando.status, "Buscando servidores... ", 255, 255, 255, True, False, True)
    Call CargarServidores
    ServersRecibidos = True
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    '###########
    ' CONSTANTES
    Call AddtoRichTextBox(frmCargando.status, "Iniciando constantes... ", 255, 255, 255, True, False, True)
    Call InicializarNombres
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
    
    UserMap = 1
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(DirExtras & "Hand.ico", vbArchive) Then Set picMouseIcon = LoadPicture(DirExtras & "Hand.ico")
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    '#######
    ' CLASES
    Call AddtoRichTextBox(frmCargando.status, "Instanciando clases... ", 255, 255, 255, True, False, True)
    Set Dialogos = New clsDialogs
    Set Audio = New clsAudio
    Set Inventario = New clsGrapchicalInventory
    Set CustomKeys = New clsCustomKeys
    Set CustomMessages = New clsCustomMessages
    Set incomingData = New clsByteQueue
    Set outgoingData = New clsByteQueue
    Set MainTimer = New clsTimer
    Set clsForos = New clsForum
    
    '##############
    ' MOTOR GRÁFICO
    Call AddtoRichTextBox(frmCargando.status, "Iniciando motor gráfico... ", 255, 255, 255, True, False, True)
    
    ' Iniciamos el Engine de DirectX 8
    Call Engine_DirectX8_Init
          
    ' Tile Engine
    If Not InitTileEngine(frmMain.hwnd, 32, 32, 8, 8) Then Call CloseClient
        
    Call Engine_DirectX8_Aditional_Init
    
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    '###################
    ' ANIMACIONES EXTRAS
    Call AddtoRichTextBox(frmCargando.status, "Creando animaciones extra... ", 255, 255, 255, True, False, True)
    Call CargarTips
    ' Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    '#############
    ' DIRECT SOUND
    Call AddtoRichTextBox(frmCargando.status, "Iniciando DirectSound... ", 255, 255, 255, True, False, True)
    
    ' Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hwnd, App.path & "\Wav\", App.path & "\Midi\")
    
    'Enable / Disable audio
    Audio.MusicActivated = ClientSetup.bMusic
    Audio.SoundActivated = ClientSetup.bSound
    Audio.SoundEffectsActivated = ClientSetup.bSoundEffects
    Audio.MusicVolume = ClientSetup.MusicVolume
    Audio.SoundVolume = ClientSetup.SoundVolume
    
    ' Inicializamos el inventario gráfico
    Call Inventario.Initialize(DirectD3D8, frmMain.PicInv, MAX_INVENTORY_SLOTS)
    
    'Call Audio.MusicMP3Play(App.path & "\MP3\" & MP3_Inicio & ".mp3")
    Call AddtoRichTextBox(frmCargando.status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.status, "                    ¡Bienvenido a Nexus AO!", 255, 255, 255, True, False, True)

    'Give the user enough time to read the welcome text
    Call Sleep(500)
    
    Unload frmCargando
    
End Sub

Private Sub LoadTimerIntervals()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/03/2011
    'Set the intervals of timers
    '***************************************************
    
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
    'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)

End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal value As String)
    '*****************************************************************
    'Writes a var to a text file
    '*****************************************************************
    writeprivateprofilestring Main, Var, value, File

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

    '*****************************************************************
    'Gets a Var from a text file
    '*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

    On Error GoTo errHnd

    Dim lPos As Long

    Dim Lx   As Long

    Dim iAsc As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")

    If (lPos <> 0) Then

        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For Lx = 0 To Len(sString) - 1

            If Not (Lx = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (Lx + 1), 1))

                If Not CMSValidateChar_(iAsc) Then Exit Function

            End If

        Next Lx
        
        'Finale
        CheckMailString = True

    End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()

    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus

    End If

End Sub

''
' Checks the command line parameters, if you are running Ao with /nores command and checks the AoUpdate parameters
'
'

Public Sub LeerLineaComandos()

    '*************************************************
    'Author: Unknown
    'Last modified: 25/11/2008 (BrianPr)
    '
    '*************************************************
    Dim t()      As String

    Dim i        As Long
    
    Dim UpToDate As Boolean

    Dim Patch    As String
    
    'Parseo los comandos
    t = Split(Command, " ")

    For i = LBound(t) To UBound(t)

        Select Case UCase$(t(i))

            Case "/NORES" 'no cambiar la resolucion
                NoRes = True

            Case "/UPTODATE"
                UpToDate = True

        End Select

    Next i
    
    #If Testeo = 0 Then
        Call AoUpdate(UpToDate, NoRes)
    #End If

End Sub

''
' Runs AoUpdate if we haven't updated yet, patches aoupdate and runs Client normally if we are updated.
'
' @param UpToDate Specifies if we have checked for updates or not
' @param NoREs Specifies if we have to set nores arg when running the client once again (if the AoUpdate is executed).

Private Sub AoUpdate(ByVal UpToDate As Boolean, ByVal NoRes As Boolean)

    '*************************************************
    'Author: BrianPr
    'Created: 25/11/2008
    'Last modified: 25/11/2008
    '
    '*************************************************
    On Error GoTo error

    Dim extraArgs As String

    If Not UpToDate Then

        'No recibe update, ejecutar AU
        'Ejecuto el AoUpdate, sino me voy
        If Dir(App.path & "\AoUpdate.exe", vbArchive) = vbNullString Then
            MsgBox "No se encuentra el archivo de actualización AoUpdate.exe por favor descarguelo y vuelva a intentar", vbCritical
            End
        Else
            FileCopy App.path & "\AoUpdate.exe", App.path & "\AoUpdateTMP.exe"
            
            If NoRes Then
                extraArgs = " /nores"

            End If
            
            Call ShellExecute(0, "Open", App.path & "\AoUpdateTMP.exe", App.EXEName & ".exe" & extraArgs, App.path, SW_SHOWNORMAL)
            End

        End If

    Else

        If FileExist(App.path & "\AoUpdateTMP.exe", vbArchive) Then Kill App.path & "\AoUpdateTMP.exe"

    End If

    Exit Sub

error:

    If Err.number = 75 Then 'Si el archivo AoUpdateTMP.exe está en uso, entonces esperamos 5 ms y volvemos a intentarlo hasta que nos deje.
        Sleep 5
        Resume
    Else
        MsgBox Err.Description & vbCrLf, vbInformation, "[ " & Err.number & " ]" & " Error "
        End

    End If

End Sub

Private Sub InicializarNombres()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
    '**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"

End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 11/27/2005
    'Removes all text from the console and dialogs
    '**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    
    Call Dialogos.RemoveAllDialogs

End Sub

Public Sub CloseClient()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    EngineRun = False
    
    Call Resolution.ResetResolution
    
    'Stop tile engine
    Call Engine_DirectX8_End
    
    'TODO// Guardar configuracion
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    
    'Si se cambio la resolucion, la reseteamos.
    If ResolucionCambiada Then Resolution.ResetResolution
    
    End

End Sub

Public Function esGM(CharIndex As Integer) As Boolean
    esGM = False

    If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer

    Dim buf As Integer

    buf = InStr(Nick, "<")

    If buf > 0 Then
        getTagPosition = buf
        Exit Function

    End If

    buf = InStr(Nick, "[")

    If buf > 0 Then
        getTagPosition = buf
        Exit Function

    End If

    getTagPosition = Len(Nick) + 2

End Function

Public Sub checkText(ByVal Text As String)

    Dim Nivel As Integer

    If Left(Text, Len(MENSAJE_FRAGSHOOTER_HAS_MATADO)) = MENSAJE_FRAGSHOOTER_HAS_MATADO Then
        EsperandoLevel = True
        Exit Sub

    End If

    EsperandoLevel = False

End Sub

Public Function getCharIndexByName(ByVal name As String) As Integer

    Dim i As Long

    For i = 1 To LastChar

        If charlist(i).Nombre = name Then
            getCharIndexByName = i
            Exit Function

        End If

    Next i

End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: 22/02/2010
    'Returns true if the post is sticky.
    '***************************************************
    Select Case ForumType

        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte

    '***************************************************
    'Author: ZaMa
    'Last Modification: 01/03/2010
    'Returns the forum alignment.
    '***************************************************
    Select Case yForumType

        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
End Function

Public Sub ResetAllInfo()
    

    Connected = False
    
    'Unload all forms except frmMain, frmConnect and frmCrearPersonaje
    Dim frm As Form

    For Each frm In Forms

        If frm.name <> frmMain.name And frm.name <> frmConnect.name And frm.name <> frmCrearPersonaje.name Then _
            Unload frm

    Next
    
    On Local Error GoTo 0
    
    ' Return to connection screen
    frmConnect.MousePointer = vbNormal

    If Not frmCrearPersonaje.Visible Then frmConnect.Visible = True
    frmMain.Visible = False
    
    'Stop audio
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone
    
    ' Reset flags
    pausa = False
    UserMeditar = False
    UserEstupido = False
    UserCiego = False
    UserDescansar = False
    UserParalizado = False
    Traveling = False
    UserNavegando = False
    bFogata = False
    bRain = False
    bFogata = False
    Comerciando = False
    bShowTutorial = False
    
    MirandoAsignarSkills = False
    MirandoCarpinteria = False
    MirandoEstadisticas = False
    MirandoForo = False
    MirandoHerreria = False
    MirandoParty = False
    
    'Delete all kind of dialogs
    Call CleanDialogs

    'Reset some char variables...
    Dim i As Long

    For i = 1 To LastChar
        charlist(i).invisible = False
    Next i

    ' Reset stats
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    SkillPoints = 0
    Alocados = 0
    
    ' Reset skills
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    ' Reset attributes
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    ' Clear inventory slots
    Inventario.ClearAllSlots

    ' Connection screen midi
    Call Audio.PlayMIDI("2.mid")

End Sub

Public Function DevolverNombreHechizo(ByVal Index As Byte) As String

    Dim i As Long
 
    For i = 1 To NumHechizos

        If i = Index Then
            DevolverNombreHechizo = Hechizos(i).Nombre
            Exit Function

        End If

    Next i

End Function

Public Function DevolverIndexHechizo(ByVal Nombre As String) As Byte

    Dim i As Long
 
    For i = 1 To NumHechizos

        If Hechizos(i).Nombre = Nombre Then
            DevolverIndexHechizo = i
            Exit Function

        End If

    Next i

End Function

Public Sub CargarHechizos()

    '********************************
    'Author: Shak
    'Last Modification:
    'Cargamos los hechizos del juego. [Solo datos necesarios]
    '********************************
    On Error GoTo errorH

    Dim PathName As String

    Dim j        As Long
 
    PathName = DirIni & "Hechizos.dat"
    NumHechizos = Val(GetVar(PathName, "INIT", "NumHechizos"))
 
    ReDim Hechizos(1 To NumHechizos) As tHechizos

    For j = 1 To NumHechizos

        With Hechizos(j)
            .Desc = GetVar(PathName, "HECHIZO" & j, "Desc")
            .PalabrasMagicas = GetVar(PathName, "HECHIZO" & j, "PalabrasMagicas")
            .Nombre = GetVar(PathName, "HECHIZO" & j, "Nombre")
            .SkillRequerido = GetVar(PathName, "HECHIZO" & j, "MinSkill")
         
            If j <> 38 And j <> 39 Then
                .EnergiaRequerida = GetVar(PathName, "HECHIZO" & j, "StaRequerido")
                 
                .HechiceroMsg = GetVar(PathName, "HECHIZO" & j, "HechizeroMsg")
                .ManaRequerida = GetVar(PathName, "HECHIZO" & j, "ManaRequerido")
             
                .PropioMsg = GetVar(PathName, "HECHIZO" & j, "PropioMsg")
             
                .TargetMsg = GetVar(PathName, "HECHIZO" & j, "TargetMsg")

            End If

        End With

    Next j
 
    Exit Sub
 
errorH:
    Call MsgBox("Error critico", vbCritical + vbOKOnly, "Nexus AO")

End Sub

Public Sub Client_Screenshot(ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long)
'*******************************
'Autor: ???
'Fecha: ???
'*******************************

On Error GoTo ErrorHandler

    Dim i As Long
    Dim Index As Long
    i = 1
    
    Set m_Jpeg = New clsJpeg
    
    '80 Quality
    m_Jpeg.Quality = 100
    
    'Sample the cImage by hDC
    m_Jpeg.SampleHDC hDC, Width, Height
    
    m_FileName = App.path & "\Fotos\NexusAO_Foto"
    
    If Dir(App.path & "\Fotos", vbDirectory) = vbNullString Then
        MkDir (App.path & "\Fotos")
    End If
    
    Do While Dir(m_FileName & Trim(str(i)) & ".jpg") <> vbNullString
        i = i + 1
        DoEvents
    Loop
    
    Index = i
    
    m_Jpeg.Comment = "Character: " & UserName & " - " & Format(Date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm AM/PM")
    
    'Save the JPG file
    m_Jpeg.SaveFile m_FileName & Trim(str(Index)) & ".jpg"
    
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Captura realizada con exito! Se guardo en " & m_FileName & Trim(str(Index)) & ".jpg", 204, 193, 155, 0, 1)
    
    Set m_Jpeg = Nothing
    
    Exit Sub

ErrorHandler:
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Error en la captura!", 204, 193, 155, 0, 1)

End Sub
