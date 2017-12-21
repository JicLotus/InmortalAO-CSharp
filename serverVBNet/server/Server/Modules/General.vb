Option Explicit On


Module General

    Public LeerNPCs As New clsIniReader

    Sub DarCuerpoDesnudo(ByVal UserIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/14/07
        'Da cuerpo desnudo a un usuario
        '***************************************************
        Dim CuerpoDesnudo As Integer
        Select Case UserList(UserIndex).Genero
            Case eGenero.Hombre
                Select Case UserList(UserIndex).Raza
                    Case eRaza.Humano
                        CuerpoDesnudo = 21
                    Case eRaza.Drow
                        CuerpoDesnudo = 32
                    Case eRaza.Elfo
                        CuerpoDesnudo = 21
                    Case eRaza.Gnomo
                        CuerpoDesnudo = 53
                    Case eRaza.Enano
                        CuerpoDesnudo = 53
                    Case eRaza.Orco
                        CuerpoDesnudo = 248
                End Select
            Case eGenero.Mujer
                Select Case UserList(UserIndex).Raza
                    Case eRaza.Humano
                        CuerpoDesnudo = 39
                    Case eRaza.Drow
                        CuerpoDesnudo = 40
                    Case eRaza.Elfo
                        CuerpoDesnudo = 39
                    Case eRaza.Gnomo
                        CuerpoDesnudo = 60
                    Case eRaza.Enano
                        CuerpoDesnudo = 60
                    Case eRaza.Orco
                        CuerpoDesnudo = 249
                End Select
        End Select

        UserList(UserIndex).cuerpo.body = CuerpoDesnudo

        UserList(UserIndex).flags.Desnudo = 1

    End Sub


    Sub Bloquear(ByVal ToMap As Boolean, ByVal sndIndex As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal b As Boolean)
        'b ahora es boolean,
        'b=true bloquea el tile en (x,y)
        'b=false desbloquea el tile en (x,y)
        'ToMap = true -> Envia los datos a todo el mapa
        'ToMap = false -> Envia los datos al user
        'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
        'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s

        If ToMap Then
            Call SendData(SendTarget.ToMap, sndIndex, PrepareMessageBlockPosition(x, Y, b))
        Else
            Call WriteBlockPosition(sndIndex, x, Y, b)
        End If

    End Sub


    Function HayAgua(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean


        If MapData(map, x, Y).Graphic Is Nothing Then
            HayAgua = False
            Exit Function
        End If

        If map > 0 And map < NumMaps + 1 And x > 0 And x < 101 And Y > 0 And Y < 101 Then
            If ((MapData(map, x, Y).Graphic(1) >= 1505 And MapData(map, x, Y).Graphic(1) <= 1520) Or
            (MapData(map, x, Y).Graphic(1) >= 5665 And MapData(map, x, Y).Graphic(1) <= 5680) Or
            (MapData(map, x, Y).Graphic(1) >= 13547 And MapData(map, x, Y).Graphic(1) <= 13562)) And
               MapData(map, x, Y).Graphic(2) = 0 Then
                HayAgua = True
            Else
                HayAgua = False
            End If
        Else
            HayAgua = False
        End If

    End Function

    Private Function HayLava(ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean
        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/12/07
        '***************************************************
        If map > 0 And map < NumMaps + 1 And x > 0 And x < 101 And Y > 0 And Y < 101 Then
            If MapData(map, x, Y).Graphic(1) >= 5837 And MapData(map, x, Y).Graphic(1) <= 5852 Then
                HayLava = True
            Else
                HayLava = False
            End If
        Else
            HayLava = False
        End If

    End Function

    Sub LimpiarMundo()
        '***************************************************
        'Author: Unknow
        'Last Modification: 04/15/2008
        '01/14/2008: Marcos Martinez (ByVal) - La funcion FOR estaba mal. En ves de i habia un 1.
        '04/15/2008: (NicoNZ) - La funcion FOR estaba mal, de la forma que se hacia tiraba error.
        '***************************************************
        On Error GoTo Errhandler

        Dim i As Integer
        Dim d As cGarbage
        d = New cGarbage

        For i = TrashCollector.Count To 1 Step -1
            d = TrashCollector(i)
            Call EraseObj(1, d.map, d.x, d.Y)
            Call TrashCollector.Remove(i)
            d = Nothing
        Next i

        Call SecurityIp.IpSecurityMantenimientoLista()

        Exit Sub

Errhandler:
        Call LogError("Error producido en el sub LimpiarMundo: " & Err.Description)
    End Sub

    Sub EnviarSpawnList(ByVal UserIndex As Integer)
        Dim k As Long
        Dim npcNames() As String

        ReDim npcNames(UBound(SpawnList))

        For k = 1 To UBound(SpawnList)
            npcNames(k) = SpawnList(k).NpcName
        Next k

        Call WriteSpawnList(UserIndex, npcNames)

    End Sub




    Sub Main()

        On Error Resume Next
        Dim F As Date

        ChDir(Application.StartupPath)
        ChDrive(Application.StartupPath)

        Ayuda.Class_Initialize()

        Arenas.arenas_iniciar()

        Call BanIpCargar()

        Prision.map = 288
        Prision.x = 53
        Prision.Y = 32

        Libertad.map = 288
        Libertad.x = 53
        Libertad.Y = 68

        NumUsers = 5

        minutos = Format(Now, "Short Time")

        IniPath = Application.StartupPath & "\"
        DatPath = Application.StartupPath & "\Dat\"

        LogError("///////////////////////////////// Se levantó el servidor /////////////////////////////////")

        LevelSkillArr(1).LevelValue = 3
        LevelSkillArr(2).LevelValue = 5
        LevelSkillArr(3).LevelValue = 7
        LevelSkillArr(4).LevelValue = 10
        LevelSkillArr(5).LevelValue = 13
        LevelSkillArr(6).LevelValue = 15
        LevelSkillArr(7).LevelValue = 17
        LevelSkillArr(8).LevelValue = 20
        LevelSkillArr(9).LevelValue = 23
        LevelSkillArr(10).LevelValue = 25
        LevelSkillArr(11).LevelValue = 27
        LevelSkillArr(12).LevelValue = 30
        LevelSkillArr(13).LevelValue = 33
        LevelSkillArr(14).LevelValue = 35
        LevelSkillArr(15).LevelValue = 37
        LevelSkillArr(16).LevelValue = 40
        LevelSkillArr(17).LevelValue = 43
        LevelSkillArr(18).LevelValue = 45
        LevelSkillArr(19).LevelValue = 47
        LevelSkillArr(20).LevelValue = 50
        LevelSkillArr(21).LevelValue = 53
        LevelSkillArr(22).LevelValue = 55
        LevelSkillArr(23).LevelValue = 57
        LevelSkillArr(24).LevelValue = 60
        LevelSkillArr(25).LevelValue = 63
        LevelSkillArr(26).LevelValue = 65
        LevelSkillArr(27).LevelValue = 67
        LevelSkillArr(28).LevelValue = 70
        LevelSkillArr(29).LevelValue = 73
        LevelSkillArr(30).LevelValue = 75
        LevelSkillArr(31).LevelValue = 77
        LevelSkillArr(32).LevelValue = 80
        LevelSkillArr(33).LevelValue = 83
        LevelSkillArr(34).LevelValue = 85
        LevelSkillArr(35).LevelValue = 87
        LevelSkillArr(36).LevelValue = 90
        LevelSkillArr(37).LevelValue = 93
        LevelSkillArr(38).LevelValue = 95
        LevelSkillArr(39).LevelValue = 97
        LevelSkillArr(40).LevelValue = 100
        LevelSkillArr(41).LevelValue = 100
        LevelSkillArr(42).LevelValue = 100
        LevelSkillArr(43).LevelValue = 100
        LevelSkillArr(44).LevelValue = 100
        LevelSkillArr(45).LevelValue = 100
        LevelSkillArr(46).LevelValue = 100
        LevelSkillArr(47).LevelValue = 100
        LevelSkillArr(48).LevelValue = 100
        LevelSkillArr(49).LevelValue = 100
        LevelSkillArr(50).LevelValue = 100


        ListaRazas(eRaza.Humano) = "Humano"
        ListaRazas(eRaza.Elfo) = "Elfo"
        ListaRazas(eRaza.Drow) = "Drow"
        ListaRazas(eRaza.Gnomo) = "Gnomo"
        ListaRazas(eRaza.Enano) = "Enano"
        ListaRazas(eRaza.Orco) = "Orco"

        ListaClases(eClass.Mago) = "Mago"
        ListaClases(eClass.Clerigo) = "Clerigo"
        ListaClases(eClass.Guerrero) = "Guerrero"
        ListaClases(eClass.Asesino) = "Asesino"
        ListaClases(eClass.Ladron) = "Ladron"
        ListaClases(eClass.Bardo) = "Bardo"
        ListaClases(eClass.Druida) = "Druida"
        ListaClases(eClass.Paladin) = "Paladin"
        ListaClases(eClass.Cazador) = "Cazador"
        ListaClases(eClass.Pescador) = "Pescador"
        ListaClases(eClass.Herrero) = "Herrero"
        ListaClases(eClass.Leñador) = "Leñador"
        ListaClases(eClass.Minero) = "Minero"
        ListaClases(eClass.Carpintero) = "Carpintero"
        ListaClases(eClass.Mercenario) = "Mercenario"
        ListaClases(eClass.Nigromante) = "Nigromante"
        ListaClases(eClass.Sastre) = "Sastre"
        ListaClases(eClass.Gladiador) = "Gladiador"

        SkillsNames(eSkill.Magia) = "Magia"
        SkillsNames(eSkill.Robar) = "Robar"
        SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
        SkillsNames(eSkill.armas) = "Combate con armas"
        SkillsNames(eSkill.Meditar) = "Meditar"
        SkillsNames(eSkill.Apuñalar) = "Apuñalar"
        SkillsNames(eSkill.Ocultarse) = "Ocultarse"
        SkillsNames(eSkill.Supervivencia) = "Supervivencia"
        SkillsNames(eSkill.Talar) = "Talar arboles"
        SkillsNames(eSkill.Comerciar) = "Comercio"
        SkillsNames(eSkill.DefensaEscudos) = "Defensa con escudos"
        SkillsNames(eSkill.Pesca) = "Pesca"
        SkillsNames(eSkill.Mineria) = "Mineria"
        SkillsNames(eSkill.Carpinteria) = "Carpinteria"
        SkillsNames(eSkill.Herreria) = "Herreria"
        SkillsNames(eSkill.Liderazgo) = "Liderazgo"
        SkillsNames(eSkill.Domar) = "Domar animales"
        SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
        SkillsNames(eSkill.artes) = "Artes Marciales"
        SkillsNames(eSkill.Navegacion) = "Navegacion"
        SkillsNames(eSkill.alquimia) = "Alquimia"
        SkillsNames(eSkill.arrojadizas) = "Armas Arrojadizas"
        SkillsNames(eSkill.botanica) = "Botanica"
        SkillsNames(eSkill.Equitacion) = "Equitacion"
        SkillsNames(eSkill.Musica) = "Musica"
        SkillsNames(eSkill.Resistencia) = "Resistencia Magica"
        SkillsNames(eSkill.Sastreria) = "Sastreria"

        ListaAtributos(eAtributos.Fuerza) = "Fuerza"
        ListaAtributos(eAtributos.Agilidad) = "Agilidad"
        ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
        ListaAtributos(eAtributos.Carisma) = "Carisma"
        ListaAtributos(eAtributos.constitucion) = "Constitucion"




        IniPath = Application.StartupPath & "\"


        MinXBorder = XMinMapSize + (XWindow \ 2)
        MaxXBorder = XMaxMapSize - (XWindow \ 2)
        MinYBorder = YMinMapSize + (YWindow \ 2)
        MaxYBorder = YMaxMapSize - (YWindow \ 2)
        Application.DoEvents()

        Console.WriteLine("Iniciando Arrays...")

        Application.DoEvents()

        Call CargarSpawnList()


        Console.WriteLine("Conectando con la Base de Datos")


        Application.DoEvents()
        Call CargarDB()


        Console.WriteLine("Cargando Clanes")

        Application.DoEvents()
        Call LoadGuildsDB()


        '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
        Console.WriteLine("Cargando Server.ini")
        Application.DoEvents()

        MaxUsers = 0
        Call LoadSini()
        Call CargaApuestas()

        '*************************************************
        Console.WriteLine("Cargando NPCs.Dat")
        Call CargaNpcsDat()
        'Call insertarNpcsDB

        Application.DoEvents()
        '*************************************************

        Console.WriteLine("Cargando Obj.Dat")
        Call LoadObjData()

        Console.WriteLine("Cargando Hechizos.Dat")
        Call CargarHechizos()


        Console.WriteLine("Cargando Objetos de Herrería")
        Call LoadArmasHerreria()
        Call LoadArmadurasHerreria()

        Console.WriteLine("Cargando Objetos de Carpintería")
        Call LoadObjCarpintero()

        Console.WriteLine("Cargando Objetos de Alquimista")
        Call LoadObjDruida()

        Console.WriteLine("Cargando Objetos de Sastre")

        Call LoadObjSastre()

        Console.WriteLine("Cargando Balance.Dat")
        Call LoadBalance()


        Console.WriteLine("Cargando Mapas")
        Call LoadMapData()


        Dim loopC As Integer

        'MaxUsers = 100
        ReDim UserList(MaxUsers)

        'Resetea las conexiones de los usuarios
        For loopC = 1 To MaxUsers
            UserList(loopC).ConnID = -1
            UserList(loopC).ConnIDValida = False
            UserList(loopC).incomingData = New clsByteQueue
            UserList(loopC).incomingData.Class_Initialize()
            UserList(loopC).outgoingData = New clsByteQueue
            UserList(loopC).outgoingData.Class_Initialize()
        Next loopC


        Call SecurityIp.InitIpTables(1000)


        Dim ctTeleport As Threading.Thread = New Threading.Thread(AddressOf PasarSegundoTelep.PasarSegundotelep)
        ctTeleport.Start()

        Dim ctThread As Threading.Thread = New Threading.Thread(AddressOf PasarSegundo.PasarSegundo)
        ctThread.Start()

        Dim aiThread As Threading.Thread = New Threading.Thread(AddressOf TimerAI.TimerAI)
        aiThread.Start()

        Dim aiAtackThread As Threading.Thread = New Threading.Thread(AddressOf TimerNpcAtaca.TimerNpcAtaca)
        aiAtackThread.Start()

        Dim gameThread As Threading.Thread = New Threading.Thread(AddressOf TimerGame.TimerGame)
        gameThread.Start()

        Dim increaseUsers As Threading.Thread = New Threading.Thread(AddressOf TimerIncreaseUsers.IncreaseUsers)
        increaseUsers.Start()

        Dim waitingConnections As Threading.Thread = New Threading.Thread(AddressOf IniciaWsApi)
        waitingConnections.Start()




        'Dim flushOutGoingData As Threading.Thread = New Threading.Thread(AddressOf TimerFlushData.flushOutGoingData)
        'flushOutGoingData.Start()

        'Dim thSendPosition As Threading.Thread = New Threading.Thread(AddressOf TimerSendPosition.sendPosition)
        'thSendPosition.Start()

        'Dim packetResend As Threading.Thread = New Threading.Thread(AddressOf PackerResend)
        'PacketResend.Start()


        tInicioServer = GetTickCount() And &H7FFFFFFF

        'DB_Conn.Execute("UPDATE `charflags` SET `Online` = 0")
        'DB_Conn.Execute("UPDATE `cuentas` SET `Online` = 0")
        'Call extra_set("EventoX2", "0")
        SendOnline()

        MsgEvento = ""

        'MapInfoArr(Prision.map).ResuSinEfecto = 1

    End Sub

    Function FileExist(ByVal file As String) As Boolean
        '*****************************************************************
        'Se fija si existe el archivo
        '*****************************************************************

        'FileExist = Len(Dir(file, FileType)) <> 0
        FileExist = System.IO.File.Exists(file)
    End Function

    Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
        '*****************************************************************
        'Gets a field from a string
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/15/2004
        'Gets a field from a delimited string
        '*****************************************************************
        Dim i As Long
        Dim LastPos As Long
        Dim CurrentPos As Long
        Dim delimiter As String

        delimiter = Chr(SepASCII)

        For i = 1 To Pos
            LastPos = CurrentPos
            CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
        Next i

        If CurrentPos = 0 Then
            ReadField = Mid(Text, LastPos + 1, Len(Text) - LastPos)
        Else
            ReadField = Mid(Text, LastPos + 1, CurrentPos - LastPos - 1)
        End If
    End Function
    Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
        '*****************************************************************
        'Gets the number of fields in a delimited string
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 07/29/2007
        '*****************************************************************
        Dim count As Long
        Dim curPos As Long
        Dim delimiter As String

        If Len(Text) = 0 Then Exit Function

        delimiter = Chr(SepASCII)

        curPos = 0

        Do
            curPos = InStr(Convert.ToInt32(curPos + 1), Text, delimiter)
            count = count + 1
        Loop While curPos <> 0

        FieldCount = count

    End Function
    Function MapaValido(ByVal map As Integer) As Boolean
        MapaValido = map >= 1 And map <= NumMaps
    End Function


    Public Sub LogCriticEvent(desc As String)
        On Error GoTo Errhandler

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\logs\Eventos.log", True)
        file.WriteLine(DateTime.Now & ": " & desc)
        file.Close()



        Exit Sub

Errhandler:

    End Sub

    Public Sub LogError(desc As String)
        On Error GoTo Errhandler

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\logs\errores.log", True)
        file.WriteLine(DateTime.Now & ": " & desc)
        file.Close()


        Exit Sub

Errhandler:
        Dim a As String
        a = Err.Description

    End Sub

    Public Sub LogGM(Nombre As String, desc As String)
        On Error GoTo Errhandler


        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\logs\GMS\" & Nombre & ".log", True)
        file.WriteLine(DateTime.Now & ": " & desc)
        file.Close()


        Exit Sub

Errhandler:

    End Sub

    Public Sub LogBruteforce(desc As String)
        On Error GoTo Errhandler


        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\logs\BruteForce.log", True)
        file.WriteLine(DateTime.Now & ": " & desc)
        file.Close()


        Exit Sub

Errhandler:

    End Sub

    Public Sub LogHackAttemp(texto As String)
        On Error GoTo Errhandler


        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\logs\HackAttemps.log", True)
        file.WriteLine(DateTime.Now & ": " & texto)
        file.Close()


        Exit Sub

Errhandler:

    End Sub

    Public Sub LogCheating(texto As String)
        On Error GoTo Errhandler


        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\logs\CH.log", True)
        file.WriteLine(DateTime.Now & ": " & texto)
        file.Close()


        Exit Sub

Errhandler:

    End Sub


    Public Sub LogCriticalHackAttemp(texto As String)
        On Error GoTo Errhandler

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\logs\CriticalHackAttemps.log", True)
        file.WriteLine(DateTime.Now & ": " & texto)
        file.Close()


        Exit Sub

Errhandler:

    End Sub

    Public Sub LogAntiCheat(texto As String)
        On Error GoTo Errhandler


        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\logs\AntiCheat.log", True)
        file.WriteLine(DateTime.Now & ": " & texto)
        file.Close()


        Exit Sub

Errhandler:

    End Sub

    Function ValidInputNP(ByVal cad As String) As Boolean
        Dim arg As String
        Dim i As Integer

        For i = 1 To 33
            arg = ReadField(i, cad, 44)

            If Len(arg) = 0 Then Exit Function
        Next i

        ValidInputNP = True

    End Function


    Sub Restart()


        'Se asegura de que los sockets estan cerrados e ignora cualquier err
        On Error Resume Next

        Dim loopC As Long

        For loopC = 1 To MaxUsers
            Call CloseSocket(loopC)
        Next

        'Initialize statistics!!


        For loopC = 1 To UBound(UserList)
            UserList(loopC).incomingData = Nothing
            UserList(loopC).outgoingData = Nothing
        Next loopC

        ReDim UserList(MaxUsers)

        For loopC = 1 To MaxUsers
            UserList(loopC).ConnID = -1
            UserList(loopC).ConnIDValida = False
            UserList(loopC).incomingData = New clsByteQueue
            UserList(loopC).incomingData.Class_Initialize()
            UserList(loopC).outgoingData = New clsByteQueue
            UserList(loopC).outgoingData.Class_Initialize()
        Next loopC

        LastUser = 0
        NumUsers = 0
        SendOnline()

        Call FreeNPCs()
        Call FreeCharIndexes()

        Call LoadSini()
        Call LoadObjData()

        Call LoadMapData()

        Call CargarHechizos()

        'Call frmMain.InitMain(0)

    End Sub



    Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
        Dim i As Integer
        For i = 1 To MAXMASCOTAS
            If UserList(UserIndex).MascotasIndex(i) > 0 Then
                If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia =
           Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
                    If Npclist(UserList(UserIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i), 0)
                End If
            End If
        Next i
    End Sub

    Public Sub EfectoFrio(ByVal UserIndex As Integer)

        Dim modifi As Integer

        If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
            UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
        Else
            If MapInfoArr(UserList(UserIndex).Pos.map).Terreno = Nieve Then
                Call WriteConsoleMsg(1, UserIndex, "¡¡Estas muriendo de frio, abrigate o moriras!!.", FontTypeNames.FONTTYPE_INFO)
                modifi = Porcentaje(UserList(UserIndex).Stats.MaxHP, 5)

                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - modifi

                If UserList(UserIndex).Stats.MinHP < 1 Then
                    Call WriteConsoleMsg(1, UserIndex, "¡¡Has muerto de frio!!.", FontTypeNames.FONTTYPE_INFO)
                    UserList(UserIndex).Stats.MinHP = 0
                    Call UserDie(UserIndex)
                End If

                Call WriteUpdateHP(UserIndex)
            Else
                If UserList(UserIndex).Stats.MinSTA > 0 Then
                    modifi = Porcentaje(UserList(UserIndex).Stats.MaxSTA, 20)
                    Call QuitarSta(UserIndex, modifi)
                    Call WriteUpdateSta(UserIndex)
                End If
            End If

            UserList(UserIndex).Counters.Frio = 0
        End If
    End Sub

    Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)

        If UserList(UserIndex).Counters.Invisibilidad < IntervaloInvisible Then
            UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad + 1
        Else
            UserList(UserIndex).Counters.Invisibilidad = 0
            UserList(UserIndex).flags.Invisible = 0
            If UserList(UserIndex).flags.Oculto = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).cuerpo.CharIndex, False))
            End If
        End If

    End Sub


    Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

        If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
            Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
        Else
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).flags.Inmovilizado = 0
        End If

    End Sub

    Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

        If UserList(UserIndex).Counters.Ceguera > 0 Then
            UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
        Else
            If UserList(UserIndex).flags.Ceguera = 1 Then
                UserList(UserIndex).flags.Ceguera = 0
                Call WriteBlindNoMore(UserIndex)
            End If
            If UserList(UserIndex).flags.Estupidez = 1 Then
                UserList(UserIndex).flags.Estupidez = 0
                Call WriteDumbNoMore(UserIndex)
            End If

        End If

    End Sub


    Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)

        If UserList(UserIndex).Counters.Paralisis > 0 Then
            UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
        Else
            UserList(UserIndex).flags.Paralizado = 0
            UserList(UserIndex).flags.Inmovilizado = 0
            Call WriteParalizeOK(UserIndex)
        End If

    End Sub

    Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

        If UserList(UserIndex).flags.Trabajando Then Exit Sub
        If UserList(UserIndex).flags.Desnudo Then
            If UserList(UserIndex).flags.Montando = 0 Then
                Exit Sub
            End If
        End If

        If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 1 And
       MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 2 And
       MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 4 Then Exit Sub


        Dim massta As Integer
        If UserList(UserIndex).Stats.MinSTA < UserList(UserIndex).Stats.MaxSTA And Not UserList(UserIndex).flags.Entrenando = 1 Then
            If UserList(UserIndex).Counters.STACounter < Intervalo Then
                UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
            Else
                EnviarStats = True
                UserList(UserIndex).Counters.STACounter = 0

                massta = RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSTA + UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia), 20))
                UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MinSTA + massta
                If UserList(UserIndex).Stats.MinSTA > UserList(UserIndex).Stats.MaxSTA Then
                    UserList(UserIndex).Stats.MinSTA = UserList(UserIndex).Stats.MaxSTA
                End If
            End If
        End If

    End Sub

    Public Sub EfectoHechizoMagico(ByVal UserIndex As Integer)


        UserList(UserIndex).Stats.eCreateTipe = 0

        UserList(UserIndex).Stats.eMaxDef = 0
        UserList(UserIndex).Stats.eMinDef = 0
        UserList(UserIndex).Stats.eMaxHit = 0
        UserList(UserIndex).Stats.eMinHit = 0

        UserList(UserIndex).Stats.dMaxDef = 0
        UserList(UserIndex).Stats.dMinDef = 0


    End Sub
    Public Sub EfectoVeneno(ByVal UserIndex As Integer)
        Dim N As Integer

        If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
            UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
        Else
            Call WriteConsoleMsg(1, UserIndex, "Estás envenenado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
            UserList(UserIndex).Counters.Veneno = 0
            N = UserList(UserIndex).flags.Envenenado - 2
            N = RandomNumber(1 * N, 5 * N)

            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - N

            If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
            Call WriteUpdateHP(UserIndex)
        End If

    End Sub
    Public Sub EfectoIncineracion(ByVal UserIndex As Integer)
        Dim N As Integer

        If UserList(UserIndex).Counters.Fuego < IntervaloVeneno Then 'IntervaloFuego Then
            UserList(UserIndex).Counters.Fuego = UserList(UserIndex).Counters.Fuego + 1
        Else
            Call WriteConsoleMsg(2, UserIndex, "Estás incendiado, si no te curas morirás.", FontTypeNames.FONTTYPE_VENENO)
            UserList(UserIndex).Counters.Fuego = 0
            N = RandomNumber(10, 20)

            UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - N

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_INCINERACION, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateCharParticle(UserList(UserIndex).cuerpo.CharIndex, 96))

            If UserList(UserIndex).Stats.MinHP < 1 Then Call UserDie(UserIndex)
            Call WriteUpdateHP(UserIndex)
        End If

    End Sub
    Public Function TieneSacri(ByVal UserIndex As Integer) As Byte
        '****************************************************************
        'Author: Leandro Mendoza (Mannakia)
        'Desc:
        'Last Modify: 21/10/10
        '****************************************************************
        On Error Resume Next
        Dim i As Long
        Dim ObjInd As Integer

        For i = 1 To MAX_INVENTORY_SLOTS
            ObjInd = UserList(UserIndex).Invent.Objeto(i).ObjIndex
            If ObjInd > 0 Then
                If UserList(UserIndex).Invent.Objeto(i).Equipped = 1 And ObjDataArr(ObjInd).EfectoMagico = eMagicType.Sacrificio Then
                    TieneSacri = CByte(i)
                    Exit Function
                End If
            End If
        Next i

        TieneSacri = 0

    End Function
    Public Sub DuracionPociones(ByVal UserIndex As Integer)

        'Controla la duracion de las pociones
        If UserList(UserIndex).flags.DuracionEfecto = 0 Then

            UserList(UserIndex).flags.TomoPocion = False
            UserList(UserIndex).flags.TipoPocion = 0
            'volvemos los atributos al estado normal
            Dim LoopX As Long
            For LoopX = 1 To NUMATRIBUTOS
                UserList(UserIndex).Stats.UserAtributos(LoopX) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopX)
            Next
            Call WriteFuerza(UserIndex)
            Call WriteAgilidad(UserIndex)

            UserList(UserIndex).flags.DuracionEfecto = -1
            Exit Sub
        End If

        UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1


    End Sub

    Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)

        If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Exit Sub

        'Sed
        If UserList(UserIndex).Stats.MinAGU > 0 Then
            If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
                UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
            Else
                UserList(UserIndex).Counters.AGUACounter = 0
                UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10

                If UserList(UserIndex).Stats.MinAGU <= 0 Then
                    UserList(UserIndex).Stats.MinAGU = 0
                    UserList(UserIndex).flags.Sed = 1
                End If

                fenviarAyS = True
            End If
        End If

        'hambre
        If UserList(UserIndex).Stats.MinHAM > 0 Then
            If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
                UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
            Else
                UserList(UserIndex).Counters.COMCounter = 0
                UserList(UserIndex).Stats.MinHAM = UserList(UserIndex).Stats.MinHAM - 10
                If UserList(UserIndex).Stats.MinHAM <= 0 Then
                    UserList(UserIndex).Stats.MinHAM = 0
                    UserList(UserIndex).flags.Hambre = 1
                End If
                fenviarAyS = True
            End If
        End If

    End Sub

    Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
        '***************************************************************************
        'Author: Leandro Mendoza(Mannakia)
        'Desc:
        'LastModify: 21/10/10
        '***************************************************************************
        If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 1 And
   MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 2 And
   MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).Trigger = 4 Then Exit Sub

        Dim mashit As Integer

        'Mannakia
        If UserList(UserIndex).Invent.MagicIndex <> 0 Then
            If ObjDataArr(UserList(UserIndex).Invent.MagicIndex).EfectoMagico = eMagicType.AceleraVida Then
                Intervalo = Intervalo - Porcentaje(Intervalo, 40)  ' nose algo razonable
            End If
        End If
        'Mannakia

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
        If UserList(UserIndex).Stats.MinHP < UserList(UserIndex).Stats.MaxHP Then
            If UserList(UserIndex).Counters.HPCounter < Intervalo Then
                UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
            Else
                mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSTA, 20))

                UserList(UserIndex).Counters.HPCounter = 0
                UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + mashit
                If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
                Call WriteConsoleMsg(1, UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
                EnviarStats = True
            End If
        End If

    End Sub

    Public Sub CargaNpcsDat()
        Dim npcfile As String

        npcfile = DatPath & "NPCs.dat"
        Call LeerNPCs.Initialize(npcfile)
    End Sub



    Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)

        'Guardar Pjs
        Call GuardarUsuarios()
        'Des Nod Kopfnickend Por las dudas, falta la aplicacion para levantarla y parece poco practico
        'If EjecutarLauncher Then Shell (Application.StartupPath & "\launcher.exe")

        'Chauuu
        'Unload frmMain

    End Sub


    Sub GuardarUsuarios()
        haciendoBK = True

        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))

        Dim i As Integer
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                Call SaveUserSQL(i)
            End If
        Next i

        Call LimpiarMundo()

        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

        haciendoBK = False
    End Sub




    Public Sub FreeNPCs()
        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all NPC Indexes
        '***************************************************
        Dim loopC As Long

        ' Free all NPC indexes
        For loopC = 1 To MAXNPCS
            Npclist(loopC).flags.NPCActive = False
        Next loopC
    End Sub

    Public Sub FreeCharIndexes()
        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all char indexes
        '***************************************************
        ' Free all char indexes (set them all to 0)
        Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
    End Sub



    Sub ControlarPortalLum(ByVal UserIndex As Integer)

        If UserList(UserIndex).Counters.CreoTeleport = True Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageDestParticle(UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY))

            MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.map = 0
            MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.x = 0
            MapData(UserList(UserIndex).flags.DondeTiroMap, UserList(UserIndex).flags.DondeTiroX, UserList(UserIndex).flags.DondeTiroY).TileExit.Y = 0
            UserList(UserIndex).flags.DondeTiroMap = 0
            UserList(UserIndex).flags.DondeTiroX = 0
            UserList(UserIndex).flags.DondeTiroY = 0
            UserList(UserIndex).flags.TiroPortalL = 0
            UserList(UserIndex).Counters.TimeTeleport = 0
            UserList(UserIndex).Counters.CreoTeleport = False
        End If

    End Sub


    Public Sub DoMetamorfosis(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer) 'Metamorfosis

        Dim tUser As Integer
        tUser = UserIndex

        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub

        UserList(tUser).cuerpo.Head = Head
        UserList(tUser).cuerpo.body = body
        UserList(tUser).cuerpo.ShieldAnim = NingunEscudo
        UserList(tUser).cuerpo.WeaponAnim = NingunArma
        UserList(tUser).cuerpo.CascoAnim = NingunCasco

        UserList(UserIndex).flags.Metamorfosis = 1

        Call ChangeUserChar(UserIndex, UserList(tUser).cuerpo.body, UserList(tUser).cuerpo.Head, UserList(tUser).cuerpo.heading, UserList(tUser).cuerpo.WeaponAnim, UserList(tUser).cuerpo.ShieldAnim, UserList(tUser).cuerpo.CascoAnim)

    End Sub

    Public Sub EfectoMetamorfosis(ByVal UserIndex As Integer) 'Metamorfosis
        If UserList(UserIndex).Counters.Metamorfosis < IntervaloInvisible * 2 Then
            UserList(UserIndex).Counters.Metamorfosis = UserList(UserIndex).Counters.Metamorfosis + 1
        Else
            UserList(UserIndex).cuerpo.Head = UserList(UserIndex).OrigChar.Head
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                UserList(UserIndex).cuerpo.body = ObjDataArr(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
            Else
                Call DarCuerpoDesnudo(UserIndex)
            End If
            If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).cuerpo.ShieldAnim = ObjDataArr(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).cuerpo.WeaponAnim = ObjDataArr(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).cuerpo.CascoAnim = ObjDataArr(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim

            UserList(UserIndex).flags.Metamorfosis = 0
            UserList(UserIndex).Counters.Metamorfosis = 0

            Call ChangeUserChar(UserIndex, UserList(UserIndex).cuerpo.body, UserList(UserIndex).cuerpo.Head, UserList(UserIndex).cuerpo.heading, UserList(UserIndex).cuerpo.WeaponAnim, UserList(UserIndex).cuerpo.ShieldAnim, UserList(UserIndex).cuerpo.CascoAnim)
        End If
    End Sub


    Public Sub limpiaritemsmundo()
        Static limpma As Integer

        If limpma > NumMaps Then
            limpma = 0
        End If

        limpma = limpma + 1

        Call limpiamundo(limpma)

    End Sub
    'Add Marius
    'Esto bugea las puertas dobles, no esta bueno.
    'Faltaria comprobar si el objeto es del mapa o agarrable
    Public Sub limpiamundo(map As Integer)
        Dim Xnn, Ynn As Byte

        If map < 0 And map > NumMaps Then Exit Sub

        'If MapInfoArr(limpma).Pk = True Then
        For Ynn = YMinMapSize To YMaxMapSize
            For Xnn = XMinMapSize To XMaxMapSize
                If MapData(map, Xnn, Ynn).ObjInfo.ObjIndex > 0 And MapData(map, Xnn, Ynn).Blocked = 0 Then
                    Call EraseObj(10000, map, Val(Xnn), Val(Ynn))
                End If
            Next Xnn
        Next Ynn
        'End If

    End Sub
    '\Add

End Module
