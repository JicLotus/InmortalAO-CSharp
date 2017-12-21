﻿Option Explicit On

Module NPCs
    '?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    '?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    '?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    '                        Modulo NPC
    '?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    '?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    '?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    'Contiene todas las rutinas necesarias para cotrolar los
    'NPCs meno la rutina de AI que se encuentra en el modulo
    'AI_NPCs para su mejor comprension.
    '?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    '?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    '?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿


    Dim list_items(2000) As Integer
    Dim list_last As Integer

    Sub List_AddItem(ByVal a As Integer)
        Dim i As Long

        If list_last = 0 Then
            list_last = 1
            list_items(1) = a
            Exit Sub
        End If

        For i = 1 To list_last
            If list_items(i) = a Then
                Exit Sub
            End If
        Next i

        list_last = list_last + 1
        list_items(list_last) = a
        'List1.AddItem a
    End Sub

    Sub List_OrdenarListar()
        Dim count As Long
        Dim i As Long
        Dim F As Long
        Dim Index As Long

        count = list_last
        Do While count <> 0
            F = 10000
            For i = 1 To list_last
                If list_items(i) < F Then
                    F = list_items(i)
                    Index = i
                End If
            Next i

            WriteVar(Application.StartupPath & "\npcusados.txt", "init", "NPC" & F, "1")
            If Not Index = list_last Then _
            list_items(Index) = list_items(list_last)

            list_last = list_last - 1
            count = count - 1
        Loop

    End Sub
    Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        Dim i As Integer

        For i = 1 To MAXMASCOTAS
            If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
                UserList(UserIndex).MascotasIndex(i) = 0
                UserList(UserIndex).MascotasType(i) = 0

                UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
                Exit For
            End If
        Next i
    End Sub

    Sub QuitarMascotaNpc(ByVal Maestro As Integer)
        Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
    End Sub

    Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)


        Try
            Dim MiNPC As npc
            MiNPC = Npclist(NpcIndex)
            'Quitamos el npc
            Call QuitarNPC(NpcIndex)

            If UserIndex > 0 Then ' Lo mato un usuario?
                If MiNPC.flags.Snd3 > 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.x, MiNPC.Pos.Y))
                End If
                UserList(UserIndex).flags.TargetNPC = 0
                UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun

                'El user que lo mato tiene mascotas?
                If UserList(UserIndex).NroMascotas > 0 Then
                    Dim t As Integer
                    For t = 1 To MAXMASCOTAS
                        If UserList(UserIndex).MascotasIndex(t) > 0 Then
                            If Npclist(UserList(UserIndex).MascotasIndex(t)).TargetNPC = NpcIndex Then
                                Call FollowAmo(UserList(UserIndex).MascotasIndex(t))
                            End If
                        End If
                    Next t
                End If

                '[KEVIN]
                ' If MiNPC.GiveEXP > MAXEXP Then MiNPC.GiveEXP = 0

                Dim experienciaGanada As Long
                experienciaGanada = (MiNPC.GiveEXP) * ModExpX

                If MiNPC.GiveEXP > 0 Then

                    If UserList(UserIndex).GrupoIndex > 0 Then
                        Call mdGrupo.ObtenerExito(UserIndex, MiNPC.GiveEXP, MiNPC.Pos.map, MiNPC.Pos.x, MiNPC.Pos.Y)
                    Else

                        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + experienciaGanada

                        If UserList(UserIndex).Stats.Exp > MAXEXP Then
                            UserList(UserIndex).Stats.Exp = MAXEXP
                        End If

                        Call WriteMsg(UserIndex, 21, CStr(MiNPC.GiveEXP))

                    End If
                    MiNPC.flags.ExpCount = 0
                End If

                '[/KEVIN]
                Call WriteConsoleMsg(2, UserIndex, "¡Has matado a la criatura! Ganaste " & experienciaGanada & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)


                If UserList(UserIndex).Stats.NPCsMuertos < 32000 Then _
            UserList(UserIndex).Stats.NPCsMuertos = UserList(UserIndex).Stats.NPCsMuertos + 1


                Call CheckUserLevel(UserIndex)

            End If ' Userindex > 0

            If MiNPC.MaestroUser = 0 Then
                'Tiramos el oro
                Call NPCTirarOro(MiNPC, UserIndex)
                'Tiramos el inventario
                Call NPC_TIRAR_ITEMS(MiNPC)
                'ReSpawn o no
                Call ReSpawnNpc(MiNPC)
            Else
                If MiNPC.IsFamiliar Then
                    Call UpdateFamiliar(MiNPC.MaestroUser, False)
                End If
            End If


            Exit Sub

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.Description & " NpcIndex: " & NpcIndex & " Userindex: " & UserIndex & " st:" & st.ToString())
        End Try

    End Sub

    Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)
        'Clear the npc's flags

        Npclist(NpcIndex).IsFamiliar = False

        With Npclist(NpcIndex).flags
            .AfectaParalisis = 0
            .AguaValida = 0
            .AttackedBy = 0
            .AttackedFirstBy = vbNullString
            .BackUp = 0
            .Domable = 0
            .Envenenado = 0
            .Incinerado = 0
            .faccion = 0
            .Follow = False
            .AtacaDoble = 0
            .LanzaSpells = 0
            .Invisible = 0
            .OldHostil = 0
            .OldMovement = 0
            .Paralizado = 0
            .Inmovilizado = 0
            .respawn = 0
            .RespawnOrigPos = 0
            .Snd1 = 0
            .Snd2 = 0
            .Snd3 = 0
            .TierraInvalida = 0
        End With
    End Sub

    Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)
        With Npclist(NpcIndex).Contadores
            .Paralisis = 0
            .TiempoExistencia = 0
        End With
    End Sub

    Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
        With Npclist(NpcIndex).cuerpo
            .body = 0
            .CascoAnim = 0
            .CharIndex = 0
            .fx = 0
            .Head = 0
            .heading = 0
            .loops = 0
            .ShieldAnim = 0
            .WeaponAnim = 0
        End With
    End Sub

    Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
        Dim j As Long

        With Npclist(NpcIndex)
            For j = 1 To .NroCriaturas
                .Criaturas(j).NpcIndex = 0
                .Criaturas(j).NpcName = vbNullString
            Next j

            .NroCriaturas = 0
        End With
    End Sub

    Sub ResetExpresiones(ByVal NpcIndex As Integer)
        Dim j As Long

        With Npclist(NpcIndex)
            For j = 1 To .NroExpresiones
                .Expresiones(j) = vbNullString
            Next j

            .NroExpresiones = 0
        End With
    End Sub

    Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
        With Npclist(NpcIndex)
            .Attackable = 0
            .CanAttack = 0
            .Comercia = 0
            .GiveEXP = 0
            .GiveGLD = 0
            .Hostile = 0
            .InvReSpawn = 0

            If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, NpcIndex)
            If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)

            .MaestroUser = 0
            .MaestroNpc = 0

            .Mascotas = 0
            .Movement = 0
            .Name = vbNullString
            .NPCtype = 0
            .Numero = 0
            .Orig.map = 0
            .Orig.x = 0
            .Orig.Y = 0
            .PoderAtaque = 0
            .PoderEvasion = 0
            .Pos.map = 0
            .Pos.x = 0
            .Pos.Y = 0
            .SkillDomar = 0
            .Target = 0
            .TargetNPC = 0
            .TipoItems = 0
            .Veneno = 0
            .Fuego = 0
            .desc = vbNullString


            Dim j As Long
            For j = 1 To .NroSpells
                .Spells(j) = 0
            Next j
        End With

        Call ResetNpcCharInfo(NpcIndex)
        Call ResetNpcCriatures(NpcIndex)
        Call ResetExpresiones(NpcIndex)
    End Sub

    Public Sub QuitarNPC(ByVal NpcIndex As Integer)

        On Error GoTo Errhandler

        With Npclist(NpcIndex)
            .flags.NPCActive = False

            If InMapBounds(.Pos.map, .Pos.x, .Pos.Y) Then
                Call EraseNPCChar(NpcIndex)
            End If
        End With

        'Nos aseguramos de que el inventario sea removido...
        'asi los lobos no volveran a tirar armaduras ;))
        Call ResetNpcInv(NpcIndex)
        Call ResetNpcFlags(NpcIndex)
        Call ResetNpcCounters(NpcIndex)

        Call ResetNpcMainInfo(NpcIndex)

        If NpcIndex = LastNPC Then
            Do Until Npclist(LastNPC).flags.NPCActive
                LastNPC = LastNPC - 1
                If LastNPC < 1 Then Exit Do
            Loop
        End If


        If NumNPCs <> 0 Then
            NumNPCs = NumNPCs - 1
        End If
        Exit Sub

Errhandler:
        Call LogError("Error en QuitarNPC")
    End Sub

    Private Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean

        If LegalPos(Pos.map, Pos.x, Pos.Y, PuedeAgua) Then
            TestSpawnTrigger =
        MapData(Pos.map, Pos.x, Pos.Y).Trigger <> 3 And
        MapData(Pos.map, Pos.x, Pos.Y).Trigger <> 2 And
        MapData(Pos.map, Pos.x, Pos.Y).Trigger <> 1
        End If

    End Function

    Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
        'Call LogTarea("Sub CrearNPC")
        'Crea un NPC del tipo NRONPC

        Dim Pos As WorldPos
        Dim NewPos As WorldPos
        Dim altpos As WorldPos
        Dim nIndex As Integer
        Dim PosicionValida As Boolean
        Dim Iteraciones As Long
        Dim PuedeAgua As Boolean
        Dim PuedeTierra As Boolean


        Dim map As Integer
        Dim x As Integer
        Dim Y As Integer

        nIndex = OpenNPC(NroNPC) 'Conseguimos un indice

        If nIndex > MAXNPCS Then Exit Sub

        PuedeAgua = Npclist(nIndex).flags.AguaValida
        PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)

        'Necesita ser respawned en un lugar especifico
        If InMapBounds(OrigPos.map, OrigPos.x, OrigPos.Y) Then

            map = OrigPos.map
            x = OrigPos.x
            Y = OrigPos.Y
            Npclist(nIndex).Orig = OrigPos
            Npclist(nIndex).Pos = OrigPos

        Else

            Pos.map = mapa 'mapa
            altpos.map = mapa
            PosicionValida = False

            Do While Not PosicionValida

                Pos.x = RandomNumber(MinXBorder, MaxXBorder)
                Pos.Y = RandomNumber(MinYBorder, MaxYBorder)

                NewPos = getNpcArroundValidPosition(Pos.map, Pos.x, Pos.Y, PuedeAgua)

                If NewPos.x = 0 And NewPos.Y = 0 Then
                    Iteraciones = Iteraciones + 1
                    If Iteraciones > 5 Then
                        'Faltaria poner al indice del npc en una cola de espera para un futuro respawn
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & mapa & " NroNpc:" & NroNPC)
                        Exit Sub
                    End If
                Else
                    Npclist(nIndex).Pos = NewPos
                    PosicionValida = True
                End If
            Loop

            'asignamos las nuevas coordenas
            map = NewPos.map
            x = Npclist(nIndex).Pos.x
            Y = Npclist(nIndex).Pos.Y
        End If


        Npclist(nIndex).StartPos = Npclist(nIndex).Pos

        'Crea el NPC
        Call MakeNPCChar(True, map, nIndex, map, x, Y)

    End Sub


    Public Function getNpcArroundValidPosition(ByVal map As Integer, ByVal randomX As Integer, ByVal randomY As Integer, ByVal PuedeAgua As Boolean) As WorldPos

        On Error GoTo ErrHandler

        Dim actualPosX As Integer = randomX
        Dim actualPosY As Integer = randomX
        Dim actualPosMap As Integer = map


        Dim xTemp As Integer
        Dim yTemp As Integer

        Dim tempPosition As WorldPos
        tempPosition.x = actualPosX
        tempPosition.Y = actualPosY
        tempPosition.map = actualPosMap


        If Not MapData(map, tempPosition.x, tempPosition.Y).UserIndex And
            LegalPosNPC(tempPosition.map, tempPosition.x, tempPosition.Y, PuedeAgua) And
             TestSpawnTrigger(tempPosition, PuedeAgua) Then

            getNpcArroundValidPosition = tempPosition
            Exit Function
        End If


        Dim limXMin As Integer
        Dim limYMin As Integer
        Dim limXMax As Integer
        Dim limYMax As Integer

        Dim cantidadBusqueda As Integer = 1

        While (actualPosX > 9 And actualPosX < 90 And actualPosY < 90 And actualPosY > 9)

            actualPosY = actualPosY - 1
            actualPosX = actualPosX - 1

            For yTemp = actualPosY To actualPosY + (cantidadBusqueda) * 2
                For xTemp = actualPosX To actualPosX + (cantidadBusqueda) * 2

                    If xTemp <= limXMax And xTemp >= limXMin And yTemp <= limXMax And yTemp >= limYMin Then
                        Continue For
                    End If

                    If Not MapData(map, xTemp, yTemp).UserIndex And
                    LegalPosNPC(actualPosMap, xTemp, yTemp, PuedeAgua) Then
                        tempPosition.x = xTemp
                        tempPosition.Y = yTemp
                        tempPosition.map = actualPosMap

                        If TestSpawnTrigger(tempPosition, PuedeAgua) Then
                            getNpcArroundValidPosition = tempPosition
                            Exit Function
                        End If
                    End If

                Next xTemp
            Next yTemp


            limXMin = actualPosX
            limYMin = actualPosY
            limXMax = xTemp
            limYMax = yTemp
            cantidadBusqueda = cantidadBusqueda + 1

        End While

        tempPosition.x = 0
        tempPosition.Y = 0
        getNpcArroundValidPosition = tempPosition

        Exit Function

ErrHandler:
        tempPosition.x = 0
        tempPosition.Y = 0
        getNpcArroundValidPosition = tempPosition


    End Function



    Public Sub MakeNPCChar(ByVal ToMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal Y As Integer)
        Dim CharIndex As Integer

        If Npclist(NpcIndex).cuerpo.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex()
            Npclist(NpcIndex).cuerpo.CharIndex = CharIndex
            CharList(CharIndex) = NpcIndex
        End If

        MapData(map, x, Y).NpcIndex = NpcIndex

        If Not ToMap Then

            Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).cuerpo.body, Npclist(NpcIndex).cuerpo.Head, Npclist(NpcIndex).cuerpo.heading, Npclist(NpcIndex).cuerpo.CharIndex, x, Y, 0, 0, 0, 0, 0, vbNullString, 0, False, False)

            Call FlushBuffer(sndIndex)
        Else
            Call AgregarNpc(NpcIndex)
        End If
    End Sub

    Public Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading)
        If NpcIndex > 0 Then
            With Npclist(NpcIndex).cuerpo
                .body = body
                .Head = Head
                .heading = heading

                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, 0, 0, 0, 0, 0))
            End With
        End If
    End Sub

    Private Sub EraseNPCChar(ByVal NpcIndex As Integer)

        If Npclist(NpcIndex).cuerpo.CharIndex <> 0 Then CharList(Npclist(NpcIndex).cuerpo.CharIndex) = 0

        If Npclist(NpcIndex).cuerpo.CharIndex = LastChar Then
            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1
                If LastChar <= 1 Then Exit Do
            Loop
        End If

        'Quitamos del mapa
        MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

        'Actualizamos los clientes
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(Npclist(NpcIndex).cuerpo.CharIndex))

        'Update la lista npc
        Npclist(NpcIndex).cuerpo.CharIndex = 0


        'update NumChars
        NumChars = NumChars - 1


    End Sub

    Public Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 06/04/2009
        '06/04/2009: ZaMa - Now npcs can force to change position with dead character
        '***************************************************

        On Error GoTo errh
        Dim nPos As WorldPos
        Dim UserIndex As Integer

        With Npclist(NpcIndex)
            nPos = .Pos
            Call HeadtoPos(nHeading, nPos)

            ' es una posicion legal
            If LegalPosNPC(.Pos.map, nPos.x, nPos.Y, .flags.AguaValida = 1) Then

                If .flags.AguaValida = 0 And HayAgua(.Pos.map, nPos.x, nPos.Y) Then Exit Sub
                If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.map, nPos.x, nPos.Y) Then Exit Sub

                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.cuerpo.CharIndex, nPos.x, nPos.Y))

                UserIndex = MapData(.Pos.map, nPos.x, nPos.Y).UserIndex
                ' Si hay un usuario a dodne se mueve el npc, entonces esta muerto
                If UserIndex > 0 Then
                    With UserList(UserIndex)
                        ' Actualizamos posicion y mapa
                        MapData(.Pos.map, .Pos.x, .Pos.Y).UserIndex = 0
                        .Pos.x = Npclist(NpcIndex).Pos.x
                        .Pos.Y = Npclist(NpcIndex).Pos.Y
                        MapData(.Pos.map, .Pos.x, .Pos.Y).UserIndex = UserIndex

                        ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
                        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).cuerpo.CharIndex, .Pos.x, .Pos.Y))
                        Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))
                    End With
                End If

                'Update map and user pos
                MapData(.Pos.map, .Pos.x, .Pos.Y).NpcIndex = 0
                .oldPos = .Pos
                .Pos = nPos
                .cuerpo.heading = nHeading
                MapData(.Pos.map, nPos.x, nPos.Y).NpcIndex = NpcIndex
                Call CheckUpdateNeededNpc(NpcIndex, nHeading)

            End If
        End With
        Exit Sub

errh:
        LogError("Error en move npc " & NpcIndex)
    End Sub

    Function NextOpenNPC() As Integer
        'Call LogTarea("Sub NextOpenNPC")

        On Error GoTo Errhandler
        Dim loopC As Long

        For loopC = 1 To MAXNPCS + 1
            If loopC > MAXNPCS Then Exit For
            If Not Npclist(loopC).flags.NPCActive Then Exit For
        Next loopC

        NextOpenNPC = loopC
        Exit Function

Errhandler:
        Call LogError("Error en NextOpenNPC")
    End Function

    Sub NpcEnvenenarUser(ByVal UserIndex As Integer)

        Dim N As Integer
        N = RandomNumber(1, 100)
        If N < 30 Then
            UserList(UserIndex).flags.Envenenado = 3
            Call WriteConsoleMsg(2, UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
        End If

    End Sub

    Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal fx As Boolean, ByVal respawn As Boolean) As Integer
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 06/15/2008
        '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
        '06/15/2008 -> Optimizé el codigo. (NicoNZ)
        '***************************************************
        Dim NewPos As WorldPos
        Dim altpos As WorldPos
        Dim nIndex As Integer
        Dim PosicionValida As Boolean
        Dim PuedeAgua As Boolean
        Dim PuedeTierra As Boolean


        Dim map As Integer
        Dim x As Integer
        Dim Y As Integer

        nIndex = OpenNPC(NpcIndex, respawn)   'Conseguimos un indice

        If nIndex > MAXNPCS Then
            SpawnNpc = 0
            Exit Function
        End If

        If nIndex = 0 Then Exit Function

        PuedeAgua = Npclist(nIndex).flags.AguaValida
        PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1

        Call ClosestLegalPos(Pos, NewPos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
        Call ClosestLegalPos(Pos, altpos, PuedeAgua)
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida

        If NewPos.x <> 0 And NewPos.Y <> 0 Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).Pos.map = NewPos.map
            Npclist(nIndex).Pos.x = NewPos.x
            Npclist(nIndex).Pos.Y = NewPos.Y
            PosicionValida = True
        Else
            If altpos.x <> 0 And altpos.Y <> 0 Then
                Npclist(nIndex).Pos.map = altpos.map
                Npclist(nIndex).Pos.x = altpos.x
                Npclist(nIndex).Pos.Y = altpos.Y
                PosicionValida = True
            Else
                PosicionValida = False
            End If
        End If


        Npclist(nIndex).StartPos.map = Npclist(nIndex).Pos.map
        Npclist(nIndex).StartPos.x = Npclist(nIndex).Pos.x
        Npclist(nIndex).StartPos.Y = Npclist(nIndex).Pos.Y

        If Not PosicionValida Then
            Call QuitarNPC(nIndex)
            SpawnNpc = 0
            Exit Function
        End If

        'asignamos las nuevas coordenas
        map = NewPos.map
        x = Npclist(nIndex).Pos.x
        Y = Npclist(nIndex).Pos.Y

        'Crea el NPC
        Call MakeNPCChar(True, map, nIndex, map, x, Y)

        If fx Then
            Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, x, Y))
            Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).cuerpo.CharIndex, FXIDs.FXWARP, 0))
        End If

        SpawnNpc = nIndex

    End Function

    Sub ReSpawnNpc(MiNPC As npc)

        If (MiNPC.flags.respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.map, MiNPC.Orig)

    End Sub

    Private Sub NPCTirarOro(ByRef MiNPC As npc, ByVal UserIndex As Integer)

        'SI EL NPC TIENE ORO LO TIRAMOS
        If UserIndex > 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiNPC.GiveGLD * ModOroX
            Call WriteMsg(UserIndex, 36, CStr(MiNPC.GiveGLD * ModOroX))
            Call WriteUpdateGold(UserIndex)
        Else
            Dim MiAux As Long
            Dim MiObj As obj
            MiAux = MiNPC.GiveGLD
            Do While MiAux > MAX_INVENTORY_OBJS
                MiObj.Amount = MAX_INVENTORY_OBJS
                MiObj.ObjIndex = iORO
                Call TirarItemAlPiso(MiNPC.Pos, MiObj)
                MiAux = MiAux - MAX_INVENTORY_OBJS
            Loop
            If MiAux > 0 Then
                MiObj.Amount = MiAux
                MiObj.ObjIndex = iORO
                Call TirarItemAlPiso(MiNPC.Pos, MiObj)
            End If
        End If
    End Sub

    Public Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal respawn As Boolean = True) As Integer

        '###################################################
        '#               ATENCION PELIGRO                  #
        '###################################################
        '
        '    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
        '
        'El que ose desafiar esta LEY, se las tendrá que ver
        'conmigo. Para leer los NPCS se deberá usar la
        'nueva clase clsIniReader.
        '
        'Alejo
        '
        '###################################################
        Dim NpcIndex As Integer
        Dim leer As clsIniReader
        Dim loopC As Long
        Dim ln As String
        Dim aux As String

        leer = LeerNPCs


        'If requested index is invalid, abort
        If Not leer.KeyExists("NPC" & NpcNumber) Then
            'OpenNPC = MAXNPCS + 1
            Exit Function
        End If

        NpcIndex = NextOpenNPC()

        If NpcIndex > MAXNPCS Then 'Limite de npcs
            OpenNPC = NpcIndex
            Exit Function
        End If

        With Npclist(NpcIndex)
            .Numero = NpcNumber
            .Name = leer.GetValue("NPC" & NpcNumber, "Name")
            .desc = leer.GetValue("NPC" & NpcNumber, "Desc")

            .Movement = Val(leer.GetValue("NPC" & NpcNumber, "Movement"))
            .flags.OldMovement = .Movement

            .flags.AguaValida = Val(leer.GetValue("NPC" & NpcNumber, "AguaValida"))
            .flags.TierraInvalida = Val(leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
            .flags.faccion = Val(leer.GetValue("NPC" & NpcNumber, "Faccion"))
            .flags.AtacaDoble = Val(leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))

            .NPCtype = Val(leer.GetValue("NPC" & NpcNumber, "NpcType"))

            .cuerpo.body = Val(leer.GetValue("NPC" & NpcNumber, "Body"))
            .cuerpo.Head = Val(leer.GetValue("NPC" & NpcNumber, "Head"))
            .cuerpo.heading = Val(leer.GetValue("NPC" & NpcNumber, "Heading"))

            .Attackable = Val(leer.GetValue("NPC" & NpcNumber, "Attackable"))
            .Comercia = Val(leer.GetValue("NPC" & NpcNumber, "Comercia"))
            .Hostile = Val(leer.GetValue("NPC" & NpcNumber, "Hostile"))
            .flags.OldHostil = .Hostile

            .GiveEXP = Val(leer.GetValue("NPC" & NpcNumber, "GiveEXP"))

            .flags.ExpCount = .GiveEXP

            .Veneno = Val(leer.GetValue("NPC" & NpcNumber, "Veneno"))

            .flags.Domable = Val(leer.GetValue("NPC" & NpcNumber, "Domable"))

            .GiveGLD = Val(leer.GetValue("NPC" & NpcNumber, "GiveGLD"))

            .PoderAtaque = Val(leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
            .PoderEvasion = Val(leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))

            .InvReSpawn = Val(leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))

            With .Stats
                .MaxHP = Val(leer.GetValue("NPC" & NpcNumber, "MaxHP"))
                .MinHP = Val(leer.GetValue("NPC" & NpcNumber, "MinHP"))
                .MaxHit = Val(leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
                .MinHit = Val(leer.GetValue("NPC" & NpcNumber, "MinHIT"))
                .def = Val(leer.GetValue("NPC" & NpcNumber, "DEF"))
                .defM = Val(leer.GetValue("NPC" & NpcNumber, "DEFm"))
                .Alineacion = Val(leer.GetValue("NPC" & NpcNumber, "Alineacion"))
            End With

            Dim Prob As Long
            .Invent.NroItems = Val(leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
            ReDim .Invent.Objeto(MAX_INVENTORY_SLOTS + 1)
            If .Invent.NroItems > 0 Then
                For loopC = 1 To .Invent.NroItems
                    ln = leer.GetValue("NPC" & NpcNumber, "Obj" & loopC)
                    .Invent.Objeto(loopC).ObjIndex = Val(ReadField(1, ln, 45))
                    .Invent.Objeto(loopC).Amount = Val(ReadField(2, ln, 45))

                    If FieldCount(ln, 45) < 3 Then
                        .Invent.Objeto(loopC).Prob = 100
                    Else
                        Prob = Val(ReadField(3, ln, 45))
                        .Invent.Objeto(loopC).Prob = Prob
                    End If
                Next loopC
            End If

            .flags.LanzaSpells = Val(leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
            If .flags.LanzaSpells > 0 Then ReDim .Spells(.flags.LanzaSpells)

            ReDim .Spells(.flags.LanzaSpells)
            For loopC = 1 To .flags.LanzaSpells
                .Spells(loopC) = Val(leer.GetValue("NPC" & NpcNumber, "Sp" & loopC))
            Next loopC

            If .NPCtype = eNPCType.Entrenador Then
                .NroCriaturas = Val(leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
                ReDim .Criaturas(.NroCriaturas)
                For loopC = 1 To .NroCriaturas
                    .Criaturas(loopC).NpcIndex = leer.GetValue("NPC" & NpcNumber, "CI" & loopC)
                    .Criaturas(loopC).NpcName = leer.GetValue("NPC" & NpcNumber, "CN" & loopC)
                Next loopC
            End If

            With .flags
                .NPCActive = True

                If respawn Then
                    .respawn = Val(leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
                Else
                    .respawn = 1
                End If

                .BackUp = Val(leer.GetValue("NPC" & NpcNumber, "BackUp"))
                .RespawnOrigPos = Val(leer.GetValue("NPC" & NpcNumber, "OrigPos"))
                .AfectaParalisis = Val(leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))

                .Snd1 = Val(leer.GetValue("NPC" & NpcNumber, "Snd1"))
                .Snd2 = Val(leer.GetValue("NPC" & NpcNumber, "Snd2"))
                .Snd3 = Val(leer.GetValue("NPC" & NpcNumber, "Snd3"))
            End With

            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
            .NroExpresiones = Val(leer.GetValue("NPC" & NpcNumber, "NROEXP"))
            If .NroExpresiones > 0 Then ReDim .Expresiones(.NroExpresiones)
            For loopC = 1 To .NroExpresiones
                .Expresiones(loopC) = leer.GetValue("NPC" & NpcNumber, "Exp" & loopC)
            Next loopC
            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

            'Tipo de items con los que comercia
            .TipoItems = Val(leer.GetValue("NPC" & NpcNumber, "TipoItems"))

            .faccion = Val(leer.GetValue("NPC" & NpcNumber, "Faccion"))

            If (NpcNumber = 610) Then
                'MsgBox .Name
            End If

        End With

        'Update contadores de NPCs
        If NpcIndex > LastNPC Then LastNPC = NpcIndex
        NumNPCs = NumNPCs + 1

        'Devuelve el nuevo Indice
        OpenNPC = NpcIndex




    End Function

    Public Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
        With Npclist(NpcIndex)
            If .flags.Follow Then
                .flags.AttackedBy = 0
                .flags.Follow = False
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            Else
                .flags.AttackedBy = NameIndex(UserName)
                .flags.Follow = True
                .Movement = TipoAI.NPCDEFENSA
                .Hostile = 0
            End If
        End With
    End Sub

    Public Sub FollowAmo(ByVal NpcIndex As Integer)
        With Npclist(NpcIndex)
            .flags.Follow = True
            If .IsFamiliar Then
                .Movement = TipoAI.NpcFamiliar
            Else
                .Movement = TipoAI.SigueAmo
            End If
            .Hostile = 0
            .Target = 0
            .TargetNPC = 0
        End With
    End Sub
    Public Function npcFamiToTipe(ByVal tipe As eMascota) As Integer
        Select Case tipe
            Case eMascota.Agua
                npcFamiToTipe = 127

            Case eMascota.Ely
                npcFamiToTipe = 132

            Case eMascota.Ent
                npcFamiToTipe = 145

            Case eMascota.Fatuo
                npcFamiToTipe = 126

            Case eMascota.Fuego
                npcFamiToTipe = 128

            Case eMascota.Lobo
                npcFamiToTipe = 133

            Case eMascota.Oso
                npcFamiToTipe = 131

            Case eMascota.Tierra
                npcFamiToTipe = 129

            Case eMascota.Tigre
                npcFamiToTipe = 130
        End Select
    End Function

End Module
