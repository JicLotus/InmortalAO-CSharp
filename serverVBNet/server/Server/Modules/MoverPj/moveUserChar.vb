Module moveUserChar

    Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)

        Dim nPos As WorldPos

        nPos = UserList(UserIndex).Pos
        Call HeadtoPos(nHeading, nPos)

        corroborarMonturaDungeon(UserIndex)
        determinarYEnviarPosicionUsuario(UserIndex, nPos, nHeading)
        disminuirCounters(UserIndex)

    End Sub

    Private Sub determinarYEnviarPosicionUsuario(ByVal UserIndex As Integer, ByVal nPos As WorldPos, ByVal nHeading As eHeading)

        Dim sailing As Boolean
        Dim indiceDeUsuarioEnLaPosicionFutura As Integer
        Dim isAdminInvi As Boolean

        sailing = PuedeAtravesarAgua(UserIndex)

        isAdminInvi = (UserList(UserIndex).flags.AdminInvisible = 1)
        indiceDeUsuarioEnLaPosicionFutura = MapData(UserList(UserIndex).Pos.map, nPos.x, nPos.Y).UserIndex

        Try
            If MoveToLegalPos(UserList(UserIndex).Pos.map, nPos.x, nPos.Y, sailing, Not sailing) Then

                If MapInfoArr(UserList(UserIndex).Pos.map).NumUsers > 1 Then
                    'determinarQueHacerConMuertos(indiceDeUsuarioEnLaPosicionFutura, isAdminInvi, nHeading)
                    determinarQueHacerAdminInvisible(UserIndex, isAdminInvi, nPos)
                End If

                actualizarPosicionesYUserIndexAlCaminar(UserIndex, isAdminInvi, indiceDeUsuarioEnLaPosicionFutura, nPos, nHeading)
            End If

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            Call LogError("Error en DeterminarYEnviarPosicionUsuario: " & ex.Message & " StackTrae: " & st.ToString())
        End Try



        'If UserList(UserIndex).Pos.x = UserList(UserIndex).client.xAnterior And UserList(UserIndex).Pos.Y = UserList(UserIndex).client.yAnterior Then
        ' Exit Sub
        ' End If

        Call WritePosUpdate(UserIndex)

        UserList(UserIndex).client.xAnterior = UserList(UserIndex).Pos.x
        UserList(UserIndex).client.yAnterior = UserList(UserIndex).Pos.Y

    End Sub

    Private Sub actualizarPosicionesYUserIndexAlCaminar(ByVal UserIndex As Integer, ByVal isAdminInvi As Boolean, ByVal indiceDeUsuarioEnLaPosicionFutura As Integer, ByVal nPos As WorldPos, ByVal nHeading As eHeading)

        ' Los admins invisibles no pueden patear caspers
        '        If Not (isAdminInvi And (indiceDeUsuarioEnLaPosicionFutura <> 0)) Then

        With UserList(UserIndex)
            MapData(.Pos.map, .Pos.x, .Pos.Y).UserIndex = 0
            .Pos = nPos
            .cuerpo.heading = nHeading
            MapData(.Pos.map, .Pos.x, .Pos.Y).UserIndex = UserIndex

            Call DoTileEvents(UserIndex, .Pos.map, .Pos.x, .Pos.Y)
        End With

        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading)
        'End If

    End Sub

    Private Sub determinarQueHacerAdminInvisible(ByVal UserIndex As Integer, ByVal isAdminInvi As Boolean, ByVal nPos As WorldPos)
        ' Si es un admin invisible, no se avisa a los demas clientes
        If Not isAdminInvi Then
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).cuerpo.CharIndex, nPos.x, nPos.Y))
        End If

    End Sub

    Private Sub determinarQueHacerConMuertos(ByVal indiceDeUsuarioEnLaPosicionFutura As Integer, ByVal isAdminInvi As Boolean, ByVal nHeading As eHeading)

        Dim CasperHeading As eHeading
        Dim CasPerPos As WorldPos

        'Si hay un usuario, y paso la validacion, entonces es un casper
        If indiceDeUsuarioEnLaPosicionFutura > 0 Then

            If Not isAdminInvi Then

                If UserList(indiceDeUsuarioEnLaPosicionFutura).flags.Muerto Then
                    CasperHeading = InvertHeading(nHeading)
                    CasPerPos = UserList(indiceDeUsuarioEnLaPosicionFutura).Pos
                    Call HeadtoPos(CasperHeading, CasPerPos)

                    With UserList(indiceDeUsuarioEnLaPosicionFutura)

                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not .flags.AdminInvisible = 1 Then
                            ' Call SendData(SendTarget.ToPCAreaButIndex, indiceDeUsuarioEnLaPosicionFutura, PrepareMessageCharacterMove(.cuerpo.CharIndex, CasPerPos.x, CasPerPos.Y))
                        End If

                        'Call WriteForceCharMove(indiceDeUsuarioEnLaPosicionFutura, CasperHeading)

                        'Update map and user pos
                        .Pos = CasPerPos
                        .cuerpo.heading = CasperHeading
                        MapData(.Pos.map, CasPerPos.x, CasPerPos.Y).UserIndex = indiceDeUsuarioEnLaPosicionFutura
                    End With

                    'Actualizamos las áreas de ser necesario
                    'Call ModAreas.CheckUpdateNeededUser(indiceDeUsuarioEnLaPosicionFutura, CasperHeading)
                End If

            End If
        End If

    End Sub

    Private Sub disminuirCounters(ByVal UserIndex As Integer)

        If UserList(UserIndex).Counters.Trabajando Then
            UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando - 1
        End If

        If UserList(UserIndex).Counters.Ocultando Then
            UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
        End If

    End Sub


    Private Sub corroborarMonturaDungeon(ByVal UserIndex As Integer)

        If UserList(UserIndex).flags.Montando = 1 And (MapInfoArr(UserList(UserIndex).Pos.map).Zona = "DUNGEON" Or UserList(UserIndex).Pos.map = Bandera_mapa) Then

            UserList(UserIndex).flags.Montando = 0
            If UserList(UserIndex).flags.Muerto = 0 Then
                UserList(UserIndex).cuerpo.Head = UserList(UserIndex).OrigChar.Head
                If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                    UserList(UserIndex).cuerpo.body = ObjDataArr(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
                Else
                    Call DarCuerpoDesnudo(UserIndex)
                End If
            Else
                UserList(UserIndex).cuerpo.body = iCuerpoMuerto
                UserList(UserIndex).cuerpo.Head = iCabezaMuerto
                UserList(UserIndex).cuerpo.ShieldAnim = NingunEscudo
                UserList(UserIndex).cuerpo.WeaponAnim = NingunArma
                UserList(UserIndex).cuerpo.CascoAnim = NingunCasco
            End If
            UserList(UserIndex).Invent.Objeto(UserList(UserIndex).Invent.MonturaSlot).Equipped = 0
            Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MonturaSlot)

            UserList(UserIndex).Invent.MonturaObjIndex = 0
            UserList(UserIndex).Invent.MonturaSlot = 0

            Call ChangeUserChar(UserIndex, UserList(UserIndex).cuerpo.body, UserList(UserIndex).cuerpo.Head, UserList(UserIndex).cuerpo.heading, UserList(UserIndex).cuerpo.WeaponAnim, UserList(UserIndex).cuerpo.ShieldAnim, UserList(UserIndex).cuerpo.CascoAnim)
            Call WriteEquitateToggle(UserIndex)

        End If

    End Sub


End Module
