Option Explicit On
Imports System.Text

Module modSendData

    Public Enum SendTarget
        ToAll = 1
        ToMap
        ToPCArea
        ToAllButIndex
        ToMapButIndex
        ToGM
        ToNPCArea
        ToGuildMembers
        ToADMINS
        ToPCAreaButIndex
        ToDiosesYclan
        ToClanArea
        ToDeadArea
        ToGrupoArea

        ToReal
        ToCaos
        Tomili

        ToImpeMap
        ToRepuMap
    End Enum

    Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)



        Dim loopC As Long
        Dim map As Integer
        Dim tempIndex As Integer

        Try

            Select Case sndRoute
                Case SendTarget.ToPCArea
                    SendToUserArea.SendToUserArea(sndIndex, sndData)
                    Exit Sub

                Case SendTarget.ToAllButIndex
                    For loopC = 1 To LastUser
                        If (UserList(loopC).ConnID <> -1) And (loopC <> sndIndex) Then
                            If UserList(loopC).flags.UserLogged Then 'Esta logeado como usuario?
                                Call EnviarDatosASlot(loopC, sndData)
                            End If
                        End If
                    Next loopC
                    Application.DoEvents()
                    Exit Sub

                Case SendTarget.ToMapButIndex
                    Call SendToMapButIndex(sndIndex, sndData)
                    Exit Sub

                Case SendTarget.ToADMINS
                    For loopC = 1 To LastUser
                        If UserList(loopC).ConnID <> -1 Then
                            If EsDIOS(loopC) Then
                                Call EnviarDatosASlot(loopC, sndData)
                            End If
                        End If
                    Next loopC
                    Application.DoEvents()
                    Exit Sub

                Case SendTarget.ToGM
                    For loopC = 1 To LastUser
                        If UserList(loopC).ConnID <> -1 Then
                            If EsCONSE(loopC) Then
                                Call EnviarDatosASlot(loopC, sndData)
                            End If
                        End If
                    Next loopC
                    Application.DoEvents()
                    Exit Sub

                Case SendTarget.ToAll
                    For loopC = 1 To LastUser
                        If UserList(loopC).ConnID <> -1 Then
                            If UserList(loopC).flags.UserLogged Then 'Esta logeado como usuario?
                                Call EnviarDatosASlot(loopC, sndData)
                            End If
                        End If
                    Next loopC
                    Application.DoEvents()
                    Exit Sub

                Case SendTarget.ToMap
                    Call SendToMap(sndIndex, sndData)
                    Exit Sub

                Case SendTarget.ToGuildMembers
                    loopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                    While loopC > 0
                        If (UserList(loopC).ConnID <> -1) Then
                            Call EnviarDatosASlot(loopC, sndData)
                        End If
                        loopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                    End While
                    Application.DoEvents()
                    Exit Sub

                Case SendTarget.ToDeadArea
                    Call SendToDeadUserArea(sndIndex, sndData)
                    Exit Sub

                Case SendTarget.ToPCAreaButIndex
                    Call SendToUserAreaButindex(sndIndex, sndData)
                    Exit Sub

                Case SendTarget.ToClanArea
                    Call SendToUserGuildArea(sndIndex, sndData)
                    Exit Sub

                Case SendTarget.ToGrupoArea
                    Call SendToUserGrupoArea(sndIndex, sndData)
                    Exit Sub

                Case SendTarget.ToNPCArea
                    Call SendToNpcArea(sndIndex, sndData)
                    Exit Sub


                Case SendTarget.ToReal
                    For loopC = 1 To LastUser
                        If (UserList(loopC).ConnID <> -1) Then
                            If UserList(loopC).flags.UserLogged And (UserList(loopC).faccion.ArmadaReal = 1 Or EsCONSE(loopC)) Then  'Esta logeado como usuario?
                                Call EnviarDatosASlot(loopC, sndData)
                            End If
                        End If
                    Next loopC
                    Application.DoEvents()
                    Exit Sub

                Case SendTarget.Tomili
                    For loopC = 1 To LastUser
                        If (UserList(loopC).ConnID <> -1) Then
                            If UserList(loopC).flags.UserLogged And (UserList(loopC).faccion.Milicia = 1 Or EsCONSE(loopC)) Then  'Esta logeado como usuario?
                                Call EnviarDatosASlot(loopC, sndData)
                            End If
                        End If
                    Next loopC
                    Application.DoEvents()
                    Exit Sub

                Case SendTarget.ToCaos
                    For loopC = 1 To LastUser
                        If (UserList(loopC).ConnID <> -1) Then
                            If UserList(loopC).flags.UserLogged And (UserList(loopC).faccion.FuerzasCaos = 1 Or EsCONSE(loopC)) Then    'Esta logeado como usuario?
                                Call EnviarDatosASlot(loopC, sndData)
                            End If
                        End If
                    Next loopC
                    Application.DoEvents()
                    Exit Sub


                Case SendTarget.ToImpeMap
                    If Not MapaValido(sndIndex) Then Exit Sub

                    For loopC = 0 To ConnGroups(sndIndex).Count - 1

                        If loopC >= ConnGroups(map).Count Then Exit Sub

                        tempIndex = ConnGroups(sndIndex)(loopC)

                        If UserList(tempIndex).ConnIDValida And (esCiuda(tempIndex) Or esArmada(tempIndex)) Then
                            Call EnviarDatosASlot(tempIndex, sndData)
                        End If
                    Next loopC
                    Application.DoEvents()
                    Exit Sub

                Case SendTarget.ToRepuMap
                    If Not MapaValido(sndIndex) Then Exit Sub

                    For loopC = 0 To ConnGroups(sndIndex).Count - 1

                        If loopC >= ConnGroups(map).Count Then Exit Sub

                        tempIndex = ConnGroups(sndIndex)(loopC)

                        If UserList(tempIndex).ConnIDValida And (esRepu(tempIndex) Or esMili(tempIndex)) Then
                            Call EnviarDatosASlot(tempIndex, sndData)
                        End If
                    Next loopC
                    Application.DoEvents()
                    Exit Sub
                    '\Add

            End Select


        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            LogError("Error en SendData: " + ex.Message + " StackTrace: " + st.ToString())
        End Try

    End Sub



    Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)


        Dim loopC As Long
        Dim TempInt As Integer
        Dim tempIndex As Integer

        Dim map As Integer
        Dim AreaX As Integer
        Dim AreaY As Integer

        Try
            map = UserList(UserIndex).Pos.map
            AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
            AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

            If Not MapaValido(map) Then Exit Sub

            For loopC = 0 To ConnGroups(map).Count - 1

                If loopC >= ConnGroups(map).Count Then Exit Sub

                tempIndex = ConnGroups(map)(loopC)

                TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
                If TempInt Then  'Esta en el area?
                    TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                    If TempInt Then

                        If tempIndex <> UserIndex Then
                            If UserList(tempIndex).ConnIDValida Then

                                Call EnviarDatosASlot(tempIndex, sdData)

                            End If
                        End If
                    End If
                End If
            Next loopC

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            LogError("Error en SendToUserAreaButindex: " + ex.Message + " StackTrace: " + st.ToString())
        End Try

    End Sub

    Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim loopC As Long
        Dim tempIndex As Integer

        Dim map As Integer
        Dim AreaX As Integer
        Dim AreaY As Integer

        map = UserList(UserIndex).Pos.map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

        If Not MapaValido(map) Then Exit Sub

        For loopC = 0 To ConnGroups(map).Count - 1

            If loopC >= ConnGroups(map).Count Then Exit Sub

            tempIndex = ConnGroups(map)(loopC)

            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    'Dead and admins read
                    If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (EsDIOS(tempIndex)) <> 0) Then
                        Call EnviarDatosASlot(tempIndex, sdData)
                    End If
                End If
            End If
        Next loopC
        Application.DoEvents()
    End Sub

    Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim loopC As Long
        Dim tempIndex As Integer

        Dim map As Integer
        Dim AreaX As Integer
        Dim AreaY As Integer

        map = UserList(UserIndex).Pos.map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

        If Not MapaValido(map) Then Exit Sub

        If UserList(UserIndex).GuildIndex = 0 Then Exit Sub

        For loopC = 0 To ConnGroups(map).Count - 1

            If loopC >= ConnGroups(map).Count Then Exit Sub

            tempIndex = ConnGroups(map)(loopC)

            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    If UserList(tempIndex).ConnIDValida And (UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex Or ((EsDIOS(tempIndex)) And (UserList(tempIndex).flags.Privilegios And PlayerType.VIP) = 0)) Then
                        Call EnviarDatosASlot(tempIndex, sdData)
                    End If
                End If
            End If
        Next loopC
        Application.DoEvents()
    End Sub

    Private Sub SendToUserGrupoArea(ByVal UserIndex As Integer, ByVal sdData As String)
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim loopC As Long
        Dim tempIndex As Integer

        Dim map As Integer
        Dim AreaX As Integer
        Dim AreaY As Integer

        map = UserList(UserIndex).Pos.map
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

        If Not MapaValido(map) Then Exit Sub

        If UserList(UserIndex).GrupoIndex = 0 Then Exit Sub

        For loopC = 0 To ConnGroups(map).Count - 1

            If loopC >= ConnGroups(map).Count Then Exit Sub

            tempIndex = ConnGroups(map)(loopC)

            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    If UserList(tempIndex).ConnIDValida And UserList(tempIndex).GrupoIndex = UserList(UserIndex).GrupoIndex Then
                        Call EnviarDatosASlot(tempIndex, sdData)
                    End If
                End If
            End If
        Next loopC
        Application.DoEvents()
    End Sub

    Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)
        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim loopC As Long
        Dim TempInt As Integer
        Dim tempIndex As Integer

        Dim map As Integer
        Dim AreaX As Integer
        Dim AreaY As Integer

        map = Npclist(NpcIndex).Pos.map
        AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
        AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY

        If Not MapaValido(map) Then Exit Sub

        For loopC = 0 To ConnGroups(map).Count - 1

            If loopC >= ConnGroups(map).Count Then Exit Sub

            tempIndex = ConnGroups(map)(loopC)

            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
            If TempInt Then  'Esta en el area?
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

                If UserList(tempIndex).Pos.map = map Then

                    If TempInt Then
                        If UserList(tempIndex).ConnIDValida Then
                            Call EnviarDatosASlot(tempIndex, sdData)
                        End If
                    End If

                End If

            End If
        Next loopC
        Application.DoEvents()
    End Sub

    Public Sub SendToAreaByPos(ByVal map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)
        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim loopC As Long
        Dim TempInt As Integer
        Dim tempIndex As Integer

        AreaX = 2 ^ (AreaX \ 9)
        AreaY = 2 ^ (AreaY \ 9)

        If Not MapaValido(map) Then Exit Sub

        For loopC = 0 To ConnGroups(map).Count - 1
            tempIndex = ConnGroups(map)(loopC)

            TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX
            If TempInt Then  'Esta en el area?
                TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY
                If TempInt Then
                    If UserList(tempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(tempIndex, sdData)
                    End If
                End If
            End If
        Next loopC
        Application.DoEvents()
    End Sub

    Public Sub SendToMap(ByVal map As Integer, ByVal sdData As String)
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 5/24/2007
        '
        '**************************************************************
        Dim loopC As Long
        Dim tempIndex As Integer

        If Not MapaValido(map) Then Exit Sub

        For loopC = 0 To ConnGroups(map).Count - 1

            If loopC >= ConnGroups(map).Count Then Exit Sub

            tempIndex = ConnGroups(map)(loopC)

            If UserList(tempIndex).ConnIDValida Then
                Call EnviarDatosASlot(tempIndex, sdData)
            End If
        Next loopC
        Application.DoEvents()
    End Sub

    Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sdData As String)
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 5/24/2007
        '
        '**************************************************************
        Dim loopC As Long
        Dim map As Integer
        Dim tempIndex As Integer

        map = UserList(UserIndex).Pos.map

        If Not MapaValido(map) Then Exit Sub

        For loopC = 0 To ConnGroups(map).Count - 1

            If loopC >= ConnGroups(map).Count Then Exit Sub

            tempIndex = ConnGroups(map)(loopC)

            If tempIndex = 0 Then Continue For

            If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
                Call EnviarDatosASlot(tempIndex, sdData)
            End If
        Next loopC
        Application.DoEvents()
    End Sub

End Module
