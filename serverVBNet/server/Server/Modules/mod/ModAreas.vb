Option Explicit On
Imports System.Collections.Generic

Module ModAreas


    Public Structure AreaInfo
        Dim AreaPerteneceX As Integer
        Dim AreaPerteneceY As Integer

        Dim AreaReciveX As Integer
        Dim AreaReciveY As Integer

        Dim AreaID As Long
    End Structure


    Public Const USER_NUEVO As Byte = 255

    'Cuidado:
    ' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
    Private CurDay As Byte
    Private CurHour As Byte

    Private AreasInfo(100, 100) As Long

    Private AreasRecive(12) As Long

    Public ConnGroups() As List(Of Integer)

    Public Sub InitAreas()
        Dim loopC As Long
        Dim LoopX As Long

        For loopC = 0 To 11
            AreasRecive(loopC) = (2 ^ loopC) Or IIf(loopC <> 0, 2 ^ (loopC - 1), 0) Or IIf(loopC <> 11, 2 ^ (loopC + 1), 0)
        Next loopC

        For loopC = 1 To 100
            For LoopX = 1 To 100
                AreasInfo(loopC, LoopX) = (loopC \ 4 + 1) * (LoopX \ 4 + 1)
            Next LoopX
        Next loopC

        ReDim ConnGroups(NumMaps + 1)

        For loopC = 1 To NumMaps
            ConnGroups(loopC) = New List(Of Integer)
        Next loopC


    End Sub


    Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte)

        Try

            If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y) Then Exit Sub

            Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, y As Long
            Dim map As Long

            With UserList(UserIndex)

                If Head = eHeading.NORTH Then
                    MinY = .Pos.Y - 11
                    MaxY = .Pos.Y - 5

                    MinX = .Pos.x - 9
                    MaxX = .Pos.x + 9
                ElseIf Head = eHeading.SOUTH Then
                    MinY = .Pos.Y + 5
                    MaxY = .Pos.Y + 11

                    MinX = .Pos.x - 9
                    MaxX = .Pos.x + 9
                ElseIf Head = eHeading.WEST Then
                    MinY = .Pos.Y - 6
                    MaxY = .Pos.Y + 6

                    MinX = .Pos.x - 13
                    MaxX = .Pos.x - 7
                ElseIf Head = eHeading.EAST Then
                    MinY = .Pos.Y - 6
                    MaxY = .Pos.Y + 6

                    MinX = .Pos.x + 7
                    MaxX = .Pos.x + 13
                ElseIf Head = USER_NUEVO Then
                    MinY = .Pos.Y - 10
                    MaxY = .Pos.Y + 10

                    MinX = .Pos.x - 12
                    MaxX = .Pos.x + 12
                End If

                If MinY < 1 Then MinY = 1
                If MinX < 1 Then MinX = 1
                If MaxY > 100 Then MaxY = 100
                If MaxX > 100 Then MaxX = 100

                map = UserList(UserIndex).Pos.map

                'Esto es para ke el cliente elimine lo "fuera de area..."
                Call WriteAreaChanged(UserIndex)

                Dim usuariosEnPantalla As New List(Of indexPosition)
                Dim npcsEnPantalla As New List(Of indexPosition)
                Dim objetosEnPantalla As New List(Of indexPosition)

                Dim temp As indexPosition

                'Actualizamos!!!
                For x = MinX To MaxX
                    For y = MinY To MaxY

                        temp.x = x
                        temp.y = y

                        If MapData(map, x, y).UserIndex Then
                            temp.index = MapData(map, x, y).UserIndex
                            usuariosEnPantalla.Add(temp)
                        End If

                        If MapData(map, x, y).NpcIndex Then
                            temp.index = MapData(map, x, y).NpcIndex
                            npcsEnPantalla.Add(temp)
                        End If

                        If MapData(map, x, y).ObjInfo.ObjIndex Then
                            temp.index = MapData(map, x, y).ObjInfo.ObjIndex
                            objetosEnPantalla.Add(temp)
                        End If

                    Next y
                Next x



                Dim indiceEnPantalla As Integer
                For indiceEnPantalla = 0 To usuariosEnPantalla.Count - 1

                    temp = usuariosEnPantalla(indiceEnPantalla)

                    If UserIndex <> temp.index Then
                        Call MakeUserChar(False, UserIndex, temp.index, map, temp.x, temp.y)
                        Call MakeUserChar(False, temp.index, UserIndex, .Pos.map, .Pos.x, .Pos.Y)

                        If UserList(temp.index).flags.Invisible Or UserList(temp.index).flags.Oculto Then
                            Call WriteSetInvisible(UserIndex, UserList(temp.index).cuerpo.CharIndex, True)
                        End If

                        If UserList(UserIndex).flags.Invisible Or UserList(UserIndex).flags.Oculto Then
                            Call WriteSetInvisible(temp.index, UserList(UserIndex).cuerpo.CharIndex, True)
                        End If

                    ElseIf Head = USER_NUEVO Then
                        Call MakeUserChar(False, UserIndex, UserIndex, map, temp.x, temp.y)
                    End If

                Next indiceEnPantalla


                For indiceEnPantalla = 0 To npcsEnPantalla.Count - 1
                    temp = npcsEnPantalla(indiceEnPantalla)
                    Call MakeNPCChar(False, UserIndex, temp.index, map, temp.x, temp.y)
                Next indiceEnPantalla


                For indiceEnPantalla = 0 To objetosEnPantalla.Count - 1
                    temp = objetosEnPantalla(indiceEnPantalla)
                    If Not EsObjetoFijo(ObjDataArr(temp.index).OBJType) And MapData(map, temp.x, temp.y).ObjEsFijo = 0 Then
                        Call WriteObjectCreate(UserIndex, temp.x, temp.y, temp.index, MapData(map, temp.x, temp.y).ObjInfo.Amount)

                        If ObjDataArr(temp.index).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(False, UserIndex, temp.x, temp.y, MapData(map, temp.x, temp.y).Blocked)
                            Call Bloquear(False, UserIndex, temp.x - 1, temp.y, MapData(map, temp.x - 1, temp.y).Blocked)
                        End If
                    End If
                Next indiceEnPantalla



                Dim TempInt As Long

                TempInt = .Pos.x \ 9
                .AreasInfo.AreaReciveX = AreasRecive(TempInt)
                .AreasInfo.AreaPerteneceX = 2 ^ TempInt

                TempInt = .Pos.Y \ 9
                .AreasInfo.AreaReciveY = AreasRecive(TempInt)
                .AreasInfo.AreaPerteneceY = 2 ^ TempInt

                If .Pos.x <> 0 And .Pos.Y <> 0 Then
                    .AreasInfo.AreaID = AreasInfo(.Pos.x, .Pos.Y)
                End If



            End With

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            Call LogError("Error en CheckUpdateNeededUser: " & ex.Message & " StackTrae: " & st.ToString())
        End Try


    End Sub

    Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        ' Se llama cuando se mueve un Npc
        '**************************************************************
        If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y) Then Exit Sub

        Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, Y As Long
        Dim TempInt As Long

        With Npclist(NpcIndex)
            MinY = .Pos.Y - 10
            MaxY = .Pos.Y + 10

            MinX = .Pos.x - 12
            MaxX = .Pos.x + 12

            If MinY < 1 Then MinY = 1
            If MinX < 1 Then MinX = 1
            If MaxY > 100 Then MaxY = 100
            If MaxX > 100 Then MaxX = 100

            'Actualizamos!!!
            If MapInfoArr(.Pos.map).NumUsers <> 0 Then
                For x = MinX To MaxX
                    For Y = MinY To MaxY
                        If MapData(.Pos.map, x, Y).UserIndex Then _
                        Call MakeNPCChar(False, MapData(.Pos.map, x, Y).UserIndex, NpcIndex, .Pos.map, .Pos.x, .Pos.Y)
                    Next Y
                Next x
            End If

            'Precalculados :P
            TempInt = .Pos.x \ 9
            .AreasInfo.AreaReciveX = AreasRecive(TempInt)
            .AreasInfo.AreaPerteneceX = 2 ^ TempInt

            TempInt = .Pos.Y \ 9
            .AreasInfo.AreaReciveY = AreasRecive(TempInt)
            .AreasInfo.AreaPerteneceY = 2 ^ TempInt

            .AreasInfo.AreaID = AreasInfo(.Pos.x, .Pos.Y)
        End With
    End Sub

    Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal map As Integer)

        Try

            ConnGroups(map).Remove(UserIndex)

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            LogError("Error en QuitarUser: " + ex.Message + " StackTrace: " + st.ToString())
        End Try

    End Sub

    Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal map As Integer)
        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: 04/01/2007
        'Modified by Juan Martín Sotuyo Dodero (Maraxus)
        '   - Now the method checks for repetead users instead of trusting parameters.
        '   - If the character is new to the map, update it
        '**************************************************************
        Dim EsNuevo As Boolean

        Try

            If Not MapaValido(map) Then Exit Sub

            EsNuevo = True

            'Prevent adding repeated users
            If EsNuevo Then
                ConnGroups(map).Add(UserIndex)
            End If

            'Update user
            UserList(UserIndex).AreasInfo.AreaID = 0

            UserList(UserIndex).AreasInfo.AreaPerteneceX = 0
            UserList(UserIndex).AreasInfo.AreaPerteneceY = 0
            UserList(UserIndex).AreasInfo.AreaReciveX = 0
            UserList(UserIndex).AreasInfo.AreaReciveY = 0

            Call CheckUpdateNeededUser(UserIndex, USER_NUEVO)

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            LogError("Error en AgregarUser: " + ex.Message + " StackTrace: " + st.ToString())
        End Try


    End Sub

    Public Sub AgregarNpc(ByVal NpcIndex As Integer)
        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Npclist(NpcIndex).AreasInfo.AreaID = 0

        Npclist(NpcIndex).AreasInfo.AreaPerteneceX = 0
        Npclist(NpcIndex).AreasInfo.AreaPerteneceY = 0
        Npclist(NpcIndex).AreasInfo.AreaReciveX = 0
        Npclist(NpcIndex).AreasInfo.AreaReciveY = 0

        Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
    End Sub



End Module
