Imports System.Threading

Module PasarSegundoTelep
    Sub PasarSegundotelep()
        Dim i As Integer, mapa As Integer, x As Integer, Y As Integer

        While True
            Try
                For i = 1 To LastUser
                    If UserList(i).Counters.CreoTeleport = True Then
                        mapa = UserList(i).flags.DondeTiroMap
                        x = UserList(i).flags.DondeTiroX
                        Y = UserList(i).flags.DondeTiroY

                        UserList(i).Counters.TimeTeleport = UserList(i).Counters.TimeTeleport + 1

                        If UserList(i).Counters.TimeTeleport = 5 Then
                            Call SendData(SendTarget.ToPCArea, i, PrepareMessageDestParticle(x, Y))
                            Call SendData(SendTarget.ToPCArea, i, PrepareMessageCreateParticle(x, Y, 34))
                            MapData(UserList(i).Pos.map, x, Y).TileExit.map = 49
                            MapData(UserList(i).Pos.map, x, Y).TileExit.x = 50
                            MapData(UserList(i).Pos.map, x, Y).TileExit.Y = 50

                        Else
                            If UserList(i).Counters.TimeTeleport = 18 Then
                                Call SendData(SendTarget.ToPCArea, i, PrepareMessageDestParticle(x, Y))
                                Call ControlarPortalLum(i)
                            End If
                        End If
                    End If
                Next i

                Thread.Sleep(1000)

            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                Call LogError("Error en PasarSegundotelep: " & ex.Message & " StackTrae: " & st.ToString())
            End Try

        End While

    End Sub
End Module
