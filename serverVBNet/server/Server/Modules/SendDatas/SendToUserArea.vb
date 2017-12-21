Module SendToUserArea

    Public Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

        Dim loopC As Long
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
                If loopC < ConnGroups(map).Count Then

                    If loopC >= ConnGroups(map).Count Then Exit Sub

                    tempIndex = ConnGroups(map)(loopC)

                    If tempIndex = 0 Then Continue For

                    If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                        If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then

                            If UserList(tempIndex).ConnIDValida Then
                                Call EnviarDatosASlot(tempIndex, sdData)
                            End If

                        End If
                    End If
                Else
                    Exit For
                End If
            Next loopC
        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            LogError("Error en SendToUserArea: " + ex.Message + " StackTrace: " + st.ToString())
        End Try




    End Sub

End Module
