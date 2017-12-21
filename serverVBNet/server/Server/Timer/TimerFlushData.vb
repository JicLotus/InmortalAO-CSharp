Module TimerFlushData

    Public Sub flushOutGoingData()

        Dim iUserIndex As Integer

        While (True)
            Try

                For iUserIndex = 1 To LastUser

                    If UserList(iUserIndex).ConnID = -1 Or Not UserList(iUserIndex).client.clientSocket.Connected Then
                        UserList(iUserIndex).client.cerrarConexionCliente()
                        Exit While
                    End If

                    UserList(iUserIndex).client.flushOutGoingData()

                Next iUserIndex



            Catch ex As Exception
                UserList(iUserIndex).client.cerrarConexionCliente()
                Exit While
            End Try

            Application.DoEvents()
            Threading.Thread.Sleep(300)

        End While

    End Sub


End Module
