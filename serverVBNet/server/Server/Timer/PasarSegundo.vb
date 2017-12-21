Imports System.Threading

Module PasarSegundo

    Public Sub PasarSegundo()

        While True
            Try
                For i = 1 To LastUser
                    If UserList(i).flags.UserLogged Then
                        If UserList(i).Counters.Saliendo Then
                            UserList(i).Counters.salir = UserList(i).Counters.salir - 1
                            If UserList(i).Counters.salir < 1 Then
                                Call WriteMsg(i, 43)
                                Call WriteDisconnect(i)
                                UserList(i).client.cerrarConexionCliente()
                            Else
                                Call WriteMsg(i, 27, CStr(UserList(i).Counters.salir))
                            End If
                        End If
                    End If
                Next i
                Thread.Sleep(1000)

            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                Call LogError("Error en PasarSegundo: " & ex.Message & " StackTrae: " & st.ToString())
            End Try
        End While

    End Sub

End Module
