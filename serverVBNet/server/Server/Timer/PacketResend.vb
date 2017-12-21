Module PacketResend

    Public Sub PackerResend()

        Dim iUserIndex As Integer

        While True
            For iUserIndex = 1 To LastUser
                With UserList(iUserIndex)
                    If Not UserList(iUserIndex).client Is Nothing Then
                        ' If UserList(iUserIndex).client.puedeEscribirOutGoing Then
                        'UserList(iUserIndex).client.puedeEscribirOutGoing = False
                        'UserList(iUserIndex).client.flushOutGoingData()
                        'UserList(iUserIndex).client.puedeEscribirOutGoing = True
                        'End If
                    End If
                End With
            Next iUserIndex

        End While


    End Sub



End Module
