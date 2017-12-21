Module TimerSendPosition


    Public Sub sendPosition()

        Dim iUserIndex As Integer

        While (True)


            For iUserIndex = 1 To LastUser
                UserList(iUserIndex).client.updatePosition()
            Next

            Threading.Thread.Sleep(400)
            Application.DoEvents()

        End While

    End Sub

End Module
