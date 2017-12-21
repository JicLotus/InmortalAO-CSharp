Imports System.Threading

Module TimerNpcAtaca

    Public Sub TimerNpcAtaca()

        Dim npc As Long

        While True

            Try
                For npc = 1 To LastNPC
                    Npclist(npc).CanAttack = 1
                Next npc
                Thread.Sleep(4000)

            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                Call LogError("Error en PasarSegundotelep: " & ex.Message & " StackTrae: " & st.ToString())
            End Try

        End While

    End Sub

End Module
