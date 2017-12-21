Imports System.Threading

Public Module TimerAI

    Public Sub TimerAI()

        Dim NpcIndex As Long
        Dim mapa As Integer

        While True
            Try
                If Not haciendoBK And Not EnPausa Then
                    'Update NPCs
                    For NpcIndex = 1 To LastNPC

                        If Npclist(NpcIndex).flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
                            If Npclist(NpcIndex).flags.Paralizado = 1 Then
                                Call EfectoParalisisNpc(NpcIndex)
                            Else
                                'Usamos AI si hay algun user en el mapa
                                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
                                    Call EfectoParalisisNpc(NpcIndex)
                                End If

                                mapa = Npclist(NpcIndex).Pos.map

                                If mapa > 0 Then
                                    If MapInfoArr(mapa).NumUsers > 0 Then
                                        If Npclist(NpcIndex).Movement <> TipoAI.ESTATICO Then
                                            Call NPCAI(NpcIndex)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next NpcIndex
                End If
                Thread.Sleep(300)

            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                Call LogError("Error en PasarSegundotelep: " & ex.Message & " StackTrae: " & st.ToString())
            End Try

        End While

    End Sub


End Module
