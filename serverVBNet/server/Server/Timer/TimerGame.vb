Imports System.Threading

Module TimerGame


    Public Sub TimerGame()


        Dim iUserIndex As Long
        Dim bEnviarStats As Boolean
        Dim bEnviarAyS As Boolean

        While True
            Try
                For iUserIndex = 1 To LastUser

                    With UserList(iUserIndex)

                        If .ConnID <> -1 Then
                            If .ConnIDValida = True And .flags.UserLogged = True Then


                                UserList(iUserIndex).client.puedeEscribirOutGoing = False

                                    bEnviarStats = False
                                    bEnviarAyS = False

                                    If .flags.Muerto = 0 Then
                                        If .flags.Paralizado = 1 Then
                                            Call EfectoParalisisUser(iUserIndex)
                                        End If

                                        If .flags.Ceguera = 1 Or .flags.Estupidez Then
                                            Call EfectoCegueEstu(iUserIndex)
                                        End If

                                        If .flags.Metamorfosis = 1 Then
                                            Call EfectoMetamorfosis(iUserIndex)
                                        End If

                                        If Not EsCONSE(iUserIndex) Then
                                            If .flags.Montando = 0 And .flags.Desnudo <> 0 Then
                                                Call EfectoFrio(iUserIndex)
                                            End If

                                            If .flags.Envenenado <> 0 Then
                                                Call EfectoVeneno(iUserIndex)
                                            End If

                                            If .flags.Incinerado <> 0 Then
                                                Call EfectoIncineracion(iUserIndex)
                                            End If

                                            If .Stats.eCreateTipe <> 0 Then
                                                Call EfectoHechizoMagico(iUserIndex)
                                            End If
                                        End If

                                        If .flags.Meditando Then
                                            Call DoMeditar(iUserIndex)
                                        End If

                                        If .flags.AdminInvisible <> 1 Then
                                            If .flags.Invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
                                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                                        End If
                                        If .flags.Trabajando Then DoTrabajar(iUserIndex)

                                        Call DuracionPociones(iUserIndex)

                                        Call HambreYSed(iUserIndex, bEnviarAyS)

                                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                                            If Not .flags.Descansar Then
                                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
                                                If bEnviarStats Then
                                                    Call WriteUpdateHP(iUserIndex)
                                                    bEnviarStats = False
                                                End If

                                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)

                                                If bEnviarStats Then
                                                    Call WriteUpdateSta(iUserIndex)
                                                    bEnviarStats = False
                                                End If

                                            Else
                                                Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
                                                If bEnviarStats Then
                                                    Call WriteUpdateHP(iUserIndex)
                                                    bEnviarStats = False
                                                End If

                                                Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
                                                If bEnviarStats Then
                                                    Call WriteUpdateSta(iUserIndex)
                                                    bEnviarStats = False
                                                End If

                                                If .Stats.MaxHP = .Stats.MinHP And .Stats.MaxSTA = .Stats.MinSTA Then
                                                    Call WriteRestOK(iUserIndex)
                                                    Call WriteConsoleMsg(1, iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                                    .flags.Descansar = False
                                                End If

                                            End If
                                        End If

                                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)

                                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                                    Else
                                        If .flags.Resucitando <> 0 Then Call DoResucitar(iUserIndex)
                                    End If

                                    If .Counters.Silenciado > 0 Then
                                        .Counters.Silenciado = .Counters.Silenciado - 1
                                    End If


                                    UserList(iUserIndex).client.puedeEscribirOutGoing = True


                                End If
                        End If


                    End With

                Next iUserIndex

                Thread.Sleep(100)
            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                LogError("Error en GameTimer: " & Err.Description & " UserIndex = " & iUserIndex & "\n Stack:" & st.ToString())
            End Try

        End While


    End Sub

End Module
