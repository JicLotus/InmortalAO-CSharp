Imports System.Threading

Module TimerIncreaseUsers

    Public Sub IncreaseUsers()

        Dim incremento As Integer
        Dim incrementar As Boolean
        incrementar = True


        While True

            If incrementar Then
                incremento = Rnd() Mod 4 + 1
                NumUsers += incremento
                incrementar = False
            Else
                NumUsers -= incremento
                incrementar = True
            End If

            SendOnline()
            Thread.Sleep(30000)
        End While

    End Sub

End Module
