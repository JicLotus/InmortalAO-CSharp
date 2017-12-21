Option Explicit On

Module Matematicas


    Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
        Porcentaje = (Total * Porc) * 0.01
    End Function

    Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
        On Error GoTo hayerror
        'Encuentra la distancia entre dos WorldPos
        Distancia = Math.Abs(wp1.x - wp2.x) + Math.Abs(wp1.Y - wp2.Y) + (Math.Abs(wp1.map - wp2.map) * 100)

        Exit Function
hayerror:
        LogError("Error en Distancia: " & Err.Description)


    End Function
    Function RangoVision(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos, ByVal tHeading As Byte) As Boolean
        Dim SignoNS As Integer
        Dim SignoEO As Integer

        Select Case tHeading
            Case eHeading.NORTH
                SignoNS = -1 : SignoEO = 0
            Case eHeading.EAST
                SignoNS = 0 : SignoEO = 1
            Case eHeading.SOUTH
                SignoNS = 1 : SignoEO = 0
            Case eHeading.WEST
                SignoEO = -1 : SignoNS = 0
        End Select

        If Math.Abs(wp1.x - wp2.x) <= RANGO_VISION_X And
       Math.Sign(wp1.x - wp2.x) = SignoEO Then
            If Math.Abs(wp1.Y - wp2.Y) <= RANGO_VISION_Y And
           Math.Sign(wp1.Y - wp2.Y) = SignoNS Then
                RangoVision = True
                Exit Function
            End If
        End If
        RangoVision = False
    End Function
    Public Function Abs(ByVal val As Long) As Long
        If val < 0 Then
            Abs = Not val + 1
        Else
            Abs = val
        End If
    End Function

    Function Distance(X1 As Object, Y1 As Object, X2 As Object, Y2 As Object) As Double

        'Encuentra la distancia entre dos puntos
        Distance = Math.Sqrt(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

    End Function

    Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero
        'Last Modify Date: 3/06/2006
        'Generates a random number in the range given - recoded to use longs and work properly with ranges
        '**************************************************************
        RandomNumber = Fix(Rnd() * (UpperBound - LowerBound + 1)) + LowerBound
    End Function

End Module
