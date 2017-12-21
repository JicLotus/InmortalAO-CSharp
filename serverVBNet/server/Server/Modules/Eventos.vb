Option Explicit On

Module Eventos
    'Add Nod kopfnickend Carrera
    Public Carrera_estado As Boolean
    Public Carrera_puestos As Byte
    Public Const MapaCarrera As Integer = 237

    Sub Carrera_Entra(ByVal UserIndex As Integer)
        On Error Resume Next
        Dim i As Integer

        If (Not Carrera_estado) Then
            Call WriteConsoleMsg(1, UserIndex, "No hay ninguna Carrera!", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        ElseIf UserList(UserIndex).Pos.map = MapaCarrera Then
            Call WriteConsoleMsg(1, UserIndex, "Ya estas adentro de la Carrera!", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        ElseIf Carrera_puestos = 255 Then
            Call WriteConsoleMsg(1, UserIndex, "La Carrera ya empezó!", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        ElseIf UserList(UserIndex).Stats.GLD < 100000 Then ' Si no tiene plata el pobre se queda afuera
            Call WriteConsoleMsg(1, UserIndex, "No tiene es dinero suficiente para entrar.", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If

        If Carrera_puestos < 21 Then
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos

            Carrera_puestos = Carrera_puestos + 1

            'Le sacamos la plata
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 100000
            Call WriteUpdateGold(UserIndex)

            FuturePos.map = MapaCarrera
            FuturePos.x = 47
            If Carrera_puestos Mod 2 Then
                FuturePos.Y = 70
            Else
                FuturePos.Y = 39
            End If

            Call ClosestLegalPos(FuturePos, NuevaPos)

            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then _
            Call WarpUserChar(UserIndex, NuevaPos.map, NuevaPos.x, NuevaPos.Y, True)

            Call WriteConsoleMsg(1, UserIndex, "Estas dentro de la Carrera!", FontTypeNames.FONTTYPE_TALK)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No hay mas lugar en la Carrera", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If
    End Sub
    '\Add

    Public Function mapasEspeciales(UserIndex As Integer) As Boolean
        mapasEspeciales = UserList(UserIndex).Pos.map = 206 Or
               UserList(UserIndex).Pos.map = 248 Or
               UserList(UserIndex).Pos.map = 249 Or
               UserList(UserIndex).Pos.map = 237 Or
               UserList(UserIndex).Pos.map = 238 Or
               UserList(UserIndex).Pos.map = 246 Or
               UserList(UserIndex).Pos.map = 247 Or
               UserList(UserIndex).Pos.map = 251 Or
               UserList(UserIndex).Pos.map = 290 Or
               UserList(UserIndex).Pos.map = 750 Or
               UserList(UserIndex).Pos.map = 751 Or
               UserList(UserIndex).Pos.map = 752 Or
               UserList(UserIndex).Pos.map = 848 Or
               UserList(UserIndex).Pos.map = 845 Or
               UserList(UserIndex).Pos.map = MapaCarrera
    End Function

End Module
