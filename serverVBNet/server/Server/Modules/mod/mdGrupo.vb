﻿Option Explicit On


Module mdGrupo

    ''
    'cantidad maxima de parties en el servidor
    Public Const MAX_PARTIES As Integer = 300

    ''
    'nivel minimo para crear Grupo
    Public Const MINGrupoLEVEL As Byte = 15

    ''
    'Cantidad maxima de gente en el grupo
    Public Const Grupo_MAXMEMBERS As Byte = 5

    'maxima diferencia de niveles permitida en un grupo
    Public Const MAXGrupoDELTALEVEL As Byte = 5

    ''
    'distancia al leader para que este acepte el ingreso
    Public Const MAXDISTANCIAINGRESOGrupo As Byte = 4

    ''
    'maxima distancia a un exito para obtener su experiencia
    Public Const Grupo_MAXDISTANCIA As Byte = 18

    ''
    'restan las muertes de los miembros?
    Public Const CASTIGOS As Boolean = False

    ''
    'Numero al que elevamos el nivel de cada miembro del Grupo
    'Esto es usado para calcular la distribución de la experiencia entre los miembros
    'Se lee del archivo de balance
    Public ExponenteNivelGrupo As Single



    Public Function NextGrupo() As Integer
        Dim i As Integer
        NextGrupo = -1
        For i = 1 To MAX_PARTIES
            If Parties(i) Is Nothing Then
                NextGrupo = i
                Exit Function
            End If
        Next i
    End Function

    Public Sub CrearGrupo(ByVal UserIndex As Integer, ByVal tU As Integer)
        Dim tInt As Integer
        If UserList(UserIndex).GrupoIndex = 0 Then
            If UserList(UserIndex).flags.Muerto = 0 Then
                tInt = mdGrupo.NextGrupo
                If tInt = -1 Then
                    Call WriteConsoleMsg(1, UserIndex, "Por el momento no se pueden crear mas grupos", FontTypeNames.FONTTYPE_GRUPO)
                    Exit Sub
                Else
                    Parties(tInt) = New clsGrupo
                    If Not Parties(tInt).NuevoMiembro(UserIndex) Then
                        Call WriteConsoleMsg(1, UserIndex, "El grupo está lleno, no puedes entrar", FontTypeNames.FONTTYPE_GRUPO)
                        Parties(tInt) = Nothing
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(1, UserIndex, "¡Has creado un grupo!", FontTypeNames.FONTTYPE_GRUPO)
                        UserList(UserIndex).GrupoIndex = tInt
                        UserList(UserIndex).GrupoSolicitud = 0
                        If Not Parties(tInt).HacerLeader(UserIndex) Then
                            Call WriteConsoleMsg(1, UserIndex, "No puedes hacerte líder.", FontTypeNames.FONTTYPE_GRUPO)
                        Else
                            Call WriteConsoleMsg(1, UserIndex, "¡Te has convertido en líder del grupo!", FontTypeNames.FONTTYPE_GRUPO)
                        End If

                        WriteGrupo(UserIndex)

                        Parties(tInt).NuevoMiembro(tU)

                        UserList(tU).GrupoIndex = tInt
                        UserList(tU).GrupoSolicitud = 0
                        WriteGrupo(tU)

                        Call WriteConsoleMsg(1, tU, "¡El grupo ha sido creado!", FontTypeNames.FONTTYPE_GRUPO)
                    End If
                End If
            Else
                Call WriteMsg(UserIndex, 0)
            End If
        Else
            Call WriteConsoleMsg(1, UserIndex, " Ya perteneces a un grupo.", FontTypeNames.FONTTYPE_GRUPO)
        End If
    End Sub

    Public Sub SolicitarIngresoAGrupo(ByVal UserIndex As Integer)
        Dim tInt As Integer

        If UserList(UserIndex).GrupoIndex > 0 Then
            Call WriteConsoleMsg(1, UserIndex, " Ya perteneces a un grupo.", FontTypeNames.FONTTYPE_GRUPO)
            UserList(UserIndex).GrupoSolicitud = 0
            Exit Sub
        End If

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMsg(UserIndex, 0)
            UserList(UserIndex).GrupoSolicitud = 0
            Exit Sub
        End If

        tInt = UserList(UserIndex).flags.TargetUser
        If tInt > 0 Then
            If UserList(tInt).GrupoIndex > 0 Then
                UserList(UserIndex).GrupoSolicitud = UserList(tInt).GrupoIndex
            Else
                Call WriteConsoleMsg(1, UserIndex, UserList(tInt).Name & " no creo ningun grupo.", FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).GrupoSolicitud = 0
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(1, UserIndex, " Para ingresar a un grupo debes hacer click Grupo y luego click sobre el lider.", FontTypeNames.FONTTYPE_GRUPO)
            UserList(UserIndex).GrupoSolicitud = 0
        End If
    End Sub

    Public Sub SalirDeGrupo(ByVal UserIndex As Integer)
        Dim PI As Integer
        PI = UserList(UserIndex).GrupoIndex
        If PI > 0 Then
            If Parties(PI).SaleMiembro(UserIndex) Then
                'sale el leader
                Parties(PI) = New clsGrupo
            Else
                UserList(UserIndex).GrupoIndex = 0
                WriteGrupo(UserIndex)
            End If
        Else
            Call WriteConsoleMsg(1, UserIndex, " No eres miembro de ningun grupo.", FontTypeNames.FONTTYPE_INFO)
        End If
    End Sub

    Public Sub ExpulsarDeGrupo(ByVal leader As Integer, ByVal OldMember As Integer)
        Dim PI As Integer
        PI = UserList(leader).GrupoIndex

        If PI = UserList(OldMember).GrupoIndex Then
            If Parties(PI).SaleMiembro(OldMember) Then
                Parties(PI) = Nothing
            Else
                UserList(OldMember).GrupoIndex = 0
                WriteGrupo(OldMember)
            End If
        Else
            Call WriteConsoleMsg(1, leader, LCase(UserList(OldMember).Name) & " no pertenece a tu grupo.", FontTypeNames.FONTTYPE_INFO)
        End If
    End Sub

    ''
    ' Determines if a user can use Grupo commands like /acceptGrupo or not.
    '
    ' @param User Specifies reference to user
    ' @return  True if the user can use Grupo commands, false if not.
    Public Function UserPuedeEjecutarComandos(ByVal User As Integer) As Boolean
        '*************************************************
        'Author: Marco Vanotti(Marco)
        'Last modified: 05/05/09
        '
        '*************************************************
        Dim PI As Integer

        PI = UserList(User).GrupoIndex

        If PI > 0 Then
            If Parties(PI).EsGrupoLeader(User) Then
                UserPuedeEjecutarComandos = True
            Else
                Call WriteConsoleMsg(1, User, "¡No eres el líder de tu Grupo!", FontTypeNames.FONTTYPE_GRUPO)
                Exit Function
            End If
        Else
            Call WriteConsoleMsg(1, User, "No eres miembro de ningun grupo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End Function

    Public Sub AprobarIngresoAGrupo(ByVal leader As Integer, ByVal NewMember As Integer)
        'el UI es el leader
        Dim PI As Integer
        Dim razon As String

        PI = UserList(leader).GrupoIndex

        If UserList(NewMember).GrupoSolicitud = PI Then
            If Not UserList(NewMember).flags.Muerto = 1 Then
                If UserList(NewMember).GrupoIndex = 0 Then
                    If Parties(PI).PuedeEntrar(NewMember, razon) Then
                        If Parties(PI).NuevoMiembro(NewMember) Then
                            Call Parties(PI).MandarMensajeAConsola(UserList(leader).Name & " ha aceptado a " & UserList(NewMember).Name & " en el grupo.", "Servidor")
                            UserList(NewMember).GrupoIndex = PI
                            UserList(NewMember).GrupoSolicitud = 0
                            WriteGrupo(NewMember)
                        Else
                            'no pudo entrar
                            'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
                            Call SendData(SendTarget.ToADMINS, leader, PrepareMessageConsoleMsg(1, " Servidor> CATASTROFE EN GRUPOS, NUEVOMIEMBRO DIO FALSE! :S ", FontTypeNames.FONTTYPE_GRUPO))
                        End If
                    Else
                        'no debe entrar
                        Call WriteConsoleMsg(1, leader, razon, FontTypeNames.FONTTYPE_GRUPO)
                    End If
                Else
                    Call WriteConsoleMsg(1, leader, UserList(NewMember).Name & " ya es miembro de otro grupo.", FontTypeNames.FONTTYPE_GRUPO)
                    Exit Sub
                End If
            Else
                Call WriteConsoleMsg(1, leader, "¡Está muerto, no puedes aceptar miembros en ese estado!", FontTypeNames.FONTTYPE_GRUPO)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(1, leader, LCase(UserList(NewMember).Name) & " no ha solicitado ingresar a tu grupo.", FontTypeNames.FONTTYPE_GRUPO)
            Exit Sub
        End If

    End Sub
    Public Sub ApruebaSolicitud(ByVal leader As Integer, ByVal NewMember As Integer)
        Dim PI As Integer
        Dim razon As String

        PI = UserList(leader).GrupoIndex
        If EsLider(leader) Then
            If UserList(leader).flags.Solicito = NewMember Then
                If Not (UserList(NewMember).flags.Muerto) Then
                    If UserList(NewMember).GrupoIndex = 0 Then
                        If Parties(PI).PuedeEntrar(NewMember, razon) Then
                            If Parties(PI).NuevoMiembro(NewMember) Then
                                UserList(NewMember).GrupoIndex = PI
                                UserList(NewMember).GrupoSolicitud = 0
                                WriteGrupo(NewMember)
                            Else
                                'no pudo entrar
                                'ACA UNO PUEDE CODIFICAR OTRO TIPO DE ERRORES...
                                Call SendData(SendTarget.ToADMINS, leader, PrepareMessageConsoleMsg(1, " Servidor> CATASTROFE EN GRUPOS, NUEVOMIEMBRO DIO FALSE! :S ", FontTypeNames.FONTTYPE_GRUPO))
                            End If
                        Else
                            'no debe entrar
                            Call WriteConsoleMsg(1, leader, razon, FontTypeNames.FONTTYPE_GRUPO)
                        End If
                    Else
                        Call WriteConsoleMsg(1, leader, UserList(NewMember).Name & " ya es miembro de otro grupo.", FontTypeNames.FONTTYPE_GRUPO)
                        Exit Sub
                    End If
                Else
                    Call WriteConsoleMsg(1, leader, "¡Está muerto, no puedes aceptar miembros en ese estado!", FontTypeNames.FONTTYPE_GRUPO)
                    Exit Sub
                End If
            Else
                Call WriteConsoleMsg(1, leader, LCase(UserList(NewMember).Name) & " no ha sido solicitado.", FontTypeNames.FONTTYPE_GRUPO)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(1, leader, "No sos lider del grupo.", FontTypeNames.FONTTYPE_GRUPO)
        End If
    End Sub
    Public Sub BroadCastGrupo(ByVal UserIndex As Integer, ByRef texto As String)
        Dim PI As Integer

        PI = UserList(UserIndex).GrupoIndex

        If PI > 0 Then
            Call Parties(PI).MandarMensajeAConsola(texto, UserList(UserIndex).Name)
        End If

    End Sub

    Public Sub OnlineGrupo(ByVal UserIndex As Integer)
        Dim PI As Integer
        Dim texto As String

        PI = UserList(UserIndex).GrupoIndex

        If PI > 0 Then
            Call Parties(PI).ObtenerMiembrosOnline(texto)
            Call WriteConsoleMsg(1, UserIndex, texto, FontTypeNames.FONTTYPE_GRUPO)
        End If


    End Sub
    Public Function NombreMiembro(ByVal UserIndex As Integer, ByVal i As Byte) As String
        NombreMiembro = Parties(UserList(UserIndex).GrupoIndex).MiembroLista(i)
    End Function


    Public Sub TransformarEnLider(ByVal OldLeader As Integer, ByVal NewLeader As Integer)
        Dim PI As Integer

        If OldLeader = NewLeader Then Exit Sub

        PI = UserList(OldLeader).GrupoIndex

        If PI = UserList(NewLeader).GrupoIndex Then
            If UserList(NewLeader).flags.Muerto = 0 Then
                If Parties(PI).HacerLeader(NewLeader) Then
                    Call Parties(PI).MandarMensajeAConsola("El nuevo líder del Grupo es " & UserList(NewLeader).Name, UserList(OldLeader).Name)
                Else
                    Call WriteConsoleMsg(1, OldLeader, "¡No se ha hecho el cambio de mando!", FontTypeNames.FONTTYPE_GRUPO)
                End If
            Else
                Call WriteConsoleMsg(1, OldLeader, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(1, OldLeader, LCase(UserList(NewLeader).Name) & " no pertenece a tu Grupo.", FontTypeNames.FONTTYPE_INFO)
        End If

    End Sub

    Public Function EsLider(ByVal UserIndex As Integer) As Byte
        If UserList(UserIndex).GrupoIndex <> 0 Then
            If Parties(UserList(UserIndex).GrupoIndex).EsGrupoLeader(UserIndex) Then
                EsLider = 1
            Else
                EsLider = 0
            End If
        Else
            EsLider = 0
        End If
    End Function


    Public Sub ObtenerExito(ByVal UserIndex As Integer, ByVal Exp As Long, mapa As Integer, x As Integer, Y As Integer)
        If Exp <= 0 Then
            If Not CASTIGOS Then Exit Sub
        End If

        Call Parties(UserList(UserIndex).GrupoIndex).ObtenerExito(Exp, mapa, x, Y)


    End Sub

    Public Function CantMiembros(ByVal UserIndex As Integer) As Integer
        CantMiembros = 0
        If UserList(UserIndex).GrupoIndex > 0 Then
            CantMiembros = Parties(UserList(UserIndex).GrupoIndex).CantMiembros
        End If

    End Function
    Public Sub LimpiarFlags(ByVal UI As Integer)
        UserList(UI).GrupoIndex = 0
        UserList(UI).GrupoSolicitud = 0

        UserList(UI).flags.Solicito = 0
    End Sub



End Module
