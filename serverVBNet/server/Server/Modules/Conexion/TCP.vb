Option Explicit On

Imports System.Text

Module TCP

    'Private Declare Function VarPtr Lib "msvbvm60.dll" Alias "VarPtr" (ByVal lpObject As String) As Integer

    Private Declare Function VarPtr Lib "vb40032.dll" Alias "VarPtr" (lpObject As Object) As Long


    Enum lStat
        Incinerado = &H1
        Envenenado = &H2
        Comerciand = &H4
        Trabajando = &H8
        Transformado = &H10
        Ciego = &H20
        Inactivo = &H40
        Silenciado = &H80
    End Enum

    Enum lStatEx
        Paralizado = &H1
        Inmovilizado = &H2
        Hombre = &H4
        Mujer = &H8
    End Enum
    Sub DarCuerpo(ByVal UserIndex As Integer)
        Dim NewBody As Integer
        Dim UserRaza As Byte
        Dim UserGenero As Byte
        UserGenero = UserList(UserIndex).Genero
        UserRaza = UserList(UserIndex).Raza
        Select Case UserGenero
            Case eGenero.Hombre
                Select Case UserRaza
                    Case eRaza.Humano
                        NewBody = 1
                    Case eRaza.Elfo
                        NewBody = 2
                    Case eRaza.Drow
                        NewBody = 3
                    Case eRaza.Enano
                        NewBody = 95
                    Case eRaza.Gnomo
                        NewBody = 52
                    Case eRaza.Orco
                        NewBody = 251
                End Select
            Case eGenero.Mujer
                Select Case UserRaza
                    Case eRaza.Humano
                        NewBody = 351
                    Case eRaza.Elfo
                        NewBody = 352
                    Case eRaza.Drow
                        NewBody = 353
                    Case eRaza.Gnomo
                        NewBody = 138
                    Case eRaza.Enano
                        NewBody = 138
                    Case eRaza.Orco
                        NewBody = 252
                End Select
        End Select
        UserList(UserIndex).cuerpo.body = NewBody
    End Sub

    Function AsciiValidos(ByVal cad As String) As Boolean
        Dim car As Byte
        Dim i As Integer

        cad = LCase$(cad)

        For i = 1 To Len(cad)
            car = Asc(Mid(cad, i, 1))

            If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
                AsciiValidos = False
                Exit Function
            End If

        Next i

        AsciiValidos = True

    End Function

    Function Numeric(ByVal cad As String) As Boolean
        Dim car As Byte
        Dim i As Integer

        cad = LCase$(cad)

        For i = 1 To Len(cad)
            car = Asc(Mid(cad, i, 1))

            If (car < 48 Or car > 57) Then
                Numeric = False
                Exit Function
            End If

        Next i

        Numeric = True

    End Function




    Function ValidateSkills(ByVal UserIndex As Integer) As Boolean

        Dim loopC As Integer

        For loopC = 1 To NUMSKILLS
            If UserList(UserIndex).Stats.UserSkills(loopC) < 0 Then UserList(UserIndex).Stats.UserSkills(loopC) = 0
            If UserList(UserIndex).Stats.UserSkills(loopC) > 100 Then UserList(UserIndex).Stats.UserSkills(loopC) = 100
        Next loopC

        ValidateSkills = True

    End Function

    Sub ConnectNewUser(ByVal UserIndex As Integer, ByRef Name As String, ByRef account As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass,
                    ByRef skills() As Byte, ByRef UserEmail As String, ByVal Hogar As eCiudad,
                    ByVal Fuerza As Byte, ByVal Agilidad As Byte, ByVal Inteligencia As Byte,
                    ByVal Carisma As Byte, ByVal constitucion As Byte, ByVal Cabeza As Integer,
                    ByVal petTipe As eMascota, ByRef petName As String)


        Name = Replace(Name, "  ", " ")

        If Not AsciiValidos(Name) Then
            Call WriteErrorMsg(UserIndex, "Nombre invalido.")
            Exit Sub
        End If

        If Len(Name) < 2 Then
            Call WriteErrorMsg(UserIndex, "El nombre es muy corto.")
            Exit Sub
        End If

        If Len(Name) > 20 Then
            Call WriteErrorMsg(UserIndex, "El nombre es muy largo.")
            Exit Sub
        End If

        If UserList(UserIndex).flags.UserLogged Then
            Call LogCheating("El usuario " & UserList(UserIndex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(UserIndex).ip)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If

        Dim loopC As Long
        Dim totalskpts As Long

        '¿Existe el personaje?
        If ExistePersonaje(Name) = True Then
            Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
            Exit Sub
        End If


        UserList(UserIndex).flags.Muerto = 0
        UserList(UserIndex).flags.Escondido = 0

        UserList(UserIndex).Name = Name
        UserList(UserIndex).Clase = UserClase
        UserList(UserIndex).Raza = UserRaza
        UserList(UserIndex).Genero = UserSexo
        UserList(UserIndex).email = UserEmail
        UserList(UserIndex).Hogar = Hogar


        ReDim UserList(UserIndex).Stats.UserAtributos(NUMATRIBUTOS + 1)

        UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = Fuerza + ModRazaArr(UserRaza).Fuerza
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = Agilidad + ModRazaArr(UserRaza).Agilidad
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = IIf(Inteligencia + ModRazaArr(UserRaza).Inteligencia < 0, 0, Inteligencia + ModRazaArr(UserRaza).Inteligencia)
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = Carisma + ModRazaArr(UserRaza).Carisma
        UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion) = constitucion + ModRazaArr(UserRaza).constitucion

        ReDim UserList(UserIndex).Stats.UserAtributosBackUP(NUMATRIBUTOS + 1)

        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Fuerza) = Fuerza + ModRazaArr(UserRaza).Fuerza
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Agilidad) = Agilidad + ModRazaArr(UserRaza).Agilidad
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Inteligencia) = IIf(Inteligencia + ModRazaArr(UserRaza).Inteligencia < 0, 0, Inteligencia + ModRazaArr(UserRaza).Inteligencia)
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.Carisma) = Carisma + ModRazaArr(UserRaza).Carisma
        UserList(UserIndex).Stats.UserAtributosBackUP(eAtributos.constitucion) = constitucion + ModRazaArr(UserRaza).constitucion

        If (Fuerza + Agilidad + Inteligencia + Carisma + constitucion) > 70 Or
       (Fuerza < 6 Or Agilidad < 6 Or Inteligencia < 6 Or Carisma < 6 Or constitucion < 6) Or
       (Fuerza > 18 Or Agilidad > 18 Or Inteligencia > 18 Or Carisma > 18 Or constitucion > 18) Then

            Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los atributos.")
            'Call BorrarUsuario(UserList(UserIndex).name)
            Call WriteErrorMsg(UserIndex, "Por favor vaya a molestar a otro servidor.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If

        ReDim UserList(UserIndex).Stats.UserSkills(NUMSKILLS + 1)

        For loopC = 1 To NUMSKILLS
            If skills(loopC - 1) >= 0 Then
                UserList(UserIndex).Stats.UserSkills(loopC) = skills(loopC - 1)
                totalskpts = totalskpts + Math.Abs(UserList(UserIndex).Stats.UserSkills(loopC))
            Else
                Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los skills.")
                'Call BorrarUsuario(UserList(UserIndex).name)
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
        Next loopC

        If totalskpts > 10 Then
            Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los skills.")
            'Call BorrarUsuario(UserList(UserIndex).name)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

        UserList(UserIndex).cuerpo.heading = eHeading.SOUTH

        Call DarCuerpo(UserIndex)
        UserList(UserIndex).cuerpo.Head = Cabeza
        UserList(UserIndex).OrigChar = UserList(UserIndex).cuerpo

        If UserClase = eClass.Mago Or UserClase = eClass.Druida Or UserClase = eClass.Cazador Then

            If Len(petName) > 30 Then
                Call WriteErrorMsg(UserIndex, "El nombre de la mascota no debe sobrepasar 30 letras.")
                Call FlushBuffer(UserIndex)
                Exit Sub
            ElseIf Len(petName) = 0 Then
                Call WriteErrorMsg(UserIndex, "Nombre de familiar o mascota invalido.")
                Call FlushBuffer(UserIndex)
                Exit Sub
            End If

            Call EntregarMascota(UserIndex, petTipe, petName)

        Else
            UserList(UserIndex).masc.TieneFamiliar = 0
            UserList(UserIndex).masc.tipo = 0
            UserList(UserIndex).masc.Nombre = ""
        End If

        UserList(UserIndex).cuerpo.WeaponAnim = NingunArma
        UserList(UserIndex).cuerpo.ShieldAnim = NingunEscudo
        UserList(UserIndex).cuerpo.CascoAnim = NingunCasco

        Dim MiInt As Long
        MiInt = UserList(UserIndex).Stats.UserAtributos(eAtributos.constitucion)

        UserList(UserIndex).Stats.MaxHP = MiInt
        UserList(UserIndex).Stats.MinHP = MiInt

        MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
        If MiInt = 1 Then MiInt = 2

        UserList(UserIndex).Stats.MaxSTA = 20 * MiInt
        UserList(UserIndex).Stats.MinSTA = 20 * MiInt

        UserList(UserIndex).Stats.MaxAGU = 100
        UserList(UserIndex).Stats.MinAGU = 100

        UserList(UserIndex).Stats.MaxHAM = 100
        UserList(UserIndex).Stats.MinHAM = 100


        '<-----------------MANA----------------------->
        If UserClase = eClass.Mago Then  'Cambio en mana inicial (ToxicWaste)
            MiInt = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * 3
            UserList(UserIndex).Stats.MaxMAN = MiInt
            UserList(UserIndex).Stats.MinMAN = MiInt
        ElseIf UserClase = eClass.Clerigo Or UserClase = eClass.Druida _
        Or UserClase = eClass.Bardo Or UserClase = eClass.Asesino _
        Or UserClase = eClass.Nigromante Then
            UserList(UserIndex).Stats.MaxMAN = 50
            UserList(UserIndex).Stats.MinMAN = 50
        Else
            UserList(UserIndex).Stats.MaxMAN = 0
            UserList(UserIndex).Stats.MinMAN = 0
        End If


        ReDim UserList(UserIndex).Stats.UserHechizos(MAXUSERHECHIZOS + 1)


        If UserClase = eClass.Mago Or UserClase = eClass.Clerigo Or
       UserClase = eClass.Druida Or UserClase = eClass.Bardo Or
       UserClase = eClass.Asesino Or UserClase = eClass.Nigromante Or
       UserClase = eClass.Paladin Then
            UserList(UserIndex).Stats.UserHechizos(1) = 2
        End If

        If UserClase = eClass.Mago Or
       UserClase = eClass.Druida Then
            UserList(UserIndex).Stats.UserHechizos(2) = 59
        End If

        If UserClase = eClass.Cazador Then
            UserList(UserIndex).Stats.UserHechizos(1) = 59
        End If


        UserList(UserIndex).flags.miPareja = ""


        UserList(UserIndex).Stats.MaxHit = 2
        UserList(UserIndex).Stats.MinHit = 1

        UserList(UserIndex).Stats.GLD = 10

        UserList(UserIndex).Stats.Exp = 0
        UserList(UserIndex).Stats.ELU = 400
        UserList(UserIndex).Stats.ELV = 1


        ReDim UserList(UserIndex).Invent.Objeto(MAX_INVENTORY_SLOTS + 1)
        '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
        UserList(UserIndex).Invent.NroItems = 4

        UserList(UserIndex).Invent.Objeto(1).ObjIndex = 573
        UserList(UserIndex).Invent.Objeto(1).Amount = 100

        UserList(UserIndex).Invent.Objeto(2).ObjIndex = 572
        UserList(UserIndex).Invent.Objeto(2).Amount = 100

        'Esto depende de la clase
        If UserList(UserIndex).Clase = eClass.Cazador Or
        UserList(UserIndex).Clase = eClass.Druida Then

            UserList(UserIndex).Invent.Objeto(3).ObjIndex = 1355
            UserList(UserIndex).Invent.Objeto(3).Amount = 1
            UserList(UserIndex).Invent.Objeto(3).Equipped = 1
        ElseIf UserList(UserIndex).Clase = eClass.Paladin Or
        UserList(UserIndex).Clase = eClass.Guerrero Or
        UserList(UserIndex).Clase = eClass.Herrero Or
        UserList(UserIndex).Clase = eClass.Mercenario Or
        UserList(UserIndex).Clase = eClass.Minero Or
        UserList(UserIndex).Clase = eClass.Clerigo Or
        UserList(UserIndex).Clase = eClass.Leñador Then

            UserList(UserIndex).Invent.Objeto(3).ObjIndex = 574
            UserList(UserIndex).Invent.Objeto(3).Amount = 1
            UserList(UserIndex).Invent.Objeto(3).Equipped = 1
        ElseIf UserList(UserIndex).Clase = eClass.Gladiador Or
        UserList(UserIndex).Clase = eClass.Bardo Then

            UserList(UserIndex).Invent.Objeto(3).ObjIndex = 1354
            UserList(UserIndex).Invent.Objeto(3).Amount = 1
            UserList(UserIndex).Invent.Objeto(3).Equipped = 1
        ElseIf UserList(UserIndex).Clase = eClass.Asesino Or
        UserList(UserIndex).Clase = eClass.Ladron Or
        UserList(UserIndex).Clase = eClass.Sastre Or
        UserList(UserIndex).Clase = eClass.Pescador Then

            UserList(UserIndex).Invent.Objeto(3).ObjIndex = 460
            UserList(UserIndex).Invent.Objeto(3).Amount = 1
            UserList(UserIndex).Invent.Objeto(3).Equipped = 1
        ElseIf UserList(UserIndex).Clase = eClass.Mago Or
        UserList(UserIndex).Clase = eClass.Nigromante Then

            UserList(UserIndex).Invent.Objeto(3).ObjIndex = 1356
            UserList(UserIndex).Invent.Objeto(3).Amount = 1
            UserList(UserIndex).Invent.Objeto(3).Equipped = 1
        End If

        Select Case UserRaza
            Case eRaza.Humano
                If UserList(UserIndex).Genero = eGenero.Hombre Then
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 463
                Else
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 1283
                End If
            Case eRaza.Elfo
                If UserList(UserIndex).Genero = eGenero.Hombre Then
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 464
                Else
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 1284
                End If
            Case eRaza.Drow
                If UserList(UserIndex).Genero = eGenero.Hombre Then
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 465
                Else
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 1285
                End If
            Case eRaza.Enano
                If UserList(UserIndex).Genero = eGenero.Hombre Then
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 562
                Else
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 563
                End If
            Case eRaza.Gnomo
                If UserList(UserIndex).Genero = eGenero.Hombre Then
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 466
                Else
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 563
                End If
            Case eRaza.Orco
                If UserList(UserIndex).Genero = eGenero.Hombre Then
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 988
                Else
                    UserList(UserIndex).Invent.Objeto(4).ObjIndex = 1087
                End If
        End Select

        UserList(UserIndex).Invent.Objeto(4).Amount = 1
        UserList(UserIndex).Invent.Objeto(4).Equipped = 1

        UserList(UserIndex).Invent.Objeto(5).ObjIndex = 461
        UserList(UserIndex).Invent.Objeto(5).Amount = 100

        If UserList(UserIndex).Clase = eClass.Cazador Or
        UserList(UserIndex).Clase = eClass.Druida Then

            UserList(UserIndex).Invent.Objeto(6).ObjIndex = 1357
            UserList(UserIndex).Invent.Objeto(6).Amount = 100
        ElseIf UserList(UserIndex).Clase = eClass.Asesino Or
        UserList(UserIndex).Clase = eClass.Ladron Then

            UserList(UserIndex).Invent.Objeto(6).ObjIndex = 576
            UserList(UserIndex).Invent.Objeto(6).Amount = 100
        End If

        UserList(UserIndex).Invent.Objeto(7).ObjIndex = 1601
        UserList(UserIndex).Invent.Objeto(7).Amount = 1

        Dim tmpObj As obj
        tmpObj.ObjIndex = 879 : tmpObj.Amount = 1
        Call MeterItemEnInventario(UserIndex, tmpObj) 'Mapa

        UserList(UserIndex).Invent.ArmourEqpSlot = 4
        UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Objeto(4).ObjIndex

        UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Objeto(3).ObjIndex
        UserList(UserIndex).Invent.WeaponEqpSlot = 3

        'Valores Default de facciones al Activar nuevo usuario
        Call ResetFacciones(UserIndex)

        If UserList(UserIndex).Hogar = 0 Then
            UserList(UserIndex).faccion.Ciudadano = 1
        ElseIf UserList(UserIndex).Hogar = 1 Then
            UserList(UserIndex).faccion.Republicano = 1
        End If

        Call SaveUserSQL(UserIndex, account, True)

        'Open User
        Call ConnectUser(UserIndex, Name, account)

    End Sub

    Sub CloseSocket(ByVal UserIndex As Integer)

        On Error GoTo Errhandler

        If UserIndex = LastUser Then
            Do Until UserList(LastUser).flags.UserLogged
                LastUser = LastUser - 1
                If LastUser < 1 Then Exit Do
            Loop
        End If

        If UserList(UserIndex).ComUsu.DestUsu > 0 Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                    ''Call WriteConsoleMsg(1, UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    Call FlushBuffer(UserList(UserIndex).ComUsu.DestUsu)
                End If
            End If
        End If

        If UserList(UserIndex).flags.UserLogged = True Then
            Call CloseUser(UserIndex)
        End If

        Call ResetUserSlot(UserIndex)

        Exit Sub

Errhandler:
        Call ResetUserSlot(UserIndex)
        Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.Description & " - UserIndex = " & UserIndex)
    End Sub


    Public Function EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String) As Long

        Dim sendingBytes As Byte()
        sendingBytes = Encoding.Default.GetBytes(Datos)

        UserList(UserIndex).client.sendInmediatePacket(sendingBytes)

        EnviarDatosASlot = 0
    End Function



    Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


        Dim x As Integer, Y As Integer
        For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
            For x = UserList(Index).Pos.x - MinXBorder + 1 To UserList(Index).Pos.x + MinXBorder - 1

                If MapData(UserList(Index).Pos.map, x, Y).UserIndex = Index2 Then
                    EstaPCarea = True
                    Exit Function
                End If

            Next x
        Next Y
        EstaPCarea = False
    End Function

    Function HayPCarea(Pos As WorldPos) As Boolean


        Dim x As Integer, Y As Integer
        For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
            For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1
                If x > 0 And Y > 0 And x < 101 And Y < 101 Then
                    If MapData(Pos.map, x, Y).UserIndex > 0 Then
                        HayPCarea = True
                        Exit Function
                    End If
                End If
            Next x
        Next Y
        HayPCarea = False
    End Function

    Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


        Dim x As Integer, Y As Integer
        For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
            For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1
                If MapData(Pos.map, x, Y).ObjInfo.ObjIndex = ObjIndex Then
                    HayOBJarea = True
                    Exit Function
                End If

            Next x
        Next Y
        HayOBJarea = False
    End Function
    Function ValidateChr(ByVal UserIndex As Integer) As Boolean

        'Add Marius es una cabeza de enano bugeada y genera que no se vea el nombre de pj
        If UserList(UserIndex).OrigChar.Head = 72 Then UserList(UserIndex).OrigChar.Head = 1
        If UserList(UserIndex).OrigChar.Head = 314 Then UserList(UserIndex).OrigChar.Head = 1
        If UserList(UserIndex).OrigChar.Head = 315 Then UserList(UserIndex).OrigChar.Head = 1
        If UserList(UserIndex).OrigChar.Head = 121 Then UserList(UserIndex).OrigChar.Head = 1
        '\Add

        '
        '
        '
        '
        '

        ValidateChr = UserList(UserIndex).cuerpo.Head <> 0 _
                And UserList(UserIndex).cuerpo.body <> 0 _
                And ValidateSkills(UserIndex)

    End Function

    Sub ConnectUser(ByVal UserIndex As Integer, ByRef Name As String, ByRef account As String)


        Try

            Dim N As Integer
            Dim tStr As String
            Dim i As Integer

            With UserList(UserIndex)


                'Reseteamos los FLAGS
                .flags.Escondido = 0
                .flags.TargetNPC = 0
                .flags.TargetNpcTipo = eNPCType.Comun
                .flags.TargetObj = 0
                .flags.TargetUser = 0
                .cuerpo.fx = 0


                '¿Existe el personaje?
                If ExistePersonaje(Name) = False Then
                    Call WriteErrorMsg(UserIndex, "El personaje no existe.")
                    Call FlushBuffer(UserIndex)
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                '¿Ya esta conectado el personaje?
                If CheckForSameName(Name) Then
                    If UserList(NameIndex(Name)).Counters.Saliendo Then
                        Call WriteErrorMsg(UserIndex, "El usuario está saliendo.")
                    Else
                        Call WriteErrorMsg(UserIndex, "Usuario Conectado.")
                    End If
                    Call FlushBuffer(UserIndex)
                    Exit Sub
                End If

                'Reseteamos los privilegios
                .flags.Privilegios = 0
                .GuildIndex = 0

                If Not LoadUserSQL(UserIndex, Name) Then
                    Call WriteErrorMsg(UserIndex, "Error al cargar el personaje.")
                    Call FlushBuffer(UserIndex)
                    Call CloseSocket(UserIndex)
                    Call LogError("Error en LoadUserSQL: Error al cargar personaje: " & Name)
                    Exit Sub
                End If

                'Add Marius Lideres Faccionarios

                If EsFaccCaos(UserIndex) Then
                    '
                    Call ResetFacciones(UserIndex, False)
                    UserList(UserIndex).faccion.FuerzasCaos = 1
                    UserList(UserIndex).faccion.Rango = 200
                ElseIf EsFaccRepu(UserIndex) Then
                    '
                    Call ResetFacciones(UserIndex, False)
                    UserList(UserIndex).faccion.Republicano = 1
                    UserList(UserIndex).faccion.Milicia = 1
                    UserList(UserIndex).faccion.Rango = 200
                ElseIf EsFaccImpe(UserIndex) Then
                    '
                    Call ResetFacciones(UserIndex, False)
                    UserList(UserIndex).faccion.Ciudadano = 1
                    UserList(UserIndex).faccion.ArmadaReal = 1
                    UserList(UserIndex).faccion.Rango = 200
                ElseIf EsCONSE(UserIndex) Then
                    '
                    Call ResetFacciones(UserIndex, False)
                    UserList(UserIndex).faccion.Renegado = 1
                End If

                If Comprobar_Si_Donador(account) > 0 Or EsFacc(UserIndex) Then
                    UserList(UserIndex).donador = True
                Else
                    UserList(UserIndex).donador = False
                End If


                UserList(UserIndex).bandera = 0

                If Not ValidateChr(UserIndex) Then
                    Call WriteErrorMsg(UserIndex, "Error en el personaje.")
                    Call FlushBuffer(UserIndex)
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If

                If .Counters.IdleCount > 0 Then
                    .Counters.IdleCount = 0
                End If

                If .Invent.EscudoEqpSlot = 0 Then .cuerpo.ShieldAnim = NingunEscudo
                If .Invent.CascoEqpSlot = 0 Then .cuerpo.CascoAnim = NingunCasco
                If .Invent.WeaponEqpSlot = 0 And .Invent.NudiEqpSlot = 0 Then .cuerpo.WeaponAnim = NingunArma

                Call UpdateUserInv(True, UserIndex, 0)
                Call UpdateUserHechizos(True, UserIndex, 0)

                If .flags.Paralizado Then
                    Call WriteParalizeOK(UserIndex, False)
                End If


                If .flags.Estupidez = 0 Then
                    Call WriteDumbNoMore(UserIndex)
                End If

                'Posicion de comienzo
                If .Pos.map = 0 Then
                    Select Case .Hogar
                        Case 0 ' Nix
                            .Pos.x = 40
                            .Pos.Y = 87
                            .Pos.map = 34
                        Case 1 ' Illindor
                            .Pos.x = 61
                            .Pos.Y = 71
                            .Pos.map = 241
                    End Select
                Else
                    If Not MapaValido(.Pos.map) Then
                        .Pos.map = 49
                        .Pos.x = 50
                        .Pos.Y = 50
                    End If
                End If


                If EsFacc(UserIndex) Then
                    .Pos.x = 50
                    .Pos.Y = 54
                    .Pos.map = 248

                    Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, "Server> " & Name & " Se ha conectado!", FontTypeNames.FONTTYPE_VENENO))
                End If




                .Name = Name
                If EsFacc(UserIndex) Then
                    Call LogGM(UserList(UserIndex).Name, "----------------------------------------------------------")
                    Call LogGM(UserList(UserIndex).Name, "Entró al juego")
                End If

                .showName = True 'Por default los nombres son visibles

                'If in the water, and has a boat, equip it!
                'Fix Marius Ya no mas caminatas sobre el agua!
                If .Invent.BarcoObjIndex > 0 And (HayAgua(.Pos.map, .Pos.x, .Pos.Y) Or .cuerpo.body = 84 Or .cuerpo.body = iFragataFantasmal) Then
                    Dim Barco As ObjData
                    'Barco = ObjDataArr(.Invent.BarcoObjIndex)
                    .cuerpo.Head = 0
                    If .flags.Muerto <> 0 Then
                        .cuerpo.body = iFragataFantasmal
                    Else
                        .cuerpo.body = 84
                    End If
                    .flags.Navegando = 1
                Else
                    .flags.Navegando = 0
                End If


                .Pos = getArroundValidPosition(UserIndex)


                'Info
                Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
                'Carga el mapa
                Call WriteChangeMap(UserIndex, .Pos.map, MapInfoArr(.Pos.map).MapVersion)
                Call WritePlayMidi(UserIndex, Val(ReadField(1, MapInfoArr(.Pos.map).Music, 45)))



                UserList(UserIndex).Counters.IdleCount = 0

                Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.x, .Pos.Y)

                Call WriteUserCharIndexInServer(UserIndex)


                Call DoTileEvents(UserIndex, .Pos.map, .Pos.x, .Pos.Y)

                Call WriteUpdateUserStats(UserIndex)

                Call WriteUpdateHungerAndThirst(UserIndex)


                Call SendMOTD(UserIndex)


                If haciendoBK Then
                    Call WritePauseToggle(UserIndex)
                    Call WriteConsoleMsg(1, UserIndex, "Servidor> Por favor espera algunos segundos, WorldSave esta ejecutandose.", FontTypeNames.FONTTYPE_SERVER)
                End If

                If EnPausa Then
                    Call WritePauseToggle(UserIndex)
                    Call WriteConsoleMsg(1, UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
                End If

                .flags.UserLogged = True

                MapInfoArr(.Pos.map).NumUsers = MapInfoArr(.Pos.map).NumUsers + 1

                If .Stats.SkillPts > 0 Then
                    Call WriteSendSkills(UserIndex)
                End If


                ReDim .MascotasIndex(0 To MAXMASCOTAS + 1)
                ReDim .MascotasType(0 To MAXMASCOTAS + 1)

                If .NroMascotas > 0 And MapInfoArr(.Pos.map).Pk Then

                    For i = 1 To MAXMASCOTAS
                        If .MascotasType(i) > 0 Then
                            .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)

                            If .MascotasIndex(i) > 0 Then
                                Npclist(.MascotasIndex(i)).MaestroUser = UserIndex
                                Call FollowAmo(.MascotasIndex(i))
                            Else
                                .MascotasIndex(i) = 0
                            End If
                        End If
                    Next i
                End If



                If .flags.Navegando = 1 Then
                    Call WriteNavigateToggle(UserIndex)
                End If

                If esRene(UserIndex) Then
                    Call WriteSafeModeOff(UserIndex)
                    .flags.Seguro = False
                Else
                    .flags.Seguro = True
                    Call WriteSafeModeOn(UserIndex)
                End If

                If .GuildIndex > 0 Then

                    .GuildIndex = GuildIndexArray(.GuildIndex)

                    If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
                        Call WriteConsoleMsg(1, UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
                    End If

                End If


                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.cuerpo.CharIndex, FXIDs.FXWARP, 0))

                Call WriteLoggedMessage(UserIndex)

                tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)

                If Len(tStr) <> 0 Then
                    Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
                End If


                Call WriteFuerza(UserIndex)
                Call WriteAgilidad(UserIndex)
                WriteMensajeSigno(UserIndex)

                WriteHora(UserIndex)


                NumUsers = NumUsers + 1

                If NumUsers > RecordUsuarios Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, "Record de usuarios conectados simultaneamente. Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_BROWNI))
                    RecordUsuarios = NumUsers

                    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", Str(RecordUsuarios))
                    Call extra_set("Record", Str(RecordUsuarios))

                End If

                Call onpj(UserIndex)

                SendOnline()

                FlushBuffer(UserIndex)

            End With

            Exit Sub

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            LogError("Error en connectUser: " + ex.Message + " StackTrace: " + st.ToString())
        End Try


    End Sub


    'ReAdd Marius
    Sub SendMOTD(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim j As Long

        'Call WriteGuildChat(UserIndex, "Mensajes de entrada:")
        For j = 1 To MotdMaxLines
            'Call WriteGuildChat(UserIndex, MOTD(j).texto)
            Call WriteConsoleMsg(1, UserIndex, MOTD(j), FontTypeNames.FONTTYPE_TALK)
        Next j

        'Add Marius Mandamos el evento activo.
        If Len(MsgEvento) > 0 Then
            Call WriteConsoleMsg(1, UserIndex, MsgEvento, FontTypeNames.FONTTYPE_TALK)
        End If
        '\Add

    End Sub
    '\ReAdd

    Sub ResetFacciones(ByVal UserIndex As Integer, Optional Muertos As Boolean = True)
        With UserList(UserIndex).faccion

            If .ArmadaReal <> 100 Then .ArmadaReal = 0
            If .FuerzasCaos <> 100 Then .FuerzasCaos = 0
            If .Milicia <> 100 Then .Milicia = 0
            .Rango = 0

            .Ciudadano = 0
            .Republicano = 0
            .Renegado = 0

            If Muertos = True Then
                .CaosMatados = 0
                .ArmadaMatados = 0
                .MilicianosMatados = 0

                .CiudadanosMatados = 0
                .RenegadosMatados = 0
                .RepublicanosMatados = 0
            End If
        End With
    End Sub

    Sub ResetContadores(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '05/20/2007 Integer - Agregue todas las variables que faltaban.
        '*************************************************
        With UserList(UserIndex).Counters
            .AGUACounter = 0
            .AttackCounter = 0
            .Ceguera = 0
            .COMCounter = 0
            .Estupidez = 0
            .Frio = 0
            .HPCounter = 0
            .IdleCount = 0
            .Invisibilidad = 0
            .Paralisis = 0
            .Pena = 0
            .PiqueteC = 0
            .STACounter = 0
            .Veneno = 0
            .Fuego = 0
            .Trabajando = 0
            .Ocultando = 0
            .Saliendo = False
            .salir = 0
            .TiempoOculto = 0
            .TimerMove = 0
            '.TimerGolpeMagia = 0
            '.TimerLanzarSpell = 0
            .TimerPuedeAtacar = 0
            '.TimerPuedeUsarArco = 0
            .TimerPuedeTrabajar = 0
            .TimerUsar = 0
        End With
    End Sub

    Sub ResetCharInfo(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
        With UserList(UserIndex).cuerpo
            .body = 0
            .CascoAnim = 0
            .CharIndex = 0
            .fx = 0
            .Head = 0
            .loops = 0
            .heading = 0
            .loops = 0
            .ShieldAnim = 0
            .WeaponAnim = 0
        End With
    End Sub

    Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
        With UserList(UserIndex)
            .Name = vbNullString
            .desc = vbNullString

            .Pos.map = 0
            .Pos.x = 0
            .Pos.Y = 0
            .ip = vbNullString
            .Clase = 0
            .email = vbNullString
            .Genero = 0
            .Hogar = 0
            .Raza = 0

            .GrupoIndex = 0
            .GrupoSolicitud = 0

            With .Stats
                .Banco = 0
                .ELV = 0
                .ELU = 0
                .Exp = 0
                .def = 0
                '.CriminalesMatados = 0
                .NPCsMuertos = 0
                .VecesMuertos = 0

                .SkillPts = 0
                .GLD = 0
                .UserAtributos(1) = 0
                .UserAtributos(2) = 0
                .UserAtributos(3) = 0
                .UserAtributos(4) = 0
                .UserAtributos(5) = 0
                .UserAtributosBackUP(1) = 0
                .UserAtributosBackUP(2) = 0
                .UserAtributosBackUP(3) = 0
                .UserAtributosBackUP(4) = 0
                .UserAtributosBackUP(5) = 0
            End With

        End With
    End Sub


    Sub ResetGuildInfo(ByVal UserIndex As Integer)
        If UserList(UserIndex).EscucheClan > 0 Then
            Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
            UserList(UserIndex).EscucheClan = 0
        End If
        If UserList(UserIndex).GuildIndex > 0 Then
            Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)
        End If
        UserList(UserIndex).GuildIndex = 0
    End Sub

    Sub ResetUserFlags(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 06/28/2008
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '06/28/2008 NicoNZ - Agrego el flag Inmovilizado
        '*************************************************
        With UserList(UserIndex).flags
            .Comerciando = False
            .ban = 0
            .Escondido = 0
            .DuracionEfecto = 0
            .TargetNPC = 0
            .TargetNpcTipo = eNPCType.Comun
            .TargetObj = 0
            .TargetObjMap = 0
            .TargetObjX = 0
            .TargetObjY = 0
            .TargetUser = 0
            .TipoPocion = 0
            .TomoPocion = False
            .Hambre = 0
            .Sed = 0
            .Descansar = False
            .ModoCombate = False
            .Navegando = 0
            .Montando = 0
            .Oculto = 0
            .Envenenado = 0
            .Metamorfosis = 0
            .Incinerado = 0
            .Invisible = 0
            .Paralizado = 0
            .Inmovilizado = 0
            .Meditando = 0
            .Trabajando = 0
            .Lingoteando = 0
            .Privilegios = 0
            .OldBody = 0
            .OldHead = 0
            .AdminInvisible = 0
            .Hechizo = 0
            .TimesWalk = 0
            .Silenciado = 0
            .AdminPerseguible = False
        End With
    End Sub

    Sub ResetUserSpells(ByVal UserIndex As Integer)
        Dim loopC As Long
        For loopC = 1 To MAXUSERHECHIZOS
            UserList(UserIndex).Stats.UserHechizos(loopC) = 0
        Next loopC
    End Sub

    Sub ResetUserPets(ByVal UserIndex As Integer)
        Dim loopC As Long

        UserList(UserIndex).NroMascotas = 0

        For loopC = 1 To MAXMASCOTAS
            UserList(UserIndex).MascotasIndex(loopC) = 0
            UserList(UserIndex).MascotasType(loopC) = 0
        Next loopC

    End Sub

    Sub ResetUserBanco(ByVal UserIndex As Integer)
        Dim loopC As Long

        For loopC = 1 To MAX_BANCOINVENTORY_SLOTS
            UserList(UserIndex).BancoInvent.Objeto(loopC).Amount = 0
            UserList(UserIndex).BancoInvent.Objeto(loopC).Equipped = 0
            UserList(UserIndex).BancoInvent.Objeto(loopC).ObjIndex = 0
        Next loopC

        UserList(UserIndex).BancoInvent.NroItems = 0
    End Sub

    Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
        With UserList(UserIndex).ComUsu
            If .DestUsu > 0 Then
                Call FinComerciarUsu(.DestUsu)
                Call FinComerciarUsu(UserIndex)
            End If
        End With
    End Sub

    Sub ResetUserSlot(ByVal UserIndex As Integer)

        offcuenta(UserIndex)

        UserList(UserIndex).ConnIDValida = False
        UserList(UserIndex).flags.UserLogged = False
        UserList(UserIndex).ConnID = -1

        'If Not UserList(UserIndex).client Is Nothing Then
        ' UserList(UserIndex).client.flushOutGoingData()
        'End If

        UserList(UserIndex) = New User()

        UserList(UserIndex).ConnIDValida = False
        UserList(UserIndex).flags.UserLogged = False
        UserList(UserIndex).ConnID = -1

        UserList(UserIndex).incomingData = New clsByteQueue
        UserList(UserIndex).incomingData.Class_Initialize()
        UserList(UserIndex).outgoingData = New clsByteQueue
        UserList(UserIndex).outgoingData.Class_Initialize()
    End Sub


    Sub CloseUser(ByVal UserIndex As Integer)
        On Error GoTo Errhandler

        Dim N As Integer
        Dim loopC As Integer
        Dim map As Integer
        Dim Name As String
        Dim i As Integer

        Dim aN As Integer

        With UserList(UserIndex)



            If mapasEspeciales(UserIndex) Then
                Call Sum(UserIndex, 49, 50, 50, True)
            End If


            Call salir_arena(UserIndex)
            Call salir_listas_espera(UserIndex)


            If EsFacc(UserIndex) Then
                'Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg(1, "Server> " & .Name & " Se ha desconectado!", FontTypeNames.FONTTYPE_SERVER))
                Call LogGM(.Name, "Salió del juego")
            End If

            If NumUsers > 0 Then NumUsers = NumUsers - 1

            SendOnline()

            aN = .flags.AtacadoPorNpc
            If aN > 0 Then
                Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                Npclist(aN).flags.AttackedBy = 0
            End If
            aN = .flags.NPCAtacado
            If aN > 0 Then
                If Npclist(aN).flags.AttackedFirstBy = .Name Then
                    Npclist(aN).flags.AttackedFirstBy = vbNullString
                End If
            End If
            .flags.AtacadoPorNpc = 0
            .flags.NPCAtacado = 0

            Call ControlarPortalLum(UserIndex)


            map = .Pos.map
            Name = UCase$(.Name)

            UserList(UserIndex).cuerpo.fx = 0
            UserList(UserIndex).cuerpo.loops = 0


            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.cuerpo.CharIndex, 0, 0))


            .flags.UserLogged = False
            .Counters.Saliendo = False

            'Le devolvemos el body y head originales
            If .flags.AdminInvisible = 1 Then
                .cuerpo.body = .flags.OldBody
                .cuerpo.Head = .flags.OldHead
                .flags.AdminInvisible = 0
            End If

            'si esta en Grupo le devolvemos la experiencia
            If UserList(UserIndex).GrupoIndex > 0 Then Call mdGrupo.SalirDeGrupo(UserIndex)

            If .flags.inDuelo = 1 Then
                Call PerderDuelo(UserIndex)
            End If

            If .masc.invocado = True Then Call desinvocarfami(UserIndex)

            Call SaveUserSQL(UserIndex)

            If MapaValido(map) Then
                If MapInfoArr(map).NumUsers > 0 Then
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(.cuerpo.CharIndex))
                End If

                'Update Map Users
                MapInfoArr(map).NumUsers = MapInfoArr(map).NumUsers - 1

                If MapInfoArr(map).NumUsers < 0 Then
                    MapInfoArr(map).NumUsers = 0
                End If
            End If

            'Borrar el personaje
            If UserList(UserIndex).cuerpo.CharIndex > 0 Then
                Call EraseUserChar(UserIndex)
            End If

            If .NroMascotas > 0 Then
                For i = 1 To MAXMASCOTAS
                    If .MascotasIndex(i) > 0 Then
                        If Npclist(.MascotasIndex(i)).flags.NPCActive Then _
                    Call QuitarNPC(.MascotasIndex(i))
                    End If
                Next i
            End If

            ' Si el usuario habia dejado un msg en la gm's queue lo borramos
            'If Ayuda.Existe(.Name) Then Call Ayuda.Quitar(.Name)

            Call offpj(UserIndex)

            Call ResetGuildInfo(UserIndex)

        End With


        Exit Sub

Errhandler:
        Dim UserName As String
        If UserIndex > 0 Then UserName = UserList(UserIndex).Name

        Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.Description &
        ".User: " & UserName & "(" & UserIndex & "). Map: " & map)

    End Sub

    Sub EntregarMascota(ByVal UserIndex As Integer, petTipe As eMascota, ByRef petName As String)
        With UserList(UserIndex)
            If .Clase = eClass.Mago Then
                If petTipe > 5 Or petTipe < 1 Then
                    petTipe = eMascota.Fuego
                End If
            Else
                If petTipe < 6 Then
                    petTipe = eMascota.Ent
                End If
            End If

            .masc.TieneFamiliar = 1
            .masc.tipo = petTipe
            .masc.Nombre = petName

            .masc.ELV = 1
            .masc.ELU = 100
            .masc.MinHP = 10
            .masc.MaxHP = 10
            Select Case petTipe
                Case eMascota.Fuego, eMascota.Tierra
                    .masc.MinHit = 5
                    .masc.MaxHit = 15

                Case eMascota.Agua
                    .masc.MinHit = 7
                    .masc.MaxHit = 20

                Case eMascota.Ely
                    .masc.MinHP = 15
                    .masc.MaxHP = 15
                    .masc.MinHit = 5
                    .masc.MaxHit = 20

                Case eMascota.Fatuo
                    .masc.MinHP = 7
                    .masc.MaxHP = 7
                    .masc.MinHit = 5
                    .masc.MaxHit = 10

            'Caza o Druida
                Case eMascota.Tigre
                    .masc.MinHP = 15
                    .masc.MaxHP = 15
                    .masc.MinHit = 10
                    .masc.MaxHit = 20

                Case eMascota.Lobo
                    .masc.MinHP = 20
                    .masc.MaxHP = 20
                    .masc.MinHit = 10
                    .masc.MaxHit = 20

                Case eMascota.Oso
                    .masc.MinHP = 20
                    .masc.MaxHP = 20
                    .masc.MinHit = 5
                    .masc.MaxHit = 30

                Case eMascota.Ent
                    .masc.MinHP = 17
                    .masc.MaxHP = 17
                    .masc.MinHit = 10
                    .masc.MaxHit = 15
            End Select

        End With
    End Sub
    Public Sub EcharPjsNoPrivilegiados()
        Dim loopC As Long

        For loopC = 1 To LastUser
            If UserList(loopC).flags.UserLogged And UserList(loopC).ConnID >= 0 And UserList(loopC).ConnIDValida Then
                If UserList(loopC).flags.Privilegios And PlayerType.User Then
                    Call SaveUserSQL(CInt(loopC))
                    Call CloseSocket(loopC)
                End If
            End If
        Next loopC

    End Sub
    Function Generate_Char_Stat(ByVal UserIndex As Integer) As Byte
        With UserList(UserIndex)
            If .flags.Envenenado > 0 Then
                Generate_Char_Stat = Generate_Char_Stat Or lStat.Envenenado
            End If

            If .flags.Trabajando = 1 Then
                Generate_Char_Stat = Generate_Char_Stat Or lStat.Trabajando
            End If

            If .flags.Silenciado = 1 Then
                Generate_Char_Stat = Generate_Char_Stat Or lStat.Silenciado
            End If

            If .flags.Ceguera = 1 Then
                Generate_Char_Stat = Generate_Char_Stat Or lStat.Ciego
            End If

            If .flags.Incinerado = 1 Then
                Generate_Char_Stat = Generate_Char_Stat Or lStat.Incinerado
            End If

            If .flags.Metamorfosis = 1 Then
                Generate_Char_Stat = Generate_Char_Stat Or lStat.Transformado
            End If

            If .flags.Comerciando = 1 Then
                Generate_Char_Stat = Generate_Char_Stat Or lStat.Comerciand
            End If

            If .Counters.IdleCount > 1 Then
                Generate_Char_Stat = Generate_Char_Stat Or lStat.Inactivo
            End If
        End With
    End Function
    Function Generate_Char_StatEx(ByVal UserIndex As Integer) As Byte

        With UserList(UserIndex)
            If .flags.Paralizado = 1 Then
                Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Paralizado
            End If

            If .flags.Inmovilizado = 1 Then
                Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Inmovilizado
            End If

            If .Genero = eGenero.Hombre Then
                Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Hombre
            Else
                Generate_Char_StatEx = Generate_Char_StatEx Or lStatEx.Mujer
            End If
        End With
    End Function


    Sub CloseSocketSL(ByVal UserIndex As Integer)

        UserList(UserIndex).client.clientSocket.Close()

    End Sub

End Module
