Option Explicit On
Imports System.IO
Imports System.Runtime.InteropServices

Module ES

    '***************************
    'Sinuhe - Map format .CSM
    '***************************
    Private Structure tMapHeader
        Dim NumeroBloqueados As Long
        Dim NumeroLayers() As Long
        Dim NumeroTriggers As Long
        Dim NumeroLuces As Long
        Dim NumeroParticulas As Long
        Dim NumeroNPCs As Long
        Dim NumeroOBJs As Long
        Dim NumeroTE As Long
    End Structure

    Private Structure tDatosBloqueados
        Dim x As Integer
        Dim Y As Integer
    End Structure

    Private Structure tDatosGrh
        Dim x As Integer
        Dim Y As Integer
        Dim GrhIndex As Long
    End Structure

    Private Structure tDatosTrigger
        Dim x As Integer
        Dim Y As Integer
        Dim Trigger As Integer
    End Structure

    Private Structure tDatosLuces
        Dim x As Integer
        Dim Y As Integer
        Dim color As Long
        Dim Rango As Byte
    End Structure

    Private Structure tDatosParticulas
        Dim x As Integer
        Dim Y As Integer
        Dim Particula As Long
    End Structure

    Private Structure tDatosNPC
        Dim x As Integer
        Dim Y As Integer
        Dim NpcIndex As Integer
    End Structure

    Private Structure tDatosObjs
        Dim x As Integer
        Dim Y As Integer
        Dim ObjIndex As Integer
        Dim ObjAmmount As Integer
    End Structure

    Private Structure tDatosTE
        Dim x As Integer
        Dim Y As Integer
        Dim DestM As Integer
        Dim DestX As Integer
        Dim DestY As Integer
    End Structure

    Private Structure tMapSize
        Dim XMax As Integer
        Dim XMin As Integer
        Dim YMax As Integer
        Dim YMin As Integer
    End Structure

    Private Structure tMapDat
        Dim map_name As String
        Dim battle_mode As Byte
        Dim backup_mode As Byte
        Dim restrict_mode As String
        Dim music_number As String
        Dim zone As String
        Dim terrain As String
        Dim ambient As String
        Dim base_light As Long
        Dim letter_grh As Long
        Dim extra1 As Long
        Dim extra2 As Long
        Dim extra3 As String
    End Structure
    Public Sub CargarSpawnList()
        Dim N As Integer, loopC As Integer
        N = Val(GetVar(Application.StartupPath & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
        ReDim SpawnList(N)
        For loopC = 1 To N
            SpawnList(loopC).NpcIndex = Val(GetVar(Application.StartupPath & "\Dat\Invokar.dat", "LIST", "NI" & loopC))
            SpawnList(loopC).NpcName = GetVar(Application.StartupPath & "\Dat\Invokar.dat", "LIST", "NN" & loopC)
        Next loopC

    End Sub

    Public Sub CargarHechizos()

        '###################################################
        '#               ATENCION PELIGRO                  #
        '###################################################
        '
        '  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
        '
        'El que ose desafiar esta LEY, se las tendrá que ver
        'con migo. Para leer Hechizos.dat se deberá usar
        'la nueva clase clsLeerInis.
        '
        'Alejo
        '
        '###################################################

        On Error GoTo Errhandler

        Console.WriteLine("Cargando Hechizos ....")

        Dim strArchivoHechizos As String
        strArchivoHechizos = getValorPropiedadDats("Hechizos")
        Dim fileHechi As New System.IO.StreamWriter(DatPath & "Hechizos.dat")
        fileHechi.Write(strArchivoHechizos)
        fileHechi.Close()


        Dim Hechizo As Integer
        Dim leer As New clsIniReader

        Call leer.Initialize(DatPath & "Hechizos.dat")

        'obtiene el numero de hechizos
        NumeroHechizos = Val(leer.GetValue("INIT", "NumeroHechizos"))

        ReDim Hechizos(NumeroHechizos + 1)

        'frmCargando.cargar.min = 0
        'frmCargando.cargar.max = NumeroHechizos
        'frmCargando.cargar.value = 0

        'Llena la lista
        For Hechizo = 1 To NumeroHechizos

            Hechizos(Hechizo).Nombre = leer.GetValue("Hechizo" & Hechizo, "Nombre")
            Hechizos(Hechizo).desc = leer.GetValue("Hechizo" & Hechizo, "Desc")
            Hechizos(Hechizo).PalabrasMagicas = leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")

            Hechizos(Hechizo).HechizeroMsg = leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            Hechizos(Hechizo).TargetMsg = leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            Hechizos(Hechizo).PropioMsg = leer.GetValue("Hechizo" & Hechizo, "PropioMsg")

            Hechizos(Hechizo).tipo = Val(leer.GetValue("Hechizo" & Hechizo, "Tipo"))

            Hechizos(Hechizo).WAV = Val(leer.GetValue("Hechizo" & Hechizo, "WAV"))

            Hechizos(Hechizo).FXgrh = Val(leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            Hechizos(Hechizo).loops = Val(leer.GetValue("Hechizo" & Hechizo, "Loops"))

            Hechizos(Hechizo).Particle = Val(leer.GetValue("Hechizo" & Hechizo, "Particle"))

            Hechizos(Hechizo).SubeHP = Val(leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            Hechizos(Hechizo).MinHP = Val(leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            Hechizos(Hechizo).MaxHP = Val(leer.GetValue("Hechizo" & Hechizo, "MaxHP"))

            Hechizos(Hechizo).SubeAgilidad = Val(leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            Hechizos(Hechizo).MinAgilidad = Val(leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            Hechizos(Hechizo).MaxAgilidad = Val(leer.GetValue("Hechizo" & Hechizo, "MaxAG"))

            Hechizos(Hechizo).SubeFuerza = Val(leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            Hechizos(Hechizo).MinFuerza = Val(leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            Hechizos(Hechizo).MaxFuerza = Val(leer.GetValue("Hechizo" & Hechizo, "MaxFU"))

            Hechizos(Hechizo).Invisibilidad = Val(leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            Hechizos(Hechizo).Paraliza = Val(leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            Hechizos(Hechizo).Inmoviliza = Val(leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            Hechizos(Hechizo).RemoverParalisis = Val(leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            Hechizos(Hechizo).RemueveInvisibilidadParcial = Val(leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))

            Hechizos(Hechizo).CuraVeneno = Val(leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            Hechizos(Hechizo).Envenena = Val(leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            Hechizos(Hechizo).Incinera = Val(leer.GetValue("Hechizo" & Hechizo, "Incinera"))
            Hechizos(Hechizo).Revivir = Val(leer.GetValue("Hechizo" & Hechizo, "Revivir"))

            'Add Marius
            Hechizos(Hechizo).ExclusivoClase = UCase$(leer.GetValue("Hechizo" & Hechizo, "ExclusivoClase"))
            '\Add

            Hechizos(Hechizo).Resurreccion = Val(leer.GetValue("Hechizo" & Hechizo, "Resurreccion"))
            Hechizos(Hechizo).ReviveFamiliar = Val(leer.GetValue("Hechizo" & Hechizo, "ResucitaFamiliar"))
            Hechizos(Hechizo).Sanacion = Val(leer.GetValue("Hechizo" & Hechizo, "Sanacion"))

            Hechizos(Hechizo).Ceguera = Val(leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            Hechizos(Hechizo).Estupidez = Val(leer.GetValue("Hechizo" & Hechizo, "Estupidez"))

            Hechizos(Hechizo).Invoca = Val(leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            Hechizos(Hechizo).NumNpc = Val(leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            Hechizos(Hechizo).Cant = Val(leer.GetValue("Hechizo" & Hechizo, "Cant"))
            Hechizos(Hechizo).Mimetiza = Val(leer.GetValue("hechizo" & Hechizo, "Mimetiza"))

            Hechizos(Hechizo).AutoLanzar = Val(leer.GetValue("hechizo" & Hechizo, "autolanzar"))
            Hechizos(Hechizo).Desencantar = Val(leer.GetValue("hechizo" & Hechizo, "desencantar"))

            Hechizos(Hechizo).HechizoDeArea = Val(leer.GetValue("hechizo" & Hechizo, "HechizoDeArea"))
            Hechizos(Hechizo).AreaEfecto = Val(leer.GetValue("hechizo" & Hechizo, "AreaEfecto"))
            Hechizos(Hechizo).Afecta = Val(leer.GetValue("hechizo" & Hechizo, "Afecta"))
            Hechizos(Hechizo).Metamorfosis = Val(leer.GetValue("hechizo" & Hechizo, "metamorfosis"))

            Hechizos(Hechizo).Certero = Val(leer.GetValue("hechizo" & Hechizo, "GolpeCertero"))

            If Hechizos(Hechizo).Metamorfosis = 1 Then
                Hechizos(Hechizo).Extrahit = Val(leer.GetValue("hechizo" & Hechizo, "ExtraHIT"))
                Hechizos(Hechizo).Extradef = Val(leer.GetValue("hechizo" & Hechizo, "ExtraDEF"))
                Hechizos(Hechizo).body = Val(leer.GetValue("hechizo" & Hechizo, "body"))
                Hechizos(Hechizo).Head = Val(leer.GetValue("hechizo" & Hechizo, "head"))
                Hechizos(Hechizo).MetaObj = Val(leer.GetValue("hechizo" & Hechizo, "MetaObj"))
            End If

            If Hechizos(Hechizo).tipo = TipoHechizo.uCreateMagic Then
                Hechizos(Hechizo).CreaAlgo = Val(leer.GetValue("Hechizo" & Hechizo, "CreaTipo"))
                Hechizos(Hechizo).MaxHit = Val(leer.GetValue("Hechizo" & Hechizo, "MaxHit"))
                Hechizos(Hechizo).MinHit = Val(leer.GetValue("Hechizo" & Hechizo, "MinHit"))

                Hechizos(Hechizo).MaxDef = Val(leer.GetValue("Hechizo" & Hechizo, "MaxDef"))
                Hechizos(Hechizo).MinDef = Val(leer.GetValue("Hechizo" & Hechizo, "MinDef"))
            End If

            Hechizos(Hechizo).MinSkill = Val(leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            Hechizos(Hechizo).ManaRequerido = Val(leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))

            Hechizos(Hechizo).StaRequerido = Val(leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))

            Hechizos(Hechizo).Anillo = Val(leer.GetValue("Hechizo" & Hechizo, "Anillo"))

            Hechizos(Hechizo).Target = Val(leer.GetValue("Hechizo" & Hechizo, "Target"))

            'frmCargando.cargar.value = frmCargando.cargar.value + 1
            Console.WriteLine("Cargando hechizos " & Hechizo & "/" & NumeroHechizos)
            Application.DoEvents()
        Next Hechizo

        leer = Nothing
        Exit Sub

Errhandler:
        MsgBox("Error cargando hechizos.dat " & Err.Number & ": " & Err.Description)

    End Sub

    'ReAdd Marius
    Sub LoadMotd()
        Dim i As Integer

        MotdMaxLines = Val(GetVar(Application.StartupPath & "\Dat\Motd.ini", "INIT", "NumLines"))

        ReDim MOTD(MotdMaxLines)
        For i = 1 To MotdMaxLines
            MOTD(i) = GetVar(Application.StartupPath & "\Dat\Motd.ini", "Motd", "Line" & i)
        Next i

    End Sub
    '\ReAdd

    'Add Marius
    Sub LoadPublicidad()
        Dim i As Integer

        PubMaxLines = Val(GetVar(Application.StartupPath & "\Dat\publicidad.ini", "INIT", "NumLines"))

        ReDim PUBLICIDAD(PubMaxLines)
        For i = 1 To PubMaxLines
            PUBLICIDAD(i) = GetVar(Application.StartupPath & "\Dat\publicidad.ini", "publicidad", "Line" & i)
        Next i

    End Sub

    Sub SendPublicidad()
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, PUBLICIDAD(RandomNumber(1, PubMaxLines)), FontTypeNames.FONTTYPE_GUILD))
    End Sub
    '\Add


    Sub LoadArmasHerreria()

        Dim N As Integer, lc As Integer

        N = Val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

        ReDim ArmasHerrero(N + 1)

        For lc = 1 To N
            ArmasHerrero(lc) = Val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
        Next lc

    End Sub

    Sub LoadArmadurasHerreria()

        Dim N As Integer, lc As Integer

        N = Val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

        ReDim ArmadurasHerrero(N + 1)

        For lc = 1 To N
            ArmadurasHerrero(lc) = Val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
        Next lc

    End Sub

    Sub LoadBalance()

        Try
            Dim i As Long

            ReDim ModClaseArr(NUMCLASES + 1)

            For i = 1 To NUMCLASES
                ModClaseArr(i).Evasion = getValorPropiedadBalance("EVASION" + ListaClases(i))
                ModClaseArr(i).AtaqueArmas = getValorPropiedadBalance("ATAQUEARMAS" + ListaClases(i))
                ModClaseArr(i).AtaqueProyectiles = getValorPropiedadBalance("ATAQUEPROYECTILES" + ListaClases(i))
                ModClaseArr(i).DañoArmas = getValorPropiedadBalance("DANIOARMAS" + ListaClases(i))
                ModClaseArr(i).DañoProyectiles = getValorPropiedadBalance("DANIOPROYECTILES" + ListaClases(i))
                ModClaseArr(i).DañoWrestling = getValorPropiedadBalance("DANIOLUCHA" + ListaClases(i))
                ModClaseArr(i).Escudo = getValorPropiedadBalance("ESCUDO" + ListaClases(i))
            Next i

            ReDim ModRazaArr(NUMRAZAS + 1)

            For i = 1 To NUMRAZAS
                ModRazaArr(i).Fuerza = getValorPropiedadBalance(ListaRazas(i) + "Fuerza")
                ModRazaArr(i).Agilidad = getValorPropiedadBalance(ListaRazas(i) + "Agilidad")
                ModRazaArr(i).Inteligencia = getValorPropiedadBalance(ListaRazas(i) + "Inteligencia")
                ModRazaArr(i).Carisma = getValorPropiedadBalance(ListaRazas(i) + "Carisma")
                ModRazaArr(i).constitucion = getValorPropiedadBalance(ListaRazas(i) + "Constitucion")
            Next i


            ReDim ModVida(NUMCLASES + 1)

            'Modificadores de Vida
            For i = 1 To NUMCLASES
                ModVida(i) = getValorPropiedadBalance("VIDA" + ListaClases(i))
            Next i

            'Modificadores de Vida
            For i = 1 To NUMCLASES
                ModMana(i) = getValorPropiedadBalance("MANA" + ListaClases(i))
            Next i


            For i = 1 To 4
                DistribucionSemienteraVida(i) = getValorPropiedadBalance("DISTRIBUCIONS" + CStr(i))
            Next i

            'Extra
            PorcentajeRecuperoMana = getValorPropiedadBalance("PorcentajeRecuperoMana")
            'Grupo
            ExponenteNivelGrupo = getValorPropiedadBalance("ExponenteNivelGrupo")

            'Intervalos
            SanaIntervaloSinDescansar = getValorPropiedadBalance("SanaIntervaloSinDescansar")
            StaminaIntervaloSinDescansar = getValorPropiedadBalance("StaminaIntervaloSinDescansar")
            SanaIntervaloDescansar = getValorPropiedadBalance("SanaIntervaloDescansar")
            StaminaIntervaloDescansar = getValorPropiedadBalance("StaminaIntervaloDescansar")

            IntervaloSed = getValorPropiedadBalance("IntervaloSed")
            IntervaloHambre = getValorPropiedadBalance("IntervaloHambre")
            IntervaloVeneno = getValorPropiedadBalance("IntervaloVeneno")
            IntervaloParalizado = getValorPropiedadBalance("IntervaloParalizado")

            IntervaloInvisible = getValorPropiedadBalance("IntervaloInvisible")
            IntervaloFrio = getValorPropiedadBalance("IntervaloFrio")
            IntervaloWavFx = getValorPropiedadBalance("IntervaloWAVFX")
            IntervaloInvocacion = getValorPropiedadBalance("IntervaloInvocacion")
            IntervaloParaConexion = getValorPropiedadBalance("IntervaloParaConexion")

            getValorPropiedadBalance("IntervaloNpcAI")
            getValorPropiedadBalance("IntervaloNpcPuedeAtacar")

            IntervaloUserPuedeTrabajar = getValorPropiedadBalance("IntervaloTrabajo")

            IntervaloGolpeUsar = getValorPropiedadBalance("IntervaloGolpeUsar")

            IntervaloCerrarConexion = getValorPropiedadBalance("IntervaloCerrarConexion")
            IntervaloUserPuedeUsar = getValorPropiedadBalance("IntervaloUserPuedeUsar")
            IntervaloUserPuedeAtacar = getValorPropiedadBalance("IntervaloUserPuedeAtacar")

            IntervaloUserPuedeCastear = getValorPropiedadBalance("IntervaloLanzaHechizo")
            IntervaloMagiaGolpe = getValorPropiedadBalance("IntervaloMagiaGolpe")
            IntervaloGolpeMagia = getValorPropiedadBalance("IntervaloGolpeMagia")
            IntervaloFlechasCazadores = getValorPropiedadBalance("IntervaloFlechasCazadores")
            IntervaloUserMove = getValorPropiedadBalance("IntervaloMover")

            IntervaloOculto = getValorPropiedadBalance("IntervaloOculto")

        Catch ex As Exception
            Dim st As New StackTrace(True)
            st = New StackTrace(ex, True)
            LogError("Error en LoadBalance: " + ex.Message + " StackTrace: " + st.ToString())
        End Try

    End Sub

    Sub LoadObjCarpintero()

        Dim N As Integer, lc As Integer

        N = Val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

        ReDim ObjCarpintero(N + 1)

        For lc = 1 To N
            ObjCarpintero(lc) = Val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
        Next lc

    End Sub

    Sub LoadObjDruida()

        Dim N As Integer, lc As Integer

        N = Val(GetVar(DatPath & "objdruida.dat", "INIT", "NumObjs"))

        ReDim ObjDruida(N + 1)

        For lc = 1 To N
            ObjDruida(lc) = Val(GetVar(DatPath & "objdruida.dat", "Obj" & lc, "Index"))
        Next lc

    End Sub

    Sub LoadObjSastre()

        Dim N As Integer, lc As Integer

        N = Val(GetVar(DatPath & "objsastre.dat", "INIT", "NumObjs"))

        ReDim ObjSastre(N + 1)

        For lc = 1 To N
            ObjSastre(lc) = Val(GetVar(DatPath & "objsastre.dat", "Obj" & lc, "Index"))
        Next lc

    End Sub




    Sub LoadObjData()

        '###################################################
        '#               ATENCION PELIGRO                  #
        '###################################################
        '
        '¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
        '
        'El que ose desafiar esta LEY, se las tendrá que ver
        'conmigo. Para leer desde el OBJ.DAT se deberá usar
        'la nueva clase clsLeerInis.
        '
        'Alejo
        '
        '###################################################

        'Call LogTarea("Sub LoadOBJData")

        On Error GoTo Errhandler

        Console.WriteLine("Cargando objetos....")

        '*****************************************************************
        'Carga la lista de objetos
        '*****************************************************************
        Dim Objeto As Integer
        Dim leer As New clsIniReader

        Dim consulta As String

        Dim cpParams As String
        Dim cpValues As String

        Call leer.Initialize(DatPath & "Obj.dat")

        'obtiene el numero de obj
        NumObjDatas = Val(leer.GetValue("INIT", "NumObjs"))

        'frmCargando.cargar.min = 0
        'frmCargando.cargar.max = NumObjDatas
        'frmCargando.cargar.value = 0


        ReDim ObjDataArr(NumObjDatas + 1)

        'Llena la lista
        For Objeto = 1 To NumObjDatas

            cpParams = ""
            cpValues = ""

            ObjDataArr(Objeto).Name = leer.GetValue("OBJ" & Objeto, "Name")

            'Pablo (ToxicWaste) Log de Objetos.
            ObjDataArr(Objeto).Log = Val(leer.GetValue("OBJ" & Objeto, "Log"))
            ObjDataArr(Objeto).NoLog = Val(leer.GetValue("OBJ" & Objeto, "NoLog"))
            '07/09/07

            ObjDataArr(Objeto).GrhIndex = Val(leer.GetValue("OBJ" & Objeto, "GrhIndex"))
            If ObjDataArr(Objeto).GrhIndex = 0 Then
                ObjDataArr(Objeto).GrhIndex = ObjDataArr(Objeto).GrhIndex
            End If

            ObjDataArr(Objeto).OBJType = Val(leer.GetValue("OBJ" & Objeto, "ObjType"))

            ObjDataArr(Objeto).Newbie = Val(leer.GetValue("OBJ" & Objeto, "Newbie"))

            ObjDataArr(Objeto).SubTipo = Val(leer.GetValue("OBJ" & Objeto, "SubTipo"))

            ObjDataArr(Objeto).EfectoMagico = Val(leer.GetValue("OBJ" & Objeto, "efectomagico"))
            ObjDataArr(Objeto).CuantoAumento = Val(leer.GetValue("OBJ" & Objeto, "cuantoaumento"))

            ObjDataArr(Objeto).Shop = 0
            ObjDataArr(Objeto).Shop = Val(leer.GetValue("OBJ" & Objeto, "Shop"))

            With ObjDataArr(Objeto)
                Select Case ObjDataArr(Objeto).OBJType
                    Case eOBJType.otItemsMagicos
                        Select Case ObjDataArr(Objeto).EfectoMagico
                            Case 1, 2, 3, 6, 7, 12, 14, 18
                                '  ObjDataArr(Objeto).CuantoAumento = val(leer.GetValue("OBJ" & Objeto, "cuantoaumento"))
                        End Select

                        If ObjDataArr(Objeto).EfectoMagico = 2 Then _
                    ObjDataArr(Objeto).QueAtributo = Val(leer.GetValue("OBJ" & Objeto, "QueAtributo"))

                        If ObjDataArr(Objeto).EfectoMagico = 3 Then _
                ObjDataArr(Objeto).QueSkill = Val(leer.GetValue("OBJ" & Objeto, "QueSkill"))

                    Case eOBJType.otArmadura

                        If .SubTipo = 1 Then
                            ObjDataArr(Objeto).CascoAnim = Val(leer.GetValue("OBJ" & Objeto, "Anim"))
                        ElseIf .SubTipo = 2 Then
                            ObjDataArr(Objeto).ShieldAnim = Val(leer.GetValue("OBJ" & Objeto, "Anim"))
                            ObjDataArr(Objeto).DosManos = Val(leer.GetValue("OBJ" & Objeto, "DosManos"))
                        End If

                        ObjDataArr(Objeto).Real = Val(leer.GetValue("OBJ" & Objeto, "Real"))
                        ObjDataArr(Objeto).Caos = Val(leer.GetValue("OBJ" & Objeto, "Caos"))
                        ObjDataArr(Objeto).Milicia = Val(leer.GetValue("OBJ" & Objeto, "Milicia"))

                        ObjDataArr(Objeto).LingH = Val(leer.GetValue("OBJ" & Objeto, "LingH"))
                        ObjDataArr(Objeto).LingP = Val(leer.GetValue("OBJ" & Objeto, "LingP"))
                        ObjDataArr(Objeto).LingO = Val(leer.GetValue("OBJ" & Objeto, "LingO"))
                        ObjDataArr(Objeto).SkHerreria = Val(leer.GetValue("OBJ" & Objeto, "SkHerreria"))

                    Case eOBJType.otNudillos
                        ObjDataArr(Objeto).LingH = Val(leer.GetValue("OBJ" & Objeto, "LingH"))
                        ObjDataArr(Objeto).LingP = Val(leer.GetValue("OBJ" & Objeto, "LingP"))
                        ObjDataArr(Objeto).LingO = Val(leer.GetValue("OBJ" & Objeto, "LingO"))
                        ObjDataArr(Objeto).SkHerreria = Val(leer.GetValue("OBJ" & Objeto, "SkHerreria"))
                        ObjDataArr(Objeto).MaxHit = Val(leer.GetValue("OBJ" & Objeto, "MaxHIT"))
                        ObjDataArr(Objeto).MinHit = Val(leer.GetValue("OBJ" & Objeto, "MinHIT"))
                        ObjDataArr(Objeto).WeaponAnim = Val(leer.GetValue("OBJ" & Objeto, "Anim"))

                    Case eOBJType.otWeapon
                        ObjDataArr(Objeto).WeaponAnim = Val(leer.GetValue("OBJ" & Objeto, "Anim"))
                        ObjDataArr(Objeto).Apuñala = Val(leer.GetValue("OBJ" & Objeto, "Apuñala"))
                        ObjDataArr(Objeto).MaxHit = Val(leer.GetValue("OBJ" & Objeto, "MaxHIT"))
                        ObjDataArr(Objeto).MinHit = Val(leer.GetValue("OBJ" & Objeto, "MinHIT"))
                        ObjDataArr(Objeto).proyectil = Val(leer.GetValue("OBJ" & Objeto, "Proyectil"))
                        ObjDataArr(Objeto).Municion = Val(leer.GetValue("OBJ" & Objeto, "Municiones"))
                        ObjDataArr(Objeto).Refuerzo = Val(leer.GetValue("OBJ" & Objeto, "Refuerzo"))

                        ObjDataArr(Objeto).LingH = Val(leer.GetValue("OBJ" & Objeto, "LingH"))
                        ObjDataArr(Objeto).LingP = Val(leer.GetValue("OBJ" & Objeto, "LingP"))
                        ObjDataArr(Objeto).LingO = Val(leer.GetValue("OBJ" & Objeto, "LingO"))
                        ObjDataArr(Objeto).SkHerreria = Val(leer.GetValue("OBJ" & Objeto, "SkHerreria"))
                        ObjDataArr(Objeto).Real = Val(leer.GetValue("OBJ" & Objeto, "Real"))
                        ObjDataArr(Objeto).Caos = Val(leer.GetValue("OBJ" & Objeto, "Caos"))
                        ObjDataArr(Objeto).DosManos = Val(leer.GetValue("OBJ" & Objeto, "DosManos"))

                    Case eOBJType.otInstrumentos
                        ObjDataArr(Objeto).Snd1 = Val(leer.GetValue("OBJ" & Objeto, "SND1"))
                        ObjDataArr(Objeto).Snd2 = Val(leer.GetValue("OBJ" & Objeto, "SND2"))
                        ObjDataArr(Objeto).Snd3 = Val(leer.GetValue("OBJ" & Objeto, "SND3"))

                        ObjDataArr(Objeto).Real = Val(leer.GetValue("OBJ" & Objeto, "Real"))
                        ObjDataArr(Objeto).Caos = Val(leer.GetValue("OBJ" & Objeto, "Caos"))
                        ObjDataArr(Objeto).Milicia = Val(leer.GetValue("OBJ" & Objeto, "Milicia"))

                    Case eOBJType.otMinerales
                        ObjDataArr(Objeto).MinSkill = Val(leer.GetValue("OBJ" & Objeto, "MinSkill"))

                    Case eOBJType.otMonturas
                        ObjDataArr(Objeto).MinSkill = Val(leer.GetValue("OBJ" & Objeto, "MinSkill"))

                    Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                        ObjDataArr(Objeto).IndexAbierta = Val(leer.GetValue("OBJ" & Objeto, "IndexAbierta"))
                        ObjDataArr(Objeto).IndexCerrada = Val(leer.GetValue("OBJ" & Objeto, "IndexCerrada"))
                        ObjDataArr(Objeto).IndexCerradaLlave = Val(leer.GetValue("OBJ" & Objeto, "IndexCerradaLlave"))

                    Case eOBJType.otPociones
                        ObjDataArr(Objeto).TipoPocion = Val(leer.GetValue("OBJ" & Objeto, "TipoPocion"))
                        ObjDataArr(Objeto).MaxModificador = Val(leer.GetValue("OBJ" & Objeto, "MaxModificador"))
                        ObjDataArr(Objeto).MinModificador = Val(leer.GetValue("OBJ" & Objeto, "MinModificador"))
                        ObjDataArr(Objeto).DuracionEfecto = Val(leer.GetValue("OBJ" & Objeto, "DuracionEfecto"))

                    Case eOBJType.otBarcos
                        ObjDataArr(Objeto).MinSkill = Val(leer.GetValue("OBJ" & Objeto, "MinSkill"))
                        ObjDataArr(Objeto).MaxHit = Val(leer.GetValue("OBJ" & Objeto, "MaxHIT"))
                        ObjDataArr(Objeto).MinHit = Val(leer.GetValue("OBJ" & Objeto, "MinHIT"))

                    Case eOBJType.otFlechas
                        ObjDataArr(Objeto).MaxHit = Val(leer.GetValue("OBJ" & Objeto, "MaxHIT"))
                        ObjDataArr(Objeto).MinHit = Val(leer.GetValue("OBJ" & Objeto, "MinHIT"))
                        ObjDataArr(Objeto).SubTipo = Val(leer.GetValue("OBJ" & Objeto, "SubTipo"))

                    Case eOBJType.otPasajes
                        ObjDataArr(Objeto).DesdeMap = Val(leer.GetValue("OBJ" & Objeto, "Desde"))
                        ObjDataArr(Objeto).HastaMap = Val(leer.GetValue("OBJ" & Objeto, "Map"))
                        ObjDataArr(Objeto).HastaX = Val(leer.GetValue("OBJ" & Objeto, "X"))
                        ObjDataArr(Objeto).HastaY = Val(leer.GetValue("OBJ" & Objeto, "Y"))
                        ObjDataArr(Objeto).CantidadSkill = Val(leer.GetValue("OBJ" & Objeto, "CantidadSkill"))

                    Case eOBJType.otContenedores
                        ObjDataArr(Objeto).CuantoAgrega = Val(leer.GetValue("OBJ" & Objeto, "CuantoAgrega"))

                End Select
            End With

            ObjDataArr(Objeto).Ropaje = Val(leer.GetValue("OBJ" & Objeto, "NumRopaje"))
            ObjDataArr(Objeto).HechizoIndex = Val(leer.GetValue("OBJ" & Objeto, "HechizoIndex"))



            If Val(leer.GetValue("OBJ" & Objeto, "SICPO")) = 1 Then
                ObjDataArr(Objeto).CPO = leer.GetValue("OBJ" & Objeto, "CPO")
            End If

            ObjDataArr(Objeto).LingoteIndex = Val(leer.GetValue("OBJ" & Objeto, "LingoteIndex"))

            ObjDataArr(Objeto).MineralIndex = Val(leer.GetValue("OBJ" & Objeto, "MineralIndex"))

            ObjDataArr(Objeto).MaxHP = Val(leer.GetValue("OBJ" & Objeto, "MaxHP"))
            ObjDataArr(Objeto).MinHP = Val(leer.GetValue("OBJ" & Objeto, "MinHP"))

            ObjDataArr(Objeto).Mujer = Val(leer.GetValue("OBJ" & Objeto, "Mujer"))
            ObjDataArr(Objeto).Hombre = Val(leer.GetValue("OBJ" & Objeto, "Hombre"))

            ObjDataArr(Objeto).MinHAM = Val(leer.GetValue("OBJ" & Objeto, "MinHam"))
            ObjDataArr(Objeto).MinSed = Val(leer.GetValue("OBJ" & Objeto, "MinAgu"))

            ObjDataArr(Objeto).MinDef = Val(leer.GetValue("OBJ" & Objeto, "MINDEF"))
            ObjDataArr(Objeto).MaxDef = Val(leer.GetValue("OBJ" & Objeto, "MAXDEF"))
            ObjDataArr(Objeto).def = (ObjDataArr(Objeto).MinDef + ObjDataArr(Objeto).MaxDef) * 0.5

            ObjDataArr(Objeto).RazaTipo = Val(leer.GetValue("OBJ" & Objeto, "RazaTipo"))
            ObjDataArr(Objeto).RazaEnana = Val(leer.GetValue("OBJ" & Objeto, "RazaEnana"))
            ObjDataArr(Objeto).MinELV = Val(leer.GetValue("OBJ" & Objeto, "MinELV"))

            ObjDataArr(Objeto).valor = Val(leer.GetValue("OBJ" & Objeto, "Valor"))

            ObjDataArr(Objeto).Crucial = Val(leer.GetValue("OBJ" & Objeto, "Crucial"))

            ObjDataArr(Objeto).Cerrada = Val(leer.GetValue("OBJ" & Objeto, "abierta"))
            If ObjDataArr(Objeto).Cerrada = 1 Then
                ObjDataArr(Objeto).Llave = Val(leer.GetValue("OBJ" & Objeto, "Llave"))
                ObjDataArr(Objeto).clave = Val(leer.GetValue("OBJ" & Objeto, "Clave"))
            End If

            'Puertas y llaves
            ObjDataArr(Objeto).clave = Val(leer.GetValue("OBJ" & Objeto, "Clave"))

            ObjDataArr(Objeto).texto = leer.GetValue("OBJ" & Objeto, "Texto")
            ObjDataArr(Objeto).GrhSecundario = Val(leer.GetValue("OBJ" & Objeto, "VGrande"))

            ObjDataArr(Objeto).Agarrable = Val(leer.GetValue("OBJ" & Objeto, "Agarrable"))
            ObjDataArr(Objeto).ForoID = leer.GetValue("OBJ" & Objeto, "ID")

            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
            Dim i As Integer
            Dim N As Integer
            Dim S As String
            i = 1 : N = 1
            S = leer.GetValue("OBJ" & Objeto, "CP" & i)

            ReDim ObjDataArr(Objeto).ClaseProhibida(NUMCLASES + 1)

            Do While Len(S) > 0

                If ClaseToEnum(S) > 0 Then ObjDataArr(Objeto).ClaseProhibida(N) = ClaseToEnum(S)

                cpParams = cpParams + "`cp" + "" & i & "" + "`, "
                cpValues = cpValues + "'" + S + "',"

            If N = NUMCLASES Then Exit Do

                N = N + 1 : i = i + 1
                S = leer.GetValue("OBJ" & Objeto, "CP" & i)

            Loop


            ObjDataArr(Objeto).ClaseTipo = Val(leer.GetValue("OBJ" & Objeto, "ClaseTipo"))

            ObjDataArr(Objeto).DefensaMagicaMax = Val(leer.GetValue("OBJ" & Objeto, "DefensaMagicaMax"))
            ObjDataArr(Objeto).DefensaMagicaMin = Val(leer.GetValue("OBJ" & Objeto, "DefensaMagicaMin"))




            ' Jose Castelli  / Resistencia Magica (RM)

            ObjDataArr(Objeto).ResistenciaMagica = Val(leer.GetValue("OBJ" & Objeto, "ResistenciaMagica"))

            ' Jose Castelli  / Resistencia Magica (RM)

            ObjDataArr(Objeto).SkCarpinteria = Val(leer.GetValue("OBJ" & Objeto, "SkCarpinteria"))

            If ObjDataArr(Objeto).SkCarpinteria > 0 Then
                ObjDataArr(Objeto).Madera = Val(leer.GetValue("OBJ" & Objeto, "Madera"))
            End If


            ObjDataArr(Objeto).SkPociones = Val(leer.GetValue("OBJ" & Objeto, "SkPociones"))

            If ObjDataArr(Objeto).SkPociones > 0 Then
                ObjDataArr(Objeto).raies = Val(leer.GetValue("OBJ" & Objeto, "Raices"))
            End If


            ObjDataArr(Objeto).SkSastreria = Val(leer.GetValue("OBJ" & Objeto, "SkSastreria"))

            If ObjDataArr(Objeto).SkSastreria > 0 Then
                ObjDataArr(Objeto).PielLobo = Val(leer.GetValue("OBJ" & Objeto, "PielLobo"))
                ObjDataArr(Objeto).PielLoboInvernal = Val(leer.GetValue("OBJ" & Objeto, "PielOsoPolar"))
                ObjDataArr(Objeto).PielOso = Val(leer.GetValue("OBJ" & Objeto, "PielOsoPardo"))
            End If

            Console.WriteLine("Cargando objetos " & Objeto & "/" & NumObjDatas)
            'frmCargando.cargar.value = frmCargando.cargar.value + 1

            'Dim anim As Integer

            'anim = val(leer.GetValue("OBJ" & Objeto, "Anim"))

            'With ObjDataArr(Objeto)


            '            consulta = "INSERT INTO `inmortalao`.`objs`(`name`,`grhIndex`,`objType`,`newbie`,`subTipo`,`efectoMagico`,`cuantoAumento`," &
            '"`shop`,`queAtributo`,`queSkill`,`anim`,`dosManos`,`real`,`caos`,`milicia`,`lingH`,`lingP`,`lingO`,`skHerreria`,`maxHit`," &
            '"`minHit`,`proyectil`,`municiones`,`refuerzo`,`snd1`,`snd2`,`snd3`,`minSkill`,`indexAbierta`,`indexCerrada`,`indexCerradaLlave`," &
            '"`tipoPocion`,`maxModificador`,`minModificador`,`duracionEfecto`,`desde`,`map`,`x`,`y`,`cantidadSkill`,`cuantoAgrega`," &
            '"`numRopaje`,`hechizoIndex`,`cpo`,`lingoteIndex`,`mineralIndex`,`maxHp`,`minHp`,`mujer`,`hombre`,`minHam`,`minAgu`," &
            '"`minDef`,`maxDef`,`razaTipo`,`razaEnana`,`minElv`,`valor`,`crucial`,`abierta`,`texto`,`grhSecundario`,`agarrable`," &
            '"`foroId`," + cpParams + "" &
            '"`claseTipo`,`resistenciaMagica`,`skCarpinteria`,`madera`,`skpociones`,`raices`,`skSastreria`,`pielLobo`," &
            '"`pielOsoPilar`,`pielOsoPardo`) values " &
            '"('" + .name + "'," & .GrhIndex & "," & .OBJType & "," & .Newbie & "," & .SubTipo & "," & .EfectoMagico & "," &
            '"" & .CuantoAumento & "," & .Shop & "," & .QueAtributo & "," & .QueSkill & "," & anim & "," & .DosManos & "," & .Real & "," & .Caos & "," & .Milicia & "," & .LingH & "," &
            '"" & .LingP & "," & .LingO & "," & .SkHerreria & "," & .maxhit & "," & .minhit & "," & .proyectil & "," & .Municion & "," & .Refuerzo & "," & .snd1 & "," & .snd2 & "," & .snd3 & "," & .MinSkill & "," & .IndexAbierta & "," & .IndexCerrada & "," & .IndexCerradaLlave & "," & .TipoPocion & "," &
            '"" & .MaxModificador & "," & .MinModificador & "," & .DuracionEfecto & "," & .DesdeMap & "," & .HastaMap & "," & .HastaX & "," & .HastaY & "," & .CantidadSkill & "," & .CuantoAgrega & "," & .Ropaje & "," & .HechizoIndex & ",'" + .CPO + "'," & .LingoteIndex & "," & .MineralIndex & "," & .maxhp & "," & .minhp & "," &
            '"" & .Mujer & "," & .Hombre & "," & .MinHAM & "," & .MinSed & "," & .MinDef & "," & .MaxDef & "," & .RazaTipo & "," & .RazaEnana & "," & .MinELV & "," & .valor & "," & .Crucial & "," & .Cerrada & ",'" + .texto + "'," & .GrhSecundario & "," & .Agarrable & "," &
            '"'" + .ForoID + "'," &
            '"" + cpValues + "" & .ClaseTipo & "," & .ResistenciaMagica & "," & .SkCarpinteria & "," &
            '"" & .Madera & "," & .SkPociones & "," & .raies & "," & .SkSastreria & "," & .PielLobo & "," & .PielLoboInvernal & "," & .PielOso & ");"

            '            'End With

            '   DB_Conn.Execute (consulta)

            Application.DoEvents()

        Next Objeto

        leer = Nothing

        Exit Sub

Errhandler:

        MsgBox("error cargando objetos " & Err.Number & ": " & Err.Description)


    End Sub


    Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

        Dim sSpaces As String ' This will hold the input that the program will retrieve
        Dim szReturn As String ' This will be the defaul value if the string is not found

        szReturn = vbNullString

        sSpaces = Space(EmptySpaces) ' This tells the computer how long the longest string can be


        GetPrivateProfileString(Main, Var, szReturn, sSpaces, EmptySpaces, file)

        GetVar = RTrim(sSpaces)
        GetVar = Left(GetVar, Len(GetVar) - 1)

    End Function



    Sub LoadMapData()

        Dim map As Integer
        Dim TempInt As Integer
        Dim tFileName As String
        Dim npcfile As String

        On Error GoTo man

        'NumMaps = Val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps")) 
        NumMaps = 851
        'NumMaps = 248
        Call InitAreas()

        '  frmCargando.cargar.min = 0
        '  frmCargando.cargar.max = NumMaps
        '  frmCargando.cargar.value = 0

        'MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        MapPath = "\Maps\"

        ReDim MapData(NumMaps + 1, XMaxMapSize + 1, YMaxMapSize + 1)
        ReDim MapInfoArr(NumMaps + 1)

        For map = 1 To NumMaps

            tFileName = Application.StartupPath & MapPath & "Mapa" & map
            Call CargarMapa(map, tFileName)


            'Add Marius Agregamos mas mapas seguros
            'Se agrego: intermundia (por decision de los users), lindos izquierda, lindos abajo, suramei abajo y bander derecha, tiama puerto
            If map = Banderbill.map Or map = Arghal.map Or
           map = Rinkel.map Or map = Lindos.map Or
           map = Suramei.map Or map = Illiandor.map Or
           map = Orac.map Or map = Tiama.map Or
           map = Ullathorpe.map Or map = Nix.map Or
           map = NuevaEsperanza.map Or
           map = NuevaEsperanzapuerto.map Or
           map = 49 Or map = 364 Or map = 64 Or map = 63 Or map = 184 Or map = 217 Then

                MapInfoArr(map).Pk = False
            End If

            Console.WriteLine("Cargando Mapas " & map & "\" & NumMaps)

            'frmCargando.cargar.value = frmCargando.cargar.value + 1
            Application.DoEvents()
        Next map

        'Add Marius Agregamos mas arenas =P Algun día tendremos WorldEdit
        'Call CargarMapa(238, tFileName)
        'Call CargarMapa(238, tFileName)
        'Call CargarMapa(238, tFileName)
        'Call CargarMapa(238, tFileName)
        '\Add

        Exit Sub

man:
        MsgBox("Error durante la carga de mapas, el mapa " & map & " contiene errores")
        Call LogError(DateTime.Now & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

    End Sub
    Public Sub CargarMapa(ByVal map As Long, ByVal MAPFl As String)
        On Error GoTo errh

        Dim fh As Integer
        Dim MH As tMapHeader
        Dim Blqs() As tDatosBloqueados
        'Dim L1() As Long
        Dim L2() As tDatosGrh
        Dim L3() As tDatosGrh
        Dim L4() As tDatosGrh
        Dim Triggers() As tDatosTrigger
        Dim Luces() As tDatosLuces
        Dim Particulas() As tDatosParticulas
        Dim Objetos() As tDatosObjs
        Dim NPCs() As tDatosNPC
        Dim TEs() As tDatosTE
        Dim MapSize As tMapSize
        Dim MapDat As tMapDat

        Dim i As Long
        Dim j As Long

        If Not FileExist(MAPFl & ".txt") Then _
        Exit Sub


        Dim lineaActual As String()

        Dim objReader As New System.IO.StreamReader(MAPFl & ".txt")

        lineaActual = objReader.ReadLine().Replace("""", "").Split(";") 'MH
        MH.NumeroBloqueados = lineaActual(0)
        ReDim MH.NumeroLayers(4)
        MH.NumeroLayers(2) = lineaActual(1)
        MH.NumeroLayers(3) = lineaActual(2)
        MH.NumeroLayers(4) = lineaActual(3)
        MH.NumeroLuces = lineaActual(4)
        MH.NumeroNPCs = lineaActual(5)
        MH.NumeroOBJs = lineaActual(6)
        MH.NumeroParticulas = lineaActual(7)
        MH.NumeroTE = lineaActual(8)
        MH.NumeroTriggers = lineaActual(9)


        lineaActual = objReader.ReadLine().Replace("""", "").Split(";") 'MapSize
        MapSize.XMax = lineaActual(0)
        MapSize.XMin = lineaActual(1)
        MapSize.YMax = lineaActual(2)
        MapSize.YMin = lineaActual(3)



        lineaActual = objReader.ReadLine().Split(";") 'MapDat
        MapDat.ambient = lineaActual(0).Replace("""", "")
        MapDat.backup_mode = lineaActual(1)
        MapDat.base_light = lineaActual(2)
        MapDat.battle_mode = 0 'lineaActual(3)
        MapDat.extra1 = lineaActual(4)
        MapDat.extra2 = lineaActual(5)
        MapDat.extra3 = lineaActual(6)
        MapDat.letter_grh = lineaActual(7)
        MapDat.map_name = lineaActual(8)
        MapDat.music_number = lineaActual(9)
        MapDat.restrict_mode = lineaActual(10)
        MapDat.terrain = lineaActual(11)
        MapDat.zone = lineaActual(12)

        'ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
        Dim L1(MapSize.XMax, MapSize.YMax) As Long

        Dim lA As String

        Dim x As Integer
        Dim Y As Integer
        For Y = MapSize.YMin To MapSize.YMax
            For x = MapSize.XMin To MapSize.XMax
                L1(x, Y) = objReader.ReadLine().Replace("""", "")
            Next x
        Next Y
        Dim vf As Long

        With MH
            If .NumeroBloqueados > 0 Then
                ReDim Blqs(.NumeroBloqueados + 1)
                For vf = 1 To .NumeroBloqueados
                    lA = objReader.ReadLine()
                    Blqs(vf).x = lA.Replace("""", "").Split(";")(0)
                    Blqs(vf).Y = lA.Replace("""", "").Split(";")(1)
                Next vf
            End If

            If .NumeroLayers(2) > 0 Then
                ReDim L2(.NumeroLayers(2) + 1)
                For vf = 1 To .NumeroLayers(2)
                    lA = objReader.ReadLine()
                    L2(vf).GrhIndex = lA.Replace("""", "").Split(";")(0)
                    L2(vf).x = lA.Replace("""", "").Split(";")(1)
                    L2(vf).Y = lA.Replace("""", "").Split(";")(2)
                Next vf
            End If

            If .NumeroLayers(3) > 0 Then
                ReDim L3(.NumeroLayers(3) + 1)
                For vf = 1 To .NumeroLayers(3)
                    lA = objReader.ReadLine()
                    L3(vf).GrhIndex = lA.Replace("""", "").Split(";")(0)
                    L3(vf).x = lA.Replace("""", "").Split(";")(1)
                    L3(vf).Y = lA.Replace("""", "").Split(";")(2)
                Next vf
            End If

            If .NumeroLayers(4) > 0 Then
                ReDim L4(.NumeroLayers(4) + 1)
                For vf = 1 To .NumeroLayers(4)
                    lA = objReader.ReadLine()
                    L4(vf).GrhIndex = lA.Replace("""", "").Split(";")(0)
                    L4(vf).x = lA.Replace("""", "").Split(";")(1)
                    L4(vf).Y = lA.Replace("""", "").Split(";")(2)
                Next vf
            End If

            If .NumeroTriggers > 0 Then
                ReDim Triggers(.NumeroTriggers + 1)
                For vf = 1 To .NumeroTriggers
                    lA = objReader.ReadLine()
                    Triggers(vf).Trigger = lA.Replace("""", "").Split(";")(0)
                    Triggers(vf).x = lA.Replace("""", "").Split(";")(1)
                    Triggers(vf).Y = lA.Replace("""", "").Split(";")(2)
                Next vf
            End If

            If .NumeroParticulas > 0 Then
                ReDim Particulas(.NumeroParticulas + 1)
                For vf = 1 To .NumeroParticulas
                    lA = objReader.ReadLine()
                    Particulas(vf).Particula = lA.Replace("""", "").Split(";")(0)
                    Particulas(vf).x = lA.Replace("""", "").Split(";")(1)
                    Particulas(vf).Y = lA.Replace("""", "").Split(";")(2)
                Next vf
            End If

            If .NumeroLuces > 0 Then
                ReDim Luces(.NumeroLuces + 1)
                For vf = 1 To .NumeroLuces
                    lA = objReader.ReadLine()
                    Luces(vf).color = lA.Replace("""", "").Split(";")(0)
                    Luces(vf).Rango = lA.Replace("""", "").Split(";")(1)
                    Luces(vf).x = lA.Replace("""", "").Split(";")(2)
                    Luces(vf).Y = lA.Replace("""", "").Split(";")(3)
                Next vf
            End If

            If .NumeroOBJs > 0 Then
                ReDim Objetos(.NumeroOBJs + 1)
                For vf = 1 To .NumeroOBJs
                    lA = objReader.ReadLine()
                    Objetos(vf).ObjAmmount = lA.Replace("""", "").Split(";")(0)
                    Objetos(vf).ObjIndex = lA.Replace("""", "").Split(";")(1)
                    Objetos(vf).x = lA.Replace("""", "").Split(";")(2)
                    Objetos(vf).Y = lA.Replace("""", "").Split(";")(3)
                Next vf
            End If

            If .NumeroNPCs > 0 Then
                ReDim NPCs(.NumeroNPCs + 1)
                For vf = 1 To .NumeroNPCs
                    lA = objReader.ReadLine()
                    NPCs(vf).NpcIndex = lA.Replace("""", "").Split(";")(0)
                    NPCs(vf).x = lA.Replace("""", "").Split(";")(1)
                    NPCs(vf).Y = lA.Replace("""", "").Split(";")(2)
                Next vf
            End If

            If .NumeroTE > 0 Then
                ReDim TEs(.NumeroTE + 1)
                For vf = 1 To .NumeroTE
                    lA = objReader.ReadLine()
                    TEs(vf).DestM = lA.Replace("""", "").Split(";")(0)
                    TEs(vf).DestX = lA.Replace("""", "").Split(";")(1)
                    TEs(vf).DestY = lA.Replace("""", "").Split(";")(2)
                    TEs(vf).x = lA.Replace("""", "").Split(";")(3)
                    TEs(vf).Y = lA.Replace("""", "").Split(";")(4)
                Next vf
            End If

        End With


        With MH
            If .NumeroBloqueados > 0 Then
                For i = 1 To .NumeroBloqueados
                    MapData(map, Blqs(i).x, Blqs(i).Y).Blocked = 1
                Next i
            End If


            If .NumeroLayers(2) > 0 Then
                For i = 1 To .NumeroLayers(2)
                    ReDim Preserve MapData(map, L2(i).x, L2(i).Y).Graphic(5)
                    MapData(map, L2(i).x, L2(i).Y).Graphic(2) = L2(i).GrhIndex
                Next i
            End If

            If .NumeroLayers(3) > 0 Then
                For i = 1 To .NumeroLayers(3)
                    ReDim Preserve MapData(map, L3(i).x, L3(i).Y).Graphic(5)
                    MapData(map, L3(i).x, L3(i).Y).Graphic(3) = L3(i).GrhIndex
                Next i
            End If

            If .NumeroLayers(4) > 0 Then
                For i = 1 To .NumeroLayers(4)
                    ReDim Preserve MapData(map, L4(i).x, L4(i).Y).Graphic(5)
                    MapData(map, L4(i).x, L4(i).Y).Graphic(4) = L4(i).GrhIndex
                Next i
            End If

            If .NumeroTriggers > 0 Then
                For i = 1 To .NumeroTriggers
                    MapData(map, Triggers(i).x, Triggers(i).Y).Trigger = Triggers(i).Trigger
                Next i
            End If

            If .NumeroOBJs > 0 Then
                For i = 1 To .NumeroOBJs
                    MapData(map, Objetos(i).x, Objetos(i).Y).ObjInfo.ObjIndex = Objetos(i).ObjIndex
                    MapData(map, Objetos(i).x, Objetos(i).Y).ObjInfo.Amount = Objetos(i).ObjAmmount

                    If ObjDataArr(Objetos(i).ObjIndex).OBJType <> eOBJType.otPuertas Then
                        MapData(map, Objetos(i).x, Objetos(i).Y).ObjEsFijo = 1
                    End If
                Next i
            End If

            If .NumeroNPCs > 0 Then
                For i = 1 To .NumeroNPCs
                    MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex = NPCs(i).NpcIndex
                    If MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex > 0 Then
                        Dim npcfile As String

                        npcfile = DatPath & "NPCs.dat"

                        If Val(GetVar(npcfile, "NPC" & MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex, "PosOrig")) = 1 Then
                            MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex = OpenNPC(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex)
                            Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Orig.map = map
                            Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Orig.x = NPCs(i).x
                            Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Orig.Y = NPCs(i).Y
                        Else
                            MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex = OpenNPC(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex)
                        End If
                        If Not MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex = 0 Then
                            Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Pos.map = map
                            Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Pos.x = NPCs(i).x
                            Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).Pos.Y = NPCs(i).Y

                            Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).StartPos.map = map
                            Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).StartPos.x = NPCs(i).x
                            Npclist(MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex).StartPos.Y = NPCs(i).Y

                            Call MakeNPCChar(True, 0, MapData(map, NPCs(i).x, NPCs(i).Y).NpcIndex, map, NPCs(i).x, NPCs(i).Y)
                        End If
                    End If
                Next i
            End If

            If .NumeroTE > 0 Then
                For i = 1 To .NumeroTE
                    MapData(map, TEs(i).x, TEs(i).Y).TileExit.map = TEs(i).DestM
                    MapData(map, TEs(i).x, TEs(i).Y).TileExit.x = TEs(i).DestX
                    MapData(map, TEs(i).x, TEs(i).Y).TileExit.Y = TEs(i).DestY
                Next i
            End If
    End With

    For j = MapSize.YMin To MapSize.YMax
    For i = MapSize.XMin To MapSize.XMax
                If L1(i, j) > 0 Then
                    ReDim Preserve MapData(map, i, j).Graphic(5)
                    MapData(map, i, j).Graphic(1) = L1(i, j)
                End If
            Next i
        Next j

        MapDat.map_name = Trim$(MapDat.map_name)

        MapInfoArr(map).Name = MapDat.map_name
        MapInfoArr(map).Music = MapDat.music_number
        MapInfoArr(map).Seguro = MapDat.extra1

        If Not (Left$(MapDat.zone, 6) = "CIUDAD") Then
            MapInfoArr(map).Pk = True
        Else
            MapInfoArr(map).Pk = False
        End If

        MapInfoArr(map).Terreno = MapDat.terrain
        MapInfoArr(map).Zona = Trim$(MapDat.zone)
        MapInfoArr(map).Restringir = MapDat.restrict_mode
        MapInfoArr(map).BackUp = MapDat.backup_mode

        Exit Sub

errh:
    Call LogError("Error cargando mapa: " & map & " ." & Err.Description)
    End Sub

    Sub LoadSini()

        Dim Temporal As Long

        'ReAdd Marius
        Call LoadMotd()
        '\ReAdd

        'Add Marius
        Call LoadPublicidad()
        '\Add

        Console.WriteLine("Cargando server.ini")

        Puerto = Val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
        RondasAutomatico = Val(GetVar(IniPath & "Server.ini", "INIT", "RondasAutomatico"))

        ModExpX = getValorPropiedadBalance("Experiencia")
        ModOroX = getValorPropiedadBalance("Oro")

        PuedenFundarClan = Val(GetVar(IniPath & "Server.ini", "INIT", "PuedenFundarClan"))
        PuedeBorrarClan = Val(GetVar(IniPath & "Server.ini", "INIT", "PuedeBorrarClan"))

        ModTrabajo = Val(GetVar(IniPath & "Server.ini", "INIT", "ModTrabajo"))
        ModSkill = Val(GetVar(IniPath & "Server.ini", "INIT", "ModSkill"))

        'AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
        'IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))

        'Lee la version correcta del cliente
        ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")
        PuedeCrearPersonajes = Val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
        ServerSoloGMs = Val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))

        RecordUsuarios = Val(GetVar(IniPath & "Server.ini", "INIT", "Record"))

        'Max users
        Temporal = Val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
        If MaxUsers = 0 Then
            MaxUsers = Temporal
            ReDim UserList(MaxUsers)
        End If

        '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
        'Se agregó en LoadBalance y en el Balance.dat
        'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))

        ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&


        Nix.map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
        Nix.x = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
        Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

        Ullathorpe.map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
        Ullathorpe.x = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
        Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

        Banderbill.map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
        Banderbill.x = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
        Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

        Arghal.map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
        Arghal.x = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
        Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")

        Illiandor.map = GetVar(DatPath & "Ciudades.dat", "Illiandor", "Mapa")
        Illiandor.x = GetVar(DatPath & "Ciudades.dat", "Illiandor", "X")
        Illiandor.Y = GetVar(DatPath & "Ciudades.dat", "Illiandor", "Y")

        Suramei.map = GetVar(DatPath & "Ciudades.dat", "Suramei", "Mapa")
        Suramei.x = GetVar(DatPath & "Ciudades.dat", "Suramei", "X")
        Suramei.Y = GetVar(DatPath & "Ciudades.dat", "Suramei", "Y")

        Lindos.map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
        Lindos.x = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
        Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

        Orac.map = GetVar(DatPath & "Ciudades.dat", "Orac", "Mapa")
        Orac.x = GetVar(DatPath & "Ciudades.dat", "Orac", "X")
        Orac.Y = GetVar(DatPath & "Ciudades.dat", "Orac", "Y")

        Rinkel.map = GetVar(DatPath & "Ciudades.dat", "Rinkel", "Mapa")
        Rinkel.x = GetVar(DatPath & "Ciudades.dat", "Rinkel", "X")
        Rinkel.Y = GetVar(DatPath & "Ciudades.dat", "Rinkel", "Y")

        Tiama.map = GetVar(DatPath & "Ciudades.dat", "Tiama", "Mapa")
        Tiama.x = GetVar(DatPath & "Ciudades.dat", "Tiama", "X")
        Tiama.Y = GetVar(DatPath & "Ciudades.dat", "Tiama", "Y")

        NuevaEsperanza.map = GetVar(DatPath & "Ciudades.dat", "NuevaEsperanza", "Mapa")
        NuevaEsperanza.x = GetVar(DatPath & "Ciudades.dat", "NuevaEsperanza", "X")
        NuevaEsperanza.Y = GetVar(DatPath & "Ciudades.dat", "NuevaEsperanza", "Y")

        NuevaEsperanzapuerto.map = GetVar(DatPath & "Ciudades.dat", "Nuevaesperanzapuerto", "Mapa")
        NuevaEsperanzapuerto.x = GetVar(DatPath & "Ciudades.dat", "Nuevaesperanzapuerto", "X")
        NuevaEsperanzapuerto.Y = GetVar(DatPath & "Ciudades.dat", "Nuevaesperanzapuerto", "Y")

    End Sub

    Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
        '*****************************************************************
        'Escribe VAR en un archivo
        '*****************************************************************

        WritePrivateProfileString(Main, Var, value, file)

    End Sub



    Sub BackUPnPc(NpcIndex As Integer)

        Dim NpcNumero As Integer
        Dim npcfile As String
        Dim loopC As Integer


        NpcNumero = Npclist(NpcIndex).Numero

        npcfile = DatPath & "bkNPCs.dat"

        'General
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).Name)
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).desc)
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", Val(Npclist(NpcIndex).cuerpo.Head))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", Val(Npclist(NpcIndex).cuerpo.body))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", Val(Npclist(NpcIndex).cuerpo.heading))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", Val(Npclist(NpcIndex).Movement))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", Val(Npclist(NpcIndex).Attackable))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", Val(Npclist(NpcIndex).Comercia))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", Val(Npclist(NpcIndex).TipoItems))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", Val(Npclist(NpcIndex).Hostile))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", Val(Npclist(NpcIndex).GiveEXP))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", Val(Npclist(NpcIndex).GiveGLD))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", Val(Npclist(NpcIndex).Hostile))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", Val(Npclist(NpcIndex).InvReSpawn))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", Val(Npclist(NpcIndex).NPCtype))


        'Stats
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", Val(Npclist(NpcIndex).Stats.Alineacion))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", Val(Npclist(NpcIndex).Stats.def))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", Val(Npclist(NpcIndex).Stats.MaxHit))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", Val(Npclist(NpcIndex).Stats.MaxHP))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", Val(Npclist(NpcIndex).Stats.MinHit))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", Val(Npclist(NpcIndex).Stats.MinHP))




        'Flags
        Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", Val(Npclist(NpcIndex).flags.respawn))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", Val(Npclist(NpcIndex).flags.BackUp))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", Val(Npclist(NpcIndex).flags.Domable))

        'Inventario
        Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", Val(Npclist(NpcIndex).Invent.NroItems))
        If Npclist(NpcIndex).Invent.NroItems > 0 Then
            For loopC = 1 To MAX_INVENTORY_SLOTS
                Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & loopC, Npclist(NpcIndex).Invent.Objeto(loopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Objeto(loopC).Amount)
            Next
        End If


    End Sub




    Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

        Call WriteVar(Application.StartupPath & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
        Call WriteVar(Application.StartupPath & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", motivo)

        Dim objReader As New System.IO.StreamWriter(Application.StartupPath & "\logs\GenteBanned.log")
        objReader.WriteLine(UserList(BannedIndex).Name)

    End Sub


    Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

        Call WriteVar(Application.StartupPath & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
        Call WriteVar(Application.StartupPath & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim objReader As New System.IO.StreamWriter(Application.StartupPath & "\logs\GenteBanned.log")
        objReader.WriteLine(BannedName)

    End Sub


    Sub ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

        Call WriteVar(Application.StartupPath & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
        Call WriteVar(Application.StartupPath & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)


        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim objReader As New System.IO.StreamWriter(Application.StartupPath & "\logs\GenteBanned.log")
        objReader.WriteLine(BannedName)

    End Sub

    Public Sub CargaApuestas()

        Apuestas.Ganancias = Val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
        Apuestas.Perdidas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
        Apuestas.Jugadas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

    End Sub


End Module
