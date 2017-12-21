Option Explicit On


Module Declaraciones

    Public Enum ALINEACION_GUILD
        ALINEACION_REPUBLICANO = 1
        ALINEACION_IMPERIAL = 2
        ALINEACION_CAOTICO = 3
        ALINEACION_RENEGADO = 4
    End Enum

    Public Const MENSAJES_TOPE_CORREO As Byte = 20 ' Cantidad de mensajes de correo

    'Add Nod Kopfnickend
    Public mins As Integer

    Public Security As New clsSecurity
    Public RondasAutomatico As Byte
    'Multiplicador de Experiencia, por Castelli
    Public ModExpX As Long
    Public ModOroX As Long
    ' ////

    Public PuedenFundarClan As Byte
    Public PuedeBorrarClan As Byte

    Public ModTrabajo As Long
    Public ModSkill As Long

    Public tHora As Byte
    Public tMinuto As Byte
    Public tSeg As Byte

    Public TrashCollector As New Collection

    Public Const MAXSPAWNATTEMPS = 60
    Public Const INFINITE_LOOPS As Integer = -1
    Public Const FXSANGRE = 14

    Public Const NO_3D_SOUND As Byte = 0

    Public Const iFragataFantasmal = 87
    Public Const iBarca = 84
    Public Const iGalera = 85
    Public Const iGaleon = 86

    Public Enum iMinerales
        HierroCrudo = 192
        PlataCruda = 193
        OroCrudo = 194
        LingoteDeHierro = 386
        LingoteDePlata = 387
        LingoteDeOro = 388
    End Enum

    Public Enum PlayerType
        User = &H1

        FaccImpe = &H2
        FaccRepu = &H4
        FaccCaos = &H8

        Conse = &H10
        Semi = &H40
        Dios = &H80
        Admin = &H100

        VIP = &H200
    End Enum

    Public Enum eClass
        Clerigo = 1
        Mago = 2
        Guerrero = 3
        Asesino = 4
        Ladron = 5
        Bardo = 6
        Druida = 7
        Gladiador = 8
        Paladin = 9
        Cazador = 10
        Pescador = 11
        Herrero = 12
        Leñador = 13
        Minero = 14
        Carpintero = 15
        Sastre = 16
        Mercenario = 17
        Nigromante = 18
    End Enum

    Public Enum eCiudad
        cUllathorpe = 1 'Imperial
        cNix            'Imperial
        cBanderbill     'Imperial
        cArghal         'Imperial

        cIlliandor      'Republicana
        cLindos         'Republicana
        cSuramei        'Republicana

        cOrac           'Coatica

        cNuevaEsperanza 'Neutral
        cRinkel         'Neutral
        cTiama          'Neutral
    End Enum

    Public Enum eRaza
        Humano = 1
        Elfo
        Drow
        Gnomo
        Enano
        Orco
    End Enum

    Enum eGenero
        Hombre = 1
        Mujer
    End Enum

    Public Const LimiteNewbie As Byte = 13

    Structure tCabecera 'Cabecera de los con
        Dim desc As String
        Dim crc As Long
        Dim MagicWord As Long
    End Structure

    Public MiCabecera As tCabecera

    'Barrin 3/10/03
    'Cambiado a 2 segundos el 30/11/07
    'Nod Kopfnickend Comento por que no se usa en ningun lado :S
    'Public Const TIEMPO_INICIOMEDITAR As Integer = 2000

    Public Const NingunEscudo As Integer = 2
    Public Const NingunCasco As Integer = 2
    Public Const NingunArma As Integer = 2

    Public Const EspadaMataDragonesIndex As Integer = 402

    Public Const MAXMASCOTASENTRENADOR As Byte = 7

    Public Enum FXIDs
        FXWARP = 1
        FXMEDITARCHICO = 4
        FXMEDITARMEDIANO = 5
        FXMEDITARGRANDE = 6
        FXMEDITARXGRANDE = 16
        FXMEDITARXXGRANDE = 34
    End Enum

    Public Const TIEMPO_CARCEL_PIQUETE As Long = 5

    ''
    ' TRIGGERS
    '
    ' @param NADA nada
    ' @param BAJOTECHO bajo techo
    ' @param trigger_2 ???
    ' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
    ' @param ZONASEGURA no se puede robar o pelear desde este trigger
    ' @param ANTIPIQUETE
    ' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
    '
    Public Enum eTrigger
        Nada = 0
        BAJOTECHO = 1
        trigger_2 = 2
        POSINVALIDA = 3
        ZONASEGURA = 4
        ANTIPIQUETE = 5
        ZONAPELEA = 6
    End Enum

    ''
    ' constantes para el trigger 6
    '
    ' @see eTrigger
    ' @param eTrigger6.TRIGGER6_PERMITE eTrigger6.TRIGGER6_PERMITE
    ' @param eTrigger6.TRIGGER6_PROHIBE eTrigger6.TRIGGER6_PROHIBE
    ' @param eTrigger6.TRIGGER6_PERMITE El trigger no aparece
    '
    Public Enum eTrigger6
        TRIGGER6_PERMITE = 1
        TRIGGER6_PROHIBE = 2
        TRIGGER6_AUSENTE = 3
    End Enum

    'TODO : Reemplazar por un enum
    Public Const Bosque As String = "BOSQUE"
    Public Const Nieve As String = "NIEVE"
    Public Const Desierto As String = "DESIERTO"
    Public Const Ciudad As String = "CIUDAD"
    Public Const Campo As String = "CAMPO"
    Public Const Dungeon As String = "DUNGEON"

    ' <<<<<< Targets >>>>>>
    Public Enum TargetType
        uUsuarios = 1
        uNPC = 2
        uUsuariosYnpc = 3
        uTerreno = 4
    End Enum

    ' <<<<<< Acciona sobre >>>>>>
    Public Enum TipoHechizo
        uPropiedades = 1
        uEstado = 2
        uInvocacion = 4
        uCreateTelep = 5
        uFamiliar = 6
        uMaterializa = 7
        uPropEsta = 8
        uCalmacion = 9
        uCreateMagic = 10
        uEquipamiento = 11
        uDetectarInvis = 12
    End Enum

    Public Const MAX_MENSAJES_FORO As Byte = 35

    Public Const MAXUSERHECHIZOS As Byte = 35


    ' TODO: Y ESTO ? LO CONOCE GD ?
    Public Const EsfuerzoTalarGeneral As Byte = 4
    Public Const EsfuerzoTalarLeñador As Byte = 2

    Public Const EsfuerzoBotanicaGeneral As Byte = 4
    Public Const EsfuerzoBotanicaDruida As Byte = 2

    Public Const EsfuerzoPescarPescador As Byte = 1
    Public Const EsfuerzoPescarGeneral As Byte = 3

    Public Const EsfuerzoExcavarMinero As Byte = 2
    Public Const EsfuerzoExcavarGeneral As Byte = 5

    Public Const FX_TELEPORT_INDEX As Integer = 1

    ' La utilidad de esto es casi nula, sólo se revisa si fue a la cabeza...
    Public Enum PartesCuerpo
        bCabeza = 1
        bPiernaIzquierda = 2
        bPiernaDerecha = 3
        bBrazoDerecho = 4
        bBrazoIzquierdo = 5
        bTorso = 6
    End Enum

    Public Const Guardias As Integer = 6

    Public Const MAXREP As Long = 6000000
    Public Const MAXORO As Long = 90000000
    Public Const MAXEXP As Long = 1999999990

    Public Const MAXUSERMATADOS As Long = 65000

    Public Const MAXATRIBUTOS As Byte = 35
    Public Const MINATRIBUTOS As Byte = 6

    Public Const LingoteHierro As Integer = 386
    Public Const LingotePlata As Integer = 387
    Public Const LingoteOro As Integer = 388
    Public Const Leña As Integer = 58
    Public Const Raiz As Integer = 888

    Public Const PielLobo As Integer = 414
    Public Const PielOso As Integer = 415
    Public Const PielLoboInvernal As Integer = 1145


    Public Const MAXNPCS As Integer = 10000
    Public Const MAXCHARS As Integer = 10000

    Public Const DAGA As Integer = 15
    Public Const FOGATA_APAG As Integer = 136
    Public Const FOGATA As Integer = 63

    Public Const ObjArboles As Integer = 4

    Public Const MARTILLO_HERRERO As Integer = 389
    Public Const SERRUCHO_CARPINTERO As Integer = 198
    Public Const HACHA_LEÑADOR As Integer = 127
    Public Const PIQUETE_MINERO As Integer = 187
    Public Const RED_PESCA As Integer = 138
    Public Const CAÑA_PESCA As Integer = 881
    Public Const TIJERAS As Integer = 885
    Public Const COSTURERO As Integer = 886
    Public Const OLLA As Integer = 887

    Public Enum eNPCType
        Comun = 0
        Revividor = 1
        GuardiaReal = 2
        Entrenador = 3
        Banquero = 4
        Noble = 5
        DRAGON = 6
        Timbero = 7
        Guardiascaos = 8
        ResucitadorNewbie = 9
        Pirata = 10
        Bot = 11
    End Enum
    '  0 NPCs Comunes
    '  1 Resucitadores
    '  2 Guardias
    '  3 Entrenadores
    '  4 Banqueros
    '  5 Facciones
    '  6 (nada)
    '  7 Transportadores
    '  8 Carceleros
    '  9 (nada)
    ' 10 Paga recompensas
    ' 11 Veterinarias
    ' 12 Apuestas
    ' 13 Presentadores de Quest
    ' 14 Objetivos de Quest
    ' 15 Centinelas
    ' 16 Subastadores
    Public Const MIN_APUÑALAR As Byte = 10

    '********** CONSTANTANTES ***********

    ''
    ' Cantidad de skills
    Public Const NUMSKILLS As Byte = 27

    ''
    ' Cantidad de Atributos
    Public Const NUMATRIBUTOS As Byte = 5

    ''
    ' Cantidad de Clases
    Public Const NUMCLASES As Byte = 18

    ''
    ' Cantidad de Razas
    Public Const NUMRAZAS As Byte = 6

    ''
    ' Valor maximo de cada skill
    Public Const MAXSKILLPOINTS As Byte = 100

    ''
    'Direccion
    '
    ' @param NORTH Norte
    ' @param EAST Este
    ' @param SOUTH Sur
    ' @param WEST Oeste
    '
    Public Enum eHeading
        NORTH = 1
        EAST = 2
        SOUTH = 3
        WEST = 4
    End Enum

    ''
    ' Cantidad maxima de mascotas
    Public Const MAXMASCOTAS As Byte = 3

    '%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
    Public Const vlASALTO As Integer = 100
    Public Const vlASESINO As Integer = 1000
    Public Const vlCAZADOR As Integer = 5
    Public Const vlNoble As Integer = 5
    Public Const vlLadron As Integer = 25
    Public Const vlProleta As Integer = 2

    '%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
    Public Const iCuerpoMuerto As Integer = 8
    Public Const iCabezaMuerto As Integer = 500


    Public Const iORO As Byte = 12
    Public Const Pescado As Byte = 139

    Public Enum PECES_POSIBLES
        PESCADO1 = 139
        PESCADO2 = 544
        PESCADO3 = 545
        PESCADO4 = 546
    End Enum

    '%%%%%%%%%% CONSTANTES DE INDICES %%%%%%%%%%%%%%%
    Public Enum eSkill
        Tacticas = 1
        armas = 2
        artes = 3
        Apuñalar = 4
        arrojadizas = 5
        Proyectiles = 6
        DefensaEscudos = 7
        Magia = 8
        Resistencia = 9
        Meditar = 10
        Ocultarse = 11
        Domar = 12
        Musica = 13
        Robar = 14
        Comerciar = 15
        Supervivencia = 16
        Liderazgo = 17
        Pesca = 18
        Mineria = 19
        Talar = 20
        botanica = 21
        Herreria = 22
        Carpinteria = 23
        alquimia = 24
        Sastreria = 25
        Navegacion = 26
        Equitacion = 27
    End Enum

    Public Const FundirMetal = 88

    Public Enum eAtributos
        Fuerza = 1
        Agilidad = 2
        Inteligencia = 3
        Carisma = 4
        constitucion = 5
    End Enum

    Public Const AdicionalHPGuerrero As Byte = 2 'HP adicionales cuando sube de nivel
    Public Const AdicionalHPCazador As Byte = 1 'HP adicionales cuando sube de nivel

    Public Const AumentoSTDef As Byte = 15
    Public Const AumentoSTLadron As Byte = AumentoSTDef + 3
    Public Const AumentoSTMago As Byte = AumentoSTDef - 1
    Public Const AumentoSTLeñador As Byte = AumentoSTDef + 23
    Public Const AumentoSTPescador As Byte = AumentoSTDef + 20
    Public Const AumentoSTMinero As Byte = AumentoSTDef + 25

    'Tamaño del mapa
    Public Const XMaxMapSize As Byte = 100
    Public Const XMinMapSize As Byte = 1
    Public Const YMaxMapSize As Byte = 100
    Public Const YMinMapSize As Byte = 1

    'Tamaño del tileset
    Public Const TileSizeX As Byte = 32
    Public Const TileSizeY As Byte = 32

    'Tamaño en Tiles de la pantalla de visualizacion
    Public Const XWindow As Byte = 17
    Public Const YWindow As Byte = 13

    'Sonidos
    Public Const SND_SWING As Byte = 2
    Public Const SND_TALAR As Byte = 13
    Public Const SND_PESCAR As Byte = 14
    Public Const SND_MINERO As Byte = 15
    Public Const SND_WARP As Byte = 3
    Public Const SND_PUERTA As Byte = 5
    Public Const SND_NIVEL As Byte = 128

    Public Const SND_USERMUERTE As Byte = 11
    Public Const SND_IMPACTO As Byte = 86 '10
    Public Const SND_IMPACTO2 As Byte = 86 '12
    Public Const SND_LEÑADOR As Byte = 13
    Public Const SND_FOGATA As Byte = 14
    Public Const SND_AVE As Byte = 21
    Public Const SND_AVE2 As Byte = 22
    Public Const SND_AVE3 As Byte = 34
    Public Const SND_GRILLO As Byte = 28
    Public Const SND_GRILLO2 As Byte = 29
    Public Const SND_SACARARMA As Byte = 25
    Public Const SND_ESCUDO As Byte = 37
    Public Const SND_BEBER As Byte = 135
    Public Const SND_REMO As Byte = 255
    Public Const SND_NEWCLAN As Byte = 44
    Public Const SND_NEWMEMBER As Byte = 43
    Public Const SND_OUT As Byte = 45
    Public Const SND_INCINERACION As Byte = 123
    Public Const SND_DROPITEM As Byte = 132
    Public Const SND_REVIVE As Byte = 204

    Public Const MARTILLOHERRERO As Byte = 150 '41
    Public Const LABUROCARPINTERO As Byte = 42

    ''
    ' Cantidad maxima de objetos por slot de inventario
    Public Const MAX_INVENTORY_OBJS As Integer = 10000

    ''
    ' Cantidad de "slots" en el inventario
    Public Const MAX_INVENTORY_SLOTS As Byte = 25

    ''
    ' Constante para indicar que se esta usando ORO
    Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1

    ' CATEGORIAS PRINCIPALES
    Public Enum eOBJType
        otUseOnce = 1           '  1 Comidas
        otWeapon = 2
        otArmadura = 3
        otArboles = 4
        otGuita = 5
        otPuertas = 6
        otContenedores = 7
        otCarteles = 8
        otLlaves = 9
        otForos = 10
        otPociones = 11
        otLibros = 12
        otBebidas = 13
        otLeña = 14
        otFuego = 15
        otESCUDO = 16
        otCASCO = 17
        otHerramientas = 18
        otTeleport = 19
        otMuebles = 20
        otItemsMagicos = 21
        otYacimiento = 22
        otMinerales = 23
        otPergaminos = 24
        otInstrumentos = 26
        otYunque = 27
        otFragua = 28
        otLingotes = 29
        otPieles = 30
        otBarcos = 31
        otFlechas = 32
        otBotellaVacia = 33
        otBotellaLlena = 34
        otManchas = 35          'No se usa
        otPasajes = 36
        otMapas = 38
        otBolsas = 39 ' Bolsas de Oro  (contienen más de 10k de oro)
        otPozos = 40 'Pozos Mágicos
        otEsposas = 41
        otRaíces = 42
        otCadáveres = 43
        otMonturas = 44
        otPuestos = 45 ' Puestos de Entrenamiento
        otNudillos = 46
        otAnillos = 47
        otCorreo = 48
        otruna = 49
        otCualquiera = 1000
    End Enum

    'Texto
    'Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
    'Public Const FONTTYPE_FIGHT As String = "~255~0~0~1~0"
    'Public Const FONTTYPE_WARNING As String = "~32~51~223~1~1"
    'Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
    'Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
    'Public Const FONTTYPE_EJECUCION As String = "~130~130~130~1~0"
    'Public Const FONTTYPE_Grupo As String = "~255~180~255~0~0"
    'Public Const FONTTYPE_VENENO As String = "~0~255~0~0~0"
    'Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
    'Public Const FONTTYPE_SERVER As String = "~0~185~0~0~0"
    'Public Const FONTTYPE_GUILDMSG As String = "~228~199~27~0~0"
    'Public Const FONTTYPE_CONSEJO As String = "~130~130~255~1~0"
    'Public Const FONTTYPE_CONSEJOCAOS As String = "~255~60~00~1~0"
    'Public Const FONTTYPE_CONSEJOVesA As String = "~0~200~255~1~0"
    'Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~50~0~1~0"
    'Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"

    'Estadisticas
    Public Const STAT_MAXELV As Byte = 50
    Public Const STAT_MAXHP As Integer = 999
    Public Const STAT_MAXSTA As Integer = 999
    Public Const STAT_MAXMAN As Integer = 9999
    Public Const STAT_MAXHIT_UNDER36 As Byte = 99
    Public Const STAT_MAXHIT_OVER36 As Integer = 999
    Public Const STAT_MAXDEF As Byte = 99



    ' **************************************************************
    ' **************************************************************
    ' ************************ TIPOS *******************************
    ' **************************************************************
    ' **************************************************************

    Public Structure tHechizo
        Dim Nombre As String
        Dim desc As String
        Dim PalabrasMagicas As String

        'Add Marius
        Dim ExclusivoClase As String
        '\Add

        Dim HechizeroMsg As String
        Dim TargetMsg As String
        Dim PropioMsg As String

        Dim HechizoDeArea As Byte
        Dim AreaEfecto As Byte
        Dim Afecta As Byte

        Dim tipo As TipoHechizo

        Dim WAV As Integer
        Dim FXgrh As Integer
        Dim loops As Byte
        Dim Particle As Integer

        Dim SubeHP As Byte
        Dim MinHP As Integer
        Dim MaxHP As Integer


        Dim SubeAgilidad As Byte
        Dim MinAgilidad As Integer
        Dim MaxAgilidad As Integer

        Dim SubeFuerza As Byte
        Dim MinFuerza As Integer
        Dim MaxFuerza As Integer

        Dim Invisibilidad As Byte
        Dim Paraliza As Byte
        Dim Inmoviliza As Byte
        Dim RemoverParalisis As Byte
        Dim CuraVeneno As Byte
        Dim Envenena As Byte
        Dim Incinera As Byte
        Dim Estupidez As Byte
        Dim Ceguera As Byte
        Dim Revivir As Byte
        Dim Resurreccion As Byte
        Dim ReviveFamiliar As Byte

        'Jose Castelli
        Dim CreaTelep As Byte
        Dim AutoLanzar As Byte
        'Jose Castelli

        'Mannakia
        Dim Desencantar As Byte
        Dim Sanacion As Byte
        Dim Certero As Byte

        Dim CreaAlgo As Byte
        Dim MinDef As Integer
        Dim MaxDef As Integer

        Dim MinHit As Byte
        Dim MaxHit As Byte
        'Mannakia

        'By jose castelli // Metamorfosis
        Dim Metamorfosis As Byte
        Dim MetaObj As Integer 'asusta no se usaba, lo use para saber el objeto que se necesita para transfomarse
        Dim Extrahit As Byte
        Dim Extradef As Byte
        Dim body As Integer
        Dim Head As Integer
        'By jose castelli  // Metamorfosis

        Dim Mimetiza As Byte
        Dim RemueveInvisibilidadParcial As Byte

        Dim Invoca As Byte
        Dim NumNpc As Integer
        Dim Cant As Integer

        Dim MinSkill As Integer
        Dim ManaRequerido As Integer

        Dim StaRequerido As Integer

        Dim Target As TargetType

        Dim Anillo As Byte
        '1 Espectral
        '2 Penumbra
    End Structure


    Public Structure LevelSkill
        Dim LevelValue As Integer
    End Structure

    Public Structure UserObj
        Dim ObjIndex As Integer
        Dim Amount As Integer
        Dim Equipped As Byte
        Dim Prob As Byte
    End Structure

    Public Structure Inventario
        Dim Objeto() As UserObj '1 To MAX_INVENTORY_SLOTS
        Dim WeaponEqpObjIndex As Integer
        Dim WeaponEqpSlot As Byte
        Dim NudiEqpSlot As Byte
        Dim NudiEqpIndex As Integer
        Dim ArmourEqpObjIndex As Integer
        Dim ArmourEqpSlot As Byte
        Dim EscudoEqpObjIndex As Integer
        Dim EscudoEqpSlot As Byte
        Dim CascoEqpObjIndex As Integer
        Dim CascoEqpSlot As Byte
        Dim MunicionEqpObjIndex As Integer
        Dim MunicionEqpSlot As Byte
        Dim MonturaObjIndex As Integer
        Dim MonturaSlot As Byte
        Dim AnilloEqpObjIndex As Integer
        Dim AnilloEqpSlot As Byte
        Dim BarcoObjIndex As Integer
        Dim BarcoSlot As Byte
        Dim MagicIndex As Integer
        Dim MagicSlot As Integer
        Dim NroItems As Integer
    End Structure

    Public Structure indexPosition
        Dim index As Long
        Dim x As Integer
        Dim y As Integer
    End Structure

    Public Structure tGrupoData
        Dim PIndex As Integer
        Dim RemXP As Double 'La exp. en el server se cuenta con Doubles
        Dim TargetUser As Integer 'Para las invitaciones
    End Structure

    Public Structure Position
        Dim x As Integer
        Dim Y As Integer
    End Structure

    Public Structure WorldPos
        Dim map As Integer
        Dim x As Integer
        Dim Y As Integer
    End Structure

    Public Structure FXdata
        Dim Nombre As String
        Dim GrhIndex As Integer
        Dim Delay As Integer
    End Structure

    Public Enum eMagicType
        ResistenciaMagica = 1
        ModificaAtributo = 2
        ModificaSkill = 3
        AceleraVida = 4
        AceleraMana = 5
        AumentaGolpe = 6
        DisminuyeGolpe = 7
        Nada = 8
        MagicasNoAtacan = 9
        Incinera = 10
        Paraliza = 11
        CarroMinerales = 12
        CaminaOculto = 13
        DañoMagico = 14
        Sacrificio = 15
        Silencio = 16
        NadieDetecta = 17
        Experto = 18
        Envenena = 19
    End Enum

    'Datos de user o npc
    Public Structure Cuerpo
        Dim CharIndex As Integer
        Dim Head As Integer
        Dim body As Integer

        Dim WeaponAnim As Integer
        Dim ShieldAnim As Integer
        Dim CascoAnim As Integer

        Dim fx As Integer
        Dim loops As Integer

        Dim heading As eHeading
    End Structure

    'Tipos de objetos
    Public Structure ObjData
        Dim Name As String 'Nombre del obj

        Dim OBJType As eOBJType 'Tipo enum que determina cuales son las caract del obj

        Dim GrhIndex As Integer ' Indice del grafico que representa el obj
        Dim GrhSecundario As Integer

        'Solo contenedores
        Dim MaxItems As Integer
        Dim Conte As Inventario
        Dim Apuñala As Byte

        Dim HechizoIndex As Integer

        Dim ForoID As String

        Dim MinHP As Integer ' Minimo puntos de vida
        Dim MaxHP As Integer ' Maximo puntos de vida

        Dim SubTipo As Byte

        Dim MineralIndex As Integer
        Dim LingoteInex As Integer


        Dim proyectil As Integer
        Dim Municion As Integer

        Dim Crucial As Byte
        Dim Newbie As Integer

        Dim DesdeMap As Long
        Dim HastaMap As Long
        Dim HastaY As Byte
        Dim HastaX As Byte
        Dim CantidadSkill As Byte

        'Pociones
        Dim TipoPocion As Byte
        Dim MaxModificador As Integer
        Dim MinModificador As Integer
        Dim DuracionEfecto As Long
        Dim MinSkill As Integer
        Dim LingoteIndex As Integer

        Dim MinHit As Integer 'Minimo golpe
        Dim MaxHit As Integer 'Maximo golpe

        Dim MinHAM As Integer
        Dim MinSed As Integer

        Dim def As Integer
        Dim MinDef As Integer ' Armaduras
        Dim MaxDef As Integer ' Armaduras

        Dim Ropaje As Integer 'Indice del grafico del ropaje

        Dim WeaponAnim As Integer ' Apunta a una anim de armas
        Dim ShieldAnim As Integer ' Apunta a una anim de escudo
        Dim CascoAnim As Integer

        Dim valor As Long     ' Precio

        Dim Cerrada As Integer
        Dim Llave As Byte
        Dim clave As Long 'si clave=llave la puerta se abre o cierra

        Dim IndexAbierta As Integer
        Dim IndexCerrada As Integer
        Dim IndexCerradaLlave As Integer

        Dim RazaTipo As Byte '1 Altas 2 Bajas 3 Orcas
        Dim RazaEnana As Byte
        Dim MinELV As Byte

        Dim QueAtributo As Byte
        Dim EfectoMagico As eMagicType
        Dim CuantoAumento As Byte
        Dim QueSkill As Byte

        'Add Marius
        Dim Shop As Byte
        '\Add

        Dim CuantoAgrega As Integer ' Para los contenedores

        Dim Mujer As Byte
        Dim Hombre As Byte

        Dim Agarrable As Byte

        Dim LingH As Integer
        Dim LingO As Integer
        Dim LingP As Integer
        Dim Madera As Integer

        Dim SkPociones As Integer
        Dim raies As Integer
        Dim PielLobo As Integer
        Dim PielOso As Integer
        Dim PielLoboInvernal As Integer

        Dim SkSastreria As Integer
        Dim SkHerreria As Integer
        Dim SkCarpinteria As Integer

        Dim texto As String

        'Clases que no tienen permitido usar este obj

        Dim ClaseProhibida() As eClass '1 To NUMCLASES

        Dim ClaseTipo As Byte

        Dim Snd1 As Integer
        Dim Snd2 As Integer
        Dim Snd3 As Integer

        Dim Real As Integer
        Dim Caos As Integer
        Dim Milicia As Integer

        Dim DefensaMagicaMax As Integer
        Dim DefensaMagicaMin As Integer

        Dim CPO As String


        Dim Refuerzo As Byte
        Dim ResistenciaMagica As Integer

        Dim Log As Byte 'es un objeto que queremos loguear? Pablo (ToxicWaste) 07/09/07
        Dim NoLog As Byte 'es un objeto que esta prohibido loguear?

        Dim DosManos As Byte
    End Structure

    Public Structure obj
        Dim ObjIndex As Integer
        Dim Amount As Integer
    End Structure


    Public Structure ModClase
        Dim Evasion As Double
        Dim AtaqueArmas As Double
        Dim AtaqueProyectiles As Double
        Dim DañoArmas As Double
        Dim DañoProyectiles As Double
        Dim DañoWrestling As Double
        Dim Escudo As Double
    End Structure

    Public Structure ModRaza
        Dim Fuerza As Single
        Dim Agilidad As Single
        Dim Inteligencia As Single
        Dim Carisma As Single
        Dim constitucion As Single
    End Structure
    '[/Pablo ToxicWaste]

    '[KEVIN]
    'Banco Objs
    Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
    '[/KEVIN]

    '[KEVIN]
    Public Structure BancoInventario
        Dim Objeto() As UserObj ' 1 To MAX_BANCOINVENTORY_SLOTS
        Dim NroItems As Integer
    End Structure
    '[/KEVIN]


    '*********************************************************
    '*********************************************************
    '*********************************************************
    '*********************************************************
    '******* T I P O S   D E    U S U A R I O S **************
    '*********************************************************
    '*********************************************************
    '*********************************************************
    '*********************************************************

    'Estadisticas de los usuarios
    Public Structure UserStats
        Dim GLD As Long 'Dinero
        Dim Banco As Long

        Dim MaxHP As Integer
        Dim MinHP As Integer

        Dim MaxSTA As Integer
        Dim MinSTA As Integer
        Dim MaxMAN As Integer
        Dim MinMAN As Integer
        Dim MaxHit As Integer
        Dim MinHit As Integer

        Dim MaxHAM As Integer
        Dim MinHAM As Integer

        Dim MaxAGU As Integer
        Dim MinAGU As Integer

        Dim def As Integer
        Dim Exp As Long
        Dim ELV As Byte
        Dim ELU As Long
        Dim UserSkills() As Byte ' 1 To NUMSKILLS
        Dim UserAtributos() As Byte ' 1 To NUMATRIBUTOS
        Dim UserAtributosBackUP() As Byte ' 1 To NUMATRIBUTOS
        Dim UserHechizos() As Integer ' 1 To MAXUSERHECHIZOS

        Dim eMinDef As Byte
        Dim eMaxDef As Byte
        Dim eMinHit As Byte
        Dim eMaxHit As Byte
        Dim eCreateTipe As Byte

        Dim dMaxDef As Integer
        Dim dMinDef As Integer

        Dim NPCsMuertos As Integer

        Dim VecesMuertos As Long

        Dim SkillPts As Integer

        Dim PuedeStaff As Byte
    End Structure

    'Flags
    Public Structure UserFlags

        Dim NoFalla As Byte

        'Castelli Casamientos
        Dim toyCasado As Byte
        Dim yaOfreci As Byte
        Dim miPareja As String
        'Castelli Casamientos

        'Mannakia Duelos 1vs1
        Dim vicDuelo As Integer 'Victima
        Dim inDuelo As Byte 'ESTA DUELEANDO ?
        Dim solDuelo As Integer 'Solicito ?

        'Sistema de Grupo
        Dim Solicito As Integer
        Dim Invito As Integer
        'Mannakia

        ' Castelli Metamorfossis
        Dim Metamorfosis As Byte
        ' Castelli Metamorfossis

        Dim DondeTiroMap As Integer
        Dim DondeTiroX As Integer
        Dim DondeTiroY As Integer
        Dim TiroPortalL As Integer

        Dim Muerto As Byte '¿Esta muerto?
        Dim Escondido As Byte '¿Esta escondido?
        Dim Comerciando As Boolean '¿Esta comerciando?
        Dim UserLogged As Boolean '¿Esta online?
        Dim accountlogged As Boolean ' ¿Cuenta online?
        Dim Meditando As Boolean
        Dim ModoCombate As Boolean

        Dim Hambre As Byte
        Dim Sed As Byte

        Dim Entrenando As Byte

        Dim Resucitando As Byte
        Dim Envenenado As Byte
        Dim Incinerado As Byte
        Dim Paralizado As Byte
        Dim Inmovilizado As Byte
        Dim Estupidez As Byte
        Dim Ceguera As Byte
        Dim Invisible As Byte
        Dim Oculto As Byte
        Dim Desnudo As Byte
        Dim Descansar As Boolean
        Dim Hechizo As Integer
        Dim TomoPocion As Boolean
        Dim TipoPocion As Byte

        Dim Trabajando As Boolean
        Dim Lingoteando As Byte

        Dim Navegando As Byte
        Dim Montando As Byte

        Dim Seguro As Boolean

        Dim DuracionEfecto As Long
        Dim TargetNPC As Integer ' Npc señalado por el usuario
        Dim TargetNpcTipo As eNPCType ' Tipo del npc señalado
        Dim NpcInv As Integer

        Dim ban As Byte

        Dim TargetUser As Integer ' Usuario señalado

        Dim TargetObj As Integer ' Obj señalado
        Dim TargetObjMap As Integer
        Dim TargetObjX As Integer
        Dim TargetObjY As Integer

        Dim TargetMap As Integer
        Dim TargetX As Integer
        Dim TargetY As Integer

        Dim TargetObjInvIndex As Integer
        Dim TargetObjInvSlot As Integer

        Dim AtacadoPorNpc As Integer
        Dim AtacadoPorUser As Integer
        Dim NPCAtacado As Integer
        Dim Privilegios As PlayerType

        Dim OldBody As Integer
        Dim OldHead As Integer
        Dim AdminInvisible As Byte
        Dim AdminPerseguible As Boolean

        Dim TimesWalk As Long

        Dim UltimoMensaje As Byte

        Dim Silenciado As Byte

    End Structure

    Public Structure UserCounters
        Dim Silenciado As Integer
        Dim Habla As Integer

        Dim CreoTeleport As Boolean
        Dim TimeTeleport As Integer

        Dim Metamorfosis As Integer

        Dim IdleCount As Long
        Dim AttackCounter As Integer
        Dim HPCounter As Integer
        Dim STACounter As Integer
        Dim Frio As Integer
        Dim COMCounter As Integer
        Dim AGUACounter As Integer
        Dim Veneno As Integer
        Dim Fuego As Integer
        Dim Paralisis As Integer
        Dim Ceguera As Integer
        Dim Estupidez As Integer

        Dim Invisibilidad As Integer
        Dim TiempoOculto As Integer

        Dim PiqueteC As Long
        Dim Pena As Long

        Dim Saliendo As Boolean
        Dim salir As Integer

        Dim IntervaloRevive As Long

        Dim TimerLanzarSpell As Long
        Dim TimerPuedeAtacar As Long
        Dim TimerPuedeUsarArco As Long
        Dim TimerPuedeTrabajar As Long
        Dim TimerUsar As Long
        Dim TimerMove As Long
        Dim TimerGolpeMagia As Long
        Dim TimerGolpeUsar As Long

        Dim Trabajando As Long  ' Para el centinela
        Dim Ocultando As Long   ' Unico trabajo no revisado por el centinela

        Dim failedUsageAttempts As Long
    End Structure

    'Cosas faccionarias.
    Public Structure tFacciones
        Dim ArmadaReal As Byte
        Dim Ciudadano As Byte

        Dim FuerzasCaos As Byte
        Dim Renegado As Byte

        Dim Milicia As Byte
        Dim Republicano As Byte

        Dim CiudadanosMatados As Integer
        Dim RenegadosMatados As Integer
        Dim RepublicanosMatados As Integer

        Dim MilicianosMatados As Integer
        Dim ArmadaMatados As Integer
        Dim CaosMatados As Integer

        Dim Rango As Byte

    End Structure

    Public Enum eMascota
        Fuego = 1
        Tierra
        Agua
        Ely
        Fatuo

        Tigre
        Lobo
        Oso
        Ent
    End Enum

    Structure Mascota
        Dim TieneFamiliar As Byte
        Dim invocado As Boolean

        Dim Nombre As String

        Dim MinHP As Integer
        Dim MaxHP As Integer

        Dim ELV As Byte
        Dim ELU As Long
        Dim Exp As Long

        Dim tipo As eMascota

        Dim MinHit As Integer
        Dim MaxHit As Integer

        Dim gDesarma As Byte
        Dim gEntorpece As Byte
        Dim gEnseguece As Byte
        Dim gParaliza As Byte
        Dim gEnvenena As Byte


        Dim Curar As Byte
        Dim Desencanta As Byte
        Dim Descargas As Byte
        Dim Paraliza As Byte
        Dim Inmoviliza As Byte
        Dim Tormentas As Byte
        Dim Misil As Byte
        Dim DetecInvi As Byte

        Dim NpcIndex As Integer
    End Structure


    Structure tCorreo
        Dim De As String
        Dim Mensaje As String
        Dim Cantidad As Integer
        Dim Item As Integer
        Dim idmsj As Long
    End Structure

    'Tipo de los Usuarios
    Public Structure User
        Dim Redundance As Byte
        Dim Name As String
        Dim account As String
        Dim IndexAccount As Long
        Dim Indexpj As Long

        Dim donador As Boolean

        Dim showName As Boolean

        Dim Matados() As Integer ' 1 To MAX_BUFFER_KILLEDS
        Dim Matados_timer() As Long '1 To MAX_BUFFER_KILLEDS


        Dim bandera As Byte

        Dim evento As Byte

        Dim cVer As Byte
        Dim Correos() As tCorreo '1 To 20
        Dim cant_mensajes As Byte ' Obvio que tope en 20


        Dim cuerpo As Cuerpo
        Dim OrigChar As Cuerpo

        Dim desc As String ' Descripcion

        Dim masc As Mascota

        Dim Clase As eClass
        Dim Raza As eRaza
        Dim Genero As eGenero
        Dim email As String
        Dim Hogar As Byte

        Dim Invent As Inventario

        Dim Pos As WorldPos
        Dim AuxPos As WorldPos

        Dim ConnIDValida As Boolean
        Dim ConnID As Long 'ID

        Dim client As handleClinet

        Dim BancoInvent As BancoInventario

        Dim Counters As UserCounters

        Dim MascotasIndex() As Integer ' 1 To MAXMASCOTAS
        Dim MascotasType() As Integer '1 To MAXMASCOTAS
        Dim NroMascotas As Integer

        Dim Stats As UserStats
        Dim flags As UserFlags

        Dim faccion As tFacciones

        Dim ip As String

        Dim ComUsu As tCOmercioUsuario

        Dim GuildIndex As Integer   'puntero al array Public de guilds
        Dim FundandoGuildAlineacion As ALINEACION_GUILD     'esto esta aca hasta que se parchee el cliente y se pongan cadenas de datos distintas para cada alineacion
        Dim EscucheClan As Integer

        Dim GrupoIndex As Integer
        Dim GrupoSolicitud As Integer


        Dim AreasInfo As AreaInfo

        'Outgoing and incoming messages
        Dim outgoingData As clsByteQueue
        Dim incomingData As clsByteQueue
    End Structure


    '*********************************************************
    '*********************************************************
    '*********************************************************
    '*********************************************************
    '**  T I P O S   D E    N P C S **************************
    '*********************************************************
    '*********************************************************
    '*********************************************************
    '*********************************************************

    Public Structure NPCStats
        Dim Alineacion As Integer
        Dim MaxHP As Long
        Dim MinHP As Long
        Dim MaxHit As Integer
        Dim MinHit As Integer
        Dim def As Integer
        Dim defM As Integer
    End Structure

    Public Structure NpcCounters
        Dim Paralisis As Integer
        Dim TiempoExistencia As Long
    End Structure

    Public Structure NPCFlags
        Dim AfectaParalisis As Byte
        Dim Domable As Integer
        Dim respawn As Byte
        Dim NPCActive As Boolean '¿Esta vivo?
        Dim Follow As Boolean
        Dim faccion As Byte
        Dim AtacaDoble As Byte
        Dim LanzaSpells As Byte

        Dim ExpCount As Long

        Dim OldMovement As TipoAI
        Dim OldHostil As Byte

        Dim AguaValida As Byte
        Dim TierraInvalida As Byte

        Dim Sound As Integer
        Dim AttackedBy As Integer
        Dim AttackedFirstBy As String
        Dim BackUp As Byte
        Dim RespawnOrigPos As Byte

        Dim Envenenado As Byte
        Dim Incinerado As Byte
        Dim Paralizado As Byte
        Dim Inmovilizado As Byte
        Dim Invisible As Byte

        Dim Snd1 As Integer
        Dim Snd2 As Integer
        Dim Snd3 As Integer
    End Structure

    Public Structure tCriaturasEntrenador
        Dim NpcIndex As Integer
        Dim NpcName As String
        Dim tmpIndex As Integer
    End Structure

    Structure tBot
        Dim UpMana As Integer
        Dim ManaMax As Integer

        Dim UpVida As Integer
        Dim VidaMax As Integer

        Dim TargetUser As Integer
        Dim TargetNPC As Integer

        Dim IntervaloAtaque As Long
        Dim IntervaloHechizo As Long

        Dim RandomDire As Byte
    End Structure


    Public Structure npc
        Dim Name As String
        Dim cuerpo As Cuerpo 'Define como se vera
        Dim desc As String

        Dim NPCtype As eNPCType
        Dim Numero As Integer
        Dim faccion As Byte
        Dim InvReSpawn As Byte

        Dim Comercia As Integer
        Dim Target As Long
        Dim TargetNPC As Long
        Dim TipoItems As Integer

        Dim Veneno As Byte
        Dim Fuego As Byte

        Dim Pos As WorldPos 'Posicion
        Dim oldPos As WorldPos
        Dim Orig As WorldPos
        Dim StartPos As WorldPos
        Dim lastHeading As eHeading

        Dim SkillDomar As Integer

        Dim Movement As TipoAI
        Dim Attackable As Byte
        Dim Hostile As Byte
        Dim PoderAtaque As Long
        Dim PoderEvasion As Long

        Dim GiveEXP As Long
        Dim GiveGLD As Long

        Dim Stats As NPCStats
        Dim flags As NPCFlags
        Dim Contadores As NpcCounters

        Dim Invent As Inventario
        Dim CanAttack As Byte

        Dim NroExpresiones As Byte
        Dim Expresiones() As String ' le da vida ;)

        Dim NroSpells As Byte
        Dim Spells() As Integer  ' le da vida ;)

        '<<<<Entrenadores>>>>>
        Dim NroCriaturas As Integer
        Dim Criaturas() As tCriaturasEntrenador
        Dim MaestroUser As Integer
        Dim MaestroNpc As Integer
        Dim Mascotas As Integer

        Dim IsFamiliar As Boolean

        Dim AreasInfo As AreaInfo
    End Structure

    '**********************************************************
    '**********************************************************
    '******************** Tipos del mapa **********************
    '**********************************************************
    '**********************************************************
    'Tile
    Public Structure MapBlock
        Dim Blocked As Byte 'Deberia ser un Boolean solo usa 2 valores 0 o 1 :S
        Dim Graphic() As Integer ' 1 To 4
        Dim UserIndex As Integer
        Dim NpcIndex As Integer

        'Des Nod Kopfnickend No s eusa para nada :S
        'BotIndex As Integer

        Dim ObjInfo As obj
        Dim ObjEsFijo As Byte

        Dim TileExit As WorldPos
        Dim Trigger As eTrigger
    End Structure

    'Info del mapa
    Structure MapInfo
        Dim NumUsers As Integer
        Dim Music As String
        Dim Name As String
        Dim StartPos As WorldPos
        Dim MapVersion As Integer
        Dim Seguro As Byte
        Dim Pk As Boolean
        Dim MagiaSinEfecto As Byte
        Dim NoEncriptarMP As Byte
        Dim InviSinEfecto As Byte
        Dim ResuSinEfecto As Byte

        Dim Terreno As String
        Dim Zona As String
        Dim Restringir As String
        Dim BackUp As Byte
    End Structure

    '********** V A R I A B L E S     P U B L I C A S ***********

    Public ULTIMAVERSION As String

    Public ListaRazas(0 To NUMRAZAS + 1) As String
    Public SkillsNames(0 To NUMSKILLS + 1) As String
    Public ListaClases(0 To NUMCLASES + 1) As String
    Public ListaAtributos(0 To NUMATRIBUTOS + 1) As String


    Public RecordUsuarios As Integer

    '
    'Directorios
    '

    ''
    'Ruta base del server, en donde esta el "server.ini"
    Public IniPath As String


    ''
    'Ruta base para los archivos de mapas
    Public MapPath As String

    ''
    'Ruta base para los DATs
    Public DatPath As String

    ''
    'Bordes del mapa
    Public MinXBorder As Byte
    Public MaxXBorder As Byte
    Public MinYBorder As Byte
    Public MaxYBorder As Byte

    ''
    'Numero de usuarios actual
    Public NumUsers As Integer
    Public LastUser As Integer
    Public LastChar As Integer
    Public NumChars As Integer
    Public LastNPC As Integer
    Public NumNPCs As Integer
    Public NumFX As Integer
    Public NumMaps As Integer
    Public NumObjDatas As Integer
    Public NumeroHechizos As Integer
    'Comentado por falta de uso
    'Public AllowMultiLogins As Byte
    'Public IdleLimit As Integer
    Public MaxUsers As Integer
    Public minutos As String
    Public haciendoBK As Boolean
    Public PuedeCrearPersonajes As Integer
    Public ServerSoloGMs As Integer


    Public EnPausa As Boolean

    '*****************ARRAYS PUBLICOS*************************
    Public UserList() As User 'USUARIOS
    Public Npclist(0 To MAXNPCS) As npc 'NPCS
    Public MapData(NumMaps, XMaxMapSize, YMaxMapSize) As MapBlock
    Public MapInfoArr() As MapInfo
    Public Hechizos() As tHechizo
    Public CharList(0 To MAXCHARS) As Integer
    Public ObjDataArr() As ObjData
    Public fx() As FXdata
    Public SpawnList() As tCriaturasEntrenador
    Public LevelSkillArr(0 To 50) As LevelSkill
    Public ArmasHerrero() As Integer
    Public ArmadurasHerrero() As Integer
    Public ObjCarpintero() As Integer
    Public ObjDruida() As Integer
    Public ObjSastre() As Integer
    Public BanIps As New Collection
    Public Parties(0 To MAX_PARTIES) As clsGrupo
    Public ModClaseArr(0 To NUMCLASES) As ModClase
    Public ModRazaArr(0 To NUMRAZAS) As ModRaza
    Public ModVida(0 To NUMCLASES) As Double
    Public ModMana(0 To NUMCLASES) As Double
    Public DistribucionSemienteraVida(0 To 4 + 1) As Integer
    '*********************************************************

    Public Nix As WorldPos
    Public Ullathorpe As WorldPos
    Public Banderbill As WorldPos
    Public Arghal As WorldPos

    Public Lindos As WorldPos
    Public Illiandor As WorldPos
    Public Suramei As WorldPos

    Public Orac As WorldPos

    Public Rinkel As WorldPos
    Public Tiama As WorldPos
    Public NuevaEsperanza As WorldPos

    Public NuevaEsperanzapuerto As WorldPos

    Public Prision As WorldPos
    Public Libertad As WorldPos

    Public Ayuda As New cCola


    Public Declare Function GetTickCount Lib "kernel32" () As Long

    'Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
    Public Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

    'Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
    Public Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal dest As IntPtr, ByVal size As IntPtr) '(ByRef destination As Any, ByVal length As Long)

End Module
