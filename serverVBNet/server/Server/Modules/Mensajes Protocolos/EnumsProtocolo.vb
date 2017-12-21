Module EnumsProtocolo
    ''
    'When we have a list of strings, we use this to separate them and prevent
    'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
    Public Const SEPARATOR As String = vbNullChar

    ''
    'The last existing client packet id.
    Public Const LAST_CLIENT_PACKET_ID As Byte = 103

    ''
    'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
    'Specially usefull to create a message once and send it over to several clients.
    Public auxiliarBuffer As New clsByteQueue


    Public Enum Stat
        Incinerado = &H1
        Envenenado = &H2
        Comerciand = &H4
        Trabajando = &H8
        Transformado = &H10
        Ciego = &H20
        Inactivo = &H40
        Silenciado = &H80
    End Enum

    Public Enum StatEx
        Paralizado = &H1
        Inmovilizado = &H2
        Hombre = &H4
        Mujer = &H8
    End Enum

    Public Enum ServerPacketID
        Logged                  ' 0
        RemoveDialogs           ' 1
        RemoveCharDialog        ' 2
        NavigateToggle          ' 3
        EquitateToggle          ' 4
        Disconnect              ' 5
        CommerceEnd             ' 6
        BankEnd                 ' 7
        CommerceInit            ' 8
        BankInit                ' 9
        UserCommerceInit        ' 10
        UserCommerceEnd         ' 11
        ShowBlacksmithForm      ' 12
        ShowCarpenterForm       ' 13
        ShowAlquimiaForm        ' 14
        ShowSastreForm          ' 15
        NPCSwing                ' 16
        NPCKillUser             ' 17
        BlockedWithShieldUser   ' 18
        BlockedWithShieldOther  ' 19
        UserSwing               ' 20
        SafeModeOn              ' 21
        SafeModeOff             ' 22
        NobilityLost            ' 23
        CantUseWhileMeditating  ' 24
        UpdateSta               ' 25
        UpdateMana              ' 26
        UpdateHP                ' 27
        UpdateGold              ' 28
        UpdateExp               ' 29
        ChangeMap               ' 30
        posUpdate               ' 31
        NPCHitUser              ' 32
        UserHitNPC              ' 33
        UserAttackedSwing       ' 34
        UserHittedByUser        ' 35
        UserHittedUser          ' 36
        ChatOverHead            ' 37
        ConsoleMsg              ' 38
        GuildChat               ' 39
        ShowMessageBox          ' 40
        UserIndexInServer       ' 41
        UserCharIndexInServer   ' 42
        CharacterCreate         ' 43
        CharacterRemove         ' 44
        CharacterMove           ' 45
        ForceCharMove           ' 46
        CharacterChange         ' 47
        CharStatus              ' 48
        ObjectCreate            ' 49
        ObjectDelete            ' 50
        BlockPosition           ' 51
        PlayMidi                ' 52
        PlayWave                ' 53
        guildList               ' 54
        AreaChanged             ' 55
        PauseToggle             ' 56
        CreateFX                ' 57
        CreateFXMap             ' 58
        UpdateUserStats         ' 59
        WorkRequestTarget       ' 60
        ChangeInventorySlot     ' 61
        ChangeBankSlot          ' 62
        ChangeSpellSlot         ' 63
        atributes               ' 64
        BlacksmithWeapons       ' 65
        BlacksmithArmors        ' 66
        CarpenterObjects        ' 67
        SastreObjects           ' 68
        AlquimiaObjects         ' 69
        RestOK                  ' 70
        ErrorMsg                ' 71
        Blind                   ' 72
        Dumb                    ' 73
        ShowSignal              ' 74
        ChangeNPCInventorySlot  ' 75
        ShowGuildFundationForm  ' 76
        ParalizeOK              ' 77
        ShowUserRequest         ' 78
        TradeOK                 ' 79
        BankOK                  ' 80
        ChangeUserTradeSlot     ' 81
        UpdateHungerAndThirst   ' 82
        MiniStats               ' 83
        AddForumMsg             ' 84
        ShowForumForm           ' 85
        SetInvisible            ' 86
        MeditateToggle          ' 87
        BlindNoMore             ' 88
        DumbNoMore              ' 89
        SendSkills              ' 90
        TrainerCreatureList     ' 91
        Pong                    ' 92
        UpdateTagAndStatus      ' 93
        SpawnList               ' 94
        ShowSOSForm             ' 95
        ShowGMPanelForm         ' 96
        UserNameList            ' 97
        AddPJ                   ' 98
        ShowAccount             ' 99
        CharacterInfo           ' 100
        GuildLeaderInfo         ' 101
        GuildDetails            ' 102
        Fuerza                  ' 103
        Agilidad                ' 104
        Subasta                 ' 105
        ParticleCreate          ' 106
        CharParticleCreate      ' 107
        DestParticle            ' 108
        DestCharParticle        ' 109
        hora                    ' 110
        Grupo                   ' 111
        ShowGrupoForm           ' 112
        Messages                ' 113
        showCorreoForm          ' 114
        AddCorreoMsg            ' 115
        ShowFamiliarForm        ' 116
        CharMsgStatus           ' 117
        MensajeSigno            ' 118
        Disconnect2             ' 119
    End Enum

    Public Enum ClientPacketID
        ConnectAccount          '0
        CreateNewAccount        '1
        LoginExistingChar       '2
        LoginNewChar            '3
        Talk                    '4
        Whisper                 '5
        Walk                    '6
        RequestPositionUpdate   '7
        Attack                  '8
        PickUp                  '9
        CombatModeToggle        '10
        SafeToggle              '11
        RequestGuildLeaderInfo  '12
        RequestEstadisticas     '13
        CommerceEnd             '14
        UserCommerceEnd         '15
        BankEnd                 '16
        UserCommerceOk          '17
        UserCommerceReject      '18
        Drop                    '19
        CastSpell               '20
        LeftClick               '21
        DoubleClick             '22
        Work                    '23
        UseItem                 '24
        CraftBlacksmith         '25
        CraftCarpenter          '26
        Craftalquimia           '27
        CraftSastre             '28
        WorkLeftClick           '29
        CreateNewGuild          '30
        SpellInfo               '31
        EquipItem               '32
        ChangeHeading           '33
        ModifySkills            '34
        Train                   '35
        CommerceBuy             '36
        BankExtractItem         '37
        CommerceSell            '38
        BankDeposit             '39
        ForumPost               '40
        MoveSpell               '41
        ClanCodexUpdate         '42
        UserCommerceOffer       '43
        GuildRequestJoinerInfo  '44
        GuildNewWebsite         '45
        GuildAcceptNewMember    '46
        GuildRejectNewMember    '47
        GuildKickMember         '48
        GuildUpdateNews         '49
        GuildMemberInfo         '50
        GuildRequestMembership  '51
        GuildRequestDetails     '52
        online                  '53
        Quit                    '54
        GuildLeave              '55
        RequestAccountState     '56
        PetStand                '57
        PetFollow               '58
        TrainList               '59
        Rest                    '60
        Meditate                '61
        Resucitate              '62
        Heal                    '63
        Help                    '64
        CommerceStart           '65
        BankStart               '66
        Enlist                  '67
        Information             '68
        Reward                  '69
        UpTime                  '70
        GrupoLeave              '71
        GrupoKick               '72
        GuildMessage            '73
        GrupoMessage            '74
        CentinelReport          '75
        GuildOnline             '76
        RoleMasterRequest       '77
        GMRequest               '78
        bugReport               '79
        ChangeDescription       '80
        Gamble                  '81
        LeaveFaction            '82
        BankExtractGold         '83
        BankTransferGold        '84
        BankDepositGold         '85
        Denounce                '86
        GuildFundate            '87
        Ping                    '88
        Casamiento              '89
        Acepto                  '90
        Divorcio                '91
        MessagesGM              '92
        Subasta                 '93
        RequestGrupo            '94
        Duelo                   '95
        BorrarMensaje           '96
        ExtraerItem             '97
        EnviarMensaje           '98
        AdoptarMascota          '99
        DelClan                 '100
        ChatFaccion             '101
        DragAndDrop             '102
        Hogar                   '103
        Participar              '104
        Pena                    '105
        RequestStats            '106 /EST
        Friends                 '107 /FADD /FDEL /FLIST /FMSG
    End Enum

    Public Enum gMessages
        GMMessage               '/GMSG
        showName                '/SHOWNAME
        OnlineArmada            '/ONLINEREAL
        OnlineCaos              '/ONLINECAOS
        OnlineMilicia           '/ONLINEMILI
        GoNearby                '/IRCERCA
        comment                 '/REM
        serverTime              '/HORA
        Where                   '/DONDE
        CreaturesInMap          '/NENE
        WarpMeToTarget          '/TELEPLOC
        WarpChar                '/TELEP
        Silence                 '/SILENCIAR
        SOSShowList             '/SHOW SOS
        SOSRemove               'SOSDONE
        GoToChar                '/IRA
        Invisible               '/INVISIBLE
        GMPanel                 '/PANELGM
        RequestUserList         'LISTUSU
        Working                 '/TRABAJANDO
        Hiding                  '/OCULTANDO
        Jail                    '/CARCEL
        KillNPC                 '/RMATA
        WarnUser                '/ADVERTENCIA
        EditChar                '/MOD
        ReviveChar              '/REVIVIR
        OnlineGM                '/ONLINEGM
        OnlineMap               '/ONLINEMAP
        Kick                    '/ECHAR
        Execute                 '/EJECUTAR
        BanChar                 '/BAN
        UnbanChar               '/UNBAN
        NPCFollow               '/SEGUIR
        SummonChar              '/SUM
        SpawnListRequest        '/CC
        SpawnCreature           'SPA
        ResetNPCInventory       '/RESETINV
        CleanWorld              '/LIMPIAR
        ServerMessage           '/RMSG
        NickToIP                '/NICK2IP
        IPToNick                '/IP2NICK
        GuildOnlineMembers      '/ONCLAN
        TeleportCreate          '/CT
        TeleportDestroy         '/DT
        SetCharDescription      '/SETDESC
        ForceMIDIToMap          '/FORCEMIDIMAP
        ForceWAVEToMap          '/FORCEWAVMAP
        TalkAsNPC               '/TALKAS
        DestroyAllItemsInArea   '/MASSDEST
        ItemsInTheFloor         '/PISO
        MakeDumb                '/ESTUPIDO
        MakeDumbNoMore          '/NOESTUPIDO
        DumpIPTables            '/DUMPSECURITY
        SetTrigger              '/TRIGGER
        AskTrigger              '/TRIGGER with no args
        BannedIPList            '/BANIPLIST
        BannedIPReload          '/BANIPRELOAD
        GuildMemberList         '/MIEMBROSCLAN
        ShowGuildMessages       '/SHOWCMSG
        GuildBan                '/BANCLAN
        BanIP                   '/BANIP
        UnbanIP                 '/UNBANIP
        CreateItem              '/CI
        DestroyItems            '/DEST
        ChaosLegionKick         '/NOCAOS
        RoyalArmyKick           '/NOREAL
        MiliciaKick             '/NOMILI
        ForceMIDIAll            '/FORCEMIDI
        ForceWAVEAll            '/FORCEWAV
        TileBlockedToggle       '/BLOQ
        KillNPCNoRespawn        '/MATA
        KillAllNearbyNPCs       '/MASSKILL
        LastIP                  '/LASTIP
        SystemMessage           '/SMSG
        CreateNPC               '/ACC
        CreateNPCWithRespawn    '/RACC
        NavigateToggle          '/NAVE
        ServerOpenToUsersToggle '/HABILITAR
        TurnOffServer           '/APAGAR
        TurnCriminal            '/CONDEN
        ResetFactions           '/RAJAR
        RemoveCharFromGuild     '/RAJARCLAN
        ToggleCentinelActivated '/CENTINELAACTIVADO
        DoBackUp                '/DOBACKUP
        Ignored                 '/IGNORADO
        CheckSlot               '/SLOT
        KickAllChars            '/ECHARTODOSPJS
        ReloadNPCs              '/RELOADNPCS
        ReloadServerIni         '/RELOADSINI
        ReloadSpells            '/RELOADHECHIZOS
        ReloadObjects           '/RELOADOBJ
        Restart                 '/REINICIAR
        SaveMap                 '/GUARDAMAPA
        ChangeMapInfoPK         '/MODMAPINFO PK
        ChangeMapInfoBackup     '/MODMAPINFO BACKUP
        ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
        ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
        ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
        ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
        ChangeMapInfoLand       '/MODMAPINFO TERRENO
        ChangeMapInfoZone       '/MODMAPINFO ZONA
        SaveChars               '/GRABAR
        CleanSOS                '/BORRAR SOS
        CancelTorneo            '/CANCELTORNEO
        CrearTorneo             '/CREARTORNEO
        Pejotas                 '/PEJOTAS
        SlashSlash              '// <comando>
    End Enum

    Public Enum FontTypeNames
        FONTTYPE_TALK '255~255~255~0~0
        FONTTYPE_FIGHT '255~0~0~1~0
        FONTTYPE_WARNING '32~51~223~1~1
        FONTTYPE_INFO '65~190~156~0~0
        FONTTYPE_VENENO '0~255~0~0~0
        FONTTYPE_GUILD '255~255~255~1~0
        FONTTYPE_TALKITALIC '255~255~255~0~1
        FONTTYPE_SERVER '0~185~0~0~0
        FONTTYPE_CLAN '228~199~27~0~0
        FONTTYPE_RED '255~0~0~0~0
        FONTTYPE_BROWNB '204~193~115~1~0
        FONTTYPE_BROWNI '204~193~115~0~1
        FONTTYPE_PRIVADO '182~226~29~0~0
        FONTTYPE_Public '139~248~244~0~1
        FONTTYPE_GRUPO '0~128~128~0~0
        FONTTYPE_FACCION '228~199~27~0~0

        FONTTYPE_FACCION_IMPE '0~80~200~1~1
        FONTTYPE_FACCION_REPU '243~147~1~1~1
        FONTTYPE_FACCION_CAOS '197~0~5~1~1
    End Enum

    Public Enum eEditOptions
        eo_Gold = 1
        eo_Experience
        eo_Body
        eo_Head
        eo_CiticensKilled
        eo_CriminalsKilled
        eo_Level
        eo_Class
        eo_Skills
        eo_SkillPointsLeft
        eo_Nobleza
        eo_Asesino
        eo_Sex
        eo_Raza
        eo_Part
    End Enum
End Module
