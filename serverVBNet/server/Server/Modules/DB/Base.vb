Option Explicit On

Module Base
    Public DB_Conn As ADODB.Connection

    Public DB_User As String    'The database username - (default "root")
    Public DB_Pass As String    'Password to your database for the corresponding username
    Public DB_Name As String    'Name of the table in the database (default "vbgore")
    Public DB_Host As String    'IP of the database server - use localhost if hosted locally! Only host remotely for multiple servers!
    Public DB_Port As Integer   'Port of the database (default "3306")


    Public Sub CargarDB()
        Dim ErrorString As String
        Dim DB_RS As New ADODB.Recordset
        On Error GoTo Errhandler

        'Add Marius Me canse de anda cambiando las constantes
        DB_User = "root" 'Trim$(GetVar(IniPath & "Server.ini", "MYSQL", "User"))
        DB_Pass = "123456" 'Trim$(GetVar(IniPath & "Server.ini", "MYSQL", "Password"))
        DB_Name = "inmortalao" 'Trim$(GetVar(IniPath & "Server.ini", "MYSQL", "Database"))
        DB_Host = "localhost" 'Trim$(GetVar(IniPath & "Server.ini", "MYSQL", "Host"))
        DB_Port = "3306" 'Val(GetVar(IniPath & "Server.ini", "MYSQL", "Port"))
        '\Add

        DB_Conn = New ADODB.Connection
        DB_Conn.ConnectionString = "DRIVER={MySQL ODBC 5.3 ANSI Driver};" &
                "SERVER=" & DB_Host & ";" &
                "DATABASE=" & DB_Name & ";" &
                "PORT=" & DB_Port & ";" &
                "UID=" & DB_User & ";" &
                "PWD=" & DB_Pass & "; OPTION=3"

        DB_Conn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        DB_Conn.Open()

        'Ejecutamos estas sentencias para asegurarnos que las tablas esten!
        DB_RS.Open("SELECT * FROM charatrib WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charbanco WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charcorreo WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charfaccion WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charflags WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charguild WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charhechizos WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charinit WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charinvent WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charmascotafami WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charskills WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM charstats WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM cuentas WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()
        DB_RS.Open("SELECT * FROM extras WHERE 0=1", DB_Conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        DB_RS.Close()

        On Error GoTo 0

        Exit Sub

Errhandler:
        MsgBox(Err.Description)
        'End

        'Add Marius

        'Refresh the errors
        DB_Conn.Errors.Refresh()

        'Get the error string if there is one
        If DB_Conn.Errors.Count > 0 Then ErrorString = DB_Conn.Errors.Item(0).Description

        'Check for known errors
        If InStr(1, ErrorString, "Access denied for user ") Then
            MsgBox("Mysql: Usuario o Contraseña incorrecta. " & Err.Description & " ErrorString: " & ErrorString)
        ElseIf InStr(1, ErrorString, "Can't connect to MySQL server on ") Then
            MsgBox("Mysql: No se pudo conectar con el server, verifique el host y el puerto. " & Err.Description & " ErrorString: " & ErrorString)
        ElseIf InStr(1, ErrorString, "Unknown database ") Then
            MsgBox("Mysql: La base de datos no existe. " & Err.Description & " ErrorString: " & ErrorString)
        ElseIf InStr(1, ErrorString, "Data source name not found and no default driver specified") Then
            MsgBox("Mysql: La base de datos invalida o no existe. " & Err.Description & " ErrorString: " & ErrorString)
        ElseIf InStr(1, ErrorString, "Table '") & InStr(1, ErrorString, "' doesn't exist") Then
            MsgBox("Mysql: Alguna tabla no existe o tiene errores. " & Err.Description & " ErrorString: " & ErrorString)
        Else
            MsgBox("Mysql: Error conectando con la base de datos... " & Err.Description & " ErrorString: " & ErrorString)
        End If


        End

    End Sub


    Public Sub insertarObjsDB()


    End Sub


    Public Function getValorPropiedadDats(ByVal nombrePropiedad) As String
        Dim RS As New ADODB.Recordset

        RS = DB_Conn.Execute("SELECT * FROM `dats` WHERE nombre='" & nombrePropiedad & "'" & " LIMIT 1")
        Dim respuesta As String
        respuesta = (RS.Fields.Item("valor").Value).ToString()
        Return respuesta

    End Function

    Public Sub insertarNpcsDB()

        Dim Index As Integer
        Dim Name As Integer
        Dim desc As Integer
        Dim Movement As Integer
        Dim AguaValida As Integer
        Dim TierraInvalida As Integer
        Dim faccion As Integer
        Dim AtacaDoble As Integer
        Dim NPCtype As Integer
        Dim body As Integer
        Dim Head As Integer
        Dim heading As Integer
        Dim Attackable As Integer
        Dim Comercia As Integer
        Dim Hostile As Integer
        Dim GiveEXP As Integer
        Dim Veneno As Integer
        Dim Domable As Integer
        Dim GiveGLD As Integer
        Dim PoderAtaque As Integer
        Dim PoderEvasion As Integer
        Dim InvReSpawn As Integer
        Dim MaxHP As Integer
        Dim MinHP As Integer
        Dim MaxHit As Integer
        Dim MinHit As Integer
        Dim def As Integer
        Dim defM As Integer
        Dim Alineacion As Integer
        Dim NroItems As Integer
        Dim LanzaSpells As Integer
        Dim sp1 As Integer
        Dim sp2 As Integer
        Dim sp3 As Integer
        Dim sp4 As Integer
        Dim sp5 As Integer
        Dim NroCriaturas As Integer
        Dim ci1 As Integer
        Dim ci2 As Integer
        Dim ci3 As Integer
        Dim ci4 As Integer
        Dim ci5 As Integer
        Dim cn1 As Integer
        Dim cn2 As Integer
        Dim cn3 As Integer
        Dim cn4 As Integer
        Dim cn5 As Integer
        Dim respawn As Integer
        Dim BackUp As Integer
        Dim originpos As Integer
        Dim AfectaParalisis As Integer
        Dim Snd1 As Integer
        Dim Snd2 As Integer
        Dim Snd3 As Integer
        Dim nroexp As Integer
        Dim exp1 As Integer
        Dim exp2 As Integer
        Dim exp3 As Integer
        Dim TipoItems As Integer

        Dim objetosParams As String
        Dim objetosValues As String

        Dim spellsParams As String
        Dim spellsValues As String

        Dim ciParams As String
        Dim ciValues As String

        Dim cnParams As String
        Dim cnValues As String

        Dim expParams As String
        Dim expValues As String

        Dim loopC As Long
        Dim ln As String
        Dim aux As String


        Dim i As Integer

        On Error GoTo errh
        Dim result As Integer

        Dim consulta As String



        For i = i To 610

            objetosParams = ""
            objetosValues = ""
            spellsParams = ""
            spellsValues = ""
            ciParams = ""
            ciValues = ""
            cnParams = ""
            cnValues = ""
            expParams = ""
            expValues = ""
            consulta = ""

            result = OpenNPC(i)


            If result <> 0 Then

                With Npclist(result)


                    If i = 610 Then
                        'MsgBox.cuerpo.body
                    End If

                    If .Invent.NroItems > 0 Then
                        For loopC = 1 To .Invent.NroItems
                            objetosValues = objetosValues + "'" & .Invent.Objeto(loopC).ObjIndex & "-" & .Invent.Objeto(loopC).Amount & "-" & .Invent.Objeto(loopC).Prob & "',"
                            objetosParams = objetosParams + "`obj" & loopC & "`,"
                        Next loopC
                    End If

                    For loopC = 1 To .flags.LanzaSpells
                        spellsValues = spellsValues + "" & .Spells(loopC) & ","
                        spellsParams = spellsParams + "`sp" & loopC & "`,"
                    Next loopC

                    If .NPCtype = eNPCType.Entrenador Then
                        For loopC = 1 To .NroCriaturas

                            ciValues = ciValues + "" & .Criaturas(loopC).NpcIndex & ","
                            ciParams = ciParams + "`ci" & loopC & "`,"

                            cnValues = cnValues + "'" + .Criaturas(loopC).NpcName + "',"
                            cnParams = cnParams + "`cn" & loopC & "`,"
                        Next loopC
                    End If

                    For loopC = 1 To .NroExpresiones
                        expValues = expValues + .Expresiones(loopC) & ","
                        expParams = expParams + "`exp" & loopC & "`,"
                    Next loopC


                    consulta = "INSERT INTO `inmortalao`.`npcs`(`id`,`index`,`name`,`desc`,`movement`,`aguavalida`,`tierrainvalida`,`faccion`,`atacadoble`,`npctype`,`body`,`head`,`heading`,`Attackable`,`comercia`,`hostile`,`giveexp`,`veneno`,`domable`,`givegld`,`poderataque`,`poderevasion`,`invrespawn`,`maxhp`,`minhp`,`maxhit`,`minhit`,`def`,`defm`,`alineacion`,`nroitems`," + objetosParams + "`lanzaspells`," + spellsParams + "`nrocriaturas`," + ciParams + cnParams + "`respawn`,`backup`,`originpos`,`afectaparalisis`,`snd1`,`snd2`,`snd3`,`nroexp`," + expParams + "`tipoitems`) values" &
                "(null," & i & ",'" + .Name + "','" + .desc + "'," & .Movement & "," & .flags.AguaValida & "," & .flags.TierraInvalida & "," &
                "" & .flags.faccion & "," & .flags.AtacaDoble & "," & .NPCtype & "," & .cuerpo.body & "," & .cuerpo.Head & "," & .cuerpo.heading & "," & .Attackable & "," &
                "" & .Comercia & "," & .Hostile & "," & .GiveEXP & "," & .Veneno & "," & .flags.Domable & "," & .GiveGLD & "," & .PoderAtaque & "," &
                "" & .PoderEvasion & "," & .InvReSpawn & "," & .Stats.MaxHP & "," & .Stats.MinHP & "," & .Stats.MaxHit & "," & .Stats.MinHit & "," & .Stats.def & "," &
                "" & .Stats.defM & "," & .Stats.Alineacion & "," & .Invent.NroItems & "," + objetosValues &
                "" & .flags.LanzaSpells & "," + spellsValues & .NroCriaturas & "," &
                "" + ciValues + cnValues &
                "" & .flags.respawn & "," & .flags.BackUp & "," & .flags.RespawnOrigPos & "," & .flags.AfectaParalisis & "," & .flags.Snd1 & "," & .flags.Snd2 & "," & .flags.Snd3 & "," &
                "" & .NroExpresiones & "," + expValues & .TipoItems & ");"

                    DB_Conn.Execute(consulta)
                End With
            End If
        Next i

        Exit Sub

errh:
        MsgBox(consulta)
        MsgBox(Err.Description)


    End Sub

    Public Function getValorPropiedadBalance(ByVal nombrePropiedad) As Double
        Dim RS As New ADODB.Recordset
        RS = DB_Conn.Execute("SELECT * FROM `balance` WHERE nombre='" & nombrePropiedad & "'" & " LIMIT 1")

        Return Convert.ToDouble(RS.Fields.Item("valor").Value)

    End Function


    Public Function ChangeBan(ByVal Name As String, ByVal Baneado As Byte) As Boolean
        Dim Orden As String
        Dim tUser As Integer
        Dim RS As New ADODB.Recordset

        ChangeBan = False
        RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'" & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function

        Orden = "UPDATE `charflags` SET "
        Orden = Orden & "Ban=" & CStr(Baneado)
        Orden = Orden & " WHERE IndexPJ=" & RS.Fields.Item("Indexpj").Value.ToString() & " LIMIT 1"

        Call DB_Conn.Execute(Orden)

        tUser = NameIndex(Name)
        If tUser > 0 Then
            Call CloseSocket(tUser)
        End If

        RS = Nothing

        ChangeBan = True
    End Function
    'Add Marius
    Public Function Pejotas(ByVal Name As String) As String
        Dim Orden As String
        Dim tUser As Integer
        Dim RS As New ADODB.Recordset

        Pejotas = ""
        RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'" & " LIMIT 1")
        If RS.BOF Or RS.EOF Then
            Pejotas = "El personaje no existe!"
            Exit Function
        End If

        'RS!Indexpj
        RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE id=" & RS.Fields.Item("id").Value.ToString() & " LIMIT 10")

        If Not (RS.BOF Or RS.EOF) Then

            Dim ii As Byte
            For ii = 1 To RS.RecordCount
                Pejotas = Pejotas & RS.Fields.Item("Nombre").Value.ToString()

                If RS.Fields.Item("Online").Value.ToString() Then Pejotas = Pejotas & " (Online)"

                Pejotas = Pejotas & " | "

                RS.MoveNext()
            Next ii
            If Len(Pejotas) <> 0 Then Pejotas = Left$(Pejotas, Len(Pejotas) - 3)
        Else
            Pejotas = "Error al mostrar personajes"
        End If


        RS = Nothing
    End Function
    '\Add


    Public Sub CerrarDB()
        On Error GoTo ErrHandle
        DB_Conn.Close()
        DB_Conn = Nothing
        Exit Sub
ErrHandle:
        Call LogError("CerrarDB " & Err.Description & " " & Err.Number)
        End

    End Sub
    Public Sub SaveUserSQL(UserIndex As Integer, Optional ByVal account As String = "", Optional insertPj As Boolean = False)

        Dim ipj As Integer

        If Not account = "" Then
            Call AddUserInAccount(UserList(UserIndex).IndexAccount)
        End If

        If insertPj Then
            ipj = Insert_New_Table(UserList(UserIndex).Name, UserList(UserIndex).IndexAccount)
        Else
            ipj = GetIndexPJ(UserList(UserIndex).Name)
        End If

        SaveUserFlags(UserIndex, ipj)
        SaveUserStats(UserIndex, ipj)
        SaveUserInit(UserIndex, ipj)
        SaveUserInv(UserIndex, ipj)
        SaveUserBank(UserIndex, ipj)
        SaveUserHechi(UserIndex, ipj)
        SaveUserAtrib(UserIndex, ipj)
        SaveUserSkill(UserIndex, ipj)
        SaveUserFami(UserIndex, ipj)
        SaveUserFaccion(UserIndex, ipj)

        Exit Sub

    End Sub

    Sub SaveUserHechi(ByVal UserIndex As Integer, ByVal ipj As Integer)
        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim mUser As User
        Dim i As Byte
        Dim str As String

        mUser = UserList(UserIndex)

        If Len(mUser.Name) = 0 Then Exit Sub

        '************************************************************************
        RS = DB_Conn.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        str = "UPDATE `charhechizos` SET"
        str = str & " IndexPJ=" & ipj

        ReDim Preserve mUser.Stats.UserHechizos(MAXUSERHECHIZOS + 1)

        For i = 1 To MAXUSERHECHIZOS
            str = str & ",H" & i & "=" & mUser.Stats.UserHechizos(i)
        Next i
        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        '************************************************************************

        Exit Sub
ErrHandle:
        Resume Next
    End Sub


    Sub SaveUserFami(ByVal UserIndex As Integer, ByVal ipj As Integer)
        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim mUser As User
        Dim i As Byte
        Dim str As String

        mUser = UserList(UserIndex)

        If Len(mUser.Name) = 0 Then Exit Sub

        '************************************************************************
        RS = DB_Conn.Execute("SELECT * FROM `charmascotafami` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        str = "UPDATE `charmascotafami` SET"
        str = str & " IndexPJ=" & ipj
        str = str & ",Tiene=" & UserList(UserIndex).masc.TieneFamiliar
        str = str & ",Nombre='" & UserList(UserIndex).masc.Nombre & "'"
        str = str & ",Tipo=" & UserList(UserIndex).masc.tipo
        str = str & ",Level=" & UserList(UserIndex).masc.ELV
        str = str & ",ELU=" & UserList(UserIndex).masc.ELU
        str = str & ",Exp=" & UserList(UserIndex).masc.Exp
        str = str & ",MinHP=" & UserList(UserIndex).masc.MinHP
        str = str & ",MaxHP=" & UserList(UserIndex).masc.MaxHP
        str = str & ",MinHIT=" & UserList(UserIndex).masc.MinHit
        str = str & ",MaxHIT=" & UserList(UserIndex).masc.MaxHit
        str = str & ",MAS1=0"
        str = str & ",MAS2=0"
        str = str & ",MAS3=0"
        str = str & ",NroMascotas=0"

        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        '************************************************************************
        Exit Sub
ErrHandle:
        Resume Next
    End Sub
    Sub SaveUserFlags(ByVal UserIndex As Integer, ByVal ipj As Integer)
        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim i As Byte
        Dim str As String

        If Len(UserList(UserIndex).Name) = 0 Then Exit Sub

        '************************************************************************
        RS = New ADODB.Recordset

        RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & ipj & " LIMIT 1")
        str = "UPDATE `charflags` SET "
        'str = str & " IndexPJ=" & ipj
        str = str & "Nombre='" & UserList(UserIndex).Name & "'"
        str = str & ",Navegando=" & UserList(UserIndex).flags.Navegando
        str = str & ",Envenenado=" & UserList(UserIndex).flags.Envenenado
        'Mod Marius Ahora funcionan xD
        str = str & ",Pena=" & UserList(UserIndex).Counters.Pena
        str = str & ",Paralizado=0"
        '\Mod
        str = str & ",Desnudo=" & UserList(UserIndex).flags.Desnudo
        str = str & ",Sed=" & UserList(UserIndex).flags.Sed
        str = str & ",Hambre=" & UserList(UserIndex).flags.Hambre
        str = str & ",Escondido=" & UserList(UserIndex).flags.Escondido
        str = str & ",Muerto=" & UserList(UserIndex).flags.Muerto
        'Add Nod Kopfnickend
        str = str & ",`desc`='" & UserList(UserIndex).desc & "'"
        '\Add
        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)

        'Grabamos Estados
        '************************************************************************
        Exit Sub
ErrHandle:
        Resume Next
    End Sub


    Sub SaveUserFaccion(ByVal UserIndex As Integer, ByVal ipj As Integer)
        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim mUser As User
        Dim i As Byte
        Dim str As String

        mUser = UserList(UserIndex)

        If Len(mUser.Name) = 0 Then Exit Sub

        '************************************************************************
        RS = DB_Conn.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        str = "UPDATE `charfaccion` SET"

        'Graba Faccion
        str = str & " IndexPJ=" & ipj
        str = str & ",EjercitoReal=" & mUser.faccion.ArmadaReal
        str = str & ",EjercitoCaos=" & mUser.faccion.FuerzasCaos
        str = str & ",EjercitoMili=" & mUser.faccion.Milicia
        str = str & ",Republicano=" & mUser.faccion.Republicano
        str = str & ",Ciudadano=" & mUser.faccion.Ciudadano
        str = str & ",Rango=" & mUser.faccion.Rango
        str = str & ",Renegado=" & mUser.faccion.Renegado
        str = str & ",CiudMatados=" & mUser.faccion.CiudadanosMatados
        str = str & ",ReneMatados=" & mUser.faccion.RenegadosMatados
        str = str & ",RepuMatados=" & mUser.faccion.RepublicanosMatados
        str = str & ",CaosMatados=" & mUser.faccion.CaosMatados
        str = str & ",ArmadaMatados=" & mUser.faccion.ArmadaMatados
        str = str & ",MiliMatados=" & mUser.faccion.MilicianosMatados
        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        '************************************************************************
        Exit Sub
ErrHandle:
        Resume Next
    End Sub
    Sub SaveUserInit(ByVal UserIndex As Integer, ByVal ipj As Integer)
        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim mUser As User
        Dim i As Byte
        Dim str As String
        Dim time As Long

        mUser = UserList(UserIndex)

        If Len(mUser.Name) = 0 Then Exit Sub


        '************************************************************************
        RS = DB_Conn.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        str = "UPDATE `charinit` SET"
        str = str & " IndexPJ=" & ipj
        str = str & ",Genero=" & mUser.Genero
        str = str & ",Raza=" & mUser.Raza
        str = str & ",Hogar=" & mUser.Hogar
        str = str & ",Clase=" & mUser.Clase
        str = str & ",Heading=" & mUser.cuerpo.heading
        str = str & ",Head=" & mUser.OrigChar.Head
        str = str & ",Body=" & mUser.cuerpo.body
        str = str & ",Arma=" & mUser.cuerpo.WeaponAnim
        str = str & ",Escudo=" & mUser.cuerpo.ShieldAnim
        str = str & ",Casco=" & mUser.cuerpo.CascoAnim
        str = str & ",LastIP='" & mUser.ip & "'"
        str = str & ",Mapa=" & mUser.Pos.map
        str = str & ",X=" & mUser.Pos.x
        str = str & ",Y=" & mUser.Pos.Y
        str = str & ",PAREJA='" & mUser.flags.miPareja & "'"


        ReDim mUser.Matados_timer(MAX_BUFFER_KILLEDS + 1)
        ReDim mUser.Matados(MAX_BUFFER_KILLEDS + 1)

        For i = 1 To MAX_BUFFER_KILLEDS
            time = mUser.Matados_timer(i) - GetTickCount
            If time > 0 And mUser.Matados(i) <> 0 Then
                str = str & ",pj" & i & "=" & mUser.Matados(i)
                str = str & ",time" & i & "=" & time
            Else
                str = str & ",pj" & i & "= 0"
                str = str & ",time" & i & "= 0"
            End If
        Next i


        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        '************************************************************************
        Exit Sub
ErrHandle:
        Resume Next
    End Sub
    Sub SaveUserInv(ByVal UserIndex As Integer, ByVal ipj As Integer)
        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim mUser As User
        Dim i As Byte
        Dim str As String

        mUser = UserList(UserIndex)

        If Len(mUser.Name) = 0 Then Exit Sub

        'Fix Marius
        mUser.Invent.MonturaSlot = 0
        '\Fix

        '************************************************************************
        RS = DB_Conn.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        str = "UPDATE `charinvent` SET"
        str = str & " IndexPJ=" & ipj
        For i = 1 To MAX_INVENTORY_SLOTS
            str = str & ",OBJ" & i & "=" & mUser.Invent.Objeto(i).ObjIndex
            str = str & ",CANT" & i & "=" & mUser.Invent.Objeto(i).Amount
        Next i
        str = str & ",CASCOSLOT=" & mUser.Invent.CascoEqpSlot
        str = str & ",ARMORSLOT=" & mUser.Invent.ArmourEqpSlot
        str = str & ",SHIELDSLOT=" & mUser.Invent.EscudoEqpSlot
        str = str & ",WEAPONSLOT=" & mUser.Invent.WeaponEqpSlot
        str = str & ",ANILLOSLOT=" & mUser.Invent.AnilloEqpSlot
        str = str & ",MUNICIONSLOT=" & mUser.Invent.MunicionEqpSlot
        str = str & ",BARCOSLOT=" & mUser.Invent.BarcoSlot
        str = str & ",NUDISLOT=" & mUser.Invent.NudiEqpSlot
        str = str & ",MONTUSLOT=" & mUser.Invent.MonturaSlot

        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        '************************************************************************
        Exit Sub

ErrHandle:
        Resume Next

    End Sub
    Sub SaveUserBank(ByVal UserIndex As Integer, ByVal ipj As Integer)
        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim mUser As User
        Dim i As Byte
        Dim str As String

        mUser = UserList(UserIndex)

        If Len(mUser.Name) = 0 Then Exit Sub


        '************************************************************************
        RS = DB_Conn.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        str = "UPDATE `charbanco` SET"
        str = str & " IndexPJ=" & ipj

        ReDim Preserve mUser.BancoInvent.Objeto(MAX_BANCOINVENTORY_SLOTS + 1)

        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            str = str & ",OBJ" & i & "=" & mUser.BancoInvent.Objeto(i).ObjIndex
            str = str & ",CANT" & i & "=" & mUser.BancoInvent.Objeto(i).Amount
        Next i
        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        '************************************************************************
        Exit Sub
ErrHandle:
        Resume Next
    End Sub
    Sub SaveUserStats(ByVal UserIndex As Integer, ByVal ipj As Integer)
        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim mUser As User
        Dim i As Byte
        Dim str As String

        mUser = UserList(UserIndex)

        If Len(mUser.Name) = 0 Then Exit Sub

        '************************************************************************
        RS = DB_Conn.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        str = "UPDATE `charstats` SET"
        str = str & " IndexPJ=" & ipj
        str = str & ",GLD=" & mUser.Stats.GLD
        str = str & ",BANCO=" & mUser.Stats.Banco
        str = str & ",MaxHP=" & mUser.Stats.MaxHP
        str = str & ",MinHP=" & mUser.Stats.MinHP
        str = str & ",MaxMAN=" & mUser.Stats.MaxMAN
        str = str & ",MinMAN=" & mUser.Stats.MinMAN
        str = str & ",MinSTA=" & mUser.Stats.MinSTA
        str = str & ",MaxSTA=" & mUser.Stats.MaxSTA
        str = str & ",MaxHIT=" & mUser.Stats.MaxHit
        str = str & ",MinHIT=" & mUser.Stats.MinHit
        str = str & ",MinAGU=" & mUser.Stats.MinAGU
        str = str & ",MinHAM=" & mUser.Stats.MinHAM
        str = str & ",SkillPtsLibres=" & mUser.Stats.SkillPts
        str = str & ",VecesMurioUsuario=" & mUser.Stats.VecesMuertos
        str = str & ",Exp=" & mUser.Stats.Exp
        str = str & ",ELV=" & mUser.Stats.ELV
        str = str & ",NpcsMuertes=" & mUser.Stats.NPCsMuertos
        str = str & ",ELU=" & mUser.Stats.ELU
        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        '************************************************************************

        Exit Sub
ErrHandle:
        Resume Next
    End Sub
    Sub SaveUserAtrib(ByVal UserIndex As Integer, ByVal ipj As Integer)

        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim mUser As User
        Dim i As Byte
        Dim str As String

        mUser = UserList(UserIndex)

        If Len(mUser.Name) = 0 Then Exit Sub


        '************************************************************************
        RS = DB_Conn.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        str = "UPDATE `charatrib` SET"
        str = str & " IndexPJ=" & ipj
        For i = 1 To NUMATRIBUTOS
            str = str & ",AT" & i & "=" & mUser.Stats.UserAtributosBackUP(i)
        Next i
        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        '************************************************************************
        Exit Sub
ErrHandle:
        Resume Next
    End Sub
    Sub SaveUserSkill(ByVal UserIndex As Integer, ByVal ipj As Integer)
        On Error GoTo ErrHandle
        Dim RS As ADODB.Recordset
        Dim mUser As User
        Dim i As Byte
        Dim str As String

        mUser = UserList(UserIndex)

        If Len(mUser.Name) = 0 Then Exit Sub

        '************************************************************************
        RS = DB_Conn.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Sub
        RS = Nothing

        str = "UPDATE `charskills` SET"
        str = str & " IndexPJ=" & ipj

        For i = 1 To NUMSKILLS
            str = str & ",SK" & i & "=" & mUser.Stats.UserSkills(i)
        Next i

        str = str & " WHERE IndexPJ=" & ipj & " LIMIT 1"
        Call DB_Conn.Execute(str)
        '************************************************************************
        Exit Sub
ErrHandle:
        Resume Next
    End Sub
    Function LoadUserSQL(UserIndex As Integer, ByVal Name As String) As Boolean
        On Error GoTo Errhandler
        Dim i As Integer
        Dim RS As New ADODB.Recordset
        Dim ipj As Integer
        Dim priv As Integer


        With UserList(UserIndex)

            '************************************************************************
            RS = DB_Conn.Execute("SELECT (IndexPJ) FROM `charflags` WHERE Nombre='" & Name & "' LIMIT 1")
            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charflags: Error al cargar charflags. Personaje: " & Name)
                Exit Function
            End If

            ipj = Convert.ToInt32(RS.Fields.Item("Indexpj").Value)
            RS = Nothing
            '************************************************************************

            .Indexpj = ipj

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & ipj & " LIMIT 1")

            If RS.BOF Or RS.EOF Then
                Call LogError("Error en LoadUserSQL/charflags: Error al cargar charflags Personaje: " & Name)
                LoadUserSQL = False
            End If

            .flags.ban = Convert.ToByte(RS.Fields.Item("ban").Value)
            .flags.Navegando = Convert.ToByte(RS.Fields.Item("Navegando").Value)
            .flags.Envenenado = Convert.ToByte(RS.Fields.Item("Envenenado").Value)
            .Counters.Pena = Convert.ToDouble(RS.Fields.Item("Pena").Value)
            .flags.Paralizado = 0

            .flags.Desnudo = Convert.ToByte(RS.Fields.Item("Desnudo").Value)
            .flags.Sed = Convert.ToByte(RS.Fields.Item("Sed").Value)
            .flags.Hambre = Convert.ToByte(RS.Fields.Item("Hambre").Value)
            .flags.Escondido = Convert.ToByte(RS.Fields.Item("Escondido").Value)
            .flags.Muerto = Convert.ToByte(RS.Fields.Item("Muerto").Value)
            'Add Nod Kopfnickend Nunca se guardaba en la DB y por ende nunca se cargaba
            .desc = RS.Fields.Item("desc").Value.ToString()
            '\Add

            priv = Convert.ToInt32(RS.Fields.Item("Privilegio").Value)
            'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
            .flags.AdminPerseguible = True

            If priv = 11 Then
                .flags.Privilegios = .flags.Privilegios Or PlayerType.VIP
            ElseIf priv = 10 Then
                .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
                .flags.AdminPerseguible = False
            ElseIf priv = 9 Then
                .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
                .flags.AdminPerseguible = False
            ElseIf priv = 8 Then
                .flags.Privilegios = .flags.Privilegios Or PlayerType.Semi
            ElseIf priv = 7 Then
                .flags.Privilegios = .flags.Privilegios Or PlayerType.Conse

                'Add Lideres faccionarios
            ElseIf priv = 3 Then
                .flags.Privilegios = .flags.Privilegios Or PlayerType.FaccCaos
            ElseIf priv = 2 Then
                .flags.Privilegios = .flags.Privilegios Or PlayerType.FaccRepu
            ElseIf priv = 1 Then
                .flags.Privilegios = .flags.Privilegios Or PlayerType.FaccImpe

            Else ' Es un pobre diablo
                .flags.Privilegios = .flags.Privilegios Or PlayerType.User
            End If

            RS = Nothing
            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & ipj & " LIMIT 1")

            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charfaccion: Error al cargar charfaccion. Personaje: " & Name)
                Exit Function
            End If

            ' Carga Faccion
            .faccion.ArmadaReal = Convert.ToByte(RS.Fields.Item("EjercitoReal").Value)
            .faccion.FuerzasCaos = Convert.ToByte(RS.Fields.Item("EjercitoCaos").Value)
            .faccion.Milicia = Convert.ToByte(RS.Fields.Item("EjercitoMili").Value)
            .faccion.Republicano = Convert.ToByte(RS.Fields.Item("Republicano").Value)
            .faccion.Ciudadano = Convert.ToByte(RS.Fields.Item("Ciudadano").Value)
            .faccion.Rango = Convert.ToByte(RS.Fields.Item("Rango").Value)
            .faccion.Renegado = Convert.ToByte(RS.Fields.Item("Renegado").Value)
            .faccion.CiudadanosMatados = Convert.ToInt32(RS.Fields.Item("CiudMatados").Value)
            .faccion.RenegadosMatados = Convert.ToInt32(RS.Fields.Item("ReneMatados").Value)
            .faccion.RepublicanosMatados = Convert.ToInt32(RS.Fields.Item("RepuMatados").Value)
            .faccion.CaosMatados = Convert.ToInt32(RS.Fields.Item("CaosMatados").Value)
            .faccion.ArmadaMatados = Convert.ToInt32(RS.Fields.Item("ArmadaMatados").Value)
            .faccion.MilicianosMatados = Convert.ToInt32(RS.Fields.Item("MiliMatados").Value)
            ' Fin Carga Faccion

            'Add Marius Un fix por un error mio xD
            If .faccion.Renegado = 1 Then
                Call ResetFacciones(UserIndex, False)
                .faccion.Renegado = 1
            End If

            RS = Nothing
            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & ipj & " LIMIT 1")

            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charatrib: Error al cargar charatrib. Personaje: " & Name)
                Exit Function
            End If

            ReDim .Stats.UserAtributos(NUMATRIBUTOS + 1)
            ReDim .Stats.UserAtributosBackUP(NUMATRIBUTOS + 1)

            For i = 1 To NUMATRIBUTOS
                .Stats.UserAtributos(i) = Convert.ToByte(RS.Fields("AT" & i).Value)
                .Stats.UserAtributosBackUP(i) = .Stats.UserAtributos(i)
            Next i

            RS = Nothing
            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & ipj & " LIMIT 1")
            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charguild: Error al cargar charguild. Personaje: " & Name)
                Exit Function
            End If

            UserList(UserIndex).GuildIndex = Convert.ToInt32(RS.Fields.Item("GuildIndex").Value)
            RS = Nothing
            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & ipj & " LIMIT 1")
            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charskills: Error al cargar charskill. Personaje: " & Name)
                Exit Function
            End If

            ReDim .Stats.UserSkills(NUMSKILLS + 1)

            For i = 1 To NUMSKILLS
                .Stats.UserSkills(i) = Convert.ToSByte(RS.Fields.Item("SK" & i).Value)
            Next i
            RS = Nothing
            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & ipj & " LIMIT 1")
            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charbanco: Error al cargar charbanco. Personaje: " & Name)
                Exit Function
            End If

            ReDim .BancoInvent.Objeto(MAX_BANCOINVENTORY_SLOTS + 1)

            For i = 1 To MAX_BANCOINVENTORY_SLOTS
                .BancoInvent.Objeto(i).ObjIndex = Convert.ToInt32(RS.Fields("OBJ" & i).Value)
                .BancoInvent.Objeto(i).Amount = Convert.ToInt32(RS.Fields("CANT" & i).Value)
            Next i
            RS = Nothing
            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & ipj & " LIMIT 1")
            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charinvent: Error al cargar charinvent. Personaje: " & Name)
                Exit Function
            End If

            ReDim .Invent.Objeto(MAX_INVENTORY_SLOTS + 1)

            For i = 1 To MAX_INVENTORY_SLOTS
                .Invent.Objeto(i).ObjIndex = Convert.ToInt32(RS.Fields("OBJ" & i).Value)
                .Invent.Objeto(i).Amount = Convert.ToInt32(RS.Fields("CANT" & i).Value)
            Next i

            .Invent.CascoEqpSlot = Convert.ToByte(RS.Fields.Item("CASCOSLOT").Value)
            .Invent.ArmourEqpSlot = Convert.ToByte(RS.Fields.Item("ARMORSLOT").Value)
            .Invent.EscudoEqpSlot = Convert.ToByte(RS.Fields.Item("SHIELDSLOT").Value)
            .Invent.WeaponEqpSlot = Convert.ToByte(RS.Fields.Item("WEAPONSLOT").Value)
            .Invent.AnilloEqpSlot = Convert.ToByte(RS.Fields.Item("ANILLOSLOT").Value)
            .Invent.MunicionEqpSlot = Convert.ToByte(RS.Fields.Item("MUNICIONSLOT").Value)
            .Invent.BarcoSlot = Convert.ToByte(RS.Fields.Item("BarcoSlot").Value)
            .Invent.NudiEqpSlot = Convert.ToByte(RS.Fields.Item("NUDISLOT").Value)
            .Invent.MonturaSlot = Convert.ToByte(RS.Fields.Item("MONTUSLOT").Value)
            RS = Nothing
            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charmascotafami` WHERE IndexPJ=" & ipj)
            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charmascotafami: Error al cargar charmascotafami. Personaje: " & Name)
                Exit Function
            End If

            .masc.TieneFamiliar = Convert.ToByte(RS.Fields.Item("Tiene").Value)
            .masc.Nombre = RS.Fields.Item("Nombre").Value.ToString()
            .masc.tipo = Convert.ToInt32(RS.Fields.Item("tipo").Value)
            .masc.ELV = Convert.ToByte(RS.Fields.Item("level").Value)
            .masc.ELU = Convert.ToDouble(RS.Fields.Item("ELU").Value)
            .masc.Exp = Convert.ToDouble(RS.Fields.Item("Exp").Value)
            .masc.MinHP = Convert.ToInt32(RS.Fields.Item("MinHP").Value)
            .masc.MaxHP = Convert.ToInt32(RS.Fields.Item("MaxHP").Value)
            .masc.MinHit = Convert.ToInt32(RS.Fields.Item("MinHit").Value)
            .masc.MaxHit = Convert.ToInt32(RS.Fields.Item("MaxHit").Value)
            RS = Nothing
            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & ipj & " LIMIT 1")
            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charhechizos: Error al cargar charhechizo. Personaje: " & Name)
                Exit Function
            End If

            ReDim .Stats.UserHechizos(MAXUSERHECHIZOS + 1)

            For i = 1 To MAXUSERHECHIZOS
                .Stats.UserHechizos(i) = Convert.ToInt32(RS.Fields("H" & i).Value)
            Next i

            RS = Nothing
            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & ipj & " LIMIT 1")

            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charstats: Error al cargar charstats. Personaje: " & Name)
                Exit Function
            End If

            .Stats.GLD = Convert.ToDouble(RS.Fields.Item("GLD").Value)
            .Stats.Banco = Convert.ToDouble(RS.Fields.Item("Banco").Value)
            .Stats.MaxHP = Convert.ToInt32(RS.Fields.Item("MaxHP").Value)
            .Stats.MinHP = Convert.ToInt32(RS.Fields.Item("MinHP").Value)
            .Stats.MinSTA = Convert.ToInt32(RS.Fields.Item("MinSTA").Value)
            .Stats.MaxSTA = Convert.ToInt32(RS.Fields.Item("MaxSTA").Value)
            .Stats.MaxMAN = Convert.ToInt32(RS.Fields.Item("MaxMAN").Value)
            .Stats.MinMAN = Convert.ToInt32(RS.Fields.Item("MinMAN").Value)
            .Stats.MaxHit = Convert.ToInt32(RS.Fields.Item("MaxHit").Value)
            .Stats.MinHit = Convert.ToInt32(RS.Fields.Item("MinHit").Value)
            .Stats.MinAGU = Convert.ToInt32(RS.Fields.Item("MinAGU").Value)
            .Stats.MinHAM = Convert.ToInt32(RS.Fields.Item("MinHAM").Value)
            .Stats.MaxAGU = 100
            .Stats.MaxHAM = 100
            .Stats.SkillPts = Convert.ToInt32(RS.Fields.Item("SkillPtsLibres").Value)
            .Stats.VecesMuertos = Convert.ToDouble(RS.Fields.Item("VecesMurioUsuario").Value)
            .Stats.Exp = Convert.ToDouble(RS.Fields.Item("Exp").Value)
            .Stats.ELV = Convert.ToInt32(RS.Fields.Item("ELV").Value)
            .Stats.NPCsMuertos = Convert.ToInt32(RS.Fields.Item("NpcsMuertes").Value)
            .Stats.ELU = Convert.ToDouble(RS.Fields.Item("ELU").Value)

            RS = Nothing

            If .Stats.MinAGU < 1 Then .flags.Sed = 1
            If .Stats.MinHAM < 1 Then .flags.Hambre = 1
            If .Stats.MinHP < 1 Then .flags.Muerto = 1

            '************************************************************************

            '************************************************************************
            RS = DB_Conn.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & ipj & " LIMIT 1")
            If RS.BOF Or RS.EOF Then
                LoadUserSQL = False
                Call LogError("Error en LoadUserSQL/charinit: Error al cargar charinit. Personaje: " & Name)
                Exit Function
            End If

            .Genero = Convert.ToInt32(RS.Fields.Item("Genero").Value)
            .Raza = Convert.ToInt32(RS.Fields.Item("Raza").Value)
            .Hogar = Convert.ToByte(RS.Fields.Item("Hogar").Value)
            .Clase = Convert.ToInt32(RS.Fields.Item("Clase").Value)
            .OrigChar.heading = Convert.ToInt32(RS.Fields.Item("heading").Value)
            .OrigChar.Head = Convert.ToInt32(RS.Fields.Item("Head").Value)
            .OrigChar.body = Convert.ToInt32(RS.Fields.Item("body").Value)
            .OrigChar.WeaponAnim = Convert.ToInt32(RS.Fields.Item("Arma").Value)
            .OrigChar.ShieldAnim = Convert.ToInt32(RS.Fields.Item("Escudo").Value)
            .OrigChar.CascoAnim = Convert.ToInt32(RS.Fields.Item("casco").Value)
            .ip = RS.Fields.Item("LastIP").Value.ToString()
            .Pos.map = Convert.ToInt32(RS.Fields.Item("mapa").Value)
            .Pos.x = Convert.ToInt32(RS.Fields.Item("x").Value)
            .Pos.Y = Convert.ToInt32(RS.Fields.Item("Y").Value)
            .flags.miPareja = RS.Fields.Item("PAREJA").Value.ToString()

            'Add Marius agregamos esto para que funcione los casamientos. (xD)
            If Len(.flags.miPareja) > 0 Then
                .flags.toyCasado = 1
            End If
            '\Add


            ReDim .Matados(MAX_BUFFER_KILLEDS + 1)
            ReDim .Matados_timer(MAX_BUFFER_KILLEDS + 1)

            'Add Marius Anti Frags
            For i = 1 To MAX_BUFFER_KILLEDS
                .Matados(i) = Convert.ToInt32(RS.Fields("pj" & i).Value)
                .Matados_timer(i) = Convert.ToDouble(RS.Fields("time" & i).Value)

                If .Matados(i) > 0 And .Matados_timer(i) > 0 Then
                    .Matados_timer(i) = GetTickCount + .Matados_timer(i)
                End If

            Next i
            '\Add

            If .flags.Muerto = 0 Then
                .cuerpo = .OrigChar
                Call VerObjetosEquipados(UserIndex)
            Else
                .cuerpo.body = iCuerpoMuerto
                .cuerpo.Head = iCabezaMuerto
                .cuerpo.WeaponAnim = NingunArma
                .cuerpo.ShieldAnim = NingunEscudo
                .cuerpo.CascoAnim = NingunCasco
            End If

            RS = Nothing

            '************************************************************************


            If Len(.desc) > 100 Then .desc = Left$(.desc, 100)

            .Stats.MaxAGU = 100
            .Stats.MaxHAM = 100



            '************************************************************************

            ReDim .Correos(20 + 1)

            RS = DB_Conn.Execute("SELECT * FROM `charcorreo` WHERE IndexPJ=" & ipj & " LIMIT 20")

            If Not (RS.BOF Or RS.EOF) Then

                Dim ii As Byte

                .cant_mensajes = RS.RecordCount



                For ii = 1 To .cant_mensajes
                    .Correos(ii).idmsj = Convert.ToDouble(RS.Fields("Idmsj").Value)
                    .Correos(ii).Mensaje = RS.Fields("Mensaje").Value.ToString()
                    .Correos(ii).De = RS.Fields("De").Value.ToString()
                    .Correos(ii).Cantidad = Convert.ToInt32(RS.Fields("Cantidad").Value)
                    .Correos(ii).Item = Convert.ToInt32(RS.Fields("Item").Value)
                    RS.MoveNext()
                Next ii
            Else
                .cant_mensajes = 0
            End If
            RS = Nothing
            '************************************************************************



            LoadUserSQL = True

        End With

        Exit Function

Errhandler:
        Call LogError("Error en LoadUserSQL. N:        " & Name & " - " & Err.Number & "-" & Err.Description)
        RS = Nothing

    End Function
    Function Add_GLD_Subast(ByRef Name As String, ByVal oro As Integer)
        Dim RS As New ADODB.Recordset
        Dim Indexpj As Integer
        Dim str As String

        Indexpj = GetIndexPJ(Name)

        If Indexpj <> 0 Then
            RS = DB_Conn.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
            If (RS.BOF Or RS.EOF) = False Then
                str = "UPDATE `charstats` SET"
                str = str & " IndexPJ=" & Indexpj
                str = str & ",GLD=" & RS.Fields.Item("GLD").Value.ToString() + oro
                str = str & ",BANCO=" & RS.Fields.Item("Banco").Value.ToString()
                str = str & ", MaxHP = " & RS.Fields.Item("MaxHP").Value.ToString()
                str = str & ",MinHP=" & RS.Fields.Item("MinHP").Value.ToString()
                str = str & ", MaxMAN = " & RS.Fields.Item("MaxMAN").Value.ToString()
                str = str & ",MinMAN=" & RS.Fields.Item("MinMAN").Value.ToString()
                str = str & ", MinSTA = " & RS.Fields.Item("MinSTA").Value.ToString()
                str = str & ",MaxSTA=" & RS.Fields.Item("MaxSTA").Value.ToString()
                str = str & ", MaxHIT = " & RS.Fields.Item("MaxHit").Value.ToString()
                str = str & ",MinHIT=" & RS.Fields.Item("MinHit").Value.ToString()
                str = str & ", MinAGU = " & RS.Fields.Item("MinAGU").Value.ToString()
                str = str & ",MinHAM=" & RS.Fields.Item("MinHAM").Value.ToString()
                str = str & ", SkillPtsLibres = " & RS.Fields.Item("SkillPtsLibres").Value.ToString()
                str = str & ",VecesMurioUsuario=" & RS.Fields.Item("VecesMurioUsuario").Value.ToString()
                str = str & ", Exp = " & RS.Fields.Item("Exp").Value.ToString()
                str = str & ",ELV=" & RS.Fields.Item("ELV").Value.ToString()
                str = str & ", NpcsMuertes = " & RS.Fields.Item("NpcsMuertes").Value.ToString()
                str = str & ",ELU=" & RS.Fields.Item("ELU").Value.ToString()
                str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"

                Call DB_Conn.Execute(str)
                RS = Nothing
            End If
        End If
    End Function
    Function Add_Bank_Gold(ByRef Name As String, ByVal oro As Long) As Boolean
        On Error GoTo LocalErr
        Dim RS As New ADODB.Recordset
        Dim Indexpj As Integer
        Dim str As String

        Indexpj = GetIndexPJ(Name)

        If Indexpj <> 0 Then
            RS = DB_Conn.Execute("Select * FROM `charstats` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
            If (RS.BOF Or RS.EOF) = False Then
                str = "UPDATE `charstats` Set"
                str = str & " IndexPJ=" & Indexpj
                str = str & ", GLD = " & RS.Fields.Item("GLD").Value.ToString()
                str = str & ", BANCO = " & RS.Fields.Item("Banco + oro").Value.ToString()
                str = str & ", MaxHP = " & RS.Fields.Item("MaxHP").Value.ToString()
                str = str & ", MinHP = " & RS.Fields.Item("MinHP").Value.ToString()
                str = str & ", MaxMAN = " & RS.Fields.Item("MaxMAN").Value.ToString()
                str = str & ", MinMAN = " & RS.Fields.Item("MinMAN").Value.ToString()
                str = str & ", MinSTA = " & RS.Fields.Item("MinSTA").Value.ToString()
                str = str & ", MaxSTA = " & RS.Fields.Item("MaxSTA").Value.ToString()
                str = str & ", MaxHIT = " & RS.Fields.Item("MaxHit").Value.ToString()
                str = str & ", MinHIT = " & RS.Fields.Item("MinHit").Value.ToString()
                str = str & ", MinAGU = " & RS.Fields.Item("MinAGU").Value.ToString()
                str = str & ", MinHAM = " & RS.Fields.Item("MinHAM").Value.ToString()
                str = str & ", SkillPtsLibres = " & RS.Fields.Item("SkillPtsLibres").Value.ToString()
                str = str & ", VecesMurioUsuario = " & RS.Fields.Item("VecesMurioUsuario").Value.ToString()
                str = str & ", Exp = " & RS.Fields.Item("Exp").Value.ToString()
                str = str & ", ELV = " & RS.Fields.Item("ELV").Value.ToString()
                str = str & ", NpcsMuertes = " & RS.Fields.Item("NpcsMuertes").Value.ToString()
                str = str & ", ELU = " & RS.Fields.Item("ELU").Value.ToString()
                str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"

                Call DB_Conn.Execute(str)
                RS = Nothing

                Add_Bank_Gold = True
                Exit Function
            End If
        End If
        Exit Function

LocalErr:
        Add_Bank_Gold = False
        Exit Function
    End Function
    Function Add_Item_Subast(ByRef Name As String, ByVal obji As Integer, ByVal Cant As Integer) As Boolean
        Dim RS As New ADODB.Recordset
        Dim Indexpj As Integer
        Dim i As Integer, j As Integer, t As Integer
        Dim b As Boolean
        Dim str As String

        Indexpj = GetIndexPJ(UCase$(Name))

        If Indexpj <> 0 Then
            RS = DB_Conn.Execute("Select * FROM `charbanco` WHERE IndexPJ=" & Indexpj & " LIMIT 1")

            If (RS.BOF Or RS.EOF) = False Then
                str = "UPDATE `charbanco` Set"
                str = str & " IndexPJ=" & Indexpj
                'Buscamos en el banco
                For i = 1 To MAX_BANCOINVENTORY_SLOTS
                    j = Convert.ToInt32(RS.Fields("OBJ" & i))
                    If j = 0 And b = False Then
                        str = str & ", obj" & i & "=" & obji
                        str = str & ", CANT" & i & "=" & Cant
                        b = True
                    Else
                        str = str & ", obj" & i & "=" & j
                        str = str & ", CANT" & i & "=" & RS.Fields("CANT" & i).ToString()
                    End If
                Next i
                str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"

                If b = True Then
                    Call DB_Conn.Execute(str)
                    RS = Nothing
                    Add_Item_Subast = True
                    Exit Function
                End If
            End If

            str = ""
            b = False
            j = 0

            RS = Nothing

            '************************************************************************
            RS = DB_Conn.Execute("Select * FROM `charinvent` WHERE IndexPJ=" & Indexpj & " LIMIT 1")
            If (RS.BOF Or RS.EOF) = False Then

                str = "UPDATE `charinvent` Set"
                str = str & " IndexPJ=" & Indexpj

                For i = 1 To MAX_INVENTORY_SLOTS
                    j = Convert.ToInt32(RS.Fields("OBJ" & i))
                    If j = 0 And b = False Then
                        str = str & ", obj" & i & "=" & obji
                        str = str & ", CANT" & i & "=" & Cant
                        b = True
                    Else
                        str = str & ", obj" & i & "=" & j
                        str = str & ", CANT" & i & "=" & RS.Fields("CANT" & i).ToString()
                    End If
                Next i
                str = str & ", CASCOSLOT = " & RS.Fields.Item("CASCOSLOT").Value.ToString()
                str = str & ", ARMORSLOT = " & RS.Fields.Item("ARMORSLOT").Value.ToString()
                str = str & ", SHIELDSLOT = " & RS.Fields.Item("SHIELDSLOT").Value.ToString()
                str = str & ", WEAPONSLOT = " & RS.Fields.Item("WEAPONSLOT").Value.ToString()
                str = str & ", ANILLOSLOT = " & RS.Fields.Item("ANILLOSLOT").Value.ToString()
                str = str & ", MUNICIONSLOT = " & RS.Fields.Item("MUNICIONSLOT").Value.ToString()
                str = str & ", BARCOSLOT = " & RS.Fields.Item("BarcoSlot").Value.ToString()
                str = str & ", NUDISLOT = " & RS.Fields.Item("NUDISLOT").Value.ToString()
                str = str & ", MONTUSLOT = " & RS.Fields.Item("MONTUSLOT").Value.ToString()

                str = str & " WHERE IndexPJ=" & Indexpj & " LIMIT 1"

                If b = True Then
                    Call DB_Conn.Execute(str)
                    Add_Item_Subast = True
                Else
                    Add_Item_Subast = False
                End If

                RS = Nothing
                Exit Function
            End If
        End If

    End Function
    Public Function BANCheckDB(ByVal Name As String) As Boolean
        Dim RS As New ADODB.Recordset
        Dim Baneado As Byte

        RS = DB_Conn.Execute("Select * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function

        Baneado = Convert.ToByte(RS.Fields.Item("ban").Value)
        BANCheckDB = (Baneado >= 1)
        RS = Nothing

    End Function

    Function ExistePersonaje(Name As String) As Boolean
        Dim RS As New ADODB.Recordset

        RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
        RS = Nothing

        ExistePersonaje = True
    End Function
    Function GetIndexPJ(Name As String) As Integer
        On Error GoTo err
        Dim RS As New ADODB.Recordset
        Dim Indexpj As Long
        Dim Index As Integer

        'Add Marius Si ya esta cargado para que buscarlo otra vez xD
        'Abria que testearlo ahora no tengo tiempo
        'index = NameIndex(Name)
        'If index > 0 And UserList(index).Indexpj <> 0 Then
        '    Indexpj = UserList(index).Indexpj
        '    Exit Function
        'End If
        '\Add

        RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(Name) & "'")
        If RS.BOF Or RS.EOF Then
            GoTo err
        Else
            GetIndexPJ = Convert.ToInt32(RS.Fields.Item("Indexpj").Value)
        End If
        RS = Nothing
        Exit Function

err:
        RS = Nothing
        GetIndexPJ = 0
        Exit Function
    End Function

    Public Sub SendOnline()
        Call extra_set("online", Str(NumUsers))
        Console.WriteLine("Online: " & NumUsers)
    End Sub

    Public Sub VerObjetosEquipados(UserIndex As Integer)

        With UserList(UserIndex).Invent
            If .CascoEqpSlot Then
                .Objeto(.CascoEqpSlot).Equipped = 1
                .CascoEqpObjIndex = .Objeto(.CascoEqpSlot).ObjIndex
                UserList(UserIndex).cuerpo.CascoAnim = ObjDataArr(.CascoEqpObjIndex).CascoAnim
            Else
                UserList(UserIndex).cuerpo.CascoAnim = NingunCasco
            End If

            If .BarcoSlot Then .BarcoObjIndex = .Objeto(.BarcoSlot).ObjIndex

            If .ArmourEqpSlot Then
                .Objeto(.ArmourEqpSlot).Equipped = 1
                .ArmourEqpObjIndex = .Objeto(.ArmourEqpSlot).ObjIndex
                UserList(UserIndex).cuerpo.body = ObjDataArr(.ArmourEqpObjIndex).Ropaje
            Else
                Call DarCuerpoDesnudo(UserIndex)
            End If

            If .WeaponEqpSlot > 0 Then
                .Objeto(.WeaponEqpSlot).Equipped = 1
                .WeaponEqpObjIndex = .Objeto(.WeaponEqpSlot).ObjIndex
                If .Objeto(.WeaponEqpSlot).ObjIndex > 0 Then UserList(UserIndex).cuerpo.WeaponAnim = ObjDataArr(.WeaponEqpObjIndex).WeaponAnim
            Else
                UserList(UserIndex).cuerpo.WeaponAnim = NingunArma
            End If

            If .EscudoEqpSlot > 0 Then
                .Objeto(.EscudoEqpSlot).Equipped = 1
                .EscudoEqpObjIndex = .Objeto(.EscudoEqpSlot).ObjIndex
                UserList(UserIndex).cuerpo.ShieldAnim = ObjDataArr(.EscudoEqpObjIndex).ShieldAnim
            Else
                UserList(UserIndex).cuerpo.ShieldAnim = NingunEscudo
            End If

            If .MunicionEqpSlot Then
                .Objeto(.MunicionEqpSlot).Equipped = 1
                .MunicionEqpObjIndex = .Objeto(.MunicionEqpSlot).ObjIndex
            End If

            If .NudiEqpSlot > 0 Then
                .Objeto(.NudiEqpSlot).Equipped = 1
                .NudiEqpIndex = .Objeto(.NudiEqpSlot).ObjIndex
                UserList(UserIndex).cuerpo.WeaponAnim = ObjDataArr(.NudiEqpIndex).WeaponAnim
            End If

            'Fix Marius
            UserList(UserIndex).Invent.MonturaSlot = 0
            '\Fix

            If UserList(UserIndex).Invent.MonturaSlot <> 0 Then
                UserList(UserIndex).Invent.MonturaObjIndex = UserList(UserIndex).Invent.Objeto(UserList(UserIndex).Invent.MonturaSlot).ObjIndex
                UserList(UserIndex).cuerpo.body = ObjDataArr(UserList(UserIndex).Invent.MonturaObjIndex).Ropaje
                UserList(UserIndex).flags.Montando = 1
                Call WriteEquitateToggle(UserIndex)
            Else
                UserList(UserIndex).flags.Montando = 0
            End If

        End With

    End Sub
    Public Function Insert_New_Table(ByRef Name As String, ByRef id As Long) As Integer
        On Error GoTo Erro
        Dim ipj As Integer

        Dim RS As ADODB.Recordset
        RS = New ADODB.Recordset

        DB_Conn.Execute("INSERT INTO `charflags` (id,Nombre) VALUES (" & id & ",'" & Name & "')")

        RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE Nombre='" & Name & "'")
        ipj = Convert.ToInt32(RS.Fields.Item("Indexpj").Value)
        RS = Nothing

        DB_Conn.Execute("INSERT INTO `charatrib` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("INSERT INTO `charbanco` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("INSERT INTO `charcorreo` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("INSERT INTO `charfaccion` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("INSERT INTO `charguild` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("UPDATE `charguild` SET GuildIndex=0 WHERE IndexPJ=" & ipj & " LIMIT 1")
        DB_Conn.Execute("UPDATE `charguild` SET AspiranteA=0 WHERE IndexPJ=" & ipj & " LIMIT 1")

        DB_Conn.Execute("INSERT INTO `charhechizos` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("INSERT INTO `charinit` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("INSERT INTO `charinvent` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("INSERT INTO `charmascotafami` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("INSERT INTO `charskills` (IndexPJ) VALUES (" & ipj & ")")
        DB_Conn.Execute("INSERT INTO `charstats` (IndexPJ) VALUES (" & ipj & ")")

        Insert_New_Table = ipj
        Exit Function
Erro:
        LogError("Insert_New_Table " & Name & " " & Err.Number & " " & Err.Description)
    End Function


    Public Sub Quitarcorreosql(ByVal idmsj As Long)
        Dim RS As New ADODB.Recordset

        RS = DB_Conn.Execute("DELETE FROM `charcorreo` WHERE Idmsj=" & idmsj & " LIMIT 1")

        RS = Nothing

    End Sub


    Public Function Cantidadmensajes(Indexpj As Integer) As Byte

        Dim RS As New ADODB.Recordset

        RS = DB_Conn.Execute("Select IndexPJ FROM `charcorreo` WHERE IndexPJ=" & Indexpj)
        Cantidadmensajes = RS.RecordCount
        RS = Nothing

    End Function

    Public Function EnviarCorreoSql(ByVal ipj As Integer, ByVal loopC As Byte, ByVal para As Integer) As Long
        Dim RS As New ADODB.Recordset
        Dim str As String
        With UserList(para)
            str = "INSERT INTO `charcorreo` SET"
            str = str & " IndexPj=" & ipj
            str = str & ",Mensaje='" & .Correos(loopC).Mensaje & "'"
            str = str & ",De='" & .Correos(loopC).De & "'"
            str = str & ",Cantidad=" & .Correos(loopC).Cantidad
            str = str & ",Item=" & .Correos(loopC).Item

            Call DB_Conn.Execute(str)

            RS = Nothing

            'Add Marius Sin esto se puede duplicar objetos
            Dim id As Long
            RS = DB_Conn.Execute("SELECT last_insert_id() as id")
            If RS.BOF Or RS.EOF Then
                'GoTo err
            Else
                EnviarCorreoSql = RS.Fields.Item("id").Value.ToString()
            End If
            '\Add

        End With
        RS = Nothing
    End Function

    Public Sub onpj(ByVal UserIndex As Integer)
        DB_Conn.Execute("UPDATE `charflags` SET `Online` = '1' WHERE `IndexPJ` = " & UserList(UserIndex).Indexpj & " LIMIT 1")
    End Sub

    Public Sub offpj(ByVal UserIndex As Integer)
        DB_Conn.Execute("UPDATE `charflags` SET `Online` = '0' WHERE `IndexPJ` = " & UserList(UserIndex).Indexpj & " LIMIT 1")
    End Sub
    Public Sub torneo_contador(ByVal UserIndex As Integer)
        DB_Conn.Execute("UPDATE charflags SET Torneos = Torneos + 1 WHERE IndexPJ = " & UserIndex & " LIMIT 1")
    End Sub

    Public Sub extra_set(ByVal Nombre As String, ByVal valor As String)
        DB_Conn.Execute("UPDATE `extras` SET `valor` = '" & valor & "' WHERE `nombre` = '" & Nombre & "' LIMIT 1")
    End Sub

    Function extra_get(Nombre As String) As String
        On Error GoTo err
        Dim RS As New ADODB.Recordset
        Dim valor As String

        RS = DB_Conn.Execute("SELECT * FROM `extras` WHERE Nombre='" & Nombre & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then
            GoTo err
        Else
            extra_get = RS.Fields.Item("valor").Value.ToString()
        End If
        RS = Nothing
        Exit Function

err:
        RS = Nothing
        extra_get = "0"
        Exit Function
    End Function

End Module
