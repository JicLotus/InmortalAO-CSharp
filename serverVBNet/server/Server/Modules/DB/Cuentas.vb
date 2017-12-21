Option Explicit On

Module Cuentas


    Public Sub agregarcreditos(ByVal id As Long)
        DB_Conn.Execute("UPDATE `cuentas` SET creditos=creditos+2 WHERE id='" & id & "' LIMIT 1")
End Sub

    Public Function AddUserInAccount(ByVal id As Long)

        DB_Conn.Execute("UPDATE `cuentas` SET numpjs=numpjs+1 WHERE id='" & id & "' LIMIT 1")

End Function
    Function NumPJs(ByVal id As Long) As Byte
        Dim RS As New ADODB.Recordset
        RS = DB_Conn.Execute("SELECT * FROM `cuentas` WHERE id=" & id & " LIMIT 1")

        If RS.BOF Or RS.EOF Then Exit Function

        NumPJs = Convert.ToByte(RS.Fields.Item("NumPJs").Value)

        RS = Nothing
End Function
    Function CuentaVerificada(cuenta As String) As String
        Dim RS As New ADODB.Recordset
        RS = DB_Conn.Execute("SELECT verificada FROM `cuentas` WHERE nombre='" & UCase$(cuenta) & "' LIMIT 1")

        If RS.BOF Or RS.EOF Then
            Exit Function
        End If

        CuentaVerificada = (RS.Fields.Item("verificada").Value).ToString()

        RS = Nothing

    End Function

    Function ExisteCuenta(cuenta As String) As Long
        Dim RS As New ADODB.Recordset
        RS = DB_Conn.Execute("SELECT * FROM `cuentas` WHERE nombre='" & UCase$(cuenta) & "' LIMIT 1")

        If RS.BOF Or RS.EOF Then
            Exit Function
        End If

        ExisteCuenta = Convert.ToDouble(RS.Fields.Item("id").Value)

        RS = Nothing

    End Function

    Public Function ConectarCuenta(ByVal UserIndex As Integer, ByVal Name As String, ByVal Pass As String, Optional ByVal ReConnect As Byte = 0, Optional accountlogged As Boolean = False)
        Dim id_account As Long

        Name = UCase$(LTrim(RTrim(Name)))

        'Existe ya la cuenta?
        id_account = ExisteCuenta(Name)
        If id_account <= 0 Then
            Call WriteMsg(UserIndex, 45)
            Call WriteDisconnect2(UserIndex)
            Call FlushBuffer(UserIndex)
            Exit Function
        End If

        Dim verificada As String

        verificada = CuentaVerificada(Name)
        If verificada <> 1 Then
            Call WriteErrorMsg(UserIndex, "La cuenta no ha sido verificada. Revise su casilla de email.")
            Call WriteDisconnect2(UserIndex)
            Call FlushBuffer(UserIndex)
            Exit Function
        End If

        If ReConnect = 0 Then
            'Es la pass correcta y la cuenta offline?
            If ComprobarPasswordCuenta(Name, Pass) = False Then
                Call WriteMsg(UserIndex, 44)
                Call WriteDisconnect2(UserIndex)
                Call FlushBuffer(UserIndex)
                Exit Function
            End If
        End If

        If ServerSoloGMs > 0 Then
            If Betatest(Name) = 0 Then
                Call WriteMsg(UserIndex, 46)
                Call WriteDisconnect2(UserIndex)
                Call FlushBuffer(UserIndex)
                Exit Function
            End If
        End If



        UserList(UserIndex).account = Name
        UserList(UserIndex).IndexAccount = id_account

        Dim numPj As Byte
        numPj = NumPJs(UserList(UserIndex).IndexAccount)

        If accountlogged = False Then
            UserList(UserIndex).flags.accountlogged = True
        End If

        Call WriteShowAccount(UserIndex)

        If numPj > 0 Then
            Dim i As Byte
            For i = 1 To numPj
                Call WriteAddPj(UserIndex, leePjSqlCuenta(UserList(UserIndex).IndexAccount, i), i)
            Next i
        End If

        oncuenta(UserIndex)

    End Function
    Function UserTypeColorAcc(ByVal ipj As Integer) As Byte
        Dim Renegado As Byte, Armada As Byte, Caos As Byte, Repu As Byte, Impe As Byte, Mili As Byte
        Dim RS As ADODB.Recordset

    RS = DB_Conn.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & ipj & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function

        Armada = Convert.ToByte(RS.Fields.Item("EjercitoReal").Value)
        Caos = Convert.ToByte(RS.Fields.Item("EjercitoCaos").Value)
        Mili = Convert.ToByte(RS.Fields.Item("EjercitoMili").Value)
        Impe = Convert.ToByte(RS.Fields.Item("Ciudadano").Value)
        Repu = Convert.ToByte(RS.Fields.Item("Republicano").Value)
        Renegado = Convert.ToByte(RS.Fields.Item("Renegado").Value)
        RS = Nothing
    
    If Renegado = 1 Then
            UserTypeColorAcc = 1
        ElseIf Armada = 1 Or Impe = 1 Then
            UserTypeColorAcc = 2
        ElseIf Caos = 1 Then
            UserTypeColorAcc = 3
        ElseIf Mili = 1 Or Repu = 1 Then
            UserTypeColorAcc = 4
        Else
            UserTypeColorAcc = 1
        End If
    End Function
    Function leePjSqlCuenta(ByVal id As Long, ByVal Index As Byte) As String

        Dim RS As New ADODB.Recordset


    RS = DB_Conn.Execute("SELECT * FROM `charflags` WHERE id=" & id & " LIMIT " & (Index - 1) & ",1")
    If RS.BOF Or RS.EOF Then Exit Function

        leePjSqlCuenta = RS.Fields.Item("Nombre").Value.ToString()

        RS = Nothing

End Function
    'Mod Nod Kopfnickend
    Function ComprobarPasswordCuenta(ByVal cuenta As String, Password As String) As Boolean

        Dim Pass As String
        Dim online As Boolean
        Dim RS As New ADODB.Recordset
        Dim i As Long
        Dim id As Integer
    
    RS = DB_Conn.Execute("SELECT * FROM `cuentas` WHERE nombre='" & UCase$(cuenta) & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
        Pass = RS.Fields.Item("Password").Value.ToString()
        online = Convert.ToBoolean(RS.Fields.Item("online").Value)
        id = Convert.ToInt32(RS.Fields.Item("id").Value)
        If Len(Pass) = 0 Then Exit Function
    RS = Nothing
    
    'Si las claves son iguales y la cuenta esta Offline
    If Pass = Password Then

            ComprobarPasswordCuenta = True

            'verificamos que no haya una pj de esa cuenta conectado
            For i = 1 To LastUser
                If UserList(i).ConnID <> -1 Then
                    If UserList(i).IndexAccount = id Then
                        ComprobarPasswordCuenta = False
                        Exit For
                    End If
                End If
            Next i
        Else
            Call LogBruteforce("Cuenta: " & cuenta & " Pass: " & Password)
            ComprobarPasswordCuenta = False
        End If

    End Function

    Function ComprobarBanCuenta(ByVal id As Integer) As Boolean

        Dim ban As Integer
        Dim RS As New ADODB.Recordset
    
    RS = DB_Conn.Execute("SELECT * FROM `cuentas` WHERE id=" & id & " LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
        ban = Convert.ToInt32(RS.Fields.Item("ban").Value)
        RS = Nothing
    
    'Si las claves son iguales y la cuenta esta Offline
    If ban > 0 Then
            ComprobarBanCuenta = True
        Else
            ComprobarBanCuenta = False
        End If

    End Function


    Function Comprobar_Si_Donador(ByVal cuenta As String) As Byte

        Dim RS As New ADODB.Recordset
    
    RS = DB_Conn.Execute("SELECT (Donador) FROM `cuentas` WHERE nombre='" & UCase$(cuenta) & "' LIMIT 1")
        If RS.BOF Or RS.EOF Then Exit Function
        Comprobar_Si_Donador = Convert.ToByte(RS.Fields.Item("donador").Value)
        RS = Nothing
    
End Function

    Function Betatest(cuenta As String) As Byte
        Dim RS As New ADODB.Recordset
    RS = DB_Conn.Execute("SELECT (betatest) FROM `cuentas` WHERE nombre='" & UCase$(cuenta) & "' LIMIT 1")
    
    If RS.BOF Or RS.EOF Then Exit Function

        Betatest = Convert.ToByte(RS.Fields.Item("Betatest").Value)

        RS = Nothing

End Function

    Public Sub oncuenta(ByVal UserIndex As Integer)
        DB_Conn.Execute("UPDATE `cuentas` SET `Online` = 1 WHERE `ID` = " & UserList(UserIndex).IndexAccount)
    End Sub

    Public Sub offcuenta(ByVal UserIndex As Integer)
        DB_Conn.Execute("UPDATE `cuentas` SET `Online` = 0 WHERE `ID` = " & UserList(UserIndex).IndexAccount)
    End Sub

End Module
