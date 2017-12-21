﻿Option Explicit On

Module modSistemaComercio



    Enum eModoComercio
        Compra = 1
        Venta = 2
    End Enum

    Public Const REDUCTOR_PRECIOVENTA As Byte = 3

    ''
    ' Makes a trade. (Buy or Sell)
    '
    ' @param Modo The trade type (sell or buy)
    ' @param UserIndex Specifies the index of the user
    ' @param NpcIndex specifies the index of the npc
    ' @param Slot Specifies which slot are you trying to sell / buy
    ' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
    Public Sub Comercio(ByVal Modo As eModoComercio, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Integer, ByVal Cantidad As Integer)
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
        '  - 06/13/08 (NicoNZ)
        '*************************************************
        Dim Precio As Long
        Dim Objeto As obj

        On Error GoTo hayerror

        If Cantidad < 1 Or Slot < 1 Then Exit Sub

        If Modo = eModoComercio.Compra Then
            If Slot > MAX_INVENTORY_SLOTS Then
                Exit Sub
            ElseIf Cantidad > MAX_INVENTORY_OBJS Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(1, UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
                Call ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)

                Call ChangeBan(UserList(UserIndex).Name, 7)

                UserList(UserIndex).flags.ban = 1
                Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Sub
            ElseIf Not Npclist(NpcIndex).Invent.Objeto(Slot).Amount > 0 Then
                Exit Sub
            End If

            If Cantidad > Npclist(NpcIndex).Invent.Objeto(Slot).Amount Then Cantidad = Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Objeto(Slot).Amount

            Objeto.Amount = Cantidad
            Objeto.ObjIndex = Npclist(NpcIndex).Invent.Objeto(Slot).ObjIndex

            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)

            Precio = CLng((ObjDataArr(Npclist(NpcIndex).Invent.Objeto(Slot).ObjIndex).valor / Descuento(UserIndex) * Cantidad) + 0.5)

            If UserList(UserIndex).Stats.GLD < Precio Then
                Call WriteConsoleMsg(1, UserIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If


            If MeterItemEnInventario(UserIndex, Objeto) = False Then
                'Call WriteConsoleMsg(1,UserIndex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                Call WriteTradeOK(UserIndex)
                Exit Sub
            End If

            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio

            Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(Slot), Cantidad)

            'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
            If ObjDataArr(Objeto.ObjIndex).OBJType = eOBJType.otLlaves Then
                Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "obj" & Slot, Objeto.ObjIndex & "-0")
            End If

        ElseIf Modo = eModoComercio.Venta Then

            If Cantidad > UserList(UserIndex).Invent.Objeto(Slot).Amount Then Cantidad = UserList(UserIndex).Invent.Objeto(Slot).Amount

            Objeto.Amount = Cantidad
            Objeto.ObjIndex = UserList(UserIndex).Invent.Objeto(Slot).ObjIndex
            If Objeto.ObjIndex = 0 Then
                Exit Sub
            ElseIf ObjDataArr(Objeto.ObjIndex).Newbie = 1 Then
                Call WriteConsoleMsg(1, UserIndex, "Lo siento, no comercio objetos para newbies.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                Call WriteTradeOK(UserIndex)
                Exit Sub

            ElseIf (Npclist(NpcIndex).TipoItems <> ObjDataArr(Objeto.ObjIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
                Call WriteConsoleMsg(1, UserIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                Call WriteTradeOK(UserIndex)
                Exit Sub
                'Add Marius
            ElseIf ObjDataArr(Objeto.ObjIndex).Shop > 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Ese objeto es demaciado valioso, no tengo tanto oro para comprartelo.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                Call WriteTradeOK(UserIndex)
                Exit Sub
                '\Add
            ElseIf ObjDataArr(Objeto.ObjIndex).Real > 0 Then
                If Npclist(NpcIndex).Name <> "SR" Then
                    Call WriteConsoleMsg(1, UserIndex, "Las armaduras de la Armada solo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                    Call WriteTradeOK(UserIndex)
                    Exit Sub
                End If
            ElseIf ObjDataArr(Objeto.ObjIndex).Caos > 0 Then
                If Npclist(NpcIndex).Name <> "SC" Then
                    Call WriteConsoleMsg(1, UserIndex, "Las armaduras de la Legión solo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                    Call WriteTradeOK(UserIndex)
                    Exit Sub
                End If
            ElseIf ObjDataArr(Objeto.ObjIndex).Milicia > 0 Then
                If Npclist(NpcIndex).Name <> "SM" Then
                    Call WriteConsoleMsg(1, UserIndex, "Las armaduras de la Milicia solo pueden ser vendidas a los sastres milicianos.", FontTypeNames.FONTTYPE_INFO)
                    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                    Call WriteTradeOK(UserIndex)
                    Exit Sub
                End If
            ElseIf UserList(UserIndex).Invent.Objeto(Slot).Amount < 0 Or Cantidad = 0 Then
                Exit Sub
            ElseIf Slot < LBound(UserList(UserIndex).Invent.Objeto) Or Slot > UBound(UserList(UserIndex).Invent.Objeto) Then
                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                Exit Sub
            End If

            Call QuitarUserInvItem(UserIndex, Slot, Cantidad)

            'Precio = Round(ObjDataArr(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
            Precio = Fix(SalePrice(ObjDataArr(Objeto.ObjIndex).valor) * Cantidad)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Precio

            If UserList(UserIndex).Stats.GLD > MAXORO Then _
            UserList(UserIndex).Stats.GLD = MAXORO

            Dim NpcSlot As Integer
            NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)

            If NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
                'Mete el obj en el slot
                Npclist(NpcIndex).Invent.Objeto(NpcSlot).ObjIndex = Objeto.ObjIndex
                Npclist(NpcIndex).Invent.Objeto(NpcSlot).Amount = Npclist(NpcIndex).Invent.Objeto(NpcSlot).Amount + Objeto.Amount
                If Npclist(NpcIndex).Invent.Objeto(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
                    Npclist(NpcIndex).Invent.Objeto(NpcSlot).Amount = MAX_INVENTORY_OBJS
                End If
            End If

        End If

        Call UpdateUserInv(True, UserIndex, 0)
        Call WriteUpdateUserStats(UserIndex)
        Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
        Call WriteTradeOK(UserIndex)

        Call SubirSkill(UserIndex, eSkill.Comerciar)


        Exit Sub

hayerror:
        LogError("Error en Comercio: " & Err.Number & " Desc: " & Err.Description)


    End Sub

    Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
        UserList(UserIndex).flags.Comerciando = True
        Call WriteCommerceInit(UserIndex)
    End Sub

    Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        SlotEnNPCInv = 1
        Do Until Npclist(NpcIndex).Invent.Objeto(SlotEnNPCInv).ObjIndex = Objeto _
      And Npclist(NpcIndex).Invent.Objeto(SlotEnNPCInv).Amount + Cantidad <= MAX_INVENTORY_OBJS

            SlotEnNPCInv = SlotEnNPCInv + 1
            If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do

        Loop

        If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then

            SlotEnNPCInv = 1

            Do Until Npclist(NpcIndex).Invent.Objeto(SlotEnNPCInv).ObjIndex = 0

                SlotEnNPCInv = SlotEnNPCInv + 1
                If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do

            Loop

            If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1

        End If

    End Function

    Private Function Descuento(ByVal UserIndex As Integer) As Single
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        Descuento = 1 + UserList(UserIndex).Stats.UserSkills(eSkill.Comerciar) / 100
    End Function

    ''
    ' Send the inventory of the Npc to the user
    '
    ' @param userIndex The index of the User
    ' @param npcIndex The index of the NPC

    Private Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '*************************************************
        'Author: Nacho (Integer)
        'Last Modified: 06/14/08
        'Last Modified By: Nicolás Ezequiel Bouhid (NicoNZ)
        '*************************************************
        Dim Slot As Byte
        Dim val As Single

        For Slot = 1 To MAX_INVENTORY_SLOTS
            If Npclist(NpcIndex).Invent.Objeto(Slot).ObjIndex > 0 Then
                Dim thisObj As obj
                thisObj.ObjIndex = Npclist(NpcIndex).Invent.Objeto(Slot).ObjIndex
                thisObj.Amount = Npclist(NpcIndex).Invent.Objeto(Slot).Amount
                val = (ObjDataArr(Npclist(NpcIndex).Invent.Objeto(Slot).ObjIndex).valor) / Descuento(UserIndex)

                Call WriteChangeNPCInventorySlot(UserIndex, Slot, thisObj, val)
            Else
                Dim DummyObj As obj
                Call WriteChangeNPCInventorySlot(UserIndex, Slot, DummyObj, 0)
            End If
        Next Slot
    End Sub

    ''
    ' Devuelve el valor de venta del objeto
    '
    ' @param valor  El valor de compra de objeto

    Public Function SalePrice(ByVal valor As Long) As Single
        '*************************************************
        'Author: Nicolás (NicoNZ)
        '
        '*************************************************
        SalePrice = valor / REDUCTOR_PRECIOVENTA
    End Function

End Module
