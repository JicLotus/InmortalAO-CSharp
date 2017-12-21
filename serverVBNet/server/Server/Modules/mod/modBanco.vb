﻿Option Explicit On

Module modBanco
    '**************************************************************
    ' modBanco.bas - Handles the character's bank accounts.
    '
    ' Implemented by Kevin Birmingham (NEB)
    ' kbneb@hotmail.com
    '**************************************************************

    '**************************************************************************
    'This program is free software; you can redistribute it and/or modify
    'it under the terms of the Affero General Public License;
    'either version 1 of the License, or any later version.
    '
    'This program is distributed in the hope that it will be useful,
    'but WITHOUT ANY WARRANTY; without even the implied warranty of
    'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    'Affero General Public License for more details.
    '
    'You should have received a copy of the Affero General Public License
    'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
    '**************************************************************************



    Sub IniciarDeposito(ByVal UserIndex As Integer, ByVal goliath As Boolean)
        On Error GoTo Errhandler
        If goliath Then
            Call WriteBankInit(UserIndex, 1)
        Else
            'Hacemos un Update del inventario del usuario
            Call UpdateBanUserInv(True, UserIndex, 0)
            'Actualizamos el dinero
            Call WriteUpdateUserStats(UserIndex)
            'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
            Call WriteBankInit(UserIndex, 0)

            UserList(UserIndex).flags.Comerciando = True

        End If

Errhandler:

    End Sub

    Sub SendBanObj(UserIndex As Integer, Slot As Byte, Objecto As UserObj)

        UserList(UserIndex).BancoInvent.Objeto(Slot) = Objecto

        Call WriteChangeBankSlot(UserIndex, Slot)

    End Sub

    Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

        Dim NullObj As UserObj
        Dim loopC As Byte

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If UserList(UserIndex).BancoInvent.Objeto(Slot).ObjIndex > 0 Then
                Call SendBanObj(UserIndex, Slot, UserList(UserIndex).BancoInvent.Objeto(Slot))
            Else
                Call SendBanObj(UserIndex, Slot, NullObj)
            End If

        Else

            'Actualiza todos los slots
            For loopC = 1 To MAX_BANCOINVENTORY_SLOTS

                'Actualiza el inventario
                If UserList(UserIndex).BancoInvent.Objeto(loopC).ObjIndex > 0 Then
                    Call SendBanObj(UserIndex, loopC, UserList(UserIndex).BancoInvent.Objeto(loopC))
                Else

                    Call SendBanObj(UserIndex, loopC, NullObj)

                End If

            Next loopC

        End If

    End Sub

    Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
        On Error GoTo Errhandler


        If Cantidad < 1 Then Exit Sub


        Call WriteUpdateUserStats(UserIndex)


        If UserList(UserIndex).BancoInvent.Objeto(i).Amount > 0 Then
            If Cantidad > UserList(UserIndex).BancoInvent.Objeto(i).Amount Then Cantidad = UserList(UserIndex).BancoInvent.Objeto(i).Amount
            'Agregamos el obj que compro al inventario
            Call UserReciveObj(UserIndex, CInt(i), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el banco
            Call UpdateBanUserInv(True, UserIndex, 0)
        End If
        'Actualizamos la ventana de comercio
        Call UpdateVentanaBanco(UserIndex)


Errhandler:

    End Sub

    Sub UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

        Dim Slot As Integer
        Dim obji As Integer


        If UserList(UserIndex).BancoInvent.Objeto(ObjIndex).Amount <= 0 Then Exit Sub

        obji = UserList(UserIndex).BancoInvent.Objeto(ObjIndex).ObjIndex


        '¿Ya tiene un objeto de este tipo?
        Slot = 1
        Do Until UserList(UserIndex).Invent.Objeto(Slot).ObjIndex = obji And
   UserList(UserIndex).Invent.Objeto(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS

            Slot = Slot + 1
            If Slot > MAX_INVENTORY_SLOTS Then
                Exit Do
            End If
        Loop

        'Sino se fija por un slot vacio
        If Slot > MAX_INVENTORY_SLOTS Then
            Slot = 1
            Do Until UserList(UserIndex).Invent.Objeto(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > MAX_INVENTORY_SLOTS Then
                    Call WriteConsoleMsg(1, UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Loop
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
        End If



        'Mete el obj en el slot
        If UserList(UserIndex).Invent.Objeto(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then

            'Menor que MAX_INV_OBJS
            UserList(UserIndex).Invent.Objeto(Slot).ObjIndex = obji
            UserList(UserIndex).Invent.Objeto(Slot).Amount = UserList(UserIndex).Invent.Objeto(Slot).Amount + Cantidad

            Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Else
            Call WriteConsoleMsg(1, UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
        End If


    End Sub

    Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)



        Dim ObjIndex As Integer
        ObjIndex = UserList(UserIndex).BancoInvent.Objeto(Slot).ObjIndex

        'Quita un Obj

        UserList(UserIndex).BancoInvent.Objeto(Slot).Amount = UserList(UserIndex).BancoInvent.Objeto(Slot).Amount - Cantidad

        If UserList(UserIndex).BancoInvent.Objeto(Slot).Amount <= 0 Then
            UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems - 1
            UserList(UserIndex).BancoInvent.Objeto(Slot).ObjIndex = 0
            UserList(UserIndex).BancoInvent.Objeto(Slot).Amount = 0
        End If



    End Sub

    Sub UpdateVentanaBanco(ByVal UserIndex As Integer)
        Call WriteBankOK(UserIndex)
    End Sub

    Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)
        On Error GoTo Errhandler
        If UserList(UserIndex).Invent.Objeto(Item).Amount > 0 And Cantidad > 0 Then
            If Cantidad > UserList(UserIndex).Invent.Objeto(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Objeto(Item).Amount

            'Agregamos el obj que deposita al banco
            Call UserDejaObj(UserIndex, CInt(Item), Cantidad)

            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)

            'Actualizamos el inventario del banco
            Call UpdateBanUserInv(True, UserIndex, 0)
        End If

        'Actualizamos la ventana del banco
        Call UpdateVentanaBanco(UserIndex)
Errhandler:
    End Sub

    Sub UserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
        Dim Slot As Integer
        Dim obji As Integer

        If Cantidad < 1 Then Exit Sub

        obji = UserList(UserIndex).Invent.Objeto(ObjIndex).ObjIndex

        '¿Ya tiene un objeto de este tipo?
        Slot = 1
        Do Until UserList(UserIndex).BancoInvent.Objeto(Slot).ObjIndex = obji And
        UserList(UserIndex).BancoInvent.Objeto(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1

            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
        Loop

        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Slot = 1
            Do Until UserList(UserIndex).BancoInvent.Objeto(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > MAX_BANCOINVENTORY_SLOTS Then
                    Call WriteConsoleMsg(1, UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Loop

            UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
        End If

        If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
            'Mete el obj en el slot
            If UserList(UserIndex).BancoInvent.Objeto(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then

                'Menor que MAX_INV_OBJS
                UserList(UserIndex).BancoInvent.Objeto(Slot).ObjIndex = obji
                UserList(UserIndex).BancoInvent.Objeto(Slot).Amount = UserList(UserIndex).BancoInvent.Objeto(Slot).Amount + Cantidad

                Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
            Else
                Call WriteConsoleMsg(1, UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End Sub

End Module
