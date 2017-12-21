Option Explicit On
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading

Module modConexion


    Public Class handleClinet
        Public clientSocket As TcpClient
        Private ctThread As Threading.Thread
        Private thUpdateadorDePoss As Thread


        Private thRecibirInformacion As Thread

        Dim userIndex As Integer
        Public puedeEscribirOutGoing As Boolean
        Public xAnterior As Byte
        Public yAnterior As Byte
        Private secondaryBuffer As clsByteQueue

        Private incomingDataBuffer As clsByteQueue

        Private tercerBuffer As clsByteQueue

        Private bytesFrom As Byte()

        Private cantidadDeBytesRecibidos As Integer

        Private conexionActiva As Boolean

        Public keepAliveByte(0) As Byte

        Public networkStream As NetworkStream

        Public Sub startClient(ByVal inClientSocket As TcpClient, ByVal userIndex As Integer)


            keepAliveByte(0) = 120


            Me.conexionActiva = True

            Me.puedeEscribirOutGoing = True
            Me.clientSocket = inClientSocket

            networkStream = clientSocket.GetStream()

            Me.userIndex = userIndex

            tercerBuffer = New clsByteQueue()
            tercerBuffer.Class_Initialize()

            secondaryBuffer = New clsByteQueue()
            secondaryBuffer.Class_Initialize()

            incomingDataBuffer = New clsByteQueue()
            incomingDataBuffer.Class_Initialize()

            ctThread = New Threading.Thread(AddressOf doChat)
            ctThread.Start()

            thRecibirInformacion = New Threading.Thread(AddressOf recibirInformacionYEscribirla)
            thRecibirInformacion.Start()

            'thUpdateadorDePoss = New Thread(AddressOf updatePosition)
            'thUpdateadorDePoss.Start()

        End Sub

        Public Sub cerrarConexionCliente()
            CloseSocket(Me.userIndex)
            clientSocket.Close()
        End Sub

        Public Sub updatePosition()

            Dim bytePacket(3) As Byte

            bytePacket(0) = ServerPacketID.posUpdate

            'While True
            If UserList(Me.userIndex).flags.UserLogged Then
                bytePacket(1) = Convert.ToByte(UserList(Me.userIndex).Pos.x)
                bytePacket(2) = Convert.ToByte(UserList(Me.userIndex).Pos.Y)

                If bytePacket(1) = Me.xAnterior And bytePacket(2) = Me.yAnterior Then
                    Exit Sub
                End If

                If bytePacket(1) <> 0 And bytePacket(2) <> 0 Then
                    sendInmediatePacketWalk(bytePacket)
                End If

                Me.xAnterior = bytePacket(1)
                Me.yAnterior = bytePacket(2)


            End If
            '    Thread.Sleep(25)
            ' End While

        End Sub


        Public Sub sendInmediatePacketWalk(data As Byte())

            If clientSocket.Connected = False Then
                Exit Sub
            End If

            tercerBuffer.WriteBlock(data)

        End Sub

        Public Sub sendInmediatePacket(data As Byte())

            If clientSocket.Connected = False Then
                Exit Sub
            End If

            secondaryBuffer.WriteBlock(data)

        End Sub

        Public Sub flushOutGoingData()

            Try

                If clientSocket.Connected = True Then

                    With UserList(Me.userIndex)

                        Dim arrayPrimaryBuffer() As Byte
                        Dim arraySecondaryBuffer() As Byte
                        Dim arrayTercerBuffer() As Byte


                        arrayPrimaryBuffer = UserList(Me.userIndex).outgoingData.get_Actual_Data
                        arraySecondaryBuffer = secondaryBuffer.get_Actual_Data
                        'Me.updatePosition()
                        arrayTercerBuffer = tercerBuffer.get_Actual_Data

                        Dim lenPrimBuf As Long
                        Dim lenSecBuf As Long
                        Dim lenTerBuf As Long

                        lenPrimBuf = arrayPrimaryBuffer.Length
                        lenSecBuf = arraySecondaryBuffer.Length
                        lenTerBuf = arrayTercerBuffer.Length


                        If lenSecBuf > 0 Then
                            networkStream.Write(arraySecondaryBuffer, 0, lenSecBuf)
                        End If

                        If lenPrimBuf > 0 Then
                            networkStream.Write(arrayPrimaryBuffer, 0, lenPrimBuf)
                        End If

                        If lenTerBuf > 0 Then
                            networkStream.Write(arrayTercerBuffer, 0, lenTerBuf)
                        End If


                        'If lenSecBuf > 0 Or lenPrimBuf > 0 Or lenTerBuf > 0 Then
                        networkStream.Write(keepAliveByte, 0, 1)
                        networkStream.Flush()
                        'End If

                    End With
                End If
            Catch ex As Exception
                Dim st As New StackTrace(True)
                st = New StackTrace(ex, True)
                Me.cerrarConexionCliente()
            End Try

        End Sub


        Private Sub recibirInformacionYEscribirla()
            'Dim networkStream As NetworkStream
            Dim bytesFrom(4096) As Byte

            While True
                Try

                    ReDim bytesFrom(4096)
                    'networkStream = clientSocket.GetStream()
                    cantidadDeBytesRecibidos = networkStream.Read(bytesFrom, 0, CInt(bytesFrom.Length))

                    If cantidadDeBytesRecibidos > 0 Then



                        ReDim Preserve bytesFrom(cantidadDeBytesRecibidos - 1)

                        With UserList(Me.userIndex)

                            Call .incomingData.WriteBlock(bytesFrom)

                            While UserList(userIndex).incomingData.length > 0 And Err.Number = 0
                                Call HandleIncomingData(Me.userIndex)
                            End While

                            If Err.Number <> 0 Then
                                Me.cerrarConexionCliente()
                                Exit While
                            End If

                        End With

                    Else

                        Me.cerrarConexionCliente()
                        Exit While

                    End If

                Catch ex As Exception
                    Me.cerrarConexionCliente()
                    Exit While
                End Try

            End While
        End Sub

        Private Sub doChat()
            While (True)
                Try

                    If UserList(Me.userIndex).ConnID = -1 Or Not clientSocket.Connected Then
                        Me.cerrarConexionCliente()
                        Exit While
                    End If

                    flushOutGoingData()


                    Thread.Sleep(100)

                Catch ex As Exception
                    Me.cerrarConexionCliente()
                    Exit While
                End Try

            End While

        End Sub
    End Class


    Public Sub IniciaWsApi()
        Call LogApiSock("IniciaWsApi")

        Dim port As Integer
        port = 7777
        Dim serverSocket As New TcpListener(Net.IPAddress.Any, port)
        Dim clientSocket As TcpClient

        Dim newIndex As Integer

        serverSocket.Start()

        While (True)
            clientSocket = serverSocket.AcceptTcpClient()

            newIndex = NextOpenUser()

            If newIndex < MaxUsers Then
                UserList(newIndex).ConnID = 1
                UserList(newIndex).ConnIDValida = True

                Dim client As New handleClinet

                'clientSocket.SendTimeout = 100
                UserList(newIndex).client = client
                client.startClient(clientSocket, newIndex)

                If newIndex > LastUser Then LastUser = newIndex
            End If

        End While


    End Sub


    Public Sub LogApiSock(ByVal str As String)
        On Error GoTo Errhandler

        Dim objReader As New System.IO.StreamWriter(Application.StartupPath & "\logs\wsapi.log")
        objReader.WriteLine(DateTime.Now & " " & DateTime.Now & " " & str)

        Exit Sub

Errhandler:

    End Sub




End Module
