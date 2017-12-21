Option Explicit On

Module modConWSAPI
    '**************************************************************
    ' WSKSOCK.bas
    '
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

    'date stamp: sept 1, 1996 (for version control, please don't remove)

    'Visual Basic 4.0 Winsock "Header"
    '   Alot of the information contained inside this file was originally
    '   obtained from ALT.WINSOCK.PROGRAMMING and most of it has since been
    '   modified in some way.
    '
    'Disclaimer: This file is public domain, updated periodically by
    '   Topaz, SigSegV@mail.utexas.edu, Use it at your own risk.
    '   Neither myself(Topaz) or anyone related to alt.programming.winsock
    '   may be held liable for its use, or misuse.
    '
    'Declare check Aug 27, 1996. (Topaz, SigSegV@mail.utexas.edu)
    '   All 16 bit declarations appear correct, even the odd ones that
    '   pass longs inplace of in_addr and char buffers. 32 bit functions
    '   also appear correct. Some are declared to return integers instead of
    '   longs (breaking MS's rules.) however after testing these functions I
    '   have come to the conclusion that they do not work properly when declared
    '   following MS's rules.
    '
    'NOTES:
    '   (1) I have never used WS_SELECT (select), therefore I must warn that I do
    '       not know if fd_set and timeval are properly defined.
    '   (2) Alot of the functions are declared with "buf as any", when calling these
    '       functions you may either pass strings, byte arrays or UDT's. For 32bit I
    '       I recommend Byte arrays and the use of memcopy to copy the data back out
    '   (3) The async functions (wsaAsync*) require the use of a message hook or
    '       message window control to capture messages sent by the winsock stack. This
    '       is not to be confused with a CallBack control, The only function that uses
    '       callbacks is WSASetBlockingHook()
    '   (4) Alot of "helper" functions are provided in the file for various things
    '       before attempting to figure out how to call a function, look and see if
    '       there is already a helper function for it.
    '   (5) Data types (hostent etc) have kept there 16bit definitions, even under 32bit
    '       windows due to the problem of them not working when redfined following the
    '       suggested rules.


    Public Const FD_SETSIZE = 64
    Public Structure fd_set
        Dim fd_count As Integer
        Dim fd_array() As Integer
    End Structure

    Public Structure timeval
        Dim tv_sec As Long
        Dim tv_usec As Long
    End Structure

    Public Structure HostEnt
        Dim h_name As Long
        Dim h_aliases As Long
        Dim h_addrtype As Integer
        Dim h_length As Integer
        Dim h_addr_list As Long
    End Structure
    Public Const hostent_size = 16

    Public Structure servent
        Dim s_name As Long
        Dim s_aliases As Long
        Dim s_port As Integer
        Dim s_proto As Long
    End Structure
    Public Const servent_size = 14

    Public Structure protoent
        Dim p_name As Long
        Dim p_aliases As Long
        Dim p_proto As Integer
    End Structure
    Public Const protoent_size = 10

    Public Const IPPROTO_TCP = 6
    Public Const IPPROTO_UDP = 17

    Public Const INADDR_NONE = &HFFFFFFFF
    Public Const INADDR_ANY = &H0

    Public Structure sockaddr
        Dim sin_family As Integer
        Dim sin_port As Integer
        Dim sin_addr As Long
        Dim sin_zero As String
    End Structure
    Public Const sockaddr_size = 16
    Public saZero As sockaddr


    Public Const WSA_DESCRIPTIONLEN = 256
    Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

    Public Const WSA_SYS_STATUS_LEN = 128
    Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

    Public Structure WSADataType
        Dim wVersion As Integer
        Dim wHighVersion As Integer
        Dim szDescription As String
        Dim szSystemStatus As String
        Dim iMaxSockets As Integer
        Dim iMaxUdpDg As Integer
        Dim lpVendorInfo As Long
    End Structure

    'Agregado por Maraxus
    Public Structure WSABUF
        Dim dwBufferLen As Long
        Dim lpBuffer As Long
    End Structure

    'Agregado por Maraxus
    Public Structure FLOWSPEC
        Dim TokenRate As Long     'In Bytes/sec
        Dim TokenBucketSize As Long     'In Bytes
        Dim PeakBandwidth As Long     'In Bytes/sec
        Dim Latency As Long     'In microseconds
        Dim DelayVariation As Long     'In microseconds
        Dim ServiceType As Integer  'Guaranteed, Predictive,
        'Best Effort, etc.
        Dim MaxSduSize As Long     'In Bytes
        Dim MinimumPolicedSize As Long     'In Bytes
    End Structure

    'Agregado por Maraxus
    Public Const WSA_FLAG_OVERLAPPED = &H1

    'Agregados por Maraxus
    Public Const CF_ACCEPT = &H0
    Public Const CF_REJECT = &H1

    'Agregado por Maraxus
    Public Const SD_RECEIVE As Long = &H0&
    Public Const SD_SEND As Long = &H1&
    Public Const SD_BOTH As Long = &H2&

    Public Const INVALID_SOCKET = -1
    Public Const SOCKET_ERROR = -1

    Public Const SOCK_STREAM = 1
    Public Const SOCK_DGRAM = 2

    Public Const MAXGETHOSTSTRUCT = 1024

    Public Const AF_INET = 2
    Public Const PF_INET = 2

    Public Structure LingerType
        Dim l_onoff As Integer
        Dim l_linger As Integer
    End Structure
    ' Windows Sockets definitions of regular Microsoft C error constants
    Public Const WSAEINTR = 10004
    Public Const WSAEBADF = 10009
    Public Const WSAEACCES = 10013
    Public Const WSAEFAULT = 10014
    Public Const WSAEINVAL = 10022
    Public Const WSAEMFILE = 10024
    ' Windows Sockets definitions of regular Berkeley error constants
    Public Const WSAEWOULDBLOCK = 10035
    Public Const WSAEINPROGRESS = 10036
    Public Const WSAEALREADY = 10037
    Public Const WSAENOTSOCK = 10038
    Public Const WSAEDESTADDRREQ = 10039
    Public Const WSAEMSGSIZE = 10040
    Public Const WSAEPROTOTYPE = 10041
    Public Const WSAENOPROTOOPT = 10042
    Public Const WSAEPROTONOSUPPORT = 10043
    Public Const WSAESOCKTNOSUPPORT = 10044
    Public Const WSAEOPNOTSUPP = 10045
    Public Const WSAEPFNOSUPPORT = 10046
    Public Const WSAEAFNOSUPPORT = 10047
    Public Const WSAEADDRINUSE = 10048
    Public Const WSAEADDRNOTAVAIL = 10049
    Public Const WSAENETDOWN = 10050
    Public Const WSAENETUNREACH = 10051
    Public Const WSAENETRESET = 10052
    Public Const WSAECONNABORTED = 10053
    Public Const WSAECONNRESET = 10054
    Public Const WSAENOBUFS = 10055
    Public Const WSAEISCONN = 10056
    Public Const WSAENOTCONN = 10057
    Public Const WSAESHUTDOWN = 10058
    Public Const WSAETOOMANYREFS = 10059
    Public Const WSAETIMEDOUT = 10060
    Public Const WSAECONNREFUSED = 10061
    Public Const WSAELOOP = 10062
    Public Const WSAENAMETOOLONG = 10063
    Public Const WSAEHOSTDOWN = 10064
    Public Const WSAEHOSTUNREACH = 10065
    Public Const WSAENOTEMPTY = 10066
    Public Const WSAEPROCLIM = 10067
    Public Const WSAEUSERS = 10068
    Public Const WSAEDQUOT = 10069
    Public Const WSAESTALE = 10070
    Public Const WSAEREMOTE = 10071
    ' Extended Windows Sockets error constant definitions
    Public Const WSASYSNOTREADY = 10091
    Public Const WSAVERNOTSUPPORTED = 10092
    Public Const WSANOTINITIALISED = 10093
    Public Const WSAHOST_NOT_FOUND = 11001
    Public Const WSATRY_AGAIN = 11002
    Public Const WSANO_RECOVERY = 11003
    Public Const WSANO_DATA = 11004
    Public Const WSANO_ADDRESS = 11004
    '---ioctl Constants
    Public Const FIONREAD = &H8004667F
    Public Const FIONBIO = &H8004667E
    Public Const FIOASYNC = &H8004667D


    '---Windows System Functions
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Object, Src As Object, ByVal cb&)
    Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Object) As Long
    '---async notification constants
    Public Const SOL_SOCKET = &HFFFF&
    Public Const SO_LINGER = &H80&
    Public Const SO_RCVBUFFER = &H1002&             ' Agregado por Maraxus
    Public Const SO_SNDBUFFER = &H1001&              ' Agregado por Maraxus
    Public Const SO_CONDITIONAL_ACCEPT = &H3002&    ' Agregado por Maraxus
    Public Const FD_READ = &H1&
    Public Const FD_WRITE = &H2&
    Public Const FD_OOB = &H4&
    Public Const FD_ACCEPT = &H8&
    Public Const FD_CONNECT = &H10&
    Public Const FD_CLOSE = &H20&
'---SOCKET FUNCTIONS
    Public Declare Function accept Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, AddrLen As Long) As Long
    Public Declare Function bind Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function apiclosesocket Lib "wsock32.dll" Alias "closesocket" (ByVal S As Long) As Long
    Public Declare Function connect Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function ioctlsocket Lib "wsock32.dll" (ByVal S As Long, ByVal Cmd As Long, argp As Long) As Long
    Public Declare Function getpeername Lib "wsock32.dll" (ByVal S As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockname Lib "wsock32.dll" (ByVal S As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, optval As Object, optlen As Long) As Long
    Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
    Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "wsock32.dll" (ByVal S As Long, ByVal backlog As Long) As Long
    Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
    Public Declare Function recv Lib "wsock32.dll" (ByVal S As Long, ByRef buf As Object, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare Function recvfrom Lib "wsock32.dll" (ByVal S As Long, buf As Object, ByVal buflen As Long, ByVal flags As Long, from As sockaddr, fromlen As Long) As Long
    Public Declare Function ws_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
    Public Declare Function send Lib "wsock32.dll" (ByVal S As Long, buf As Object, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare Function sendto Lib "wsock32.dll" (ByVal S As Long, buf As Object, ByVal buflen As Long, ByVal flags As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
    Public Declare Function setsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, optval As Object, ByVal optlen As Long) As Long
    Public Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As Long
    Public Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
'---DATABASE FUNCTIONS
    Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
    Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
    Public Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
    Public Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Public Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Long) As Long
    Public Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
    Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
    Public Declare Sub WSASetLastError Lib "wsock32.dll" (ByVal iError As Long)
    Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
    Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
    Public Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long
    Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
    Public Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Object, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Object, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Object, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Number As Long, buf As Object, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Object, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Object, ByVal buflen As Long) As Long
    Public Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
    Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Public Declare Function WSARecvEx Lib "wsock32.dll" (ByVal S As Long, buf As Object, ByVal buflen As Long, ByVal flags As Long) As Long
    'Agregado por Maraxus
    Declare Function WSAAccept Lib "ws2_32.DLL" (ByVal S As Long, pSockAddr As sockaddr, AddrLen As Long, ByVal lpfnCondition As Long, ByVal dwCallbackData As Long) As Long
    Public Const SOMAXCONN As Long = &H7FFFFFFF            ' Agregado por Maraxus


    'SOME STUFF I ADDED
    Public MySocket%
    Public SockReadBuffer$
    Public Const WSA_NoName = "Unknown"
    Public WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled


    Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
        If (lParam And &HFFFF&) > &H7FFF Then
            WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
        Else
            WSAGetSelectEvent = lParam And &HFFFF&
        End If
    End Function



    Sub EndWinsock()
        Dim Ret&
        If WSAIsBlocking() Then
            Ret = WSACancelBlockingCall()
        End If
        Ret = WSACleanup()
        WSAStartedUp = False
    End Sub

    Public Function GetAscIP(ByVal inn As Long) As String
#If Win32 Then
        Dim nStr&
#Else
        Dim nStr%
#End If
        Dim lpStr&
        Dim retString As String

        lpStr = inet_ntoa(inn)
        If lpStr Then
            nStr = lstrlen(lpStr)
            If nStr > 32 Then nStr = 32
            MemCopy(retString, lpStr, nStr)
            retString = Left$(retString, nStr)
            GetAscIP = retString
        Else
            GetAscIP = "255.255.255.255"
        End If
    End Function


    'this function should work on 16 and 32 bit systems
    Function GetWSAErrorString(ByVal errnum&) As String
        On Error Resume Next
        Select Case errnum
            Case 10004 : GetWSAErrorString = "Interrupted system call."
            Case 10009 : GetWSAErrorString = "Bad file number."
            Case 10013 : GetWSAErrorString = "Permission Denied."
            Case 10014 : GetWSAErrorString = "Bad Address."
            Case 10022 : GetWSAErrorString = "Invalid Argument."
            Case 10024 : GetWSAErrorString = "Too many open files."
            Case 10035 : GetWSAErrorString = "Operation would block."
            Case 10036 : GetWSAErrorString = "Operation now in progress."
            Case 10037 : GetWSAErrorString = "Operation already in progress."
            Case 10038 : GetWSAErrorString = "Socket operation on nonsocket."
            Case 10039 : GetWSAErrorString = "Destination address required."
            Case 10040 : GetWSAErrorString = "Message too long."
            Case 10041 : GetWSAErrorString = "Protocol wrong type for socket."
            Case 10042 : GetWSAErrorString = "Protocol not available."
            Case 10043 : GetWSAErrorString = "Protocol not supported."
            Case 10044 : GetWSAErrorString = "Socket type not supported."
            Case 10045 : GetWSAErrorString = "Operation not supported on socket."
            Case 10046 : GetWSAErrorString = "Protocol family not supported."
            Case 10047 : GetWSAErrorString = "Address family not supported by protocol family."
            Case 10048 : GetWSAErrorString = "Address already in use."
            Case 10049 : GetWSAErrorString = "Can't assign requested address."
            Case 10050 : GetWSAErrorString = "Network is down."
            Case 10051 : GetWSAErrorString = "Network is unreachable."
            Case 10052 : GetWSAErrorString = "Network dropped connection."
            Case 10053 : GetWSAErrorString = "Software caused connection abort."
            Case 10054 : GetWSAErrorString = "Connection reset by peer."
            Case 10055 : GetWSAErrorString = "No buffer space available."
            Case 10056 : GetWSAErrorString = "Socket is already connected."
            Case 10057 : GetWSAErrorString = "Socket is not connected."
            Case 10058 : GetWSAErrorString = "Can't send after socket shutdown."
            Case 10059 : GetWSAErrorString = "Too many references: can't splice."
            Case 10060 : GetWSAErrorString = "Connection timed out."
            Case 10061 : GetWSAErrorString = "Connection refused."
            Case 10062 : GetWSAErrorString = "Too many levels of symbolic links."
            Case 10063 : GetWSAErrorString = "File name too long."
            Case 10064 : GetWSAErrorString = "Host is down."
            Case 10065 : GetWSAErrorString = "No route to host."
            Case 10066 : GetWSAErrorString = "Directory not empty."
            Case 10067 : GetWSAErrorString = "Too many processes."
            Case 10068 : GetWSAErrorString = "Too many users."
            Case 10069 : GetWSAErrorString = "Disk quota exceeded."
            Case 10070 : GetWSAErrorString = "Stale NFS file handle."
            Case 10071 : GetWSAErrorString = "Too many levels of remote in path."
            Case 10091 : GetWSAErrorString = "Network subsystem is unusable."
            Case 10092 : GetWSAErrorString = "Winsock DLL cannot support this application."
            Case 10093 : GetWSAErrorString = "Winsock not initialized."
            Case 10101 : GetWSAErrorString = "Disconnect."
            Case 11001 : GetWSAErrorString = "Host not found."
            Case 11002 : GetWSAErrorString = "Nonauthoritative host not found."
            Case 11003 : GetWSAErrorString = "Nonrecoverable error."
            Case 11004 : GetWSAErrorString = "Valid name, no data record of requested type."
            Case Else
        End Select
    End Function




    Public Function StartWinsock(sDescription As String) As Boolean
        Dim StartupData As WSADataType
        If Not WSAStartedUp Then
            If Not WSAStartup(&H202, StartupData) Then  'Use sockets v2.2 instead of 1.1 (Maraxus)
                WSAStartedUp = True
                sDescription = StartupData.szDescription
            Else
                WSAStartedUp = False
            End If
        End If
        StartWinsock = WSAStartedUp
    End Function



End Module
