Attribute VB_Name = "mdlTcpIPFramework"
Public Const COMMAND_ERROR = -1
Public Const RECV_ERROR = -1
Public Const NO_ERROR = 0

Public socketId As Long

'Global Variables for WINSOCK
Global state As Integer

Function CloseConnection() As Long

    CloseConnection = closesocket(socketId)
    
End Function

Function EndIt() As Long

    'Shutdown Winsock DLL
    EndIt = WSACleanup()

End Function

Function StartIt() As Long

    Dim StartUpInfo As WSAData
    
    'Version 1.1 (1*256 + 1) = 257
    'version 2.0 (2*256 + 0) = 512

    Version = GetConfigValue("FBS_Winsock_Ver") 'Originally used const - Winsock_Ver '512 'ActiveCell.FormulaR1C1
    
    'Initialize Winsock DLL
    StartIt = WSAStartup(Version, StartUpInfo)

End Function
 
Function OpenSocket(ByVal hostName As String, ByVal PortNumber As Integer) As Integer
   
    Dim I_SocketAddress As sockaddr_in
    Dim ipAddress As Long
    
    ipAddress = inet_addr(hostName)

    'Create a new socket
    socketId = socket(AF_INET, SOCK_STREAM, 0)
    If socketId = SOCKET_ERROR Then
        'MsgBox ("ERROR: socket = " + Str$(socketId))
        OpenSocket = COMMAND_ERROR
        Exit Function
    End If

    'Open a connection to a server

    I_SocketAddress.sin_family = AF_INET
    I_SocketAddress.sin_port = htons(PortNumber)
    I_SocketAddress.sin_addr = ipAddress
    I_SocketAddress.sin_zero = String$(8, 0)

    x = connect(socketId, I_SocketAddress, Len(I_SocketAddress))
    If socketId = SOCKET_ERROR Then
        'MsgBox ("ERROR: connect = " + Str$(x))
        OpenSocket = COMMAND_ERROR
        Exit Function
    End If

    OpenSocket = socketId

End Function

Function SendCommand(ByVal command As String) As Integer

    Dim strSend As String
    
    strSend = command + vbCrLf
    
    Count = send(socketId, ByVal strSend, Len(strSend), 0)
    
    If Count = SOCKET_ERROR Then
        'MsgBox ("ERROR: send = " + Str$(count))
        SendCommand = COMMAND_ERROR
        Exit Function
    End If
    
    SendCommand = NO_ERROR

End Function

Function RecvAscii(ByRef dataBuf As String, ByVal maxLength As Integer) As Integer

    Dim c As String * 1
    Dim length As Integer
    
    dataBuf = Space(maxLength)
    
    DoEvents
    
    Count = recv(socketId, dataBuf, Len(dataBuf), 0)
    
    If Count < 1 Then
        RecvAscii = RECV_ERROR
        dataBuf = Chr$(0)
    Else
        RecvAscii = NO_ERROR
        'Debug.Print Trim(recvBuf) 'dataBuf
    End If
    
End Function
