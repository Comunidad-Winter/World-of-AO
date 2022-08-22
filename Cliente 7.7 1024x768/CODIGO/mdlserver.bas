Attribute VB_Name = "mdlserver"
Public Const SERVER_PORT As Long = 7668
Public Const GWL_WNDPROC = (-4)

Public Type HostEnt

    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long

End Type

Public Declare Sub CopyMemoryIP _
                    Lib "kernel32" _
                        Alias "RtlMoveMemory" (hpvDest As Any, _
                                               ByVal hpvSource As Long, _
                                               ByVal cbCopy As Long)
Declare Function SetWindowLong _
                  Lib "user32" _
                      Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                              ByVal nIndex As Long, _
                                              ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc _
                  Lib "user32" _
                      Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                               ByVal hwnd As Long, _
                                               ByVal msg As Long, _
                                               ByVal wParam As Long, _
                                               ByVal lParam As Long) As Long
Public listenSocket As Long
Public Sockets As New Collection
Public IPAddresses As New Collection
Public SocketHandle As Long
Private PrevProc As Long
Dim Resto As String

'Function to retrieve the IP address
Public Function GetIPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim Host As HostEnt
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim i As Integer
    Dim sIPAddr As String

    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)

    CopyMemoryIP Host, lpHost, Len(Host)
    CopyMemoryIP dwIPAddr, Host.hAddrList, 4
    ReDim tmpIPAddr(1 To Host.hLen)
    CopyMemoryIP tmpIPAddr(1), dwIPAddr, Host.hLen

    For i = 1 To Host.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    GetIPAddress = mid$(sIPAddr, 1, Len(sIPAddr) - 1)

End Function

Public Sub StartSubclass(f As Form)
    PrevProc = SetWindowLong(f.hwnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Public Sub StopSubclass(f As Form)

    If PrevProc <> 0 Then SetWindowLong f.hwnd, GWL_WNDPROC, PrevProc

End Sub

Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

    If uMsg = WINSOCK_MESSAGE Then
        SocketHandle = wParam
        ProcessMessage wParam, lParam
    Else
        WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)

    End If

End Function

'Process socket messages
Public Function ProcessMessage(ByVal wParam As Long, _
                               ByVal lParam As Long)    'wParam = Socket Handle, lParam = connection message

    Select Case lParam

    Case FD_ACCEPT
        Dim tempSocket As Long, tempAddr As sockaddr
        tempSocket = accept(wParam, tempAddr, Len(tempAddr))
        AddSocket tempSocket, getascip(tempAddr.sin_addr)

        'AddText " connected..."
    Case FD_WRITE

    Case FD_READ
        Dim sTemp As String, lRet As Long, szBuf As String
        Do
            szBuf = String(256, 0)
            lRet = recv(wParam, ByVal szBuf, Len(szBuf), 0)

            If lRet > 0 Then sTemp = sTemp + Left$(szBuf, lRet)
        Loop Until lRet <= 0

        If LenB(sTemp) > 0 Then
            Separar sTemp

        End If

        'SendDataw wParam, "From Server: Thank you for sending me a message!"
    Case Else    'FD_CLOSE
        FormP.CerrarConexion
        FormP.Escuchar
        'AddText " disconnected..."
        RemoveSocket wParam

    End Select

End Function

Sub Separar(Cadena As String)

    On Error GoTo Final

    Dim Dato As String
    Cadena = Resto & Cadena
    Resto = ""

    Do Until Cadena = ""

        If Len(Replace(Cadena, "t#z@", ".")) < Len(Cadena) Then
            Dato = Left(Cadena, InStr(Cadena, "t#z@") - 1)
            FormP.DataArrival (Dato)
            Cadena = Right(Cadena, Len(Cadena) - Len(Dato) - 4)
        Else
            Resto = Cadena
            Exit Sub

        End If

    Loop
    Exit Sub
Final:

End Sub

'Here are some functions to keep a list of all the connections to the server
Public Sub AddSocket(ByVal s As Long, ByVal FromIP As String)
    On Local Error Resume Next

    'If FromIP <> "92.43.20.27" Then
    FormP.CerrarConexion
    'Form1.Escuchar
    'AddText " disconnected..."
    'RemoveSocket wParam
    Exit Sub
    'End If

    IPAddresses.Add FromIP, CStr(s)
    Sockets.Add s, CStr(s)

End Sub

Public Sub RemoveSocket(ByVal s As Long)
    On Local Error Resume Next
    IPAddresses.Remove CStr(s)
    Sockets.Remove CStr(s)

End Sub

