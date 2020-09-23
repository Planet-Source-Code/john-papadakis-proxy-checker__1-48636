Attribute VB_Name = "Module1"
Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Long
    iMaxUdpDg As Long
    lpVendorInfo As Long
End Type
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequested As Integer, lpWSAData _
    As WSADATA) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Const AF_INET = 2
Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal length As Long, ByVal _
    protocol As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source _
    As Any, ByVal length As Long)
Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal _
    lpString2 As Any) As Long

Public Const ICC_INTERNET_CLASSES = &H800
Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
    ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x _
    As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, _
    ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Const WC_IPADDRESS = "SysIPAddress32"
Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg _
    As Long, wParam As Any, lParam As Any) As Long
Public Const IPM_ISBLANK = &H469

 Function MakeLongOfInternetAddress(ByVal sIP As String) As Long

    Dim sFirst As String
    Dim sSecond As String
    Dim sThird As String
    Dim sForth As String
    Dim sAddress() As String
    
    
    sAddress = Split(Trim(sIP), ".")
        
    
    sFirst = Hex$(CLng(sAddress(0)))
    
    sFirst = String$(2 - Len(sFirst), "0") & sFirst
    
    sSecond = Hex$(CLng(sAddress(1)))
    sSecond = String$(2 - Len(sSecond), "0") & sSecond
    sThird = Hex$(CLng(sAddress(2)))
    sThird = String$(2 - Len(sThird), "0") & sThird
    sForth = Hex$(CLng(sAddress(3)))
    sForth = String$(2 - Len(sForth), "0") & sForth
    
    MakeLongOfInternetAddress = CLng("&H" & sFirst & sSecond & sThird & sForth)
   
End Function



 Function ResolveHostName(ByVal sIP As String) As String
    Dim ipAddress_h As Long
    Dim ipAddress_n As Long
    Dim sockinfo As WSADATA
    Dim pHostinfo As Long
    Dim hostinfo As HOSTENT
    Dim domainName As String
    Dim retval As Long
    
    ipAddress_h = MakeLongOfInternetAddress(sIP)
 
    retval = WSAStartup(&H202, sockinfo)
    If retval <> 0 Then
        Debug.Print "ERROR: Attempt to open Winsock failed: error"; retval
        Exit Function
    End If
    
    
    ipAddress_n = htonl(ipAddress_h)
    
    pHostinfo = gethostbyaddr(ipAddress_n, 4, AF_INET)
    If pHostinfo = 0 Then
        Debug.Print "Could not find a host with the specified IP address. " & sIP
    Else
    
        CopyMemory hostinfo, ByVal pHostinfo, Len(hostinfo)
    
        domainName = Space(lstrlen(hostinfo.h_name))
        retval = lstrcpy(domainName, hostinfo.h_name)
    End If
    
    ResolveHostName = domainName
  
    retval = WSACleanup()
End Function





Public Function CheckConnection() As Boolean
Dim result As Boolean
    result = InternetGetConnectedState(0&, 0&)
    If result = False Then
        CheckConnection = False
    Else
        CheckConnection = True
    End If
End Function

