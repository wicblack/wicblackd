Attribute VB_Name = "modCrackd"
#Const FinalMode = 1
Option Explicit

Public Type TypeTibiaKey
 key(15) As Byte
End Type
' firstPacketByte is the first byte of the packet array
' crackd.dll functions expect to find the rest of the packet bytes after firstPacketByte
' firstKeyByte is the first byte of the 16bytes-key array
' crackd.dll functions expect to find the other 15 bytes of the key after firstKeyByte
' crackd.dll functions expect to receive a packet with size (8*n) + 2 bytes
'   (fill with random trash if required)
' crackd.dll functions expect to receive the packet size (in bytes)
'   in the first two bytes of the packet array
#If FinalMode Then

'Public Declare Function EncipherTibia Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte) As Long

'Public Declare Function DecipherTibia Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte) As Long

Public Declare Function EncipherTibiaProtected Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtected Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long
      
Public Declare Function EncipherTibiaProtectedSP Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtectedSP Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function GetTibiaCRC Lib _
    "crackd.dll" (ByRef firstPacketByte As Byte, ByVal uboundpacketMinus6 As Long) As Long
    
Public Declare Function BlackdForceWrite Lib _
    "crackd.dll" (ByVal Address As Long, ByRef mybuffer As Byte, ByVal mybuffersize As Long, ByVal hwndClientWindow As Long) As Long
    
    

#Else

'Public Declare Function EncipherTibia Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte) As Long

'Public Declare Function DecipherTibia Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte) As Long
    
Public Declare Function EncipherTibiaProtected Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtected Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long
    
Public Declare Function EncipherTibiaProtectedSP Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function DecipherTibiaProtectedSP Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, _
    ByRef firstKeyByte As Byte, ByVal uboundpacket As Long, ByVal uboundkey As Long) As Long

Public Declare Function GetTibiaCRC Lib _
    "C:\blackdProxy\crackd.dll" (ByRef firstPacketByte As Byte, ByVal uboundpacketMinus6 As Long) As Long

Public Declare Function BlackdForceWrite Lib _
    "C:\blackdProxy\crackd.dll" (ByVal Address As Long, ByRef mybuffer As Byte, ByVal mybuffersize As Long, ByVal hwndClientWindow As Long) As Long
    
'Public Declare Function GetTibiaCRC2 Lib _
'    "C:\blackdProxy\crackd2.dll" (ByRef firstPacketByte As Byte, ByVal uboundpacketMinus6 As Long) As Long
    
#End If

Public Declare Sub RtlMoveMemory Lib "Kernel32" ( _
    lpDest As Any, _
    lpSource As Any, _
    ByVal ByValcbCopy As Long)
    
Public packetKey() As TypeTibiaKey
Public loginPacketKey() As TypeTibiaKey
Public gotFirstLoginPacket() As Boolean
Public UseCrackd As Boolean
Public adrConnectionKey As Long
Public adrSelectedCharIndex As Long
Public adrLastPacket As Long
Public adrCharListPtr As Long
Public debugStrangeFail As String
Public MAXCHARACTERLEN As Long
Public manualDebugOrder As Long
Public GameServerDictionary As scripting.Dictionary  ' A dictionary server (string) -> IP (string)


Public Sub JustReadPID(idConnection As Integer)
 ' should be only used at login stage, in tibia 7.63+
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim status As Long
  Dim abyte As Byte
  Dim sucess As Long
  Dim i As Integer
  sucess = -3
  If (GameConnected(idConnection) = True) Then
    ' keys will be only read at login
    Exit Sub
  End If
  ProcessID(idConnection) = 0
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      status = Memory_ReadLong(adrConnected, tibiaclient)
      If (status <> 0) Then ' doing login
        sucess = 0
        ProcessID(idConnection) = tibiaclient
        Exit Do
      End If
    End If
  Loop
End Sub

Public Function readLoginTibiaKeyAtPID(idConnection As Integer, ProcessID As Long) As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim abyte As Byte
  Dim i As Integer
  If (ProcessID = -1) Then
    readLoginTibiaKeyAtPID = -1
  Else
    For i = 0 To 15
      abyte = Memory_ReadByte(adrConnectionKey + i, ProcessID)
      loginPacketKey(idConnection).key(i) = abyte
    Next i
    readLoginTibiaKeyAtPID = 0
  End If
  Exit Function
goterr:
  readLoginTibiaKeyAtPID = -1
End Function

Public Function readTibiaKeyAtPID(idConnection As Integer, ProcessID As Long) As Long
  Dim abyte As Byte
  Dim i As Integer
    For i = 0 To 15
      abyte = Memory_ReadByte(adrConnectionKey + i, ProcessID)
      packetKey(idConnection).key(i) = abyte
    Next i
  readTibiaKeyAtPID = 0
End Function

Public Function CompareLastPacket(ByVal pid As Long, ByRef packet() As Byte) As Boolean
  Dim res As Boolean
  Dim i As Long
  Dim lngp As Long
  Dim b As Byte
  Dim errmessage As String
  On Error GoTo cantdoit
  If UBound(packet) < 1 Then
    CompareLastPacket = False
    Exit Function
  End If
  res = True
  lngp = GetTheLong(packet(0), packet(1)) + 1
  For i = 0 To lngp
    b = Memory_ReadByte(adrLastPacket + i, pid)
    If b <> packet(i) Then
      res = False
      Exit For
    End If
  Next i
  CompareLastPacket = res
  Exit Function
cantdoit:
  errmessage = "Function failure : CompareLastPacket failed : Error number " & CStr(Err.Number) & " : " & Err.Description
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & errmessage
  LogOnFile "errors.txt", errmessage
  CompareLastPacket = False
End Function

Public Function GetLastPacket(pid As Long, lngp As Long) As String
  Dim res As Boolean
  Dim i As Long
  Dim b As Byte
  Dim errmessage As String
  Dim packetR() As Byte
  On Error GoTo cantdoit
  ReDim packetR(lngp)
  For i = 0 To lngp
    b = Memory_ReadByte(adrLastPacket + i, pid)
    packetR(i) = b
  Next i
  GetLastPacket = frmMain.showAsStr2(packetR, 0)
  Exit Function
cantdoit:
  GetLastPacket = "ERROR"
End Function

Public Sub UpdateProcessIDbyLastPacket(ByVal idConnection As Integer, ByRef packet() As Byte, Optional strIP As String = "")
  Dim tibiaclient As Long
  'Dim hWndDesktop As Long
  Dim status As Byte
  Dim abyte As Byte
  Dim sucess As Long
  Dim i As Integer
  Dim errmessage As String
  On Error GoTo goterr
  sucess = -2
  ProcessID(idConnection) = 0
  If AlternativeBinding <> 0 Then
    If strIP <> "" Then
      ProcessID(idConnection) = GetProcessIdFromIP(strIP)
      Exit Sub
    End If
  End If
  'hWndDesktop = GetDesktopWindow()
  tibiaclient = 0
  Do

    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
      
      If CompareLastPacket(tibiaclient, packet) = True Then
        ProcessID(idConnection) = tibiaclient
        sucess = 0
        Exit Do
      End If
    End If
  Loop
  If (sucess = 0) Then
    Exit Sub
  End If
  errmessage = "Warning on function UpdateProcessIDbyLastPacket : could not find match"
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & errmessage
  LogOnFile "errors.txt", errmessage
  Exit Sub
goterr:
  errmessage = "Function failure : UpdateProcessIDbyLastPacket could not match idconnection<->pid : Error number " & CStr(Err.Number) & " : " & Err.Description
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & errmessage
  LogOnFile "errors.txt", errmessage
End Sub


Public Function GiveProcessIDbyLastPacket(ByRef packet() As Byte, Optional strIP As String = "", Optional FromIP As String = "?", Optional part As String = "LOGIN1") As Long
  Dim tibiaclient As Long
  Dim hWndDesktop As Long
  Dim status As Byte
  Dim abyte As Byte
  Dim sucess As Long
  Dim i As Integer
  Dim errmessage As String
  Dim res As Long
  Dim comparing1 As String
  Dim comparing2 As String
  Dim tcount As Long
  Dim trivialRes As Long
  Dim packetSizeForComparing As Long
  #If FinalMode = 1 Then
  On Error GoTo goterr
  #End If
  tcount = 0
  
  If AlternativeBinding <> 0 Then
    If strIP <> "" Then
      GiveProcessIDbyLastPacket = GetProcessIdFromIP(strIP)
      Exit Function
    End If
  End If
  debugStrangeFail = ""
  
  debugStrangeFail = "WARNING on GiveProcessIDbyLastPacket . Doing a complete report:"
  res = 0
  sucess = -2
  packetSizeForComparing = GetTheLong(packet(0), packet(1)) ' fix since 11.7 : lets only compare first subpacket
  
  comparing1 = frmMain.showAsStr2(packet, 0, packetSizeForComparing + 1)
  
  'hWndDesktop = GetDesktopWindow()
  'debugStrangeFail = debugStrangeFail & vbCrLf & "GetDesktopWindow() returned " & CStr(hWndDesktop)
  debugStrangeFail = debugStrangeFail & vbCrLf & "Now trying to determine what client sent the packet that Blackd Proxy just received."
  debugStrangeFail = debugStrangeFail & vbCrLf & "BLACKDPROXY RECEIVED, from ip [" & FromIP & "] at " & part & " :" & comparing1
  tibiaclient = 0
  Do

    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      debugStrangeFail = debugStrangeFail & vbCrLf & "Found a total of " & CStr(tcount) & " Tibia client(s) opened"
      Exit Do
    Else
      trivialRes = tibiaclient
      tcount = tcount + 1
      comparing2 = GetLastPacket(tibiaclient, packetSizeForComparing + 1)
      debugStrangeFail = debugStrangeFail & vbCrLf & "CLIENT #" & CStr(tcount) & " HAVE SENT :" & comparing2
      If (comparing1 = comparing2) Then
        debugStrangeFail = debugStrangeFail & vbCrLf & " ...MATCH at pid " & CStr(tibiaclient)
        res = tibiaclient
        sucess = 0
        'Exit Do
      Else
        debugStrangeFail = debugStrangeFail & vbCrLf & " ...FAIL! at pid " & CStr(tibiaclient)
      End If
    End If
  Loop
  If (sucess = 0) Then
    debugStrangeFail = debugStrangeFail & vbCrLf & "Function worked fine."
  Else
    debugStrangeFail = debugStrangeFail & vbCrLf & "Function failed!"
    If tcount = 1 Then
        debugStrangeFail = debugStrangeFail & vbCrLf & "However, there is a trivial match since only 1 client was detected: " & CStr(trivialRes)
        sucess = 0
        res = trivialRes
    Else
        debugStrangeFail = debugStrangeFail & vbCrLf & "Please report to daniel@blackdtools.com"
    End If
  End If
  If (sucess = 0) Then
    GiveProcessIDbyLastPacket = res
    Exit Function
  End If
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & debugStrangeFail
  LogOnFile "errors.txt", debugStrangeFail
  GiveProcessIDbyLastPacket = 0
  Exit Function
goterr:
  errmessage = "Function failure : GiveProcessIDbyLastPacket could not match idconnection<->pid : Error number " & CStr(Err.Number) & " : " & Err.Description
  frmMain.txtPackets.Text = frmMain.txtPackets.Text & vbCrLf & errmessage
  LogOnFile "errors.txt", errmessage
  GiveProcessIDbyLastPacket = 0
End Function

Public Sub AddGameServer(ByVal ServerName As String, ByVal serverIPport As String)
  On Error GoTo goterr
  ' add item to dictionary
  Dim res As Boolean
  GameServerDictionary.item(ServerName) = serverIPport
  Exit Sub
goterr:
  LogOnFile "errors.txt", "Get error at AddGameServer : " & Err.Description
End Sub

Public Function GetIPandPortfromServerName(ByVal ServerName As String) As String
  On Error GoTo goterr
  ' get the IPandport from server name
  Dim aRes As String
  Dim res As Boolean
  Dim strBuildIt As String
  Dim b(3) As Byte
  Dim i As Long
  Dim lastI As Long
  Dim strTmp As String
  Dim pos1 As Long
  Dim pos2 As Long
  Dim pos3 As Long
  If GameServerDictionary.Exists(ServerName) = True Then
    GetIPandPortfromServerName = GameServerDictionary.item(ServerName)
  Else
    strTmp = GetIPofTibiaServer(ServerName)
    lastI = Len(strTmp)
    ' search the 3 points of the IP
    pos1 = InStr(1, strTmp, ".")
    If pos1 > 0 Then
        pos2 = InStr(pos1 + 1, strTmp, ".")
    Else
        GetIPandPortfromServerName = ""
        Exit Function
    End If
    If pos2 > 0 Then
        pos3 = InStr(pos2 + 1, strTmp, ".")
    Else
        GetIPandPortfromServerName = ""
        Exit Function
    End If
    b(0) = CByte(CLng(Left$(strTmp, pos1 - 1)))
    b(1) = CByte(CLng(Mid$(strTmp, pos1 + 1, pos2 - pos1 - 1)))
    b(2) = CByte(CLng(Mid$(strTmp, pos2 + 1, pos3 - pos2 - 1)))
    b(3) = CByte(CLng(Right$(strTmp, lastI - pos3)))
    strBuildIt = fixThreeDigits(b(0)) & "." & fixThreeDigits(b(1)) & "." & _
     fixThreeDigits(b(2)) & "." & fixThreeDigits(b(3)) & ":7171"
    GetIPandPortfromServerName = strBuildIt
  End If
  Exit Function
goterr:
  LogOnFile "errors.txt", "Got error at GetIPandPortfromServerName (" & ServerName & " ): " & Err.Description
  GetIPandPortfromServerName = ""
End Function


'Public Function GetProcessIdByAccount(strAccount As String) As Long
'    Dim res As Long
'    res = GetProcessIdFromAccount(strAccount)
'    If res <= 0 Then
'        res = -1
'    End If
'    GetProcessIdByAccount = res
'End Function

Public Function GetProcessIdByManualDebug() As Long
   Dim tibiaclient As Long
   Dim bc As Byte
   Dim c As Long
   c = 0
   tibiaclient = 0
   
   Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
        Debug.Print "#" & CStr(c) & " : " & CStr(tibiaclient)
        If manualDebugOrder = c Then
            GetProcessIdByManualDebug = tibiaclient
            Exit Function
        End If
        c = c + 1
    End If
  Loop

  GetProcessIdByManualDebug = -1
End Function

Public Function GetProcessIdByAdrConnected() As Long
   Dim tibiaclient As Long
   Dim bc As Byte
   Dim foundcount As Long
   Dim lastfound As Long
   Dim totalclients As Long
   Dim cantbeother As Long
   Dim cantbeotherBYTE As Byte
   foundcount = 0
   totalclients = 0
   cantbeother = 0
   Do
    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      If foundcount = 1 Then
        GetProcessIdByAdrConnected = lastfound
      ElseIf foundcount = 0 Then
        If totalclients = 1 Then
            Debug.Print "Warning: only 1 tibiaclient, with connection status " & GoodHex(cantbeotherBYTE)
            GetProcessIdByAdrConnected = cantbeother
        Else
            GetProcessIdByAdrConnected = -1
        End If
      Else
        GetProcessIdByAdrConnected = -2
      End If
      Exit Function
    Else
        totalclients = totalclients + 1
        bc = Memory_ReadLong(adrConnected, tibiaclient)
        If TibiaVersionLong >= 980 Then
            If ((bc = &H5) Or (bc = &H6) Or (bc = &H8)) Then
                lastfound = tibiaclient
                foundcount = foundcount + 1
            End If
        Else
            If ((bc = &H5) Or (bc = &H6)) Then
                lastfound = tibiaclient
                foundcount = foundcount + 1
            End If
        End If
        If totalclients = 1 Then
            cantbeother = tibiaclient
            cantbeotherBYTE = bc
        End If
    End If
  Loop

  GetProcessIdByAdrConnected = -1
End Function
Public Function UpdateCharListFromMemory(idConnection As Integer, maxr As Integer) As Long
  Dim tibiaclient As Long
  Dim curradr As Long
  Dim readSoFar As Long
  Dim partialRead As Long
  Dim strName As String
  Dim strServerName As String
  Dim thestart As Long
  Dim b As Byte
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim strIPandPort As String
  Dim currRead As Long
  Dim maxread As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  If TibiaVersionLong >= 971 Then
    UpdateCharListFromMemory = UpdateCharListFromMemory2(idConnection, maxr)
    Exit Function
  End If
  currRead = 0
  maxread = CLng(maxr) + 1
  tibiaclient = ProcessID(idConnection)
  ResetCharList2 idConnection
  thestart = Memory_ReadLong(adrCharListPtr, tibiaclient)
  curradr = thestart
  readSoFar = 0
continueIt:
  strName = ""
  partialRead = 0
  Do
    b = Memory_ReadByte(curradr, tibiaclient, True)
    If b = &H0 Then
      Exit Do
    Else
      strName = strName & Chr(b)
      readSoFar = readSoFar + 1
      partialRead = partialRead + 1
      curradr = curradr + 1
    End If
    If readSoFar > 10000 Then
      UpdateCharListFromMemory = -1
      Exit Function
    End If
    If partialRead = MAXCHARACTERLEN Then
      Exit Do
    End If
  Loop
  curradr = curradr + MAXCHARACTERLEN - partialRead
  strServerName = ""
  partialRead = 0
  Do
    b = Memory_ReadByte(curradr, tibiaclient, True)
    If b = &H0 Then
      Exit Do
    Else
      strServerName = strServerName & Chr(b)
      readSoFar = readSoFar + 1
      partialRead = partialRead + 1
      curradr = curradr + 1
    End If
    If readSoFar > 10000 Then
      UpdateCharListFromMemory = -1
      Exit Function
    End If
    If partialRead = MAXCHARACTERLEN Then
      Exit Do
    End If
  Loop
  If strServerName = "" Or (currRead >= maxread) Then
    If currRead >= maxread Then
      UpdateCharListFromMemory = 0
    Else
      LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory : could not read all (" & CStr(currRead) & "/" & CStr(maxread) & ")"
      UpdateCharListFromMemory = -1
    End If
    Exit Function
  Else
    Debug.Print strName & " : " & strServerName
    curradr = curradr + 54 - partialRead
    strIPandPort = GetIPandPortfromServerName(strServerName)
    If strIPandPort = "" Then
      LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory : can't get IP of server '" & strServerName & "'"
      UpdateCharListFromMemory = -1
      Exit Function
    End If
    servIP1 = CByte(CLng(Mid$(strIPandPort, 1, 3)))
    servIP2 = CByte(CLng(Mid$(strIPandPort, 5, 3)))
    servIP3 = CByte(CLng(Mid$(strIPandPort, 9, 3)))
    servIP4 = CByte(CLng(Mid$(strIPandPort, 13, 3)))
    servPort = CLng(Right$(strIPandPort, Len(strIPandPort) - 16))
    AddCharServer2 idConnection, strName, strServerName, servIP1, servIP2, servIP3, servIP4, servPort
    currRead = currRead + 1
  End If
  b = Memory_ReadByte(curradr, tibiaclient, True)
  If b = &H0 Then
    UpdateCharListFromMemory = 0
    Exit Function
  Else
    GoTo continueIt
  End If
  Exit Function
goterr:
  LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory : " & Err.Description
  UpdateCharListFromMemory = -1
End Function


Public Function UpdateCharListFromMemory2(idConnection As Integer, maxr As Integer) As Long
  Dim tibiaclient As Long
  Dim curradr As Long
  Dim readSoFar As Long
  Dim partialRead As Long
  Dim strName As String
  Dim strServerName As String
  Dim thestart As Long
  Dim b As Byte
  Dim servIP1 As Byte
  Dim servIP2 As Byte
  Dim servIP3 As Byte
  Dim servIP4 As Byte
  Dim servPort As Long
  Dim strIPandPort As String
  Dim currRead As Long
  Dim maxread As Long
  Dim MAXCHARACTERLEN2 As Long
  Dim charCount As Long
  Dim namesize As Long
  Dim nametype As Byte
  Dim i As Long
  Dim remoteAddress As Long
  Dim badd(3) As Byte
  
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  'Debug.Print "You should not use this function since Tibia 9.71"
  'UpdateCharListFromMemory2 = -1
  'Exit Function
  
  ' the new list does not display clear names but it should be enough for our small purpose...
  MAXCHARACTERLEN2 = 28
  currRead = 0
  maxread = CLng(maxr) + 1
  tibiaclient = ProcessID(idConnection)
  ResetCharList2 idConnection
  thestart = Memory_ReadLong(adrCharListPtr, tibiaclient)
  curradr = thestart
  readSoFar = 0
  charCount = 0
continueIt:
  strName = ""
  partialRead = 0
  curradr = curradr + 4 ' skip 4 strange bytes

  namesize = Memory_ReadByte(curradr + 16, tibiaclient, True)
  nametype = Memory_ReadByte(curradr + 20, tibiaclient, True)
  If nametype = &HF Then
    Do
      b = Memory_ReadByte(curradr, tibiaclient, True)
      If b = &H0 Then
        Exit Do
      Else
        strName = strName & Chr(b)
        readSoFar = readSoFar + 1
        partialRead = partialRead + 1
        curradr = curradr + 1
      End If
      If readSoFar > 10000 Then
        UpdateCharListFromMemory2 = -1
        Exit Function
      End If
      If partialRead = MAXCHARACTERLEN2 Then
        Exit Do
      End If
    Loop
  Else
'    badd(0) = Memory_ReadByte(curradr, tibiaclient, True)
'    badd(1) = Memory_ReadByte(curradr + 1, tibiaclient, True)
'    badd(2) = Memory_ReadByte(curradr + 2, tibiaclient, True)
'    badd(3) = Memory_ReadByte(curradr + 3, tibiaclient, True)
    remoteAddress = Memory_ReadLong(curradr, tibiaclient, True)
'    Debug.Print "character name stored in remote address = " & remoteAddress
'    Debug.Print GoodHex(badd(0)) & " " & GoodHex(badd(1)) & " " & GoodHex(badd(2)) & " " & GoodHex(badd(3))
    partialRead = 0
    strName = readMemoryString(tibiaclient, remoteAddress, CLng(namesize), True)
  End If
  charCount = charCount + 1
  'strname = "#" & CStr(charCount)
 ' Debug.Print "Got name type " & GoodHex(nametype) & " :" & strName
  curradr = curradr + MAXCHARACTERLEN2 - partialRead
  strServerName = ""
  partialRead = 0
  Do
    b = Memory_ReadByte(curradr, tibiaclient, True)
    If b = &H0 Then
      Exit Do
    Else
      strServerName = strServerName & Chr(b)
      readSoFar = readSoFar + 1
      partialRead = partialRead + 1
      curradr = curradr + 1
    End If
    If readSoFar > 10000 Then
      UpdateCharListFromMemory2 = -1
      Exit Function
    End If
    If partialRead = MAXCHARACTERLEN2 Then
      Exit Do
    End If
  Loop
  
  'Debug.Print "Got server:" & strServerName
  If strServerName = "" Or (currRead >= maxread) Then
    If currRead >= maxread Then
      UpdateCharListFromMemory2 = 0
    Else
      LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory2 : could not read all (" & CStr(currRead) & "/" & CStr(maxread) & ")"
      UpdateCharListFromMemory2 = -1
    End If
    Exit Function
  Else
    'Debug.Print "size=" & CLng(namesize) & " type= " & GoodHex(nametype); " : " & strname & " : " & strServerName
    curradr = curradr + 40 - partialRead
    strIPandPort = GetIPandPortfromServerName(strServerName)
    If strIPandPort = "" Then
      LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory2 : can't get IP of server '" & strServerName & "'"
      UpdateCharListFromMemory2 = -1
      Exit Function
    End If
    servIP1 = CByte(CLng(Mid$(strIPandPort, 1, 3)))
    servIP2 = CByte(CLng(Mid$(strIPandPort, 5, 3)))
    servIP3 = CByte(CLng(Mid$(strIPandPort, 9, 3)))
    servIP4 = CByte(CLng(Mid$(strIPandPort, 13, 3)))
    servPort = CLng(Right$(strIPandPort, Len(strIPandPort) - 16))
    AddCharServer2 idConnection, strName, strServerName, servIP1, servIP2, servIP3, servIP4, servPort
    currRead = currRead + 1
  End If
  'b = Memory_ReadByte(curradr + 4, tibiaclient, True)
  If currRead >= maxread Then
    UpdateCharListFromMemory2 = 0
    Exit Function
  Else
    GoTo continueIt
  End If
  Exit Function
goterr:
  LogOnFile "errors.txt", "Got error at UpdateCharListFromMemory2 : " & Err.Description
  UpdateCharListFromMemory2 = -1
End Function





Public Function GetProcessIDfromCharList2(ByVal idConnection As Long) As Long
    '...
    ' PENDIENTE DE PROGRAMAR
   Dim strName As String
   Dim readSoFar As Long
   Dim partialRead As Long
   Dim curradr As Long
   Dim thestart As Long
   Dim strName2 As String
   Dim tibiaclient As Long
   Dim b As Byte
   If idConnection = 0 Then
    GetProcessIDfromCharList2 = -1
    Exit Function
   End If
   Do

    tibiaclient = FindWindowEx(0, tibiaclient, tibiaclassname, vbNullString)
    If tibiaclient = 0 Then
      Exit Do
    Else
    
    
        thestart = Memory_ReadLong(adrCharListPtr, tibiaclient)
        curradr = thestart
        strName = ""
        partialRead = 0
        Do
            b = Memory_ReadByte(curradr, tibiaclient, True)
            If b = &H0 Then
                Exit Do
            Else
                strName = strName & Chr(b)
                readSoFar = readSoFar + 1
                partialRead = partialRead + 1
                curradr = curradr + 1
            End If
            If readSoFar > 10000 Then
                Exit Do
            Else
                If partialRead = MAXCHARACTERLEN Then
                    Exit Do
                End If
            End If
        Loop
        strName2 = CharacterList2(idConnection).item(0).CharacterName
        If strName = strName2 Then
            GetProcessIDfromCharList2 = tibiaclient
            Exit Function
        End If
    End If
  Loop
    
    
    GetProcessIDfromCharList2 = -1
End Function

Public Function fixThreeDigits(n As Byte) As String
  Dim res As String
  res = ""
  If (n < 100) Then
    res = "0"
  End If
  If (n < 10) Then
    res = res & "0"
  End If
  res = res & CStr(CInt(n))
  fixThreeDigits = res
End Function



