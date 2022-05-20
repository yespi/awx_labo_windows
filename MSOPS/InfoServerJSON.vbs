on error resume next
Dim OFSO
set OFSO = CreateObject("Scripting.FileSystemObject")
Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set oNetWork = WScript.CreateObject("WScript.Network")
Dim strPath
strPath = OFSO.GetAbsolutePathName(".")

'##########################################################################
'
' Funciones
'
'##########################################################################


Class JSONStringEncoder

    Private m_RegExp
    
    Sub Class_Initialize()
        Set m_RegExp = Nothing
    End Sub
    
    Function Encode(ByVal Str)

        Dim Parts(): ReDim Parts(3)
        Dim NextPartIndex: NextPartIndex = 0
        Dim AnchorIndex: AnchorIndex = 1
        Dim CharCode, Escaped
        Dim Match, MatchIndex
        Dim RegExp: Set RegExp = m_RegExp

        If RegExp Is Nothing Then
            Set RegExp = New RegExp
            ' See https://github.com/douglascrockford/JSON-js/blob/43d7836c8ec9b31a02a31ae0c400bdae04d3650d/json2.js#L196
            RegExp.Pattern = "[\\\""\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]"
            RegExp.Global = True
            Set m_RegExp = RegExp
        End If
        For Each Match In RegExp.Execute(Str)
            MatchIndex = Match.FirstIndex + 1
            If NextPartIndex > UBound(Parts) Then ReDim Preserve Parts(UBound(Parts) * 2)
            Parts(NextPartIndex) = Mid(Str, AnchorIndex, MatchIndex - AnchorIndex): NextPartIndex = NextPartIndex + 1
            CharCode = AscW(Mid(Str, MatchIndex, 1))
            Select Case CharCode
                Case 34  : Escaped = "\"""
                Case 10  : Escaped = "\n"
                Case 13  : Escaped = "\r"
                Case 92  : Escaped = "\\"
                Case 8   : Escaped = "\b"
                Case Else: Escaped = "\u" & Right("0000" & Hex(CharCode), 4)
            End Select
            If NextPartIndex > UBound(Parts) Then ReDim Preserve Parts(UBound(Parts) * 2)
            Parts(NextPartIndex) = Escaped: NextPartIndex = NextPartIndex + 1
            AnchorIndex = MatchIndex + 1
        Next
        If AnchorIndex = 1 Then Encode = """" & Str & """": Exit Function
        If NextPartIndex > UBound(Parts) Then ReDim Preserve Parts(UBound(Parts) * 2)
        Parts(NextPartIndex) = Mid(Str, AnchorIndex): NextPartIndex = NextPartIndex + 1
        ReDim Preserve Parts(NextPartIndex - 1)
	Encode = """" & Join(Parts, "") & """"
    End Function

End Class

Dim TheJSONStringEncoder: Set TheJSONStringEncoder = New JSONStringEncoder

Function EncodeJSONString(ByVal Str)
    EncodeJSONString = TheJSONStringEncoder.Encode(Str) 
End Function

Function EncodeJSONMember(ByVal Key, Value)
    EncodeJSONMember = EncodeJSONString(Key) & ":" & JSONStringify(Value)
End Function

Public Function JSONStringify(Thing) 

    Dim Key, Item, Index, NextIndex, Arr()
    Dim VarKind: VarKind = VarType(Thing)
    Select Case VarKind
        Case vbNull, vbEmpty: JSONStringify = "null"
        Case vbDate: JSONStringify = EncodeJSONString(FormatISODateTime(Thing))
        Case vbString: JSONStringify = EncodeJSONString(Thing)
        Case vbBoolean: If Thing Then JSONStringify = "true" Else JSONStringify = "false"
        Case vbObject
            If Thing Is Nothing Then
                JSONStringify = "null"
            Else
                If TypeName(Thing) = "Dictionary" Then
                    If Thing.Count = 0 Then JSONStringify = "{}": Exit Function
                    ReDim Arr(Thing.Count - 1)
                    Index = 0
                    For Each Key In Thing.Keys
                        Arr(Index) = EncodeJSONMember(Key, Thing(Key))
                        Index = Index + 1
                    Next
                    JSONStringify = parsea("{" & Join(Arr, ",") & "}")
                Else
                    ReDim Arr(3)
                    NextIndex = 0
                    For Each Item In Thing
                        If NextIndex > UBound(Arr) Then ReDim Preserve Arr(UBound(Arr) * 2)
                        Arr(NextIndex) = JSONStringify(Item)
                        NextIndex = NextIndex + 1
                    Next
                    ReDim Preserve Arr(NextIndex - 1)
                    JSONStringify = parsea("[" & Join(Arr, ",") & "]")
                End If
            End If
        Case Else
            If vbArray = (VarKind And vbArray) Then
                For Index = LBound(Thing) To UBound(Thing)
                    If Len(JSONStringify) > 0 Then JSONStringify = JSONStringify & ","
                    JSONStringify = parsea(JSONStringify & JSONStringify(Thing(Index)))
                Next
                JSONStringify = parsea("[" & JSONStringify & "]")
            ElseIf IsNumeric(Thing) Then
                JSONStringify = parsea(CStr(Thing))
            Else
                JSONStringify = parsea(EncodeJSONString(CStr(Thing)))
            End If
    End Select

End Function

Function parsea (str)
	str = replace(str,"á","a")
	str = replace(str,"é","e")
	str = replace(str,"í","i")
	str = replace(str,"ó","o")
	str = replace(str,"ú","u")
	str = replace(str,"Á","A")
	str = replace(str,"É","E")
	str = replace(str,"Í","I")
	str = replace(str,"Ó","O")
	str = replace(str,"Ú","U")
	str = replace(str,"à","a")
	str = replace(str,"è","e")
	str = replace(str,"ì","i")
	str = replace(str,"ò","o")
	str = replace(str,"ù","u")
	str = replace(str,"â","a")
	str = replace(str,"ê","e")
	str = replace(str,"î","i")
	str = replace(str,"ô","o")
	str = replace(str,"û","u")
	str = replace(str,"Â","A")
	str = replace(str,"Ê","E")
	str = replace(str,"Î","I")
	str = replace(str,"Ô","O")
	str = replace(str,"Û","U")
	str = replace(str,"ä","a")
	str = replace(str,"ë","e")
	str = replace(str,"ï","i")
	str = replace(str,"ö","o")
	str = replace(str,"ü","u")
	str = replace(str,"Ä","A")
	str = replace(str,"Ë","E")
	str = replace(str,"Ï","I")
	str = replace(str,"Ö","O")
	str = replace(str,"Ü","U")
	str = replace(str,"ç","c")
	str = replace(str,"Ç","C")
	str = replace(str,"®","")
	str = replace(str,"ã","a")
	str = replace(str,"ß","")
	str = replace(str,"·","-")
	str = replace(str,"Ð","D")
	str = replace(str,"€","E")
	str = replace(str,"ñ","ny")
	str = replace(str,"Ñ","NY")
	str = replace(str,"º","o")
	str = replace(str,"ª","a")
	str = replace(str,"Ý","Y")
	str = replace(str,"¡","i")
	str = replace(str,"¾","")
	str = replace(str,"¦","-")
	str = replace(str,"'","\'")
	'
	Parsea=str
End Function
'========================================================================== 
' PutLog - Guarda una línea en un fichero de texto. Si no existe, lo crea
'========================================================================== 
sub putlog (strFich,strTexto)
	Const ForReading=1
	Const ForWriting=2
	Const ForAppending=8

	If not oFSO.FileExists(strFich) Then
		Set oFile = oFSO.CreateTextFile(strFich, True)
		oFile.WriteLine(strTexto) 
		oFile.Close 
	else
		Set oFile = oFSO.OpenTextFile(strFich, ForAppending)
		oFile.WriteLine(strTexto) 
		oFile.Close
	end if
	wscript.echo strFich
	wscript.echo strTexto
end sub


Function WMIDateToString(varWMIDate)
 ' Get date string in mm/dd/yyyy hh:nn:ss format.
 ' WMIDateToString = Mid(varWMIDate, 5, 2) & "/" & Mid(varWMIDate, 7, 2) & "/" & Left(varWMIDate, 4) & " " & _
 '      Mid(varWMIDate, 9, 2) & ":" & Mid(varWMIDate, 11, 2) & ":" & Mid(varWMIDate, 13, 2)
 ' Get date string in yyyymmddhhnnss format.
   WMIDateToString = Left(varWMIDate, 4) & Mid(varWMIDate, 5, 2) & Mid(varWMIDate, 7, 2) & _
	Mid(varWMIDate, 9, 2) & Mid(varWMIDate, 11, 2) & Mid(varWMIDate, 13, 2)
End Function

Function DateToString(vNow)
  '\\ Create Timestamp
' vNow = Now()
 vMthStr = CStr(Month(vNow))
 vDayStr = CStr(Day(vNow))
 vHourStr = CStr(Hour(vNow))
 vMinuteStr = CStr(Minute(vNow))
 vSecondStr = CStr(Second(vNow))

 ' añadimos los 0 que falten
 if len(vHourStr)=1 then vHourStr= "0" & vHourStr
 If Len(vMthStr) = 1 Then vMthStr = "0" & vMthStr
 If Len(vDayStr) = 1 Then vDayStr = "0" & vDayStr
 If Len(vMinuteStr) = 1 Then vMinuteStr = "0" & vMinuteStr
 If Len(vSecondStr) = 1 Then vSecondStr = "0" & vSecondStr

 ' Get date string in yyyymmddhhnnss format.
   DateToString= Year(vNow) & vMthStr & vDayStr & vHourStr & vMinuteStr & vSecondStr
End Function

'Ejecutamos comandos MS-DOS
Function Ejecuta (strcommand)
	Dim osh,oExec
	Set osh=CreateObject("wscript.shell")

	strOs=check_os()
	if instr(strOs,"w2003") then
		strcommand=replace(ucase(strcommand),"PROGRAM FILES","PROGRA~1")
	end if
	strErr=""
	Set oExec = osh.Exec( strCommand)
	Do While oExec.Status = 0
	  WScript.Sleep 100
	Loop
	strErr=oExec.StdErr.ReadAll
	if strErr<>"" then
		strResult="ERROR: " & strErr
	else
		strResult=oExec.StdOut.ReadAll
	end if
	Ejecuta=replace(strResult,vbCrLf,"")
end function

Function Check_Os
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

	vOs="No_Detectado"
	For Each objOperatingSystem in colOperatingSystems
		if instr(objOperatingSystem.Caption,"2022") then vOs="w2022"
		if instr(objOperatingSystem.Caption,"2019") then vOs="w2019"
		if instr(objOperatingSystem.Caption,"2016") then vOs="w2016"
		if instr(objOperatingSystem.Caption,"2012") then vOs="w2012"
		if instr(objOperatingSystem.Caption,"2008") then vOs="w2008"
		if instr(objOperatingSystem.Caption,"2003") then vOs="w2003"
	Next
	Check_Os=vOs
end function

function LastLogon(usuario)

	on error resume next
	llogon="Never"
	Set objUser = GetObject("WinNT://./" & usuario)
	lLogon = objUser.LastLogin
    If Err <> 0 Then
        lLogon="Never"
    End If
    On Error GoTo 0
	LastLogon=replace(llogon,"/","-")
end function
'***********************

'******* Convierte el fichero a UTF8, para que se pueda importar bien
function utf8(strfileIn,strfileOut)

	Set stream = CreateObject("ADODB.Stream")
	stream.Open
	stream.Type = 2 'text
	stream.Charset = "utf-8"
	stream.LoadFromFile strFileIn
	text = stream.ReadText
	stream.Close

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(strFileOut, 2, True, True)
	f.Write text
	f.Close
	utf8=true
end function

function fqdn()
	on error resume next
	strDomain=""
	Set sysInfo    = CreateObject("ADSystemInfo")
	Set wshNetwork = CreateObject("WScript.Network")
	strDomain=sysInfo.DomainDNSName
	if strdomain="" then
		strFQDN=wshNetwork.ComputerName
	else
		strFQDN=wshNetwork.ComputerName & "." & sysInfo.DomainDNSName
	end if
	fqdn=strFQDN
end function


'***********************

'##########################################################################
'
' MAIN - Creamos el fichero JSON resultante
'
'##########################################################################

'***********************
'\\ Create strHostname

Set objWshNetwork = WScript.CreateObject("WScript.Network")
If IsObject( objWshNetwork ) Then
	strHostname=LCase( objWshNetwork.ComputerName )
Else
	strHostname=LCase( objWshShell.ExpandEnvironmentStrings("%COMPUTERNAME%") )
End If

'\\ Create Timestamp

strTimestamp=DateToString(Now())

'\\ Create System
  '****** BIOS
  Set colBIOS = oWMI.ExecQuery ("Select * from Win32_BIOS")
  For each objBIOS in colBIOS
 	BIOS = Trim(objBIOS.Version) &" "& objBIOS.SMBIOSBIOSVersion
	BIOSSerial = objBIOS.SerialNumber
  Next

  '****** CPU
  Set colItems = oWMI.ExecQuery("SELECT * FROM Win32_Processor", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
  For Each objItem In colItems
	  strCPUAddressWidth=objItem.AddressWidth
	  strCPUArchitecture=objItem.Architecture
	  strCPUCaption=objItem.Caption
	  strCPUStatus=objItem.CpuStatus
	  strCPUDeviceID=objItem.DeviceID
	  strCPUFamily=objItem.Family
	  strCPUManufacturer=objItem.Manufacturer
	  strCPUMaxClockSpeed=objItem.MaxClockSpeed
	  strCPUName=objItem.Name
	  strCPUCores = objItem.NumberOfCores
	  strCPUNumberOfLogicalProcessors=objItem.NumberOfLogicalProcessors
  Next

  '***** Sistema operativo y Memoria
  MB = 1024 
  Set colItems = oWMI.ExecQuery ("Select * from Win32_OperatingSystem")
  For Each objItem In colItems
	strOSBootDevice=objItem.BootDevice
	strOSBuild=objItem.BuildNumber
	strOSCaption=trim(objItem.Caption)
	strOSCodeSet=objItem.CodeSet
	strOSCountryCode=objItem.CountryCode
	strOSCurrentTimeZone=objItem.CurrentTimeZone
	strOSInstallDate=WMIDateToString(objItem.InstallDate)
	strOSLastBootTime=WMIDateToString(objItem.LastBootUpTime)
	strOSLocalDateTime=WMIDateToString(objItem.LocalDateTime)
	intRAMFree=int(objItem.FreePhysicalMemory/MB)
	intRAMPaging=int(objItem.FreeSpaceInPagingFiles/MB)
	intRAMVFree = int(objItem.FreeVirtualMemory/MB)
        intRAMTotal = int(objItem.TotalVisibleMemorySize/MB)
	strOSManufacturer=objItem.Manufacturer
	strOSName=objItem.Name
	strOSNumberOfLicensedUsers=objItem.NumberOfLicensedUsers
	strOSNumberOfProcesses=objItem.NumberOfProcesses
	strOSNumberOfUsers=objItem.NumberOfUsers
	strOSSerialNumber=objItem.SerialNumber
	strOSSystemDevice=objItem.SystemDevice
	strOSSystemDirectory=objItem.SystemDirectory
	strOSVersion=objItem.Version
	strOSWindowsDirectory=objItem.WindowsDirectory
  next

  Set colItems = oWMI.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)
  For Each objItem in colItems
    intRAMCommitLimit=int(objItem.CommitLimit / MB /MB /MB)
    intRAMCommit=int(objItem.CommittedBytes / MB /MB /MB)
  Next

strFQDN=fqdn()

Set oStaD = CreateObject("Scripting.Dictionary")
oStaD.Add "Hostname",strHostname
oStaD.Add "CI",strHostname
oStaD.Add "FQDN",strFQDN
oStaD.Add "Timestamp",strTimestamp
oStaD.Add "distribution",""
oStaD.Add "OS",CreateObject("Scripting.Dictionary")
oStaD.Item("OS").Add "BootDevice", strOSBootDevice
oStaD.Item("OS").Add "Build", strOSBuild
oStaD.Item("OS").Add "Caption", strOSCaption
oStaD.Item("OS").Add "CodeSet", strOSCodeSet
oStaD.Item("OS").Add "CountryCode", strOSCountryCode
oStaD.Item("OS").Add "CurrentTimeZone", strOSCurrentTimeZone
oStaD.Item("OS").Add "InstallDate", strOSInstallDate
oStaD.Item("OS").Add "LastBootTime", strOSLastBootTime
oStaD.Item("OS").Add "LocalDateTime", strOSLocalDateTime
oStaD.Item("OS").Add "Manufacturer",strOSManufacturer
oStaD.Item("OS").Add "Name",strOSName
oStaD.Item("OS").Add "NumberofLicensedUsers",strOSNumberOfLicensedUsers
oStaD.Item("OS").Add "NumberofProcesses",strOSNumberOfProcesses
oStaD.Item("OS").Add "NumberofUsers",strOSNumberOfUsers
oStaD.Item("OS").Add "SerialNumber",strOSSerialNumber
oStaD.Item("OS").Add "SystemDevice",strOSSystemDevice
oStaD.Item("OS").Add "SystemDirectory",strOSSystemDirectory
oStaD.Item("OS").Add "Version",strOSVersion
oStaD.Item("OS").Add "WindowsDirectory",strOSWindowsDirectory


oStaD.Add "CPU",CreateObject("Scripting.Dictionary")
oStaD.Item("CPU").Add "addresswidth", strCPUAddressWidth
oStaD.Item("CPU").Add "architecture", strCPUArchitecture
oStaD.Item("CPU").Add "caption", strCPUcaption
oStaD.Item("CPU").Add "status", strCPUstatus
oStaD.Item("CPU").Add "deviceID", strCPUdeviceID
oStaD.Item("CPU").Add "family", strCPUfamily
oStaD.Item("CPU").Add "manufacturer", strCPUManufacturer
oStaD.Item("CPU").Add "maxClockspeed", strCPUMaxClockSpeed
oStaD.Item("CPU").Add "name", strCPUName
oStaD.Item("CPU").Add "cores",strcpuCores
oStaD.Item("CPU").Add "vcpus",strCPUNumberOfLogicalProcessors
oStaD.Add "memory",CreateObject("Scripting.Dictionary")
oStaD.Item("memory").Add "total",intRAMTotal
oStaD.Item("memory").Add "free",intRAMFree
oStaD.Item("memory").Add "free virtual",intRAMVFree
oStaD.Item("memory").Add "Paging",intRAMPaging
oStaD.Item("memory").Add "Commit_GB",intRAMCommit
oStaD.Item("memory").Add "CommitLimit_GB",intRAMCommitLimit

'***********************
'\\ Networking

oStaD.Add "networks",CreateObject("Scripting.Dictionary")


dim ArrNetIF()
Set InterfaceName = oWMI.ExecQuery ("Select * From Win32_NetworkAdapter Where NetConnectionStatus = 2")
cont=0
For Each objItem in InterfaceName
    strAdptName=objItem.NetConnectionID
    oStaD.Item("networks").Add strAdptName,CreateObject("Scripting.Dictionary")
    ReDim Preserve ArrNetIF(cont + 1)
    ArrNetIF(cont)=strAdptName
    cont=cont+1
next

Set colAdapters = oWMI.ExecQuery ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
cont=1 
For Each oAdpt in colAdapters
	'Usaremos "n" para etiquetar adaptadores, Ips, etc
	n=cont:if cont=0 then n=""

	   strAdptName=ArrNetIF(cont-1)
	   strAdptDescription=oAdpt.Description
	   oStaD.Item("networks").Item(strAdptName).Add "description",oAdpt.Description

	   If Not IsNull(oAdpt.IPAddress) Then
		  For i = 0 To UBound(oAdpt.IPAddress)
		 ii=i:if i=0 then ii=""
		 oStaD.Item("networks").Item(strAdptName).Add "ip"&ii,oAdpt.IPAddress(i)
		  Next
	   End If
	 
	   If Not IsNull(oAdpt.IPSubnet) Then
		  For i = 0 To UBound(oAdpt.IPSubnet)
		 ii=i:if i=0 then ii=""
	'	wscript.echo strAdptName & "subnet"&ii & " - " & IPSubnet(i)
		 oStaD.Item("networks").Item(strAdptName).Add "subnet"&ii,oAdpt.IPSubnet(i)
		 
		  Next
	   End If
	 
	   If Not IsNull(oAdpt.DefaultIPGateway) Then
		  For i = 0 To UBound(oAdpt.DefaultIPGateway)
		 ii=i:if i=0 then ii=""
		 oStaD.Item("networks").Item(strAdptName).Add "DefaultGW"&ii,oAdpt.DefaultIPGateway(i)
		  Next
	   End If
	 
	   If Not IsNull(oAdpt.DNSServerSearchOrder) Then
		  For i = 0 To UBound(oAdpt.DNSServerSearchOrder)
		 ii=i:if i=0 then ii=""
		 oStaD.Item("networks").Item(strAdptName).Add "DNS"&ii,oAdpt.DNSServerSearchOrder(i)
		  Next
	   End If
	 
	   If Not IsNull(oAdpt.DNSDomainSuffixSearchOrder) Then
		  For i = 0 To UBound(oAdpt.DNSDomainSuffixSearchOrder)
		 ii=i:if i=0 then ii=""
		 oStaD.Item("networks").Item(strAdptName).Add "DNSSuffix"&ii,oAdpt.DNSDomainSuffixSearchOrder(i)
		  Next
	   End If
	 
	   oStaD.Item("networks").Item(strAdptName).Add "DHDCPEnabled",oAdpt.DHCPEnabled
	   oStaD.Item("networks").Item(strAdptName).Add "DHCPServer",oAdpt.DHCPServer
	 
	   If Not IsNull(oAdpt.DHCPLeaseObtained) Then
		  utcLeaseObtained = oAdpt.DHCPLeaseObtained
		  strLeaseObtained = WMIDateToString(utcLeaseObtained)
	   Else
		  strLeaseObtained = ""
	   End If
	   oStaD.Item("networks").Item(strAdptName).Add "DHCPlease",strLeaseObtained
	 
	   If Not IsNull(oAdpt.DHCPLeaseExpires) Then
		  utcLeaseExpires = oAdpt.DHCPLeaseExpires
		  strLeaseExpires = WMIDateToString(utcLeaseExpires)
	   Else
		  strLeaseExpires = ""
	   End If
	   oStaD.Item("networks").Item(strAdptName).Add "DHCPleaseExpires",strLeaseExpires
	   oStaD.Item("networks").Item(strAdptName).Add "PrimaryWins",oAdpt.WINSPrimaryServer
	   oStaD.Item("networks").Item(strAdptName).Add "SecondaryWins",oAdpt.WINSSecondaryServer
	   cont = cont + 1
Next

oStaD.Item("networks").Add "routes",CreateObject("Scripting.Dictionary")
cont=1
Set colItems = oWMI.ExecQuery("Select * from Win32_IP4RouteTable")
For Each objItem in colItems
    oStaD.Item("networks").Item("routes").Add "route_"&cont,CreateObject("Scripting.Dictionary")
    oStaD.Item("networks").Item("routes").Item("route_"&cont).Add "destination",objItem.Destination
    oStaD.Item("networks").Item("routes").Item("route_"&cont).Add "mask",objItem.Mask
    oStaD.Item("networks").Item("routes").Item("route_"&cont).Add "gateway",objItem.NextHop
    cont=cont+1
Next

'***************************************************************
'// Users

oStaD.Add "users",CreateObject("Scripting.Dictionary")
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery ("Select * from Win32_UserAccount Where LocalAccount = True")

dim aGroups()
netaccounts=0
For Each objItem in colItems
    redim aGroups(0)
	strItem=replace(objItem.name,".","")
	if netaccounts=0 then '''' Generamos el NET ACCOUNTS
		set objUser = GetObject("WinNT://./" & strItem & ",user")
		oStaD.Add "users_policy",CreateObject("Scripting.Dictionary")
		minPassAge=int(objUser.minPasswordAge/86400)
		maxPassAge=int(objUser.maxPasswordAge/86400)
		if ((minPassAge)<0) then 
			minPassAge=0
		end if
		if ((minPassAge)>100000) then 
			minPassAge=0
		end if
		if ((maxPassAge)<0) then 
			maxPassAge=0
		end if
		if ((maxPassAge)>100000) then 
			maxPassAge=0
		end if
		oStaD.Item("users_policy").Add "MinPasswordAge", minPassAge
		oStaD.Item("users_policy").Add "MaxPasswordAge", maxPassAge
		oStaD.Item("users_policy").Add "MinPasswordlength", objUser.MinPasswordlength
		oStaD.Item("users_policy").Add "PasswordHistoryLength", objUser.PasswordHistoryLength
		oStaD.Item("users_policy").Add "LockoutObservationInterval", objUser.LockoutObservationInterval / 60
		netaccounts=1
		set objUser=Nothing
	end if
	
    oStaD.Item("users").Add strItem,CreateObject("Scripting.Dictionary")
    oStaD.Item("users").Item(strItem).Add "description", objItem.Description
    oStaD.Item("users").Item(strItem).Add "disabled", objItem.disabled
    oStaD.Item("users").Item(strItem).Add "domain", objItem.Domain
    oStaD.Item("users").Item(strItem).Add "fullname", objItem.FullName
    oStaD.Item("users").Item(strItem).Add "local_account", objItem.LocalAccount
    oStaD.Item("users").Item(strItem).Add "lockout", objItem.lockout
    oStaD.Item("users").Item(strItem).Add "name", objItem.name
    oStaD.Item("users").Item(strItem).Add "password_changeable", objItem.PasswordChangeable
    oStaD.Item("users").Item(strItem).Add "password_expires", objItem.PasswordExpires
    oStaD.Item("users").Item(strItem).Add "password_required", objItem.PasswordRequired
    oStaD.Item("users").Item(strItem).Add "status", objItem.status
	on error goto 0
    ultlogon=LastLogon(objItem.name)
	oStaD.Item("users").Item(strItem).Add "LastLogon", ultlogon

    cont=0
    Set colGroups = GetObject("WinNT://" & strComputer & "")
    colGroups.Filter = Array("group")
    For Each objGroup In colGroups
    	For Each objUser in objGroup.Members
	        If objUser.name = objItem.Name Then
		    ReDim Preserve aGroups(cont)
		    aGroups(cont)=objGroup.Name
		    cont=cont+1
	        End If
	    Next
    Next
    if cont>0 then     oStaD.Item("users").Item(strItem).Add "groups", aGroups
Next

'*******************************************************************
'// Disk Drives

'//// Logical Drives
oStaD.Add "filesystems",CreateObject("Scripting.Dictionary")
oStaD.Item("filesystems").Add "logicaldrives",CreateObject("Scripting.Dictionary")

Const wbemFlagReturnImmediately = &h10 
Const wbemFlagForwardOnly = &h20 
Set colItems = oWMI.ExecQuery("SELECT * FROM Win32_LogicalDisk where description='Local Fixed Disk' OR 	description='Disco fijo local'", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly) 
	
For Each objItem In colItems 
	tfreespc=round((((objItem.FreeSpace/1024)/1024/1024)),2) & " GB"
	tsize=round((((objItem.Size/1024)/1024/1024)),2) & " GB"
	PerCent = objItem.FreeSpace/objItem.Size 
	tfreepct=round(PerCent * 100,2) & " %"

	oStaD.Item("filesystems").Item("logicaldrives").Add objItem.Caption,CreateObject("Scripting.Dictionary")
	oStaD.Item("filesystems").Item("logicaldrives").Item(objItem.Caption).Add "description",objItem.Description
	oStaD.Item("filesystems").Item("logicaldrives").Item(objItem.Caption).Add "freespace",tfreespc
	oStaD.Item("filesystems").Item("logicaldrives").Item(objItem.Caption).Add "size",tsize
	oStaD.Item("filesystems").Item("logicaldrives").Item(objItem.Caption).Add "freepercent",tfreepct
Next 

'//// Physical Drives
oStaD.Item("filesystems").Add "physicaldrives",CreateObject("Scripting.Dictionary")
Set colDiskDrives = oWMI.ExecQuery ("Select * from Win32_DiskDrive")
For each objItem in colDiskDrives    
	if InStr(1,objItem.Model, "HITACHI", 1)  = 0 AND InStr(1,objItem.Model, "NETAPP", 1)  = 0 AND _
  	   InStr(1,objItem.Model, "CLARION", 1)  = 0 then
		strItem=replace(objItem.name,".","")
		strItem=replace(strItem,"\","")
		oStaD.Item("filesystems").Item("physicaldrives").Add strItem,CreateObject("Scripting.Dictionary")
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "caption",objItem.Caption
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "model",objItem.Model
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "deviceid",objItem.DeviceID
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "interfacetype",objItem.InterfaceType
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "mediatype",objItem.MediaType
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "partitions",objItem.partitions
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "pnpdeviceid",objItem.PNPDeviceID
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "scsibus",objItem.SCSIBus
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "scsilogicalunit",objItem.SCSILogicalUnit
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "scsiport",objItem.SCSIPort
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "scsitargetid",objItem.SCSITargetId
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "signature",objItem.Signature
		tsize=round((((objItem.Size/1024)/1024/1024)),2) & " GB"
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "size",objItem.Size
		oStaD.Item("filesystems").Item("physicaldrives").Item(strItem).Add "status",objItem.status
	End If
Next


on error resume next
strPathDatos="c:\ts_data\INFO_SERVER\"
'***********************************************************
'/// Borramos los ficheros JSON antiguos
Const DeleteReadOnly = TRUE
oFSO.DeleteFile(strPathDatos & "*.json"), DeleteReadOnly
on error goto 0

'***********************************************************
'/// Creamos el fichero JSON
If Not oFSO.FolderExists(strPathDatos) Then
    BuildFullPath oFSO.GetParentFolderName(strPathDatos)
    oFSO.CreateFolder strPathDatos
End If
strFile=strPathDatos & strHostname & "@" & strHostname & "@WINDOWS@OS@"& strTimestamp &".json"
putlog strFile,JSONStringify(oStaD)

utf8 strFile,strFile
