Dim objFSO, newDIR, fso, server, FreeSpace, warnspace, dateineu, speicherstorage, strComputer, space

'Logfiles
Dim strSafeDate, strSafeTime, strDateTime, strLogFileName

'Datum etc
Dim ausgabe
Dim datum
Dim zeit

'allgemein infos
Set wshnet = CreateObject("WScript.Network")
strComputer = wshnet.Computername

Set objShell = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.Filesystemobject")
Set colDrives = fso.Drives

'Auslesen der Windows-Hardwareinformationen
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk",,48)

For Each objItem in colItems
	if objItem.Caption = "C:" then festplatteC= objItem.Caption & " " & Round(objItem.FreeSpace /1024 /1024 /1024, 2) & " / " & Round(objItem.Size /1024 /1024 /1024,2) & " GByte"
	if objItem.Caption = "D:" then festplatteD= objItem.Caption & " " & Round(objItem.FreeSpace /1024 /1024 /1024, 2) & " / " & Round(objItem.Size /1024 /1024 /1024,2) & " GByte"
	if objItem.Caption = "W:" then festplatteW= objItem.Caption & " " & Round(objItem.FreeSpace /1024 /1024 /1024, 2) & " / " & Round(objItem.Size /1024 /1024 /1024,2) & " GByte"
	if objItem.Caption = "Z:" then festplatteZ= objItem.Caption & " " & Round(objItem.FreeSpace /1024 /1024 /1024, 2) & " / " & Round(objItem.Size /1024 /1024 /1024,2) & " GByte"
	if objItem.Caption = "H:" then festplatteH= objItem.Caption & " " & Round(objItem.FreeSpace /1024 /1024 /1024, 2) & " / " & Round(objItem.Size /1024 /1024 /1024,2) & " GByte"
	if objItem.Caption = "X:" then festplatteX= objItem.Caption & " " & Round(objItem.FreeSpace /1024 /1024 /1024, 2) & " / " & Round(objItem.Size /1024 /1024 /1024,2) & " GByte"
	if objItem.Caption = "Y:" then festplatteY= objItem.Caption & " " & Round(objItem.FreeSpace /1024 /1024 /1024, 2) & " / " & Round(objItem.Size /1024 /1024 /1024,2) & " GByte"
	if objItem.Caption = "E:" then festplatteE= objItem.Caption & " " & Round(objItem.FreeSpace /1024 /1024 /1024, 2) & " / " & Round(objItem.Size /1024 /1024 /1024,2) & " GByte"

Next

' Gesammten Speicher ermitteln für warnspace
space = Round(fso.GetDrive("C:\").AvailableSpace/1024/1024/1024,2)

warnspace = 10

' Datum und timestamp für msgBox und txt
set ausgabe = WScript.CreateObject("WScript.Shell")
datum = Date
zeit = Time
timestamp = datum & " "  & zeit

strSafeDate = Right("0" & DatePart("d",Date), 2) & "-" & Right("0" & DatePart("m",Date), 2) & "-" & DatePart("yyyy",Date)
strSafeTime = Right("0" & Hour(Now), 2)  & "-" &  Right("0" & Minute(Now), 2) & "-" & Right("0" & Second(Now), 2)
strDateTime = "Speicherplatz_[" & strSafeDate & "]_[" & strSafeTime & "]"

' Abspeichern + Ordner Directory
ziel = objShell.ExpandEnvironmentStrings("%userprofile%") & "\Desktop\StorageSystem\speicher"

' Textbox und Warnung + Logfile automatisch erstellen
IF space >= warnspace Then
	
	Meldung = "Es ist GENUG Speicher frei:" & VbCr & VbCr

	'For Each objDrive in colDrives
	'	IF fso.DriveExists(objDrive.DriveLetter &":\") then Add "Festplatte " & festplatte & objDrive.DriveLetter End if
	'Next
	
	IF fso.DriveExists("C:\") then Add "Festplatte " & festplatteC End if
	IF fso.DriveExists("D:\") then Add "Festplatte " & festplatteD End if
	IF fso.DriveExists("W:\") then Add "Festplatte " & festplatteW End if
	IF fso.DriveExists("Z:\") then Add "Festplatte " & festplatteZ End if
	IF fso.DriveExists("H:\") then Add "Festplatte " & festplatteH End if
	IF fso.DriveExists("X:\") then Add "Festplatte " & festplatteX End if
	IF fso.DriveExists("Y:\") then Add "Festplatte " & festplatteY End if
	IF fso.DriveExists("E:\") then Add "Festplatte " & festplatteE End if
	
	Add " "
	Add "Timestamp: " & timestamp

	' Ordner erstellen
	folderaim = objShell.ExpandEnvironmentStrings("%userprofile%") & "\Desktop\StorageSystem\treiber"
	folderfile = objShell.ExpandEnvironmentStrings("%userprofile%") & "\Desktop\StorageSystem\treiber\hp.txt"
	
	IF Not fso.FolderExists(folderaim) Then 
		fso.CreateFolder folderaim
	End IF
	
	' Datei hineinkopieren
	'IF Not fso.FileExists(folderfile) Then
		'fso.CopyFile "Z:\2019_2020\PFE\IMD5\hp.txt", objShell.ExpandEnvironmentStrings("%userprofile%") & "\Desktop\StorageSystem\treiber\", True
	'End IF
	
	GetInfo "localhost"

Else
	
	Meldung = "Es ist ZU WENIG Speicher frei:" & VbCr & VbCr
	
	IF fso.DriveExists("C:\") then Add "Festplatte " & festplatteC End if
	IF fso.DriveExists("D:\") then Add "Festplatte " & festplatteD End if
	IF fso.DriveExists("W:\") then Add "Festplatte " & festplatteW End if
	IF fso.DriveExists("Z:\") then Add "Festplatte " & festplatteZ End if
	IF fso.DriveExists("H:\") then Add "Festplatte " & festplatteH End if
	IF fso.DriveExists("X:\") then Add "Festplatte " & festplatteX End if
	IF fso.DriveExists("Y:\") then Add "Festplatte " & festplatteY End if
	IF fso.DriveExists("E:\") then Add "Festplatte " & festplatteE End if

	Add " "
	Add "Es wurde automatisch eine Logfile erstellt!"
	Add " "
	Add "Timestamp: " & timestamp
	
	'msgBox space & " GB frei - WARNUNG"
	
	IF Not fso.FolderExists(ziel) Then 
		fso.CreateFolder ziel
	End IF
	
	strLogFileName = ziel & "\" & strDateTime & ".txt"
	CreateLog strLogFileName, strDateTime
	
End IF


' Text box ausgabe
MsgBox Meldung,,"Ergebnis:"

Sub Add(text)
	' fügt Text hinzu
	Meldung = Meldung & text & vbCrLf
End Sub

Sub CreateLog(strLogFileName,strEventInfo)
	'http://msdn.microsoft.com/en-us/library/5t9b5c0c(v=vs.84).aspx
	Dim objFSO, objTextFile

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.CreateTextFile(strLogFileName, True)

	' Textblock
	objTextFile.WriteLine(strEventInfo)
	objTextFile.WriteLine " "

	IF fso.DriveExists("C:\") then objTextFile.WriteLine "Festplatte " & festplatteC End if
	IF fso.DriveExists("D:\") then objTextFile.WriteLine "Festplatte " & festplatteD End if
	IF fso.DriveExists("W:\") then objTextFile.WriteLine "Festplatte " & festplatteW End if
	IF fso.DriveExists("Z:\") then objTextFile.WriteLine "Festplatte " & festplatteZ End if
	IF fso.DriveExists("H:\") then objTextFile.WriteLine "Festplatte " & festplatteH End if
	IF fso.DriveExists("X:\") then objTextFile.WriteLine "Festplatte " & festplatteX End if
	IF fso.DriveExists("Y:\") then objTextFile.WriteLine "Festplatte " & festplatteY End if
	IF fso.DriveExists("E:\") then objTextFile.WriteLine "Festplatte " & festplatteE End if
	
	objTextFile.WriteLine " "
	'objTextFile.WriteLine "Timestamp: " & timestamp
	
	objTextFile.Close
	
End Sub



Function GetInfo( strComputer )
     Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
     Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
     For Each objItem in colItems
	 
	 
		IF objItem.Manufacturer = "Hewlett-Packard" Then
			
			IF Not fso.FileExists(folderfile) Then
				fso.CopyFile "Z:\2019_2020\PFE\IMD5\hp.txt", objShell.ExpandEnvironmentStrings("%userprofile%") & "\Desktop\StorageSystem\treiber\", True
			End IF
				
		End IF
		
		IF Trim(Left(Trim(objItem.Model), InStr(1, Trim(objItem.Model), " ", vbTextCompare))) = "DELL" Then
		
			IF Not fso.FileExists(folderfile) Then
				fso.CopyFile "Z:\2019_2020\PFE\IMD5\dell.txt", objShell.ExpandEnvironmentStrings("%userprofile%") & "\Desktop\StorageSystem\treiber\", True
			End IF
		End IF
		  
     next

End Function

' alles was in das csv ding reinkommt:
'dateineu.writeline space & " GB frei - erstellt am " & now & timstamp

set drive = nothing
 
set fso = nothing