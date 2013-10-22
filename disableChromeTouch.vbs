Const ForReading = 1
Const ForWriting = 2 
Const touchFlag = "      ""enabled_labs_experiments"": [ ""touch-events@2"" ],"

localStateFile = findLocalStateFile()
strSplitText = readToArray(localStateFile)
call quitIfLineFound(strSplitText, touchFlag)
call killChrome()
call writeChromeLocalStateFile(localStateFile,strSplitText)

function findLocalStateFile()
	Const staffLocalStateFile = "\Google\ChromeData\Local State"
	Const otherLocalStateFile = "\Google\Chrome\User Data\Local State"
	
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	Set oShell = CreateObject("WScript.Shell")
	
	strAppData = oShell.ExpandEnvironmentStrings("%APPDATA%")
	strLocalAppData = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
	
	if (objFSO.fileExists(strAppData & staffLocalstatefile)) then
	findLocalStateFile = strAppData & staffLocalstatefile
	end if
	
	if (objFSO.fileExists(strLocalAppData & otherLocalstatefile)) then
	findLocalStateFile = strLocalAppData & otherLocalstatefile
	end if
	
	if findLocalStateFile = "" then
		msgbox("Could not find localstate file or chrome not installed. Please contact IT for assistance")
		wscript.quit
	end if
end function

function killChrome()
	Set oShell = CreateObject("WScript.Shell")
	set service = GetObject ("winmgmts:")
	chromeRunning = 0
	for each Process in Service.InstancesOf ("Win32_Process")
		If Process.Name = "chrome.exe" then
			chromeRunning = 1
		End If
	next
	if chromeRunning = 1 then
	oShell.Run "cmd /c TASKKILL /F /IM CHROME.EXE /T" 
	end if
end function

function readToArray(localStateFile)
	Const ForReading = 1
	Const ForWriting = 2 
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	Set objFile = objFSO.OpenTextFile(localStateFile, ForReading) 
	strText = objFile.ReadAll
	readToArray = split(strText, vbCrLf)
	objFile.Close 
end function

function quitIfLineFound(strSplitText, lineToFind)
	intTextLen = CInt(UBound(strSplitText))
	for b = 0 to intTextLen
	if inStr(strSplitText(b), lineToFind) then
		wscript.quit
	end if
	next
end function

function writeChromeLocalStateFile(localStateFile,strSplitText)
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	Set objFile = objFSO.OpenTextFile(localStateFile, ForWriting) 
	intTextLen = CInt(UBound(strSplitText))
	for c = 0 to intTextLen
		if inStr(strSplitText(c), "browser"": {") then
			objFile.WriteLine strSplitText(c)
			objFile.WriteLine touchFlag
		else
			if strSplitText(c) <> "" then
				if inStr(strSplitText(c), "touch-events@") then
					'do nothing
				else
					objFile.WriteLine strSplitText(c)
				end if
			end if
		end if
	next
	objFile.Close 
end function
