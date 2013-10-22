const staffLocalStateFile = "\Google\ChromeData\Local State"
const otherLocalStateFile = "\Google\Chrome\User Data\Local State"
Const ForReading = 1
Const ForWriting = 2 
Const touchFlag = "      ""enabled_labs_experiments"": [ ""touch-events@2"" ],"

Set oShell = CreateObject("WScript.Shell")
strAppData = oShell.ExpandEnvironmentStrings("%APPDATA%")
strLocalAppData = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
Set objFSO = CreateObject("Scripting.FileSystemObject") 

if (objFSO.fileExists(strAppData & staffLocalstatefile)) then
localStateFile = strAppData & staffLocalstatefile
continue = 1
end if

if (objFSO.fileExists(strLocalAppData & otherLocalstatefile)) then
localStateFile = strLocalAppData & otherLocalstatefile
continue = 1
end if

if continue <> 1 then
wscript.quit
end if


Set objFile = objFSO.OpenTextFile(localStateFile, ForReading) 

strText = objFile.ReadAll
strSplitText = split(strText, vbCrLf)
intTextLen = CInt(UBound(strSplitText))

objFile.Close 

for b = 0 to intTextLen
if inStr(strSplitText(b), "touch-events@2") then
	wscript.quit
end if
next

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

Set objFile = objFSO.OpenTextFile(localStateFile, ForWriting) 

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




