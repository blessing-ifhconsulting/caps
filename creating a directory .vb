
' NewFolder.vbs
' Free example VBScript to create a folder (Simple)
' Author Guy Thomas http://computerperformance.co.uk/
' Version 2.4 - September 2010
' ------------------------------------------------' 
Option Explicit
Dim objFSO, objFolder, strDirectory
strDirectory = "c:\logs" 
' Create FileSystemObject. So we can apply .createFolder method
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Here is the key line Create a Folder, using strDirectory 
Set objFolder = objFSO.CreateFolder(strDirectory)
WScript.Echo "Just created " & strDirectory 
WScript.Quit 
' End of free example VBScript to create a folder