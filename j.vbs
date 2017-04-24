dim prm
prm = "72"

Dim objNetwork
Set objNetwork = CreateObject("WScript.Network")

Set oDrives = objNetwork.EnumNetworkDrives
dim test
mappingDrives = false

For i = 0 to oDrives.Count - 1 Step 2   
   if oDrives.iTem(i) = "I:" then mappingDrives = true   
Next

if not(mappingDrives) then
	objNetwork.MapNetworkDrive "F:", "\\adsamba03\ext-pdonzel\liens\e2\dtd\custom\entities"
	objNetwork.MapNetworkDrive "G:", "\\adsamba03\ext-pdonzel\liens\e1\apps\eip\share\config"
	objNetwork.MapNetworkDrive "H:", "\\adsamba03\ext-pdonzel\liens\e1\apps\eip"

	objNetwork.MapNetworkDrive "I:", "\\adsamba03\ext-pdonzel"

	objNetwork.MapNetworkDrive "L:", "\\adsamba03\ext-pdonzel\liens\e10\progs\webbudev\uaur"

	objNetwork.MapNetworkDrive "P:", "\\nas003\eflusers\ext-pdonzel"
	objNetwork.MapNetworkDrive "Q:", "\\adsamba03\ext-pdonzel\bas\Q"

	objNetwork.MapNetworkDrive "R:", "\\adsamba03\ext-pdonzel\liens\e3"
	objNetwork.MapNetworkDrive "S:", "\\adsamba03\ext-pdonzel\liens\e8"

	objNetwork.MapNetworkDrive "X:", "\\adsamba03\ext-pdonzel\excel"
else
on error resume next
	objNetwork.RemoveNetworkDrive "J:"
on error goto 0
end if

Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

strFolder = "i:\jira\siechaine-" & prm 

If Not oFSO.FolderExists(strFolder) Then
  oFSO.CreateFolder strFolder
End If

objNetwork.MapNetworkDrive "J:", "\\adsamba03\ext-pdonzel\jira\siechaine-" & prm 