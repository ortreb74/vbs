dim prm
prm = "72"

'Connexion du lecteur reseau personnel UNIX
Set WshNetWork = CreateObject("Wscript.Network")

WshNetWork.MapNetworkDrive "J:", "\\adsamba03\ext-pdonzel\jira\siechaine-" & prm


WshNetWork.MapNetworkDrive "F:", "\\adsamba03\ext-pdonzel\liens\e2\dtd\custom\entities"
WshNetWork.MapNetworkDrive "G:", "\\adsamba03\ext-pdonzel\liens\e1\apps\eip\share\config"
WshNetWork.MapNetworkDrive "H:", "\\adsamba03\ext-pdonzel\liens\e1\apps\eip"

WshNetWork.MapNetworkDrive "I:", "\\adsamba03\ext-pdonzel"

WshNetWork.MapNetworkDrive "L:", "\\adsamba03\ext-pdonzel\liens\e10\progs\webbudev\uaur"

objNetwork.MapNetworkDrive "P:", "\\nas003\eflusers\ext-pdonzel"
WshNetWork.MapNetworkDrive "Q:", "\\adsamba03\ext-pdonzel\bas\Q"

WshNetWork.MapNetworkDrive "R:", "\\adsamba03\ext-pdonzel\liens\e3"
WshNetWork.MapNetworkDrive "S:", "\\adsamba03\ext-pdonzel\liens\e8"

WshNetWork.MapNetworkDrive "X:", "\\adsamba03\ext-pdonzel\excel"

