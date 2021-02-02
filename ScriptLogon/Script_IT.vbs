Option Explicit

Dim WSHNetwork, oNet, Printers, user

set WSHNetwork = WScript.CreateObject("WScript.Network")
Set oNet = CreateObject("Wscript.Network")

'contrôle de l'utilisateur
user = WshNetwork.UserName

' et transformation du username en majuscule
user = UCase(user)


 Select Case user
   
  Case "TIM.BROWN"

   oNet.MapNetworkDrive "E:", "\\ICT158-SRV03-1\Services\IT"
   WScript.Echo "User Name = " & user
   Printers = "\\ICT158-SRV03-1\HP_IT"
   WSHNetwork.AddWindowsPrinterConnection Printers

   WSHNetwork.SetDefaultPrinter Printers

 End Select

Set Printers = nothing
WSCript.Quit 
