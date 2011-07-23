Const mPerLat = 111320 ' Meters per degree latitude
Const mPerLon = 85160 ' Meters per degree longitude at 44 degrees north
ToRad = 4*Atn(1)/180 ' PI/180

fcheckexchange
function fcheckexchange()
Set googleEarth = CreateObject("GoogleEarth.ApplicationGE")

Dim oSocket, iErr, sSocketText
sSocketText = ""
Set oSocket = CreateObject("Socket.TCP")
oSocket.DoTelnetEmulation = True
oSocket.TelnetEmulation = "TTY"
oSocket.Host = "localhost:5554"
On Error Resume Next 
oSocket.Open
iErr = Err.Number
 If iErr <> 0 Then 
   fCheckExchange = 0
   Exit Function 
  End If 
 sSocketText = oSocket.GetLine
 WScript.Echo sSocketText
 sSocketText = oSocket.GetLine
 WScript.Echo sSocketText
 while true
	set camPos = googleEarth.getCamera(true)
	
	horzDist = Sin(camPos.Tilt*ToRad)*camPos.Range
	latOff = Cos(camPos.Azimuth*ToRad)*horzDist
	lonOff = Sin(camPos.Azimuth*ToRad)*horzDist
	latitude = camPos.FocusPointLatitude-latOff/mPerLat
	longitude = camPos.FocusPointLongitude-lonOff/mPerLon
	
	WScript.Echo longitude&" "&latitude
	oSocket.SendLine "geo fix "&longitude&" "&latitude
	sSocketText = oSocket.GetLine
	WScript.Echo sSocketText
	WScript.sleep 200
 wend
 oSocket.Close
 On Error GoTo 0
 fCheckExchange = 0
end function

function sendCoords(latitude,longitude)
end function
