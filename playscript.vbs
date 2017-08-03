' Spotifly random playlist
' open spotify and play random playlist when time < 1200
' close spotify when time > 1200
' inspired by 8aller and mrpixeltech on
'https://community.spotify.com/t5/Desktop-Linux-Windows-Web-Player/Is-there-a-command-to-play-a-playlist-using-command-line/td-p/952251

' TODO: only can play the last playlist, URI only open target playlist

Randomize
Dim timenow, selected_playlist
' determine path of spotify
Const spotify_path = "{}:\Users\{YOUR USER NAME}\AppData\Roaming\Spotify.exe"

Function selectPlaylist()
	Dim range, min, which
	' random number 1-4
	range = 4
	min = 1
	which = Int((range - min + 1) * Rnd() + min)
	' change your favourvite playlist in below (playlist URI)
	Select Case which
		Case 1
			' weekly discovery
			selectPlaylist = "spotify:user:spotify:playlist:37i9dQZEVXcSNrR35Ak4pi"
		Case 2
			' classic cafe time
			selectPlaylist = "spotify:user:sonymusictaiwan:playlist:7ABu8w2IP4kNPOaqouzzH5"
		Case 3
			' Mandopop
			selectPlaylist = "spotify:user:spotify:playlist:37i9dQZF1DWWqC43bGTcPc"
		Case 4
			' pop chillout
			selectPlaylist = "spotify:user:spotify:playlist:37i9dQZF1DXbIGqYf7WDxP"
		Case Else
			' last play list
			selectPlaylist = spotify_path
	End Select
End Function

Function startSpotifly()
	selected_playlist = selectPlaylist() ' random select playlist
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.Run(selected_playlist)
	' extra time delay for potential update
	WScript.sleep 10000
	' trigger focus on song in playlist by searching playlist
	WshShell.SendKeys "^f"
	WScript.sleep 200
	' escape the search
	WshShell.SendKeys "{ESC}"
	WScript.sleep 200
	' select a song to play by select all songs
	WshShell.SendKeys "^a"
	WScript.sleep 200
	' select one song
	WshShell.SendKeys "{UP}"
	WScript.sleep 200
	' play song in playlist
	WshShell.SendKeys "{ENTER}"
	WScript.sleep 100
	' play song in shuffle
	WshShell.SendKeys "^{RIGHT}"
End Function

Function stopSpotifly()
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.Run("TASKKILL /IM Spotify.exe")
End Function

' get current time
timenow = Hour(Now())
If (timenow < 12) Then
	' if it is morning
	startSpotifly()
Else
	' else if it is afternoon / evening
	stopSpotifly()
End If
