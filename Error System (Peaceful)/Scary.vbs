Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

oPlayer.URL = "s09tgo.mp3"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend