Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

oPlayer.URL = "C:\Windows\Media\Windows Unlock.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 2
Wend

oPlayer.URL = "C:\Windows\Media\Windows Logoff Sound.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 20
Wend

oPlayer.URL = "C:\Windows\Media\Windows User Account Control.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 10
Wend

oPlayer.URL = "C:\Windows\Media\Windows Shutdown.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 2
Wend

oPlayer.URL = "C:\Windows\Media\Windows Logoff Sound.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 6
Wend

oPlayer.URL = "C:\Windows\Media\Windows Shutdown.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 1
Wend

oPlayer.URL = "C:\Windows\Media\Windows User Account Control.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 2
Wend

oPlayer.URL = "C:\Windows\Media\Windows User Account Control.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 2
Wend

oPlayer.URL = "C:\Windows\Media\Windows Unlock.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 2
Wend

oPlayer.URL = "C:\Windows\Media\Windows Unlock.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 2
Wend

oPlayer.URL = "C:\Windows\Media\Windows Unlock.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 2
Wend

oPlayer.URL = "C:\Windows\Media\Windows Shutdown.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 2
Wend

oPlayer.URL = "C:\Windows\Media\Windows Shutdown.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 1
Wend

oPlayer.URL = "C:\Windows\Media\Windows Shutdown.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 1
Wend