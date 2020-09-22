<div align="center">

## To play RTTTL \(nokia ring tone\) tunes


</div>

### Description

If you have a nokia mobile phone and looked at ring tones you will have come across RTTTL, the text format for the tunes.

This is a stand alone module with one public function PlayRTTTL. You give it a tune as a string in RTTTL format and it plays it using beeps.

Note that this only works on NT as the Beep function is different on windows.

If you are wondering what it could be used for, here is an example, at work we have written a phone book system for staff extension numbers and when you click on an entry you see details about the person and a picture. I wanted to let staff also give themselves a theme song that would play when you clicked on them. Since there are hundreds of RTTTL tunes available on the internet I decided to use that format as it is easily edited by users and saved to the database, and users can add new ones whenever they like.

The code could have been written better, but I wanted to keep it in a self contained single module that you could plug and play into any project.

This has nothing to do with Nokia mobile phones, it just uses the same format for the tunes.

If you have not seen them, this is an example of the format:

Simpsons:d=4,o=5,b=160:c.6,e6,f#6,8a6,g.6,e6,c6,8a,8f#, 8f#,8f#,2g,8p,8p,8f#,8f#,8f#,8g,a#.,8c6,8c6,8c6,c6

The Simpsons are probably copyrighted so don't use that one at home kids :)
 
### More Info
 
It takes a string containing the RTTTL tune

The Beeps are synchronous so be prepared to wait while it is playing.

You could avoid this by creating an exe that takes the RTTTL as a command line and shelling that from within your program.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BarryDunne](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/barrydunne.md)
**Level**          |Beginner
**User Rating**    |5.0 (35 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/barrydunne-to-play-rtttl-nokia-ring-tone-tunes__1-5645/archive/master.zip)





### Source Code

```
Option Explicit
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private colFrequencies As Collection
Public Sub PlayRTTTL(ByVal RTTTL As String)
 Dim colNotes As Collection
 Dim i As Long
 Set colNotes = GetNotesFromRTTTL(RTTTL)
 For i = 1 To colNotes.Count
  PlayNote Trim$(Left$(colNotes(i), 5)), Val(Mid$(colNotes(i), 5))
 Next i
End Sub
Private Sub PlayNote(ByVal sNote As String, ByVal lDuration As Long)
 On Error GoTo PlayNote_err
 Dim lFrequency As Long
 If colFrequencies Is Nothing Then
  Set colFrequencies = New Collection
  colFrequencies.Add 32.703, "C2"
  colFrequencies.Add 34.648, "C#2"
  colFrequencies.Add 36.708, "D2"
  colFrequencies.Add 38.891, "D#2"
  colFrequencies.Add 41.203, "E2"
  colFrequencies.Add 43.654, "F2"
  colFrequencies.Add 46.249, "F#2"
  colFrequencies.Add 48.999, "G2"
  colFrequencies.Add 51.913, "G#2"
  colFrequencies.Add 55, "A2"
  colFrequencies.Add 58.27, "A#2"
  colFrequencies.Add 61.735, "B2"
  colFrequencies.Add 65.406, "C3"
  colFrequencies.Add 69.296, "C#3"
  colFrequencies.Add 73.416, "D3"
  colFrequencies.Add 77.782, "D#3"
  colFrequencies.Add 82.407, "E3"
  colFrequencies.Add 87.307, "F3"
  colFrequencies.Add 92.499, "F#3"
  colFrequencies.Add 97.999, "G3"
  colFrequencies.Add 103.826, "G#3"
  colFrequencies.Add 110, "A3"
  colFrequencies.Add 116.541, "A#3"
  colFrequencies.Add 123.471, "B3"
  colFrequencies.Add 130.813, "C4"
  colFrequencies.Add 138.591, "C#4"
  colFrequencies.Add 146.832, "D4"
  colFrequencies.Add 155.564, "D#4"
  colFrequencies.Add 164.814, "E4"
  colFrequencies.Add 174.614, "F4"
  colFrequencies.Add 184.997, "F#4"
  colFrequencies.Add 195.998, "G4"
  colFrequencies.Add 207.652, "G#4"
  colFrequencies.Add 220, "A4"
  colFrequencies.Add 233.082, "A#4"
  colFrequencies.Add 246.942, "B4"
  colFrequencies.Add 261.626, "C5"
  colFrequencies.Add 277.183, "C#5"
  colFrequencies.Add 293.665, "D5"
  colFrequencies.Add 311.127, "D#5"
  colFrequencies.Add 329.628, "E5"
  colFrequencies.Add 349.228, "F5"
  colFrequencies.Add 369.994, "F#5"
  colFrequencies.Add 391.995, "G5"
  colFrequencies.Add 415.305, "G#5"
  colFrequencies.Add 440, "A5"
  colFrequencies.Add 466.164, "A#5"
  colFrequencies.Add 493.883, "B5"
  colFrequencies.Add 523.251, "C6"
  colFrequencies.Add 554.365, "C#6"
  colFrequencies.Add 587.33, "D6"
  colFrequencies.Add 622.254, "D#6"
  colFrequencies.Add 659.255, "E6"
  colFrequencies.Add 698.457, "F6"
  colFrequencies.Add 739.989, "F#6"
  colFrequencies.Add 783.991, "G6"
  colFrequencies.Add 830.609, "G#6"
  colFrequencies.Add 880, "A6"
  colFrequencies.Add 932.328, "A#6"
  colFrequencies.Add 987.767, "B6"
  colFrequencies.Add 1046.502, "C7"
  colFrequencies.Add 1108.731, "C#7"
  colFrequencies.Add 1174.659, "D7"
  colFrequencies.Add 1244.508, "D#7"
  colFrequencies.Add 1318.51, "E7"
  colFrequencies.Add 1396.913, "F7"
  colFrequencies.Add 1479.978, "F#7"
  colFrequencies.Add 1567.982, "G7"
  colFrequencies.Add 1661.219, "G#7"
  colFrequencies.Add 1760, "A7"
  colFrequencies.Add 1864.655, "A#7"
  colFrequencies.Add 1975.533, "B7"
  colFrequencies.Add 2093.005, "C8"
  colFrequencies.Add 2217.461, "C#8"
  colFrequencies.Add 2349.318, "D8"
  colFrequencies.Add 2489.016, "D#8"
  colFrequencies.Add 2637.021, "E8"
  colFrequencies.Add 2793.826, "F8"
  colFrequencies.Add 2959.956, "F#8"
  colFrequencies.Add 3135.964, "G8"
  colFrequencies.Add 3322.438, "G#8"
  colFrequencies.Add 3520, "A8"
  colFrequencies.Add 3729.31, "A#8"
  colFrequencies.Add 3951.066, "B8"
  colFrequencies.Add 4186.009, "C9"
  colFrequencies.Add 4434.922, "C#9"
  colFrequencies.Add 4698.637, "D9"
  colFrequencies.Add 4978.032, "D#9"
  colFrequencies.Add 5274.042, "E9"
  colFrequencies.Add 5587.652, "F9"
  colFrequencies.Add 5919.912, "F#9"
  colFrequencies.Add 6271.928, "G9"
  colFrequencies.Add 6644.876, "G#9"
  colFrequencies.Add 7040, "A9"
  colFrequencies.Add 7458.62, "A#9"
  colFrequencies.Add 7902.133, "B9"
  colFrequencies.Add 8372.019, "C10"
  colFrequencies.Add 8869.845, "C#10"
  colFrequencies.Add 9397.273, "D10"
  colFrequencies.Add 9956.064, "D#10"
  colFrequencies.Add 10548.083, "E10"
  colFrequencies.Add 11175.305, "F10"
  colFrequencies.Add 11839.823, "F#10"
  colFrequencies.Add 12543.855, "G10"
  colFrequencies.Add 13289.752, "G#10"
 End If
 DoEvents
 If UCase$(Mid$(sNote, 1, 1)) = "P" Then 'pause
  Sleep lDuration
 Else
  lFrequency = CLng(colFrequencies(UCase$(sNote)))
  Beep lFrequency, lDuration
 End If
 Exit Sub
PlayNote_err:
 Debug.Print Err.Number & ": " & Err.Description
End Sub
Private Function GetNotesFromRTTTL(ByVal RTTTL As String) As Collection
 Dim lDefDuration As Long
 Dim lDefScale As Long
 Dim lBPM As Long
 Dim lStart As Long
 Dim sNote As String
 Dim lDuration As Long
 Set GetNotesFromRTTTL = New Collection
 'Get default values
 lDefDuration = GetDefaultFromRTTTL(RTTTL, "d", 4)
 lDefScale = GetDefaultFromRTTTL(RTTTL, "o", 6)
 lBPM = GetDefaultFromRTTTL(RTTTL, "b", 63)
 'Find first note
 lStart = InStr(1, RTTTL, ":")
 If InStr(lStart + 1, RTTTL, ":") > 0 Then
  lStart = InStr(lStart + 1, RTTTL, ":")
 End If
 lStart = lStart + 1
 'Parse notes
 Do Until lStart = 1
  sNote = GetNoteNameFromRTTTL(RTTTL, lStart, lDefScale)
  lDuration = GetNoteDurationFromRTTTL(RTTTL, lStart, lDefDuration, lBPM)
  GetNotesFromRTTTL.Add sNote & Space$(5 - Len(sNote)) & lDuration
  lStart = InStr(lStart + 1, RTTTL, ",") + 1
 Loop
End Function
Private Function GetDefaultFromRTTTL(ByVal RTTTL As String, ByVal sType As String, lDefault As Long) As Long
 Dim lPos As Long
 lPos = InStr(1, RTTTL, sType & "=")
 If lPos > 0 Then
  Do While IsNumeric(Mid$(RTTTL, lPos + 2, 1))
   GetDefaultFromRTTTL = GetDefaultFromRTTTL * 10 + Val(Mid$(RTTTL, lPos + 2, 1))
   lPos = lPos + 1
  Loop
 Else
  GetDefaultFromRTTTL = lDefault
 End If
End Function
Private Function GetNoteNameFromRTTTL(ByVal RTTTL As String, ByVal lStart As Long, ByVal lDefScale As Long) As String
 Dim lPos As Long
 Dim sTemp As String
 lPos = InStr(lStart, RTTTL, ",")
 If lPos > 0 Then
  sTemp = UCase$(Mid$(RTTTL, lStart, lPos - lStart))
 Else
  sTemp = UCase$(Mid$(RTTTL, lStart))
 End If
 sTemp = Trim$(sTemp)
 If Len(sTemp) = 0 Then
  Exit Function
 End If
 'Remove duration, if any
 Do While IsNumeric(Left$(sTemp, 1))
  sTemp = Mid$(sTemp, 2)
 Loop
 'Remove any dots
 sTemp = FindAndReplace(sTemp, ".", "")
 GetNoteNameFromRTTTL = sTemp
 'Add default scale if not given
 If Mid$(sTemp, 2, 1) = "#" Then
  If Len(sTemp) = 2 Then
   GetNoteNameFromRTTTL = sTemp & lDefScale
  End If
 Else
  If Len(sTemp) = 1 Then
   GetNoteNameFromRTTTL = sTemp & lDefScale
  End If
 End If
End Function
Private Function GetNoteDurationFromRTTTL(ByVal RTTTL As String, ByVal lStart As Long, ByVal lDefDuration As Long, ByVal lBPM As Long) As Long
 Dim lPos As Long
 Dim sTemp As String
 Dim lDur As Long
 lPos = InStr(lStart, RTTTL, ",")
 If lPos > 0 Then
  sTemp = UCase$(Mid$(RTTTL, lStart, lPos - lStart))
 Else
  sTemp = UCase$(Mid$(RTTTL, lStart))
 End If
 If Len(sTemp) = 0 Then
  Exit Function
 End If
 'See if any duration given for note
 lPos = 1
 If IsNumeric(Mid$(sTemp, lPos, 1)) Then
  Do While IsNumeric(Mid$(sTemp, lPos, 1))
   lDur = lDur & Mid$(sTemp, lPos, 1)
   lPos = lPos + 1
  Loop
 Else
  lDur = lDefDuration
 End If
 GetNoteDurationFromRTTTL = (4 * 60000) / (lBPM * lDur)
 'check for a .
 If InStr(1, sTemp, ".") > 0 Then
  GetNoteDurationFromRTTTL = GetNoteDurationFromRTTTL * 1.5
 End If
End Function
Private Function FindAndReplace(ByVal sOriginal As String, ByVal sFind As String, ByVal sReplace As String, Optional ByVal bCaseSensitive As Boolean = True) As String
 Dim lPos As Long
 FindAndReplace = sOriginal
 If Len(sFind) = 0 Then
  Exit Function
 End If
 If bCaseSensitive Then
  lPos = InStr(1, sOriginal, sFind, vbBinaryCompare)
 Else
  lPos = InStr(1, sOriginal, sFind, vbTextCompare)
 End If
 Do While lPos > 0
  FindAndReplace = Mid$(FindAndReplace, 1, lPos - 1) & sReplace & Mid$(FindAndReplace, lPos + Len(sFind))
  If bCaseSensitive Then
   lPos = InStr(lPos + Len(sReplace), FindAndReplace, sFind, vbBinaryCompare)
  Else
   lPos = InStr(lPos + Len(sReplace), FindAndReplace, sFind, vbTextCompare)
  End If
 Loop
End Function
```

