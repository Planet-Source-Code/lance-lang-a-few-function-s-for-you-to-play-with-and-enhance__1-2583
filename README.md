<div align="center">

## A Few Function's For You To Play With And Enhance\.


</div>

### Description

A Few function's For You To Play With.. Trap Mouse In A Form, Random Object/Form Color's, A Wacked Screen Closing Special Effect, And Download File's Via The Internet..
 
### More Info
 
These Are Just Basic Function's For All That Don't Know The Basic's.. Nothing Special...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lance Lang](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lance-lang.md)
**Level**          |Unknown
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lance-lang-a-few-function-s-for-you-to-play-with-and-enhance__1-2583/archive/master.zip)

### API Declarations

```
' Mouse Trap Declaration's, Toss These Into A Module
Option Explicit
Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Declare Function ClipCursor Lib "user32" _
(lpRect As Any) As Long
```


### Source Code

```
' Trapping And Releaseing Mouse Routine's -----Start
Public Function LetMouseGo(Frm2LetMouseGo As Object)
  Dim erg As Long
  Dim NewRect As RECT
  With NewRect
    .Left = 0&
    .Top = 0&
    .Right = Screen.Width / Screen.TwipsPerPixelX
    .Bottom = Screen.Height / Screen.TwipsPerPixelY
  End With
  erg& = ClipCursor(NewRect)
'Be Sure To Add
'
' Private Sub Form_Unload(Cancel As Integer)
' LetMouseGo Me
' End Sub
'
'To The Form That You Trap Incase They Ctrl-alt-Del Or X
'Out Of The Program, Otherwise, There Mouse Will Still Be
'Trapped In The Form Square!!
End Function
Public Function TrapMouse(Frm2MouseTrap As Object)
  Dim x As Long, y As Long, erg As Long
  Dim NewRect As RECT
  x& = Screen.TwipsPerPixelX
  y& = Screen.TwipsPerPixelY
  With NewRect
    .Left = Frm2MouseTrap.Left / x&
    .Top = Frm2MouseTrap.Top / y&
    .Right = .Left + Frm2MouseTrap.Width / x&
    .Bottom = .Top + Frm2MouseTrap.Height / y&
  End With
  erg& = ClipCursor(NewRect)
End Function
' Trapping And Releaseing Mouse Routine's -----End
' Random ForeColor Or BackColor Or FillColor On Form Or Object's ---Start
Public Function RandColor(ObjectToFlash As Object, ForeColorBackColorOrFillColor As Object)
  Dim c(2) As Byte
  For x = 0 To 2
    Randomize
    c(x) = Int((255 - 0 + 1) * Rnd + 0)
  Next x
  ObjectToFlash.ForeColorBackColorOrFillColor = RGB(c(0), c(1), c(2))
End Function
' Random ForeColor Or BackColor Or FillColor On Form Or Object's ---End
'Special Closing Affect ---Start
Public Function WickedFormClose(Form2Close As Object)
    GotoVal = (Form2Close.Height / 12)
    For Gointo = 1 To GotoVal
      DoEvents
        Form2Close.Height = Form2Close.Height - 50
        Form2Close.Top = (Screen.Height - Form2Close.Height) \ 2
        Form2Close.Width = Form2Close.Width - 50
        Form2Close.Left = (Screen.Width - Form2Close.Width) \ 2
        If Form2Close.Width <= 50 Then Unload Form2Close
        If Form2Close.Height <= 50 Then Unload Form2Close
      Next Gointo
Unload Form2Close
End Function
'Special Closing Affect ---End
'Retrieve File Off A WebPage Internet ---Start
' Usage Example
' GetInterNetFile "http://somewhere.com/ifsomething/", "test.zip", "c:"
' Note: You Have To Put A Microsoft Internet Transfer Control On The Form!
Public Function GetInterNetFile(Location As String, Filename As String, DirToSaveAt As String)
Dim mocha As String
mocha = Location & Filename
Dim bData() As Byte
Dim intFile As Integer
intFile = FreeFile()
bData() = Inet1.OpenURL(mocha, icByteArray)
Open DirToSaveAt & "\" & Filename For Binary Access Write _
As #intFile
Put #intFile, , bData()
Close #intFile
End Function
'Retrieve File Off The Internet ---End
' Yea, I know These Are Probably Crapily Coded But I'm Just Trying
' To Show The New People To VB Some Little Need (pointless)
' Thing's To Play Around With!!
```

