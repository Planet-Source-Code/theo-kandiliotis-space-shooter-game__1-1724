<div align="center">

## Space Shooter Game


</div>

### Description

Its a space shooter game, ok for all you people who can't read the address is

www.geocities.com/baja/cliffs/8036/space.html
 
### More Info
 
LOOKS LIKE A LOT BUT MOST IS JUST STEP BY STEP INSTRUCTIONS:)

There are files that are needed. Go to www.geocities.com/baja/cliffs/8036/ space.html to get them

Its eaisier to download it but heres the code.

There are a lot of objects(mostly pictures and sprites)

a few bugs that are still being worked on. Nothing serious:)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Theo Kandiliotis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/theo-kandiliotis.md)
**Level**          |Unknown
**User Rating**    |6.0 (600 globes from 100 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/theo-kandiliotis-space-shooter-game__1-1724/archive/master.zip)

### API Declarations

```
'The coordinates,velocity, and size of the
'stars that are moving on the background
Dim x(30), Y(30), pace(30), size(30) As Integer
'The coordinates of the spaceship sprite
'(the good guy...)
Dim x2, y2, xn As Integer
'Variable that defines the movement of the
'good guy (takes the values "Left" and
'"Right")
Dim movement As String
'Variable that is used to create the
'"laser blaster" effect when the good guy
'shoots.At first,it takes the value 1 and
'a line is drawn on the form ,either blue
'or white.Then it takes the value 2 and a
'new line line is drawn (the other color)
'Then a black line is drawn in order to
'erase the shot (the background of the form
'is black,too)
Dim shot As Integer
'The X coordinate of the line that represents
'the good guy's gunfire
Dim xshot As Integer
'Variables that define the movement of the
'5 rows of aliens.They take the values "Left"
'and "Right"
Dim Mi1, Mi2, Mi3, Mi4, Mi5 As String
'X and Y Coordinates of the top alien
Dim xi1, yi1 As Integer
'X and Y coordinates of the aliens in
'rows 2-5
Dim xi2(2), xi3(3), xi4(4), xi5(5) As Integer
Dim yi2(2), yi3(3), yi4(4), yi5(5) As Integer
'Variable that is used to define which alien
'is beign shot each time so that the frames
'of the explosion are painted at the right
'spot of the form ( is assigned a X
'coordinate of an alien,ie xinv=xi2(2) )
Dim xinv, yinv As Variant
'Integer that defines where the laser beem
'will stop (it's a Y coordinate in fact)
Dim upper As Integer
'Array that defines which aliens are dead.
'Counting starts from the top row.ie if
'the good guy kills the second alien from
'the right in the 5th row,the array element
'kill(14) will be assigned the value True
Dim kill(15) As Boolean
'Integer that counts how many aliens are
'dead.If it gets the value 15 (meaning the
'good huy cleared the level) the game starts
'over again.
Dim dead As Integer
'Boolean type variable that's set to True
'while an explosion is painted on the form
'It is used to halt some other procedures
'in order to prevent flickering and bugs
Dim boom As Boolean
'Defines which frame of the explosion must
'be painted the next time the Timer event
'occurs (that's every 0,001 sec).
Dim explosion As Integer
'Color contants that are used with the Shot
'variable to create the "laser blaster"
'effect when the good guy shoots
Dim COL1, COL2 As ColorConstants
'Guess what this one means ...
Dim YouDied As Boolean
'Variable holding the score.In fact it is
'the number of aliens that are dead(variable
'Dead) multiplied by 1000
Dim score
'Integer that is assigned a random value from
'0 to 10 and defines the delay before the
'next movement command is executed
Dim Wait As Integer
```


### Source Code

```
'The procedure that runs after Form_Load.It
'It is used to give starting values to all
'the variables that must be reset every
'time a new level begins
Private Sub Form_Activate()
'Set starting values for the movement of each row
'of aliens.All the aliens in one row move towards
'the same direction and change direction when the
'far right or far left alien hits the edge of the
'form.This is why killing the aliens on the far
'right and far left slows down their vertical movement
'start the background midi.
Mi1 = "LEFT"
Mi2 = "RIGHT"
Mi3 = "LEFT"
Mi4 = "RIGHT"
Mi5 = "LEFT"
Timer1.Enabled = True
Cls
dead = 0
Form1.KeyPreview = True
Randomize
'This code sets the coordinates,velocity
'and size of the 30 small circles that
'contantly move on the background.
 For i = 1 To 30
 x(i) = Int(Form1.Width * Rnd)
 Y(i) = Int(Form1.Height * Rnd)
 pace(i) = Int(500 - (Int(Rnd * 499)))
 size(i) = 14 - (13 * Rnd)
 Next
'Set starting values for the coordinates
'of the spaceship sprite
 x2 = 3760
 y2 = 5600
'Set starting values for the coordinates of
'the 15 aliens.The syntax
' For I=1 to N
' X=(Container.Width * (N-I)/N)-(Control.Width/2)
' Next
'can be used to horizontaly center N identical
'controls in a container (Form,picture box etc)
 xi1 = (Form1.Width / 2) - (sprINVADER5.Width / 2)
 yi1 = 1000
 For i = 1 To 2
 yi2(i) = yi1 + sprINVADER5.Height + 50
 Next
 xi2(1) = (Form1.Width / 2) - (Form1.Width / 8) - (sprINVADER5.Width / 2)
 xi2(2) = (Form1.Width / 2) + (Form1.Width / 8) - (sprINVADER5.Width / 2)
 For i = 1 To 3
 yi3(i) = yi2(1) + sprINVADER5.Height + 50
 xi3(i) = ((Form1.Width * (4 - i) / 4) - (sprINVADER5.Width / 2))
 Next
 For i = 1 To 4
 yi4(i) = yi3(1) + sprINVADER5.Height + 200
 xi4(i) = ((Form1.Width * (5 - i) / 5) - (sprINVADER5.Width) / 2)
 Next
 For i = 1 To 5
 yi5(i) = yi4(1) + sprINVADER5.Height + 300
 xi5(i) = ((Form1.Width * (6 - i) / 6) - (sprINVADER5.Width) / 2)
 Next
End Sub
'The procedure that would normally run when
'the user hits the cursor keys or the space bar
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeySpace
Call fire
Case vbKeyLeft
movement = "Left"
Case vbKeyRight
movement = "Right"
Case Else
movement = ""
End Select
End Sub
'This procedure is here in case you click the X
'button on the upper right of the form instead
'of the exit button.It makes sure that the text
'file closes before the application ends
Private Sub Form_Unload(Cancel As Integer)
cmdExit_Click
End Sub
'The procedure that executes every 0,001 sec and
'controls just about everything goes on in the
'game.The general idea is that event procedures all
'all around the project alter the values of FLAG
'variables like BOOM,SHOT,X2 etc and this procedure
'uses those values to do whatever the event is about.
Private Sub Timer1_Timer()
'---------------------------------------
'Check if the top alien has crushed on the ship
If (y2 - (yi1 + sprINVADER5.Height)) < -100 And Abs((xi1 + (sprINVADER5.Width / 2)) - (x2 + sprSHIP.Width / 2)) < (sprINVADER5.Width / 2) Then
'and if so
Timer1.Enabled = False
For q = 1 To 6
PaintPicture imgKABOOM(q).Picture, x2 + 600, y2
For l = 1 To 100000: Next
Next
Form_Activate
'halt everything (Timer1.Enabled=false),
'paint the 6 frames of the explosion on the good
'guy ( that is now a dead good guy) and start over
'a new level (Form_Activate)
End If
'Check if an alien from the second row has crushed
'on the ship
For i = 1 To 2
If (y2 - (yi2(i) + sprINVADER5.Height)) < -100 And Abs((xi2(i) + (sprINVADER5.Width / 2)) - (x2 + sprSHIP.Width / 2)) < (sprINVADER5.Width / 2) Then
Timer1.Enabled = False
For q = 1 To 6
PaintPicture imgKABOOM(q).Picture, x2 + 600, y2
For l = 1 To 100000: Next
Next
Form_Activate
End If
Next
'Check if an alien from the 3rd row has crushed
'on the ship
For i = 1 To 3
If (y2 - (yi3(i) + sprINVADER5.Height)) < -100 And Abs((xi3(i) + (sprINVADER5.Width / 2)) - (x2 + sprSHIP.Width / 2)) < (sprINVADER5.Width / 2) Then
Timer1.Enabled = False
For q = 1 To 6
PaintPicture imgKABOOM(q).Picture, x2 + 600, y2
For l = 1 To 100000: Next
Next
Form_Activate
End If
Next
'Check if an alien from the 4th row has crushed
'on the ship
For i = 1 To 4
If (y2 - (yi4(i) + sprINVADER5.Height)) < -100 And Abs((xi4(i) + (sprINVADER5.Width / 2)) - (x2 + sprSHIP.Width / 2)) < (sprINVADER5.Width / 2) Then
Timer1.Enabled = False
For q = 1 To 6
PaintPicture imgKABOOM(q).Picture, x2 + 600, y2
For l = 1 To 100000: Next
Next
Form_Activate
End If
Next
'Check if an alien from the lower row has crushed
'on the ship
For i = 1 To 5
If (y2 - (yi5(i) + sprINVADER5.Height)) < -100 And Abs((xi5(i) + (sprINVADER5.Width / 2)) - (x2 + sprSHIP.Width / 2)) < (sprINVADER5.Width / 2) Then
Timer1.Enabled = False
For q = 1 To 6
PaintPicture imgKABOOM(q).Picture, x2 + 600, y2
For l = 1 To 100000: Next
Next
Form_Activate
End If
Next
'If alien #1 is not dead...
Select Case kill(1)
Case False
Select Case xi1
Case Is > -230
'and he's on the visible part of the Form
'** (when an alien is shot,his X coordinate is given
' a very low value (-5000) so that if by any
' chance his sprite is painted on the form,you wont
' see him)
Select Case Mi1
'Then move him towards the direction that the
'Mi1 variable points out
Case "LEFT"
xi1 = xi1 - 200
If xi1 < 0 Then Mi1 = "RIGHT": yi1 = yi1 + 155
Case "RIGHT"
xi1 = xi1 + 200
If xi1 > Form1.Width - sprINVADER5.Width Then Mi1 = "LEFT": yi1 = yi1 + 155
End Select
End Select
End Select
'Move the 2nd row aliens
Select Case Mi2
Case "LEFT"
If kill(2) = False And xi2(1) > 100 Then xi2(1) = xi2(1) - 200
If kill(3) = False And xi2(2) > 100 Then xi2(2) = xi2(2) - 200
'** (only if they're alive...)
If xi2(1) < 200 And kill(2) = False And xi2(1) > -100 Then
Mi2 = "RIGHT"
For q = 1 To 2
If boom = True And xinv = xi2(q) Then Exit For
yi2(q) = yi2(q) + 155
Next
End If
If xi2(2) < 200 And kill(3) = False And xi2(2) > -100 Then
Mi2 = "RIGHT"
For q = 1 To 2
If boom = True And xinv = xi2(q) Then Exit For
yi2(q) = yi2(q) + 155
Next
End If
GoTo 2
Case "RIGHT"
If kill(2) = False And xi2(1) > -100 And xi2(1) < (Form1.Width - sprINVADER5.Width) Then xi2(1) = xi2(1) + 200
If kill(3) = False And xi2(2) > -100 And xi2(2) < (Form1.Width - sprINVADER5.Width) Then xi2(2) = xi2(2) + 200
If xi2(1) > (Form1.Width - sprINVADER5.Width) And kill(2) = False Then
Mi2 = "LEFT"
For q = 1 To 2
If boom = True And xinv = xi2(q) Then Exit For
yi2(q) = yi2(q) + 155
Next
End If
If xi2(2) > (Form1.Width - sprINVADER5.Width) And kill(3) = False Then
Mi2 = "LEFT"
For q = 1 To 2
If boom = True And xinv = xi2(q) Then Exit For
yi2(q) = yi2(q) + 155
Next
End If
2 End Select
'Move the third row aliens
Select Case Mi3
Case "LEFT"
For i = 1 To 3
If kill(i + 3) = False And xi3(i) > 100 Then xi3(i) = xi3(i) - 200
If xi3(i) < 200 And kill(3 + i) = False And xi3(i) > -100 Then
Mi3 = "RIGHT"
For q = 1 To 3
If boom = True And xinv = xi3(q) Then Exit For
yi3(q) = yi3(q) + 155
Next
End If
Next
GoTo 3
Case "RIGHT"
For i = 1 To 3
If kill(3 + i) = False And xi3(i) > -100 And xi3(i) < (Form1.Width - sprINVADER5.Width) Then xi3(i) = xi3(i) + 200
If xi3(i) > (Form1.Width - sprINVADER5.Width) And kill(3 + i) = False Then
Mi3 = "LEFT"
For q = 1 To 3
If boom = True And xinv = xi3(q) Then Exit For
yi3(q) = yi3(q) + 155
Next
End If
Next
3 End Select
'Move the fourth row aliens
Select Case Mi4
Case "LEFT"
For i = 1 To 4
If kill(6 + i) = False And xi4(i) > 100 Then xi4(i) = xi4(i) - 200
If xi4(i) < 200 And kill(6 + i) = False And xi4(i) > -100 Then
Mi4 = "RIGHT"
For q = 1 To 4
If boom = True And xinv = xi4(q) Then Exit For
yi4(q) = yi4(q) + 155
Next
End If
Next
GoTo 4
Case "RIGHT"
For i = 1 To 4
If kill(6 + i) = False And xi4(i) > -100 And xi4(i) < (Form1.Width - sprINVADER5.Width) Then xi4(i) = xi4(i) + 200
If xi4(i) > (Form1.Width - sprINVADER5.Width) And kill(6 + i) = False Then
Mi4 = "LEFT"
For q = 1 To 4
If boom = True And xinv = xi4(q) Then Exit For
yi4(q) = yi4(q) + 155
Next
End If
Next
4 End Select
'Move the 5th row aliens
Select Case Mi5
Case "LEFT"
For i = 1 To 5
If kill(10 + i) = False And xi5(i) > 100 Then xi5(i) = xi5(i) - 200
If xi5(i) < 200 And kill(10 + i) = False And xi5(i) > -100 Then
Mi5 = "RIGHT"
For q = 1 To 5
If boom = True And xinv = xi5(q) Then Exit For
yi5(q) = yi5(q) + 155
Next
End If
Next
GoTo 5
Case "RIGHT"
For i = 1 To 5
If kill(10 + i) = False And xi5(i) > -100 And xi5(i) < (Form1.Width - sprINVADER5.Width) Then xi5(i) = xi5(i) + 200
If xi5(i) > (Form1.Width - sprINVADER5.Width) And kill(10 + i) = False Then
Mi5 = "LEFT"
For q = 1 To 5
If boom = True And xinv = xi5(q) Then Exit For
yi5(q) = yi5(q) + 155
Next
End If
Next
5 End Select
'If the good guy killed 15 aliens,start
'a new level by calling the Form_Activate procedure
Select Case dead
Case 15
dead = 0: Cls
Call Form_Activate
End Select
'If there is an explosion going on somewhere on the
'form (Boom=true) then paint the correct frame of
'it (the variable Explosion represents the number
'of the frame that must be currently painted)
Select Case boom
Case True
Select Case explosion
Case 1
'If the explosion has just started (explosion=1)
'then don't paint any frame of it,just erase the
'alien that has been shot by painting the label
'lblBlank on him.
'lnlBlank is a black label used every now and then
'to erase something from the form.It's Left and
'Top properties have been assigned the X and Y
'coordinates of the alien that's being killed
'as you read these lines.This was done from the
'Kaboom procedure,where the Boom variable was
'assigned the value True and the whole explosion
'buisness began
lblBlank.Visible = True
lblBlank.Visible = False
Case Is <> 1
'If the explosion has already started before the
'current timer event then paint the current frame
'of the explosion on the alien that was just killed.
'Xinv and Yinv are the coordinates of the alien
'that get's his butt kicked every time the good
'guy hits bullseye.
Select Case xinv
Case xi1
PaintPicture imgKABOOM(explosion).Picture, xi1, yi1 - 100
Case xi2(1)
PaintPicture imgKABOOM(explosion).Picture, xi2(1) + 200, yi2(1) - 100
Case xi2(2)
PaintPicture imgKABOOM(explosion).Picture, xi2(2) + 200, yi2(2) - 100
Case xi3(1)
PaintPicture imgKABOOM(explosion).Picture, xi3(1) + 150, yi3(1) - 100
Case xi3(2)
PaintPicture imgKABOOM(explosion).Picture, xi3(2) + 150, yi3(2) - 100
Case xi3(3)
PaintPicture imgKABOOM(explosion).Picture, xi3(3) + 150, yi3(3) - 100
Case xi4(1)
PaintPicture imgKABOOM(explosion).Picture, xi4(1) - 30, yi4(1) - 100
Case xi4(2)
PaintPicture imgKABOOM(explosion).Picture, xi4(2) - 30, yi4(2) - 100
Case xi4(3)
PaintPicture imgKABOOM(explosion).Picture, xi4(3) - 30, yi4(3) - 100
Case xi4(4)
PaintPicture imgKABOOM(explosion).Picture, xi4(4) - 30, yi4(4) - 100
Case xi5(1)
PaintPicture imgKABOOM(explosion).Picture, xi5(1) - 30, yi5(1) - 100
Case xi5(2)
PaintPicture imgKABOOM(explosion).Picture, xi5(2) - 30, yi5(2) - 100
Case xi5(3)
PaintPicture imgKABOOM(explosion).Picture, xi5(3) - 30, yi5(3) - 100
Case xi5(4)
PaintPicture imgKABOOM(explosion).Picture, xi5(4) - 30, yi5(4) - 100
Case xi5(5)
PaintPicture imgKABOOM(explosion).Picture, xi5(5) - 30, yi5(5) - 100
End Select
End Select
'Add 1 to the value of Explosion so that a new frame
'of it will be painted in the next timer event
explosion = explosion + 1
Select Case explosion
Case 7
'if all the frames of the explosion have been painted
'then the alien whose coordinates are the same
'with the values of Xinv and Yinv is officialy dead
'and the fun ends (Boom=false)
explosion = 0
boom = False
If YouDied = True Then Form_Activate
'Let's see that score increase...
score = score + 1000
'Onother one bites the dust.(15-dead) to go
dead = dead + 1
'paint the blank label on the exact spot where
'the explosion frames were painted incase there
'is some smoke left floating there.
'**(There is still a little bug and sometimes
' (the smoke remains there.Oh well...)
lblBlank.Left = xinv
lblBlank.Top = yinv - 100
lblBlank.Visible = True
lblBlank.Visible = False
'If an alien was just killed,then give his X coordinate
'a very low value (-5000) so that he wont be visible
'on the form
If kill(1) = True Then xi1 = -5000: kill(1) = False
For i = 1 To 2
If kill(i + 1) = True Then xi2(i) = -5000: kill(i + 1) = False: Exit For
Next
For i = 1 To 3
If kill(3 + i) = True Then xi3(i) = -5000: kill(3 + i) = False: Exit For
Next
If kill(7) = True Then xi4(1) = -5000: kill(7) = False
If kill(8) = True Then xi4(2) = -5000:: kill(8) = False
If kill(9) = True Then xi4(3) = -5000: kill(9) = False
If kill(10) = True Then xi4(4) = -5000: kill(10) = False
For i = 1 To 5
If kill(10 + i) = True Then xi5(i) = -5000: kill(10 + i) = False
Next
End Select
Case False
End Select
'If the good guy has just fired his blaster,draw
'a line .The upper bound is determined by how
'high the alien that will be killed is.(if this
'shot wont kill an alien then it will go all the
'way up to the upper edge of the form).In fact
'each time the good guy shoots ,3 lines are drawn
'one after the other on the exact same spot and
'with the exact same length.The first is either
'blue or white,the second is either white or
'blue and the third is black so that
'the "shot ring" will be erased.This is how the
'cool "laser blaster" effect is done
Select Case shot
Case 1
Line (xshot, y2)-(xshot, upper), COL2
shot = 2
Case 2
Line (xshot, y2)-(xshot, upper), BackColor
shot = 0
End Select
'Move the spaceship according to the value of the
'variable Movement.As in most shoot'em up games,the
'ship continues moving towards the direction it has
'been last guided to until it meets the edge of the
'form or onother direction is given.Normally this
'would be done by the user hitting a cursor key
'but in this demo it's done by inputing a new
'keyword command from the file DEMO.DAT
Select Case movement
Case "Left"
x2 = x2 - 210
If x2 < -320 Then x2 = -320
PaintPicture sprSHIP.Picture, x2, y2
Case "Right"
x2 = x2 + 210
If x2 > Form1.Width - 1700 Then x2 = Form1.Width - 1700
PaintPicture sprSHIP.Picture, x2, y2
End Select
'Move the star field that's on the background.
For i = 1 To 30
Circle (x(i), Y(i)), size(i), BackColor
Y(i) = Y(i) + pace(i)
'If a star reaches the bottom of the form then
'it is assigned a new X coordinate and in the next
'timer event it will start falling again ( Y(i)=0)
If Y(i) >= Form1.Height Then Y(i) = 0: x(i) = Int(Form1.Width * Rnd)
Select Case pace(i)
'There are 30 circles of various sizes.Each one
'moves with a different speed and according to his
'speed it is painted with a different shade of grey.
'Stars that move slow are painted almost black 'cause
'they are located deep in space.The ones that move
'quick are almost white because they zip by near
'the camera,they're also bigger than then slow ones.
Case Is <= 200
clr = &H404040
Case Is <= 300
clr = &H808080
Case Is <= 400
clr = &HC0C0C0
Case Else
clr = &HFFFFFF
End Select
Circle (x(i), Y(i)), size(i), clr
Next
'Paint the good guy
PaintPicture sprSHIP.Picture, x2, y2
'Paint the aliens that are not dead,in other
'words the ones for whom Kill(#of alien)=false
'Remember 1<= #of alien <=15
If kill(1) = False Then PaintPicture sprINVADER5.Picture, xi1, yi1
For i = 1 To 2
If kill(1 + i) = False Then PaintPicture sprINVADER5.Picture, xi2(i), yi2(i)
Next
For i = 1 To 3
If kill(3 + i) = False Then PaintPicture sprINVADER5.Picture, xi3(i), yi3(i)
Next
For i = 1 To 4
If kill(6 + i) = False Then PaintPicture sprINVADER5.Picture, xi4(i), yi4(i)
Next
For i = 1 To 5
If kill(10 + i) = False Then PaintPicture sprINVADER5.Picture, xi5(i), yi5(i)
Next
End Sub
'Pull that trigger...
Private Sub fire()
Select Case boom
Case True
Exit Sub
'...but if an explosion is going on,don't fire,
'there's no need to waste ammo
End Select
upper = 1000
Select Case shot
Case Is <> 0
'Also if the previous shot hasn't yet been painted
'in all of it's 3 colors,don't fire onother one
Exit Sub
End Select
shot = 1
'Play the sound fx using API's
xshot = x2 + 1062
'Xshot is the X coordinate of the laser beem.The
'weapon of the ship is in the middle,so it's
'Xshot = x2 + 1062 (x2 is the X coordinate of the ship)
'Keep in mind that the sprite has a lot of space
'in each side of the actual ship otherwise it would
'leave trails on the form when it moves
'
'Check if the Xshot is anywhere near an alien
'and "kill" him if so ( kill(#of alien)=True).He
'may not yet be officialy dead since the
'explosion hasn't been painted on him,but that's
'a matter of milliseconds if his KILL array element
'is set to True.
If Abs(xi1 + (sprINVADER5.Width / 2) - xshot) < (sprINVADER5.Width / 4) Then kill(1) = True Else kill(1) = False
If Abs(xi2(1) + (sprINVADER5.Width / 2) - xshot) < (sprINVADER5.Width / 3) Then kill(2) = True Else kill(2) = False
If Abs(xi2(2) + (sprINVADER5.Width / 2) - xshot) < (sprINVADER5.Width / 3) Then kill(3) = True Else kill(3) = False
For i = 1 To 3
If Abs(xi3(i) + (sprINVADER5.Width / 2) - xshot) < (sprINVADER5.Width / 2) Then kill(3 + i) = True: Exit For Else kill(3 + i) = False
Next
For i = 1 To 4
If Abs(xi4(i) + (sprINVADER5.Width / 2) - xshot) < (sprINVADER5.Width / 2) Then kill(6 + i) = True: Exit For Else kill(6 + i) = False
Next
For i = 1 To 5
If Abs(xi5(i) + (sprINVADER5.Width / 2) - xshot) < (sprINVADER5.Width / 3) Then kill(10 + i) = True: Exit For Else kill(10 + i) = False
Next
If kill(1) = True Then xinv = xi1: yinv = yi1: upper = yi1: Call kaboom
If kill(2) = True Then xinv = xi2(1): yinv = yi2(1): upper = yi2(1): Call kaboom:
If kill(3) = True Then xinv = xi2(2): yinv = yi2(2): upper = yi2(2): Call kaboom:
For i = 1 To 3
If kill(i + 3) = True Then xinv = xi3(i): yinv = yi3(i): upper = yi3(i): Call kaboom:
Next
For i = 1 To 4
If kill(6 + i) = True Then xinv = xi4(i): yinv = yi4(i): upper = yi4(i): Call kaboom:
Next
For i = 1 To 5
If kill(10 + i) = True Then xinv = xi5(i): yinv = yi5(i): upper = yi5(i): Call kaboom:
Next
aa = 0
'Check if the shot "killed" some more aliens
'that were above the unlucky one and "ressurect"
'them since laser blasters don't go through
'metal.Only railguns in Quake2...
For i = 15 To 1 Step -1
If kill(i) = True Then aa = i: Exit For
Next
For i = 1 To 15
If i <> aa Then kill(i) = False
Next
11 Line (x2 + 1062, y2)-(xshot, upper), BackColor
'Randomly choose a color for the first of three
'laser beems ...
D = Int(2 * Rnd)
Select Case D
Case 0
COL1 = &HFFFFFF
COL2 = &HFF0000
Case Else
COL1 = &HFF0000
COL2 = &HFFFFFF
End Select
'and paint it...If this one was blue,the second will
'be white and vise versa.The third is always the
'black one that "erases" the trail of the shot
Line (x2 + 1062, y2)-(x2 + 1062, upper), COL1
End Sub
'The procedure that runs everytime you shoot an alien.
'He still has no idea what's in store for him but
'these lines of code will fix him up good
Private Sub kaboom()
'Place the blank label on the unlucky alien so that
'he's erased on the next timer event,before the
'frames of the explosion start showing
lblBlank.Top = yinv
lblBlank.Left = xinv
boom = True
'...and start the fun by setting the Explosion variable
'to 1.In the next timer event a beautiful explosion
'will go off on an ugly alien's face
explosion = 1
End Sub
```

