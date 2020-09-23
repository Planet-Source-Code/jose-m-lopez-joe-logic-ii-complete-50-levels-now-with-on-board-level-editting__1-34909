Attribute VB_Name = "Processing"
Option Explicit
'*************************************************************************************************************************************************************************************************************************
     'Sleep Declaration
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Sleep x' x = milliseconds
'*************************************************************************************************************************************************************************************************************************

'*************************************************************************************************************************************************************************************************************************
'Midi Background music
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public IsMusicOn As Boolean
Public RetValue As Long
'*************************************************************************************************************************************************************************************************************************


Public intResponse As Integer 'Message boxes
Public picCell(11) As Picture 'Picture array
Public picCellSmall(11) As Picture 'Picture array

Public strMasterArray() As String 'Holds entire game
Public intMasterFileSize As Integer 'Holds game size
Public CellContents(20, 15) As Integer 'Screen cell contents array
Public CellCharactr(20, 15) As String 'Screen cell contents array
Public StoredCellContents(10, 20, 15) As Integer 'Small Storage Screen cell contents array
Public StoredCellCharactr(10, 20, 15) As String 'Small Storage Screen cell contents array
Public xCorner(20, 15) As Integer 'Stores Cell corner position
Public yCorner(20, 15) As Integer 'Stores Cell corner position

Public arrMoves(10000) As String  'Stores Moves
Public intMoveCount As Integer 'Holds Move Count
Public intLevel As Integer 'Holds current level number
Public intMovesMaxLineSize As Integer
Public intMovesMaxLineQnty As Integer
Public intMaxLevelQnty As Integer

Dim intLevelStartLine As Integer
Public JoeX As Integer  'Store Joe's x cell position
Public JoeY As Integer 'Store Joe's y cell position
Public TempColumn As Integer  'Store Temporary x cell position for falling objects
Public TempRow As Integer 'Store Temporary y cell position for falling objects
Public TempObject As Integer  'Holds Temporarily a falling object's value
Public TempCharacter As String   'Holds Temporarily a falling object's Letter

Public EditObject As Integer  'Holds an object's value in Edit Mode
Public EditX As Integer  'Holds an object's value in Edit Mode
Public EditY As Integer  'Holds an object's value in Edit Mode
Public stFileName As String 'Hold current open filename
Public sngSolutionDelay As Double 'Solutions delay value according to solution menu

Public booWithBlock As Boolean 'Whether or not little Joe is carrying a block
Public booLevelAvailable As Boolean
'This booBusy is needed because without it when Joe is falling down
'an appreciable distance, he can also travel right or left if the respective key
'is held down. This variable prevents this.
Public booBusy As Boolean
Public booClunk As Boolean

'*********************************************************
'********Picture Name Constants***************************
'*********************************************************
Public Const cEmpty As Integer = 1
Public Const cBox As Integer = 2
Public Const cRoll  As Integer = 3
Public Const cRail  As Integer = 4
Public Const cBrick  As Integer = 5
Public Const cTransport1  As Integer = 6
Public Const cTransport2  As Integer = 7
Public Const cJoeLeft  As Integer = 8
Public Const cJoeRight  As Integer = 9
Public Const cExit  As Integer = 10
Public Const cRailWalkOnce  As Integer = 11
'*********************************************************
'*********************************************************
'*********************************************************
    
    '*********************************************************
    '**********Special Location Values************************
    '*********************************************************
    Public Transport1X   As Integer
    Public Transport2X   As Integer
    Public Transport1Y   As Integer
    Public Transport2Y   As Integer
    Public ExitColumn As Integer
    Public ExitRow As Integer
    '*********************************************************
    '*********************************************************
    '*********************************************************

'*********************************************************
'*********************************************************
'*********************************************************


'*********************************************************
'***Sound Effects ****************************************
'*********************************************************
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Global Const SND_SYNC = &H0
    Global Const SND_ASYNC = &H1
'Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
'Public Const SND_ASYNC = &H1
'Public Const SND_LOOP = &H8
'Public Const SND_NODEFAULT = &H2
'Public Const SND_SYNC = &H0
'Public Const SND_NOSTOP = &H10
'Public Const SND_MEMORY = &H4
'Public RetValue As Long
'*********************************************************
'*********************************************************

'*********************************************************
Public dummy As Integer 'For calling Fx function
Public dblPreviousSoundEffectTime
Public intPreviousSoundEffect
'*********************************************************

'*********************************************************
'***Sound Effects Constants*******************************
'*********************************************************
Public Const cWalk As Integer = 1
Public Const cClimb As Integer = 2
Public Const cWalkWithBox As Integer = 3
Public Const cClimbWithBox As Integer = 4
Public Const cJoeTurn As Integer = 5
Public Const cJoeFall As Integer = 6
Public Const cJoeLand As Integer = 7

Public Const cBallRoll As Integer = 8
Public Const cBallFall As Integer = 9
Public Const cBallLand As Integer = 10

Public Const cBoxUp As Integer = 11
Public Const cBoxDown As Integer = 12
Public Const cBoxSlide As Integer = 13
Public Const cBoxFall As Integer = 14
Public Const cBoxLand As Integer = 15

Public Const cGrayRail As Integer = 16
Public Const cTransport As Integer = 17
Public Const cExiting As Integer = 18
Public Const cBallStop As Integer = 19
Public Const cClunk As Integer = 20
'*********************************************************



Public Function fncPlayFx(FxNumber As Integer) As Integer
'Sound Effects Functtion
'All Sound Effects go through here
Dim x As Integer, y As Integer, intFallLength As Integer
If Form1.mnuSoundEffects.Checked = False Then GoTo lblExit

Select Case FxNumber
    Case cWalk
        If Timer > (dblPreviousSoundEffectTime + 0.1) Then
        Call sndPlaySound("_JoeWalk.wav", SND_ASYNC)
        Else
        GoTo lblExit:
        End If
    Case cClimb: Call sndPlaySound("_JoeWalk.wav", SND_ASYNC)
    Case cWalkWithBox: Call sndPlaySound("_JoeWalkWithBox.wav", SND_ASYNC)
    Case cClimbWithBox: Call sndPlaySound("_JoeClimbWithBox.wav", SND_ASYNC)
    Case cJoeTurn: Call sndPlaySound("_JoeTurn.wav", SND_ASYNC)
    Case cJoeFall
        If intPreviousSoundEffect = cJoeFall Then GoTo lblExit 'Already falling
        'Get fall length
        x = TempColumn: y = TempRow + 1
        intFallLength = 0
        Do While CellContents(x, y) = cEmpty
        y = y + 1: intFallLength = intFallLength + 1
        Loop
            'Select Fx according to fall length
            Select Case intFallLength
            Case 0, 1, 2
            Case 3, 4, 5: Call sndPlaySound("_JoeFall1.wav", SND_ASYNC)
            Case 6 To 10: Call sndPlaySound("_JoeFall2.wav", SND_ASYNC)
            Case 10 To 18: Call sndPlaySound("_JoeFall3.wav", SND_ASYNC)
            End Select
    Case cBallRoll: Call sndPlaySound("_BallRoll.wav", SND_ASYNC)
    Case cBallStop: Call sndPlaySound("_BallStop.wav", SND_ASYNC)
    Case cBoxUp: Call sndPlaySound("_BoxUp.wav", SND_ASYNC)
    Case cBoxDown: Call sndPlaySound("_BoxDown.wav", SND_ASYNC)
    Case cBoxSlide: Call sndPlaySound("_BoxSlide.wav", SND_ASYNC)
    Case cBoxFall
        If intPreviousSoundEffect = cBoxFall Then GoTo lblExit 'Already falling
        'Get fall length
        x = TempColumn: y = TempRow + 1
        intFallLength = 0
        Do While CellContents(x, y) = cEmpty
        y = y + 1: intFallLength = intFallLength + 1
        Loop
            'Select Fx according to fall length
            Select Case intFallLength
            Case 0, 1, 2
            Case 3, 4, 5: Call sndPlaySound("_BoxFall1.wav", SND_ASYNC)
            Case 6 To 10: Call sndPlaySound("_BoxFall2.wav", SND_ASYNC): booClunk = True
            Case 10 To 18: Call sndPlaySound("_BoxFall3.wav", SND_ASYNC): booClunk = True
            End Select
    Case cTransport: Call sndPlaySound("_Transport.wav", SND_ASYNC)
    Case cGrayRail: Call sndPlaySound("_GrayRail.wav", SND_ASYNC)
    Case cClunk: Call sndPlaySound("_Clunk.wav", SND_ASYNC)
    Case cExiting: Call sndPlaySound("_Exit.wav", SND_ASYNC)
End Select


dblPreviousSoundEffectTime = Timer
intPreviousSoundEffect = FxNumber
lblExit:
End Function

Public Sub rtnLoadMasterArray()
Dim tmpText As String, tmpCount As Integer
Dim x As Integer
'Enter:
'stFilename holding file to open and load from
'Leave:
'strMasterArray() holding entire game
'intMasterFileSize holding strMasterArray() size

'Get filesize
Open stFileName For Input As #1
    tmpCount = 0
    Do While Not EOF(1)
    Line Input #1, tmpText
    tmpCount = tmpCount + 1
    Loop
Close
intMasterFileSize = tmpCount
ReDim strMasterArray(intMasterFileSize)
'Load master array
Open stFileName For Input As #1
    For x = 1 To intMasterFileSize
    Line Input #1, tmpText
    strMasterArray(x) = tmpText
    Next x
Close
End Sub


Public Sub rtnInitialize()
'Enter:
'intLevel holding requested level number to display
'strMasterArray() holding entire game
'intMasterFileSize holding strMasterArray() size

'Get the text level file and decode it into the CellContents array
Form1.Enabled = False 'Prevent unwanted events while busy
Form1.Caption = "Joe Logic           Level " & intLevel & "  " & stFileName 'Title

'*****************************************************
'***Determine if next level is available**************
'*****************************************************
rtnCheckIfLevelAvailable
'intLevelStartLine now holds level's first line needed in rtnLoadLevelCellArrayslFromMasterArray
    'If not available then End of Game!
    If booLevelAvailable = False Then
    Form1.mnuFile.Enabled = False
    Form1.mnuEdit.Enabled = False
    Form1.mnuLevel.Enabled = False
    Form1.mnuSolution.Enabled = False
    Form1.mnuHelp.Enabled = False
    Form1.tmrFinale.Enabled = True
    Exit Sub
    End If
'*****************************************************
'*****************************************************
'*****************************************************

rtnLoadLevelCellArrayslFromMasterArray
Form1.Enabled = False 'Prevent unwanted events while busy

rtnRefreshFromCellArrays
Form1.Enabled = False 'Prevent unwanted events while busy

If Form1.mnuEditEditMode.Checked = True Then GoTo lblOnEditMode
rtnAllRollsAndBoxesFall
Form1.Enabled = False 'Prevent unwanted events while busy
lblOnEditMode:
    
    'Reset solution/moves array
    For intMoveCount = 1 To 10000
    arrMoves(intMoveCount) = "0"
    Next intMoveCount
intMoveCount = 0 'Reset Count
booWithBlock = False
Form1.Enabled = True
Form1.Picture1.SetFocus
End Sub

Public Sub rtnCheckIfLevelAvailable()
Dim tmpText As String, x As Integer
booLevelAvailable = False
For x = 1 To intMasterFileSize
tmpText = strMasterArray(x)
    If tmpText = "level" & Trim(Str(intLevel)) Then
    booLevelAvailable = True
    intLevelStartLine = x + 1
    Exit For
    End If
Next x
End Sub

Public Sub rtnLoadLevelCellArrayslFromMasterArray()
Dim tmpText As String, stChar As String
Dim intLine As Integer, intChar As Integer
Dim x As Integer
    
        
        
'*****************************************************
'***********Read Level Line By Line*******************
'*****************************************************
        intLine = 0
        For x = intLevelStartLine To intLevelStartLine + 14     'y cells
        tmpText = strMasterArray(x)
        intLine = intLine + 1
            For intChar = 1 To 20   'x cells
            stChar = Mid$(tmpText, intChar, 1)
                Select Case stChar
                Case " "
                CellContents(intChar, intLine) = cEmpty
                CellCharactr(intChar, intLine) = stChar
                Case "B"
                CellContents(intChar, intLine) = cBox
                CellCharactr(intChar, intLine) = stChar
                Case "O"
                CellContents(intChar, intLine) = cRoll
                CellCharactr(intChar, intLine) = stChar
                Case "="
                CellContents(intChar, intLine) = cRail
                CellCharactr(intChar, intLine) = stChar
                Case "#"
                CellContents(intChar, intLine) = cBrick
                CellCharactr(intChar, intLine) = stChar
                Case "T"
                CellContents(intChar, intLine) = cEmpty
                CellCharactr(intChar, intLine) = stChar
                Case "t"
                CellContents(intChar, intLine) = cEmpty
                CellCharactr(intChar, intLine) = stChar
                Case "j"
                CellContents(intChar, intLine) = cJoeLeft
                CellCharactr(intChar, intLine) = stChar
                Case "J"
                CellContents(intChar, intLine) = cJoeRight
                CellCharactr(intChar, intLine) = stChar
                Case "X"
                CellContents(intChar, intLine) = cEmpty
                CellCharactr(intChar, intLine) = stChar
                Case "."
                CellContents(intChar, intLine) = cRailWalkOnce
                CellCharactr(intChar, intLine) = stChar
                End Select
            Next intChar
        Next x
booWithBlock = False
'*****************************************************
'*****************************************************
'*****************************************************
'*****************************************************

'       empty = " " 1
'         box = "B" 2
'        roll = "O" 3
'        rail = "=" 4
'       brick = "#" 5
'   transport1 = "T" 6
'   transport2 = "t" 7
'     joeleft = "j" 8
'    joeright = "J" 9
'        exit = "X" 10
'railwalkonce = "." 11
End Sub

Public Sub rtnRefreshFromCellArrays()
Dim x As Integer, y As Integer
Dim intLine As Integer, intChar As Integer

Form1.Enabled = False 'Prevent unwanted events while busy
    'Turn off transports
    Form1.imgTransport(1).Visible = False
    Form1.imgTransport(2).Visible = False
    Form1.imgExit.Visible = False

    'Impossible values indicating these objects are currently non existent
    Transport1X = 100
    Transport2X = 100
    Transport1Y = 100
    Transport2Y = 100
    ExitColumn = 100
    ExitRow = 100

'Paint screen from arrays
'Blocks x = 0 to 19    y = 0 to 14    20x15 Blocks   Total 300 Blocks
For y = 1 To 15
For x = 1 To 20
Form1.Picture1.PaintPicture picCell(CellContents(x, y)), xCorner(x, y), yCorner(x, y)
Next x
Next y

For intLine = 1 To 15
For intChar = 1 To 20
'Debug.Print CellCharactr(intChar, intLine);
    Select Case CellCharactr(intChar, intLine)
    Case "T"
                Transport1X = intChar
                Transport1Y = intLine
                Form1.imgTransport(1).Visible = True
                Form1.imgTransport(1).Left = xCorner(intChar, intLine)
                Form1.imgTransport(1).Top = yCorner(intChar, intLine)
    Case "t"
                Transport2X = intChar
                Transport2Y = intLine
                Form1.imgTransport(2).Visible = True
                Form1.imgTransport(2).Left = xCorner(intChar, intLine)
                Form1.imgTransport(2).Top = yCorner(intChar, intLine)
    Case "j"
                JoeX = intChar
                JoeY = intLine
    Case "J"
                JoeX = intChar
                JoeY = intLine
    Case "X"
                ExitColumn = intChar
                ExitRow = intLine
                Form1.imgExit.Visible = True
                Form1.imgExit.Left = xCorner(intChar, intLine)
                Form1.imgExit.Top = yCorner(intChar, intLine)
    End Select
Next intChar
'Debug.Print
Next intLine

Form1.Enabled = True
Form1.Picture1.Refresh
Form1.Picture1.SetFocus
End Sub

'Enters WITHOUT box
Public Sub rtnRightKey()
With Form1
'***********************************************************************
'if bumping a roll
If CellContents(JoeX + 1, JoeY) = cRoll And CellContents(JoeX, JoeY) = cJoeRight Then
TempObject = cRoll
TempColumn = JoeX + 1
TempRow = JoeY
rtnRollRight
Exit Sub
End If
'***********************************************************************

   
'***********************************************************************
'If Joe was facing left, Just turn Joe Right and Exit
If CellContents(JoeX, JoeY) = cJoeLeft Then
'Paint Joe Right
.Picture1.PaintPicture picCell(cJoeRight), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
dummy = fncPlayFx(cJoeTurn)
    'Update arrays
    CellContents(JoeX, JoeY) = cJoeRight
    If CellCharactr(JoeX, JoeY) <> "T" And CellCharactr(JoeX, JoeY) <> "t" Then
    CellCharactr(JoeX, JoeY) = "J"
    End If
Exit Sub
End If
'***********************************************************************
   
'If right block empty then move right
If CellContents(JoeX + 1, JoeY) = cEmpty Then
'Paint JoeRight where he will now be
.Picture1.PaintPicture picCell(cJoeRight), xCorner(JoeX + 1, JoeY), yCorner(JoeX + 1, JoeY)
'Paint empty where he was
.Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
dummy = fncPlayFx(cWalk)
    CellContents(JoeX, JoeY) = cEmpty 'Update Arrays
    If CellCharactr(JoeX, JoeY) <> "T" And CellCharactr(JoeX, JoeY) <> "t" Then
    CellCharactr(JoeX, JoeY) = " " 'Update Arrays
    End If

    'Joes new row position
    JoeX = JoeX + 1
    CellContents(JoeX, JoeY) = cJoeRight 'Update Arrays
    If CellCharactr(JoeX, JoeY) <> "T" And CellCharactr(JoeX, JoeY) <> "t" Then
    CellCharactr(JoeX, JoeY) = "J" 'Update Arrays
    End If

       
   
   'If previous block was a RailWalkOnce then
    If CellContents(JoeX - 1, JoeY + 1) = cRailWalkOnce Then
   'Make it disappear
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY + 1), yCorner(JoeX - 1, JoeY + 1)
    CellContents(JoeX - 1, JoeY + 1) = cEmpty 'Update Arrays
    CellCharactr(JoeX - 1, JoeY + 1) = " " 'Update Arrays
    dummy = fncPlayFx(cGrayRail)
    End If
            
    'Check if Transport and process
    If Transport1X = JoeX And Transport1Y = JoeY Then
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellContents(Transport2X, Transport2Y) = cJoeRight 'Update
    JoeX = Transport2X
    JoeY = Transport2Y
    dummy = fncPlayFx(cTransport)
    rtnRefreshFromCellArrays
    Exit Sub
    Else
        If Transport2X = JoeX And Transport2Y = JoeY Then
        CellContents(JoeX, JoeY) = cEmpty 'Update
        CellContents(Transport1X, Transport1Y) = cJoeRight 'Update
        JoeX = Transport1X
        JoeY = Transport1Y
        dummy = fncPlayFx(cTransport)
        rtnRefreshFromCellArrays
        Exit Sub
        End If
    End If

    'Check if Exit
    If ExitColumn = JoeX And ExitRow = JoeY Then
    intLevel = intLevel + 1
    If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
dummy = fncPlayFx(cExiting)
    rtnInitialize
    Exit Sub
    End If
       
    'Check if should fall
    TempObject = cJoeRight
    TempColumn = JoeX
    TempRow = JoeY
    rtnCheckIfFallAndFall
    
    'Check if Exit
    If ExitColumn = JoeX And ExitRow = JoeY Then
    intLevel = intLevel + 1
    If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
    dummy = fncPlayFx(cExiting)
    rtnInitialize
    Exit Sub
    End If
    
End If

End With
End Sub

Public Sub rtnLeftKey()
'Enters WITHOUT box
With Form1
'***********************************************************************
'if bumping a roll
If CellContents(JoeX - 1, JoeY) = cRoll And CellContents(JoeX, JoeY) = cJoeLeft Then
TempObject = cRoll
TempColumn = JoeX - 1
TempRow = JoeY
rtnRollLeft
Exit Sub
End If
'***********************************************************************
   
'***********************************************************************
'If Joe was facing right Just turn Joe Left and Exit
 If CellContents(JoeX, JoeY) = cJoeRight Then
'Paint Joe Left
.Picture1.PaintPicture picCell(cJoeLeft), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
 dummy = fncPlayFx(cJoeTurn)
    CellContents(JoeX, JoeY) = cJoeLeft 'Update
    If CellCharactr(JoeX, JoeY) <> "T" And CellCharactr(JoeX, JoeY) <> "t" Then
    CellCharactr(JoeX, JoeY) = "j" 'Update Arrays
    End If
Exit Sub
End If
'***********************************************************************
 

'***********************************************************************
'If left block empty then move left
 If CellContents(JoeX - 1, JoeY) = cEmpty Then
'Paint  This Picture  atPixelxof  Column,  RowPixelyofColumn,  Row
.Picture1.PaintPicture picCell(cJoeLeft), xCorner(JoeX - 1, JoeY), yCorner(JoeX - 1, JoeY)
.Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
dummy = fncPlayFx(cWalk)
 
    CellContents(JoeX, JoeY) = cEmpty 'Update
    If CellCharactr(JoeX, JoeY) <> "T" And CellCharactr(JoeX, JoeY) <> "t" Then
    CellCharactr(JoeX, JoeY) = " " 'Update Arrays
    End If
    JoeX = JoeX - 1
    CellContents(JoeX, JoeY) = cJoeLeft 'Update
    If CellCharactr(JoeX, JoeY) <> "T" And CellCharactr(JoeX, JoeY) <> "t" Then
    CellCharactr(JoeX, JoeY) = "j" 'Update Arrays
    End If

   'If block was a RailWalkOnce
    If CellContents(JoeX + 1, JoeY + 1) = cRailWalkOnce Then
   'Make it disappear
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY + 1), yCorner(JoeX + 1, JoeY + 1)
    CellContents(JoeX + 1, JoeY + 1) = cEmpty
    CellCharactr(JoeX + 1, JoeY + 1) = " "
    dummy = fncPlayFx(cGrayRail)
    End If
           
    'Check if Transport and process
    If Transport1X = JoeX And Transport1Y = JoeY Then
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellContents(Transport2X, Transport2Y) = cJoeLeft 'Update
    JoeX = Transport2X
    JoeY = Transport2Y
    dummy = fncPlayFx(cTransport)
    rtnRefreshFromCellArrays
    Exit Sub
    Else
        If Transport2X = JoeX And Transport2Y = JoeY Then
        CellContents(JoeX, JoeY) = cEmpty 'Update
        CellContents(Transport1X, Transport1Y) = cJoeLeft 'Update
        JoeX = Transport1X
        JoeY = Transport1Y
        dummy = fncPlayFx(cTransport)
        rtnRefreshFromCellArrays
        Exit Sub
        End If
    End If

    'Check if Exit and process
    If ExitColumn = JoeX And ExitRow = JoeY Then
    intLevel = intLevel + 1
    If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
    dummy = fncPlayFx(cExiting)
    rtnInitialize
    Exit Sub
    End If
       
    'Check if should fall
    TempObject = cJoeLeft
    TempColumn = JoeX
    TempRow = JoeY
    rtnCheckIfFallAndFall
    
    'Check if Exit and process
    If ExitColumn = JoeX And ExitRow = JoeY Then
    intLevel = intLevel + 1
    If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
    dummy = fncPlayFx(cExiting)
    rtnInitialize
    Exit Sub
    End If

End If

End With
End Sub

Public Sub rtnRightKeyWithBox()
'Enters WITH box
With Form1
   
'***********************************************************************
'If Joe was facing left
'Just turn Joe Right and Exit
If CellContents(JoeX, JoeY) = cJoeLeft Then
'       Paint           This Picture at Pxlx   of     Column,  Row  Pixely  of   Column,  Row
.Picture1.PaintPicture picCell(cJoeRight), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
CellContents(JoeX, JoeY) = cJoeRight 'Update
CellCharactr(JoeX, JoeY) = "J"
dummy = fncPlayFx(cJoeTurn)
Exit Sub
End If
'***********************************************************************
   
   
   
'***********************************************************************
'Non obstructed right walk
'If right block empty and right top block is empty
If CellContents(JoeX + 1, JoeY) = cEmpty And CellContents(JoeX + 1, JoeY - 1) = cEmpty Then
'Move Joe to the left once
JoeX = JoeX + 1
'Paint JoeRight at Adjacent Right cell and Clear Current cell
.Picture1.PaintPicture picCell(cJoeRight), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
.Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY), yCorner(JoeX - 1, JoeY)
'Paint Box at Adjacent Top Right cell and Clear Current Top cell
.Picture1.PaintPicture picCell(cBox), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
.Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY - 1), yCorner(JoeX - 1, JoeY - 1)
'Update cell where Joe will be
CellContents(JoeX, JoeY) = cJoeRight
CellCharactr(JoeX, JoeY) = "J"
'Update cell where Joe was to empty
CellContents(JoeX - 1, JoeY) = cEmpty
CellCharactr(JoeX - 1, JoeY) = " "
'Update cell where Joe's box will be
CellContents(JoeX, JoeY - 1) = cBox
CellCharactr(JoeX, JoeY - 1) = "B"
'Update cell where Joe's box was
CellContents(JoeX - 1, JoeY - 1) = cEmpty
CellCharactr(JoeX - 1, JoeY - 1) = " "
dummy = fncPlayFx(cWalkWithBox)
    '***********************************************************************
    'If block was a RailWalkOnce
    If CellContents(JoeX - 1, JoeY + 1) = cRailWalkOnce Then
    'Make it disappear
    .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY + 1), yCorner(JoeX - 1, JoeY + 1)
    CellContents(JoeX - 1, JoeY + 1) = cEmpty
    CellCharactr(JoeX - 1, JoeY + 1) = " "
    dummy = fncPlayFx(cGrayRail)
    End If
   '***********************************************************************
       
    '***********************************************************************
    'Check if Exit and process
    If ExitColumn = JoeX And ExitRow = JoeY Then
    intLevel = intLevel + 1
    If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
    dummy = fncPlayFx(cExiting)
    rtnInitialize
    Exit Sub
    End If
    '***********************************************************************
       
       
       '***********************************************************************
       'If lower cell is empty then fall
        If CellContents(JoeX, JoeY + 1) = cEmpty Then
       'Check if should fall
        Dim tempBlockColumn As Integer, tempBlockRow As Integer
        TempObject = cJoeRight
        TempColumn = JoeX
        TempRow = JoeY
            tempBlockColumn = JoeX
            tempBlockRow = JoeY - 1
            rtnCheckIfFallAndFall
        TempObject = cBox
        TempColumn = tempBlockColumn
        TempRow = tempBlockRow
        rtnCheckIfFallAndFall
       '***********************************************************************
       
       '********************************************
       'Check if Exit and process
        If ExitColumn = JoeX And ExitRow = JoeY Then
        intLevel = intLevel + 1
        If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
        dummy = fncPlayFx(cExiting)
        rtnInitialize
        Exit Sub
        End If
        End If
       '********************************************
    
    Exit Sub
End If
    '***********************************************************************
   
'***********************************************************************
'Top obstructed right walk
'If right block empty and right top block is not empty
If CellContents(JoeX + 1, JoeY) = cEmpty And CellContents(JoeX + 1, JoeY - 1) <> cEmpty Then
'Paint JoeRight at Adjacent Right cell and Clear Current cell
.Picture1.PaintPicture picCell(cJoeRight), xCorner(JoeX + 1, JoeY), yCorner(JoeX + 1, JoeY)
.Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)

CellContents(JoeX + 1, JoeY) = cJoeRight 'Update
If CellCharactr(JoeX + 1, JoeY) <> "T" And CellCharactr(JoeX + 1, JoeY) <> "t" Then
CellCharactr(JoeX + 1, JoeY) = "J"
End If

CellContents(JoeX, JoeY) = cEmpty  'Update
CellCharactr(JoeX, JoeY) = " "
JoeX = JoeX + 1
dummy = fncPlayFx(cWalkWithBox)
    '***********************************************************************
    'If block was a RailWalkOnce
    If CellContents(JoeX - 1, JoeY + 1) = cRailWalkOnce Then
    'Make it disappear
    .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY + 1), yCorner(JoeX - 1, JoeY + 1)
    CellContents(JoeX - 1, JoeY + 1) = cEmpty
    CellCharactr(JoeX - 1, JoeY + 1) = " "
    dummy = fncPlayFx(cGrayRail)
    End If
   '***********************************************************************
'Box fall
TempObject = cBox
TempColumn = JoeX - 1
TempRow = JoeY - 1
rtnCheckIfFallAndFall
booWithBlock = False
'***********************************************************************
       
         '***************************************************************
         'Check if Transport and process
         If Transport1X = JoeX And Transport1Y = JoeY Then
         CellContents(JoeX, JoeY) = cEmpty 'Update
         CellContents(Transport2X, Transport2Y) = cJoeRight 'Update
         JoeX = Transport2X
         JoeY = Transport2Y
         dummy = fncPlayFx(cTransport)
         rtnRefreshFromCellArrays
         Else
            If Transport2X = JoeX And Transport2Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport1X, Transport1Y) = cJoeRight 'Update
            JoeX = Transport1X
            JoeY = Transport1Y
            dummy = fncPlayFx(cTransport)
            rtnRefreshFromCellArrays
            End If
        End If
       '***************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
            dummy = fncPlayFx(cExiting)
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
           'Joe fall if nothing underneath
            If CellContents(JoeX, JoeY + 1) = cEmpty Then
            TempObject = cJoeRight
            TempColumn = JoeX
            TempRow = JoeY
            rtnCheckIfFallAndFall
            End If
       '***********************************************************************
       '***********************************************************************
    
    Exit Sub
    End If
   '***********************************************************************
   '***********************************************************************
   '***********************************************************************


End With
End Sub

Public Sub rtnLeftKeyWithBox()
'Enters WITH box
With Form1
   '***********************************************************************
   'Just turn joe
   'if left block not empty or joe facing right
   'then just turn Joe if not already facing left
    If CellContents(JoeX - 1, JoeY) <> cEmpty Or CellContents(JoeX, JoeY) = cJoeRight Then
   '       Paint           This Picture at Pxlx   of     Column,  Row  Pixely  of   Column,  Row
   .Picture1.PaintPicture picCell(cJoeLeft), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cJoeLeft 'Update
    CellCharactr(JoeX, JoeY) = "j"
    dummy = fncPlayFx(cJoeTurn)
    Exit Sub
    End If
   '***********************************************************************
    
   '***********************************************************************
   '***********************************************************************3
   '***********************************************************************
   'Non obstructed left walk
   'If left block empty then move left            and             left top block is empty
    If CellContents(JoeX - 1, JoeY) = cEmpty And CellContents(JoeX - 1, JoeY - 1) = cEmpty Then
    'Move Joe to the left once
    JoeX = JoeX - 1
   'Paint Joeleft at Adjacent left cell and Clear Current cell
   .Picture1.PaintPicture picCell(cJoeLeft), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY), yCorner(JoeX + 1, JoeY)
   'Paint Box at Adjacent Top left cell and Clear Current Top cell
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY - 1), yCorner(JoeX + 1, JoeY - 1)
    CellContents(JoeX, JoeY) = cJoeLeft  'Update
    CellCharactr(JoeX, JoeY) = "j"
    'Update Adjacent left cell
    CellContents(JoeX + 1, JoeY) = cEmpty
    CellCharactr(JoeX + 1, JoeY) = " "
    
    CellContents(JoeX, JoeY - 1) = cBox  'Update
    CellCharactr(JoeX, JoeY - 1) = "B"
    
    CellContents(JoeX + 1, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = " "
    dummy = fncPlayFx(cWalkWithBox)
       
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX + 1, JoeY + 1) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY + 1), yCorner(JoeX + 1, JoeY + 1)
             CellContents(JoeX + 1, JoeY + 1) = cEmpty
             CellCharactr(JoeX + 1, JoeY + 1) = " "
        dummy = fncPlayFx(cGrayRail)
        End If
       '***********************************************************************
       '***********************************************************************
       
           '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
            dummy = fncPlayFx(cExiting)
            rtnInitialize
            Exit Sub
            End If
           '***********************************************************************
       
       
       '***********************************************************************
       'If lower cell is empty then fall
        If CellContents(JoeX, JoeY + 1) = cEmpty Then
       'Check if should fall
        Dim tempBlockColumn As Integer, tempBlockRow As Integer
        TempObject = cJoeLeft
        TempColumn = JoeX
        TempRow = JoeY
            tempBlockColumn = JoeX
            tempBlockRow = JoeY - 1
            rtnCheckIfFallAndFall
        TempObject = cBox
        TempColumn = tempBlockColumn
        TempRow = tempBlockRow
        rtnCheckIfFallAndFall
       '***********************************************************************
            
           '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
            dummy = fncPlayFx(cExiting)
            rtnInitialize
            Exit Sub
            End If
           '***********************************************************************
        End If
    Exit Sub
    End If
   '***********************************************************************
   '***********************************************************************
   '***********************************************************************


   '***********************************************************************
   '***********************************************************************
   '***********************************************************************
   'Top obstructed left walk
   'If left block empty                            and         left top block is not empty
    If CellContents(JoeX - 1, JoeY) = cEmpty And CellContents(JoeX - 1, JoeY - 1) <> cEmpty Then
       'Paint Joeleft at Adjacent left cell and Clear Current cell
       .Picture1.PaintPicture picCell(cJoeLeft), xCorner(JoeX - 1, JoeY), yCorner(JoeX - 1, JoeY)
       .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
    

        CellContents(JoeX - 1, JoeY) = cJoeLeft 'Update
        If CellCharactr(JoeX - 1, JoeY) <> "T" And CellCharactr(JoeX - 1, JoeY) <> "t" Then
        CellCharactr(JoeX - 1, JoeY) = "j"
        End If

        CellContents(JoeX, JoeY) = cEmpty  'Update
        CellCharactr(JoeX, JoeY) = " "
        JoeX = JoeX - 1
        dummy = fncPlayFx(cWalkWithBox)
    
    '***********************************************************************
    'If block was a RailWalkOnce
    If CellContents(JoeX + 1, JoeY + 1) = cRailWalkOnce Then
    'Make it disappear
    .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY + 1), yCorner(JoeX + 1, JoeY + 1)
    CellContents(JoeX + 1, JoeY + 1) = cEmpty
    CellCharactr(JoeX + 1, JoeY + 1) = " "
    dummy = fncPlayFx(cGrayRail)
    End If
   '***********************************************************************
       
       '***********************************************************************
       'Joe fall if nothing underneath
        TempObject = cBox
        TempColumn = JoeX + 1
        TempRow = JoeY - 1
        rtnCheckIfFallAndFall
        booWithBlock = False
       '***********************************************************************
           
            'Check if Transport and process
            If Transport1X = JoeX And Transport1Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport2X, Transport2Y) = cJoeLeft 'Update
            JoeX = Transport2X
            JoeY = Transport2Y
            dummy = fncPlayFx(cTransport)
            rtnRefreshFromCellArrays
            Else
            If Transport2X = JoeX And Transport2Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport1X, Transport1Y) = cJoeLeft 'Update
            JoeX = Transport1X
            JoeY = Transport1Y
            dummy = fncPlayFx(cTransport)
            rtnRefreshFromCellArrays
            End If
            End If
           
       '***********************************************************************
           
       '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
            dummy = fncPlayFx(cExiting)
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
           'If lower block is empty then fall
            If CellContents(JoeX, JoeY + 1) = cEmpty Then
           'Joe fall if nothing underneath
            TempObject = cJoeLeft
            TempColumn = JoeX
            TempRow = JoeY
            rtnCheckIfFallAndFall
            End If
           '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
            dummy = fncPlayFx(cExiting)
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
       '***********************************************************************
    Exit Sub
    End If
   '***********************************************************************
   '***********************************************************************
   '***********************************************************************
End With
End Sub


Public Sub rtnUpKey()
'Enters WITHOUT box
With Form1
Select Case CellContents(JoeX, JoeY)
  Case cJoeRight
   'If Adjacent Right Top Block is Empty and Top Block is empty and Adjacent Right Block is not  Empty then
    If CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 1) = cEmpty And CellContents(JoeX + 1, JoeY) <> cEmpty Then
   
   'Then Jump up
   'Paint and clear Top Block
   .Picture1.PaintPicture picCell(cJoeRight), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    JoeY = JoeY - 1
    CellContents(JoeX, JoeY) = cJoeRight 'Update
    CellCharactr(JoeX, JoeY) = "J"
   'Paint Adjacent Right Top Block
   .Picture1.PaintPicture picCell(cJoeRight), xCorner(JoeX + 1, JoeY), yCorner(JoeX + 1, JoeY)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    JoeX = JoeX + 1
TempCharacter = CellCharactr(JoeX, JoeY) 'Used to recover "T" & "t" on transports
CellContents(JoeX, JoeY) = cJoeRight 'Update Arrays
CellCharactr(JoeX, JoeY) = "J" 'Update Arrays
If TempCharacter = "T" Then CellCharactr(JoeX, JoeY) = "T"
If TempCharacter = "t" Then CellCharactr(JoeX, JoeY) = "t"
dummy = fncPlayFx(cClimb)
       
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX - 1, JoeY + 2) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY + 2), yCorner(JoeX - 1, JoeY + 2)
             CellContents(JoeX - 1, JoeY + 2) = cEmpty
             CellCharactr(JoeX - 1, JoeY + 2) = " "
        dummy = fncPlayFx(cGrayRail)
        End If
       '***********************************************************************
       '***********************************************************************
       
            'Check if Transport and process
            If Transport1X = JoeX And Transport1Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport2X, Transport2Y) = cJoeRight 'Update
            JoeX = Transport2X
            JoeY = Transport2Y
            dummy = fncPlayFx(cTransport)
            rtnRefreshFromCellArrays
            Else
            If Transport2X = JoeX And Transport2Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport1X, Transport1Y) = cJoeRight 'Update
            JoeX = Transport1X
            JoeY = Transport1Y
            dummy = fncPlayFx(cTransport)
            rtnRefreshFromCellArrays
            End If
            End If
       
       
       '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
            dummy = fncPlayFx(cExiting)
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
    'If lower block is empty then fall
    'If CellContents(JoeX, JoeY + 1) = cEmpty Then
    'Joe fall if nothing underneath
    'TempObject = cJoeRight
    'TempColumn = JoeX
    'TempRow = JoeY
    'rtnCheckIfFallAndFall
    'End If
    
    End If

  Case cJoeLeft
   'If    Adjacent Left  Top Block is Empty             and                  Top Block is empty          and    Adjacent Left  Block is not  Empty         then
    If CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 1) = cEmpty And CellContents(JoeX - 1, JoeY) <> cEmpty Then
   'Paint and clear Top Block
   .Picture1.PaintPicture picCell(cJoeLeft), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
   JoeY = JoeY - 1
    CellContents(JoeX, JoeY) = cJoeLeft 'Update
    CellCharactr(JoeX, JoeY) = "j"
   'Paint Adjacent Right Top Block
   .Picture1.PaintPicture picCell(cJoeLeft), xCorner(JoeX - 1, JoeY), yCorner(JoeX - 1, JoeY)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    JoeX = JoeX - 1
TempCharacter = CellCharactr(JoeX, JoeY) 'Used to recover "T" & "t" on transports
CellContents(JoeX, JoeY) = cJoeLeft 'Update
CellCharactr(JoeX, JoeY) = "j"
If TempCharacter = "T" Then CellCharactr(JoeX, JoeY) = "T"
If TempCharacter = "t" Then CellCharactr(JoeX, JoeY) = "t"
dummy = fncPlayFx(cClimb)
       
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX + 1, JoeY + 2) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY + 2), yCorner(JoeX + 1, JoeY + 2)
             CellContents(JoeX + 1, JoeY + 2) = cEmpty
             CellCharactr(JoeX + 1, JoeY + 2) = " "
             dummy = fncPlayFx(cGrayRail)
        End If
       '***********************************************************************
       '***********************************************************************
           
            'Check if Transport and process
            If Transport1X = JoeX And Transport1Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport2X, Transport2Y) = cJoeLeft 'Update
            JoeX = Transport2X
            JoeY = Transport2Y
            dummy = fncPlayFx(cGrayRail)
            rtnRefreshFromCellArrays
            Else
            If Transport2X = JoeX And Transport2Y = JoeY Then
            CellContents(JoeX, JoeY) = cEmpty 'Update
            CellContents(Transport1X, Transport1Y) = cJoeLeft 'Update
            JoeX = Transport1X
            JoeY = Transport1Y
            dummy = fncPlayFx(cTransport)
            rtnRefreshFromCellArrays
            End If
            End If
           
       '***********************************************************************
            'Check if Exit and process
            If ExitColumn = JoeX And ExitRow = JoeY Then
            intLevel = intLevel + 1
            If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
            dummy = fncPlayFx(cExiting)
            rtnInitialize
            Exit Sub
            End If
       '***********************************************************************
    End If

End Select
End With
End Sub

Public Sub rtnUpKeyWithBox()
'Enters WITH box
With Form1
Select Case CellContents(JoeX, JoeY)
  Case cJoeRight
   'If    Adjacent Right Top Top Block is Empty         and       Adjacent Right Top Block is Empty          and          Top Top Block is empty              and    right block is not empty                 then
    If CellContents(JoeX + 1, JoeY - 2) = cEmpty And CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 2) = cEmpty And CellContents(JoeX + 1, JoeY) <> cEmpty Then
   'Paint step 1
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX, JoeY - 2), yCorner(JoeX, JoeY - 2)
   .Picture1.PaintPicture picCell(cJoeRight), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
   
   'Paint step 2
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX + 1, JoeY - 2), yCorner(JoeX + 1, JoeY - 2)
   .Picture1.PaintPicture picCell(cJoeRight), xCorner(JoeX + 1, JoeY - 1), yCorner(JoeX + 1, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY - 2), yCorner(JoeX, JoeY - 2)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
    
    CellContents(JoeX, JoeY - 2) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 2) = " "
    
    CellContents(JoeX + 1, JoeY - 1) = cJoeRight 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = "J"
    
    CellContents(JoeX + 1, JoeY - 2) = cBox 'Update
    CellCharactr(JoeX + 1, JoeY - 2) = "B"
    JoeX = JoeX + 1
    JoeY = JoeY - 1
    dummy = fncPlayFx(cClimbWithBox)
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX - 1, JoeY + 2) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY + 2), yCorner(JoeX - 1, JoeY + 2)
             CellContents(JoeX - 1, JoeY + 2) = cEmpty
             CellCharactr(JoeX - 1, JoeY + 2) = " "
             dummy = fncPlayFx(cGrayRail)
        End If
       '***********************************************************************
       '***********************************************************************
    End If


  Case cJoeLeft
   'If    Adjacent left Top Top Block is Empty         and       Adjacent left Top Block is Empty          and          Top Top Block is empty              and    left block is not empty                 then
    If CellContents(JoeX - 1, JoeY - 2) = cEmpty And CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 2) = cEmpty And CellContents(JoeX - 1, JoeY) <> cEmpty Then
   'Paint step 1
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX, JoeY - 2), yCorner(JoeX, JoeY - 2)
   .Picture1.PaintPicture picCell(cJoeLeft), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY), yCorner(JoeX, JoeY)
   
   'Paint step 2
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX - 1, JoeY - 2), yCorner(JoeX - 1, JoeY - 2)
   .Picture1.PaintPicture picCell(cJoeLeft), xCorner(JoeX - 1, JoeY - 1), yCorner(JoeX - 1, JoeY - 1)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY - 2), yCorner(JoeX, JoeY - 2)
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY) = cEmpty 'Update
    CellCharactr(JoeX, JoeY) = " "
    
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
    
    CellContents(JoeX, JoeY - 2) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 2) = " "
    
    CellContents(JoeX - 1, JoeY - 1) = cJoeLeft 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = "j"
    
    CellContents(JoeX - 1, JoeY - 2) = cBox 'Update
    CellCharactr(JoeX - 1, JoeY - 2) = "B"
    JoeX = JoeX - 1
    JoeY = JoeY - 1
    dummy = fncPlayFx(cClimbWithBox)
       '***********************************************************************
       '***********************************************************************
       'If block was a RailWalkOnce
        If CellContents(JoeX + 1, JoeY + 2) = cRailWalkOnce Then
            'Make it disappear
            .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY + 2), yCorner(JoeX + 1, JoeY + 2)
             CellContents(JoeX + 1, JoeY + 2) = cEmpty
             CellCharactr(JoeX + 1, JoeY + 2) = " "
             dummy = fncPlayFx(cGrayRail)
        End If
       '***********************************************************************
       '***********************************************************************
    End If

End Select
End With
End Sub

Public Sub rtnDownKey()
'Enters WITHOUT box
With Form1
Select Case CellContents(JoeX, JoeY)
  Case cJoeRight
   'Pick up box if available
   'If    Adjacent Right Top Block is Empty             and                  Top Block is empty          and    Adjacent Right Block is a box         then
    If CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 1) = cEmpty And CellContents(JoeX + 1, JoeY) = cBox Then
   'Pick up box
    
    dummy = fncPlayFx(cBoxUp)
   'Paint Adjacent Right Top Box and Clear Adjacent Right Box
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX + 1, JoeY - 1), yCorner(JoeX + 1, JoeY - 1)
    CellContents(JoeX + 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY), yCorner(JoeX + 1, JoeY)
    CellContents(JoeX + 1, JoeY) = cEmpty 'Update
    CellCharactr(JoeX + 1, JoeY) = " "

   'Paint Top Box and Clear Adjacent Right Top Box
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY - 1), yCorner(JoeX + 1, JoeY - 1)
    CellContents(JoeX + 1, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = " "
    booWithBlock = True
    End If

  Case cJoeLeft
   'Pick up box if available
   'If    Adjacent Right Top Block is Empty             and                  Top Block is empty          and    Adjacent Right Block is a box         then
    If CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX, JoeY - 1) = cEmpty And CellContents(JoeX - 1, JoeY) = cBox Then
   'Pick up box

    dummy = fncPlayFx(cBoxUp)
   'Paint Adjacent Left Top Box and Clear Adjacent Left Box
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX - 1, JoeY - 1), yCorner(JoeX - 1, JoeY - 1)
    CellContents(JoeX - 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY), yCorner(JoeX - 1, JoeY)
    CellContents(JoeX - 1, JoeY) = cEmpty 'Update
    CellCharactr(JoeX - 1, JoeY) = " "
   
   'Paint Top Box and Clear Adjacent Left Top Box
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY - 1), yCorner(JoeX - 1, JoeY - 1)
    CellContents(JoeX - 1, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = " "
    booWithBlock = True
    End If

End Select
End With

End Sub

Public Sub rtnDownKeyWithBox()
'Enters WITH box
With Form1
Select Case CellContents(JoeX, JoeY)
  Case cJoeRight
   'Place box down
   'If    Adjacent Right Top Block is Empty             and    Adjacent Right Block is empty         then
    If CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX + 1, JoeY) = cEmpty Then
    dummy = fncPlayFx(cBoxDown)
   
   'Paint Adjacent Right Top Box and Clear Top Box
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX + 1, JoeY - 1), yCorner(JoeX + 1, JoeY - 1)
    CellContents(JoeX + 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
   
   'Paint Adjacent Right Box and Clear Adjacent Right Top Box
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX + 1, JoeY), yCorner(JoeX + 1, JoeY)
    CellContents(JoeX + 1, JoeY) = cBox 'Update
    CellCharactr(JoeX + 1, JoeY) = "B"
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX + 1, JoeY - 1), yCorner(JoeX + 1, JoeY - 1)
    CellContents(JoeX + 1, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX + 1, JoeY - 1) = " "
    booWithBlock = False
    
    'Box fall is nothing underneath
    TempObject = cBox
    TempColumn = JoeX + 1
    TempRow = JoeY
    rtnCheckIfFallAndFall
    Exit Sub
    End If

       'Place box on top of object
       'If    Adjacent Right Top Block is Empty             and    Adjacent Right Block is not empty         then
        If CellContents(JoeX + 1, JoeY - 1) = cEmpty And CellContents(JoeX + 1, JoeY) <> cEmpty Then
        dummy = fncPlayFx(cBoxSlide)
  
       'Paint Adjacent Right Top Box and Clear Top Box
       .Picture1.PaintPicture picCell(cBox), xCorner(JoeX + 1, JoeY - 1), yCorner(JoeX + 1, JoeY - 1)
        CellContents(JoeX + 1, JoeY - 1) = cBox 'Update
        CellCharactr(JoeX + 1, JoeY - 1) = "B"
       .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
        CellContents(JoeX, JoeY - 1) = cEmpty 'Update
        CellCharactr(JoeX, JoeY - 1) = " "
        booWithBlock = False
        Exit Sub
        End If
  
  Case cJoeLeft
   'If    Adjacent Left Top Block is Empty              and    Adjacent Left Block is empty         then
    If CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX - 1, JoeY) = cEmpty Then
   'Place box down
    dummy = fncPlayFx(cBoxDown)
   
   'Paint Adjacent Left Top Box and Clear Top Box
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX - 1, JoeY - 1), yCorner(JoeX - 1, JoeY - 1)
    CellContents(JoeX - 1, JoeY - 1) = cBox 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = "B"
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
    CellContents(JoeX, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX, JoeY - 1) = " "
   
   'Paint Adjacent Left Box and Clear Adjacent Left Top Box
   .Picture1.PaintPicture picCell(cBox), xCorner(JoeX - 1, JoeY), yCorner(JoeX - 1, JoeY)
    CellContents(JoeX - 1, JoeY) = cBox 'Update
    CellCharactr(JoeX - 1, JoeY) = "B"
   .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX - 1, JoeY - 1), yCorner(JoeX - 1, JoeY - 1)
    CellContents(JoeX - 1, JoeY - 1) = cEmpty 'Update
    CellCharactr(JoeX - 1, JoeY - 1) = " "
    booWithBlock = False
    
    'Box fall is nothing underneath
    TempObject = cBox
    TempColumn = JoeX - 1
    TempRow = JoeY
    rtnCheckIfFallAndFall
    Exit Sub
    End If

       'Place box on top of object
       'If    Adjacent Left Top Block is Empty             and    Adjacent Left Block is not empty         then
        If CellContents(JoeX - 1, JoeY - 1) = cEmpty And CellContents(JoeX - 1, JoeY) <> cEmpty Then
        dummy = fncPlayFx(cBoxSlide)
  
       'Paint Adjacent Left Top Box and Clear Top Box
       .Picture1.PaintPicture picCell(cBox), xCorner(JoeX - 1, JoeY - 1), yCorner(JoeX - 1, JoeY - 1)
        CellContents(JoeX - 1, JoeY - 1) = cBox 'Update
        CellCharactr(JoeX - 1, JoeY - 1) = "B"
       .Picture1.PaintPicture picCell(cEmpty), xCorner(JoeX, JoeY - 1), yCorner(JoeX, JoeY - 1)
        CellContents(JoeX, JoeY - 1) = cEmpty 'Update
        CellCharactr(JoeX, JoeY - 1) = " "
        booWithBlock = False
        Exit Sub
        End If



End Select
End With

End Sub


Public Sub rtnCheckIfFallAndFall()
'Enter with
'TempObject holding  object # to fall down
'TempColumn holding object's x cell value 1 - 20
'TempRow holding object's y cell value 1 - 15
With Form1
        'If Joe or Box fall then sound effect
        'Not on Ball fall because roll,fall,roll,fall,etc on stairs
        'may not sound right
        If CellContents(TempColumn, TempRow + 1) = cEmpty Then
              Select Case TempObject
              Case cBox
              dummy = fncPlayFx(cBoxFall)
              Case cJoeRight
              dummy = fncPlayFx(cJoeFall)
              Case cJoeLeft
              dummy = fncPlayFx(cJoeFall)
              End Select
        End If
        
        'Fall if nothing underneath
        If CellContents(TempColumn, TempRow + 1) = cEmpty Then
           '***********************************************************************
            Do While CellContents(TempColumn, TempRow + 1) = cEmpty
           '       Paint           This Picture  at   Pixelx of   Column,  Row          Pixely of Column,  Row
           .Picture1.PaintPicture picCell(TempObject), xCorner(TempColumn, TempRow + 1), yCorner(TempColumn, TempRow + 1)
           .Picture1.PaintPicture picCell(cEmpty), xCorner(TempColumn, TempRow), yCorner(TempColumn, TempRow)
           .Picture1.Refresh
            CellContents(TempColumn, TempRow) = cEmpty 'Update
            CellCharactr(TempColumn, TempRow) = " "
            TempRow = TempRow + 1
            If TempObject = cJoeRight Then JoeY = JoeY + 1
            If TempObject = cJoeLeft Then JoeY = JoeY + 1
            CellContents(TempColumn, TempRow) = TempObject 'Update
            
              Select Case TempObject
              Case cBox
              CellCharactr(TempColumn, TempRow) = "B"
              Case cJoeRight
              CellCharactr(TempColumn, TempRow) = "J"
              Case cJoeLeft
              CellCharactr(TempColumn, TempRow) = "j"
              Case cRoll
              CellCharactr(TempColumn, TempRow) = "O"
              End Select
            Loop
           '***********************************************************************
        End If
        If TempObject = cBox And CellContents(TempColumn, TempRow + 1) = cJoeRight And booClunk = True Then
        dummy = fncPlayFx(cClunk)
        booClunk = False
        End If
        If TempObject = cBox And CellContents(TempColumn, TempRow + 1) = cJoeLeft And booClunk = True Then
        dummy = fncPlayFx(cClunk)
        booClunk = False
        End If
End With
End Sub

Public Sub rtnAllRollsAndBoxesFall()
Dim x As Integer, y As Integer
'All Boxes Fall Down
'TempObject holding value of object to fall down
'TempColumn holding object's x cell value
'TempRow holding object's y cell value
For y = 15 To 1 Step -1
For x = 1 To 20
If CellContents(x, y) = cRoll Or CellContents(x, y) = cBox Then
   TempObject = CellContents(x, y)
   TempColumn = x
   TempRow = y
   rtnCheckIfFallAndFall
End If
Next x
Next y

End Sub

Public Sub rtnRollRight()
'Enter with
'TempObject holding  object # to fall down
'TempColumn holding object's x cell value 1 - 20
'TempRow holding object's y cell value 1 - 15
With Form1
        If CellContents(TempColumn + 1, TempRow) = cEmpty Then
           '***********************************************************************
            'Roll right while front cell empty
            dummy = fncPlayFx(cBallRoll)
            Do While CellContents(TempColumn + 1, TempRow) = cEmpty
           'Paint front cell where object is now
           .Picture1.PaintPicture picCell(TempObject), xCorner(TempColumn + 1, TempRow), yCorner(TempColumn + 1, TempRow)
           'Clear cell where object was
           .Picture1.PaintPicture picCell(cEmpty), xCorner(TempColumn, TempRow), yCorner(TempColumn, TempRow)
           .Picture1.Refresh
            'Update cell arrays where object was
            CellContents(TempColumn, TempRow) = cEmpty 'Update
                If CellCharactr(TempColumn, TempRow) <> "T" And CellCharactr(TempColumn, TempRow) <> "t" Then
                CellCharactr(TempColumn, TempRow) = " "
                End If
            'Update cell arrays where object is now
            TempColumn = TempColumn + 1
            CellContents(TempColumn, TempRow) = TempObject 'Update
                If CellCharactr(TempColumn, TempRow) <> "T" And CellCharactr(TempColumn, TempRow) <> "t" Then
                  Select Case TempObject
                  Case cBox
                  CellCharactr(TempColumn, TempRow) = "B" 'Ball
                  Case cJoeRight
                  CellCharactr(TempColumn, TempRow) = "J" 'Joe Right
                  Case cJoeLeft
                  CellCharactr(TempColumn, TempRow) = "j" 'Joe Left
                  Case cRoll
                  CellCharactr(TempColumn, TempRow) = "O" 'Ball
                  End Select
                End If
            
               '***********************************************************************
               rtnCheckIfFallAndFall
               '***********************************************************************
            Loop
           '***********************************************************************
        dummy = fncPlayFx(cBallStop)
        End If
rtnAllRollsAndBoxesFall
End With
End Sub

Public Sub rtnRollLeft()
'Enter with
'TempObject holding  object # to roll
'TempColumn holding object's x cell value 1 - 20
'TempRow holding object's y cell value 1 - 15
With Form1
    If CellContents(TempColumn - 1, TempRow) = cEmpty Then
           '***********************************************************************
            'Roll left while front cell empty
            dummy = fncPlayFx(cBallRoll)
            Do While CellContents(TempColumn - 1, TempRow) = cEmpty
           'Paint front cell where object is now
           .Picture1.PaintPicture picCell(TempObject), xCorner(TempColumn - 1, TempRow), yCorner(TempColumn - 1, TempRow)
           'Clear cell where object was
           .Picture1.PaintPicture picCell(cEmpty), xCorner(TempColumn, TempRow), yCorner(TempColumn, TempRow)
           .Picture1.Refresh
            'Update cell arrays where object was
                If CellCharactr(TempColumn, TempRow) <> "T" And CellCharactr(TempColumn, TempRow) <> "t" Then
                CellCharactr(TempColumn, TempRow) = " "
                End If
            CellContents(TempColumn, TempRow) = cEmpty
            'Update cell arrays where object is now
            TempColumn = TempColumn - 1
            CellContents(TempColumn, TempRow) = TempObject
                If CellCharactr(TempColumn, TempRow) <> "T" And CellCharactr(TempColumn, TempRow) <> "t" Then
                  Select Case TempObject
                  Case cBox
                  CellCharactr(TempColumn, TempRow) = "B" 'Ball
                  Case cJoeRight
                  CellCharactr(TempColumn, TempRow) = "J" 'Joe Right
                  Case cJoeLeft
                  CellCharactr(TempColumn, TempRow) = "j" 'Joe Left
                  Case cRoll
                  CellCharactr(TempColumn, TempRow) = "O" 'Ball
                  End Select
                End If
               '***********************************************************************
               'Fall while bottom cell empty
               rtnCheckIfFallAndFall
               '***********************************************************************
            Loop
           '***********************************************************************
        dummy = fncPlayFx(cBallStop)
        End If
rtnAllRollsAndBoxesFall
End With
End Sub

Public Sub rtnFinale()
Dim x As Integer, y As Integer
With Form1
Sleep 2000 'Wait for exit wave to finish
.imgExit.Visible = False
.imgTransport(1).Visible = False
.imgTransport(2).Visible = False
.Picture1.Cls
.Picture1.BackColor = vbBlack
.Picture1.Picture = LoadPicture("")
.Picture1.ForeColor = vbWhite
.Picture1.FontSize = 50
.Picture1.Font = "arial"
.Picture1.Print "Congratulations!"
.Picture1.Print "You Win!"
.Picture1.Refresh

Sleep 2000 'Wait a little more
Call sndPlaySound("_Fanfare.wav", SND_ASYNC)
y = 7
For x = 1 To 10
.Picture1.Cls
.Picture1.PaintPicture .imgFinalMrJoeRight.Picture, xCorner(x, y), yCorner(x, y)
.Picture1.PaintPicture .imgFinalMrsJoeLeft.Picture, xCorner(21 - x, y), yCorner(21 - x, y)
.Picture1.Refresh
Sleep 500
Next x
 
 Sleep 5000  ' Pause 5 seconds
End With

End Sub

