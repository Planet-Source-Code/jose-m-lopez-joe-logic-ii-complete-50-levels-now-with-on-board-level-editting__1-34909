VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Joe Logic"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   791
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFinale 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10200
      Top             =   7200
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   9720
      Top             =   7200
   End
   Begin VB.PictureBox pbxLevelSelect 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   4320
      ScaleHeight     =   2415
      ScaleWidth      =   1425
      TabIndex        =   29
      Top             =   840
      Visible         =   0   'False
      Width           =   1425
      Begin VB.ListBox listLevelSelect 
         BackColor       =   &H00000080&
         ForeColor       =   &H00FFFFFF&
         Height          =   2205
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   0
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   28
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton cmdEnd 
         BackColor       =   &H00FFFF80&
         Caption         =   "End"
         Height          =   375
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image imgTransport 
         Height          =   450
         Index           =   1
         Left            =   3240
         Picture         =   "Play Area.frx":0000
         Top             =   960
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgTransport 
         Height          =   450
         Index           =   2
         Left            =   3240
         Picture         =   "Play Area.frx":0B8A
         Top             =   1560
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgExit 
         Height          =   450
         Left            =   3240
         Picture         =   "Play Area.frx":1714
         Top             =   2160
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   9240
      Top             =   7200
   End
   Begin VB.Frame frameEditLevels 
      Height          =   735
      Left            =   240
      TabIndex        =   22
      Top             =   6840
      Width           =   3135
      Begin VB.CommandButton cmdEditUpDownStore 
         Caption         =   "/\"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdEditUpDownStore 
         Caption         =   "update"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdEditUpDownStore 
         Caption         =   "\/"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Update Level"
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame FrameSave 
      Caption         =   "Save As"
      Height          =   4215
      Left            =   480
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdCancelSave 
         Caption         =   "Cancel"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton cmdSaveAs 
         Caption         =   "Save"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   1695
      End
      Begin VB.FileListBox fileSave 
         Height          =   2430
         Left            =   120
         Pattern         =   "*.joe"
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtSave 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblSave 
         Caption         =   "Label1"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   1815
      End
   End
   Begin VB.Frame frameOpen 
      Caption         =   "Open"
      Height          =   3975
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdCancelOpen 
         Caption         =   "Cancel"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   1695
      End
      Begin VB.FileListBox fileOpen 
         Height          =   2430
         Left            =   120
         Pattern         =   "*.joe"
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblOpen 
         Caption         =   "Label1"
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   10
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   11
      Top             =   6480
      Width           =   975
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   9
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   10
      Top             =   5760
      Width           =   975
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   8
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   9
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   7
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   6
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   5
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   4
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox pbxSmall 
      AutoRedraw      =   -1  'True
      Height          =   135
      Left            =   10800
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   840
      Width           =   135
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   3
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   2
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.PictureBox pbxWork 
      AutoRedraw      =   -1  'True
      Height          =   735
      Index           =   1
      Left            =   9120
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   6960
      Width           =   5415
      Begin VB.Image Image1 
         Height          =   450
         Index           =   11
         Left            =   4800
         Picture         =   "Play Area.frx":229E
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   10
         Left            =   4320
         Picture         =   "Play Area.frx":2AA0
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   9
         Left            =   3840
         Picture         =   "Play Area.frx":362A
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   8
         Left            =   3360
         Picture         =   "Play Area.frx":3E2C
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   7
         Left            =   2880
         Picture         =   "Play Area.frx":462E
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   6
         Left            =   2400
         Picture         =   "Play Area.frx":51B8
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   5
         Left            =   1920
         Picture         =   "Play Area.frx":5D42
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   4
         Left            =   1440
         Picture         =   "Play Area.frx":64F4
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   3
         Left            =   960
         Picture         =   "Play Area.frx":6CF6
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   2
         Left            =   480
         Picture         =   "Play Area.frx":7800
         Top             =   0
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   1
         Left            =   0
         Picture         =   "Play Area.frx":8002
         Top             =   0
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10800
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgFinalMrsJoeLeft 
      Height          =   870
      Left            =   10680
      Picture         =   "Play Area.frx":8804
      Top             =   3720
      Width           =   585
   End
   Begin VB.Image imgFinalMrJoeRight 
      Height          =   870
      Left            =   10680
      Picture         =   "Play Area.frx":A376
      Top             =   2400
      Width           =   585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   ""
      End
      Begin VB.Menu mnuMusic 
         Caption         =   "Music"
      End
      Begin VB.Menu mnuSoundEffects 
         Caption         =   "Sound Effects"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   ""
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditEditMode 
         Caption         =   "Edit Mode"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuUpdateMoves 
         Caption         =   "Update Solution Level"
      End
   End
   Begin VB.Menu mnuLevel 
      Caption         =   "Level"
      Begin VB.Menu mnuLevelUp 
         Caption         =   "Level Up"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuLevelDown 
         Caption         =   "Level Down"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLevelManual 
         Caption         =   "Level Select"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuLevelRestart 
         Caption         =   "Restart Level"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuSolution 
      Caption         =   "Solution"
      Begin VB.Menu mnuSolutionSpeed0 
         Caption         =   "Fast"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuSolutionSpeed1 
         Caption         =   "Medium"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSolutionSpeed2 
         Caption         =   "Slow"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSolutionStart 
         Caption         =   "Solution Start"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuSolutionStop 
         Caption         =   "Solution Stop"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuRegister 
         Caption         =   "Register"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim x As Integer, y As Integer, Z As Integer
Form1.Show
dblPreviousSoundEffectTime = Timer
booClunk = False 'Whether or not a box landing on Joe
                 'should play "clunk.wav" sound
                 'This is so the clunk is not played in short jumps
                 'like going down stairs
intMovesMaxLineSize = 50 'Maximum Length of each solution line
intMovesMaxLineQnty = 200 'Maximum Quantity of solution lines
intMaxLevelQnty = 100 'Maximum levels allowed to play.
                      'If intMaxLevelQnty is the same or greater than
                      'the number of levels in the
                      'game file then reaching the final level exit
                      'ends the game with the fanfare.
                      'If intMaxLevelQnty is smaller than (for shareware)
                      'the game file's levels then the game will not be
                      'allowed to get past level intMaxLevelQnty
                      'Reaching the exit in intMaxLevelQnty
                      'restarts that level
Picture1.Visible = True 'Main game screen
frameOpen.Visible = False 'Game open dialog
FrameSave.Visible = False 'Game save dialog
ChDir App.Path 'Only this directory will be accessible
Form1.Width = 9150
Form1.Height = 7470
Picture1.Top = 2
Picture1.Left = 2

'Set Your Normal Game Objects, Ball, Block, JoeRight/Left, etc
Set picCell(1) = Image1(1) 'empty = " " 1
Set picCell(2) = Image1(2) 'box = "B" 2
Set picCell(3) = Image1(3) 'roll = "O" 3
Set picCell(4) = Image1(4) 'rail = "=" 4
Set picCell(5) = Image1(5) 'brick = "#" 5
Set picCell(6) = Image1(6) 'transport = "T" 6  icon
Set picCell(7) = Image1(7) 'transport = "t" 7  icon
Set picCell(8) = Image1(8) 'joeleft = "j" 8
Set picCell(9) = Image1(9) 'joeright = "J" 9
Set picCell(10) = Image1(10) '        exit = "X" 10  icon
Set picCell(11) = Image1(11) 'railwalkonce = "." 11

'On edit mode our user can store his main editting screen
'to any of **10** small storage picture boxes.
'These picture boxes show a small representaion
'of the stored screens.

    'Set the scale the same as the main screen so our user can view
    'small copies of the main editting screen that he/she stores
    For x = 1 To 10
    pbxWork(x).ScaleWidth = 600
    pbxWork(x).ScaleHeight = 450
    Next x
'Set Your Small Game Objects for the **10** Small Storage picture boxes.
'These are used to paint the objects including "empty"
'in the small user storeage pictureboxes. Because of the tiny size of these
'objects we only paint an average color instead of a tiny undecernable object.
pbxSmall.Picture = LoadPicture("")
pbxSmall.BackColor = vbBlack
pbxSmall.Picture = pbxSmall.Image
Set picCellSmall(1) = pbxSmall.Picture  'empty = " " 1

pbxSmall.Picture = LoadPicture("")
pbxSmall.BackColor = RGB(204, 153, 0)
pbxSmall.Picture = pbxSmall.Image
Set picCellSmall(2) = pbxSmall.Picture 'box = "B" 2

pbxSmall.Picture = LoadPicture("")
pbxSmall.BackColor = vbBlue
pbxSmall.Picture = pbxSmall.Image
Set picCellSmall(3) = pbxSmall.Picture 'roll = "O" 3

pbxSmall.Picture = LoadPicture("")
pbxSmall.BackColor = vbGreen
pbxSmall.Picture = pbxSmall.Image
Set picCellSmall(4) = pbxSmall.Picture 'rail = "=" 4

pbxSmall.Picture = LoadPicture("")
pbxSmall.BackColor = vbRed
pbxSmall.Picture = pbxSmall.Image
Set picCellSmall(5) = pbxSmall.Picture 'brick = "#" 5
Set picCellSmall(10) = pbxSmall.Picture '        exit = "X" 10  icon

pbxSmall.Picture = LoadPicture("")
pbxSmall.BackColor = vbWhite
pbxSmall.Picture = pbxSmall.Image
Set picCellSmall(6) = pbxSmall.Picture 'transport = "T" 6  icon

pbxSmall.Picture = LoadPicture("")
pbxSmall.BackColor = vbWhite
pbxSmall.Picture = pbxSmall.Image
Set picCellSmall(7) = pbxSmall.Picture 'transport = "t" 7  icon

pbxSmall.Picture = LoadPicture("")
pbxSmall.BackColor = vbMagenta
pbxSmall.Picture = pbxSmall.Image
Set picCellSmall(8) = pbxSmall.Picture 'joeleft = "j" 8
Set picCellSmall(9) = pbxSmall.Picture 'joeright = "J" 9

pbxSmall.Picture = LoadPicture("")
pbxSmall.BackColor = vbButtonFace
pbxSmall.Picture = pbxSmall.Image
Set picCellSmall(11) = pbxSmall.Picture 'railwalkonce = "." 11

'Pixel Log
'Log the x, y pixel corner of each and every cell on our main screen
For y = 1 To 15
For x = 1 To 20
xCorner(x, y) = (x - 1) * 30
yCorner(x, y) = (y - 1) * 30
Next x, y

'***************************************************
'***************************************************
'Load "Empty" array representing a main screen
'with no objects except border,joe, and empties
'Then paint our small user storeage pictureboxes with this
Dim emptyCellContents(20, 15) As Integer
Dim emptyCellCharactr(20, 15) As String
For x = 1 To 20
emptyCellContents(x, 1) = cBrick
emptyCellCharactr(x, 1) = "#"
emptyCellContents(x, 15) = cBrick
emptyCellCharactr(x, 15) = "#"
Next x
For y = 1 To 15
emptyCellContents(1, y) = cBrick
emptyCellCharactr(1, y) = "#"
emptyCellContents(20, y) = cBrick
emptyCellCharactr(20, y) = "#"
Next y
For y = 2 To 14
For x = 2 To 19
emptyCellContents(x, y) = cEmpty
emptyCellCharactr(x, y) = " "
Next x, y
emptyCellContents(7, 14) = cJoeRight
emptyCellCharactr(7, 14) = "J"

'Load the 10 work arrays with our "Empty" user storeage arrays
'(small picture boxes)
For Z = 1 To 10
For y = 1 To 15
For x = 1 To 20
StoredCellContents(Z, x, y) = emptyCellContents(x, y)
StoredCellCharactr(Z, x, y) = emptyCellCharactr(x, y)
Next x, y, Z
'***************************************************
'***************************************************


sngSolutionDelay = 0 'Set Default Solution Speed
mnuSolutionStart.Enabled = True
mnuSolutionStop.Enabled = False
mnuSolutionSpeed0.Checked = True 'Fast Solution Speed

'Load game from HD
stFileName = "levels.joe" 'Main Play File
rtnLoadMasterArray
intLevel = 1 'Starting Level
rtnInitialize 'Do it
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IsMusicOn = False
    'Midi Stop
    RetValue = mciSendString("CLOSE BackgroundMusic", "", 0, 0)
Unload Form1
End
End Sub

Private Sub cmdEnd_Click()
Unload Form1
End
End Sub

'**********************************
'***********Open*******************
'**********************************
Private Sub mnuFileOpen_Click()
'Menu Open new game
mnuFile.Enabled = False
mnuEdit.Enabled = False
mnuLevel.Enabled = False
mnuSolution.Enabled = False
mnuHelp.Enabled = False
Picture1.Visible = False
frameOpen.Visible = True
fileOpen.Refresh
lblOpen.Caption = "Please select by doubleclicking"
End Sub
Private Sub fileOpen_DblClick()
'User just selected new game
mnuFile.Enabled = True
mnuEdit.Enabled = True
mnuLevel.Enabled = True
mnuSolution.Enabled = True
mnuHelp.Enabled = True
Picture1.Visible = True
frameOpen.Visible = False
stFileName = fileOpen.FileName
rtnLoadMasterArray
intLevel = 1
rtnInitialize
End Sub
Private Sub cmdCancelOpen_Click()
'User just canceled from selecting new game
mnuFile.Enabled = True
mnuEdit.Enabled = True
mnuLevel.Enabled = True
mnuSolution.Enabled = True
mnuHelp.Enabled = True
Picture1.Visible = True
frameOpen.Visible = False
End Sub





Private Sub mnuNew_Click()
'Menu New game from scratch
'One empty level
Dim strMessage As String
Dim x As Integer, y As Integer
ReDim strMasterArray(18)

'Load Master Array with an empty level
intLevel = 1
intMasterFileSize = 18
 
 strMasterArray(1) = "level1"
 strMasterArray(2) = "####################"
 strMasterArray(3) = "#                  #"
 strMasterArray(4) = "#                  #"
 strMasterArray(5) = "#                  #"
 strMasterArray(6) = "#                  #"
 strMasterArray(7) = "#                  #"
 strMasterArray(8) = "#                  #"
 strMasterArray(9) = "#                  #"
strMasterArray(10) = "#                  #"
strMasterArray(11) = "#                  #"
strMasterArray(12) = "#                  #"
strMasterArray(13) = "#                  #"
strMasterArray(14) = "#                  #"
strMasterArray(15) = "#     J            #"
strMasterArray(16) = "####################"
strMasterArray(17) = "moves1"
strMasterArray(18) = "0"

stFileName = "NewGameFile.joe"
rtnInitialize
End Sub

Private Sub mnuSave_Click()
'Menu Save current game
Dim x As Integer
    'Save array to current filename
    Open stFileName For Output As #1
        For x = 1 To intMasterFileSize
        Print #1, strMasterArray(x)
        Next x
        Close
End Sub

Private Sub mnuSaveAs_Click()
'Menu Save current game as
mnuFile.Enabled = False
mnuEdit.Enabled = False
mnuLevel.Enabled = False
mnuSolution.Enabled = False
mnuHelp.Enabled = False
Picture1.Visible = False
FrameSave.Visible = True
fileSave.Refresh
lblSave.Caption = "Doubleclick a filename or enter in textbox"
txtSave.Text = ""
End Sub
Private Sub fileSave_DblClick()
'User just selected an existing file to possibly save as
txtSave.Text = fileSave.FileName
End Sub
Private Sub cmdSaveAs_Click()
'User just clicked save as command button
Dim tmpCount As Integer
Dim tmpText As String
Dim x As Integer
    'Check if a filename entered and check extension
    If txtSave.Text = "" Or Right(txtSave.Text, 4) <> ".joe" Then
    intResponse = MsgBox("Please enter filename and .joe extension.", vbOKOnly, "Filename Error! ")
    Exit Sub
    End If
stFileName = txtSave.Text
    
    'Save array to user selected filename
    Open stFileName For Output As #1
        For x = 1 To intMasterFileSize
        Print #1, strMasterArray(x)
        Next x
        Close

'Menu maintenance and click prevention
mnuFile.Enabled = True
mnuEdit.Enabled = True
mnuLevel.Enabled = True
mnuSolution.Enabled = True
mnuHelp.Enabled = True
Picture1.Visible = True
FrameSave.Visible = False
rtnInitialize
End Sub

Private Sub cmdCancelSave_Click()
'User just canceled from saving game

'Menu maintenance
mnuFile.Enabled = True
mnuEdit.Enabled = True
mnuLevel.Enabled = True
mnuSolution.Enabled = True
mnuHelp.Enabled = True
Picture1.Visible = True
FrameSave.Visible = False
End Sub

Private Sub mnuMusic_Click()
'Menu Music
mnuMusic.Checked = Not mnuMusic.Checked
    If mnuMusic.Checked = True Then
    IsMusicOn = True
    'Get Midi And Play
    RetValue = mciSendString("OPEN _BackGround.mid TYPE SEQUENCER ALIAS BackgroundMusic", "", 0, 0)
    Else
    IsMusicOn = False
    'Midi Stop
    RetValue = mciSendString("CLOSE BackgroundMusic", "", 0, 0)
    End If
End Sub

Private Sub mnuSoundEffects_Click()
'Menu Sound Effects
mnuSoundEffects.Checked = Not mnuSoundEffects.Checked
End Sub

Private Sub mnuedit_click()
'Menu maintenance
mnuUpdateMoves.Caption = "Update Solution Level " & intLevel
End Sub

Private Sub mnuUpdateMoves_Click()
'Menu Update Moves
Dim x As Integer, y As Integer, tmpintCount As Integer
Dim tmpstrSolutionLines() As String, tmpText As String
Dim response As Integer, tmpintSolutionLinesCount As Integer
'***************************************
'Tell user what this routine does
'***************************************
Dim strMessage As String
strMessage = "All moves you have made since you started Level " & intLevel & vbCrLf
strMessage = strMessage & "will become its solution in memory." & vbCrLf
response = MsgBox(strMessage, vbOKCancel, "Update Solution")
    If response = vbCancel Then
    Exit Sub
    End If
'***************************************
'***************************************


'***************************************
'Convert all new solution moves to line format and
'Load temporarily tmpstrSolutionLines() with the new lines
'Load temporarily tmpintSolutionLinesCount with the new lines count
'MasterArray unmodified
'***************************************
ReDim tmpstrSolutionLines(intMovesMaxLineQnty) 'max possible number if lines
    tmpintCount = 0: tmpText = ""
    For y = 1 To intMovesMaxLineQnty
        For x = 1 To intMovesMaxLineSize
        tmpintCount = tmpintCount + 1
        tmpText = tmpText & arrMoves(tmpintCount)
            If arrMoves(tmpintCount) = "0" Then
            tmpstrSolutionLines(y) = tmpText
            tmpintSolutionLinesCount = y
            GoTo lblOne
            End If
        Next x
    tmpstrSolutionLines(y) = tmpText
    tmpText = ""
    Next y
response = MsgBox("Error, No End of Moves '0' encountered. Possibly over limit of 10000 moves! Game will end.", vbOKOnly, "Error!")
Unload Form1
End
'***************************************
'***************************************

'***************************************
'Count from beginning to current moves label
'Count from next moves label to end
'Load tmpintQntyLinesBefore with line count (beginning to current moves label)
'Load tmpintQntyLinesAfter with line count, (next moves label to end, 0 if none)
'MasterArray unmodified
'***************************************
lblOne:
Dim tmpintQntyLinesBefore As Integer, tmpintQntyLinesAfter As Integer
Dim tmpintAfterIndex As Integer
For x = 1 To intMasterFileSize
    If strMasterArray(x) = "moves" & Trim(Str(intLevel)) Then
    'Quantity of lines before current solution including label
    tmpintQntyLinesBefore = x
    GoTo lblPartTwo
    End If
Next x
response = MsgBox("Error, Moves" & Trim(Str(intLevel)) & " not found.", vbOKOnly, "Error!")
Unload Form1
End

lblPartTwo:
For x = 1 To intMasterFileSize
    If strMasterArray(x) = "moves" & Trim(Str(intLevel + 1)) Then
    'Start line after current solution at move label
    tmpintAfterIndex = x
    'Quantity of lines after current solution including label
    tmpintQntyLinesAfter = intMasterFileSize - x + 1
    GoTo lblTwo
    End If
Next x
'No more solutions after current solution
tmpintQntyLinesAfter = 0
'***************************************
'***************************************

'***************************************
'Calculate New MasterFileSize
'Load MasterArray with updated master file
'***************************************
lblTwo:
Dim tmpintNewMasterFileSize As Integer
Dim tmpstrMasterArray()
tmpintNewMasterFileSize = tmpintQntyLinesBefore + tmpintSolutionLinesCount + tmpintQntyLinesAfter
ReDim tmpstrMasterArray(tmpintNewMasterFileSize)
    
tmpintCount = 0
For x = 1 To tmpintQntyLinesBefore
tmpintCount = tmpintCount + 1
tmpstrMasterArray(x) = strMasterArray(x)
Next x

For x = 1 To tmpintSolutionLinesCount
tmpintCount = tmpintCount + 1
tmpstrMasterArray(tmpintCount) = tmpstrSolutionLines(x)
Next x

If tmpintQntyLinesAfter > 0 Then
For x = tmpintAfterIndex To tmpintAfterIndex + tmpintQntyLinesAfter - 1
tmpintCount = tmpintCount + 1
tmpstrMasterArray(tmpintCount) = strMasterArray(x)
Next x
End If

ReDim strMasterArray(tmpintNewMasterFileSize)
For x = 1 To tmpintNewMasterFileSize
strMasterArray(x) = tmpstrMasterArray(x)
Next x
intMasterFileSize = tmpintNewMasterFileSize
'***************************************
'***************************************
End Sub

Private Sub mnuFileExit_Click()
'End
Unload Form1
End
End Sub

Private Sub mnuEditEditMode_Click()
'Menu Edit
Dim x As Integer, y As Integer, Z As Integer
Dim intLine As Integer, intChar As Integer
Picture1.Picture = Picture1.Image
mnuEditEditMode.Checked = Not mnuEditEditMode.Checked
    
    'Enter into edit mode
    If mnuEditEditMode.Checked = True Then
    'menu maintenance
    mnuFile.Enabled = False
    mnuLevel.Enabled = False
    mnuSolution.Enabled = False
    mnuHelp.Enabled = False
    mnuUpdateMoves.Enabled = False
    Form1.Width = 10245
    Form1.Height = 8235
    Form1.Caption = "Edit Mode            Level" & Str(intLevel) & "  " & stFileName
        'Paint small storage picture boxes
        For Z = 1 To 10
        For y = 1 To 15
        For x = 1 To 20
        pbxWork(Z).PaintPicture picCellSmall(StoredCellContents(Z, x, y)), xCorner(x, y), yCorner(x, y)
        pbxWork(Z).Refresh
        Next x, y, Z
    Label1.Caption = "Update level " & intLevel & " in memory"
    
    'Enter into normal run mode
    Else
    'Prompt user if updating current level with freshly editted screen is desired
    intResponse = MsgBox("Would you like to update Level " & intLevel & " with the one you just editted?", vbYesNo, "Update Level " & intLevel)
        
        'Update
        If intResponse = vbYes Then
        Dim tmpText As String, tmpCount As Integer
        tmpText = ""
        For y = 1 To 15       'y cells
            For x = 1 To 20   'x cells
            tmpText = tmpText & CellCharactr(x, y)
            Next x
            '    ((point to prevlvl last line)  =line#)
            strMasterArray(((intLevel - 1) * 16) + y + 1) = tmpText
            tmpText = ""
        Next y
        End If
    
    'menu maintenance
    mnuFile.Enabled = True
    mnuLevel.Enabled = True
    mnuSolution.Enabled = True
    mnuHelp.Enabled = True
    mnuUpdateMoves.Enabled = True
    Form1.Width = 9150
    Form1.Height = 7470
            rtnInitialize
    End If

EditObject = 1 'cEmpty
'rtnRefreshFromCellArrays
'rtnAllRollsAndBoxesFall
End Sub

Private Sub mnuLevelRestart_Click()
'Menu Restart level
rtnInitialize
End Sub

Private Sub mnuLevelUp_Click()
'Menu Level Up
Dim tmpText As String
    'Go Up One Level
    intLevel = intLevel + 1
    If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
    
rtnCheckIfLevelAvailable

If booLevelAvailable = True Then 'Level found so load it
rtnInitialize
Exit Sub
Else 'There wasn't a higher level so load the current one
intLevel = intLevel - 1
rtnInitialize
Exit Sub
End If

End Sub

Private Sub mnuLevelDown_Click()
'Menu Level Down
    'Go Down One Level
    intLevel = intLevel - 1
        If intLevel = 0 Then
        intLevel = intLevel + 1
        Exit Sub
        End If
rtnInitialize
End Sub

Private Sub mnuLevelManual_Click()
'Menu Manual Level Select
Dim x As Integer
Dim tmpintLevel As Integer
Picture1.Enabled = False
pbxLevelSelect.Visible = True
'Enable and disable menu to avoid conflicts
mnuFile.Enabled = False
mnuEdit.Enabled = False
mnuSolution.Enabled = False
mnuLevel.Enabled = False
mnuHelp.Enabled = False
listLevelSelect.Clear
listLevelSelect.AddItem "Level Select"

'booLevelAvailable = True
tmpintLevel = intLevel
intLevel = 0
Do
intLevel = intLevel + 1
rtnCheckIfLevelAvailable
    If booLevelAvailable = False Then Exit Do
    listLevelSelect.AddItem "Level " & Trim(Str(intLevel))
    If intLevel >= intMaxLevelQnty Then Exit Do
Loop
intLevel = tmpintLevel
End Sub

Private Sub listLevelSelect_DblClick()
'Manual Level Select
    'User selected the label (list 0) which is just like cancel
    If listLevelSelect.ListIndex = 0 Then GoTo lblOne
intLevel = listLevelSelect.ListIndex 'Level selected
lblOne:
    'menu maintenance
Picture1.Enabled = True
pbxLevelSelect.Visible = False
mnuFile.Enabled = True
mnuEdit.Enabled = True
mnuSolution.Enabled = True
mnuLevel.Enabled = True
mnuHelp.Enabled = True
rtnInitialize
End Sub


Private Sub mnuSolutionSpeed0_Click()
'Set Solution Speed
sngSolutionDelay = 0
    'menu maintenance
mnuSolutionSpeed0.Checked = True
mnuSolutionSpeed1.Checked = False
mnuSolutionSpeed2.Checked = False
End Sub

Private Sub mnuSolutionSpeed1_Click()
'Set Solution Speed
sngSolutionDelay = 0.1
    'menu maintenance
mnuSolutionSpeed0.Checked = False
mnuSolutionSpeed1.Checked = True
mnuSolutionSpeed2.Checked = False
End Sub
Private Sub mnuSolutionSpeed2_Click()
'Set Solution Speed
sngSolutionDelay = 0.5
    'menu maintenance
mnuSolutionSpeed0.Checked = False
mnuSolutionSpeed1.Checked = False
mnuSolutionSpeed2.Checked = True
End Sub

Private Sub mnuSolutionStart_Click()
'Display Solution
Dim stLine As String, stChar As String
Dim x As Integer, y As Integer
Dim booSolutionFound As Boolean
Dim sngTimerAppend As Single
Dim tmpCount As Integer

    'menu maintenance
mnuFile.Enabled = False
mnuEdit.Enabled = False
mnuSolutionStart.Enabled = False
mnuSolutionStop.Enabled = True
mnuLevel.Enabled = False
mnuHelp.Enabled = False
'First initiate level because we may not be in the beginning
rtnInitialize



    'Find the correct solution. (moves)
    For x = 1 To intMasterFileSize
    stLine = strMasterArray(x)
        If stLine = "moves" & Trim(Str(intLevel)) Then
        tmpCount = x
        GoTo lblOne
        End If
    Next x
        intResponse = MsgBox("Solution was not found. Please check game file.", vbOKOnly, "Error!")
        Exit Sub
    
lblOne:
    'Solution found
    'Now tmpCount points to the label  "moves**"
    'the next line/lines will contain the solution
    'So let's process
    For y = 1 To 50 'Max possible number of solution lines
    stLine = strMasterArray(tmpCount + y)
        For x = 1 To Len(stLine)
        stChar = Mid$(stLine, x, 1) 'Get a solution keypress direction
     
            'if we find a zero that our que to exit, end of solution
            If stChar = "0" Then
            Close
            mnuSolutionStart.Enabled = True
            GoTo lblExitSub:
            End If
    
                'Find out which key/direction was used and goto respective procedure.
                'If little joe is carrying a block then send him to respective procedure also.
                Select Case stChar
                Case "1"
                If booWithBlock = False Then
                rtnLeftKey
                Else
                rtnLeftKeyWithBox
                End If
    
                Case "2"
                If booWithBlock = False Then
                rtnDownKey
                Else
                rtnDownKeyWithBox
                End If
    
                Case "3"
                If booWithBlock = False Then
                rtnRightKey
                Else
                rtnRightKeyWithBox
                End If
    
                Case "5"
                If booWithBlock = False Then
                rtnUpKey
                Else
                rtnUpKeyWithBox
                End If
                End Select


'Here is our delay according to user selected speed on solution menu
sngTimerAppend = Timer
Do While Timer < sngTimerAppend + sngSolutionDelay
DoEvents
Loop

Picture1.Refresh
DoEvents
    'If user selects stop solution then exit
    If mnuSolutionStart.Enabled = True Then
    Close
    GoTo lblExitSub:
    End If
Next x
Next y


'Reset all before leaving
lblExitSub:
    'menu maintenance
mnuSolutionStart.Enabled = True
mnuFile.Enabled = True
mnuEdit.Enabled = True
mnuEdit.Enabled = True
mnuSolutionStart.Enabled = True
mnuSolutionStop.Enabled = False
mnuLevel.Enabled = True
mnuHelp.Enabled = True
'Picture1.Enabled = True
Picture1.SetFocus
Exit Sub
End Sub

Private Sub mnuSolutionStop_Click()
'Stop Solution
    'menu maintenance
mnuSolutionStart.Enabled = True
mnuSolutionStop.Enabled = False
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
'OK, A keypress! Let's process this keypress!
Dim stChar As String
Dim dblTimerStart As Double
If mnuEditEditMode.Checked = True Then Exit Sub 'Exit if in edit mode
If mnuSolutionStart.Enabled = False Then Exit Sub 'Exit if playing solution
If booBusy = True Then Exit Sub 'Exit if busy falling or ball rolling
Picture1.Refresh

'Find out which key/direction was used and goto respective procedure.
'If little joe is carrying a block then send him to respective procedure also.
Select Case KeyCode
    
    Case vbKeyLeft
    stChar = "1"
    GoSub lblStoreMove: 'Increment move count and store move in case of a Menu Update Moves
    If booWithBlock = False Then
    booBusy = True
    rtnLeftKey
    booBusy = False
    Exit Sub
    Else
    booBusy = True
    rtnLeftKeyWithBox
    booBusy = False
    Exit Sub
    End If
    
    Case vbKeyDown
    stChar = "2"
    GoSub lblStoreMove: 'Increment move count and store move in case of a Menu Update Moves
    If booWithBlock = False Then
    booBusy = True
    rtnDownKey
    booBusy = False
    Exit Sub
    Else
    booBusy = True
    rtnDownKeyWithBox
    booBusy = False
    Exit Sub
    End If

    Case vbKeyRight
    stChar = "3"
    GoSub lblStoreMove: 'Increment move count and store move in case of a Menu Update Moves
    If booWithBlock = False Then
    booBusy = True
    rtnRightKey
    booBusy = False
    Exit Sub
    Else
    booBusy = True
    rtnRightKeyWithBox
    booBusy = False
    Exit Sub
    End If

    Case vbKeyUp
    stChar = "5"
    GoSub lblStoreMove: 'Increment move count and store move in case of a Menu Update Moves
    If booWithBlock = False Then
    booBusy = True
    rtnUpKey
    booBusy = False
    Exit Sub
    Else
    booBusy = True
    rtnUpKeyWithBox
    booBusy = False
    Exit Sub
    End If

End Select

Exit Sub

lblStoreMove:
'Increment move count and store move in case of a Menu Update Moves
intMoveCount = intMoveCount + 1
arrMoves(intMoveCount) = stChar
Return

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Must be in edit mode.
'This procedure places the
'current user selected object (joe, block, roll, etc)
'on the screen according to mouse cursor and left button click


If mnuEditEditMode.Checked = False Then Exit Sub 'Not on edit mode. Let's Exit

Dim NewObject As Integer
Dim PreviousLetter As String, NewLetter As String

    
   'Left button was pressed so update arrays etc.
    If Button = 1 Then
       PreviousLetter = CellCharactr(EditX, EditY) 'Letter designation for object previously in the location we wish to change
       NewObject = EditObject 'Temporary hold new object
                    
                    'Contents requested to be change
    Select Case PreviousLetter
       
       Case "T", "t", "J", "j", "X" 'Transport1,Transport2,JoeRight,JoeLeft,Exit
           'Do Nothing
           'You can not write over these.
           'These can only be changed by placing them elsewhere.
       
       Case " ", "B", "O", "=", "#", "." 'Empty,Box,Ball,Rail,Brick,RailWalkOnce
           'These can be changed so lets update arrays
          
          Select Case NewObject
            Case cEmpty
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = " " 'Update new location letter
            Case cBox
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "B" 'Update new location letter
            Case cRoll
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "O" 'Update new location letter
            Case cRail
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "=" 'Update new location letter
            Case cBrick
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "#" 'Update new location letter
            Case cRailWalkOnce
              CellContents(EditX, EditY) = NewObject 'Update new location object
              CellCharactr(EditX, EditY) = "." 'Update new location letter
            Case cTransport1
                  'Dummy numbers (100) are initial default values pointing
                  'to an impossible location. They
                  'indicate whether or not transports were
                  'already in use so as to update there previous location.
                  If Transport1X <> 100 And Transport1Y <> 100 Then
                  CellCharactr(Transport1X, Transport1Y) = " " 'Update previous location letter
                  CellContents(Transport1X, Transport1Y) = cEmpty 'Update previous location object
                  End If
              CellContents(EditX, EditY) = cEmpty 'Update new location object
              CellCharactr(EditX, EditY) = "T" 'Update new location letter
              Transport1X = EditX 'Update Transports x matrix value
              Transport1Y = EditY 'Update Transports y matrix value
              Form1.imgTransport(1).Visible = True 'Make Transport visible
              Form1.imgTransport(1).Left = xCorner(EditX, EditY) 'Place in proper location
              Form1.imgTransport(1).Top = yCorner(EditX, EditY) 'Place in proper location
            
            Case cTransport2
                  'Dummy numbers (100) are initial default values pointing
                  'to an impossible location. They
                  'indicate whether or not transports were
                  'already in use so as to update there previous location.
              If Transport2X <> 100 And Transport2Y <> 100 Then
              CellCharactr(Transport2X, Transport2Y) = " " 'Update previous location letter
              CellContents(Transport2X, Transport2Y) = cEmpty 'Update previous location object
              End If
              CellContents(EditX, EditY) = cEmpty 'Update new location object
              CellCharactr(EditX, EditY) = "t" 'Update new location letter
              Transport2X = EditX 'Update Transports x matrix value
              Transport2Y = EditY 'Update Transports y matrix value
              Form1.imgTransport(2).Visible = True 'Make Transport visible
              Form1.imgTransport(2).Left = xCorner(EditX, EditY) 'Place in proper location
              Form1.imgTransport(2).Top = yCorner(EditX, EditY) 'Place in proper location
            
            Case cExit
                  'Dummy numbers (100) are initial default values pointing
                  'to an impossible location. They
                  'indicate whether or not an exit was
                  'already in play area so as to update its previous location.
                  If ExitColumn <> 100 And ExitRow <> 100 Then
                  CellCharactr(ExitColumn, ExitRow) = " " 'Update previous location letter
                  CellContents(ExitColumn, ExitRow) = cEmpty 'Update previous location object
                  End If
              CellContents(EditX, EditY) = cEmpty 'Update new location object
              CellCharactr(EditX, EditY) = "X" 'Update new location letter
              ExitColumn = EditX 'Update Exit x matrix value
              ExitRow = EditY 'Update Exit y matrix value
              Form1.imgExit.Visible = True 'Make Exit visible
              Form1.imgExit.Left = xCorner(EditX, EditY) 'Place in proper location
              Form1.imgExit.Top = yCorner(EditX, EditY) 'Place in proper location
            
            Case cJoeRight
              CellContents(JoeX, JoeY) = cEmpty 'Update previous location object
              CellCharactr(JoeX, JoeY) = " " 'Update previous location letter
              CellContents(EditX, EditY) = cJoeRight 'Update new location object
              CellCharactr(EditX, EditY) = "J" 'Update new location letter
              JoeX = EditX 'Update Joe's x matrix value
              JoeY = EditY 'Update Joe's y matrix value
            
            Case cJoeLeft
              CellContents(JoeX, JoeY) = cEmpty 'Update previous location object
              CellCharactr(JoeX, JoeY) = " " 'Update previous location letter
              CellContents(EditX, EditY) = cJoeLeft 'Update new location object
              CellCharactr(EditX, EditY) = "j" 'Update new location letter
              JoeX = EditX 'Update Joe's x matrix value
              JoeY = EditY 'Update Joe's y matrix value
          
          End Select
    End Select
End If
'       empty = " " 1
'         box = "B" 2
'        roll = "O" 3
'        rail = "=" 4
'       brick = "#" 5
'   transport = "T" 6
'   transport = "t" 7
'     joeleft = "j" 8
'    joeright = "J" 9
'        exit = "X" 10
'railwalkonce = "." 11
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Must be in edit mode
'This procedure paints and clears
'the current user selected  object (joe, block, roll, etc)
'as you move the mouse across the screen.
'This is so user can see current object which will be placed and
'updated to screen and arrays if screen is left clicked.

If mnuEditEditMode.Checked = False Then Exit Sub 'Not on edit mode. Let's Exit



Dim XX As Integer, YY As Integer
'Set these variables to correspond with the mouse cursor
'and picture placement
Rem cell matrix is 20 by 15!
x = x + 15
y = y + 15
x = (x / 30): y = (y / 30)

'Just making sure you are in the play area
XX = x: YY = y
If XX > 20 Then XX = 20
If XX < 1 Then XX = 1
If YY > 15 Then YY = 15
If YY < 1 Then YY = 1
    
    'Repaint only if on another cell block to prevent redundancy
    If XX <> EditX Or YY <> EditY Then
    rtnRefreshFromCellArrays
    Picture1.PaintPicture picCell(EditObject), xCorner(XX, YY), yCorner(XX, YY)
    End If

'Update cell x, y pointers,  20 x 15
EditX = XX: EditY = YY
End Sub

Private Sub pbxWork_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'Edit mode only
'User clicked on one of the small storage picture boxes.
'If user right clicked then transfer main screen to clicked small picture box.
'If user left clicked then transfer clicked small picture box screen to main screen.

Dim XX As Integer, YY As Integer
Form1.Enabled = False 'Prevent unwanted events while busy

Select Case Button
    Case 1
    'Contents from small to main
    For YY = 1 To 15
    For XX = 1 To 20
    CellContents(XX, YY) = StoredCellContents(Index, XX, YY)
    CellCharactr(XX, YY) = StoredCellCharactr(Index, XX, YY)
    pbxWork(Index).PaintPicture picCellSmall(CellContents(XX, YY)), xCorner(XX, YY), yCorner(XX, YY)
    Next XX
    Next YY
    rtnRefreshFromCellArrays
    rtnAllRollsAndBoxesFall
    
    Case 2
    'Contents from main to small
    'Paint screen from arrays
    'Blocks x = 0 to 19    y = 0 to 14    20x15 Blocks   Total 300 Blocks
    If Index = 0 Then GoTo lblExit
    For YY = 1 To 15
    For XX = 1 To 20
    StoredCellContents(Index, XX, YY) = CellContents(XX, YY)
    StoredCellCharactr(Index, XX, YY) = CellCharactr(XX, YY)
    pbxWork(Index).PaintPicture picCellSmall(CellContents(XX, YY)), xCorner(XX, YY), yCorner(XX, YY)
    Next XX
    Next YY
    Form1.pbxWork(Index).Refresh
lblExit:
End Select
Form1.Enabled = True
End Sub

Private Sub Image1_Click(Index As Integer)
EditObject = Index 'Object
dummy = fncPlayFx(cWalk) 'Just to let us know another object was selectrd
End Sub

Private Sub cmdEditUpDownStore_Click(Index As Integer)
'Edit mode only
'User just clicked on Level up/down or on Update (the freshly user editted level to the current level in memory)
Dim tmpText As String, tmpCount As Integer
Dim tmpMasterPointer As Integer
Dim tmpOldFileSize As Integer
Dim tmpLastLevelsLastLine As Integer
Dim tmpArray() As String
Dim x As Integer, y As Integer
Select Case Index
    Case 0 'store editted level to game memory
        tmpText = ""
        For y = 1 To 15       'y cells
            For x = 1 To 20   'x cells
            tmpText = tmpText & CellCharactr(x, y)
            Next x
         '((point to previous level's last line) + 1 + y) =
            strMasterArray(((intLevel - 1) * 16) + 1 + y) = tmpText
            tmpText = ""
        Next y
    
    Case 1 'down
    'Go Down One Level
    intLevel = intLevel - 1
    Label1.Caption = "Update level " & intLevel & " in memory"
        If intLevel = 0 Then
        intLevel = intLevel + 1
        Label1.Caption = "Update level " & intLevel & " in memory"
        Exit Sub
        Label1.Caption = "Update level " & intLevel & " in memory"
        End If

    Case 2 'Up
    'Go Up One Level
    intLevel = intLevel + 1
    If intLevel > intMaxLevelQnty Then intLevel = intLevel - 1
    rtnCheckIfLevelAvailable
        If booLevelAvailable = True Then
        'Level found so load it
        rtnInitialize
        Label1.Caption = "Update level " & intLevel & " in memory"
        Exit Sub
    
        'There wasn't a higher level so make new "empty" one

        Else
        'Load variables
        tmpOldFileSize = intMasterFileSize
        intMasterFileSize = intMasterFileSize + 18
        tmpLastLevelsLastLine = (intLevel - 1) * 16
        ReDim tmpArray(tmpOldFileSize)
        'Copy master array and Redim to new size
        For x = 1 To tmpOldFileSize
        tmpArray(x) = strMasterArray(x)
        Next x
        ReDim strMasterArray(intMasterFileSize)
        'Transfer existing levels
        tmpMasterPointer = 0
        For x = 1 To tmpLastLevelsLastLine
        tmpMasterPointer = tmpMasterPointer + 1
        strMasterArray(tmpMasterPointer) = tmpArray(x)
        Next x
        'Add the new level
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "level" & Trim(Str(intLevel))
        'Add new level
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "####################"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#                  #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "#     J            #"
         tmpMasterPointer = tmpMasterPointer + 1: strMasterArray(tmpMasterPointer) = "####################"
        'Transfer existing solutions
        For x = tmpLastLevelsLastLine + 1 To tmpOldFileSize
        tmpMasterPointer = tmpMasterPointer + 1
        strMasterArray(tmpMasterPointer) = tmpArray(x)
        Next x
        'Add new solution with no moves, "0" only
        tmpMasterPointer = tmpMasterPointer + 1
        strMasterArray(tmpMasterPointer) = "moves" & Trim(Str(intLevel))
        tmpMasterPointer = tmpMasterPointer + 1
        strMasterArray(tmpMasterPointer) = "0"
        'Flash "New Level Added"
        Label1.Caption = "Update level " & intLevel & " in memory"
        Label2.BackColor = vbGreen
        Label2.Caption = "New Level Added!"
        Timer1.Enabled = True
        End If
End Select
rtnInitialize
End Sub

Private Sub Timer1_Timer()
'Turn off the label "New Level Added"
Label2.BackColor = vbButtonFace
Label2.Caption = ""
Timer1.Enabled = False
End Sub


Private Sub Timer2_Timer()
'Repeat background music when finished playing
    Dim MCIStatusLen As Integer
    Dim MCIStatus As String
    
    ' check status of background music
    If IsMusicOn = True Then
        ' see if the music is still playing
        MCIStatusLen = 15
        MCIStatus = String(MCIStatusLen + 1, " ")
        RetValue = mciSendString("STATUS BackgroundMusic MODE", MCIStatus, MCIStatusLen, 0)
        If UCase(Left$(MCIStatus, 7)) = "STOPPED" Then
            ' restart music from the beginning again
            RetValue = mciSendString("PLAY BackgroundMusic FROM 0", "", 0, 0)
        End If
    End If
End Sub

Private Sub tmrFinale_Timer()
'User finished all levels so go to finale
tmrFinale.Enabled = False
rtnFinale
cmdEnd.Visible = True
Form1.Enabled = True
End Sub
