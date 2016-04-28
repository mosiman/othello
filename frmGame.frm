VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello"
   ClientHeight    =   7380
   ClientLeft      =   3975
   ClientTop       =   1995
   ClientWidth     =   11175
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11175
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   6120
   End
   Begin MSFlexGridLib.MSFlexGrid grdBoard 
      Height          =   5145
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9075
      _Version        =   393216
      Rows            =   10
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
   End
   Begin VB.Label lblGameTimer 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "0:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label lblP1Name 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Points:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Points:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblPoints2 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblPoints1 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgP2CellPic 
      Height          =   495
      Left            =   12120
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image imgP1CellPic 
      Height          =   495
      Left            =   12720
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblP2Name 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Image imgPlayer2 
      Height          =   975
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Image imgPlayer1 
      Height          =   975
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblTurn 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Shape shpRectangle 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   1
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Shape shpRectangle 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   2
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Image imgBackground 
      Height          =   7455
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuHighScore 
         Caption         =   "View Leaderboards"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Const GRDSIZE = 10
Dim gameRestart As Integer 'Keeps track of whether or not to display messagebox.
Dim gameExit As Integer
Dim gameTimer As Integer
Dim gameWinner As Integer
Dim theBoard(0 To GRDSIZE - 1, 0 To GRDSIZE - 1) As Integer '2D array interpretation of board.
Dim playArr(0 To GRDSIZE - 1, 0 To GRDSIZE - 1) As Boolean 'Keeps track of possibly playable pieces.
Dim currPlayer As Integer 'keeps track of current player.

Private Sub Form_Load()
    Dim isPlayable As Boolean
    Dim K As Integer, X As Integer
    
    Randomize
    
    'Initialize a lot of stuff.
    
    isPlayable = False
    gameTimer = 0
    lblGameTimer.Visible = False
    tmrTimer.Enabled = False
    
    If Player(2).isComputer < 1 Then 'Enable/disable high score viewing in 2P mode.
        mnuHighScore.Enabled = False
    Else
        mnuHighScore.Enabled = True
    End If
    
    'Set / display pictures and names from PlayerSelect form.
    
    imgPlayer1.Picture = Player(1).Avatar
    imgPlayer2.Picture = Player(2).Avatar
    lblP1Name.Caption = Player(1).name
    lblP2Name.Caption = Player(2).name
    
    imgP1CellPic.Width = frmGame.TextWidth("XXXXx..")
    imgP1CellPic.Height = frmGame.TextWidth("XXXX...")
    imgP2CellPic.Width = imgP1CellPic.Width
    imgP2CellPic.Height = imgP2CellPic.Height
    
    imgP1CellPic.Picture = Player(1).Avatar
    imgP2CellPic.Picture = Player(2).Avatar

    'Set the current player and initial score.
    
    currPlayer = 1
    Player(1).points = 2
    Player(2).points = 2
    
    'Initialize 2D game board array.
    
    For K = 0 To GRDSIZE - 1
        For X = 0 To GRDSIZE - 1
            theBoard(K, X) = 0
        Next X
    Next K
    
    grdBoard.Clear
    grdBoard.Enabled = True
    
    For K = 0 To GRDSIZE - 1
        grdBoard.ColWidth(K) = frmGame.TextWidth("XXXXx..")
        grdBoard.RowHeight(K) = frmGame.TextWidth("XXXX..")
    Next K
    
    'Place starting pieces.
    
    boardPlace grdBoard, theBoard(), 1, 4, 4
    boardPlace grdBoard, theBoard(), 1, 5, 5
    boardPlace grdBoard, theBoard(), 2, 4, 5
    boardPlace grdBoard, theBoard(), 2, 5, 4
    
    updateBoard theBoard(), frmGame, isPlayable, currPlayer, gameTimer, gameWinner
End Sub



Private Sub grdBoard_Click()
    Dim K As Integer
    Dim validMove As Boolean
    Dim startNextTurn As Boolean
    Dim start As Boolean
    Dim isPlayable As Boolean
    Dim pointsGrabbed As Integer
    Dim firstPass As Boolean
    
    firstPass = True
    isPlayable = False
    pointsGrabbed = 0
    
    'Replace appropriate enemy pieces if appropriate.
    
    If theBoard(grdBoard.Col, grdBoard.Row) = 0 Then
        
        'Search all directions for valid moves.
        
        For K = 1 To 8
            validMove = False
            start = True
            searchReplace theBoard(), currPlayer, validMove, grdBoard.Col, grdBoard.Row, K, start
            If validMove = True Then
                startNextTurn = True
            End If
            
        Next K
    End If

    
    ' Switch players, update board, count points if move is valid.
    
    If startNextTurn Then
    
        'Start timer if it has not been started and it is 1P mode.
    
        If gameTimer = 0 And Player(2).isComputer > 0 Then
            lblGameTimer.Visible = True
            tmrTimer.Enabled = True
        End If
        
        nextTurn currPlayer
        updateBoard theBoard(), frmGame, isPlayable, currPlayer, gameTimer, gameWinner
        
        'If there are still valid moves, commence CPU turn (if 1P mode).
        
        If Player(currPlayer).isComputer > 0 And isPlayable Then
            ComputerTurn theBoard(), currPlayer, isPlayable
            nextTurn currPlayer
            updateBoard theBoard(), frmGame, isPlayable, currPlayer, gameTimer, gameWinner
        End If
        
        If grdBoard.Enabled = False Then
            gameRestart = MsgBox("Do you wish to play again?", vbQuestion + vbYesNo, "Game Over")
            
            If gameRestart = vbYes Then
                Form_Load
            Else
                Load frmPlayerSelect
                frmPlayerSelect.Show
                Unload Me
            End If
        End If
                
    End If
    
    
    

End Sub


Private Sub mnuAbout_Click()
    'Modally display About form.
    Load frmAbout
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
'Exits the current game and returns user to the main / option screen.
    gameExit = MsgBox("Are you sure you want to quit?", vbYesNo + vbQuestion, "Exit Game?")
    
    If gameExit = vbYes Then
        Load frmPlayerSelect
        frmPlayerSelect.Show
        Unload Me
    End If
End Sub

Private Sub mnuHighScore_Click()
'Displays the leaderboard.
    Load frmHighScores
    frmHighScores.Show vbModal
End Sub

Private Sub mnuNewGame_Click()
'Starts a new game.
    Dim UserIn As Integer
    
    UserIn = MsgBox("Are you sure you want to start a new game? All progress will be lost.", vbExclamation + vbYesNo, "New Game")
    
    If UserIn = vbYes Then
        Form_Load
    End If
End Sub

Private Sub mnuOptions_Click()
'Exits the current game and returns user to the main / option screen.
    Dim UserIn As Integer
    
    UserIn = MsgBox("Are you sure you want to return to the options screen? All progress will be lost.", vbExclamation + vbYesNo, "Back To Options Screen")
    
    If UserIn = vbYes Then
        Load frmPlayerSelect
        frmPlayerSelect.Show
        Unload Me
    End If
End Sub


Private Sub tmrTimer_Timer()
    Dim TimeConv As String
    gameTimer = gameTimer + 1
    
    TimeConv = Str$(gameTimer \ 60) & ":" & Format$(Str$(gameTimer Mod 60), "00") 'Converts seconds to m:ss.
    
    lblGameTimer.Caption = TimeConv
    
    If gameTimer = 7201 Then
        tmrTimer.Enabled = False
        grdBoard.Enabled = False
        
        MsgBox "This game has taken too long (> 2 hours). No winner is decided."
    End If
End Sub
