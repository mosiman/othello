VERSION 5.00
Begin VB.Form frmPlayerSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Select"
   ClientHeight    =   7695
   ClientLeft      =   3975
   ClientTop       =   2220
   ClientWidth     =   11355
   Icon            =   "frmPlayerSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11355
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9960
      TabIndex        =   16
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame fraComputer 
      Caption         =   "Computer"
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   8175
      Begin VB.OptionButton optHardBot 
         Caption         =   "Hard"
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton optEasyBot 
         Caption         =   "Easy"
         Height          =   495
         Left            =   480
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Image imgHardBot 
         Height          =   1200
         Left            =   2520
         Picture         =   "frmPlayerSelect.frx":0442
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblBotName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   17
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Image imgEasyBot 
         Height          =   1200
         Left            =   2520
         Picture         =   "frmPlayerSelect.frx":0884
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1200
      End
   End
   Begin VB.Frame fraBackgrounds 
      Caption         =   "Background Image"
      Height          =   4455
      Left            =   8400
      TabIndex        =   12
      Top             =   2520
      Width           =   2775
      Begin VB.Image imgBackground 
         Height          =   975
         Index           =   2
         Left            =   720
         Picture         =   "frmPlayerSelect.frx":0CC6
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Image imgBackground 
         Height          =   975
         Index           =   1
         Left            =   720
         Picture         =   "frmPlayerSelect.frx":16C70
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image imgBackground 
         Height          =   975
         Index           =   0
         Left            =   720
         Picture         =   "frmPlayerSelect.frx":58B62
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Play"
      Height          =   495
      Left            =   8400
      TabIndex        =   9
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame fraGameMode 
      Caption         =   "Game Mode"
      Height          =   2175
      Left            =   8400
      TabIndex        =   4
      Top             =   240
      Width           =   2775
      Begin VB.OptionButton optSinglePlayer 
         Caption         =   "Single Player (Against Computer)"
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optTwoPlayer 
         Caption         =   "Two Player"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraPlayerTwo 
      Caption         =   "Player 2"
      ClipControls    =   0   'False
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   8175
      Begin VB.TextBox txtP2Name 
         Height          =   495
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblP2AvatarName 
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Image imgP2Avatar 
         Height          =   1095
         Index           =   3
         Left            =   5040
         Picture         =   "frmPlayerSelect.frx":355BA4
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Image imgP2Avatar 
         Height          =   1095
         Index           =   2
         Left            =   3480
         Picture         =   "frmPlayerSelect.frx":355EAE
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Image imgP2Avatar 
         Height          =   1095
         Index           =   1
         Left            =   1800
         Picture         =   "frmPlayerSelect.frx":3561B8
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Image imgP2Avatar 
         Height          =   1095
         Index           =   0
         Left            =   240
         Picture         =   "frmPlayerSelect.frx":3564C2
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   495
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraPlayerOne 
      Caption         =   "Player 1"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.TextBox txtP1Name 
         Height          =   495
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblP1AvatarName 
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image imgP1Avatar 
         Height          =   1095
         Index           =   3
         Left            =   5040
         Picture         =   "frmPlayerSelect.frx":3567CC
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image imgP1Avatar 
         Height          =   1095
         Index           =   2
         Left            =   3360
         Picture         =   "frmPlayerSelect.frx":356AD6
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image imgP1Avatar 
         Height          =   1095
         Index           =   1
         Left            =   1800
         Picture         =   "frmPlayerSelect.frx":356DE0
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image imgP1Avatar 
         Height          =   1080
         Index           =   0
         Left            =   240
         Picture         =   "frmPlayerSelect.frx":3570EA
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuHighScores 
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
Attribute VB_Name = "frmPlayerSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Avatar1 As Integer
Dim Avatar2 As Integer
Dim selectedBackground As New StdPicture

Private Sub cmdExit_Click()
    Dim UserIn As Integer
    
    UserIn = MsgBox("Are you sure you want to exit?", vbInformation + vbYesNo, "Are you sure?")
    
    If UserIn = vbYes Then
        Beep
        End
    End If
    
End Sub

Private Sub cmdStartt_Click()
    If optTwoPlayer.Value = True Then
        If txtP1Name.Text <> "" And txtP2Name.Text <> "" Then
            If Player(1).Avatar <> 0 And Player(2).Avatar <> 0 Then
                If selectedBackground <> 0 Then
                    If UCase$(txtP1Name.Text) <> UCase$(txtP2Name.Text) Then
                
                        Player(1).name = txtP1Name.Text
                        
                        If Player(2).isComputer = 0 Then
                            Player(2).name = txtP2Name.Text
                        Else
                            Player(2).name = lblBotName.Caption
                        End If
    
                        Load frmGame
                        Set frmGame.imgBackground.Picture = selectedBackground
                        frmGame.Show
                    
                        Unload Me
                    Else
                        MsgBox "Both players cannot share the same name", vbExclamation, "Change Names"
                    End If
                Else
                    MsgBox "Please select a background!", vbExclamation, "Select Background"
                End If
            Else
                MsgBox "Please select avatars for both players!", vbExclamation, "Select Avatar"
            End If
        Else
            MsgBox "Please enter names for both players!", vbExclamation, "Enter Name"
        End If
    End If
    
End Sub

Private Sub cmdStart_Click()
    If selectedBackground <> 0 Then
        If optTwoPlayer.Value = True Then
            If (UCase$(Trim$(txtP1Name.Text)) <> UCase$(Trim$(txtP2Name.Text))) And (txtP1Name.Text <> "" And txtP2Name.Text <> "") Then
                If Player(1).Avatar <> 0 And Player(2).Avatar <> 0 Then
                    Player(1).name = txtP1Name.Text
                    Player(2).name = txtP2Name.Text
                    
                    Load frmGame
                    Set frmGame.imgBackground.Picture = selectedBackground
                    frmGame.Show
                
                    Unload Me
                Else
                    MsgBox "Please select avatars for both players!", vbExclamation, "Select Avatars"
                End If
            Else
                MsgBox "Player names should not be blank and must be unique!", vbExclamation, "Invalid Names"
            End If
        Else
            If (UCase$(Trim$(txtP1Name.Text)) <> "EASY BOT" And UCase$(Trim$(txtP1Name.Text)) <> "HARD BOT") And Trim$(txtP1Name.Text) <> "" Then
                If Player(1).Avatar <> 0 Then
                    Player(1).name = txtP1Name.Text
                    Player(2).name = lblBotName.Caption
                    
                    If Player(2).isComputer = 1 Then
                        Set Player(2).Avatar = imgEasyBot.Picture
                    Else
                        Set Player(2).Avatar = imgHardBot.Picture
                    End If
                    
                    Load frmGame
                    Set frmGame.imgBackground.Picture = selectedBackground
                    frmGame.Show
                
                    Unload Me
                Else
                    MsgBox "Please select an avatar!", vbExclamation, "Select Avatar"
                End If
            Else
                MsgBox "You cannot name yourself 'Easy Bot', 'Hard Bot', or blank!", vbExclamation, "Invalid Name"
            End If
        End If
    Else
        MsgBox "Please select a background image!", vbExclamation, "Select Background"
    End If
                
End Sub

Private Sub Form_Load()
    Dim K As Integer
    
    'Initialize Player type.
    For K = 1 To 2
        Player(K).isComputer = 0 ' because two player is enabled by default.
        Set Player(K).Avatar = Nothing
        Player(K).name = ""
        Player(K).points = 0
    Next K
        
    
End Sub


Private Sub imgBackground_Click(index As Integer)
    Dim K As Integer
    
    Set selectedBackground = imgBackground(index).Picture
    
    For K = 0 To imgBackground.ubound
        imgBackground(K).BorderStyle = 0
    Next K
    
    imgBackground(index).BorderStyle = 1
    
        
End Sub

'Check if image selected
Private Sub imgP1Avatar_Click(index As Integer)
    Dim Msg As String
    Dim K As Integer
    
    For K = 0 To 3
        imgP1Avatar(K).BorderStyle = 0
        imgP2Avatar(K).Enabled = True
    Next K
    imgP1Avatar(index).BorderStyle = 1
    
    imgP2Avatar(index).Enabled = False
    
    Set Player(1).Avatar = imgP1Avatar(index).Picture
        
End Sub

Private Sub imgP2Avatar_Click(index As Integer)
    Dim Msg As String
    Dim K As Integer
    
    For K = 0 To 3
        imgP2Avatar(K).BorderStyle = 0
        imgP1Avatar(K).Enabled = True
    Next K
    imgP2Avatar(index).BorderStyle = 1
    
    imgP1Avatar(index).Enabled = False
    
    Set Player(2).Avatar = imgP2Avatar(index).Picture
End Sub

Private Sub mnuExit_Click()
    Dim gameExit As Integer
    
    gameExit = MsgBox("Are you sure you want to quit?", vbYesNo + vbQuestion, "Exit Game?")
    
    If gameExit = vbYes Then
        Beep
        End
    End If
End Sub

Private Sub mnuHighScores_Click()
    Load frmHighScores
    frmHighScores.Show vbModal
End Sub

Private Sub optEasyBot_Click()
    imgEasyBot.Visible = True
    imgHardBot.Visible = False
    lblBotName.Caption = "Easy Bot"
    Player(2).isComputer = 1
End Sub

Private Sub optHardBot_Click()
    Dim K As Integer
    
    imgHardBot.Visible = True
    imgEasyBot.Visible = False
    lblBotName.Caption = "Hard Bot"
    Player(2).isComputer = 2
    
    For K = 0 To 3
        If imgP2Avatar(K).BorderStyle = 1 Then
            imgP1Avatar(K).Enabled = False
        End If
    Next K
    
End Sub

Private Sub optSinglePlayer_Click()
    Dim K As Integer
    
    fraComputer.Visible = True
    fraPlayerTwo.Visible = False
    optEasyBot.Value = True
    imgEasyBot.Visible = True
    lblBotName.Caption = "Easy Bot"
    Player(2).isComputer = 1
    
    For K = 0 To 3
        imgP1Avatar(K).Enabled = True
    Next K
End Sub

Private Sub optTwoPlayer_Click()
    Dim K As Integer
    
    fraComputer.Visible = False
    fraPlayerTwo.Visible = True
    Player(2).isComputer = 0
    
    For K = 0 To 3
        If imgP2Avatar(K).BorderStyle = 1 Then
            imgP1Avatar(K).Enabled = False
        End If
    Next K
End Sub
