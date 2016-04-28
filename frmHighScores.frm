VERSION 5.00
Begin VB.Form frmHighScores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   4320
   ClientLeft      =   5940
   ClientTop       =   3660
   ClientWidth     =   7890
   Icon            =   "frmHighScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7890
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Const NUMRECS = 6
    Dim K As Integer
    Dim TimeConv As String
    
    ReadFileREC HighScores(), App.Path & "\highscore.rec", lenHighScore, NUMRECS
    
    picDisplay.Print "High Scores"
    picDisplay.Print
    
    picDisplay.Print "Name", "", Spc(1); "Time", "Score"
    
    For K = 1 To 51
        picDisplay.Print "-";
    Next K
    picDisplay.Print "-"
    
    For K = 1 To 5
        If HighScores(K).intTime < 32767 Then
            TimeConv = Str$(HighScores(K).intTime \ 60) & ":" & Format$(Str$(HighScores(K).intTime Mod 60), "00") ' Converts seconds to mm:ss
            picDisplay.Print Trim$(HighScores(K).name), "", TimeConv, Str$(HighScores(K).points)
        End If
    Next K
End Sub

