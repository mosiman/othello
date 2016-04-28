VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   6150
   ClientTop       =   3540
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmeTimer 
      Interval        =   1000
      Left            =   480
      Top             =   3600
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Dillon Chan"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   4
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Image imgLogo 
         Height          =   1665
         Left            =   360
         Picture         =   "frmSplash.frx":0442
         Stretch         =   -1  'True
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 2.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5580
         TabIndex        =   1
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Othello"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   2
         Top             =   1140
         Width           =   2205
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timerCount As Integer
Option Explicit

Private Sub Form_Load()
    timerCount = 0
End Sub

Private Sub lblPlatform_Click()

End Sub

Private Sub tmeTimer_Timer()
    timerCount = timerCount + 1
    If timerCount = 2 Then
        Load frmPlayerSelect
        frmPlayerSelect.Show
        Unload Me
    End If
End Sub
