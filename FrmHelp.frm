VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cascading Words Help"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Next  >"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton CmdPre 
      Caption         =   "<  &Previous"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label LblInfo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Lblno 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   4580
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   4440
      Y2              =   4440
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClose_Click()
  FrmHelp.Hide
End Sub

Private Sub CmdNext_Click()
  Disp = Disp + 1
  Call Display_Help
End Sub

Private Sub CmdPre_Click()
  If Disp > 1 Then Disp = Disp - 1
  Call Display_Help
End Sub

Private Sub Form_Activate()
  Disp = 1
  Call Display_Help
End Sub

Public Sub Display_Help()
Dim mtx As String
mtx = ""
If Disp = 1 Then
  CmdPre.Visible = False
  Lblno = 1
  mtx = "Cascading words is a game designed for people who can already type,"
  mtx = mtx & " but who wish to increase their typing speed, and at the"
  mtx = mtx & " same time improve their spelling."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "Once a new game has been selected, words appear in boxes at"
  mtx = mtx & " the top of the screen and cascade down towards the bottom blue line."
  mtx = mtx & " As well as normal words the boxes may contain a mixture of"
  mtx = mtx & " numbers, punctuation and abbreviations."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "There are 4 levels of play, the level chosen dictates the"
  mtx = mtx & " speed the words cascade downwards and the maximum length of the words."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "Once the box containing the word reaches the blue line it"
  mtx = mtx & " stops and any other box will then stop above this box."
  mtx = mtx & " Eventually the boxes will build up towards the top"
  mtx = mtx & " of the screen."
  LblInfo = mtx
End If

If Disp = 2 Then
  CmdPre.Visible = True
  CmdNext.Visible = True
  Lblno = 2
  mtx = mtx & "Once two words appear on the top line and cannot cascade"
  mtx = mtx & " downwards, due to other boxes below, the game ends."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "As soon as the words appear they can be eliminated from the screen"
  mtx = mtx & " by typing the word exactly as it appears in the box plus the space bar."
  mtx = mtx & " As the boxes build up the skill is not only typing words quickly"
  mtx = mtx & " to eliminate them, but choosing a word which will allow the"
  mtx = mtx & " greatest number of other words above to carry on cascading downwards."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "At the end of the game your score is the number of completed words typed."
  mtx = mtx & " You are also given a report listing the time taken,"
  mtx = mtx & " the number of characters typed, (this includes any letters typed as"
  mtx = mtx & " part of an incomplete word) the number of errors and your words per minute,"
  mtx = mtx & " (based on 5 characters per word, including spaces)"
  
  LblInfo = mtx
End If

If Disp = 3 Then
  CmdNext.Visible = False
  Lblno = 3
  mtx = mtx & "If errors have occurred these can be displayed by clicking on the"
  mtx = mtx & " Show Errors button."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "There are four levels of play and the top three scores for each"
  mtx = mtx & " level are recorded in a high score table. If you obtain"
  mtx = mtx & " a score that qualifies you will be asked to enter your name."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "The speed of the game will increase after each 100 words have appeared."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "If any character is typed incorrectly while entering your word the"
  mtx = mtx & " Backspace key can be used to delete it."
  mtx = mtx & vbCr & vbLf & vbLf & vbLf
  mtx = mtx & "To exit the game select Game - Exit from the Game menu. This will ensure"
  mtx = mtx & " that any changes to the high score table are saved."
  LblInfo = mtx
End If
End Sub
