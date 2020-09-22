VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Millionaire Information"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   Icon            =   "MillFrmHelp.frx":0000
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
Option Explicit

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
  CmdPre.Visible = True
  CmdNext.Visible = True
  Call Display_Help
End Sub

Public Sub Display_Help()
Dim mtx As String
mtx = ""
If Disp = 1 Then
  CmdPre.Visible = False
  Lblno = 1
  mtx = "Who Wants to be a Millionaire is, as you might expect, based on the TV "
  mtx = mtx & "programme of the same name. "
  mtx = mtx & "You are presented with a question and four possible answers, "
  mtx = mtx & "only one of which is correct. The object of the game is to try to win £1,000,000 by answering "
  mtx = mtx & "15 questions correctly."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "To help you on your way you have three lifelines, 50/50, Phone a Friend and Ask the Audience. "
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "The 50/50 option is completely random and will take away two wrong answers, leaving the correct "
  mtx = mtx & "answer and the one remaining wrong answer."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "The Phone a Friend option randomly selects one of four 'Friends', who will attempt to help you "
  mtx = mtx & "answer the question correctly."
  LblInfo = mtx
End If

If Disp = 2 Then
  CmdPre.Visible = True
  CmdNext.Visible = True
  Lblno = 2
  
  mtx = mtx & "Obviously the harder the questions get the chance of the friend knowing the correct answer diminishes. "
  mtx = mtx & "The four friends, John, Martin, Julie and Mary have different levels of intelligence. "
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "As you continue to play the game you will realise which of the 'friends' is the most, and least reliable. "
  mtx = mtx & " The answers offered by your selected 'friend' are not always going to be correct. "
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "The Ask the Audience option is also based on an intelligence factor, which decreases as the question values increase. "
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "During the early part of the game the Ask the Audience option should be very helpful, but the percentages will level "
  mtx = mtx & "off as questions with a value of between £32,000 and £1,000,000 appear. "
  LblInfo = mtx
End If

If Disp = 4 Then
  CmdNext.Visible = False
  Lblno = 3
  mtx = mtx & "This version contains a database of 4,000 different questions and answers. "
  mtx = mtx & "The programme also creates a shuffle file which is updated to prevent repetition of questions. "
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "I have made every effort to ensure that all 4,000 questions are different, that the answers are correct, and that all questions and answers have been "
  mtx = mtx & "spell checked."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "However I cannot be held responsible for any question that you may think has the wrong answer."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "If you do suspect that an answer to a question is wrong I would like to hear from you. Also if you detect any "
  mtx = mtx & "typo errors please could you report them to me. "
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "My email address is davek606@aol.com"
  
  LblInfo = mtx
End If

If Disp = 3 Then
  CmdNext.Visible = False
  Lblno = 3
  
  mtx = mtx & "Although this version only contains 109 questions the programme creates a shuffle file which is updated to prevent repetition of questions. "
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "A much more extensive version which contains a database of 4,000 questions and answers is available for just £5. "
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "Obviously with 4,000 different questions the larger version presents a much more challenging quiz, and will certainly increase your general knowledge considerably."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "If you are interested in obtaining the larger version please email me at davek606@aol.com "
  mtx = mtx & "and I will send you the necessary information."
  mtx = mtx & vbCr & vbLf & vbLf
  mtx = mtx & "I hope you enjoy playing the game..."
  LblInfo = mtx
End If


End Sub
