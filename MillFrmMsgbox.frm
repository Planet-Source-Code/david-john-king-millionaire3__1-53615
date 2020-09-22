VERSION 5.00
Begin VB.Form FrmMsgbox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MillFrmMsgbox.frx":0000
   ScaleHeight     =   1440
   ScaleWidth      =   3510
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Height          =   340
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton CmdNo 
      BackColor       =   &H008080FF&
      Caption         =   "No"
      Height          =   340
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton CmdYes 
      BackColor       =   &H0080FF80&
      Caption         =   "Yes"
      Height          =   340
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.Label Lblmsg 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Imatype 
      Height          =   540
      Left            =   120
      Picture         =   "MillFrmMsgbox.frx":0721
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "FrmMsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdNo_Click()
  retval = 0
  FrmMsgbox.Hide
End Sub

Private Sub CmdOK_Click()
  retval = 0
  FrmMsgbox.Hide
End Sub

Private Sub CmdYes_Click()
  retval = 6
  FrmMsgbox.Hide
End Sub

Private Sub Form_Activate()
  Me.Left = (FrmBoard2.Left + (FrmBoard2.Width / 2)) - (FrmMsgbox.Width / 2)
  Me.Top = (FrmBoard2.Top + (FrmBoard2.Height / 2)) - (FrmMsgbox.Height / 2)
  If cQ = 16 Then Me.Top = Me.Top + 1920
End Sub
