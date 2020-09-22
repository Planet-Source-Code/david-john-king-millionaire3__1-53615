VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MillFrmAbout.frx":0000
   ScaleHeight     =   1545
   ScaleWidth      =   3675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.Label Lblinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   795
      TabIndex        =   1
      Top             =   190
      Width           =   2655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOK_Click()
  FrmAbout.Hide
End Sub

