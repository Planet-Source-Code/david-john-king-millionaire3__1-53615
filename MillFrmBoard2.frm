VERSION 5.00
Begin VB.Form FrmBoard2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Who Wants to be a Millionaire - By D J King"
   ClientHeight    =   7470
   ClientLeft      =   6045
   ClientTop       =   630
   ClientWidth     =   9285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   Icon            =   "MillFrmBoard2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MillFrmBoard2.frx":0ECA
   ScaleHeight     =   7470
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   1920
      Pattern         =   "*.mil*"
      TabIndex        =   414
      Top             =   2160
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Cmdfook 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   413
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Tmr003 
      Enabled         =   0   'False
      Interval        =   130
      Left            =   7560
      Top             =   7200
   End
   Begin VB.CommandButton Cmdexit3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6840
      TabIndex        =   412
      Top             =   6720
      Width           =   735
   End
   Begin VB.Timer Tmr002 
      Enabled         =   0   'False
      Left            =   8160
      Top             =   7200
   End
   Begin VB.Timer Tmr001 
      Enabled         =   0   'False
      Interval        =   1400
      Left            =   8760
      Top             =   7200
   End
   Begin VB.Label Lblfoh 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Select Question File to Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   416
      Top             =   1680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Lblfileopen 
      BackColor       =   &H00C0FFC0&
      Height          =   3105
      Left            =   1800
      TabIndex        =   415
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Image Imaincorrect 
      Height          =   1020
      Left            =   1920
      Picture         =   "MillFrmBoard2.frx":130CB
      Top             =   2640
      Width           =   3120
   End
   Begin VB.Image Imacorrect 
      Height          =   1020
      Left            =   2040
      Picture         =   "MillFrmBoard2.frx":1454C
      Top             =   1200
      Width           =   3120
   End
   Begin VB.Image ImaWin 
      Height          =   2805
      Left            =   2160
      Picture         =   "MillFrmBoard2.frx":156E9
      Top             =   1200
      Width           =   4530
   End
   Begin VB.Image Imano 
      Height          =   495
      Left            =   4800
      Picture         =   "MillFrmBoard2.frx":1A3FA
      Top             =   3840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Imayes 
      Height          =   495
      Left            =   3600
      Picture         =   "MillFrmBoard2.frx":1A986
      Top             =   3840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Imafa 
      Height          =   450
      Left            =   2640
      Picture         =   "MillFrmBoard2.frx":1AF17
      Top             =   3360
      Width           =   4020
   End
   Begin VB.Image ImaWW 
      Height          =   375
      Left            =   8040
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label LblQuit 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Take the Money"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3720
      TabIndex        =   411
      Top             =   6860
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Imapafper 
      Height          =   1095
      Left            =   2400
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Lblpafin 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1455
      Left            =   7200
      TabIndex        =   410
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Lblpafout 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   6840
      TabIndex        =   409
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   99
      Left            =   0
      TabIndex        =   408
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   98
      Left            =   0
      TabIndex        =   407
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   97
      Left            =   0
      TabIndex        =   406
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   96
      Left            =   0
      TabIndex        =   405
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   95
      Left            =   0
      TabIndex        =   404
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   94
      Left            =   0
      TabIndex        =   403
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   93
      Left            =   0
      TabIndex        =   402
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   92
      Left            =   0
      TabIndex        =   401
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   91
      Left            =   0
      TabIndex        =   400
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   90
      Left            =   0
      TabIndex        =   399
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   89
      Left            =   0
      TabIndex        =   398
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   88
      Left            =   0
      TabIndex        =   397
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   87
      Left            =   0
      TabIndex        =   396
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   86
      Left            =   0
      TabIndex        =   395
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   85
      Left            =   0
      TabIndex        =   394
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   84
      Left            =   0
      TabIndex        =   393
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   83
      Left            =   0
      TabIndex        =   392
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   82
      Left            =   0
      TabIndex        =   391
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   81
      Left            =   0
      TabIndex        =   390
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   80
      Left            =   0
      TabIndex        =   389
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   79
      Left            =   0
      TabIndex        =   388
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   78
      Left            =   0
      TabIndex        =   387
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   77
      Left            =   0
      TabIndex        =   386
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   76
      Left            =   0
      TabIndex        =   385
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   75
      Left            =   0
      TabIndex        =   384
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   74
      Left            =   0
      TabIndex        =   383
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   73
      Left            =   0
      TabIndex        =   382
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   72
      Left            =   0
      TabIndex        =   381
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   71
      Left            =   0
      TabIndex        =   380
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   70
      Left            =   0
      TabIndex        =   379
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   69
      Left            =   0
      TabIndex        =   378
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   68
      Left            =   0
      TabIndex        =   377
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   67
      Left            =   0
      TabIndex        =   376
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   66
      Left            =   0
      TabIndex        =   375
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   65
      Left            =   0
      TabIndex        =   374
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   64
      Left            =   0
      TabIndex        =   373
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   63
      Left            =   0
      TabIndex        =   372
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   62
      Left            =   0
      TabIndex        =   371
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   61
      Left            =   0
      TabIndex        =   370
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   60
      Left            =   0
      TabIndex        =   369
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   59
      Left            =   0
      TabIndex        =   368
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   58
      Left            =   0
      TabIndex        =   367
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   57
      Left            =   0
      TabIndex        =   366
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   56
      Left            =   0
      TabIndex        =   365
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   55
      Left            =   0
      TabIndex        =   364
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   54
      Left            =   0
      TabIndex        =   363
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   53
      Left            =   0
      TabIndex        =   362
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   52
      Left            =   0
      TabIndex        =   361
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   51
      Left            =   0
      TabIndex        =   360
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   50
      Left            =   0
      TabIndex        =   359
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   49
      Left            =   0
      TabIndex        =   358
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   48
      Left            =   0
      TabIndex        =   357
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   47
      Left            =   0
      TabIndex        =   356
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   46
      Left            =   0
      TabIndex        =   355
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   45
      Left            =   0
      TabIndex        =   354
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   44
      Left            =   0
      TabIndex        =   353
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   43
      Left            =   0
      TabIndex        =   352
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   42
      Left            =   0
      TabIndex        =   351
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   41
      Left            =   0
      TabIndex        =   350
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   40
      Left            =   0
      TabIndex        =   349
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   39
      Left            =   0
      TabIndex        =   348
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   38
      Left            =   0
      TabIndex        =   347
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   37
      Left            =   0
      TabIndex        =   346
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   36
      Left            =   0
      TabIndex        =   345
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   35
      Left            =   0
      TabIndex        =   344
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   34
      Left            =   0
      TabIndex        =   343
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   33
      Left            =   0
      TabIndex        =   342
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   32
      Left            =   0
      TabIndex        =   341
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   31
      Left            =   0
      TabIndex        =   340
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   30
      Left            =   0
      TabIndex        =   339
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   29
      Left            =   0
      TabIndex        =   338
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   28
      Left            =   0
      TabIndex        =   337
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   27
      Left            =   0
      TabIndex        =   336
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   26
      Left            =   0
      TabIndex        =   335
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   25
      Left            =   0
      TabIndex        =   334
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   24
      Left            =   0
      TabIndex        =   333
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   23
      Left            =   0
      TabIndex        =   332
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   22
      Left            =   0
      TabIndex        =   331
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   21
      Left            =   0
      TabIndex        =   330
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   20
      Left            =   0
      TabIndex        =   329
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   19
      Left            =   0
      TabIndex        =   328
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   18
      Left            =   0
      TabIndex        =   327
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   17
      Left            =   0
      TabIndex        =   326
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   16
      Left            =   0
      TabIndex        =   325
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   15
      Left            =   0
      TabIndex        =   324
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   14
      Left            =   0
      TabIndex        =   323
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   13
      Left            =   0
      TabIndex        =   322
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   12
      Left            =   0
      TabIndex        =   321
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   11
      Left            =   0
      TabIndex        =   320
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   10
      Left            =   0
      TabIndex        =   319
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   9
      Left            =   0
      TabIndex        =   318
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   8
      Left            =   0
      TabIndex        =   317
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   7
      Left            =   0
      TabIndex        =   316
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   6
      Left            =   0
      TabIndex        =   315
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   5
      Left            =   0
      TabIndex        =   314
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   4
      Left            =   0
      TabIndex        =   313
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   3
      Left            =   0
      TabIndex        =   312
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   2
      Left            =   0
      TabIndex        =   311
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   1
      Left            =   0
      TabIndex        =   310
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcold 
      Height          =   15
      Index           =   0
      Left            =   1200
      TabIndex        =   309
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   99
      Left            =   0
      TabIndex        =   308
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   98
      Left            =   0
      TabIndex        =   307
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   97
      Left            =   0
      TabIndex        =   306
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   96
      Left            =   0
      TabIndex        =   305
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   95
      Left            =   0
      TabIndex        =   304
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   94
      Left            =   0
      TabIndex        =   303
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   93
      Left            =   0
      TabIndex        =   302
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   92
      Left            =   0
      TabIndex        =   301
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   91
      Left            =   0
      TabIndex        =   300
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   90
      Left            =   0
      TabIndex        =   299
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   89
      Left            =   0
      TabIndex        =   298
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   88
      Left            =   0
      TabIndex        =   297
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   87
      Left            =   0
      TabIndex        =   296
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   86
      Left            =   0
      TabIndex        =   295
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   85
      Left            =   0
      TabIndex        =   294
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   84
      Left            =   0
      TabIndex        =   293
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   83
      Left            =   0
      TabIndex        =   292
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   82
      Left            =   0
      TabIndex        =   291
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   81
      Left            =   0
      TabIndex        =   290
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   80
      Left            =   0
      TabIndex        =   289
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   79
      Left            =   0
      TabIndex        =   288
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   78
      Left            =   0
      TabIndex        =   287
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   77
      Left            =   0
      TabIndex        =   286
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   76
      Left            =   0
      TabIndex        =   285
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   75
      Left            =   0
      TabIndex        =   284
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   74
      Left            =   0
      TabIndex        =   283
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   73
      Left            =   0
      TabIndex        =   282
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   72
      Left            =   0
      TabIndex        =   281
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   71
      Left            =   0
      TabIndex        =   280
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   70
      Left            =   0
      TabIndex        =   279
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   69
      Left            =   0
      TabIndex        =   278
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   68
      Left            =   0
      TabIndex        =   277
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   67
      Left            =   0
      TabIndex        =   276
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   66
      Left            =   0
      TabIndex        =   275
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   65
      Left            =   0
      TabIndex        =   274
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   64
      Left            =   0
      TabIndex        =   273
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   63
      Left            =   0
      TabIndex        =   272
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   62
      Left            =   0
      TabIndex        =   271
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   61
      Left            =   0
      TabIndex        =   270
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   60
      Left            =   0
      TabIndex        =   269
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   59
      Left            =   0
      TabIndex        =   268
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   58
      Left            =   0
      TabIndex        =   267
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   57
      Left            =   0
      TabIndex        =   266
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   56
      Left            =   0
      TabIndex        =   265
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   55
      Left            =   0
      TabIndex        =   264
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   54
      Left            =   0
      TabIndex        =   263
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   53
      Left            =   0
      TabIndex        =   262
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   52
      Left            =   0
      TabIndex        =   261
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   51
      Left            =   0
      TabIndex        =   260
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   50
      Left            =   0
      TabIndex        =   259
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   49
      Left            =   0
      TabIndex        =   258
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   48
      Left            =   0
      TabIndex        =   257
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   47
      Left            =   0
      TabIndex        =   256
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   46
      Left            =   0
      TabIndex        =   255
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   45
      Left            =   0
      TabIndex        =   254
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   44
      Left            =   0
      TabIndex        =   253
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   43
      Left            =   0
      TabIndex        =   252
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   42
      Left            =   0
      TabIndex        =   251
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   41
      Left            =   0
      TabIndex        =   250
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   40
      Left            =   0
      TabIndex        =   249
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   39
      Left            =   0
      TabIndex        =   248
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   38
      Left            =   0
      TabIndex        =   247
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   37
      Left            =   0
      TabIndex        =   246
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   36
      Left            =   0
      TabIndex        =   245
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   35
      Left            =   0
      TabIndex        =   244
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   34
      Left            =   0
      TabIndex        =   243
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   33
      Left            =   0
      TabIndex        =   242
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   32
      Left            =   0
      TabIndex        =   241
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   31
      Left            =   0
      TabIndex        =   240
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   30
      Left            =   0
      TabIndex        =   239
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   29
      Left            =   0
      TabIndex        =   238
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   28
      Left            =   0
      TabIndex        =   237
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   27
      Left            =   0
      TabIndex        =   236
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   26
      Left            =   0
      TabIndex        =   235
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   25
      Left            =   0
      TabIndex        =   234
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   24
      Left            =   0
      TabIndex        =   233
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   23
      Left            =   0
      TabIndex        =   232
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   22
      Left            =   0
      TabIndex        =   231
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   21
      Left            =   0
      TabIndex        =   230
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   20
      Left            =   0
      TabIndex        =   229
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   19
      Left            =   0
      TabIndex        =   228
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   18
      Left            =   0
      TabIndex        =   227
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   17
      Left            =   0
      TabIndex        =   226
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   16
      Left            =   0
      TabIndex        =   225
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   15
      Left            =   0
      TabIndex        =   224
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   14
      Left            =   0
      TabIndex        =   223
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   13
      Left            =   0
      TabIndex        =   222
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   12
      Left            =   0
      TabIndex        =   221
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   11
      Left            =   0
      TabIndex        =   220
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   10
      Left            =   0
      TabIndex        =   219
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   9
      Left            =   0
      TabIndex        =   218
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   8
      Left            =   0
      TabIndex        =   217
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   7
      Left            =   0
      TabIndex        =   216
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   6
      Left            =   0
      TabIndex        =   215
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   5
      Left            =   0
      TabIndex        =   214
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   4
      Left            =   0
      TabIndex        =   213
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   3
      Left            =   0
      TabIndex        =   212
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   2
      Left            =   0
      TabIndex        =   211
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   1
      Left            =   0
      TabIndex        =   210
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolc 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   0
      Left            =   120
      TabIndex        =   209
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   99
      Left            =   1000
      TabIndex        =   208
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   98
      Left            =   1000
      TabIndex        =   207
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   97
      Left            =   1000
      TabIndex        =   206
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   96
      Left            =   1000
      TabIndex        =   205
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   95
      Left            =   1000
      TabIndex        =   204
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   94
      Left            =   1000
      TabIndex        =   203
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   93
      Left            =   1000
      TabIndex        =   202
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   92
      Left            =   1000
      TabIndex        =   201
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   91
      Left            =   1000
      TabIndex        =   200
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   90
      Left            =   1000
      TabIndex        =   199
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   89
      Left            =   1000
      TabIndex        =   198
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   88
      Left            =   1000
      TabIndex        =   197
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   87
      Left            =   1000
      TabIndex        =   196
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   86
      Left            =   1000
      TabIndex        =   195
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   85
      Left            =   1000
      TabIndex        =   194
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   84
      Left            =   1000
      TabIndex        =   193
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   83
      Left            =   1000
      TabIndex        =   192
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   82
      Left            =   1000
      TabIndex        =   191
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   81
      Left            =   1000
      TabIndex        =   190
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   80
      Left            =   1000
      TabIndex        =   189
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   79
      Left            =   1000
      TabIndex        =   188
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   78
      Left            =   1000
      TabIndex        =   187
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   77
      Left            =   1000
      TabIndex        =   186
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   76
      Left            =   1000
      TabIndex        =   185
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   75
      Left            =   1000
      TabIndex        =   184
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   74
      Left            =   1000
      TabIndex        =   183
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   73
      Left            =   1000
      TabIndex        =   182
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   72
      Left            =   1000
      TabIndex        =   181
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   71
      Left            =   1000
      TabIndex        =   180
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   70
      Left            =   1000
      TabIndex        =   179
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   69
      Left            =   1000
      TabIndex        =   178
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   68
      Left            =   1000
      TabIndex        =   177
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   67
      Left            =   1000
      TabIndex        =   176
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   66
      Left            =   1000
      TabIndex        =   175
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   65
      Left            =   1000
      TabIndex        =   174
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   64
      Left            =   1000
      TabIndex        =   173
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   63
      Left            =   1000
      TabIndex        =   172
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   62
      Left            =   1000
      TabIndex        =   171
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   61
      Left            =   1000
      TabIndex        =   170
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   60
      Left            =   1000
      TabIndex        =   169
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   59
      Left            =   1000
      TabIndex        =   168
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   58
      Left            =   1000
      TabIndex        =   167
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   57
      Left            =   1000
      TabIndex        =   166
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   56
      Left            =   1000
      TabIndex        =   165
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   55
      Left            =   1000
      TabIndex        =   164
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   54
      Left            =   1000
      TabIndex        =   163
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   53
      Left            =   1000
      TabIndex        =   162
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   52
      Left            =   1000
      TabIndex        =   161
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   51
      Left            =   1000
      TabIndex        =   160
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   50
      Left            =   1000
      TabIndex        =   159
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   49
      Left            =   1000
      TabIndex        =   158
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   48
      Left            =   1000
      TabIndex        =   157
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   47
      Left            =   1000
      TabIndex        =   156
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   46
      Left            =   1000
      TabIndex        =   155
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   45
      Left            =   1000
      TabIndex        =   154
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   44
      Left            =   1000
      TabIndex        =   153
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   43
      Left            =   1000
      TabIndex        =   152
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   42
      Left            =   1000
      TabIndex        =   151
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   41
      Left            =   1000
      TabIndex        =   150
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   40
      Left            =   1000
      TabIndex        =   149
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   39
      Left            =   1000
      TabIndex        =   148
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   38
      Left            =   1000
      TabIndex        =   147
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   37
      Left            =   1000
      TabIndex        =   146
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   36
      Left            =   1000
      TabIndex        =   145
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   35
      Left            =   1000
      TabIndex        =   144
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   34
      Left            =   1000
      TabIndex        =   143
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   33
      Left            =   1000
      TabIndex        =   142
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   32
      Left            =   1000
      TabIndex        =   141
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   31
      Left            =   1000
      TabIndex        =   140
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   30
      Left            =   1000
      TabIndex        =   139
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   29
      Left            =   1000
      TabIndex        =   138
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   28
      Left            =   1000
      TabIndex        =   137
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   27
      Left            =   1000
      TabIndex        =   136
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   26
      Left            =   1000
      TabIndex        =   135
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   25
      Left            =   1000
      TabIndex        =   134
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   24
      Left            =   1000
      TabIndex        =   133
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   23
      Left            =   1000
      TabIndex        =   132
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   22
      Left            =   1000
      TabIndex        =   131
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   21
      Left            =   1000
      TabIndex        =   130
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   20
      Left            =   1000
      TabIndex        =   129
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   19
      Left            =   1000
      TabIndex        =   128
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   18
      Left            =   1000
      TabIndex        =   127
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   17
      Left            =   1000
      TabIndex        =   126
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   16
      Left            =   1000
      TabIndex        =   125
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   15
      Left            =   1000
      TabIndex        =   124
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   14
      Left            =   1000
      TabIndex        =   123
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   13
      Left            =   1000
      TabIndex        =   122
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   12
      Left            =   1000
      TabIndex        =   121
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   11
      Left            =   1000
      TabIndex        =   120
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   10
      Left            =   1000
      TabIndex        =   119
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   9
      Left            =   1000
      TabIndex        =   118
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   8
      Left            =   1000
      TabIndex        =   117
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   7
      Left            =   1000
      TabIndex        =   116
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   6
      Left            =   1000
      TabIndex        =   115
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   5
      Left            =   1000
      TabIndex        =   114
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   4
      Left            =   1000
      TabIndex        =   113
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   3
      Left            =   1000
      TabIndex        =   112
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   2
      Left            =   1000
      TabIndex        =   111
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   1
      Left            =   1000
      TabIndex        =   110
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Lblcolb 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   0
      Left            =   1000
      TabIndex        =   109
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   99
      Left            =   2000
      TabIndex        =   108
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   98
      Left            =   2000
      TabIndex        =   107
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   97
      Left            =   2000
      TabIndex        =   106
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   96
      Left            =   2000
      TabIndex        =   105
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   95
      Left            =   2000
      TabIndex        =   104
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   94
      Left            =   3480
      TabIndex        =   103
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   93
      Left            =   2000
      TabIndex        =   102
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   92
      Left            =   2000
      TabIndex        =   101
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   91
      Left            =   2000
      TabIndex        =   100
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   90
      Left            =   2000
      TabIndex        =   99
      Top             =   600
      Width           =   400
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   89
      Left            =   2000
      TabIndex        =   98
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   88
      Left            =   2000
      TabIndex        =   97
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   87
      Left            =   2000
      TabIndex        =   96
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   86
      Left            =   2000
      TabIndex        =   95
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   85
      Left            =   2000
      TabIndex        =   94
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   84
      Left            =   2000
      TabIndex        =   93
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   83
      Left            =   2000
      TabIndex        =   92
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   82
      Left            =   2000
      TabIndex        =   91
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   81
      Left            =   2000
      TabIndex        =   90
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   80
      Left            =   2000
      TabIndex        =   89
      Top             =   600
      Width           =   400
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   79
      Left            =   2000
      TabIndex        =   88
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   78
      Left            =   2000
      TabIndex        =   87
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   77
      Left            =   2000
      TabIndex        =   86
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   76
      Left            =   2000
      TabIndex        =   85
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   75
      Left            =   2000
      TabIndex        =   84
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   74
      Left            =   2000
      TabIndex        =   83
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   73
      Left            =   2000
      TabIndex        =   82
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   72
      Left            =   2000
      TabIndex        =   81
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   71
      Left            =   2000
      TabIndex        =   80
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   70
      Left            =   2000
      TabIndex        =   79
      Top             =   600
      Width           =   400
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   69
      Left            =   2000
      TabIndex        =   78
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   68
      Left            =   2000
      TabIndex        =   77
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   67
      Left            =   2000
      TabIndex        =   76
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   66
      Left            =   2000
      TabIndex        =   75
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   65
      Left            =   2000
      TabIndex        =   74
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   64
      Left            =   2000
      TabIndex        =   73
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   63
      Left            =   2000
      TabIndex        =   72
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   62
      Left            =   2000
      TabIndex        =   71
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   61
      Left            =   2000
      TabIndex        =   70
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   60
      Left            =   2000
      TabIndex        =   69
      Top             =   600
      Width           =   400
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   59
      Left            =   2000
      TabIndex        =   68
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   58
      Left            =   2000
      TabIndex        =   67
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   57
      Left            =   2000
      TabIndex        =   66
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   56
      Left            =   2000
      TabIndex        =   65
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   55
      Left            =   2000
      TabIndex        =   64
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   54
      Left            =   2000
      TabIndex        =   63
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   53
      Left            =   2000
      TabIndex        =   62
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   52
      Left            =   2000
      TabIndex        =   61
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   51
      Left            =   2000
      TabIndex        =   60
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   50
      Left            =   2000
      TabIndex        =   59
      Top             =   600
      Width           =   400
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   49
      Left            =   2000
      TabIndex        =   58
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   48
      Left            =   2000
      TabIndex        =   57
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   47
      Left            =   2000
      TabIndex        =   56
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   46
      Left            =   2000
      TabIndex        =   55
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   45
      Left            =   2000
      TabIndex        =   54
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   44
      Left            =   2000
      TabIndex        =   53
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   43
      Left            =   2000
      TabIndex        =   52
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   42
      Left            =   2000
      TabIndex        =   51
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   41
      Left            =   2000
      TabIndex        =   50
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   40
      Left            =   2000
      TabIndex        =   49
      Top             =   600
      Width           =   400
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   39
      Left            =   2000
      TabIndex        =   48
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   38
      Left            =   2000
      TabIndex        =   47
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   37
      Left            =   2000
      TabIndex        =   46
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   36
      Left            =   2000
      TabIndex        =   45
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   35
      Left            =   2000
      TabIndex        =   44
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   34
      Left            =   2000
      TabIndex        =   43
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   33
      Left            =   2000
      TabIndex        =   42
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   32
      Left            =   2000
      TabIndex        =   41
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   31
      Left            =   2000
      TabIndex        =   40
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   30
      Left            =   2000
      TabIndex        =   39
      Top             =   600
      Width           =   400
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   29
      Left            =   2000
      TabIndex        =   38
      Top             =   0
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   28
      Left            =   2000
      TabIndex        =   37
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   27
      Left            =   2000
      TabIndex        =   36
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   26
      Left            =   2000
      TabIndex        =   35
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   25
      Left            =   2000
      TabIndex        =   34
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   24
      Left            =   2000
      TabIndex        =   33
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   23
      Left            =   2000
      TabIndex        =   32
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   22
      Left            =   2000
      TabIndex        =   31
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   21
      Left            =   2000
      TabIndex        =   30
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   20
      Left            =   2000
      TabIndex        =   29
      Top             =   600
      Width           =   400
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   19
      Left            =   3480
      TabIndex        =   28
      Top             =   2640
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   18
      Left            =   3480
      TabIndex        =   27
      Top             =   2760
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   17
      Left            =   2000
      TabIndex        =   26
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   16
      Left            =   2000
      TabIndex        =   25
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   15
      Left            =   2000
      TabIndex        =   24
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   14
      Left            =   2000
      TabIndex        =   23
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   13
      Left            =   2000
      TabIndex        =   22
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   12
      Left            =   2000
      TabIndex        =   21
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   11
      Left            =   2000
      TabIndex        =   20
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   10
      Left            =   2000
      TabIndex        =   19
      Top             =   600
      Width           =   400
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   9
      Left            =   3480
      TabIndex        =   18
      Top             =   2880
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   8
      Left            =   3480
      TabIndex        =   17
      Top             =   3000
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   7
      Left            =   3480
      TabIndex        =   16
      Top             =   3120
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   6
      Left            =   3480
      TabIndex        =   15
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   5
      Left            =   3480
      TabIndex        =   14
      Top             =   3360
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   4
      Left            =   3480
      TabIndex        =   13
      Top             =   3480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   3
      Left            =   3480
      TabIndex        =   12
      Top             =   3480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   2
      Left            =   3480
      TabIndex        =   11
      Top             =   3480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   15
      Index           =   1
      Left            =   3480
      TabIndex        =   10
      Top             =   3480
      Width           =   405
   End
   Begin VB.Label Lblcola 
      BackColor       =   &H000080FF&
      Height          =   10
      Index           =   0
      Left            =   3480
      TabIndex        =   9
      Top             =   3480
      Width           =   400
   End
   Begin VB.Label Lblg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   4920
      TabIndex        =   8
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label Lblg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   4440
      TabIndex        =   7
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label Lblg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   3960
      TabIndex        =   6
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label Lblg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   230
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      Top             =   1300
      Width           =   400
   End
   Begin VB.Image Imagraph 
      Height          =   2535
      Left            =   4320
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   15
      Left            =   4920
      Top             =   1560
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   14
      Left            =   4800
      Top             =   1680
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   13
      Left            =   4920
      Top             =   1800
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   12
      Left            =   4920
      Top             =   1920
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   11
      Left            =   4920
      Top             =   2040
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   10
      Left            =   4920
      Top             =   2280
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   9
      Left            =   4920
      Top             =   2400
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   8
      Left            =   4920
      Top             =   2520
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   7
      Left            =   4920
      Top             =   2760
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   6
      Left            =   4920
      Top             =   3000
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   5
      Left            =   4920
      Top             =   3120
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   4
      Left            =   4920
      Top             =   3240
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   3
      Left            =   4920
      Top             =   3360
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   2
      Left            =   4920
      Top             =   3600
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   1
      Left            =   4800
      Top             =   3720
      Width           =   1845
   End
   Begin VB.Image Imamo 
      Height          =   195
      Index           =   0
      Left            =   4440
      Top             =   3840
      Width           =   1845
   End
   Begin VB.Image Ima5050 
      Height          =   750
      Left            =   840
      Picture         =   "MillFrmBoard2.frx":20D91
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Image Imapaf 
      Height          =   750
      Left            =   840
      Picture         =   "MillFrmBoard2.frx":216A7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Image Imaata 
      Height          =   750
      Left            =   840
      Picture         =   "MillFrmBoard2.frx":2206A
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Label Lbld 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5430
      TabIndex        =   4
      Top             =   5920
      Width           =   2700
   End
   Begin VB.Label Lblc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   5920
      Width           =   2655
   End
   Begin VB.Label Lblb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5430
      TabIndex        =   2
      Top             =   5310
      Width           =   2700
   End
   Begin VB.Label Lbla 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   5310
      Width           =   2655
   End
   Begin VB.Image Imaansc 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   625
      Picture         =   "MillFrmBoard2.frx":22A49
      Top             =   5760
      Width           =   4005
   End
   Begin VB.Image Imaansd 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   4640
      Picture         =   "MillFrmBoard2.frx":29553
      Top             =   5760
      Width           =   4020
   End
   Begin VB.Image Imaansb 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   4640
      Picture         =   "MillFrmBoard2.frx":3005D
      Top             =   5160
      Width           =   4020
   End
   Begin VB.Image Imaansa 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   625
      Picture         =   "MillFrmBoard2.frx":36B67
      Top             =   5160
      Width           =   4020
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   600
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Label LblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   4455
      Width           =   6375
   End
   Begin VB.Image Imaquest 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   600
      Picture         =   "MillFrmBoard2.frx":3D671
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   8055
   End
   Begin VB.Menu mgame 
      Caption         =   "Game  "
      Begin VB.Menu mexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu moptions 
      Caption         =   "Options  "
      Begin VB.Menu mansconfirm 
         Caption         =   "Answer Confirmation"
         Begin VB.Menu mconon 
            Caption         =   "On"
         End
         Begin VB.Menu mconoff 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu md2 
         Caption         =   "-"
      End
      Begin VB.Menu msound 
         Caption         =   "Sound"
         Begin VB.Menu mson 
            Caption         =   "On"
         End
         Begin VB.Menu md3 
            Caption         =   "-"
         End
         Begin VB.Menu msoff 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu md4 
         Caption         =   "-"
      End
      Begin VB.Menu mbso 
         Caption         =   "Background Sound"
         Begin VB.Menu mtr1 
            Caption         =   "Track1"
         End
         Begin VB.Menu md5 
            Caption         =   "-"
         End
         Begin VB.Menu mtr2 
            Caption         =   "Track2"
         End
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   " Help"
      Begin VB.Menu mabout 
         Caption         =   "About"
      End
      Begin VB.Menu md6 
         Caption         =   "-"
      End
      Begin VB.Menu minfo 
         Caption         =   "Information"
      End
   End
End
Attribute VB_Name = "FrmBoard2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim son, xx, yy, zz, co, gpass As Integer
Dim pos, aco, bco, cco, dco, n, uldiff As Integer
Dim k As Integer
Dim amo As Integer
Dim temp, txt As String

Private Sub Cmdfook_Click()
  FileInputName = File1.FileName
  Pass = 1
End Sub

Private Sub Form_Load()
Dim xyz As String
Randomize Timer
Version = 3
FrmMsgbox.Picture = FrmGraphics!Imamb1.Picture
mexit.Enabled = True
mconon.Enabled = False
mson.Enabled = False
Cmdexit3.Visible = False
Imacorrect.Visible = False
Imaincorrect.Visible = False

psound = 1
FinalAon = 1
FinalA = 0
Lifeline = 0

LfLine = 1
sloop = 1
cata = 0: cpaf = 0
s5050 = 0: spaf = 0: sata = 0
Ima5050.Top = 1980
Ima5050.Left = 840
Ima5050.Picture = FrmGraphics!Ima5050.Picture
Imapaf.Top = 2760
Imapaf.Left = 840
Imapaf.Picture = FrmGraphics!Imapaf.Picture
Imaata.Top = 3540
Imaata.Left = 840
Imaata.Picture = FrmGraphics!Imaata.Picture
Imafa.Top = 3400
Imafa.Left = 2640
Imafa.Picture = FrmGraphics!Imafinans.Picture
Imafa.Visible = False

' Set Background colours for graph
' Dark Blue to light blue
Dim yy, n As Integer
yy = &HFFFF10
For n = 0 To 99
  Lblcola(n).BackColor = yy
  Lblcolb(n).BackColor = yy
  Lblcolc(n).BackColor = yy
  Lblcold(n).BackColor = yy
  yy = yy + 2
Next n

' Mauve - Blue
'yy = &HFF80D0
'For n = 0 To 99
'  Lblcola(n).BackColor = yy
'  Lblcolb(n).BackColor = yy
'  Lblcolc(n).BackColor = yy
'  Lblcold(n).BackColor = yy
'  yy = yy - 1
'Next n

LblQuestion.ForeColor = &HFFFFFF
LblQuestion.Left = 1450
LblQuestion.Width = 6350
Imaansa.Left = 625
Imaansb.Left = 4640
Imaansc.Left = 625
Imaansd.Left = 4640
' setup imawin position
ImaWin.Visible = False
ImaWin.Top = 1210
ImaWin.Left = 2160
Imapafper.Top = 1410
Imapafper.Left = 2600
Lblpafout.Top = 1260
Lblpafout.Height = 2970
Lblpafout.Left = 2100
Lblpafin.Top = 2640
Lblpafin.Height = 1455
Lblpafin.Left = 2220
Imagraph.Left = 4425
Imagraph.Top = 1240
Lblfileopen.Width = 4090
Lblfileopen.Top = 1240
Lblfileopen.Left = 2480
Lblfileopen.Visible = False
Lblfoh.Width = 3495
Lblfoh.Top = 1320
Lblfoh.Left = 2780
Lblfoh.Visible = False
File1.Width = 3855
File1.Top = 1640
File1.Left = 2600
File1.Visible = False
Cmdfook.Width = 855
Cmdfook.Top = 3930
Cmdfook.Left = 4040
Cmdfook.Visible = False
' Setup Pictures
ImaWW.Picture = FrmGraphics!Imaw.Picture
ImaWW.Top = 6550
ImaWW.Left = 7990
' Set position for Y & N Graphics
Imayes.Left = 3620
Imayes.Top = 3780
Imano.Left = 4800
Imano.Top = 3780
Dim xx As Integer
xx = 3960
Lbla.Left = 1640: Lbla.Top = 5285: Lbla.Width = 2750
Lblb.Left = 5400: Lblb.Top = 5285: Lblb.Width = 2750
Lblc.Left = 1640: Lblc.Top = 5900: Lblc.Width = 2750
Lbld.Left = 5400: Lbld.Top = 5900: Lbld.Width = 2750

For n = 1 To 15
  Imamo(n).Top = xx
  Imamo(n).Width = 1845
  Imamo(n).Height = 195
  Imamo(n).Left = 6760
  Imamo(n).Picture = FrmGraphics!Imamo(n).Picture
  xx = xx - 190
Next n

' Set Graph positions
xx = 4550
For n = 0 To 3
  Lblg(n).Left = xx
  Lblg(n).Height = 230
  Lblg(n).Width = 400
  Lblg(n).Top = 1300
  xx = xx + 480
  If n = 1 Then xx = xx + 10
Next n

xx = 3865
For n = 0 To 99
  Lblcola(n).Height = 24
  Lblcola(n).Top = xx
  Lblcola(n).Left = 4580
  Lblcola(n).Width = 360
  xx = xx - 23
Next n

xx = 3865
For n = 0 To 99
  Lblcolb(n).Height = 24
  Lblcolb(n).Top = xx
  Lblcolb(n).Left = 5060
  Lblcolb(n).Width = 360
  xx = xx - 23
Next n

xx = 3865
For n = 0 To 99
  Lblcolc(n).Height = 24
  Lblcolc(n).Top = xx
  Lblcolc(n).Left = 5550
  Lblcolc(n).Width = 360
  xx = xx - 23
Next n

xx = 3865
For n = 0 To 99
  Lblcold(n).Height = 24
  Lblcold(n).Top = xx
  Lblcold(n).Left = 6030
  Lblcold(n).Width = 360
  xx = xx - 23
Next n

L(1) = 80: U(1) = 90
L(2) = 75: U(2) = 85
L(3) = 65: U(3) = 75
L(4) = 55: U(4) = 65
L(5) = 37: U(5) = 47
L(6) = 36: U(6) = 42
L(7) = 30: U(7) = 38
L(8) = 22: U(8) = 33
L(9) = 20: U(9) = 27
L(10) = 16: U(10) = 24
L(11) = 13: U(11) = 19
L(12) = 12: U(12) = 16
L(13) = 9: U(13) = 14
L(14) = 7: U(14) = 11
L(15) = 4: U(15) = 7

Qacd(1) = "WGABZTDYC"
Qacd(2) = "AZBDYCRDR"
Qacd(3) = "DCBHQOINA"
Qacd(4) = "XYDACHKKB"
Qacd(5) = "BDZHHAZCZ"
Qacd(6) = "CAJTDTBTD"
Qacd(7) = "MBPCXDAPQ"
Qacd(8) = "QZCJABYWD"
Qacd(9) = "UKCABHDLK"

PickSt01 = "1314171921232426313233394344454951525658"
PickSt01 = PickSt01 & "61626567727476778385868993949597"
xx = 1
For n = 1 To 57
  CodNum(n) = Val(Mid(PickSt01, xx, 2))
  xx = xx + 2
Next n

mAmo(1) = "£100"
mAmo(2) = "£200"
mAmo(3) = "£300"
mAmo(4) = "£500"
mAmo(5) = "£1,000"
mAmo(6) = "£2,000"
mAmo(7) = "£4,000"
mAmo(8) = "£8,000"
mAmo(9) = "£16,000"
mAmo(10) = "£32,000"
mAmo(11) = "£64,000"
mAmo(12) = "£125,000"
mAmo(13) = "£250,000"
mAmo(14) = "£500,000"
mAmo(15) = "£1,000,000"

pafName(1) = "John"
pafName(2) = "Martin"
pafName(3) = "Julie"
pafName(4) = "Mary"

c(1, 1) = 100: c(1, 2) = 100: c(1, 3) = 100: c(1, 4) = 100
c(2, 1) = 100: c(2, 2) = 100: c(2, 3) = 100: c(2, 4) = 100
c(3, 1) = 99: c(3, 2) = 100: c(3, 3) = 100: c(3, 4) = 100
c(4, 1) = 99: c(4, 2) = 99: c(4, 3) = 99: c(4, 4) = 100
c(5, 1) = 92: c(5, 2) = 94: c(5, 3) = 93: c(5, 4) = 98
c(6, 1) = 89: c(6, 2) = 94: c(6, 3) = 92: c(6, 4) = 96
c(7, 1) = 83: c(7, 2) = 90: c(7, 3) = 87: c(7, 4) = 92
c(8, 1) = 72: c(8, 2) = 81: c(8, 3) = 78: c(8, 4) = 85
c(9, 1) = 61: c(9, 2) = 72: c(8, 3) = 68: c(9, 4) = 75
c(10, 1) = 40: c(10, 2) = 60: c(10, 3) = 52: c(10, 4) = 65
c(11, 1) = 30: c(11, 2) = 50: c(11, 3) = 42: c(11, 4) = 55
c(12, 1) = 20: c(12, 2) = 40: c(12, 3) = 32: c(12, 4) = 45
c(13, 1) = 12: c(13, 2) = 28: c(13, 3) = 25: c(13, 4) = 32
c(14, 1) = 7: c(14, 2) = 20: c(14, 3) = 15: c(14, 4) = 25
c(15, 1) = 4: c(15, 2) = 7: c(15, 3) = 6: c(15, 4) = 10

Imagraph.Picture = FrmGraphics!Imagraph.Picture

FrmBoard2.Show
Dim d1, d2, d3, d4, p1, p2 As Integer
d1 = Int(FrmBoard2.ScaleWidth / 2)
d2 = Int(Imacorrect.Width) / 2
p1 = d1 - d2 - FrmBoard2.Left
d3 = Int(FrmBoard2.ScaleHeight / 2)
d4 = Int(Imacorrect.Height) / 2
p2 = d3 - d4 - FrmBoard2.Top

Imacorrect.Top = FrmBoard2.Top + p2 + 98
Imacorrect.Left = FrmBoard2.Left + p1
Imaincorrect.Top = FrmBoard2.Top + p2 + 98
Imaincorrect.Left = FrmBoard2.Left + p1

If Version <> 1 Then
  Call Questions
  Call Check_Question_File
End If

If Version = 1 Then
  Lblfileopen.Visible = True
  Lblfoh.Visible = True
  File1.Visible = True
  Cmdfook.Visible = True
loadinput:
  cpth = Trim(App.Path)
  Pass = 0
  Do
    DoEvents
    File1.Visible = True
  Loop Until Pass = 1

  FileInputName = File1.FileName

  If FileInputName = "" Then
    msgtxt = "No Question File selected"
    msgtxt = msgtxt & Chr(13) & Chr(10)
    msgtxt = msgtxt & Chr(13) & Chr(10)
    msgtxt = msgtxt & "Click on Question File to highlight,"
    msgtxt = msgtxt & Chr(13) & Chr(10)
    msgtxt = msgtxt & "then click on Open."
    optval = vbExclamation + vbOKOnly
    retval = MsgBox(msgtxt, optval, "Question File Error")
    File1.SetFocus
    GoTo loadinput
  End If

  cpth = cpth & "\" & FileInputName
  Call Input_Question_File
  ' Check there is at least 1 question in each category
  Dim ww As Integer
  ww = 0
  Call Check_Question_Categorys(ww)

  If ww = 1 Then
    msgtxt = "Input File Error!"
    msgtxt = msgtxt & Chr(13) & Chr(13)
    msgtxt = msgtxt & "There must be at least 1 question assigned"
    msgtxt = msgtxt & Chr(13)
    msgtxt = msgtxt & "to each of the 15 money values"
    msgtxt = msgtxt & Chr(13) & Chr(13)
    msgtxt = msgtxt & "Use Millionaire_Qeditor to add more"
    msgtxt = msgtxt & Chr(13)
    msgtxt = msgtxt & "questions to " & FileInputName
    optval = vbExclamation + vbOKOnly
    retval = MsgBox(msgtxt, optval, "Question File Error")
    File1.SetFocus
    GoTo loadinput
  End If
  Call Check_Question_File
  Lblfileopen.Visible = False
  File1.Visible = False
  Lblfoh.Visible = False
  Cmdfook.Visible = False
End If

If Version = 3 Then
  cpth = Trim(App.Path)
  xyz = ""
  xyz = Left(UserName, (Len(UserName) - 1))
  If Trim(xyz) = "" Then
    xyz = GetIPHostName()
  End If
  If Trim(xyz) = "" Then xyz = ""
  cpth = cpth & "\" & xyz & "Shuffle.txt"

  If FileExists(cpth) = True Then
    'Read Input from File
    Open cpth For Input As #1
 
    For n = 1 To NumOfQuest
      Input #1, Sh(n)
    Next n
    For n = 1 To 15
      Input #1, Cn(n)
    Next n
    For n = 1 To 15
      Input #1, Pn(n)
    Next n
  
    Close 1
  End If
  Dim Remake As Integer
  Remake = 0
  'Remake = 1
  If FileExists(cpth) = False Or Remake = 1 Then
    'MsgBox "File does not exist - Cpth =" & cpth
    Call Create_Shuffle_File
  End If
End If

Pass = 1
Call New_Game

End Sub

Public Sub New_Game()
Dim n, son As Integer
tpass = 0: Ttm = 0
LblQuit.Visible = False
' reset icons
Ima5050.Picture = FrmGraphics!Ima5050.Picture
Imapaf.Picture = FrmGraphics!Imapaf.Picture
Imaata.Picture = FrmGraphics!Imaata.Picture
' Delete Phone a Friend
Imapafper.Visible = False
Lblpafout.Visible = False
Lblpafin.Visible = False
' Delete Graph
For n = 0 To 99
  Lblcola(n).Visible = False
  Lblcolb(n).Visible = False
  Lblcolc(n).Visible = False
  Lblcold(n).Visible = False
Next n
For n = 0 To 3
  Lblg(n).Visible = False
Next n
Imagraph.Visible = False
If psound = 1 Then
  ' disable background sound selection
  son = 0
  If mtr1.Enabled = True Then son = 1
  If mtr2.Enabled = True Then son = 2
  mtr1.Enabled = False
  mtr2.Enabled = False
  wavefile = Trim(App.Path & "\letsplay.wav")
  Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
  Tmr001.Interval = 6400
  Tmr001.Enabled = True
  Do
    DoEvents
  Loop Until tpass = 1
  ' enable background sound selection
  If son = 1 Then mtr1.Enabled = True
  If son = 2 Then mtr2.Enabled = True
End If
Do
  DoEvents

Loop Until Pass = 1

' Start of New Game
For n = 0 To 99
  Lblcola(n).Visible = False
  Lblcolb(n).Visible = False
  Lblcolc(n).Visible = False
  Lblcold(n).Visible = False
Next n
For n = 0 To 3
  Lblg(n).Visible = False
Next n

s5050 = 0: spaf = 0: sata = 0
Choice = 0

cQ = 1

If Version <> 3 Then
  Call Shuffle_Questions
End If

' Main Loop
Do
  Bs = 1
  If cQ > 1 Then
    If psound = 1 Then
      ' disable background sound selection
      son = 0
      If mtr1.Enabled = True Then son = 1
      If mtr2.Enabled = True Then son = 2
      mtr1.Enabled = False
      mtr2.Enabled = False
      wavefile = Trim(App.Path & "\nextq.wav")
      Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
      t2Pass = 0
      Tmr002.Interval = 3700
      Tmr002.Enabled = True
      Do
        DoEvents
      Loop Until t2Pass = 1
      ' enable background sound selection
      If son = 1 Then mtr1.Enabled = True
      If son = 2 Then mtr2.Enabled = True
    End If
  End If
  If sloop = 1 Then wavefile = Trim(App.Path & "\mloop01.wav")
  If sloop = 2 Then wavefile = Trim(App.Path & "\mloop02.wav")
  If psound = 1 Then Call PlaySound(wavefile, 0, SND_ASYNC + SND_LOOP)
  If cata = 1 Then
    ' Delete Graph
    For n = 0 To 99
      Lblcola(n).Visible = False
      Lblcolb(n).Visible = False
      Lblcolc(n).Visible = False
      Lblcold(n).Visible = False
    Next n
    For n = 0 To 3
      Lblg(n).Visible = False
    Next n
    Imagraph.Visible = False
    cata = 0
  End If
  If cpaf = 1 Then
    Imapafper.Visible = False
    Lblpafout.Visible = False
    Lblpafin.Visible = False
    cpaf = 0
  End If
  ' backcolor to black - forecolor to white
  Imaansa.Picture = FrmGraphics!Imaansa.Picture
  Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
  Imaansb.Picture = FrmGraphics!Imaansb.Picture
  Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
  Imaansc.Picture = FrmGraphics!Imaansc.Picture
  Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
  Imaansd.Picture = FrmGraphics!Imaansd.Picture
  Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  Act1 = 0: Act2 = 0: Act3 = 0: Act4 = 0
  Bs = 0
  c5050 = 0
  Call Select_question
  Pass = 0
  'MsgBox "lifeline set"
  Lifeline = 1
  If cQ > 1 Then LblQuit.Visible = True
  Do
    DoEvents
    
    If Choice = CorPos Then
      Bs = 1
      Pass = 1
      Lifeline = 0
      LblQuit.Visible = False
      ' Pause for answer sound
      If psound = 1 And cQ > 9 Then
        wavefile = Trim(App.Path & "\answer.wav")
        Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
        t2Pass = 0
        Tmr002.Interval = 2600
        Tmr002.Enabled = True
        Do
         DoEvents
        Loop Until t2Pass = 1
      End If
      If CorPos = "1" Then
        Imaansa.Picture = FrmGraphics!Imaansag.Picture
        Lbla.BackColor = &HFF00&: Lbla.ForeColor = &H0&
      End If
      If CorPos = "2" Then
        Imaansb.Picture = FrmGraphics!Imaansbg.Picture
        Lblb.BackColor = &HFF00&: Lblb.ForeColor = &H0&
      End If
      If CorPos = "3" Then
        Imaansc.Picture = FrmGraphics!Imaanscg.Picture
        Lblc.BackColor = &HFF00&: Lblc.ForeColor = &H0&
      End If
      If CorPos = "4" Then
        Imaansd.Picture = FrmGraphics!Imaansdg.Picture
        Lbld.BackColor = &HFF00&: Lbld.ForeColor = &H0&
      End If
      Imacorrect.Visible = True
      
      'Play Correct Sound
      If psound = 1 Then
        wavefile = Trim(App.Path & "\correct2.wav")
        Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
        t2Pass = 0
        Tmr002.Interval = 820
        Tmr002.Enabled = True
        Do
         DoEvents
        Loop Until t2Pass = 1
      End If
      
    End If
    If Choice <> CorPos And Choice <> 0 Then
      ' Wrong Answer Given
      Bs = 1
      'MsgBox "lifeline off"
      Lifeline = 0
      LblQuit.Visible = False
      ' Pause for answer sound
      If psound = 1 And cQ > 9 Then
        wavefile = Trim(App.Path & "\answer.wav")
        Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
        t2Pass = 0
        Tmr002.Interval = 2600
        Tmr002.Enabled = True
        Do
         DoEvents
        Loop Until t2Pass = 1
      End If
      Call Clear_Top_Screen
      Imaincorrect.Visible = True
      ' Play wrong answer sound
      If psound = 1 Then
        wavefile = Trim(App.Path & "\wrong2.wav")
        Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
        t2Pass = 0
        Tmr002.Interval = 1220
        Tmr002.Enabled = True
        Do
         DoEvents
        Loop Until t2Pass = 1
      End If
      
      Tmr003.Enabled = True
      t3Pass = 1: Switch = 0
      Do
        DoEvents
      Loop Until t3Pass = 15
      
      Tmr003.Enabled = False
      If CorPos = "1" Then
        Imaansa.Picture = FrmGraphics!Imaansag.Picture
        Lbla.BackColor = &HFF00&: Lbla.ForeColor = &H0&
      End If
      If CorPos = "2" Then
        Imaansb.Picture = FrmGraphics!Imaansbg.Picture
        Lblb.BackColor = &HFF00&: Lblb.ForeColor = &H0&
      End If
      If CorPos = "3" Then
        Imaansc.Picture = FrmGraphics!Imaanscg.Picture
        Lblc.BackColor = &HFF00&: Lblc.ForeColor = &H0&
      End If
      If CorPos = "4" Then
        Imaansd.Picture = FrmGraphics!Imaansdg.Picture
        Lbld.BackColor = &HFF00&: Lbld.ForeColor = &H0&
      End If
      
      'Bs = 0
      t2Pass = 0
      Tmr002.Interval = 2500
      Tmr002.Enabled = True
      Do
        DoEvents
      Loop Until t2Pass = 1
      Call End_Of_Game
    End If
  Loop Until Pass = 1
  
  For n = 1 To 15
    Imamo(n).Picture = FrmGraphics!Imamo(n).Picture
  Next n
  Imamo(cQ).Picture = FrmGraphics!Imamos(cQ)
  cQ = cQ + 1: Choice = 0
  Call Clear_Top_Screen
  tpass = 0: Tmr001.Interval = 1500
  Tmr001.Enabled = True
  Do
    DoEvents
  Loop Until tpass = 1
  Imacorrect.Visible = False
  LblQuestion.Visible = True
  LblQuestion.FontSize = 24
  LblQuestion.ForeColor = &HFFFF&
  LblQuestion.Caption = " *  " & mAmo(cQ - 1) & "  *"
  Imaansa.Picture = FrmGraphics!Imaansa.Picture
  Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
  Imaansb.Picture = FrmGraphics!Imaansb.Picture
  Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
  Imaansc.Picture = FrmGraphics!Imaansc.Picture
  Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
  Imaansd.Picture = FrmGraphics!Imaansd.Picture
  Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  Lbla.Visible = False
  Lblb.Visible = False
  Lblc.Visible = False
  Lbld.Visible = False
  Bs = 1
  tpass = 0: Tmr001.Interval = 2000
  Tmr001.Enabled = True
  Do
    DoEvents
  Loop Until tpass = 1
  ' cQ = 16
  If cQ = 16 Then
    ' Million Pound Winner
    Imamo(15).Picture = FrmGraphics!Imamos(15)
    Bs = 1
    '
    ' Add Millionaire display routine here
    '
    LblQuestion.Visible = False
    Call Clear_Top_Screen
    ImaWin.Visible = True
    retval = 0
    FrmMsgbox.Caption = "Millionaire"
    FrmMsgbox!Imatype = FrmGraphics!Imaq.Picture
    FrmMsgbox!CmdNo.Visible = True
    FrmMsgbox!CmdYes.Visible = True
    FrmMsgbox!CmdOK.Visible = False
    msgtxt = "Play Again"
    FrmMsgbox!Lblmsg = msgtxt
    FrmMsgbox.Show 1
    If retval <> 6 Then
      wavefile = Trim(App.Path & "\" & vbNullString)
      Call PlaySound(wavefile, 0, SND_ASYNC + SND_PURGE)
      Call Exit_Game
    End If
    ImaWin.Visible = False
    Call Clear_Screen
    Pass = 1
    Imamo(15).Picture = FrmGraphics!Imamo(15).Picture
    LblQuestion.FontSize = 10
    LblQuestion.ForeColor = &HFFFFFF
    LblQuestion.Visible = False
    Call New_Game
  End If
  LblQuestion.FontSize = 10
  LblQuestion.ForeColor = &HFFFFFF
  LblQuestion.Visible = False
  
Loop
End Sub

Public Sub End_Of_Game()
  ' Display Winings after wrong answer loss
  Bs = 1
  Imaincorrect.Visible = False
  Call Calculate_Winnings
  If CorPos = "1" Then
    Imaansa.Picture = FrmGraphics!Imaansag.Picture
    Lbla.BackColor = &HFF00&: Lbla.ForeColor = &H0
  End If
  If CorPos = "2" Then
    Imaansb.Picture = FrmGraphics!Imaansbg.Picture
    Lblb.BackColor = &HFF00&: Lblb.ForeColor = &H0
  End If
  If CorPos = "3" Then
    Imaansc.Picture = FrmGraphics!Imaanscg.Picture
    Lblc.BackColor = &HFF00&: Lblc.ForeColor = &H0
  End If
  If CorPos = "4" Then
    Imaansd.Picture = FrmGraphics!Imaansdg.Picture
    Lbld.BackColor = &HFF00&: Lbld.ForeColor = &H0
  End If
  retval = 0
  FrmMsgbox.Caption = "Millionaire"
  FrmMsgbox!Imatype = FrmGraphics!Imai.Picture
  FrmMsgbox!CmdNo.Visible = False
  FrmMsgbox!CmdYes.Visible = False
  FrmMsgbox!CmdOK.Visible = True
  msgtxt = "You leave with " & MoWon & "  "
  FrmMsgbox!Lblmsg = msgtxt
  FrmMsgbox.Show 1
  Call Clear_Screen
  retval = 0
  FrmMsgbox.Caption = "Millionaire"
  FrmMsgbox!Imatype = FrmGraphics!Imaq.Picture
  FrmMsgbox!CmdNo.Visible = True
  FrmMsgbox!CmdYes.Visible = True
  FrmMsgbox!CmdOK.Visible = False
  msgtxt = "Play Again"
  FrmMsgbox!Lblmsg = msgtxt
  FrmMsgbox.Show 1

  If retval <> 6 Then
    wavefile = Trim(App.Path & "\" & vbNullString)
    Call PlaySound(wavefile, 0, SND_ASYNC + SND_PURGE)
    Call Exit_Game
  End If
  Pass = 1
  Call New_Game
End Sub

Public Sub Calculate_Winnings()
  MoWon = mAmo(cQ - 1)
  If MoWon = "" Then MoWon = "£0"
  If Ttm = 0 Then
    If cQ < 6 Then
      MoWon = "£0"
    End If
    If cQ > 5 And cQ < 11 Then
      MoWon = "£1,000"
    End If
    If cQ > 10 Then
      MoWon = "£32,000"
    End If
  End If
End Sub



Public Sub Clear_Top_Screen()
  ' Delete Phone a Friend
  Lblpafout.Visible = False
  Lblpafin.Visible = False
  Imapafper.Visible = False
  ' Delete Graph
  For n = 0 To 99
    Lblcola(n).Visible = False
    Lblcolb(n).Visible = False
    Lblcolc(n).Visible = False
    Lblcold(n).Visible = False
  Next n
  For n = 0 To 3
    Lblg(n).Visible = False
  Next n
  Imagraph.Visible = False
End Sub

Public Sub Clear_Screen()
  Imaansa.Picture = FrmGraphics!Imaansa.Picture
  Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
  Imaansb.Picture = FrmGraphics!Imaansb.Picture
  Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
  Imaansc.Picture = FrmGraphics!Imaansc.Picture
  Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
  Imaansd.Picture = FrmGraphics!Imaansd.Picture
  Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  For n = 1 To 15
    Imamo(n).Picture = FrmGraphics!Imamo(n).Picture
  Next n
  LblQuestion.Visible = False
  Lbla.Visible = False
  Lblb.Visible = False
  Lblc.Visible = False
  Lbld.Visible = False
  s5050 = 1: spaf = 1: sata = 1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Bs = 1 Then Exit Sub
  If Choice <> 0 Then Exit Sub
  If FinalA = 0 Then
    Imaansa.Picture = FrmGraphics!Imaansa.Picture
    Lbla.BackColor = &H0&:: Lbla.ForeColor = &HFFFFFF
    Imaansb.Picture = FrmGraphics!Imaansb.Picture
    Lblb.BackColor = &H0&:: Lblb.ForeColor = &HFFFFFF
    Imaansc.Picture = FrmGraphics!Imaansc.Picture
    Lblc.BackColor = &H0&:: Lblc.ForeColor = &HFFFFFF
    Imaansd.Picture = FrmGraphics!Imaansd.Picture
    Lbld.BackColor = &H0&:: Lbld.ForeColor = &HFFFFFF
    LblQuit.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
  Imano.Picture = FrmGraphics!Imanoff.Picture
  Imayes.Picture = FrmGraphics!Imayoff.Picture
End Sub

Private Sub Ima5050_Click()
  If LfLine = 0 Or Lifeline = 0 Then Exit Sub
  If FinalA = 1 Then Exit Sub
  If s5050 = 1 Then
    ' Enter msgbox here to say Lifeline already used
    Exit Sub
  End If
  ' disable other Lifelines
  LfLine = 0
  s5050 = 1: c5050 = 1
  Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  co = 0
  Do
    xx = 1 + Int(Rnd * 4)
    If xx = 1 And Act1 = 1 And co < 2 And CorPos <> 1 Then
      co = co + 1: Act1 = 0: Lbla.Visible = False
    End If
    If xx = 2 And Act2 = 1 And co < 2 And CorPos <> 2 Then
      co = co + 1: Act2 = 0: Lblb.Visible = False
    End If
    If xx = 3 And Act3 = 1 And co < 2 And CorPos <> 3 Then
      co = co + 1: Act3 = 0: Lblc.Visible = False
    End If
    If xx = 4 And Act4 = 1 And co < 2 And CorPos <> 4 Then
      co = co + 1: Act4 = 0: Lbld.Visible = False
    End If
  Loop Until co > 1
  If psound = 1 Then
    wavefile = Trim(App.Path & "\5050.wav")
    Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
    tpass = 0: Tmr001.Interval = 1400
    Tmr001.Enabled = True
    Do
      DoEvents
    Loop Until tpass = 1
    If sloop = 1 Then wavefile = Trim(App.Path & "\mloop01.wav")
    If sloop = 2 Then wavefile = Trim(App.Path & "\mloop02.wav")
    Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
  End If
  ' enable other Lifelines
   LfLine = 1
End Sub

Private Sub Ima5050_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalA = 1 Then Exit Sub
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
  If s5050 = 1 Then Exit Sub
  Ima5050.Picture = FrmGraphics!Ima5050h.Picture
  
End Sub

Private Sub Imaansa_Click()
  If Act1 = 0 Or Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalAon = 1 Then
    Call Final_Answer
    ' Once returned if yes then continue / if no then exit sub
    If FinalYN = 0 Then Exit Sub
  End If
  Choice = 1
End Sub

Private Sub Imaansa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Act1 = 0 Or LfLine = 0 Then Exit Sub
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansa.Picture = FrmGraphics!Imaansas.Picture
    Lbla.BackColor = &H8080FF: Lbla.ForeColor = &H0
    Imaansb.Picture = FrmGraphics!Imaansb.Picture
    Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
    Imaansc.Picture = FrmGraphics!Imaansc.Picture
    Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
    Imaansd.Picture = FrmGraphics!Imaansd.Picture
    Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
End Sub


Private Sub Imaansb_Click()
  If Act2 = 0 Or Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalAon = 1 Then
    Call Final_Answer
    ' Once returned if yes then continue / if no then exit sub
    If FinalYN = 0 Then Exit Sub
  End If
  Choice = 2
End Sub

Private Sub Imaansb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Act2 = 0 Or LfLine = 0 Then Exit Sub
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansb.Picture = FrmGraphics!Imaansbs.Picture
    Lblb.BackColor = &H8080FF:: Lblb.ForeColor = &H0
    Imaansa.Picture = FrmGraphics!Imaansa.Picture
    Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
    Imaansc.Picture = FrmGraphics!Imaansc.Picture
    Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
    Imaansd.Picture = FrmGraphics!Imaansd.Picture
    Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
End Sub

Private Sub Imaansc_Click()
  If Act3 = 0 Or Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalAon = 1 Then
    Call Final_Answer
    ' Once returned if yes then continue / if no then exit sub
    If FinalYN = 0 Then Exit Sub
  End If
  Choice = 3
End Sub

Private Sub Imaansc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Act3 = 0 Or LfLine = 0 Then Exit Sub
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansc.Picture = FrmGraphics!Imaanscs.Picture
    Lblc.BackColor = &H8080FF: Lblc.ForeColor = &H0
    Imaansa.Picture = FrmGraphics!Imaansa.Picture
    Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
    Imaansb.Picture = FrmGraphics!Imaansb.Picture
    Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
    Imaansd.Picture = FrmGraphics!Imaansd.Picture
    Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
End Sub

Private Sub Imaansd_Click()
  If Act4 = 0 Or Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalAon = 1 Then
    Call Final_Answer
    ' Once returned if yes then continue / if no then exit sub
    If FinalYN = 0 Then Exit Sub
  End If
  Choice = 4
End Sub

Private Sub Imaansd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Act4 = 0 Or LfLine = 0 Then Exit Sub
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansd.Picture = FrmGraphics!Imaansds.Picture
    Lbld.BackColor = &H8080FF: Lbld.ForeColor = &H0
    Imaansa.Picture = FrmGraphics!Imaansa.Picture
    Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
    Imaansb.Picture = FrmGraphics!Imaansb.Picture
    Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
    Imaansc.Picture = FrmGraphics!Imaansc.Picture
    Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
End Sub

Private Sub Imaata_Click()
  If LfLine = 0 Or Lifeline = 0 Then Exit Sub
  If FinalA = 1 Then Exit Sub
  If sata = 1 Then
    ' Enter msgbox here to say Lifeline already used
    Exit Sub
  End If
  Imaata.Picture = FrmGraphics!Imaatas.Picture
  ' disable other Lifelines
  LfLine = 0
  ' disable background sound selection
  son = 0
  If mtr1.Enabled = True Then son = 1
  If mtr2.Enabled = True Then son = 2
  mtr1.Enabled = False
  mtr2.Enabled = False
  sata = 1:  cata = 1
  ' Show Blank graph
  Imagraph.Visible = True
  
  ' cQ = 15
  pos = 0
  For n = 1 To 15
    If cQ = n Then pos = n: Exit For
  Next n
  
  aco = 0: bco = 0: cco = 0: dco = 0
  uldiff = U(pos) - L(pos)
  
  xx = L(pos) + Int(Rnd * uldiff)
  If CorPos = "1" Then aco = aco + xx
  If CorPos = "2" Then bco = bco + xx
  If CorPos = "3" Then cco = cco + xx
  If CorPos = "4" Then dco = dco + xx
  amo = 100 - xx
  If c5050 = 0 Then
    For n = 1 To amo
      xx = 1 + Int(Rnd * 4)
      If xx = 1 Then aco = aco + 1
      If xx = 2 Then bco = bco + 1
      If xx = 3 Then cco = cco + 1
      If xx = 4 Then dco = dco + 1
    Next n
  End If
  If c5050 = 1 Then
    For n = 1 To amo
      gpass = 0
      Do
        xx = 1 + Int(Rnd * 4)
        If xx = 1 And Act1 = 1 Then gpass = 1
        If xx = 2 And Act2 = 1 Then gpass = 1
        If xx = 3 And Act3 = 1 Then gpass = 1
        If xx = 4 And Act4 = 1 Then gpass = 1
      Loop Until gpass = 1
      If xx = 1 Then aco = aco + 1
      If xx = 2 Then bco = bco + 1
      If xx = 3 Then cco = cco + 1
      If xx = 4 Then dco = dco + 1
    Next n
  
  End If

  If psound = 1 Then
    wavefile = Trim(App.Path & "\askaud01.wav")
    Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
    tpass = 0: Tmr001.Interval = 5550
    Tmr001.Enabled = True
    Do
      DoEvents
    Loop Until tpass = 1
  End If
  
   For n = 0 To 3
    Lblg(n).Visible = True
  Next n
  For n = 0 To aco - 1
    Lblcola(n).Visible = True
  Next n
  For n = 0 To bco - 1
    Lblcolb(n).Visible = True
  Next n
  For n = 0 To cco - 1
    Lblcolc(n).Visible = True
  Next n
  For n = 0 To dco - 1
    Lblcold(n).Visible = True
  Next n
  Lblg(0).Caption = aco & "%"
  Lblg(1).Caption = bco & "%"
  Lblg(2).Caption = cco & "%"
  Lblg(3).Caption = dco & "%"
  If sloop = 1 Then wavefile = Trim(App.Path & "\mloop01.wav")
  If sloop = 2 Then wavefile = Trim(App.Path & "\mloop02.wav")
  If psound = 1 Then Call PlaySound(wavefile, 0, SND_ASYNC + SND_LOOP)
  
  'MsgBox aco & "% " & bco & "% " & cco & "% " & dco & "% "
  ' enable background sound selection
  If son = 1 Then mtr1.Enabled = True
  If son = 2 Then mtr2.Enabled = True
   ' enable other Lifelines
  LfLine = 1
End Sub

Private Sub Imaata_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalA = 1 Then Exit Sub
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 1 Then Exit Sub
  Imaata.Picture = FrmGraphics!Imaatah.Picture
End Sub

Private Sub Imafa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Imano.Picture = FrmGraphics!Imanoff.Picture
  Imayes.Picture = FrmGraphics!Imayoff.Picture
End Sub

Private Sub Imano_Click()
  Fapass = 1
End Sub

Private Sub Imano_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Imano.Picture = FrmGraphics!Imanon.Picture
  Imayes.Picture = FrmGraphics!Imayoff.Picture
End Sub

Private Sub Imapaf_Click()
  If LfLine = 0 Or Lifeline = 0 Then Exit Sub
  If FinalA = 1 Then Exit Sub
  If spaf = 1 Then
    ' Enter msgbox here to say Lifeline already used
    Exit Sub
  End If
  Imapaf.Picture = FrmGraphics!Imapafs.Picture
  ' disable other Lifelines
  LfLine = 0
  ' disable background sound selection
  son = 0
  If mtr1.Enabled = True Then son = 1
  If mtr2.Enabled = True Then son = 2
  mtr1.Enabled = False
  mtr2.Enabled = False
  
  spaf = 1: cpaf = 1
  
  ' Play Phone and pause using tmr002
  If psound = 1 Then
    wavefile = Trim(App.Path & "\phone.wav")
    Call PlaySound(wavefile, 0, SND_ASYNC + SND_LOOP)
    t2Pass = 0
    Tmr002.Interval = 3000 + Int(Rnd * 2500)
    Tmr002.Enabled = True
    Do
      DoEvents
    Loop Until t2Pass = 1
    t2Pass = 0
    Tmr002.Interval = 700
    Tmr002.Enabled = True
    Do
      DoEvents
    Loop Until t2Pass = 1
    If sloop = 1 Then wavefile = Trim(App.Path & "\mloop01.wav")
    If sloop = 2 Then wavefile = Trim(App.Path & "\mloop02.wav")
    Call PlaySound(wavefile, 0, SND_ASYNC + SND_LOOP)
  End If
  
  Dim Cho As String
  Dim PafPer As String
  xx = 1 + Int(Rnd * 4)
  If xx = 1 Then
    PafPer = "John"
    Imapafper.Picture = FrmGraphics!Imajohn.Picture
  End If
  If xx = 2 Then
    PafPer = "Martin"
    Imapafper.Picture = FrmGraphics!Imamartin.Picture
  End If
  If xx = 3 Then
    PafPer = "Julie"
    Imapafper.Picture = FrmGraphics!Imajulie.Picture
  End If
  If xx = 4 Then
    PafPer = "Mary"
    Imapafper.Picture = FrmGraphics!Imamary.Picture
  End If
  Imapafper.Visible = True
  ' cQ = 14
  xx = 1 + Int(Rnd * 100)
  If PafPer = "John" Then
    If xx <= c(cQ, 1) Then
      ' Know the answer
      Cho = CorPos
    Else
      ' Guess the answer (1 in 4 chance)
      gpass = 0
      Do
        xx = 1 + Int(Rnd * 4)
        If xx = 1 And Act1 = 1 Then gpass = 1
        If xx = 2 And Act2 = 1 Then gpass = 1
        If xx = 3 And Act3 = 1 Then gpass = 1
        If xx = 4 And Act4 = 1 Then gpass = 1
      Loop Until gpass = 1
      Cho = Trim(Str$(xx))
    End If
  End If
  If PafPer = "Martin" Then
    If xx <= c(cQ, 2) Then
      ' Know the answer
      Cho = CorPos
    Else
      ' Guess the answer (1 in 2 chance)
      xx = 1 + Int(Rnd * 2)
      If xx = 1 Then
        Cho = CorPos
      Else
        gpass = 0
        Do
          yy = 1 + Int(Rnd * 4)
          If yy = 1 And Act1 = 1 Then gpass = 1
          If yy = 2 And Act2 = 1 Then gpass = 1
          If yy = 3 And Act3 = 1 Then gpass = 1
          If yy = 4 And Act4 = 1 Then gpass = 1
        Loop Until gpass = 1
        Cho = Trim(Str$(yy))
      End If
    End If
  End If
  If PafPer = "Julie" Then
    If xx <= c(cQ, 3) Then
      ' Know the answer
      Cho = CorPos
    Else
      ' Guess the answer (1 in 3 chance)
      xx = 1 + Int(Rnd * 3)
      If xx = 1 Then
        Cho = CorPos
      Else
        gpass = 0
        Do
          yy = 1 + Int(Rnd * 4)
          If yy = 1 And Act1 = 1 Then gpass = 1
          If yy = 2 And Act2 = 1 Then gpass = 1
          If yy = 3 And Act3 = 1 Then gpass = 1
          If yy = 4 And Act4 = 1 Then gpass = 1
        Loop Until gpass = 1
        Cho = Trim(Str$(yy))
      End If
    End If
  End If
  If PafPer = "Mary" Then
    If xx <= c(cQ, 4) Then
      ' Know the answer
      Cho = CorPos
    Else
      ' Guess the answer (1 in 2 chance)
      xx = 1 + Int(Rnd * 2)
      If xx = 1 Then
        Cho = CorPos
      Else
        gpass = 0
        Do
          yy = 1 + Int(Rnd * 4)
          If yy = 1 And Act1 = 1 Then gpass = 1
          If yy = 2 And Act2 = 1 Then gpass = 1
          If yy = 3 And Act3 = 1 Then gpass = 1
          If yy = 4 And Act4 = 1 Then gpass = 1
        Loop Until gpass = 1
        Cho = Trim(Str$(yy))
      End If
    End If
  End If
  
  'If cQ < 8 And Cho = CorPos Then
  If Cho = CorPos Then
    xx = 3
    If cQ > 5 Then
      xx = Int(Rnd * 4)
    End If
    If cQ = 15 Then
      xx = Int(Rnd * 2)
    End If
    If cQ > 10 And cQ <> 15 Then
      xx = Int(Rnd * 2)
      yy = Int(Rnd * 99)
      If yy > 93 Then
        xx = 3
      End If
      If PafPer = "Mary" And yy > 88 Then
        xx = 3
      End If
    End If
    'xx = 0
    yy = Int(Rnd * 3)
    If yy = 0 Then temp = "80%"
    If yy = 1 Then temp = "90%"
    If yy = 2 Then temp = "95%"
    If cQ > 12 Then
      If yy = 0 Then temp = "70%"
      If yy = 1 Then temp = "75%"
      If yy = 2 Then temp = "85%"
    End If
    txt = ""
    If xx = 0 Then
      txt = txt & "I'm not absolutely positive but I'm about " & temp & " sure the answer is " & ans(Cho)
    End If
    If xx = 2 Then
      txt = txt & "I know this one, It's definitely " & ans(Cho)
    End If
    If xx = 3 Then
      txt = txt & "It's easy when you know the answer: It's definitely " & ans(Cho)
    End If
    If xx = 1 Then
      zz = Int(Rnd * 2)
      If zz = 0 Then
        txt = txt & "I'm not sure but I think the answer is " & ans(Cho)
      End If
      If zz = 1 Then
        txt = txt & "I think the answer is " & ans(Cho)
      End If
    End If
  End If
  If Cho <> CorPos Then
    xx = Int(Rnd * 3)
    yy = Int(Rnd * 2)
    If cQ > 9 And xx <> 0 Then
      xx = Int(Rnd * 2)
    End If
    txt = ""
    If xx = 0 And yy = 0 Then
      txt = "Sorry, I don't know the answer to that question, and I wouldn't want to guess"
    End If
    If xx = 0 And yy = 1 Then
      txt = "Wow! That's a tough question... I don't think I know the answer"
      zz = Int(Rnd * 2)
      If zz = 1 Then txt = txt & " - You're on your own, Good luck!"
    End If
    If xx = 1 Then
      txt = txt & "I'm not sure but I think the answer is " & ans(Cho)
    End If
    If xx = 2 Then
      txt = txt & "I think the answer is " & ans(Cho)
    End If
  End If
  Lblpafin.Caption = txt
  
  Lblpafout.Visible = True
  Lblpafin.Visible = True
 
  ' enable background sound selection
  If son = 1 Then mtr1.Enabled = True
  If son = 2 Then mtr2.Enabled = True
   ' enable lifelines
  LfLine = 1
End Sub

Private Sub Imapaf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalA = 1 Then Exit Sub
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
  If spaf = 1 Then Exit Sub
  Imapaf.Picture = FrmGraphics!Imapafh.Picture
End Sub

Private Sub Imaquest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansa.Picture = FrmGraphics!Imaansa.Picture
    Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
    Imaansb.Picture = FrmGraphics!Imaansb.Picture
    Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
    Imaansc.Picture = FrmGraphics!Imaansc.Picture
    Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
    Imaansd.Picture = FrmGraphics!Imaansd.Picture
    Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
End Sub

Private Sub Imayes_Click()
  Fapass = 1
  FinalYN = 1
End Sub

Private Sub Imayes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Imayes.Picture = FrmGraphics!Imayon.Picture
  Imano.Picture = FrmGraphics!Imanoff.Picture
End Sub

Private Sub Lbla_Click()
  If Act1 = 0 Or Lifeline = 0 Or LfLine = 0 Then Exit Sub
  ' Add rountine to call Final Answer
  If FinalAon = 1 Then
    Call Final_Answer
    ' Once returned if yes then continue / if no then exit sub
    If FinalYN = 0 Then Exit Sub
  End If
  Imaansa.Picture = FrmGraphics!Imaansas.Picture
  Lbla.BackColor = &H8080FF: Lbla.ForeColor = &H0
  Choice = 1
End Sub

Private Sub Lbla_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Act1 = 0 Or LfLine = 0 Then Exit Sub
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansa.Picture = FrmGraphics!Imaansas.Picture
    Lbla.BackColor = &H8080FF: Lbla.ForeColor = &H0
    Imaansb.Picture = FrmGraphics!Imaansb.Picture
    Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
    Imaansc.Picture = FrmGraphics!Imaansc.Picture
    Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
    Imaansd.Picture = FrmGraphics!Imaansd.Picture
    Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
End Sub

Private Sub Lblb_Click()
  If Act2 = 0 Or Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalAon = 1 Then
    Call Final_Answer
    ' Once returned if yes then continue / if no then exit sub
    If FinalYN = 0 Then Exit Sub
  End If
  Imaansb.Picture = FrmGraphics!Imaansbs.Picture
  Lblb.BackColor = &H8080FF: Lblb.ForeColor = &H0
  Choice = 2
End Sub

Private Sub Lblb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Act2 = 0 Or LfLine = 0 Then Exit Sub
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansb.Picture = FrmGraphics!Imaansbs.Picture
    Lblb.BackColor = &H8080FF: Lblb.ForeColor = &H0
    Imaansa.Picture = FrmGraphics!Imaansa.Picture
    Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
    Imaansc.Picture = FrmGraphics!Imaansc.Picture
    Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
    Imaansd.Picture = FrmGraphics!Imaansd.Picture
    Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
End Sub

Private Sub Lblc_Click()
  If Act3 = 0 Or Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalAon = 1 Then
    Call Final_Answer
    ' Once returned if yes then continue / if no then exit sub
    If FinalYN = 0 Then Exit Sub
  End If
  Imaansc.Picture = FrmGraphics!Imaanscs.Picture
  Lblc.BackColor = &H8080FF: Lblc.ForeColor = &H0
  Choice = 3
End Sub

Private Sub Lblc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Act3 = 0 Or LfLine = 0 Then Exit Sub
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansc.Picture = FrmGraphics!Imaanscs.Picture
    Lblc.BackColor = &H8080FF: Lblc.ForeColor = &H0
    Imaansa.Picture = FrmGraphics!Imaansa.Picture
    Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
    Imaansb.Picture = FrmGraphics!Imaansb.Picture
    Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
    Imaansd.Picture = FrmGraphics!Imaansd.Picture
    Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
End Sub

Private Sub Lbld_Click()
  If Act4 = 0 Or Lifeline = 0 Or LfLine = 0 Then Exit Sub
  If FinalAon = 1 Then
    Call Final_Answer
    ' Once returned if yes then continue / if no then exit sub
    If FinalYN = 0 Then Exit Sub
  End If
  Imaansd.Picture = FrmGraphics!Imaansds.Picture
  Lbld.BackColor = &H8080FF: Lbld.ForeColor = &H0
  Choice = 4
End Sub

Private Sub Lbld_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Act4 = 0 Or LfLine = 0 Then Exit Sub
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansd.Picture = FrmGraphics!Imaansds.Picture
    Lbld.BackColor = &H8080FF: Lbld.ForeColor = &H0
    Imaansa.Picture = FrmGraphics!Imaansa.Picture
    Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
    Imaansb.Picture = FrmGraphics!Imaansb.Picture
    Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
    Imaansc.Picture = FrmGraphics!Imaansc.Picture
    Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
End Sub

Private Sub LblQuestion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Bs = 1 Then Exit Sub
  If FinalA = 0 Then
    Imaansa.Picture = FrmGraphics!Imaansa.Picture
    Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
    Imaansb.Picture = FrmGraphics!Imaansb.Picture
    Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
    Imaansc.Picture = FrmGraphics!Imaansc.Picture
    Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
    Imaansd.Picture = FrmGraphics!Imaansd.Picture
    Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  End If
  If s5050 = 0 Then
    Ima5050.Picture = FrmGraphics!Ima5050.Picture
  Else
    Ima5050.Picture = FrmGraphics!Ima5050s.Picture
  End If
  If spaf = 0 Then
    Imapaf.Picture = FrmGraphics!Imapaf.Picture
  Else
    Imapaf.Picture = FrmGraphics!Imapafs.Picture
  End If
  If sata = 0 Then
    Imaata.Picture = FrmGraphics!Imaata.Picture
  Else
    Imaata.Picture = FrmGraphics!Imaatas.Picture
  End If
  Imayes.Picture = FrmGraphics!Imayoff.Picture
  Imano.Picture = FrmGraphics!Imanoff.Picture
End Sub

Private Sub LblQuit_Click()
  If LfLine = 0 Or FinalA = 1 Then Exit Sub
  MoWon = mAmo(cQ - 1)
  retval = 0
  FrmMsgbox.Caption = "Millionaire"
  FrmMsgbox!Imatype = FrmGraphics!Imaq.Picture
  FrmMsgbox!CmdNo.Visible = True
  FrmMsgbox!CmdYes.Visible = True
  FrmMsgbox!CmdOK.Visible = False
  msgtxt = "Are you sure you wish to quit" & vbLf
  msgtxt = msgtxt & "and leave with " & MoWon
  FrmMsgbox!Lblmsg = msgtxt
  FrmMsgbox.Show 1
  If retval = 6 Then
    LblQuit.Visible = False
    Ttm = 1
    Call Clear_Top_Screen
    Call End_Of_Game
  End If
  Exit Sub
End Sub

Private Sub LblQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If FinalA = 1 Or LfLine = 0 Then Exit Sub
  LblQuit.ForeColor = &HFF&
  'LblQuit.BackColor = &HFFFFFF
End Sub

Private Sub mabout_Click()
  msgtxt = "Who Wants to be a Millionaire" & vbLf
  msgtxt = msgtxt & "     was written by D J King"
  msgtxt = msgtxt & vbLf & "         ( Copyright 2002 )"
  FrmAbout!LblInfo.Caption = msgtxt
  FrmAbout.Show 1
End Sub

Private Sub mconoff_Click()
  FinalAon = 0
  mconoff.Enabled = False
  mconon.Enabled = True
End Sub

Private Sub mconon_Click()
  FinalAon = 1
  mconon.Enabled = False
  mconoff.Enabled = True
End Sub

Private Sub mexit_Click()
  retval = 0
  FrmMsgbox.Caption = "Millionaire"
  FrmMsgbox!Imatype = FrmGraphics!Imaq.Picture
  FrmMsgbox!CmdNo.Visible = True
  FrmMsgbox!CmdYes.Visible = True
  FrmMsgbox!CmdOK.Visible = False
  msgtxt = "Do you really want to Exit the game and leave with nothing"
  FrmMsgbox!Lblmsg = msgtxt
  FrmMsgbox.Show 1
  If retval = 6 Then
    wavefile = Trim(App.Path & "\" & vbNullString)
    Call PlaySound(wavefile, 0, SND_ASYNC + SND_PURGE)
    Call Exit_Game
  End If
End Sub

Private Sub minfo_Click()
  FrmHelp.Show 1
End Sub

Private Sub msoff_Click()
  psound = 0
  wavefile = Trim(App.Path & "\" & vbNullString)
  Call PlaySound(wavefile, 0, SND_ASYNC + SND_PURGE)
  mson.Enabled = True
  msoff.Enabled = False
End Sub

Private Sub mson_Click()
  psound = 1
  If sloop = 1 Then wavefile = Trim(App.Path & "\mloop01.wav")
  If sloop = 2 Then wavefile = Trim(App.Path & "\mloop02.wav")
  If psound = 1 Then Call PlaySound(wavefile, 0, SND_ASYNC + SND_LOOP)
  mson.Enabled = False
  msoff.Enabled = True
End Sub

Private Sub mtr1_Click()
  sloop = 1
  mtr1.Enabled = False
  mtr2.Enabled = True
  wavefile = Trim(App.Path & "\mloop01.wav")
  If psound = 1 Then Call PlaySound(wavefile, 0, SND_ASYNC + SND_LOOP)
End Sub

Private Sub mtr2_Click()
  sloop = 2
  mtr2.Enabled = False
  mtr1.Enabled = True
  wavefile = Trim(App.Path & "\mloop02.wav")
  If psound = 1 Then Call PlaySound(wavefile, 0, SND_ASYNC + SND_LOOP)
End Sub

Private Sub Tmr001_Timer()
  tpass = 1
  Tmr001.Enabled = False
End Sub

Private Sub Tmr002_Timer()
  t2Pass = 1
  Tmr002.Enabled = False
End Sub

Public Sub Check_Question_Categorys(ww)

Dim vco(15)

For n = 1 To 15
  vco(n) = 0
Next n

For n = 1 To NumOfQuest
  xx = Val(Left(QA(n), 2))
  vco(xx) = vco(xx) + 1
Next n

For n = 1 To 15
  If vco(n) < 1 Then
    ww = 1
  End If
Next n

End Sub

Public Sub Create_Shuffle_File()

For n = 1 To 15
  Cn(n) = 0: Pn(n) = 1
Next n

For n = 1 To NumOfQuest
  xx = Val(Left(QA(n), 2))
  Cn(xx) = Cn(xx) + 1
Next n

For n = 1 To NumOfQuest
  Sh(n) = n
Next n

' Swap routine
For n = 1 To 15
  For k = 1 To NumOfQuest
    xx = k
    Do
      yy = 1 + Int(Rnd * NumOfQuest)
    Loop Until yy <> xx
    ' Swap
    temp = Sh(xx): Sh(xx) = Sh(yy): Sh(yy) = temp
  Next k
Next n

Open cpth For Output As #1
 
  For n = 1 To NumOfQuest
    Print #1, Sh(n)
  Next n
  For n = 1 To 15
    Print #1, Cn(n)
  Next n
  For n = 1 To 15
    Print #1, Pn(n)
  Next n
  
Close 1

End Sub

Public Sub Final_Answer()
  FinalYN = 0
  FinalA = 1
  If cpaf = 1 Then
    Lblpafout.Visible = False
    Lblpafin.Visible = False
    Imapafper.Visible = False
  End If
  If cata = 1 Then
    ' Delete Graph
    For n = 0 To 99
      Lblcola(n).Visible = False
      Lblcolb(n).Visible = False
      Lblcolc(n).Visible = False
      Lblcold(n).Visible = False
    Next n
    For n = 0 To 3
      Lblg(n).Visible = False
    Next n
    Imagraph.Visible = False
  End If
  Imafa.Picture = FrmGraphics!Imafinanson.Picture
  Imafa.Visible = True
  Imayes.Visible = True
  Imano.Visible = True
  
  Fapass = 0
  Do
    DoEvents
    
  Loop Until Fapass = 1
  If cpaf = 1 Then
    ' Show Paf
    Lblpafout.Visible = True
    Lblpafin.Visible = True
    Imapafper.Visible = True
  End If
  If cata = 1 Then
    ' Show Graph
    For n = 0 To aco - 1
      Lblcola(n).Visible = True
    Next n
    For n = 0 To bco - 1
      Lblcolb(n).Visible = True
    Next n
    For n = 0 To cco - 1
      Lblcolc(n).Visible = True
    Next n
    For n = 0 To dco - 1
      Lblcold(n).Visible = True
    Next n
    For n = 0 To 3
      Lblg(n).Visible = True
    Next n
    Imagraph.Visible = True
  End If
  FinalA = 0
  Imafa.Visible = False
  Imayes.Visible = False
  Imano.Visible = False
  
End Sub

Public Sub Exit_Game()
  If Version = 3 Then Call Output_Shuffle_File
  Unload FrmBoard2
  Unload FrmAbout
  Unload FrmGraphics
  Unload FrmMsgbox
  End
End Sub

Private Sub Tmr003_Timer()
' MsgBox t3Pass
If Switch = 1 Then
  If CorPos = "1" Then
    Imaansa.Picture = FrmGraphics!Imaansag.Picture
    Lbla.BackColor = &HFF00&: Lbla.ForeColor = &H0
  End If
  If CorPos = "2" Then
    Imaansb.Picture = FrmGraphics!Imaansbg.Picture
    Lblb.BackColor = &HFF00&: Lblb.ForeColor = &H0
  End If
  If CorPos = "3" Then
    Imaansc.Picture = FrmGraphics!Imaanscg.Picture
    Lblc.BackColor = &HFF00&: Lblc.ForeColor = &H0
  End If
  If CorPos = "4" Then
    Imaansd.Picture = FrmGraphics!Imaansdg.Picture
    Lbld.BackColor = &HFF00&: Lbld.ForeColor = &H0
  End If
End If
If Switch = 0 Then
  If CorPos = "1" Then
    Imaansa.Picture = FrmGraphics!Imaansab.Picture
    Lbla.BackColor = &H0&: Lbla.ForeColor = &HFFFFFF
  End If
  If CorPos = "2" Then
    Imaansb.Picture = FrmGraphics!Imaansbb.Picture
    Lblb.BackColor = &H0&: Lblb.ForeColor = &HFFFFFF
  End If
  If CorPos = "3" Then
    Imaansc.Picture = FrmGraphics!Imaanscb.Picture
    Lblc.BackColor = &H0&: Lblc.ForeColor = &HFFFFFF
  End If
  If CorPos = "4" Then
    Imaansd.Picture = FrmGraphics!Imaansdb.Picture
    Lbld.BackColor = &H0&: Lbld.ForeColor = &HFFFFFF
  End If
End If
t3Pass = t3Pass + 1
If Switch = 0 Then
  Switch = 1
Else
  Switch = 0
End If
End Sub

Public Sub Questions()

QA(1) = "01|A 'goatee' is a small type of what|Beard|Fork|Goat|Cucumber|sk62"
QA(2) = "01|One of the most senior judges in Britain is the 'Master of the...' what|Rolls|Muffins|Teacakes|Crumpets|sk44"
QA(3) = "01|Proverbially, what is rubbed into the wound to make things worse|Salt|Chocolate|Mayonnaise|Vinegar|sk21"
QA(4) = "01|What do diners in a restaurant use to take away their leftovers|Doggy bag|Piggy bag|Kitty bag|Bunny bag|sk85"
QA(5) = "01|What is the name of Channel 4's popular words and numbers game|Countdown|Meltdown|Lowdown|Eiderdown|se62"
QA(6) = "01|What name is given to a country's song played on official occasions|National anthem|National curriculum|National debt|National gallery|sk13"
QA(7) = "01|Which of these is a place in Australia|Alice Springs|Susie Leaps|Mary Jumps|Felicity Hops|sg13"
QA(8) = "01|Which of these is a popular toy|Pea-shooter|Carrot-firer|Cabbage-launcher|Turnip-zapper|sk56"
QA(9) = "01|Which of these is a type of dance|Cancan|Bambam|Tintin|Dondon|sk13"
QA(10) = "01|Which of these is a type of nut|Chestnut|Elbownut|Stomachnut|Shouldernut|sn56"
QA(11) = "01|Which of these means an eccentric person|Crackpot|Jackpot|Flowerpot|Chamberpot|sk85"
QA(12) = "01|Which of these phrases means to tease|Take the Mickey|Take the Minnie|Take the Goofy|Take the Donald|sk56"
QA(13) = "01|Which town in southwest England shares its name with something found in a bathroom|Bath|Plug|Sink|Bidet|sg56"
QA(14) = "02|Complete the saying used to draw attention to an innuendo: 'Nudge, nudge...'|Wink, wink|Wiggle, wiggle|Whisper, whisper|Woof, woof|sk85"
QA(15) = "02|Glucose is a form of which substance|Sugar|Salt|Oil|Acid|sn94"
QA(16) = "02|If someone is described as 'poker-faced', how are they looking|Expressionless|Sad|Happy|Excited|sk94"
QA(17) = "02|On which of these items would you be most likely to find a 'latch'|Door|Saucepan|Bed|Light bulb|sk39"
QA(18) = "02|What is the name for the last line of a joke|Punch line|Post line|Tick line|Flag line|sk39"
QA(19) = "02|What kind of food is sage|Herb|Fungus|Bean|Wine|sk56"
QA(20) = "02|Where on the body would a bracelet normally be worn|Wrist|Waist|Ankle|Neck|sk21"
QA(21) = "02|Which of these is a computer accessory|Mouse mat|Doormat|Beer mat|Table mat|sk94"
QA(22) = "02|Which of these is something that is very much in demand|Hot ticket|Boiling coupon|Blazing token|Scalding stub|sk62"
QA(23) = "02|Which of these words describes a person who carries out menial tasks for others|Dogsbody|Dogma|Doggerel|Doggy bag|sk44"
QA(24) = "02|Which of these would normally be worn on the feet|Flip-flops|Boleros|Trilbys|Slacks|sk62"
QA(25) = "02|Which part of the eye shares its name with a school student|Pupil|Iris|Retina|Cornea|sk94"
QA(26) = "03|'Plates of meat' is rhyming slang for what|Feet|Wheat|Treat|Heat|sk44"
QA(27) = "03|As what was Liberace famous|Pianist|Soldier|Bullfighter|Footballer|sk56"
QA(28) = "03|The phrase 'twenty-twenty' describes perfect... what|Vision|IQ|Hearing|Heart rate|sk56"
QA(29) = "03|What colour are hotels in Monopoly|Green|Red|Blue|Brown|sk72"
QA(30) = "03|What colour are houses in Monopoly|Green|Red|Blue|Brown|ss85"
QA(31) = "03|What do you do if you 'grease someone's palm'|Bribe someone|Shake hands|Give a warning|Tell a fortune|sk85"
QA(32) = "03|What is a yak|Long-haired Tibetan ox|Muslim veil|Boat|Sweet potato|sk56"
QA(33) = "03|What is the correct name for a set of drums|Kit|Gear|Pack|Tackle|se94"
QA(34) = "03|What is the surname of the brothers who make up the Bee Gees|Gibb|Gates|Gallagher|Gilbert|se56"
QA(35) = "03|When eaten, which part of a cow is known as tripe|Stomach lining|Rump|Neck|Tongue|sk77"
QA(36) = "03|Which of these is a county in the Republic of Ireland|Mayo|Febo|Juno|Sepo|sk39"
QA(37) = "04|How many sides does a rectangle have|Three|Four|Five|Six|nk67"
QA(38) = "04|Little Jimmy Osmond topped the UK charts with 'Long-Haired Lover from...' where|Liverpool|Lincoln|Lowestoft|Luton|se39"
QA(39) = "04|What do the five coloured rings on the Olympic flag represent|Continents|Countries|Centuries|Sports|sk21"
QA(40) = "04|Which Soho street was a centre of London fashion in the 1960s|Carnaby Street|Waldour Street|Berwick Street|Dean Street|sg62"
QA(41) = "04|Which bank has a black horse as its symbol|Lloyds TSB|Midland|Barclays|National Westminster|sk21"
QA(42) = "04|Which of these creatures is also the name of a spy|Mole|Badger|Fox|Ferret|sk39"
QA(43) = "04|Which of these is an alternative name for a 'dilemma'|Quandary|Quoit|Quince|Quarf|sk94"
QA(44) = "04|Which of these was a long-running TV quiz show|Mastermind|Master Plan|Masterpiece|Masterwork|se44"
QA(45) = "04|Which word describes someone who has an appealing manner or appearance on television|Telegenic|Telegraphic|Telepathic|Telescopic|sk39"
QA(46) = "04|Who created the animated series 'Creature Comforts'|Nick Park|Nick Lane|Nick Hancock|Nick Driver|se13"
QA(47) = "05|A blouse and a biscuit are named after which soldier|Garibaldi|Schwarzkopf|Napoleon|Wellington|sk94"
QA(48) = "05|For which national Rugby Union side has Gavin Hastings been a record scorer|Scotland|England|Wales|Ireland|ss13"
QA(49) = "05|In which TV comedy drama did Ian McShane play a roguish antiques dealer|Lovejoy|Fish|Spender|Boon|se94"
QA(50) = "05|What is the name of the little helicopter created by Sarah Ferguson|Budgie|Bertie|Barry|Billy|sk62"
QA(51) = "05|What nationality is the Formula One racing driver Mika Hakkinen|Finnish|Swedish|Danish|Norwegian|sk77"
QA(52) = "05|Which British island group lies just off the coast of Normandy|Channel Islands|Farne Islands|Orkney Islands|Shetland Islands|sg44"
QA(53) = "05|Which of these was a Greek god|Hermes|Benetton|Gucci|Prada|sk21"
QA(54) = "05|Which of these was an open-air venue for chariot racing in ancient Rome|Hippodrome|Palindrome|Velodrome|Aerodrome|sh56"
QA(55) = "05|Which of these women was a famous 19th-century cook|Mrs Beeton|Mrs Mills|Mrs Gaskell|Mrs Siddons|sh85"
QA(56) = "06|A quiver is a container for what|Arrows|Sword|Wine|Tea|sk56"
QA(57) = "06|Leonardo DiCaprio starred in which of these films|The Beach|The Sea of Sand|The Cruel Sea|The Swimmer|se44"
QA(58) = "06|On TV, Oz Clarke is best known as an expert on what subject|Wine|Horse racing|DIY|Fashion|se56"
QA(59) = "06|The Ewoks are a race of little people in which popular fantasy film|The Wizard of Oz|Chitty Chitty Bang Bang|Return of The Jedi|The Borrowers|se32"
QA(60) = "06|The Munchkins are a race of little people in which popular fantasy film|The Wizard of Oz|Chitty Chitty Bang Bang|Return of The Jedi|The Borrowers|se56"
QA(61) = "06|What is the name for the outer part of a citrus fruit|Zest|Zip|Zap|Zing|sn77"
QA(62) = "06|What type of creature is a pollack|Fish|Deer|Chicken|Beetle|sn62"
QA(63) = "06|Which of these is a type of ancient burial mound|Barrow|Coniston|Storth|Kendal|sh77"
QA(64) = "07|In which country is the city of Mecca|Saudi Arabia|India|Yemen|Israel|sg13"
QA(65) = "07|The 'sockeye' is a species of which fish|Salmon|Mackerel|Tuna|Eel|sn94"
QA(66) = "07|The Mariana Trench is the deepest part of which ocean|Pacific|Indian|Atlantic|Arctic|sg85"
QA(67) = "07|Typically, which type of literary form is an elegy|Poem|Play|Novel|Diary|sl62"
QA(68) = "07|Where in the world is El Salvador|Central America|Central Europe|Central Africa|Central Asia|sg56"
QA(69) = "07|Which of these European countries does not have a monarch|France|Spain|Norway|Sweden|sk21"
QA(70) = "07|Who played Granville in the TV sitcom 'Open All Hours'|David Jason|Nicholas Lyndhurst|Ronnie Barker|Richard Beckinsale|se39"
QA(71) = "08|In 2000, which author released his novel 'The Plant' on the Internet|Stephen King|James Herbert|Michael Crichton|John Grisham|sl21"
QA(72) = "08|In which industry are trade union branches known as 'chapels'|Printing|Mining|Shipbuilding|Farming|sk39"
QA(73) = "08|Where is T.S. Eliot's play 'Murder in the Cathedral' set|Canterbury Cathedral|York Minster|Salisbury Cathedral|St. Paul's Cathedral|se21"
QA(74) = "08|Who was the last Tsar of Russia|Nicholas II|Alexander III|Alexei|Feodor III|sg62"
QA(75) = "08|Who was the original presenter of the TV series 'Call My Bluff'|Frank Muir|Robin Ray|Robert Robinson|Hughie Green|se67"
QA(76) = "08|Who won a light-heavyweight boxing Olympic gold medal in 1960|Cassius Clay|James Boyd|George Foreman|Evander Holyfield|ss21"
QA(77) = "08|Who wrote 'Christine'|Dean Koontz|James Herbert|Stephen King|John Grisham|sl83"
QA(78) = "09|What is the more common name for the condition 'hypermetropia'|Long-sightedness|Sore throat|Migraine|Toothache|sn94"
QA(79) = "09|What was made by the Bessemer process|Steel|Glass|Paper|Bread|sn13"
QA(80) = "09|Which author is the creator and executive producer of the TV series 'E.R.'|Michael Crichton|Jack Kerouac|Joyce Carol Oates|Walt Whitman|se21"
QA(81) = "09|Which pastry is traditionally used to make the Greek dish baklava|Filo|Puff|Choux|Suet|sk94"
QA(82) = "09|Which pop singer sang the theme to the 1974 Bond film 'The Man With The Golden Gun'|Lulu|Shirley Bassey|Tom Jones|Sheena Easton|se21"
QA(83) = "09|Who wrote 'The Ballad of Reading Gaol'|Oscar Wilde|Robert Browning|Wilfred Owen|John Masefield|sl85"
QA(84) = "10|The word 'cataract' refers to what kind of geographical feature|Waterfall|Mountain|Lake|Cave|sg62"
QA(85) = "10|What is a 'Black-Eyed Susan'|Flower|Bird|Spider|Butterfly|sn77"
QA(86) = "10|Which London building stands on the site of Newgate Prison|Old Bailey|Westminster Cathedral|Battersea Power Station|Harrods|sk44"
QA(87) = "10|Which football team won the League and FA Cup double in 1961|Chelsea|Tottenham Hotspur|Arsenal|Crystal Palace|ss94"
QA(88) = "10|Which of these is the name of a type of small yellowish-red ant|Pharaoh|Pyramid|Cleopatra|Mummy|sn21"
QA(89) = "10|Which of these is the title of a Hitchcock film|Rope|Candlestick|Lead Piping|Dagger|se85"
QA(90) = "11|Which author wrote 'A Case of Need' under the pseudonym of Jeffery Hudson|Michael Crichton|Ian Fleming|P.D. James|Walt Whitman|sl62"
QA(91) = "11|Which body of water links the Black Sea and Mediterranean Sea|Sea of Marmara|Tyrrhenian Sea|Ligurian Sea|Ionian Sea|sg56"
QA(92) = "11|Who won the Tour De France in 1989 and 1990|Pedro Delgado|Miguel Indurain|Greg LeMond|Lance Armstrong|ss93"
QA(93) = "11|Who wrote 'Empire of the Sun'|J.G. Ballard|Michael Tolkin|Stephen Gallagher|L.M. Montgomery|sl94"
QA(94) = "11|Who wrote 'Ironweed'|William Kennedy|Fay Weldon|Jeanette Winterson|Stephen Gallagher|sl85"
QA(95) = "12|In which sport did Eric Heiden win five gold medals at the 1980 Olympic Winter Games in Lake Placid|Speed skating|Skiing|Curling|Ski Jumping|ss85"
QA(96) = "12|What sort of creature is a capercaillie|Bird|Fish|Snake|Lizard|sn39"
QA(97) = "12|What was the name of the first yacht to win sailing's America's Cup in 1851|America|Atlanta|Australia|Endeavour|ss13"
QA(98) = "12|What was the title of John Berger's 1972 Booker Prize winning novel|G.|Y.|K.|M.|sl21"
QA(99) = "13|In mathematics, which prefix refers to 10 to the power of minus 9|Nano|Femto|Atto|Pico|sn77"
QA(100) = "13|The Egyptian goddess Bastet was depicted with the head of which creature|Cat|Jackal|Falcon|Crocodile|sh85"
QA(101) = "13|Who was the first King of Belgium|Leopold I|Albert|Frederick I|Baudouin|sh13"
QA(102) = "13|Who was the first man to break the 27 minute barrier in the 10,000 metres|Emil Zatopek|Lasse Viren|Yobes Ondieki|Brahim Boutayeb|ss45"
QA(103) = "14|As of 2003, who is the only person in Olympic history to have earned gold medals in both Summer and Winter sports|Eddie Eagan|Billy Fiske|Gillis Grafström|Pierre Brunet|ss39"
QA(104) = "14|What instrument was played by jazz musician Chet Baker|Trumpet|Piano|Saxophone|Guitar|se85"
QA(105) = "14|What was the former name of Burkina Faso in Africa|Upper Volta|Ivory Coast|Gold Coast|Dahomey|sg56"
QA(106) = "14|Which of these scientists was a co-discoverer of nuclear fission|Otto Hahn|James Chadwick|Ernest Rutherford|Robert Goddard|sn44"
QA(107) = "15|Calamine, used as an ointment, contains a carbonate of which element|Zinc|Calcium|Magnesium|Sodium|sn56"
QA(108) = "15|Which French author wrote 'The Outsider'|Albert Camus|Simone de Beauvoir|Marguerite Duras|Jean-Paul Sartre|sl77"
QA(109) = "15|Which town was the capital of the British king Cymbeline|Colchester|Cambridge|Chester|Coventry|sh21"

NumOfQuest = 109
End Sub
