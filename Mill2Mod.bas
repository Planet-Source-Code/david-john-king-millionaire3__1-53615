Attribute VB_Name = "Mill2Mod"
Option Explicit
' For checking existance of a file
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
' For Sleeping
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' ***********************************************************************************

' For Playing Sound - specify wavefile name / path to file / call subroutine
' wavefile = Trim(App.Path & "\?????.wav")
' If psound = 1 Then Call PlaySound(wavefile, 0, SND_ASYNC + SND_NODEFAULT)
'
Public Declare Function PlaySound _
       Lib "winmm.dll" Alias "PlaySoundA" _
       (ByVal lpszName As String, _
       ByVal hModule As Long, _
       ByVal dwFlags As Long) As Long

Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10
Public Const SND_SYNC = &H0
Public Const SND_ALIAS = &H10000
Public Const SND_APPLICATION = &H80
Public Const SND_ALIAS_ID = &H110000
Public Const SND_FILENAME = &H20000
Public Const SND_NOWAIT = &H2000
Public Const SND_PURGE = &H40
Public Const SND_RESOURCE = &H40004
Global wavefile As String ' Name of wavefile to play
Global psound As Integer ' Sound Off / On var
' *************************************************************

Global Pass As Integer ' var to control pause loops
Global t2Pass As Integer
Global t3Pass As Integer
Global Switch As Integer
Global QA(109) As String ' Question & Answer Strings
Global Sh(109) As Integer ' Shuffle id's
Global Cn(15) As Integer
Global Pn(15) As Integer
Global L(15) As Integer ' Lower Percentage for ATA
Global U(15) As Integer ' Upper Percentage for ATA
Global Qacd(9) As String
Global CodNum(57) As Integer
Global PickSt01 As String ' var for checking correct code for a b c or d
Global mAmo(15) As String ' Money awarded
Global NumOfQuest As Integer ' Number of Questions
Global cQ As Integer ' Count of Current question being asked
Global cQnum ' Number of Question and Answer selected
Global cQuest As String ' Question string for lblquestion
Global cAnsA As String ' Answer A for lbla
Global cAnsB As String ' Answer B for lblb
Global cAnsC As String ' Answer C for lblc
Global cAnsD As String ' Answer D for lbld
Global CorPos As String ' Correct Answer A B C or D
Global Choice As Integer ' Players Guess A B C or D
Global s5050 As Integer ' 50/50 var to restrict to 1
Global c5050 As Integer ' var = 1 if 50/50 selected - used for graph
Global spaf As Integer ' PAF var to restrict to 1
Global cpaf As Integer ' PAF var to delete PAF Lable
Global sata As Integer ' ATA var to restriict to 1
Global cata As Integer ' ATA var to erase graph
Global Act1 As Integer ' var to control activation of answer A for 50/50
Global Act2 As Integer ' var to control activation of answer B for 50/50
Global Act3 As Integer ' var to control activation of answer C for 50/50
Global Act4 As Integer ' var to control activation of answer D for 50/50
Global tpass As Integer ' var used in tmr001 to control wait times
Global pafName(4) As String ' Names of the Phone a Friends
Global c(15, 4) As Integer ' vars used for % chance for PAF's
Global ans(4) As String ' vars A B C & D
Global Bs As Integer ' var to stop highlighting on mouse moves
Global sID As String ' Sort ID = "s" then sort answers
Global Lifeline As Integer ' Var to control activation of Lifelines
Global FileInputName As String ' Name of Input Question File
Global msgtxt As String ' Var used for Msgbox text
Global optval As String ' Var used for Msgbox text
Global retval As Integer ' Var used for Msgbox text
Global cpth As String ' Current Path to Millionaire directory
Global Ttm As Integer ' Var for Taking the money
Global Fapass As Integer ' Var used for final answer
Global FinalA As Integer ' Var fot final answer
Global FinalAon As Integer ' Var for final answer = 1 then final answer on
Global FinalYN As Integer ' Var for final answer 0 = no 1 = yes
Global LfLine As Integer ' controls selection of lifelines
Global MoWon As String ' Amount of money won
Global aco As Integer
Global bco As Integer
Global cco As Integer
Global dco As Integer
Global sloop As Integer ' Sound Loop var
Global Version As Integer ' Version Indicator
Global Disp As Integer ' Used for help display


' Wave Files Used
' Letsplay.wav / mloop01.wav / mloop02.wav / Answer.wav
' 5050.wav / Askaud01.wav / phone.wav / nextq.wav

Public Sub Main()
  FrmBoard2.Show
End Sub


Public Function FileExists(strPath As String) As Boolean
  Dim Exist As Integer
  If PathFileExists(strPath) = 1 Then
    FileExists = True
  Else
    FileExists = False
  End If
End Function

Public Sub Shuffle_Questions()
Dim n, k, xx, yy As Integer
Dim temp As String
  For n = 1 To 15
    For k = 1 To NumOfQuest
      xx = k
      Do
        yy = 1 + Int(Rnd * NumOfQuest)
      Loop Until yy <> xx
      ' Swap
      temp = QA(xx): QA(xx) = QA(yy): QA(yy) = temp
    Next k
  Next n

End Sub

Public Sub Select_question()
Dim ch, Tcor, Tcor3 As String
Dim CCAnsa, CCAnsb, CCAnsc, CCAnsd As String
Dim Tcor1, Tcor2, ql As Integer
Dim Hit, temp, st, ls, n, j, xx As Integer
temp = cQ
ch = Trim(Str$(cQ))
If Len(ch) = 1 Then ch = "0" & ch

For n = 1 To NumOfQuest
  If Left(QA(n), 2) = ch Then cQnum = n: Exit For
Next n

If Version = 3 Then
  Hit = 0
  For n = 1 To NumOfQuest
     If Left(QA(Sh(n)), 2) = ch Then
      Hit = Hit + 1
      If Hit = Pn(ch) Then
        cQnum = Sh(n)
        Pn(ch) = Pn(ch) + 1
        ' Check that pn(?) is < cn(?) if it is OK
        ' - if not reset pn(?) = 1
        If Pn(ch) > Cn(ch) Then
          Pn(ch) = 1
        End If
        Exit For
      End If
    End If
  Next n
End If

' cQnum = number of Question and Answer selected
st = 4: ls = 0
For n = st To Len(QA(cQnum))
  If Mid(QA(cQnum), n, 1) = "|" Then
    Exit For
  End If
  ls = ls + 1
Next n

cQuest = Mid(QA(cQnum), st, ls) & "?"

st = n + 1

ls = 0
For n = st To Len(QA(cQnum))
  If Mid(QA(cQnum), n, 1) = "|" Then
    Exit For
  End If
  ls = ls + 1
Next n

cAnsA = Mid(QA(cQnum), st, ls)

st = n + 1

ls = 0
For n = st To Len(QA(cQnum))
  If Mid(QA(cQnum), n, 1) = "|" Then
    Exit For
  End If
  ls = ls + 1
Next n

cAnsB = Mid(QA(cQnum), st, ls)

st = n + 1

ls = 0
For n = st To Len(QA(cQnum))
  If Mid(QA(cQnum), n, 1) = "|" Then
    Exit For
  End If
  ls = ls + 1
Next n

cAnsC = Mid(QA(cQnum), st, ls)

st = n + 1

ls = 0
For n = st To Len(QA(cQnum))
  If Mid(QA(cQnum), n, 1) = "|" Then
    Exit For
  End If
  ls = ls + 1
Next n

cAnsD = Mid(QA(cQnum), st, ls)

Tcor = Right(QA(cQnum), 2)
Tcor1 = Val(Left(Tcor, 1))
Tcor2 = Val(Right(Tcor, 1))
Tcor3 = Mid(Qacd(Tcor1), Tcor2, 1)

If Tcor3 = "A" Then CorPos = 1
If Tcor3 = "B" Then CorPos = 2
If Tcor3 = "C" Then CorPos = 3
If Tcor3 = "D" Then CorPos = 4

ql = Len(QA(cQnum)) - 3
sID = Mid(QA(cQnum), ql, 1)
'MsgBox sID

If sID = "s" Then
  Dim i(4)

  For n = 1 To 4: i(n) = 0: Next n

  j = 1
  For n = 1 To 4
    If j = 5 Then Exit For
    Do
      xx = 1 + Int(Rnd * 4)
    Loop Until i(xx) = 0
    i(xx) = j: j = j + 1
  Next n

  For j = 1 To 1
    If i(1) = CorPos Then CorPos = 1: Exit For
    If i(2) = CorPos Then CorPos = 2: Exit For
    If i(3) = CorPos Then CorPos = 3: Exit For
    If i(4) = CorPos Then CorPos = 4: Exit For
  Next j

  CCAnsa = cAnsA
  CCAnsb = cAnsB
  CCAnsc = cAnsC
  CCAnsd = cAnsD

  ' Shuffle Answers
  If i(1) = 1 Then cAnsA = CCAnsa
  If i(1) = 2 Then cAnsA = CCAnsb
  If i(1) = 3 Then cAnsA = CCAnsc
  If i(1) = 4 Then cAnsA = CCAnsd

  If i(2) = 1 Then cAnsB = CCAnsa
  If i(2) = 2 Then cAnsB = CCAnsb
  If i(2) = 3 Then cAnsB = CCAnsc
  If i(2) = 4 Then cAnsB = CCAnsd

  If i(3) = 1 Then cAnsC = CCAnsa
  If i(3) = 2 Then cAnsC = CCAnsb
  If i(3) = 3 Then cAnsC = CCAnsc
  If i(3) = 4 Then cAnsC = CCAnsd

  If i(4) = 1 Then cAnsD = CCAnsa
  If i(4) = 2 Then cAnsD = CCAnsb
  If i(4) = 3 Then cAnsD = CCAnsc
  If i(4) = 4 Then cAnsD = CCAnsd
End If

FrmBoard2!LblQuestion.Caption = cQuest
FrmBoard2!LblQuestion.Visible = True
FrmBoard2!Lbla.Visible = True
FrmBoard2!Lblb.Visible = True
FrmBoard2!Lblc.Visible = True
FrmBoard2!Lbld.Visible = True
FrmBoard2!Lbla.Caption = cAnsA: ans(1) = "A : '" & cAnsA & "'"
FrmBoard2!Lblb.Caption = cAnsB: ans(2) = "B : '" & cAnsB & "'"
FrmBoard2!Lblc.Caption = cAnsC: ans(3) = "C : '" & cAnsC & "'"
FrmBoard2!Lbld.Caption = cAnsD: ans(4) = "D : '" & cAnsD & "'"

Act1 = 1: Act2 = 1: Act3 = 1: Act4 = 1

End Sub
Public Sub Check_Question_File()
Dim ch1 As String
Dim n, j, k, co, xx As Integer
ch1 = "0123456789"
For n = 1 To NumOfQuest
  ' check first 2 digits are between 0-9
  co = 0
  For j = 1 To 2
    For k = 1 To Len(ch1)
      If Mid(QA(n), j, 1) = Mid(ch1, k, 1) Then
        co = co + 1: Exit For
      End If
    Next k
  Next j
  If co <> 2 Then
    GoTo FileError
  End If
  co = 0
  ' Check for 6 x |
  For j = 1 To Len(QA(n)) - 2
    If Mid(QA(n), j, 1) = "|" Then
      co = co + 1
    End If
  Next j
  If co <> 6 Then
    GoTo FileError
  End If
  ' Check sort = s or n
  co = 0
  xx = Len(QA(n))
  'MsgBox Mid(QA(n), xx - 3, 1)
  If Mid(QA(n), xx - 3, 1) = "s" Or Mid(QA(n), xx - 3, 1) = "n" Then
    co = 1
  End If
  If co <> 1 Then
    GoTo FileError
  End If
  ' Check for Category code
  ' MsgBox Mid(QA(n), xx - 1, 1)
  co = 0
  If Mid(QA(n), xx - 2, 1) = "e" Then co = 1
  If Mid(QA(n), xx - 2, 1) = "g" Then co = 1
  If Mid(QA(n), xx - 2, 1) = "h" Then co = 1
  If Mid(QA(n), xx - 2, 1) = "l" Then co = 1
  If Mid(QA(n), xx - 2, 1) = "n" Then co = 1
  If Mid(QA(n), xx - 2, 1) = "s" Then co = 1
  If Mid(QA(n), xx - 2, 1) = "k" Then co = 1
  If co <> 1 Then
    GoTo FileError
  End If
  
  ' Check for A-D from code string
  'MsgBox Mid(QA(n), xx - 1, 2)
  co = 0
  For j = 1 To 57
    If CodNum(j) = Mid(QA(n), xx - 1, 2) Then
      'MsgBox CodNum(j)
      co = 1: Exit For
    End If
  Next j
  
  If co <> 1 Then
    GoTo FileError
  End If
  
Next n
Exit Sub

FileError:
msgtxt = "Question Input File Corrupt"
optval = vbCritical + vbOKOnly
retval = MsgBox(msgtxt, optval, "Input Error")
End

End Sub

Public Sub Output_Shuffle_File()
Dim n As Integer
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

Public Sub Input_Question_File()
Dim i As Integer
Open cpth For Input As #1
 
  i = 1
  Do While Not EOF(1)
    Line Input #1, QA(i)
    i = i + 1
  Loop
  NumOfQuest = i - 1
  Close #1

End Sub
