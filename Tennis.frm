VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label P2PW 
      AutoSize        =   -1  'True
      Caption         =   "Points Won"
      Height          =   195
      Left            =   3000
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label P2DF 
      AutoSize        =   -1  'True
      Caption         =   "Double Faults"
      Height          =   195
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label P2A 
      AutoSize        =   -1  'True
      Caption         =   "Aces"
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.Label P1PW 
      AutoSize        =   -1  'True
      Caption         =   "Points Won"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label P1DF 
      AutoSize        =   -1  'True
      Caption         =   "Double Faults"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label P1A 
      AutoSize        =   -1  'True
      Caption         =   "Aces"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Score"
      Height          =   195
      Left            =   2070
      TabIndex        =   1
      Top             =   600
      Width           =   435
   End
   Begin VB.Label lblPlayers 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Players"
      Height          =   195
      Left            =   2025
      TabIndex        =   0
      Top             =   240
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnPlay_Click()
Dim Toss As Integer
Dim Rno As Integer
Dim GameSet As Integer
Dim P1Fail As Integer
Dim P2Fail As Integer
Dim n As String
Dim l, s1, s2, s3, s4, s5, s6, s7, s8, s9, s10 As Integer
Open "players.txt" For Input As #1
l = l + 1
Do Until EOF(1)
Input #1, n
Input #1, s1
Input #1, s2
Input #1, s3
Input #1, s4
Input #1, s5
Input #1, s6
Input #1, s7
Input #1, s8
Input #1, s9
Input #1, s10
If l = 1 Then
P1Name = n
P1Serve = s1
P1Volley = s2
P1HVolley = s3
P1OHead = s4
P1Lob = s5
P1DShot = s6
P1FHand = s7
P1BHand = s8
P1Speed = s9
P1Strategy = s10
l = l + 1
End If
If l = 2 Then
P2Name = n
P2Serve = s1
P2Volley = s2
P2HVolley = s3
P2OHead = s4
P2Lob = s5
P2DShot = s6
P2FHand = s7
P2BHand = s8
P2Speed = s9
P2Strategy = s10
End If
Loop
Close 1

 P1Fault = 0
 P1Score = 0
 P1Game = 0
 P1Set = 0
 P1Shot = 0
 P1PWon = 0
P1Aces = 0
P1DFault = 0
 
 P2Fault = 0
 P2Score = 0
 P2Game = 0
 P2Set = 0
 P2Shot = 0
 P2PWon = 0
P2Aces = 0
P2DFault = 0
 
 Score = ""
 Serve = 0
 Match = 0
 
TB = 0
TBCount = 0


lblScore.Caption = ""
 
lblPlayers.Caption = P1Name + " vs. " + P2Name
Toss:
Toss = Int(2 * Rnd) + 1
If Toss = 1 Then
Serve = 1
GoTo P1Server
Else
Serve = 2
GoTo P2Server
End If

P1Server:

If P1Serve >= 0 And P1Serve <= 4 Then
Rno = Int(10 * Rnd) + 1
Else
Rno = Int(P1Serve * Rnd) + 1
End If

If (P1Serve >= 0 And P1Serve <= 4) And (Rno >= 1 And Rno <= 9) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 5 And P1Serve <= 10) And (Rno >= 1 And Rno <= 4) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 11 And P1Serve <= 13) And (Rno >= 1 And Rno <= 5) Then
GoTo P1ServeFail

ElseIf (P1Serve = 14 Or P1Serve = 15) And (Rno >= 1 And Rno <= 6) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 16 And P1Serve <= 18) And (Rno >= 1 And Rno <= 7) Then
GoTo P1ServeFail

ElseIf (P1Serve = 19 Or P1Serve = 20) And (Rno >= 1 And Rno <= 8) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 21 And P1Serve <= 23) And (Rno >= 1 And Rno <= 9) Then
GoTo P1ServeFail

ElseIf (P1Serve = 24 Or P1Serve = 25) And (Rno >= 1 And Rno <= 10) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 26 And P1Serve <= 28) And (Rno >= 1 And Rno <= 11) Then
GoTo P1ServeFail

ElseIf (P1Serve = 29 Or P1Serve = 30) And (Rno >= 1 And Rno <= 12) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 31 And P1Serve <= 33 And P1Fault = 0) And (Rno >= 1 And Rno <= 13) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 31 And P1Serve <= 33 And P1Fault = 1) And (Rno >= 1 And Rno <= 7) Then
GoTo P1ServeFail

ElseIf (P1Serve = 34 Or P1Serve = 35 And P1Fault = 0) And (Rno >= 1 And Rno <= 14) Then
GoTo P1ServeFail

ElseIf (P1Serve = 34 Or P1Serve = 35 And P1Fault = 1) And (Rno >= 1 And Rno <= 7) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 36 And P1Serve <= 38 And P1Fault = 0) And (Rno >= 1 And Rno <= 15) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 36 And P1Serve <= 38 And P1Fault = 1) And (Rno >= 1 And Rno <= 8) Then
GoTo P1ServeFail

ElseIf (P1Serve = 39 Or P1Serve = 40 And P1Fault = 0) And (Rno >= 1 And Rno <= 16) Then
GoTo P1ServeFail

ElseIf (P1Serve = 39 Or P1Serve = 40 And P1Fault = 1) And (Rno >= 1 And Rno <= 8) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 41 And P1Serve <= 43 And P1Fault = 0) And (Rno >= 1 And Rno <= 17) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 41 And P1Serve <= 43 And P1Fault = 1) And (Rno >= 1 And Rno <= 9) Then
GoTo P1ServeFail

ElseIf (P1Serve = 44 Or P1Serve = 45 And P1Fault = 0) And (Rno >= 1 And Rno <= 18) Then
GoTo P1ServeFail

ElseIf (P1Serve = 44 Or P1Serve = 45 And P1Fault = 1) And (Rno >= 1 And Rno <= 9) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 46 And P1Serve <= 48 And P1Fault = 0) And (Rno >= 1 And Rno <= 19) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 46 And P1Serve <= 48 And P1Fault = 1) And (Rno >= 1 And Rno <= 10) Then
GoTo P1ServeFail

ElseIf (P1Serve = 49 Or P1Serve = 50 And P1Fault = 0) And (Rno >= 1 And Rno <= 20) Then
GoTo P1ServeFail

ElseIf (P1Serve = 49 Or P1Serve = 50 And P1Fault = 1) And (Rno >= 1 And Rno <= 10) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 51 And P1Serve <= 55 And P1Fault = 0) And (Rno >= 1 And Rno <= 19) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 51 And P1Serve <= 55 And P1Fault = 1) And (Rno >= 1 And Rno <= 10) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 56 And P1Serve <= 60 And P1Fault = 0) And (Rno >= 1 And Rno <= 18) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 56 And P1Serve <= 60 And P1Fault = 1) And (Rno >= 1 And Rno <= 9) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 61 And P1Serve <= 65 And P1Fault = 0) And (Rno >= 1 And Rno <= 17) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 61 And P1Serve <= 65 And P1Fault = 1) And (Rno >= 1 And Rno <= 9) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 66 And P1Serve <= 70 And P1Fault = 0) And (Rno >= 1 And Rno <= 16) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 66 And P1Serve <= 70 And P1Fault = 1) And (Rno >= 1 And Rno <= 8) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 71 And P1Serve <= 75 And P1Fault = 0) And (Rno >= 1 And Rno <= 15) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 71 And P1Serve <= 75 And P1Fault = 1) And (Rno >= 1 And Rno <= 8) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 76 And P1Serve <= 80 And P1Fault = 0) And (Rno >= 1 And Rno <= 14) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 76 And P1Serve <= 80 And P1Fault = 1) And (Rno >= 1 And Rno <= 7) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 81 And P1Serve <= 85 And P1Fault = 0) And (Rno >= 1 And Rno <= 13) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 81 And P1Serve <= 85 And P1Fault = 1) And (Rno >= 1 And Rno <= 7) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 86 And P1Serve <= 90 And P1Fault = 0) And (Rno >= 1 And Rno <= 12) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 86 And P1Serve <= 90 And P1Fault = 1) And (Rno >= 1 And Rno <= 6) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 91 And P1Serve <= 95 And P1Fault = 0) And (Rno >= 1 And Rno <= 11) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 91 And P1Serve <= 95 And P1Fault = 1) And (Rno >= 1 And Rno <= 6) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 96 And P1Serve <= 100 And P1Fault = 0) And (Rno >= 1 And Rno <= 10) Then
GoTo P1ServeFail

ElseIf (P1Serve >= 96 And P1Serve <= 100 And P1Fault = 1) And (Rno >= 1 And Rno <= 5) Then
GoTo P1ServeFail

Else

If (P1Serve >= 31 And P1Serve <= 33 And P1Fault = 0) And (Rno = 14) Then
GoTo P1Ace

ElseIf (P1Serve = 34 Or P1Serve = 35) And (P1Fault = 0 And Rno = 15) Then
GoTo P1Ace

ElseIf (P1Serve >= 36 And P1Serve <= 38 And P1Fault = 0) And (Rno = 16) Then
GoTo P1Ace

ElseIf (P1Serve = 39 Or P1Serve = 40) And (P1Fault = 0 And Rno = 17) Then
GoTo P1Ace

ElseIf (P1Serve >= 41 And P1Serve <= 43 And P1Fault = 0) And (Rno = 18) Then
GoTo P1Ace

ElseIf (P1Serve = 44 Or P1Serve = 45) And (P1Fault = 0 And Rno = 19) Then
GoTo P1Ace

ElseIf (P1Serve >= 46 And P1Serve <= 48 And P1Fault = 0) And (Rno = 20) Then
GoTo P1Ace

ElseIf (P1Serve = 49 Or P1Serve = 50) And (P1Fault = 0 And Rno = 21) Then
GoTo P1Ace

ElseIf (P1Serve >= 51 And P1Serve <= 55 And P1Fault = 0) And (Rno = 20 Or Rno = 21 Or Rno = 22) Then
GoTo P1Ace

ElseIf (P1Serve >= 51 And P1Serve <= 55 And P1Fault = 1) And (Rno = 11) Then
GoTo P1Ace

ElseIf (P1Serve >= 56 And P1Serve <= 60 And P1Fault = 0) And (Rno = 19 Or Rno = 20 Or Rno = 21 Or Rno = 22) Then
GoTo P1Ace

ElseIf (P1Serve >= 56 And P1Serve <= 60 And P1Fault = 1) And (Rno = 10 Or Rno = 11) Then
GoTo P1Ace

ElseIf (P1Serve >= 61 And P1Serve <= 65 And P1Fault = 0) And (Rno >= 18 And Rno <= 23) Then
GoTo P1Ace

ElseIf (P1Serve >= 61 And P1Serve <= 65 And P1Fault = 1) And (Rno >= 10 And Rno <= 12) Then
GoTo P1Ace

ElseIf (P1Serve >= 66 And P1Serve <= 70 And P1Fault = 0) And (Rno >= 17 And Rno <= 24) Then
GoTo P1Ace

ElseIf (P1Serve >= 66 And P1Serve <= 70 And P1Fault = 1) And (Rno >= 9 And Rno <= 12) Then
GoTo P1Ace

ElseIf (P1Serve >= 71 And P1Serve <= 75 And P1Fault = 0) And (Rno >= 16 And Rno <= 25) Then
GoTo P1Ace

ElseIf (P1Serve >= 71 And P1Serve <= 75 And P1Fault = 1) And (Rno >= 9 And Rno <= 13) Then
GoTo P1Ace

ElseIf (P1Serve >= 76 And P1Serve <= 80 And P1Fault = 0) And (Rno >= 15 And Rno <= 26) Then
GoTo P1Ace

ElseIf (P1Serve >= 76 And P1Serve <= 80 And P1Fault = 1) And (Rno >= 8 And Rno <= 13) Then
GoTo P1Ace

ElseIf (P1Serve >= 81 And P1Serve <= 85 And P1Fault = 0) And (Rno >= 14 And Rno <= 27) Then
GoTo P1Ace

ElseIf (P1Serve >= 81 And P1Serve <= 85 And P1Fault = 1) And (Rno >= 8 And Rno <= 14) Then
GoTo P1Ace

ElseIf (P1Serve >= 86 And P1Serve <= 90 And P1Fault = 0) And (Rno >= 13 And Rno <= 28) Then
GoTo P1Ace

ElseIf (P1Serve >= 86 And P1Serve <= 90 And P1Fault = 1) And (Rno >= 7 And Rno <= 14) Then
GoTo P1Ace

ElseIf (P1Serve >= 91 And P1Serve <= 95 And P1Fault = 0) And (Rno >= 12 And Rno <= 29) Then
GoTo P1Ace

ElseIf (P1Serve >= 91 And P1Serve <= 95 And P1Fault = 1) And (Rno >= 7 And Rno <= 15) Then
GoTo P1Ace

ElseIf (P1Serve >= 91 And P1Serve <= 95 And P1Fault = 0) And (Rno >= 12 And Rno <= 29) Then
GoTo P1Ace

ElseIf (P1Serve >= 91 And P1Serve <= 95 And P1Fault = 1) And (Rno >= 6 And Rno <= 15) Then
GoTo P1Ace

Else
P1Fault = 0
GoTo P2Move
End If
End If

P1ServeFail:
P1Fault = P1Fault + 1
If P1Fault = 2 Then
P1DFault = P1DFault + 1
P2PWon = P2PWon + 1
P1Fault = 0
P2Score = P2Score + 1
If TB = 1 Then
TBCount = TBCount + 1
If P2Score >= 7 And P2Score - P1Score >= 2 Then
P2Set = P2Set + 1
P2Game = P2Game + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + "(" + Str(P1Score) + "-" + Str(P2Score) + ")" + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
TB = 0
TBCount = 0
End If
Else
If P2Score >= 4 And P2Score - P1Score >= 2 Then
P2Game = P2Game + 1
P1Score = 0
P2Score = 0
P1Fault = 0
P2Fault = 0
GameSet = 1
End If
If P2Game - P1Game > 1 And P2Game = 6 Or P2Game - P1Game >= 1 And P2Game = 7 Then
P2Set = P2Set + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
End If
End If
Else
GoTo P1Server
End If

If P1Game = 6 And P2Game = 6 Then
TB = 1
End If

If GameSet = 1 And (P1Set = 3 Or P2Set = 3) Then
GameSet = 0
Match = 1
lblScore.Caption = Score
P1PW.Caption = Str(P1PWon)
P1A.Caption = Str(P1Aces)
P1DF.Caption = Str(P1DFault)
P2PW.Caption = Str(P2PWon)
P2A.Caption = Str(P2Aces)
P2DF.Caption = Str(P2DFault)
Exit Sub
End If

If TB = 1 Then
If TBCount = 0 Then
Serve = 2
GoTo P2Server
ElseIf TBCount = 1 Then
Serve = 1
GoTo P1Server
ElseIf TBCount = 2 Then
TBCount = 0
Serve = 2
GoTo P2Server
End If
End If

If GameSet = 1 And Match = 0 And TB = 0 Then
GameSet = 0
If Serve = 1 Then
Serve = 2
GoTo P2Server
Else
Serve = 1
GoTo P1Server
End If
ElseIf GameSet = 0 And Match = 0 And TB = 0 Then
If Serve = 1 Then
GoTo P1Server
Else
GoTo P2Server
End If
End If

P1Ace:
P1Fault = 0
P1Aces = P1Aces + 1
P1PWon = P1PWon + 1
P1Score = P1Score + 1
If TB = 1 Then
TBCount = TBCount + 1
If P1Score >= 7 And P1Score - P2Score >= 2 Then
P1Set = P1Set + 1
P1Game = P1Game + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + "(" + Str(P1Score) + "-" + Str(P2Score) + ")" + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
TB = 0
TBCount = 0
End If
Else
If P1Score >= 4 And P1Score - P2Score >= 2 Then
P1Game = P1Game + 1
P1Score = 0
P2Score = 0
P2Fault = 0
GameSet = 1
P1Fault = 0
End If
If P1Game - P2Game > 1 And P1Game = 6 Or P1Game - P2Game >= 1 And P1Game = 7 Then
P1Set = P1Set + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
End If
End If

If P1Game = 6 And P2Game = 6 Then
TB = 1
End If


If GameSet = 1 And (P1Set = 3 Or P2Set = 3) Then
GameSet = 0
Match = 1
lblScore.Caption = Score
P1PW.Caption = Str(P1PWon)
P1A.Caption = Str(P1Aces)
P1DF.Caption = Str(P1DFault)
P2PW.Caption = Str(P2PWon)
P2A.Caption = Str(P2Aces)
P2DF.Caption = Str(P2DFault)
Exit Sub
End If

If TB = 1 Then
If TBCount = 1 Then
Serve = 1
GoTo P1Server
ElseIf TBCount = 2 Then
TBCount = 0
Serve = 2
GoTo P2Server
End If
End If


If GameSet = 1 And Match = 0 And TB = 0 Then
GameSet = 0
If Serve = 1 Then
Serve = 2
GoTo P2Server
Else
Serve = 1
GoTo P1Server
End If
ElseIf GameSet = 0 And Match = 0 And TB = 0 Then
If Serve = 1 Then
GoTo P1Server
Else
GoTo P2Server
End If
End If

P1Move:
If P1Strategy = 1 Then
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(6 * Rnd) + 1
If Rno = 1 Then
P1Shot = P1Volley
ElseIf Rno = 2 Then
P1Shot = P1HVolley
ElseIf Rno = 3 Then
P1Shot = P1OHead
ElseIf Rno = 4 Then
P1Shot = P1Lob
ElseIf Rno = 5 Then
P1Shot = P1BHand
ElseIf Rno = 6 Then
P1Shot = P1Speed
End If
ElseIf Rno = 2 Then
P1Shot = P1FHand
Else
P1Shot = P1DShot
End If
End If

If P1Strategy = 2 Then
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(6 * Rnd) + 1
If Rno = 1 Then
P1Shot = P1Volley
ElseIf Rno = 2 Then
P1Shot = P1HVolley
ElseIf Rno = 3 Then
P1Shot = P1OHead
ElseIf Rno = 4 Then
P1Shot = P1Lob
ElseIf Rno = 5 Then
P1Shot = P1DShot
ElseIf Rno = 6 Then
P1Shot = P1FHand
End If
ElseIf Rno = 2 Then
P1Shot = P1BHand
Else
P1Shot = P1Speed
End If
End If

If P1Strategy = 3 Then
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(6 * Rnd) + 1
If Rno = 1 Then
P1Shot = P1OHead
ElseIf Rno = 2 Then
P1Shot = P1Lob
ElseIf Rno = 3 Then
P1Shot = P1DShot
ElseIf Rno = 4 Then
P1Shot = P1FHand
ElseIf Rno = 5 Then
P1Shot = P1BHand
ElseIf Rno = 6 Then
P1Shot = P1Speed
End If
ElseIf Rno = 2 Then
P1Shot = P1Volley
Else
P1Shot = P1HVolley
End If
End If

If P1Strategy = 4 Then
Rno = Int(8 * Rnd) + 1
If Rno = 1 Then
P1Shot = P1Volley
ElseIf Rno = 2 Then
P1Shot = P1HVolley
ElseIf Rno = 3 Then
P1Shot = P1OHead
ElseIf Rno = 4 Then
P1Shot = P1Lob
ElseIf Rno = 5 Then
P1Shot = P1DShot
ElseIf Rno = 6 Then
P1Shot = P1FHand
ElseIf Rno = 7 Then
P1Shot = P1BHand
ElseIf Rno = 8 Then
P1Shot = P1Speed
End If
End If

If P1Shot >= 0 And P1Shot <= 4 Then
Rno = Int(10 * Rnd) + 1
Else
Rno = Int(P1Shot * Rnd) + 1
End If

If (P1Shot >= 0 And P1Shot <= 4) And (Rno >= 1 And Rno <= 9) Then
 P1Fail = 1

ElseIf (P1Shot >= 5 And P1Shot <= 10) And (Rno >= 1 And Rno <= 4) Then
 P1Fail = 1

ElseIf (P1Shot >= 11 And P1Shot <= 13) And (Rno >= 1 And Rno <= 5) Then
 P1Fail = 1

ElseIf (P1Shot = 14 Or P1Shot = 15) And (Rno >= 1 And Rno <= 6) Then
 P1Fail = 1

ElseIf (P1Shot >= 16 And P1Shot <= 18) And (Rno >= 1 And Rno <= 7) Then
 P1Fail = 1

ElseIf (P1Shot = 19 Or P1Shot = 20) And (Rno >= 1 And Rno <= 8) Then
 P1Fail = 1

ElseIf (P1Shot >= 21 And P1Shot <= 23) And (Rno >= 1 And Rno <= 9) Then
 P1Fail = 1

ElseIf (P1Shot = 24 Or P1Shot = 25) And (Rno >= 1 And Rno <= 10) Then
 P1Fail = 1

ElseIf (P1Shot >= 26 And P1Shot <= 28) And (Rno >= 1 And Rno <= 11) Then
 P1Fail = 1

ElseIf (P1Shot = 29 Or P1Shot = 30) And (Rno >= 1 And Rno <= 12) Then
 P1Fail = 1

ElseIf (P1Shot >= 31 And P1Shot <= 33) And (Rno >= 1 And Rno <= 13) Then
 P1Fail = 1

ElseIf (P1Shot = 34 Or P1Shot = 35) And (Rno >= 1 And Rno <= 14) Then
 P1Fail = 1

ElseIf (P1Shot >= 36 And P1Shot <= 38) And (Rno >= 1 And Rno <= 15) Then
 P1Fail = 1

ElseIf (P1Shot = 39 Or P1Shot = 40) And (Rno >= 1 And Rno <= 16) Then
 P1Fail = 1

ElseIf (P1Shot >= 41 And P1Shot <= 43) And (Rno >= 1 And Rno <= 17) Then
 P1Fail = 1

ElseIf (P1Shot = 44 Or P1Shot = 45) And (Rno >= 1 And Rno <= 18) Then
 P1Fail = 1

ElseIf (P1Shot >= 46 And P1Shot <= 48) And (Rno >= 1 And Rno <= 19) Then
 P1Fail = 1

ElseIf (P1Shot = 49 Or P1Shot = 50) And (Rno >= 1 And Rno <= 20) Then
 P1Fail = 1


ElseIf (P1Shot >= 51 And P1Shot <= 55) And (Rno >= 1 And Rno <= 19) Then
P1Fail = 1

ElseIf (P1Shot >= 56 And P1Shot <= 60) And (Rno >= 1 And Rno <= 18) Then
P1Fail = 1

ElseIf (P1Shot >= 61 And P1Shot <= 65) And (Rno >= 1 And Rno <= 17) Then
P1Fail = 1

ElseIf (P1Shot >= 66 And P1Shot <= 70) And (Rno >= 1 And Rno <= 16) Then
P1Fail = 1

ElseIf (P1Shot >= 71 And P1Shot <= 75) And (Rno >= 1 And Rno <= 15) Then
P1Fail = 1

ElseIf (P1Shot >= 76 And P1Shot <= 80) And (Rno >= 1 And Rno <= 14) Then
P1Fail = 1

ElseIf (P1Shot >= 81 And P1Shot <= 85) And (Rno >= 1 And Rno <= 13) Then
P1Fail = 1

ElseIf (P1Shot >= 86 And P1Shot <= 90) And (Rno >= 1 And Rno <= 12) Then
P1Fail = 1

ElseIf (P1Shot >= 91 And P1Shot <= 95) And (Rno >= 1 And Rno <= 11) Then
P1Fail = 1

ElseIf (P1Shot >= 96 And P1Shot <= 100) And (Rno >= 1 And Rno <= 10) Then
P1Fail = 1

Else
GoTo P2Move
End If

If P1Fail = 1 Then
P2PWon = P2PWon + 1
P1Fail = 0
P2Score = P2Score + 1
If TB = 1 Then
TBCount = TBCount + 1
If P2Score >= 7 And P2Score - P1Score >= 2 Then
P2Set = P2Set + 1
P2Game = P2Game + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + "(" + Str(P1Score) + "-" + Str(P2Score) + ")" + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
TB = 0
TBCount = 0
End If
Else
If P2Score >= 4 And P2Score - P1Score >= 2 Then
P2Game = P2Game + 1
P1Score = 0
P2Score = 0
P2Fault = 0
P1Fault = 0
GameSet = 1
End If
If P2Game - P1Game > 1 And P2Game = 6 Or P2Game - P1Game >= 1 And P2Game = 7 Then
P2Set = P2Set + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
End If
End If
End If

If P1Game = 6 And P2Game = 6 Then
TB = 1
End If


If GameSet = 1 And (P1Set = 3 Or P2Set = 3) Then
GameSet = 0
Match = 1
lblScore.Caption = Score
P1PW.Caption = Str(P1PWon)
P1A.Caption = Str(P1Aces)
P1DF.Caption = Str(P1DFault)
P2PW.Caption = Str(P2PWon)
P2A.Caption = Str(P2Aces)
P2DF.Caption = Str(P2DFault)
Exit Sub
End If

If TB = 1 Then
If TBCount = 0 And Serve = 1 Then
Serve = 2
GoTo P2Server
ElseIf TBCount = 0 And Serve = 2 Then
Serve = 1
GoTo P1Server
ElseIf TBCount = 1 And Serve = 1 Then
Serve = 1
GoTo P1Server
ElseIf TBCount = 1 And Serve = 2 Then
Serve = 2
GoTo P2Server
ElseIf TBCount = 2 And Serve = 1 Then
TBCount = 0
Serve = 2
GoTo P2Server
ElseIf TBCount = 2 And Serve = 2 Then
TBCount = 0
Serve = 1
GoTo P1Server
End If
End If



If GameSet = 1 And Match = 0 And TB = 0 Then
GameSet = 0
If Serve = 1 Then
Serve = 2
GoTo P2Server
Else
Serve = 1
GoTo P1Server
End If
ElseIf GameSet = 0 And Match = 0 And TB = 0 Then
If Serve = 1 Then
GoTo P1Server
Else
GoTo P2Server
End If
End If

P2Server:
If P2Serve >= 0 And P2Serve <= 4 Then
Rno = Int(10 * Rnd) + 1
Else
Rno = Int(P2Serve * Rnd) + 1
End If

If (P2Serve >= 0 And P2Serve <= 4) And (Rno >= 1 And Rno <= 9) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 5 And P2Serve <= 10) And (Rno >= 1 And Rno <= 4) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 11 And P2Serve <= 13) And (Rno >= 1 And Rno <= 5) Then
GoTo P2ServeFail

ElseIf (P2Serve = 14 Or P2Serve = 15) And (Rno >= 1 And Rno <= 6) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 16 And P2Serve <= 18) And (Rno >= 1 And Rno <= 7) Then
GoTo P2ServeFail

ElseIf (P2Serve = 19 Or P2Serve = 20) And (Rno >= 1 And Rno <= 8) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 21 And P2Serve <= 23) And (Rno >= 1 And Rno <= 9) Then
GoTo P2ServeFail

ElseIf (P2Serve = 24 Or P2Serve = 25) And (Rno >= 1 And Rno <= 10) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 26 And P2Serve <= 28) And (Rno >= 1 And Rno <= 11) Then
GoTo P2ServeFail

ElseIf (P2Serve = 29 Or P2Serve = 30) And (Rno >= 1 And Rno <= 12) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 31 And P2Serve <= 33 And P2Fault = 0) And (Rno >= 1 And Rno <= 13) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 31 And P2Serve <= 33 And P2Fault = 1) And (Rno >= 1 And Rno <= 7) Then
GoTo P2ServeFail

ElseIf (P2Serve = 34 Or P2Serve = 35 And P2Fault = 0) And (Rno >= 1 And Rno <= 14) Then
GoTo P2ServeFail

ElseIf (P2Serve = 34 Or P2Serve = 35 And P2Fault = 1) And (Rno >= 1 And Rno <= 7) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 36 And P2Serve <= 38 And P2Fault = 0) And (Rno >= 1 And Rno <= 15) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 36 And P2Serve <= 38 And P2Fault = 1) And (Rno >= 1 And Rno <= 8) Then
GoTo P2ServeFail

ElseIf (P2Serve = 39 Or P2Serve = 40 And P2Fault = 0) And (Rno >= 1 And Rno <= 16) Then
GoTo P2ServeFail

ElseIf (P2Serve = 39 Or P2Serve = 40 And P2Fault = 1) And (Rno >= 1 And Rno <= 8) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 41 And P2Serve <= 43 And P2Fault = 0) And (Rno >= 1 And Rno <= 17) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 41 And P2Serve <= 43 And P2Fault = 1) And (Rno >= 1 And Rno <= 9) Then
GoTo P2ServeFail

ElseIf (P2Serve = 44 Or P2Serve = 45 And P2Fault = 0) And (Rno >= 1 And Rno <= 18) Then
GoTo P2ServeFail

ElseIf (P2Serve = 44 Or P2Serve = 45 And P2Fault = 1) And (Rno >= 1 And Rno <= 9) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 46 And P2Serve <= 48 And P2Fault = 0) And (Rno >= 1 And Rno <= 19) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 46 And P2Serve <= 48 And P2Fault = 1) And (Rno >= 1 And Rno <= 10) Then
GoTo P2ServeFail

ElseIf (P2Serve = 49 Or P2Serve = 50 And P2Fault = 0) And (Rno >= 1 And Rno <= 20) Then
GoTo P2ServeFail

ElseIf (P2Serve = 49 Or P2Serve = 50 And P2Fault = 1) And (Rno >= 1 And Rno <= 10) Then
GoTo P2ServeFail


ElseIf (P2Serve >= 51 And P2Serve <= 55 And P2Fault = 0) And (Rno >= 1 And Rno <= 19) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 51 And P2Serve <= 55 And P2Fault = 1) And (Rno >= 1 And Rno <= 10) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 56 And P2Serve <= 60 And P2Fault = 0) And (Rno >= 1 And Rno <= 18) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 56 And P2Serve <= 60 And P2Fault = 1) And (Rno >= 1 And Rno <= 9) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 61 And P2Serve <= 65 And P2Fault = 0) And (Rno >= 1 And Rno <= 17) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 61 And P2Serve <= 65 And P2Fault = 1) And (Rno >= 1 And Rno <= 9) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 66 And P2Serve <= 70 And P2Fault = 0) And (Rno >= 1 And Rno <= 16) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 66 And P2Serve <= 70 And P2Fault = 1) And (Rno >= 1 And Rno <= 8) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 71 And P2Serve <= 75 And P2Fault = 0) And (Rno >= 1 And Rno <= 15) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 71 And P2Serve <= 75 And P2Fault = 1) And (Rno >= 1 And Rno <= 8) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 76 And P2Serve <= 80 And P2Fault = 0) And (Rno >= 1 And Rno <= 14) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 76 And P2Serve <= 80 And P2Fault = 1) And (Rno >= 1 And Rno <= 7) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 81 And P2Serve <= 85 And P2Fault = 0) And (Rno >= 1 And Rno <= 13) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 81 And P2Serve <= 85 And P2Fault = 1) And (Rno >= 1 And Rno <= 7) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 86 And P2Serve <= 90 And P2Fault = 0) And (Rno >= 1 And Rno <= 12) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 86 And P2Serve <= 90 And P2Fault = 1) And (Rno >= 1 And Rno <= 6) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 91 And P2Serve <= 95 And P2Fault = 0) And (Rno >= 1 And Rno <= 11) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 91 And P2Serve <= 95 And P2Fault = 1) And (Rno >= 1 And Rno <= 6) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 96 And P2Serve <= 100 And P2Fault = 0) And (Rno >= 1 And Rno <= 10) Then
GoTo P2ServeFail

ElseIf (P2Serve >= 96 And P2Serve <= 100 And P2Fault = 1) And (Rno >= 1 And Rno <= 5) Then
GoTo P2ServeFail


Else

If (P2Serve >= 31 And P2Serve <= 33 And P2Fault = 0) And (Rno = 14) Then
GoTo P2Ace

ElseIf (P2Serve = 34 Or P2Serve = 35) And (P2Fault = 0 And Rno = 15) Then
GoTo P2Ace

ElseIf (P2Serve >= 36 And P2Serve <= 38 And P2Fault = 0) And (Rno = 16) Then
GoTo P2Ace

ElseIf (P2Serve = 39 Or P2Serve = 40) And (P2Fault = 0 And Rno = 17) Then
GoTo P2Ace

ElseIf (P2Serve >= 41 And P2Serve <= 43 And P2Fault = 0) And (Rno = 18) Then
GoTo P2Ace

ElseIf (P2Serve = 44 Or P2Serve = 45) And (P2Fault = 0 And Rno = 19) Then
GoTo P2Ace

ElseIf (P2Serve >= 46 And P2Serve <= 48 And P2Fault = 0) And (Rno = 20) Then
GoTo P2Ace

ElseIf (P2Serve = 49 Or P2Serve = 50) And (P2Fault = 0 And Rno = 21) Then
GoTo P2Ace

ElseIf (P2Serve >= 51 And P2Serve <= 55 And P2Fault = 0) And (Rno = 20 Or Rno = 21 Or Rno = 22) Then
GoTo P2Ace

ElseIf (P2Serve >= 51 And P2Serve <= 55 And P2Fault = 1) And (Rno = 11) Then
GoTo P2Ace

ElseIf (P2Serve >= 56 And P2Serve <= 60 And P2Fault = 0) And (Rno = 19 Or Rno = 20 Or Rno = 21 Or Rno = 22) Then
GoTo P2Ace

ElseIf (P2Serve >= 56 And P2Serve <= 60 And P2Fault = 1) And (Rno = 10 Or Rno = 11) Then
GoTo P2Ace

ElseIf (P2Serve >= 61 And P2Serve <= 65 And P2Fault = 0) And (Rno >= 18 And Rno <= 23) Then
GoTo P2Ace

ElseIf (P2Serve >= 61 And P2Serve <= 65 And P2Fault = 1) And (Rno >= 10 And Rno <= 12) Then
GoTo P2Ace

ElseIf (P2Serve >= 66 And P2Serve <= 70 And P2Fault = 0) And (Rno >= 17 And Rno <= 24) Then
GoTo P2Ace

ElseIf (P2Serve >= 66 And P2Serve <= 70 And P2Fault = 1) And (Rno >= 9 And Rno <= 12) Then
GoTo P2Ace

ElseIf (P2Serve >= 71 And P2Serve <= 75 And P2Fault = 0) And (Rno >= 16 And Rno <= 25) Then
GoTo P2Ace

ElseIf (P2Serve >= 71 And P2Serve <= 75 And P2Fault = 1) And (Rno >= 9 And Rno <= 13) Then
GoTo P2Ace

ElseIf (P2Serve >= 76 And P2Serve <= 80 And P2Fault = 0) And (Rno >= 15 And Rno <= 26) Then
GoTo P2Ace

ElseIf (P2Serve >= 76 And P2Serve <= 80 And P2Fault = 1) And (Rno >= 8 And Rno <= 13) Then
GoTo P2Ace

ElseIf (P2Serve >= 81 And P2Serve <= 85 And P2Fault = 0) And (Rno >= 14 And Rno <= 27) Then
GoTo P2Ace

ElseIf (P2Serve >= 81 And P2Serve <= 85 And P2Fault = 1) And (Rno >= 8 And Rno <= 14) Then
GoTo P2Ace

ElseIf (P2Serve >= 86 And P2Serve <= 90 And P2Fault = 0) And (Rno >= 13 And Rno <= 28) Then
GoTo P2Ace

ElseIf (P2Serve >= 86 And P2Serve <= 90 And P2Fault = 1) And (Rno >= 7 And Rno <= 14) Then
GoTo P2Ace

ElseIf (P2Serve >= 91 And P2Serve <= 95 And P2Fault = 0) And (Rno >= 12 And Rno <= 29) Then
GoTo P2Ace

ElseIf (P2Serve >= 91 And P2Serve <= 95 And P2Fault = 1) And (Rno >= 7 And Rno <= 15) Then
GoTo P2Ace

ElseIf (P2Serve >= 91 And P2Serve <= 95 And P2Fault = 0) And (Rno >= 12 And Rno <= 29) Then
GoTo P2Ace

ElseIf (P2Serve >= 91 And P2Serve <= 95 And P2Fault = 1) And (Rno >= 6 And Rno <= 15) Then
GoTo P2Ace

Else
P2Fault = 0
GoTo P1Move
End If
End If

P2ServeFail:
P2Fault = P2Fault + 1
If P2Fault = 2 Then
P2DFault = P2DFault + 1
P1PWon = P1PWon + 1
P2Fault = 0
P1Score = P1Score + 1
If TB = 1 Then
TBCount = TBCount + 1
If P1Score >= 7 And P1Score - P2Score >= 2 Then
P1Set = P1Set + 1
P1Game = P1Game + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + "(" + Str(P1Score) + "-" + Str(P2Score) + ")" + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
TB = 0
TBCount = 0
End If
Else
If P1Score >= 4 And P1Score - P2Score >= 2 Then
P1Game = P1Game + 1
P1Score = 0
P2Score = 0
P1Fault = 0
P2Fault = 0
GameSet = 1
End If
If P1Game - P2Game > 1 And P1Game = 6 Or P1Game - P2Game >= 1 And P1Game = 7 Then
P1Set = P1Set + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
End If
End If

Else
GoTo P2Server
End If

If P1Game = 6 And P2Game = 6 Then
TB = 1
End If

If GameSet = 1 And (P1Set = 3 Or P2Set = 3) Then
GameSet = 0
Match = 1
lblScore.Caption = Score
P1PW.Caption = Str(P1PWon)
P1A.Caption = Str(P1Aces)
P1DF.Caption = Str(P1DFault)
P2PW.Caption = Str(P2PWon)
P2A.Caption = Str(P2Aces)
P2DF.Caption = Str(P2DFault)
Exit Sub
End If

If TB = 1 Then
If TBCount = 0 Then
Serve = 1
GoTo P1Server
ElseIf TBCount = 1 Then
Serve = 2
GoTo P2Server
ElseIf TBCount = 2 Then
TBCount = 0
Serve = 1
GoTo P1Server
End If
End If

If GameSet = 1 And Match = 0 And TB = 0 Then
GameSet = 0
If Serve = 1 Then
Serve = 2
GoTo P2Server
Else
Serve = 1
GoTo P1Server
End If
ElseIf GameSet = 0 And Match = 0 And TB = 0 Then
If Serve = 1 Then
GoTo P1Server
Else
GoTo P2Server
End If
End If

P2Ace:
P2Fault = 0
P2Aces = P2Aces + 1
P2PWon = P2PWon + 1
P2Score = P2Score + 1
If TB = 1 Then
TBCount = TBCount + 1
If P2Score >= 7 And P2Score - P1Score >= 2 Then
P2Set = P2Set + 1
P2Game = P2Game + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + "(" + Str(P1Score) + "-" + Str(P2Score) + ")" + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
TB = 0
TBCount = 0
End If
Else
If P2Score >= 4 And P2Score - P1Score >= 2 Then
P2Game = P2Game + 1
P1Score = 0
P2Score = 0
P2Fault = 0
GameSet = 1
P1Fault = 0
End If
If P2Game - P1Game > 1 And P2Game = 6 Or P2Game - P1Game >= 1 And P2Game = 7 Then
P2Set = P2Set + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
End If
End If

If P1Game = 6 And P2Game = 6 Then
TB = 1
End If


If GameSet = 1 And (P1Set = 3 Or P2Set = 3) Then
GameSet = 0
Match = 1
lblScore.Caption = Score
P1PW.Caption = Str(P1PWon)
P1A.Caption = Str(P1Aces)
P1DF.Caption = Str(P1DFault)
P2PW.Caption = Str(P2PWon)
P2A.Caption = Str(P2Aces)
P2DF.Caption = Str(P2DFault)
Exit Sub
End If

If TB = 1 Then
If TBCount = 1 Then
Serve = 2
GoTo P2Server
ElseIf TBCount = 2 Then
TBCount = 0
Serve = 1
GoTo P1Server
End If
End If


If GameSet = 1 And Match = 0 And TB = 0 Then
GameSet = 0
If Serve = 1 Then
Serve = 2
GoTo P2Server
Else
Serve = 1
GoTo P1Server
End If
ElseIf GameSet = 0 And Match = 0 And TB = 0 Then
If Serve = 1 Then
GoTo P1Server
Else
GoTo P2Server
End If
End If

P2Move:
If P2Strategy = 1 Then
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(6 * Rnd) + 1
If Rno = 1 Then
P2Shot = P2Volley
ElseIf Rno = 2 Then
P2Shot = P2HVolley
ElseIf Rno = 3 Then
P2Shot = P2OHead
ElseIf Rno = 4 Then
P2Shot = P2Lob
ElseIf Rno = 5 Then
P2Shot = P2BHand
ElseIf Rno = 6 Then
P2Shot = P2Speed
End If
ElseIf Rno = 2 Then
P2Shot = P2FHand
Else
P2Shot = P2DShot
End If
End If

If P2Strategy = 2 Then
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(6 * Rnd) + 1
If Rno = 1 Then
P2Shot = P2Volley
ElseIf Rno = 2 Then
P2Shot = P2HVolley
ElseIf Rno = 3 Then
P2Shot = P2OHead
ElseIf Rno = 4 Then
P2Shot = P2Lob
ElseIf Rno = 5 Then
P2Shot = P2DShot
ElseIf Rno = 6 Then
P2Shot = P2FHand
End If
ElseIf Rno = 2 Then
P2Shot = P2BHand
Else
P2Shot = P2Speed
End If
End If

If P2Strategy = 3 Then
Rno = Int(3 * Rnd) + 1
If Rno = 1 Then
Rno = Int(6 * Rnd) + 1
If Rno = 1 Then
P2Shot = P2OHead
ElseIf Rno = 2 Then
P2Shot = P2Lob
ElseIf Rno = 3 Then
P2Shot = P2DShot
ElseIf Rno = 4 Then
P2Shot = P2FHand
ElseIf Rno = 5 Then
P2Shot = P2BHand
ElseIf Rno = 6 Then
P2Shot = P2Speed
End If
ElseIf Rno = 2 Then
P2Shot = P2Volley
Else
P2Shot = P2HVolley
End If
End If

If P2Strategy = 4 Then
Rno = Int(8 * Rnd) + 1
If Rno = 1 Then
P2Shot = P2Volley
ElseIf Rno = 2 Then
P2Shot = P2HVolley
ElseIf Rno = 3 Then
P2Shot = P2OHead
ElseIf Rno = 4 Then
P2Shot = P2Lob
ElseIf Rno = 5 Then
P2Shot = P2DShot
ElseIf Rno = 6 Then
P2Shot = P2FHand
ElseIf Rno = 7 Then
P2Shot = P2BHand
ElseIf Rno = 8 Then
P2Shot = P2Speed
End If
End If

If P2Shot >= 0 And P2Shot <= 4 Then
Rno = Int(10 * Rnd) + 1
Else
Rno = Int(P2Shot * Rnd) + 1
End If

If (P2Shot >= 0 And P2Shot <= 4) And (Rno >= 1 And Rno <= 9) Then
 P2Fail = 1

ElseIf (P2Shot >= 5 And P2Shot <= 10) And (Rno >= 1 And Rno <= 4) Then
 P2Fail = 1

ElseIf (P2Shot >= 11 And P2Shot <= 13) And (Rno >= 1 And Rno <= 5) Then
 P2Fail = 1

ElseIf (P2Shot = 14 Or P2Shot = 15) And (Rno >= 1 And Rno <= 6) Then
 P2Fail = 1

ElseIf (P2Shot >= 16 And P2Shot <= 18) And (Rno >= 1 And Rno <= 7) Then
 P2Fail = 1

ElseIf (P2Shot = 19 Or P2Shot = 20) And (Rno >= 1 And Rno <= 8) Then
 P2Fail = 1

ElseIf (P2Shot >= 21 And P2Shot <= 23) And (Rno >= 1 And Rno <= 9) Then
 P2Fail = 1

ElseIf (P2Shot = 24 Or P2Shot = 25) And (Rno >= 1 And Rno <= 10) Then
 P2Fail = 1

ElseIf (P2Shot >= 26 And P2Shot <= 28) And (Rno >= 1 And Rno <= 11) Then
 P2Fail = 1

ElseIf (P2Shot = 29 Or P2Shot = 30) And (Rno >= 1 And Rno <= 12) Then
 P2Fail = 1

ElseIf (P2Shot >= 31 And P2Shot <= 33) And (Rno >= 1 And Rno <= 13) Then
 P2Fail = 1

ElseIf (P2Shot = 34 Or P2Shot = 35) And (Rno >= 1 And Rno <= 14) Then
 P2Fail = 1

ElseIf (P2Shot >= 36 And P2Shot <= 38) And (Rno >= 1 And Rno <= 15) Then
 P2Fail = 1

ElseIf (P2Shot = 39 Or P2Shot = 40) And (Rno >= 1 And Rno <= 16) Then
 P2Fail = 1

ElseIf (P2Shot >= 41 And P2Shot <= 43) And (Rno >= 1 And Rno <= 17) Then
 P2Fail = 1

ElseIf (P2Shot = 44 Or P2Shot = 45) And (Rno >= 1 And Rno <= 18) Then
 P2Fail = 1

ElseIf (P2Shot >= 46 And P2Shot <= 48) And (Rno >= 1 And Rno <= 19) Then
 P2Fail = 1

ElseIf (P2Shot = 49 Or P2Shot = 50) And (Rno >= 1 And Rno <= 20) Then
 P2Fail = 1


ElseIf (P2Shot >= 51 And P2Shot <= 55) And (Rno >= 1 And Rno <= 19) Then
P2Fail = 1

ElseIf (P2Shot >= 56 And P2Shot <= 60) And (Rno >= 1 And Rno <= 18) Then
P2Fail = 1

ElseIf (P2Shot >= 61 And P2Shot <= 65) And (Rno >= 1 And Rno <= 17) Then
P2Fail = 1

ElseIf (P2Shot >= 66 And P2Shot <= 70) And (Rno >= 1 And Rno <= 16) Then
P2Fail = 1

ElseIf (P2Shot >= 71 And P2Shot <= 75) And (Rno >= 1 And Rno <= 15) Then
P2Fail = 1

ElseIf (P2Shot >= 76 And P2Shot <= 80) And (Rno >= 1 And Rno <= 14) Then
P2Fail = 1

ElseIf (P2Shot >= 81 And P2Shot <= 85) And (Rno >= 1 And Rno <= 13) Then
P2Fail = 1

ElseIf (P2Shot >= 86 And P2Shot <= 90) And (Rno >= 1 And Rno <= 12) Then
P2Fail = 1

ElseIf (P2Shot >= 91 And P2Shot <= 95) And (Rno >= 1 And Rno <= 11) Then
P2Fail = 1

ElseIf (P2Shot >= 96 And P2Shot <= 100) And (Rno >= 1 And Rno <= 10) Then
P2Fail = 1

Else
GoTo P1Move
End If

If P2Fail = 1 Then
P1PWon = P1PWon + 1
P2Fail = 0
P1Score = P1Score + 1
If TB = 1 Then
TBCount = TBCount + 1
If P1Score >= 7 And P1Score - P2Score >= 2 Then
P1Set = P1Set + 1
P1Game = P1Game + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + "(" + Str(P1Score) + "-" + Str(P2Score) + ")" + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
TB = 0
TBCount = 0
End If
Else
If P1Score >= 4 And P1Score - P2Score >= 2 Then
P1Game = P1Game + 1
P1Score = 0
P2Score = 0
P2Fault = 0
P1Fault = 0
GameSet = 1
End If
If P1Game - P2Game > 1 And P1Game = 6 Or P1Game - P2Game >= 1 And P1Game = 7 Then
P1Set = P1Set + 1
Score = Score + Str(P1Game) + "-" + Str(P2Game) + " "
P1Fault = 0
P2Fault = 0
P2Score = 0
P1Score = 0
P2Game = 0
P1Game = 0
GameSet = 1
End If
End If
End If

If P1Game = 6 And P2Game = 6 Then
TB = 1
End If


If GameSet = 1 And (P1Set = 3 Or P2Set = 3) Then
GameSet = 0
Match = 1
lblScore.Caption = Score
P1PW.Caption = Str(P1PWon)
P1A.Caption = Str(P1Aces)
P1DF.Caption = Str(P1DFault)
P2PW.Caption = Str(P2PWon)
P2A.Caption = Str(P2Aces)
P2DF.Caption = Str(P2DFault)
Exit Sub
End If

If TB = 1 Then
If TBCount = 0 And Serve = 1 Then
Serve = 2
GoTo P2Server
ElseIf TBCount = 0 And Serve = 2 Then
Serve = 1
GoTo P1Server
ElseIf TBCount = 1 And Serve = 1 Then
Serve = 1
GoTo P1Server
ElseIf TBCount = 1 And Serve = 2 Then
Serve = 2
GoTo P2Server
ElseIf TBCount = 2 And Serve = 1 Then
TBCount = 0
Serve = 2
GoTo P2Server
ElseIf TBCount = 2 And Serve = 2 Then
TBCount = 0
Serve = 1
GoTo P1Server
End If
End If



If GameSet = 1 And Match = 0 And TB = 0 Then
GameSet = 0
If Serve = 1 Then
Serve = 2
GoTo P2Server
Else
Serve = 1
GoTo P1Server
End If
ElseIf GameSet = 0 And Match = 0 And TB = 0 Then
If Serve = 1 Then
GoTo P1Server
Else
GoTo P2Server
End If
End If
End Sub
Private Sub Form_Load()
Randomize Timer
End Sub

