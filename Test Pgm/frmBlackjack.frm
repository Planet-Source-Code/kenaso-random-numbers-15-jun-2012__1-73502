VERSION 5.00
Begin VB.Form frmBlackjack 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4740
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   4605
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   315
      Left            =   3105
      TabIndex        =   43
      Top             =   4230
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3885
      TabIndex        =   42
      Top             =   4230
      Width           =   615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top"
      Height          =   315
      Left            =   3105
      TabIndex        =   41
      Top             =   4230
      Width           =   615
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   17
      Left            =   2820
      TabIndex        =   40
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   900
      Width           =   690
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   16
      Left            =   2820
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   675
      Width           =   690
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   15
      Left            =   1425
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   900
      Width           =   1140
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   14
      Left            =   1425
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   675
      Width           =   1140
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   0
      Left            =   1425
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   450
      Width           =   1140
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   13
      Left            =   3075
      TabIndex        =   35
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   12
      Left            =   3075
      TabIndex        =   34
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3060
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   11
      Left            =   3075
      TabIndex        =   33
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   10
      Left            =   3075
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2505
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   9
      Left            =   3075
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2205
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   8
      Left            =   3075
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1905
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   7
      Left            =   675
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3615
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   6
      Left            =   675
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   5
      Left            =   675
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3060
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   4
      Left            =   675
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   675
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2505
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   2
      Left            =   675
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2205
      Width           =   1215
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   675
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1905
      Width           =   1215
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Times played"
      Height          =   390
      Index           =   6
      Left            =   3675
      TabIndex        =   22
      Top             =   1275
      Width           =   615
   End
   Begin VB.Label lblMisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Card values"
      Height          =   390
      Index           =   5
      Left            =   2325
      TabIndex        =   21
      Top             =   1275
      Width           =   765
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Times played"
      Height          =   390
      Index           =   4
      Left            =   1350
      TabIndex        =   20
      Top             =   1275
      Width           =   540
   End
   Begin VB.Label lblMisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Card values"
      Height          =   390
      Index           =   3
      Left            =   0
      TabIndex        =   19
      Top             =   1275
      Width           =   765
   End
   Begin VB.Line Line2 
      X1              =   150
      X2              =   4425
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblPause 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   75
      TabIndex        =   18
      Top             =   4095
      Width           =   2460
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "of  50,000"
      Height          =   195
      Index           =   23
      Left            =   2790
      TabIndex        =   17
      Top             =   450
      Width           =   720
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   22
      Left            =   150
      TabIndex        =   16
      Top             =   450
      Width           =   1140
   End
   Begin VB.Line Line1 
      X1              =   2100
      X2              =   2100
      Y1              =   1200
      Y2              =   3960
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "King"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   2385
      TabIndex        =   15
      Top             =   3315
      Width           =   570
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Queen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   2355
      TabIndex        =   14
      Top             =   3015
      Width           =   570
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jack"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   2385
      TabIndex        =   13
      Top             =   2715
      Width           =   570
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   2385
      TabIndex        =   12
      Top             =   2460
      Width           =   570
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   2385
      TabIndex        =   11
      Top             =   2160
      Width           =   570
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   2385
      TabIndex        =   10
      Top             =   1860
      Width           =   570
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   180
      TabIndex        =   9
      Top             =   3570
      Width           =   345
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   180
      TabIndex        =   8
      Top             =   3315
      Width           =   345
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   180
      TabIndex        =   7
      Top             =   3015
      Width           =   345
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   2715
      Width           =   345
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   180
      TabIndex        =   5
      Top             =   2460
      Width           =   345
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   2160
      Width           =   345
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1860
      Width           =   345
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   375
      TabIndex        =   2
      Top             =   900
      Width           =   915
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wins"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   375
      TabIndex        =   1
      Top             =   675
      Width           =   915
   End
   Begin VB.Label lblMisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blackjack Demo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   -75
      Width           =   4290
   End
End
Attribute VB_Name = "frmBlackjack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mhFile          As Long
Private mlngAlgorithm   As Long
Private mblnStopPressed As Boolean
Private mblnExitPgm     As Boolean
Private mobjPrng        As Object

Private Const MAX_GAMES As Long = 50000    ' number of games
Private Const MAX_SIZE  As Long = 100000   ' size of random number array

Public Sub Algorithm(ByVal lngData As Long)
    
    mlngAlgorithm = lngData
    
    With frmBlackjack
        Select Case mlngAlgorithm
               Case 0: .Caption = "Blackjack - CryptoAPI Prng"
               Case 1: .Caption = "Blackjack - Isaac Prng"
               Case 2: .Caption = "Blackjack - KISS Prng"
               Case 3: .Caption = "Blackjack - MWC-SHR3 Prng"
               Case 4: .Caption = "Blackjack - Mother-of-All Prng"
               Case 5: .Caption = "Blackjack - MT19937 Prng"
               Case 6: .Caption = "Blackjack - MT11231-A Prng"
               Case 7: .Caption = "Blackjack - MT11231-B Prng"
               Case 8: .Caption = "Blackjack - TT800 Prng"
        End Select
    End With
    
End Sub

Private Sub cmdExit_Click()
    
    DoEvents
    mblnStopPressed = True
    mblnExitPgm = True
    mobjPrng.StopProcessing = True
    DoEvents
    
    Close #mhFile
    
    DoEvents
    Unload Me

End Sub

Private Sub cmdStart_Click()

    cmdStop.Visible = True
    cmdStart.Visible = False
    mblnStopPressed = False
    mobjPrng.StopProcessing = False
    PlayBlackjack
    
End Sub

Private Sub cmdStop_Click()

    DoEvents
    cmdStart.Visible = True
    cmdStop.Visible = False
    
    DoEvents
    mblnStopPressed = True
    mobjPrng.StopProcessing = True
    DoEvents
    
End Sub

Private Sub Form_Load()
  
    Select Case mlngAlgorithm
           Case 0: Set mobjPrng = New kiPrng.cPrng
           Case 1: Set mobjPrng = New kiPrng.cIsaac
           Case 2: Set mobjPrng = New kiPrng.cKISS
           Case 3: Set mobjPrng = New kiPrng.cMWC
           Case 4: Set mobjPrng = New kiPrng.cMother
           Case 5: Set mobjPrng = New kiPrng.cMT19937
           Case 6: Set mobjPrng = New kiPrng.cMT11231A
           Case 7: Set mobjPrng = New kiPrng.cMT11231B
           Case 8: Set mobjPrng = New kiPrng.cTT800
    End Select
    
    mblnStopPressed = False
    mblnExitPgm = False
    mobjPrng.StopProcessing = mblnStopPressed
    
    With frmBlackjack
        .cmdStart.Visible = True
        .cmdStop.Visible = False
        .lblPause.Visible = False
        .lblPause.Caption = "Creating " & Format$(MAX_SIZE, "#,0") & vbNewLine & " random numbers"
        ' center on the screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless               ' display the screen with no flicker
        .Refresh
    End With

    cmdStart.SetFocus
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
    DoEvents
    Set mobjPrng = Nothing       ' free object from memory
    DoEvents
    
    Unload frmBlackjack          ' deactivate form object
    Set frmBlackjack = Nothing   ' free form object from memory
    
    If mblnExitPgm Then
        End
    End If
    
End Sub

' ---------------------------------------------------------------------------
' In our test, there is no dealer
'
' Below is an excert of rules of blackjack.
'==========================================================================
' Rules for Blackjack (http://www.blackjackinfo.com/)
'
' A blackjack, or natural, is a total of 21 in your first two cards.
'
' basic premise of game is that you want to have a hand value that
' is closer to 21 than that of dealer, without going over 21.
'
' Once all bets are made, dealer will deal cards to players.
' He'll make two passes around table starting at his left (your right)
' so that players and dealer have two cards each.
'
' dealer must play his hand in a specific way, with no choices allowed.
' There are two popular rule variations that determine what totals dealer
' must draw to.   In any given casino, you can tell which rule is in effect
' by looking at blackjack tabletop.  It should be clearly labeled with
' one of these rules:
'
'    "Dealer stands on all 17s":  This is most common rule.  In this
'    case, dealer must continue to take cards ("hit") until his total
'    is 17 or greater.  An Ace in dealer's hand is always counted as
'    11 if possible without dealer going over 21.  For example, (Ace,8)
'    would be 19 and dealer would stop drawing cards ("stand").  Also,
'    (Ace,6) is 17 and again dealer will stand.  (Ace,5) is only 16, so
'    dealer would hit.  He will continue to draw cards until hand's
'    value is 17 or more.  For example, (Ace,5,7) is only 13 so he hits
'    again.  (Ace,5,7,5) makes 18 so he would stop ("stand") at that point.
'
'    "Dealer hits soft 17":  Some casinos use this rule variation instead.
'    This rule is identical except for what happens when dealer has a
'    soft total of 17.  Hands such as (Ace,6),  (Ace,5,Ace), and (Ace,2,4)
'    are all examples of soft 17.  dealer hits these hands, and stands
'    on soft 18 or higher, or hard 17 or higher.  When this rule is used,
'    house advantage against players is slightly increased.
'
' Again, dealer has no choices to make in play of his hand.  He
' cannot split pairs, but must instead simply hit until he reaches at least
' 17 or busts by going over 21.
'
' ---------------------------------------------------------------------------
Private Sub PlayBlackjack()

    Dim blnGoodFinish       As Boolean
    Dim strTemp             As String
    Dim strPath             As String
    Dim strFormat           As String
    Dim astrHands()         As String
    Dim lngIdx              As Long
    Dim lngIndex            As Long
    Dim lngCard1            As Long
    Dim lngCard2            As Long
    Dim lngNextCard         As Long
    Dim lngPoint            As Long
    Dim lngTmpValue         As Long
    Dim lngWin              As Long
    Dim lngLose             As Long
    Dim lngCardCnt          As Long
    Dim lngMax              As Long
    Dim alngTotals(1 To 13) As Long
    Dim adblData()          As Double
    Dim vntCurrCard         As Variant
    
    Const FILE_NAME As String = "Blackjack.txt"
                                     
    Const ROUTINE_NAME As String = "PlayBlackjack"

    On Error GoTo PlayBlackjack_Error

    Erase alngTotals()   ' Always start with empty arrays
    Erase astrHands()
    Erase adblData()
    vntCurrCard = Empty  ' Always start with empty variants
    
    blnGoodFinish = False   ' Preset to bad finish
    lngWin = 0
    lngLose = 0
    lngMax = (MAX_SIZE + 10)
    
    ReDim astrHands(MAX_GAMES)
    
    ' Clear display totals
    For lngIdx = 0 To txtCount.Count - 1
        txtCount(lngIdx).Text = "0"
    Next lngIdx
    
    ' prefill array
    For lngIdx = 1 To 13
        alngTotals(lngIdx) = 0
    Next lngIdx
    
    DoEvents
    If mblnStopPressed Then
        GoTo PlayBlackjack_CleanUp
    End If
        
    strPath = FormatPath()
    strPath = strPath & FILE_NAME
    
    ' Write start time to test result file.
    mhFile = FreeFile
    Open strPath For Output As #mhFile
    Print #mhFile, Format$(MAX_GAMES, "#,0") & " hands of Blackjack"
    Print #mhFile, " "
    
    ' save version information
    Print #mhFile, "Version: " & mobjPrng.Version
    Print #mhFile, " "
    
    Print #mhFile, " Start:  " & Now()
    Close #mhFile
    
    DoEvents
    If mblnStopPressed Then
        GoTo PlayBlackjack_CleanUp
    End If
        
    ' Start playing
    For lngIndex = 1 To MAX_GAMES
        
        DoEvents
        If mblnStopPressed Then
            Exit For    ' exit For..Next loop
        End If
        
        ' is it time to get some more numbers?
        If (lngMax + 4) >= MAX_SIZE Then
            
            Erase adblData()
            lblPause.Visible = True
            DoEvents
            
            ' refill random data array
            Select Case mlngAlgorithm
                   Case 0: adblData() = mobjPrng.BuildRndData(MAX_SIZE, ePRNG_DBL_ARRAY, False)
                   Case 1: adblData() = mobjPrng.ISAAC_Prng(MAX_SIZE, True)
                   Case 2: adblData() = mobjPrng.KISS_Prng(MAX_SIZE, True)
                   Case 3: adblData() = mobjPrng.MWC_Prng(MAX_SIZE, True)
                   Case 4: adblData() = mobjPrng.MOA_Prng(MAX_SIZE, True)
                   Case 5: adblData() = mobjPrng.MT_Prng(MAX_SIZE, True)
                   Case 6: adblData() = mobjPrng.MTA_Prng(MAX_SIZE, True)
                   Case 7: adblData() = mobjPrng.MTB_Prng(MAX_SIZE, True)
                   Case 8: adblData() = mobjPrng.TT800_Prng(MAX_SIZE, True)
            End Select
            lblPause.Visible = False
            lngMax = 1                              ' reset index counter
        End If
          
        DoEvents
        If mblnStopPressed Then
            Exit For    ' exit For..Next loop
        End If
        
        ' update game count
        txtCount(0).Text = Format$(lngIndex, "#,0")
        
        ' initialize variables
        strTemp = vbNullString
        lngCardCnt = 0
        lngPoint = 0
        
        ' zero based array.  Fill with 14 zeroes.
        vntCurrCard = Empty
        vntCurrCard = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
              
        ' Get first two cards
        lngCard1 = Int((Abs(adblData(lngMax)) * 13) + 1)
        lngCard2 = Int((Abs(adblData(lngMax + 1)) * 13) + 1)
          
        ' Point accumulation.  See if one
        ' of the cards is an Ace.
        lngCard1 = CardValue(lngCard1, True, strFormat)
        strTemp = strTemp & strFormat    ' format output record
        
        lngCard2 = CardValue(lngCard2, True, strFormat)
        strTemp = strTemp & strFormat    ' format output record
        
        If lngCard1 = 11 And lngCard2 = 11 Then
            lngPoint = 1 + 1          ' I do not split in this demo
        
        ' Do not take a card 17 or greater
        ElseIf lngCard1 = 11 And lngCard2 > 6 Then
            lngPoint = 11 + lngCard2
            
        ' Do not take a card 17 or greater
        ElseIf lngCard1 > 6 And lngCard2 = 11 Then
            lngPoint = 11 + lngCard1
                
        ' Accrure card totals
        Else
            lngPoint = lngCard1 + lngCard2
        End If
        
        ' final accummulations
        alngTotals(lngCard1) = alngTotals(lngCard1) + 1
        alngTotals(lngCard2) = alngTotals(lngCard2) + 1
        
        ' keep track of cards played.  no more than 4 of each kind
        vntCurrCard(lngCard1) = vntCurrCard(lngCard1) + 1
        vntCurrCard(lngCard2) = vntCurrCard(lngCard2) + 1
        
        lngMax = lngMax + 2  ' track number of elements
        lngCardCnt = lngCardCnt + 1
        
        ' update display
        UpdateDisplay lngCard1
        UpdateDisplay lngCard2
        
        DoEvents
        If mblnStopPressed Then
            Exit For    ' exit For..Next loop
        End If
        
        Do
            DoEvents
            If mblnStopPressed Then
                Exit Do
            End If
            
            lngTmpValue = 0
            
            If lngCardCnt = 5 And lngPoint <= 21 Then
                ' update display
                lngWin = lngWin + 1
                UpdateWin lngWin
                Exit Do
            Else
                ' see what we have
                Select Case lngPoint
                       
                       Case 21       ' We won
                            ' update display
                            lngWin = lngWin + 1
                            UpdateWin lngWin
                            Exit Do
                         
                       Case Is > 21  ' We lost
                            ' update display
                            lngLose = lngLose + 1
                            Call UpdateLose(lngLose)
                            Exit Do
                         
                       Case Is < 17    ' less than 17, get a hit
                            ' Get next card
                            lngMax = lngMax + 1          ' track number of elements
                            lngNextCard = Int((Abs(adblData(lngMax)) * 13) + 1)
                            vntCurrCard(lngNextCard) = vntCurrCard(lngNextCard) + 1
                          
                            ' Process if we do not have 4 of a kind
                            If vntCurrCard(lngNextCard) < 5 Then
                                lngTmpValue = CardValue(lngNextCard, False, strFormat)
                                strTemp = strTemp & strFormat      ' format output record
                                lngPoint = lngPoint + lngTmpValue  ' accumulate total
                                alngTotals(lngNextCard) = alngTotals(lngNextCard) + 1
                                lngCardCnt = lngCardCnt + 1
                                               
                                ' update display
                                UpdateDisplay lngNextCard
                            End If
      
                       Case Else       ' We won over 17 and less than 21
                            ' update display
                            lngWin = lngWin + 1
                            UpdateWin lngWin
                            Exit Do
                End Select
            End If
            
            ' is it time to get some more numbers?
            If (lngMax + 4) >= MAX_SIZE Then
            
                Erase adblData()
                lblPause.Visible = True
                DoEvents
                
                ' refill random data array
                Select Case mlngAlgorithm
                       Case 0: adblData() = mobjPrng.BuildRndData(MAX_SIZE, ePRNG_DBL_ARRAY, False)
                       Case 1: adblData() = mobjPrng.ISAAC_Prng(MAX_SIZE, True)
                       Case 2: adblData() = mobjPrng.KISS_Prng(MAX_SIZE, True)
                       Case 3: adblData() = mobjPrng.MWC_Prng(MAX_SIZE, True)
                       Case 4: adblData() = mobjPrng.MOA_Prng(MAX_SIZE, True)
                       Case 5: adblData() = mobjPrng.MT_Prng(MAX_SIZE, True)
                       Case 6: adblData() = mobjPrng.MTA_Prng(MAX_SIZE, True)
                       Case 7: adblData() = mobjPrng.MTB_Prng(MAX_SIZE, True)
                       Case 8: adblData() = mobjPrng.TT800_Prng(MAX_SIZE, True)
                End Select
                lblPause.Visible = False
                lngMax = 1                          ' reset index counter
            
                DoEvents
                If mblnStopPressed Then
                    Exit Do
                End If
            End If
        
        Loop
        
        DoEvents
        If mblnStopPressed Then
            Exit For    ' exit For..Next loop
        End If
        
        DoEvents
        If lngPoint > 21 Then
            astrHands(lngIndex - 1) = "(-)  " & CStr(lngPoint) & " = " & strTemp ' with total points
        Else
            astrHands(lngIndex - 1) = "(w)  " & CStr(lngPoint) & " = " & strTemp ' with total points
        End If
        
    Next lngIndex
    
    DoEvents
    If mblnStopPressed Then
        GoTo PlayBlackjack_CleanUp
    End If
        
    ' Write finish time to test result file.
    mhFile = FreeFile
    Open strPath For Append As #mhFile
    
    Print #mhFile, "Finish:  " & Now()
    Print #mhFile, " "
    Print #mhFile, "Wins     " & Format$(Format$(lngWin, "#,0"), "@@@@@@@@@@") & _
                         " (" & Format$(lngWin / MAX_SIZE, "Percent") & ")"
    Print #mhFile, "Losses   " & Format$(Format$(lngLose, "#,0"), "@@@@@@@@@@") & _
                         " (" & Format$(lngLose / MAX_SIZE, "Percent") & ")"
    Print #mhFile, " "
    Print #mhFile, "   Card      Times"
    Print #mhFile, "  Values     Played"
       
    ' write totals to output file
    For lngIndex = 1 To 13
        Select Case lngIndex
               Case 1:   Print #mhFile, "    Ace.  " & Format$(Format$(alngTotals(lngIndex), "#,0"), "@@@@@@@@@")
               Case 11:  Print #mhFile, "   Jack.  " & Format$(Format$(alngTotals(lngIndex), "#,0"), "@@@@@@@@@")
               Case 12:  Print #mhFile, "  Queen.  " & Format$(Format$(alngTotals(lngIndex), "#,0"), "@@@@@@@@@")
               Case 13:  Print #mhFile, "   King.  " & Format$(Format$(alngTotals(lngIndex), "#,0"), "@@@@@@@@@")
               Case Else:  Print #mhFile, Format$(lngIndex, "@@@@@@@") & ".  " & _
                                         Format$(Format$(alngTotals(lngIndex), "#,0"), "@@@@@@@@@")
        End Select
    Next lngIndex
    
    Print #mhFile, " "
    Print #mhFile, String$(30, "-")
    
    ' Write each of hands to output file.
    ' Remove everything above this in output file and
    ' import into MS Access to sort.  Excel is limited to
    ' 256 columns x 65536 rows.  Uncomment next loop
    ' to write data to output file.  This file will exceed
    ' 20mb.
    For lngIndex = 0 To MAX_GAMES - 1
        Print #mhFile, astrHands(lngIndex)
    Next lngIndex
    
    Print #mhFile, " "
    Close #mhFile
    blnGoodFinish = True
    
    cmdStop_Click
    
PlayBlackjack_CleanUp:
    Erase alngTotals()    ' Always empty arrays when not needed
    Erase astrHands()
    Erase adblData()
    vntCurrCard = Empty  ' Always empty variants when not needed
    
    If blnGoodFinish Then
        MsgBox "See file " & vbNewLine & strPath
    Else
        For lngIndex = 0 To 17
            txtCount(lngIndex).Text = "0"
        Next lngIndex
    End If
    
    On Error GoTo 0
    Exit Sub

PlayBlackjack_Error:
    MsgBox "frmBlackjack:" & ROUTINE_NAME & vbNewLine & Err.Description
    Resume PlayBlackjack_CleanUp
    
End Sub

Private Sub UpdateDisplay(lngIndex As Long)

    Dim lngCurrValue As Long
    
    If mblnStopPressed Then
        Exit Sub
    End If
        
    With frmBlackjack
        lngCurrValue = CLng(Val(.txtCount(lngIndex).Text)) + 1
        .txtCount(lngIndex).Text = CStr(lngCurrValue)
    End With
    
End Sub
Private Function CardValue(ByVal lngCard As Long, _
                           ByVal blnFirstTwoCards As Boolean, _
                           ByRef strFormat As String) As Long
    
    strFormat = vbNullString
    
    DoEvents
    If mblnStopPressed Then
        Exit Function
    End If
        
    Select Case lngCard
           Case 2 To 10
                CardValue = lngCard
                strFormat = Format$(CardValue, "@@@@")
           Case 11
                strFormat = Format$("  Jk", "@@@@")
                CardValue = 10
           Case 12
                strFormat = Format$("  Qn", "@@@@")
                CardValue = 10
           Case 13
                strFormat = Format$("  Kg", "@@@@")
                CardValue = 10
           Case Else
                DoEvents
                If blnFirstTwoCards Then
                    CardValue = 11
                Else
                    CardValue = 1
                End If
                strFormat = Format$(CStr(CardValue), "@@@@")
    End Select


End Function

Private Sub UpdateLose(ByVal lngValue As Long)

    txtCount(15).Text = Format$(lngValue, "#,0")
    txtCount(17).Text = Format$(lngValue / MAX_GAMES, "Percent")

End Sub

Private Sub UpdateWin(ByVal lngValue As Long)

    txtCount(14).Text = Format$(lngValue, "#,0")
    txtCount(16).Text = Format$(lngValue / MAX_GAMES, "Percent")

End Sub

