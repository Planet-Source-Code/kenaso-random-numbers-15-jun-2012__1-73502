VERSION 5.00
Begin VB.Form frmCraps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rnd Testing - Craps"
   ClientHeight    =   5745
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   4125
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   315
      Left            =   2640
      TabIndex        =   59
      Top             =   5265
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3420
      TabIndex        =   58
      Top             =   5265
      Width           =   615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top"
      Height          =   315
      Left            =   2640
      TabIndex        =   57
      Top             =   5265
      Width           =   615
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   25
      Left            =   2775
      TabIndex        =   56
      Top             =   900
      Width           =   915
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   24
      Left            =   2775
      TabIndex        =   55
      Top             =   675
      Width           =   915
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nbr of games"
      Height          =   390
      Index           =   6
      Left            =   3150
      TabIndex        =   54
      Top             =   1275
      Width           =   615
   End
   Begin VB.Label lblMisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Throws per game"
      Height          =   390
      Index           =   5
      Left            =   2100
      TabIndex        =   53
      Top             =   1275
      Width           =   765
   End
   Begin VB.Label lblMisc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nbr of games"
      Height          =   390
      Index           =   4
      Left            =   1140
      TabIndex        =   52
      Top             =   1275
      Width           =   540
   End
   Begin VB.Label lblMisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Throws per game"
      Height          =   390
      Index           =   3
      Left            =   105
      TabIndex        =   51
      Top             =   1275
      Width           =   765
   End
   Begin VB.Line Line2 
      X1              =   75
      X2              =   3975
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
      Left            =   60
      TabIndex        =   50
      Top             =   5085
      Width           =   2460
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   615
      TabIndex        =   49
      Top             =   2625
      Width           =   1095
   End
   Begin VB.Label lblThrow 
      BackStyle       =   0  'Transparent
      Caption         =   " of  200,000"
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
      Index           =   23
      Left            =   2700
      TabIndex        =   48
      Top             =   450
      Width           =   1065
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
      Height          =   240
      Index           =   22
      Left            =   375
      TabIndex        =   47
      Top             =   450
      Width           =   915
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   23
      Left            =   1500
      TabIndex        =   46
      Top             =   450
      Width           =   1140
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   22
      Left            =   2700
      TabIndex        =   45
      Top             =   4725
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   1875
      X2              =   1875
      Y1              =   1200
      Y2              =   4950
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   21
      Left            =   2700
      TabIndex        =   44
      Top             =   4425
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   20
      Left            =   2700
      TabIndex        =   43
      Top             =   4125
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   19
      Left            =   2700
      TabIndex        =   42
      Top             =   3825
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   18
      Left            =   2700
      TabIndex        =   41
      Top             =   3525
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   17
      Left            =   2700
      TabIndex        =   40
      Top             =   3225
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   16
      Left            =   2700
      TabIndex        =   39
      Top             =   2925
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   15
      Left            =   2700
      TabIndex        =   38
      Top             =   2625
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   14
      Left            =   2700
      TabIndex        =   37
      Top             =   2325
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   13
      Left            =   2700
      TabIndex        =   36
      Top             =   2025
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   12
      Left            =   2700
      TabIndex        =   35
      Top             =   1725
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   11
      Left            =   615
      TabIndex        =   34
      Top             =   4425
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   10
      Left            =   615
      TabIndex        =   33
      Top             =   4125
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   615
      TabIndex        =   32
      Top             =   3825
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   8
      Left            =   615
      TabIndex        =   31
      Top             =   3525
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   615
      TabIndex        =   30
      Top             =   3225
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   615
      TabIndex        =   29
      Top             =   2925
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   615
      TabIndex        =   28
      Top             =   2325
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   615
      TabIndex        =   27
      Top             =   2025
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   615
      TabIndex        =   26
      Top             =   1725
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1500
      TabIndex        =   25
      Top             =   900
      Width           =   1140
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   1500
      TabIndex        =   24
      Top             =   675
      Width           =   1140
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Over 20"
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
      Index           =   21
      Left            =   1725
      TabIndex        =   23
      Top             =   4725
      Width           =   915
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "20"
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
      Index           =   20
      Left            =   2100
      TabIndex        =   22
      Top             =   4425
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "19"
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
      Index           =   19
      Left            =   2100
      TabIndex        =   21
      Top             =   4125
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Index           =   18
      Left            =   2100
      TabIndex        =   20
      Top             =   3825
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "17"
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
      Index           =   17
      Left            =   2100
      TabIndex        =   19
      Top             =   3525
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
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
      Index           =   16
      Left            =   2100
      TabIndex        =   18
      Top             =   3225
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "15"
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
      Index           =   15
      Left            =   2100
      TabIndex        =   17
      Top             =   2925
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "14"
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
      Index           =   14
      Left            =   2100
      TabIndex        =   16
      Top             =   2625
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "13"
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
      Left            =   2100
      TabIndex        =   15
      Top             =   2325
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "12"
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
      Left            =   2100
      TabIndex        =   14
      Top             =   2025
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "11"
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
      Left            =   2100
      TabIndex        =   13
      Top             =   1725
      Width           =   540
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
      Left            =   105
      TabIndex        =   12
      Top             =   4425
      Width           =   540
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
      Left            =   105
      TabIndex        =   11
      Top             =   4125
      Width           =   540
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
      Left            =   105
      TabIndex        =   10
      Top             =   3825
      Width           =   540
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
      Left            =   105
      TabIndex        =   9
      Top             =   3525
      Width           =   540
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
      Left            =   105
      TabIndex        =   8
      Top             =   3225
      Width           =   540
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
      Left            =   105
      TabIndex        =   7
      Top             =   2925
      Width           =   540
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
      Left            =   105
      TabIndex        =   6
      Top             =   2625
      Width           =   540
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
      Left            =   105
      TabIndex        =   5
      Top             =   2325
      Width           =   540
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
      Left            =   105
      TabIndex        =   4
      Top             =   2025
      Width           =   540
   End
   Begin VB.Label lblThrow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Left            =   105
      TabIndex        =   3
      Top             =   1725
      Width           =   540
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
      Caption         =   "Craps Demo"
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
      Width           =   3615
   End
End
Attribute VB_Name = "frmCraps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mhFile          As Long
Private mlngAlgorithm   As Long
Private mblnExitPgm     As Boolean
Private mblnStopPressed As Boolean
Private mobjPrng        As Object

Private Const MAX_GAMES As Long = 50000   ' number of games
Private Const MAX_SIZE  As Long = 100000  ' size of array of random numbers.

Public Sub Algorithm(ByVal lngData As Long)
    
    mlngAlgorithm = lngData
    
    With frmCraps
        Select Case mlngAlgorithm
               Case 0: .Caption = "Craps - CryptoAPI Prng"
               Case 1: .Caption = "Craps - Isaac Prng"
               Case 2: .Caption = "Craps - KISS Prng"
               Case 3: .Caption = "Craps - MWC-SHR3 Prng"
               Case 4: .Caption = "Craps - Mother-of-All Prng"
               Case 5: .Caption = "Craps - MT19937 Prng"
               Case 6: .Caption = "Craps - MT11231-A Prng"
               Case 7: .Caption = "Craps - MT11231-B Prng"
               Case 8: .Caption = "Craps - TT800 Prng"
        End Select
    End With
    
End Sub

Private Sub cmdExit_Click()
    
    DoEvents
    mblnStopPressed = True
    mblnExitPgm = True
    
    DoEvents
    mobjPrng.StopProcessing = True
    
    DoEvents
    Unload Me

End Sub

Private Sub cmdStart_Click()

    cmdStop.Visible = True
    cmdStart.Visible = False
    mblnStopPressed = False
    mobjPrng.StopProcessing = False
    PlayCraps
    
End Sub

Private Sub cmdStop_Click()

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
    
    With frmCraps
        .cmdStart.Visible = True
        .cmdStop.Visible = False
        .lblPause.Visible = False
        .lblPause.Caption = "Creating " & Format$(MAX_SIZE, "#,0") & vbNewLine & " random numbers"
        .lblThrow(23) = "of   " & Format$(MAX_GAMES, "#,0")
        ' center on the screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless               ' display the screen with no flicker
        .Refresh
    End With
    
    cmdStart.SetFocus
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
    Set mobjPrng = Nothing       ' free object from memory
    Unload frmCraps              ' deactivate form object
    Set frmCraps = Nothing       ' free form object from memory
    
    If mblnExitPgm Then
        End
    End If
    
End Sub

' ---------------------------------------------------------------------------
' Test9:  RANDOM NUMBER GENERATOR
'         Rename this routine to "Main" and then press F5 to execute.
'
' This test uses frmCraps.
'
' Excert from DieHard Tests.txt:
'
'      "This is the CRAPS TEST. It plays 200,000 games of craps, finds
'      the number of wins and the number of throws necessary to end
'      each game.  The number of wins should be (very close to) a
'      normal with mean 200000p and variance 200000p(1-p), with
'      p=244/495.  Throws necessary to complete the game can vary
'      from 1 to infinity, but counts for all>21 are lumped with 21.
'      A chi-square test is made on the no.-of-throws cell counts.
'      Each 32-bit integer from the test file provides the value for
'      the throw of a die, by floating to [0,1), multiplying by 6
'      and taking 1 plus the integer part of the result."
'
' NOTE: Most of the tests in DIEHARD return a p-value, which
'      should be uniform on [0,1) if the input file contains truly
'      independent random bits.   Those p-values are obtained by
'      p=F(X), where F is the assumed distribution of the sample
'      random variable X---often normal. But that assumed F is just
'      an asymptotic approximation, for which the fit will be worst
'      in the tails. Thus you should not be surprised with
'      occasional p-values near 0 or 1, such as .0012 or .9983.
'      When a bit stream really FAILS BIG, you will get p's of 0 or
'      1 to six or more places.  By all means, do not, as a
'      Statistician might, think that a p < .025 or p> .975 means
'      that the RNG has "failed the test at the .05 level".  Such
'      p 's happen among the hundreds that DIEHARD produces, even
'      with good RNG's.  So keep in mind that " p happens".
'
'======================================================================
' Rules for Craps (http://www.gambling-systems.om/crapsbasics.html)
'
'                   Craps - Pass line bet
' When it is your turn to throw the craps dice, you must determine
' whether to bet the pass line or the don't pass line. Most shooters,
' as well as most of the other craps players at the table, will bet
' the pass line, as it is the basic wager of craps.
'
' The pass line wager is an even money bet that wins if you either
' roll a total of 7 or 11 on the come-out roll, or if you throw a
' 4, 5, 6, 8, 9 or 10 on the come-out roll and repeat that number
' before you roll a 7. The pass line bet loses if the come-out roll
' is a 2, 3, or 12 (known as "craps") or when a 7 is rolled before
' the established point number is repeated.
'
' If you successfully complete a pass, - that is, if you repeat an
' established point number before throwing a 7-, you get to roll the
' dice again. Only when you seven-out will the stickman push the
' dice to the next player in succession.
'
' Once you have established a point, if you roll a number other than
' your point or a 7, it is disregard as far as pass line bets are
' concerned, although these additional rolls do affect other bets
' that can be made at the craps table.
'
' As an example, suppose you have established a point of 8 on the
' come-out roll. If you next throw a 3, then a 5, a 9, and a 10, these
' numbers will be ignored for pass line bets. But if you then roll 7,
' you will lose your pass line wager, since the 7 came up before your
' point number.
'
' Out of 990 decisions at the craps table you can expect to lose 14
' decisions more than you win. That makes the house advantage at craps
' 1.41%. In other words, out of every $100 that you wager at the craps
' table, you can expect to lose $1.41. Of course this is in the long
' run. You can win because in the relatively short time you will be
' playing, there will be fluctuations in this house edge, so at times
' things will be going in your favor at the craps table.
'
' A pass line bet can be made at any time during a shooter's roll, even
' after he has established a point. However, a bet placed on the pass
' line after a point has been established is a very poor wager, since
' you have missed the opportunity to win on the come-out roll when the
' shooter throws a 7 or an 11. The only way you can now win is if the
' shooter repeats his point before he sevens-out.
' ---------------------------------------------------------------------------
Private Sub PlayCraps()

    Dim blnGameOver As Boolean
    Dim lngIndex    As Long
    Dim lngDice1    As Long
    Dim lngDice2    As Long
    Dim lngPoint    As Long
    Dim lngHold     As Long
    Dim lngWin      As Long
    Dim lngLose     As Long
    Dim lngThrow    As Long
    Dim lngMax      As Long
    Dim strPath     As String
    Dim adblData()  As Double
    Dim arNbrOfTries(1 To 21)  As Long
    
    Const FILE_NAME As String = "Craps.txt"
    
    lngHold = 0
    lngWin = 0
    lngLose = 0
    lngMax = 1
    Erase adblData()
    
    ' Clear display totals
    For lngIndex = 0 To lblCount.Count - 1
        lblCount(lngIndex).Caption = "0"
    Next lngIndex
    
    ' prefill the array
    For lngIndex = 1 To 21
        arNbrOfTries(lngIndex) = 0
    Next
    
    DoEvents
    If mblnStopPressed Then
        GoTo Finished
    End If
        
    strPath = FormatPath()
    strPath = strPath & FILE_NAME
    
    ' Write the start time to the test result file.
    mhFile = FreeFile
    Open strPath For Output As #mhFile
    Print #mhFile, Format$(MAX_GAMES, "#,0") & " Crap Games"
    Print #mhFile, ""
    
    ' save the version information
    Print #mhFile, "Version: " & mobjPrng.Version
    Print #mhFile, ""
    
    Print #mhFile, " Start:  " & Now()
    Close #mhFile
    
    ' Fill with random numbers
    lblPause.Visible = True
    DoEvents
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
    
    DoEvents
    If mblnStopPressed Then
        GoTo Finished
    End If
        
    ' Generate random numbers between 1 and 6
    For lngIndex = 1 To MAX_GAMES
        
        blnGameOver = False
        lngHold = 0
        lngThrow = 0
        
        DoEvents
        If mblnStopPressed Then
            Exit For    ' exit For..Next loop
        End If
        
        Do While Not blnGameOver
        
            DoEvents
            If mblnStopPressed Then
                Exit Do
            End If
        
            ' is it time to get some more numbers?
            If (lngMax + 4) >= MAX_SIZE Then
                
                Erase adblData()
                lblPause.Caption = "Creating " & Format$(MAX_SIZE, "#,0") & " random numbers"
                lblPause.Visible = True
                DoEvents
                
                ' refill the random data array
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
                lngMax = 1                        ' reset the index counter
                lblPause.Visible = False          ' hide the red banner
            End If
            
            DoEvents
            If mblnStopPressed Then
                Exit Do
            End If
        
            ' Calculate the dice
            lngDice1 = Int((Abs(adblData(lngMax)) * 6) + 1)
            lngDice2 = Int((Abs(adblData(lngMax + 1)) * 6) + 1)
            lngMax = lngMax + 2
            
            lngPoint = lngDice1 + lngDice2   ' this is our point to make
            lngThrow = lngThrow + 1          ' track the number of throws
                    
            ' this is the first throw
            If lngThrow = 1 Then
                        
                ' Evaluate our point
                Select Case lngPoint
                       Case 7, 11               ' if 7 or 11, we win
                            lngWin = lngWin + 1
                            blnGameOver = True
                       Case 2, 3, 12            ' if 2, 3, or 12, we lose (Craps)
                            lngLose = lngLose + 1
                            blnGameOver = True
                       Case 4, 5, 6, 8, 9, 10   ' This is our point to make
                            lngHold = lngPoint
                End Select
            Else
                ' this is not the first try
                Select Case lngPoint
                       Case 7                   ' we lose
                            lngLose = lngLose + 1
                            blnGameOver = True
                       Case Is = lngHold        ' we win
                            lngWin = lngWin + 1
                            blnGameOver = True
                       Case Else
                            ' continue trying
                End Select
            End If
                    
            ' keep track of the number of throws to end of game
            If blnGameOver Then
                If lngThrow > 20 Then
                    lngThrow = 21
                End If
                                
                ' update form display
                With frmCraps
                     .lblCount(0) = Format$(lngWin, "#,0")
                     .lblCount(1) = Format$(lngLose, "#,0")
                     .lblCount(24) = Format$(lngWin / MAX_SIZE, "Percent")
                     .lblCount(25) = Format$(lngLose / MAX_SIZE, "Percent")
                     arNbrOfTries(lngThrow) = arNbrOfTries(lngThrow) + 1
                     .lblCount(lngThrow + 1) = Format$(arNbrOfTries(lngThrow), "#,0")
                     .lblCount(23) = Format$(lngIndex, "#,0")   ' game number
                End With
                
                DoEvents  ' allow other processes to function
            End If
        
            DoEvents
            If mblnStopPressed Then
                Exit Do
            End If
        
        Loop
    
        DoEvents
        If mblnStopPressed Then
            Exit For    ' exit For..Next loop
        End If
        
    Next lngIndex
  
    DoEvents
    If mblnStopPressed Then
        GoTo Finished
    End If
        
    ' Write the finish time to the test result file.
    mhFile = FreeFile
    Open strPath For Append As #mhFile
    
    Print #mhFile, "Finish:  " & Now()
    Print #mhFile, ""
    Print #mhFile, "Wins     " & Format$(Format$(lngWin, "#,0"), "@@@@@@@@@@") & _
                         " (" & Format$(lngWin / MAX_SIZE, "Percent") & ")"
    Print #mhFile, "Losses   " & Format$(Format$(lngLose, "#,0"), "@@@@@@@@@@") & _
                         " (" & Format$(lngLose / MAX_SIZE, "Percent") & ")"
    Print #mhFile, ""
    Print #mhFile, "   Throws    Game"
    Print #mhFile, "  per game   Count"
       
    For lngIndex = 1 To 21
        If lngIndex = 21 Then
            Print #mhFile, "Over " & Format$(lngIndex - 1, "@@") & ".    " & _
                          Format$(Format$(arNbrOfTries(lngIndex), "#,0"), "@@@@@@@")
        Else
            Print #mhFile, Format$(lngIndex, "@@@@@@@") & ".    " & _
                           Format$(Format$(arNbrOfTries(lngIndex), "#,0"), "@@@@@@@")
        End If
    Next
    
    Print #mhFile, ""
    Close #mhFile
    
    cmdStop_Click
    MsgBox "See file " & vbNewLine & strPath
    Exit Sub
    
Finished:
    Close #mhFile
    
    For lngIndex = 0 To 25
        lblCount(lngIndex).Caption = "0"
    Next lngIndex

    mhFile = FreeFile
    Open strPath For Output As #mhFile
    Close #mhFile
    
End Sub

