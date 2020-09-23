VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6540
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   6585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   650
      Index           =   0
      Left            =   4320
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Start the wiping process"
      Top             =   5760
      Width           =   650
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   650
      Index           =   1
      Left            =   4335
      Picture         =   "frmMain.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Stop the active process"
      Top             =   5760
      Width           =   650
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   650
      Index           =   3
      Left            =   5760
      Picture         =   "frmMain.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Terminate this application"
      Top             =   5760
      Width           =   650
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   650
      Index           =   2
      Left            =   5040
      Picture         =   "frmMain.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Display About Screen"
      Top             =   5760
      Width           =   650
   End
   Begin VB.Frame fraProgress 
      Height          =   1260
      Left            =   90
      TabIndex        =   38
      Top             =   4365
      Width           =   6375
      Begin VB.PictureBox picProgressBar 
         BackColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   150
         ScaleHeight     =   330
         ScaleWidth      =   6030
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   750
         Width           =   6090
      End
      Begin VB.Label lblProgress 
         BackStyle       =   0  'Transparent
         Caption         =   "lblProgress(1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   195
         TabIndex        =   41
         Top             =   420
         Width           =   6045
      End
      Begin VB.Label lblProgress 
         BackStyle       =   0  'Transparent
         Caption         =   "lblProgress(0)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   40
         Top             =   120
         Width           =   6090
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4350
      Left            =   90
      TabIndex        =   20
      Top             =   0
      Width           =   6360
      Begin VB.PictureBox picTitle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1943
         Picture         =   "frmMain.frx":106A
         ScaleHeight     =   420
         ScaleWidth      =   2475
         TabIndex        =   29
         Top             =   60
         Width           =   2475
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   2910
         Picture         =   "frmMain.frx":1530
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   27
         Top             =   615
         Width           =   510
      End
      Begin VB.Frame fraRecType 
         Height          =   1200
         Left            =   2235
         TabIndex        =   26
         Top             =   1530
         Width           =   1875
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   75
            ScaleHeight     =   900
            ScaleWidth      =   1665
            TabIndex        =   30
            Top             =   180
            Width           =   1665
            Begin VB.OptionButton optType 
               Caption         =   "Continuous"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   75
               TabIndex        =   5
               Top             =   630
               Width           =   1185
            End
            Begin VB.OptionButton optType 
               Caption         =   "Fixed - 80 chars"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   75
               TabIndex        =   4
               ToolTipText     =   "Minimum 2 bytes"
               Top             =   330
               Value           =   -1  'True
               Width           =   1485
            End
            Begin VB.Label lblTitle 
               Caption         =   "Type of record"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   36
               Top             =   0
               Width           =   1275
            End
         End
      End
      Begin VB.Frame fraRecDef 
         Height          =   1530
         Left            =   135
         TabIndex        =   25
         Top             =   2790
         Width           =   6090
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Left            =   75
            ScaleHeight     =   1140
            ScaleWidth      =   1920
            TabIndex        =   32
            Top             =   165
            Width           =   1920
            Begin VB.TextBox txtLength 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   525
               TabIndex        =   10
               Text            =   "0"
               Top             =   720
               Width           =   855
            End
            Begin VB.OptionButton optRecLength 
               Caption         =   "Custom length"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   165
               TabIndex        =   9
               Top             =   420
               Width           =   1635
            End
            Begin VB.OptionButton optRecLength 
               Caption         =   "Predefined length"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   165
               TabIndex        =   8
               Top             =   90
               Value           =   -1  'True
               Width           =   1650
            End
         End
         Begin VB.CheckBox chkDiehard 
            Caption         =   "Build 11mb (approx) binary file for randomness testing."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2250
            TabIndex        =   11
            Top             =   195
            Width           =   3300
         End
         Begin VB.ComboBox cboCustom 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2205
            TabIndex        =   12
            Text            =   "cboCustom"
            Top             =   1095
            Width           =   3585
         End
         Begin VB.ComboBox cboSize 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2205
            TabIndex        =   19
            Text            =   "cboSize"
            Top             =   1095
            Width           =   3585
         End
      End
      Begin VB.Frame fraChars 
         Height          =   1200
         Left            =   135
         TabIndex        =   24
         Top             =   1530
         Width           =   2025
         Begin VB.TextBox txtASCII 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   300
            MaxLength       =   3
            TabIndex        =   17
            Text            =   "0"
            Top             =   435
            Width           =   540
         End
         Begin VB.TextBox txtASCII 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   300
            MaxLength       =   3
            TabIndex        =   18
            Text            =   "255"
            Top             =   810
            Width           =   540
         End
         Begin VB.Label lblTitle 
            Caption         =   "Decimal ranges"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   300
            TabIndex        =   35
            Top             =   180
            Width           =   1515
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Low (0-255)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   975
            TabIndex        =   2
            Top             =   510
            Width           =   915
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "High (0-255)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   975
            TabIndex        =   3
            Top             =   885
            Width           =   900
         End
      End
      Begin VB.Frame fraHexConv 
         Height          =   1200
         Left            =   4185
         TabIndex        =   23
         Top             =   1530
         Width           =   2040
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   960
            Left            =   75
            ScaleHeight     =   960
            ScaleWidth      =   1890
            TabIndex        =   31
            Top             =   180
            Width           =   1890
            Begin VB.OptionButton optHex 
               Caption         =   "Yes (Two Char)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   165
               TabIndex        =   6
               Top             =   330
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.OptionButton optHex 
               Caption         =   "No (Single Char)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   150
               TabIndex        =   7
               Top             =   630
               Width           =   1575
            End
            Begin VB.Label lblTitle 
               Caption         =   "Convert to hex"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   360
               TabIndex        =   37
               Top             =   0
               Width           =   1275
            End
         End
      End
      Begin VB.Frame fraDest 
         Height          =   990
         Left            =   135
         TabIndex        =   22
         Top             =   480
         Width           =   2040
         Begin VB.ComboBox cboDrive 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   0
            Text            =   "cboDrive"
            Top             =   495
            Width           =   1515
         End
         Begin VB.Label lblTitle 
            Caption         =   "Destination drive"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   1515
         End
      End
      Begin VB.Frame fraPrng 
         Height          =   990
         Left            =   4185
         TabIndex        =   21
         Top             =   480
         Width           =   2040
         Begin VB.ComboBox cboPrng 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   225
            TabIndex        =   1
            Text            =   "cboPrng"
            Top             =   495
            Width           =   1650
         End
         Begin VB.Label lblTitle 
            Caption         =   "PRNG Method"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   1515
         End
      End
      Begin VB.Label lblAuthor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kenneth Ives"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2355
         TabIndex        =   28
         Top             =   1230
         Width           =   1650
      End
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "The closer the range variances, the longer it will take to create the data."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   2
      Left            =   150
      TabIndex        =   42
      Top             =   5790
      Width           =   2730
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmMain
'                by Kenneth Ives  kenaso@tx.rr.com
'
' Description:   This is the main screen to determine how the user wants to
'                create their test file.
'
' NOTE:    This application is slow due to the formatting and file creation
'          processes, not the generation of the data.  The primary purpose
'          of this application is to introduce you to more secure ways of
'          creating random values.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@tx.rr.com
'              Created Module
' 29-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Separated out CryptoAPI in BuildDiehard(), BuildContinuous(),
'              BuildFixedLength() routines because it can produce a byte
'              array directly thus creating the output data faster.
' 03-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Set Stop Processing flag for class object if main flag has
'              been set.  See  BuildContinuous(), BuildFixedLength(),
'              BuildDiehard() routines.
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
'              See ProgressBar() routine.
' 15-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote ElapsedTime() routine.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME    As String = "frmMain"
  Private Const PROGRESS_MSG   As String = "Do you have enough space?"
  Private Const KB_1           As Long = 1024         ' 1 Kilobyte (Kibibyte)
  Private Const KB_4           As Long = 4096         ' 4 kb
  Private Const KB_16          As Long = 16384        ' 16 kb
  Private Const KB_32          As Long = 32768        ' 32 kb
  Private Const KB_64          As Long = 65536        ' 64 kb
  Private Const MB_1           As Long = 1048576      ' 1 Megabyte (Mebibyte)
  Private Const GB_1           As Long = 1073741824   ' 1 Gigabyte (Gigibyte)
  Private Const BASE_LENGTH    As Long = 80

' ***************************************************************************
' Module variables
' ***************************************************************************
  Private mstrFilename         As String
  Private mstrPrefix           As String
  Private mstrDriveLtr         As String
  Private mintSizeIndex        As Integer
  Private mintDriveC           As Integer
  Private mlngAlgorithm        As Long
  Private mcurNumberOfBytes    As Currency
  Private mblnHexConvert       As Boolean
  Private mblnUseKeyboardChars As Boolean
  Private mblnFixed            As Boolean
  Private mblnByteSize         As Boolean
  Private mblnKB_Size          As Boolean
  Private mblnMB_Size          As Boolean
  Private mblnGB_Size          As Boolean
  Private mblnPredefined       As Boolean
  Private mblnDiehardTest      As Boolean
  Private mobjBigFiles         As cBigFiles
  Private mobjKeyEdit          As cKeyEdit
  Private mobjFSO              As Scripting.FileSystemObject
  Private mobjPrng             As Object
    
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' This is a rough translation of the GetTickCount API. The
  ' tick count of a PC is only valid for the first 49.7 days
  ' since the last reboot.  When you capture the tick count,
  ' you are capturing the total number of milliseconds elapsed
  ' since the last reboot.  The elapsed time is stored as a
  ' DWORD value. Therefore, the time will wrap around to zero
  ' if the system is run continuously for 49.7 days.
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


' ***************************************************************************
' Routine:       LoadComboBoxes
'
' Description:   This routine will preload the combo box with the most
'                common selections.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-Apr-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 23-DEC-1999  Kenneth Ives  kenaso@tx.rr.com
'              Routine created by kenaso@tx.rr.com
' 30 SEP 2000  Kenneth Ives  kenaso@tx.rr.com
'              Added the available drive letters
' 15 DEC 2001  Kenneth Ives  kenaso@tx.rr.com
'              Entered accurate byte counts
' ***************************************************************************
Private Sub LoadComboBoxes()

    ' Set switches accordingly
    mblnByteSize = True
    mblnKB_Size = False
    mblnMB_Size = False
    mblnGB_Size = False
    cboPrng.Clear

    ' Load the PRNG combo box
    With cboPrng
         .AddItem "CryptoAPI"      ' 0
         .AddItem "ISAAC"          ' 1
         .AddItem "KISS"           ' 2
         .AddItem "MWC"            ' 3
         .AddItem "Mother of All"  ' 4
         .AddItem "MT19937"        ' 5
         .AddItem "MT11231A"       ' 6
         .AddItem "MT11231B"       ' 7
         .AddItem "TT800"          ' 8
         .ListIndex = 0            ' CryptoAPI
    End With

    mlngAlgorithm = 0   ' CryptoAPI algorithm
    ChoosePrefix        ' Get default file prefix
    cboCustom.Clear
    
    ' Load the Custom size combo box
    With cboCustom
         .AddItem "  Bytes [ Minimum 6 bytes ]"
         .AddItem "  Kilobytes [ Length  x  " & Format$(KB_1, "#,##0") & " ]"
         .AddItem "  Megabytes [ Length  x  " & Format$(MB_1, "#,##0") & " ]"
         .AddItem "  Gigabytes [ Length  x  " & Format$(GB_1, "#,##0") & " ]"
         .ListIndex = 0
    End With

    txtLength.Text = "0"
    cboSize.Clear

    ' Load the Size combo box
    With cboSize
         .AddItem " 1 kb (1024 bytes)"
         .AddItem " 2 kb (2048 bytes)"
         .AddItem " 4 kb (4096 bytes)"
         .AddItem " 6 kb (6144 bytes)"
         .AddItem " 8 kb (8192 bytes)"
         .AddItem " 10 kb (10,240 bytes)"
         .AddItem " 16 kb (16,384 bytes)"
         .AddItem " 32 kb (32,768 bytes)"
         .AddItem " 64 kb (65,536 bytes)"
         .AddItem " 128 kb (131,072 bytes)"
         .AddItem " 256 kb (262,144 bytes)"
         .AddItem " 512 kb (524,288 bytes)"
         .AddItem " 1 mb (1,048,576 bytes)"
         .AddItem " 2 mb (2,097,152 bytes)"
         .AddItem " 3 mb (3,145,728 bytes)"
         .AddItem " 4 mb (4,194,304 bytes)"
         .AddItem " 5 mb (5,242,880 bytes)"
         .AddItem " 6 mb (6,291,456 bytes)"
         .AddItem " 7 mb (7,340,032 bytes)"
         .AddItem " 8 mb (8,392,704 bytes)"
         .AddItem " 9 mb (9,437,184 bytes)"
         .AddItem " 10 mb (10,485,760 bytes)"
         .AddItem " 1.44 mb (1,457,664 bytes)"
         .AddItem " 720 kb (730,112 bytes)"
         .ListIndex = 0
    End With

End Sub

Private Sub Load_cboDrive()

    Dim objDrives As Drives
    Dim objDrive  As Drive

    mintDriveC = -1

    ' Load the available drives combo box
    Set objDrives = mobjFSO.Drives  ' get available drive letters
    cboDrive.Clear

    For Each objDrive In objDrives

        ' save drive letter to the combobox
        Select Case objDrive.DriveType
               
               Case Fixed
                    cboDrive.AddItem objDrive.DriveLetter & ":\"
                    
                    ' look for drive C:
                    If InStr(1, objDrive.DriveLetter, "C", vbTextCompare) > 0 Then
                        mintDriveC = cboDrive.ListCount
                    End If
        
               Case Removable
                    ' see if this is a flash drive
                    If objDrive.IsReady Then
                        If objDrive.TotalSize > (MB_1 * 4) Then
                            cboDrive.AddItem objDrive.DriveLetter & ":\"
                        End If
                    End If
        End Select
        
    Next

    ' Set the combo box displays to the first item
    cboDrive.ListIndex = 0

    Set objDrives = Nothing
    Set objDrive = Nothing

End Sub

Private Sub cboCustom_Click()

    lblProgress(0).Caption = PROGRESS_MSG
    lblProgress(1).Caption = vbNullString
    txtLength.Text = "0"
    
    ' Set switches accordingly
    If cboCustom.ListIndex = 0 Then
        mblnByteSize = True
        mblnKB_Size = False
        mblnMB_Size = False
        mblnGB_Size = False
        txtLength.MaxLength = 6
    ElseIf cboCustom.ListIndex = 1 Then
        mblnByteSize = False
        mblnKB_Size = True
        mblnMB_Size = False
        mblnGB_Size = False
        txtLength.MaxLength = 3
    ElseIf cboCustom.ListIndex = 2 Then
        mblnByteSize = False
        mblnKB_Size = False
        mblnMB_Size = True
        mblnGB_Size = False
        txtLength.MaxLength = 3
    ElseIf cboCustom.ListIndex = 3 Then
        mblnByteSize = False
        mblnKB_Size = False
        mblnMB_Size = False
        mblnGB_Size = True
        txtLength.MaxLength = 3
    End If
        
End Sub

Private Sub cboDrive_Click()

    Dim strMsg   As String
    Dim objDrive As Drive

    ' Capture the visible drive letter on the list
    mstrDriveLtr = cboDrive.Text
    strMsg = "Drive " & mstrDriveLtr & " is not available at this time."
    lblProgress(0).Caption = PROGRESS_MSG
    lblProgress(1).Caption = vbNullString

    ' Make sure the destination drive is available.
    If mobjFSO.DriveExists(mstrDriveLtr) Then
        Set objDrive = mobjFSO.GetDrive(mstrDriveLtr)
        If Not objDrive.IsReady Then
            MsgBox strMsg, vbExclamation Or vbOKOnly, "Not available"
            cboDrive.ListIndex = mintDriveC
        End If
    Else
        MsgBox strMsg, vbExclamation Or vbOKOnly, "Not available"
        cboDrive.ListIndex = mintDriveC
    End If

    Set objDrive = Nothing

End Sub

Private Sub cboPrng_Click()

    ' Capture the index of item selected
    mlngAlgorithm = cboPrng.ListIndex
    
    ChoosePrefix   ' Select file prefix name
    lblProgress(0).Caption = PROGRESS_MSG
    lblProgress(1).Caption = vbNullString
    
End Sub

Private Sub cboSize_Click()

    Const ONE_KB As Currency = 1024@
    
    ' Capture the index and title of item selected
    mintSizeIndex = cboSize.ListIndex
    mblnPredefined = True
    
    lblProgress(0).Caption = PROGRESS_MSG
    lblProgress(1).Caption = vbNullString

    ' Based on the index selection,
    ' get file length and filename
    Select Case mintSizeIndex
           Case 0:    mstrFilename = "_1KB":   mcurNumberOfBytes = ONE_KB
           Case 1:    mstrFilename = "_2KB":   mcurNumberOfBytes = ONE_KB * 2@
           Case 2:    mstrFilename = "_4KB":   mcurNumberOfBytes = ONE_KB * 4@
           Case 3:    mstrFilename = "_6KB":   mcurNumberOfBytes = ONE_KB * 6@
           Case 4:    mstrFilename = "_8KB":   mcurNumberOfBytes = ONE_KB * 8@
           Case 5:    mstrFilename = "_10KB":  mcurNumberOfBytes = ONE_KB * 10@
           Case 6:    mstrFilename = "_16KB":  mcurNumberOfBytes = ONE_KB * 16@
           Case 7:    mstrFilename = "_32KB":  mcurNumberOfBytes = ONE_KB * 32@
           Case 8:    mstrFilename = "_64KB":  mcurNumberOfBytes = ONE_KB * 64@
           Case 9:    mstrFilename = "_128KB": mcurNumberOfBytes = ONE_KB * 128@
           Case 10:   mstrFilename = "_256KB": mcurNumberOfBytes = ONE_KB * 256@
           Case 11:   mstrFilename = "_512KB": mcurNumberOfBytes = ONE_KB * 512@
           Case 12:   mstrFilename = "_1MB":   mcurNumberOfBytes = ONE_KB * ONE_KB
           Case 13:   mstrFilename = "_2MB":   mcurNumberOfBytes = ONE_KB * 2048@
           Case 14:   mstrFilename = "_3MB":   mcurNumberOfBytes = ONE_KB * 3072@
           Case 15:   mstrFilename = "_4MB":   mcurNumberOfBytes = ONE_KB * 4096@
           Case 16:   mstrFilename = "_5MB":   mcurNumberOfBytes = ONE_KB * 5120@
           Case 17:   mstrFilename = "_6MB":   mcurNumberOfBytes = ONE_KB * 6144@
           Case 18:   mstrFilename = "_7MB":   mcurNumberOfBytes = ONE_KB * 7168@
           Case 19:   mstrFilename = "_8MB":   mcurNumberOfBytes = ONE_KB * 8192@
           Case 20:   mstrFilename = "_9MB":   mcurNumberOfBytes = ONE_KB * 9216@
           Case 21:   mstrFilename = "_10MB":  mcurNumberOfBytes = ONE_KB * 10240@
           Case 22:   mstrFilename = "_144MB": mcurNumberOfBytes = 1457664@
           Case 23:   mstrFilename = "_720KB": mcurNumberOfBytes = 730112@
           Case Else: mstrFilename = vbNullString:       mcurNumberOfBytes = 0@
    End Select

    If Len(Trim$(mstrFilename)) > 0 Then
    
        If mblnHexConvert Then
            mstrFilename = mstrPrefix & mstrFilename & ".txt"
        Else
            mstrFilename = mstrPrefix & mstrFilename & ".bin"
        End If
    End If
    
End Sub

Private Sub chkDiehard_Click()

    lblProgress(0).Caption = PROGRESS_MSG
    lblProgress(1).Caption = vbNullString
    
    If chkDiehard.Value = vbChecked Then
        
        optType_Click 0
        
        With frmMain
             If mblnHexConvert Then
                 .optHex(0).Value = False
                 .optHex(1).Value = True
                 mblnHexConvert = False
             End If
            
             If mblnFixed Then
                 .optType(0).Value = False
                 .optType(1).Value = True
                 mblnFixed = False
             End If
            
             .txtASCII(0).Text = "0"
             .txtASCII(1).Text = "255"
             .cboSize.Visible = True
             .cboSize.Enabled = False
             .cboCustom.Visible = False
             .cboCustom.Enabled = False
             .optRecLength(0).Enabled = False
             .optRecLength(1).Enabled = False
             .fraChars.Enabled = False
             .fraHexConv.Enabled = False
             .fraRecType.Enabled = False
             .fraHexConv.Enabled = False
        End With

        mblnDiehardTest = True

    Else

        With frmMain
             .fraChars.Enabled = True
             .fraHexConv.Enabled = True
             .fraRecType.Enabled = True
             .fraHexConv.Enabled = True
             .cboSize.Enabled = True
             .optRecLength(0).Enabled = True
             .optRecLength(1).Enabled = True
        End With

        mblnDiehardTest = False

    End If

End Sub

Private Sub StartProcessing()

    Dim curRecordLength As Currency
    Dim curFreeSpace    As Currency
    Dim curTemp         As Currency
    Dim intLow          As Integer
    Dim intHigh         As Integer
    Dim hFile           As Long
    Dim lngStartTime    As Long
    Dim strMsg          As String
    Dim strPath         As String
    Dim strElapsed      As String
    Dim objDrive        As Drive

    Const ROUTINE_NAME As String = "StartProcessing"

    On Error GoTo StartProcessing_Error
       
    mcurNumberOfBytes = 0@
    curFreeSpace = 0@
    strMsg = vbNullString   ' clear the message area
    intLow = Val(txtASCII(0).Text)

    ' Verify we have a file size selected
    If mblnPredefined And Not mblnDiehardTest Then
        cboSize_Click  ' verify we have data
    End If

    ' Test record sizes
    If mblnPredefined Then

        If mblnDiehardTest Then
            mstrFilename = mstrPrefix & ".bin"
            mcurNumberOfBytes = 11468800@
        End If

    Else
        ' Custom sizes
        If Len(Trim$(txtLength.Text)) = 0 Or txtLength.Text = "0" Then
            InfoMsg "Will not build a zero byte file."
            GoTo StartProcessing_CleanUp
        End If

        ' Test record sizes for values
        If Val(txtLength.Text) < 6 And mblnByteSize Then
            InfoMsg "Minimum file size is 6 bytes."
            GoTo StartProcessing_CleanUp
        End If

        ' See if a valid record length has been entered
        ' Test input if user opted to custimize the file length
        ' calculate the file size
        If mblnByteSize Then                          ' bytes
            
            curRecordLength = CCur(txtLength.Text)
            
            If mblnHexConvert Then
                mstrFilename = mstrPrefix & "_" & Trim$(txtLength.Text) & "Bytes.txt"
            Else
                mstrFilename = mstrPrefix & "_" & Trim$(txtLength.Text) & "Bytes.bin"
            End If
            
        ElseIf mblnKB_Size Then                      ' kilobytes
            curTemp = CCur(txtLength.Text)
            curRecordLength = (curTemp * KB_1)

            If mblnHexConvert Then
                mstrFilename = mstrPrefix & "_" & Trim$(txtLength.Text) & "KB.txt"
            Else
                mstrFilename = mstrPrefix & "_" & Trim$(txtLength.Text) & "KB.bin"
            End If
            
        ElseIf mblnMB_Size Then                      ' megabytes
            curTemp = CCur(txtLength.Text)
            curRecordLength = (curTemp * MB_1)

            If mblnHexConvert Then
                mstrFilename = mstrPrefix & "_" & Trim$(txtLength.Text) & "MB.txt"
            Else
                mstrFilename = mstrPrefix & "_" & Trim$(txtLength.Text) & "MB.bin"
            End If

        ElseIf mblnGB_Size Then                      ' gigabytes
            curTemp = CCur(txtLength.Text)
            curRecordLength = (curTemp * GB_1)

            If mblnHexConvert Then
                mstrFilename = mstrPrefix & "_" & Trim$(txtLength.Text) & "GB.txt"
            Else
                mstrFilename = mstrPrefix & "_" & Trim$(txtLength.Text) & "GB.bin"
            End If
        End If

        mcurNumberOfBytes = curRecordLength

    End If

    Set objDrive = mobjFSO.GetDrive(mstrDriveLtr)  ' Capture drive information
    curFreeSpace = CCur(objDrive.FreeSpace)        ' free space
    
    ' See if the drive is available
    If objDrive.IsReady Then
        
        strPath = QualifyPath(App.Path) & "Data_Output"
        
        ' See if there is a temp folder availalble
        If Not mobjFSO.FolderExists(strPath) Then
            mobjFSO.CreateFolder strPath
        End If
        
        strPath = QualifyPath(strPath)
        mstrFilename = strPath & mstrFilename
        
    Else
        InfoMsg "Drive " & mstrDriveLtr & " is not available at this time."
        GoTo StartProcessing_CleanUp
    End If
    
    Set objDrive = Nothing   ' free object form memory
    
    ' Test available disk space against the file size requested
    If mcurNumberOfBytes > curFreeSpace Then
        strMsg = "Your requested file size is "
        strMsg = strMsg & Format$(mcurNumberOfBytes, "#,##0") & " bytes." & Space$(5)
        strMsg = strMsg & vbNewLine & vbNewLine
        strMsg = strMsg & "This exceeds the desitnation drive (" & mstrDriveLtr & ") whose total" & vbNewLine
        strMsg = strMsg & "free space is only " & Format$(curFreeSpace, "#,##0") & " bytes."
        strMsg = strMsg & vbNewLine & vbNewLine
        strMsg = strMsg & "Try another drive or reduce your file size."

        InfoMsg strMsg
        GoTo StartProcessing_CleanUp
    End If

    ' Format the filename to point to the root directory of the destination drive
    intLow = Val(txtASCII(0).Text)
    intHigh = Val(txtASCII(1).Text)
    lblProgress(0).Caption = mstrFilename
    UpdateDisplay 0
    LockDownCtls
    
    ' If the file already exists
    ' then make sure it is empty
    If IsPathValid(mstrFilename) Then
        hFile = FreeFile
        Open mstrFilename For Output As #hFile
        Close #hFile
    End If
    
    lngStartTime = GetTickCount()  ' starting time
    
    ' Build the test file
    If mblnFixed Then

        ' create fixed length records
        If BuildFixedLength(intLow, intHigh) Then
            UpdateDisplay mcurNumberOfBytes       ' Update progress display
        Else
            GoTo StartProcessing_CleanUp
        End If
    Else
        If mblnDiehardTest Then
        
            If BuildDiehard() Then
                UpdateDisplay mcurNumberOfBytes   ' Update progress display
            Else
                GoTo StartProcessing_CleanUp
            End If
        Else
            ' build one contiguous record
            If BuildContinuous(intLow, intHigh) Then
                UpdateDisplay mcurNumberOfBytes   ' Update progress display
            Else
                GoTo StartProcessing_CleanUp
            End If
        End If
    End If

    strElapsed = ElapsedTime(GetTickCount() - lngStartTime)
    
    ' Reset mouse pointer and display message
    InfoMsg mstrFilename & Space$(4) & vbNewLine & _
            "Elapsed time:  " & strElapsed
    
StartProcessing_CleanUp:
    ' Reset mouse pointer and display message
    Screen.MousePointer = vbDefault
    Set objDrive = Nothing   ' free object form memory
    
    ' the Stop Button is pressed
    DoEvents
    If gblnStopProcessing Then
        
        ' If the file already exists
        ' then make sure it is empty
        If IsPathValid(mstrFilename) Then
                    
            CloseAllFiles  ' Verify all opened files have been closed
            
            hFile = FreeFile
            Open mstrFilename For Output As #hFile
            Close #hFile
                        
            ' Now remove unwanted file
            DoEvents
            On Error Resume Next
            Kill mstrFilename
            On Error GoTo 0
        
        End If
    End If
    
    mcurNumberOfBytes = 0@
    curFreeSpace = 0@
    strMsg = vbNullString   ' clear the message area
    intLow = Val(txtASCII(0).Text)
    Load_cboDrive
    Refresh
    DoEvents
    
    On Error GoTo 0
    Exit Sub

StartProcessing_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume StartProcessing_CleanUp

End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index
            
           Case 0   ' Go button
                gblnStopProcessing = False  ' reset the STOP flag
                LockDownCtls
                ResetProgressMsg
                StartProcessing
                ResetProgressMsg
                UnlockCtls
                
           Case 1   ' Stop button
                gblnStopProcessing = True   ' reset the STOP flag
                ResetProgressMsg
                UnlockCtls
                Screen.MousePointer = vbDefault  ' Reset mouse pointer
           
           Case 2   ' Show About screen
                frmMain.Hide
                frmAbout.DisplayAbout
           
           Case Else  ' End application
                gblnStopProcessing = True
                TerminateProgram
    End Select
    
End Sub

Private Sub Form_Load()

    Set mobjFSO = New Scripting.FileSystemObject
    Set mobjKeyEdit = New cKeyEdit
    
    LoadComboBoxes
    Load_cboDrive
    
    ' initialize variables
    mintSizeIndex = 0
    mstrFilename = "_1KB"
    mblnFixed = True
    mblnHexConvert = True
    mblnPredefined = True
    mblnDiehardTest = False
    mblnUseKeyboardChars = True
    
    ' Initialize screen
    With frmMain
        .Caption = gstrVersion
        .lblProgress(0).Caption = PROGRESS_MSG
        .lblProgress(1).Caption = vbNullString
        .chkDiehard.Value = vbUnchecked
        .optType(0).Value = True
        .optType(1).Value = False
        .optHex(0).Value = True
        .optHex(1).Value = False
        .optRecLength(0).Value = True
        .optRecLength(1).Value = False
        .txtLength.Text = 0
        .txtLength.Enabled = False
        .cboSize.Enabled = True
        .cboSize.Visible = True
        .cboCustom.Enabled = False
        .cboCustom.Visible = False
        
        cboDrive_Click                 ' get current drive information
        cboSize_Click
        
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless               ' display the screen with no flicker
        .Refresh
    End With
    
    DisableX frmMain
    mobjKeyEdit.CenterCaption frmMain
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' If class object is still
    ' active then shut it down
    If Not mobjPrng Is Nothing Then
        mobjPrng.StopProcessing = True  ' Stop any processing
        DoEvents                        ' Allow time to stop
    End If
    
    Set mobjPrng = Nothing      ' free class objects from memory
    Set mobjBigFiles = Nothing
    Set mobjFSO = Nothing
    Set mobjKeyEdit = Nothing
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub optHex_Click(Index As Integer)
    
    ' Set the appropriate switch based on option choice
    mblnHexConvert = CBool(optHex(0).Value)
    
End Sub

Private Sub optRecLength_Click(Index As Integer)

    ' Set the appropriate switch based on option choice
    mblnPredefined = CBool(optRecLength(0).Value)

    Select Case Index
           Case 0
                txtLength.Text = "0"
                txtLength.Enabled = False
                cboSize.Enabled = True
                cboSize.Visible = True
                cboCustom.Visible = False
                cboCustom.Enabled = False
           Case Else
                txtLength.Enabled = True
                txtLength.Text = "0"
                txtLength.SetFocus
                cboSize.Visible = False
                cboSize.Enabled = False
                cboCustom.Enabled = True
                cboCustom.Visible = True
    End Select

End Sub

Private Sub optType_Click(Index As Integer)
    
    ' Set the appropriate switch based on option choice
    mblnFixed = CBool(optType(0).Value)
    
End Sub

Private Sub txtASCII_Change(Index As Integer)

    ' Prevent user from pasting a non-numeric value
    ' into this textbox
    If Not IsNumeric(txtASCII(Index).Text) Then
        txtASCII(Index).Text = vbNullString
    End If

End Sub

Private Sub txtASCII_GotFocus(Index As Integer)
    ' Highlight all the text in the box
    mobjKeyEdit.TextBoxFocus txtASCII(Index)
End Sub

Private Sub txtASCII_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    ' Process any key combinations
    mobjKeyEdit.TextBoxKeyDown txtASCII(Index), KeyCode, Shift
End Sub

Private Sub txtASCII_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Save only numeric and the backspace character
    mobjKeyEdit.ProcessNumericOnly KeyAscii
End Sub

Private Sub txtASCII_LostFocus(Index As Integer)

    ' Test for data
    If Len(Trim$(txtASCII(Index).Text)) = 0 Then
        txtASCII(Index).Text = 0
    End If

    ' Evaluate low and high values
    If Val(txtASCII(0).Text) < 0 Or Val(txtASCII(0).Text) > 255 Then
        MsgBox "Low decimal value must be a positive value of 0-255.", _
               vbOKOnly Or vbExclamation, "Invalid decimal input"
        txtASCII(0).SetFocus
        Exit Sub
    End If

    If Val(txtASCII(1).Text) < 0 Or Val(txtASCII(1).Text) > 255 Then
        MsgBox "High decimal value must be a positive value of 0-255.", _
               vbOKOnly Or vbExclamation, "Invalid decimal input"
        txtASCII(1).SetFocus
        Exit Sub
    End If

End Sub

Private Sub txtLength_Change()

    ' Prevent user from pasting a non-numeric value
    ' into this textbox
    If Not IsNumeric(txtLength.Text) Then
        txtLength.Text = vbNullString
    End If

End Sub

Private Sub txtLength_GotFocus()
    ' Highlight all the text in the box
    mobjKeyEdit.TextBoxFocus txtLength
End Sub

Private Sub txtLength_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Process any key combinations
    mobjKeyEdit.TextBoxKeyDown txtLength, KeyCode, Shift
End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)
    ' Save only numeric and the backspace character
    mobjKeyEdit.ProcessNumericOnly KeyAscii
End Sub

' ***************************************************************************
' Routine:       BuildDiehard
'
' Description:   This routine will build an 11mb (11,468,800 bytes) binary
'                file for use in DieHard and ENT randomness testing.
'
' Test file size:  2,867,200 32-bit random integers (11,468,800 bytes)
'                  11,468,800 bytes = x * y * 4
'                                     |   |   |__ 4 bytes = 1 long integer
'                                     |   |__ # of writes to output file
'                                     |__ # of long integers per array
'
'                When finished, copy to \Diehard folder and execute the
'                application Diehard.exe   You will be prompted for this
'                filename and then for the new output name.  I use MT_DH.TXT
'                to represent Mersenne Twister diehard test output. Use the
'                ".BIN" file as the test file for both Diehard and ENT testing.
'
'                Diehard by George Marsaglia
'                http://stat.fsu.edu/pub/diehard/
'                Click on the link labeled "Windows Software (732 Kb)".
'
'                ENT test site - http://www.fourmilab.ch/random/
'                Scroll down and click on the link "Download Random.zip"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-SEP-2002  Kenneth Ives  kenaso@tx.rr.com
'              Created routine
' 21-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Function BuildDiehard() As Boolean

    Dim hFile        As Long
    Dim lngAmtLeft   As Long
    Dim lngArraySize As Long
    Dim lngPointer   As Long
    Dim alngData()   As Long
    Dim abytData()   As Byte
        
    Const FILE_SIZE    As Long = 11468800    ' Max output file size
    Const ROUTINE_NAME As String = "BuildDiehard"

    On Error GoTo BuildDiehard_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
      
    Screen.MousePointer = vbHourglass
    
    Set mobjBigFiles = New cBigFiles
    
    lngPointer = 1     ' init pointer for output file
    Erase alngData()   ' empty arrays
    Erase abytData()
    
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
    
    mobjPrng.StopProcessing = gblnStopProcessing
    
    hFile = FreeFile                         ' capture first free file handle
    Open mstrFilename For Output As #hFile   ' Create an empty file
    Close #hFile                             ' close the file

    hFile = FreeFile                                  ' capture first free file handle
    Open mstrFilename For Binary Access Write As #hFile   ' re-open file in binary mode

    If mlngAlgorithm = 0 Then
        lngAmtLeft = FILE_SIZE        ' Number of bytes needed
    Else
        lngAmtLeft = (FILE_SIZE \ 4)  ' Number of long integers needed
    End If
    
    Do
        If gblnStopProcessing Then
            DoEvents
            mobjPrng.StopProcessing = True
            DoEvents
            Exit Do   ' Exit Do..Loop
        End If
        
        ' Calc array size
        Select Case lngAmtLeft
               Case Is >= KB_32: lngArraySize = KB_32
               Case Else:        lngArraySize = lngAmtLeft
        End Select
        
        Select Case mlngAlgorithm
               Case 0: abytData() = mobjPrng.BuildRndData(lngArraySize, , False)
               Case 1: alngData() = mobjPrng.ISAAC_Prng(lngArraySize, False)
               Case 2: alngData() = mobjPrng.KISS_Prng(lngArraySize, False)
               Case 3: alngData() = mobjPrng.MWC_Prng(lngArraySize, False)
               Case 4: alngData() = mobjPrng.MOA_Prng(lngArraySize, False)
               Case 5: alngData() = mobjPrng.MT_Prng(lngArraySize, False)
               Case 6: alngData() = mobjPrng.MTA_Prng(lngArraySize, False)
               Case 7: alngData() = mobjPrng.MTB_Prng(lngArraySize, False)
               Case 8: alngData() = mobjPrng.TT800_Prng(lngArraySize, False)
        End Select
        
        If mlngAlgorithm = 0 Then
            ReDim Preserve abytData(lngArraySize - 1)         ' Resize array
            Put #hFile, lngPointer, abytData()                ' Write data to output file
            lngPointer = lngPointer + UBound(abytData)        ' Update pointer for output file
        Else
            ReDim Preserve alngData(lngArraySize - 1)         ' Resize array
            Put #hFile, lngPointer, alngData()                ' Write data to output file
            lngPointer = lngPointer + (UBound(alngData) * 4)  ' Update pointer for output file
        End If
        
        lngAmtLeft = lngAmtLeft - (lngArraySize - 1)          ' Calculate what is left
        UpdateDisplay lngPointer                              ' Update progress bar display
            
        If lngAmtLeft <= 1 Then
            Exit Do
        End If
    
    Loop
    
BuildDiehard_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        mobjPrng.StopProcessing = gblnStopProcessing
        DoEvents
        BuildDiehard = False
    Else
        BuildDiehard = True
    End If
    
    CloseAllFiles            ' Verify all opened files have been closed
    Set mobjPrng = Nothing   ' Free objects from memory when not needed
    Erase alngData()         ' empty arrays
    Erase abytData()
    
    Screen.MousePointer = vbDefault  ' Reset mouse cursor
    On Error GoTo 0
    Exit Function

    ' User pressed the STOP button or an error occurred
BuildDiehard_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume BuildDiehard_CleanUp

End Function

' ***************************************************************************
' Routine:       BuildFixedLength
'
' Description:   This routine will build a file with fixed length records.
'                These will be 80 bytes in length.
'
' Parameters:    intLow - Lowest byte value requested
'                intHigh - Highest byte value requested
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@tx.rr.com
'              Created routine
' 21-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Function BuildFixedLength(ByVal intLow As Integer, _
                                  ByVal intHigh As Integer) As Boolean
    
    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim hFile        As Long
    Dim lngIdx       As Long
    Dim lngIndex     As Long
    Dim lngBlockSize As Long
    Dim alngRand()   As Long
    Dim abytRand()   As Byte
    Dim abytData()   As Byte
    
    Const ROUTINE_NAME As String = "BuildFixedLength"
        
    On Error GoTo BuildFixedLength_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
      
    Set mobjBigFiles = New cBigFiles
    
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
    
    mobjPrng.StopProcessing = gblnStopProcessing
    
    hFile = FreeFile                        ' capture first free file handle
    Open mstrFilename For Output As #hFile  ' Create an empty file
    Close #hFile                            ' close the file

    ' Open a new file
    If Not mobjBigFiles.OpenReadWrite(mstrFilename, hFile) Then
        gblnStopProcessing = True
        GoTo BuildFixedLength_CleanUp
    End If
    
    curAmtLeft = mcurNumberOfBytes
    curFilePos = 0@
    
    If mlngAlgorithm = 0 Then
        
        Do While curAmtLeft > 0@
        
            If gblnStopProcessing Then
                DoEvents
                mobjPrng.StopProcessing = True
                DoEvents
                Exit Do   ' Exit Do..Loop
            End If
        
            Erase abytRand()   ' Start with an empty array
            abytRand() = mobjPrng.BuildWithinRange(KB_64, intLow, intHigh, , False)
        
            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit Do    ' exit Do..Loop
            End If
        
            For lngIndex = 0 To (UBound(abytRand) - 1) Step BASE_LENGTH
                                            
                lngBlockSize = GetBlockSize(curAmtLeft)
                If lngBlockSize > BASE_LENGTH Then
                    lngBlockSize = BASE_LENGTH
                End If
                                            
                Erase abytData()                                                  ' Empty array
                ReDim abytData(lngBlockSize)                                      ' Size array
                CopyMemory abytData(0), abytRand(lngIndex), lngBlockSize          ' Copy data from byte array to another
                ReDim Preserve abytData(lngBlockSize - 1)                         ' Resize array
                UpdateFixedLengthFile abytData(), curAmtLeft, curFilePos, hFile   ' Update output file
                
                ' An error occurred or user opted to STOP processing
                DoEvents
                If gblnStopProcessing Then
                    Exit For    ' exit For..Next loop
                End If
            
                UpdateDisplay curFilePos   ' Update progress bar display
                
                If curAmtLeft <= 0@ Then
                    Exit For   ' exit For..Next loop
                End If
                
            Next lngIndex
            
            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit Do    ' exit Do..Loop
            End If
            
            If curAmtLeft <= 0@ Then
                Exit Do    ' exit Do..Loop
            End If
            
        Loop
    
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo BuildFixedLength_CleanUp
        End If
    
    Else
        
        Do While curAmtLeft > 0@
        
            If gblnStopProcessing Then
                DoEvents
                mobjPrng.StopProcessing = True
                DoEvents
                Exit Do   ' Exit Do..Loop
            End If
        
            lngIdx = 0
            Erase abytData()
            ReDim abytData(KB_64)
            
            Do While lngIdx < KB_64
                
                Erase alngRand()   ' Start with an empty array
                    
                ' Fill byte array with random values
                Select Case mlngAlgorithm
                       Case 1: alngRand() = mobjPrng.ISAAC_Prng(KB_16, False)
                       Case 2: alngRand() = mobjPrng.KISS_Prng(KB_16, False)
                       Case 3: alngRand() = mobjPrng.MWC_Prng(KB_16, False)
                       Case 4: alngRand() = mobjPrng.MOA_Prng(KB_16, False)
                       Case 5: alngRand() = mobjPrng.MT_Prng(KB_16, False)
                       Case 6: alngRand() = mobjPrng.MTA_Prng(KB_16, False)
                       Case 7: alngRand() = mobjPrng.MTB_Prng(KB_16, False)
                       Case 8: alngRand() = mobjPrng.TT800_Prng(KB_16, False)
                End Select
                
                ' An error occurred or user opted to STOP processing
                If gblnStopProcessing Then
                    Exit Do    ' exit Do..Loop
                End If
        
                ReDim abytRand(KB_64)  ' Size byte array
                
                ' Convert long array to byte array
                CopyMemory abytRand(0), alngRand(0), KB_64
            
                For lngIndex = 0 To UBound(abytRand) - 1
                                                
                    Select Case abytRand(lngIndex)
                           Case intLow To intHigh
                                abytData(lngIdx) = abytRand(lngIndex)
                                lngIdx = lngIdx + 1
                    End Select
                    
                    ' If enough data collected then exit loop
                    If lngIdx = KB_64 Then
                        Exit For    ' exit For..Next loop
                    End If
                                                                            
                Next lngIndex
        
                ' An error occurred or user opted to STOP processing
                If gblnStopProcessing Then
                    Exit Do    ' exit Do..Loop
                End If
        
                ' If enough data collected then exit loop
                If lngIdx = KB_64 Then
                    Exit Do    ' exit Do..Loop
                End If
                                                                            
            Loop
        
            ReDim abytRand(KB_64)  ' Size byte array
                
            CopyMemory abytRand(0), abytData(0), KB_64
        
            For lngIndex = 0 To (UBound(abytRand) - 1) Step BASE_LENGTH
                                            
                lngBlockSize = GetBlockSize(curAmtLeft)
                If lngBlockSize > BASE_LENGTH Then
                    lngBlockSize = BASE_LENGTH
                End If
                                            
                Erase abytData()                                                  ' Empty array
                ReDim abytData(lngBlockSize)                                      ' Size array
                CopyMemory abytData(0), abytRand(lngIndex), lngBlockSize          ' Copy data from byte array to another
                ReDim Preserve abytData(lngBlockSize - 1)                         ' Resize array
                UpdateFixedLengthFile abytData(), curAmtLeft, curFilePos, hFile   ' Update output file
                
                UpdateDisplay curFilePos   ' Update progress bar display
                
                If curAmtLeft <= 0@ Then
                    Exit For   ' exit For..Next loop
                End If
                
            Next lngIndex
            
            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit Do    ' exit Do..Loop
            End If
            
            If curAmtLeft <= 0@ Then
                Exit Do    ' exit Do..Loop
            End If
            
        Loop
    
    End If
    
BuildFixedLength_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        mobjPrng.StopProcessing = gblnStopProcessing
        DoEvents
        BuildFixedLength = False
    Else
        BuildFixedLength = True
    End If
    
    mobjBigFiles.API_CloseFile hFile  ' Close open file
    
    CloseAllFiles            ' Verify all opened files have been closed
    Set mobjPrng = Nothing   ' Free objects from memory when not needed
    Set mobjBigFiles = Nothing
    
    On Error GoTo 0
    Exit Function
    
BuildFixedLength_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume BuildFixedLength_CleanUp
  
End Function

Private Sub UpdateFixedLengthFile(ByRef abytData() As Byte, _
                                  ByRef curAmtLeft As Currency, _
                                  ByRef curFilePos As Currency, _
                                  ByVal hFile As Long)
    
    ' Called by BuildFixedLength()
    
    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngPointer As Long
    Dim strRecord  As String
    
    lngLength = UBound(abytData) - 1  ' Capture data length
    lngPointer = 1                    ' set to first positionin output string
    strRecord = Space$(MAX_SIZE)      ' preload output string with spaces
    
    ' build the output record
    If mblnHexConvert Then
    
        ' Loop thru the byte array and build the output record
        For lngIndex = 0 To lngLength + 1
          
            ' load with the hex values
            Mid$(strRecord, lngPointer, 2) = Right$("00" & Hex$(abytData(lngIndex)), 2)
            lngPointer = lngPointer + 2
        
        Next lngIndex
        
        strRecord = UCase$(Trim$(strRecord))
        strRecord = Left$(strRecord, lngLength) & vbNewLine
        
    Else
    
        strRecord = ByteArrayToString(abytData())
        Mid$(strRecord, lngLength, 2) = vbNewLine
        strRecord = UCase$(strRecord)
        
    End If
    
    Erase abytData()
    abytData() = StringToByteArray(strRecord)
    
    ' Write to target file
    If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytData()) Then
        gblnStopProcessing = True
    End If
               
    curFilePos = curFilePos + CCur(UBound(abytData) + 1)  ' Adjust pointers accordingly
    curAmtLeft = curAmtLeft - CCur(UBound(abytData) + 1)  ' calc how much is left
    
End Sub

' ***************************************************************************
' Routine:       BuildContinuous
'
' Description:   This routine will build a file with one contiguous record.
'                Sometimes referred to a variable length record.
'
' Parameters:    intLow - Lowest byte value requested
'                intHigh - Highest byte value requested
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-DEC-1999  Kenneth Ives  kenaso@tx.rr.com
'              Created routine
' 21-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Function BuildContinuous(ByVal intLow As Integer, _
                                 ByVal intHigh As Integer) As Boolean

    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim hFile        As Long
    Dim lngIdx       As Long
    Dim lngIndex     As Long
    Dim alngRand()   As Long
    Dim abytRand()   As Byte
    Dim abytData()   As Byte
    
    Const ROUTINE_NAME As String = "BuildContinuous"
        
    On Error GoTo BuildContinuous_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
      
    Set mobjBigFiles = New cBigFiles
    
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
    
    mobjPrng.StopProcessing = gblnStopProcessing
    
    hFile = FreeFile                        ' capture first free file handle
    Open mstrFilename For Output As #hFile  ' Create an empty file
    Close #hFile                            ' close the file

    ' Open a new file
    If Not mobjBigFiles.OpenReadWrite(mstrFilename, hFile) Then
        gblnStopProcessing = True
    End If
    
    curAmtLeft = mcurNumberOfBytes
    curFilePos = 0@
    
    If mlngAlgorithm = 0 Then
        
        Do While curAmtLeft > 0
            
            If gblnStopProcessing Then
                DoEvents
                mobjPrng.StopProcessing = True
                DoEvents
                Exit Do   ' Exit Do..Loop
            End If
        
            Erase abytData()   ' Start with an empty array
            abytData() = mobjPrng.BuildWithinRange(KB_64, intLow, intHigh, , False)
        
            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit Do    ' exit Do..Loop
            End If
        
            If KB_64 > curAmtLeft Then
                ReDim Preserve abytData(curAmtLeft)
            End If
            
            UpdateContinuousFile abytData(), curAmtLeft, curFilePos, hFile   ' Update output file
            
            UpdateDisplay curFilePos   ' Update progress bar display
                
            If curAmtLeft <= 0@ Then
                Exit Do    ' exit Do..Loop
            End If
        
        Loop
        
    Else
        
        Do While curAmtLeft > 0@
        
            If gblnStopProcessing Then
                DoEvents
                mobjPrng.StopProcessing = True
                DoEvents
                Exit Do   ' Exit Do..Loop
            End If
        
            lngIdx = 0
            ReDim abytData(KB_64)
            
            Do While lngIdx < KB_64
                
                Erase alngRand()   ' Start with an empty array
                    
                ' Fill byte array with random values
                Select Case mlngAlgorithm
                       Case 1: alngRand() = mobjPrng.ISAAC_Prng(KB_16, False)
                       Case 2: alngRand() = mobjPrng.KISS_Prng(KB_16, False)
                       Case 3: alngRand() = mobjPrng.MWC_Prng(KB_16, False)
                       Case 4: alngRand() = mobjPrng.MOA_Prng(KB_16, False)
                       Case 5: alngRand() = mobjPrng.MT_Prng(KB_16, False)
                       Case 6: alngRand() = mobjPrng.MTA_Prng(KB_16, False)
                       Case 7: alngRand() = mobjPrng.MTB_Prng(KB_16, False)
                       Case 8: alngRand() = mobjPrng.TT800_Prng(KB_16, False)
                End Select
                
                ' An error occurred or user opted to STOP processing
                If gblnStopProcessing Then
                    Exit Do    ' exit Do..Loop
                End If
        
                ReDim abytRand(KB_64)  ' Size byte array
                
                ' Convert long array to byte array
                CopyMemory abytRand(0), alngRand(0), KB_64
            
                For lngIndex = 0 To UBound(abytRand) - 1
                                                
                    Select Case abytRand(lngIndex)
                           Case intLow To intHigh
                                abytData(lngIdx) = abytRand(lngIndex)
                                lngIdx = lngIdx + 1
                    End Select
                        
                    ' If enough data collected then exit loop
                    If lngIdx = KB_64 Then
                        Exit For    ' exit For..Next loop
                    End If
                                                                            
                Next lngIndex
        
                ' An error occurred or user opted to STOP processing
                If gblnStopProcessing Then
                    Exit Do    ' exit Do..Loop
                End If
        
                ' If enough data collected then exit loop
                If lngIdx = KB_64 Then
                    Exit Do    ' exit Do..Loop
                End If
                                                                            
            Loop
        
            If KB_64 > curAmtLeft Then
                ReDim Preserve abytData(curAmtLeft)
            End If
            
            UpdateContinuousFile abytData(), curAmtLeft, curFilePos, hFile   ' Update output file
            
            UpdateDisplay curFilePos   ' Update progress bar display
                
            If curAmtLeft <= 0@ Then
                Exit Do    ' exit Do..Loop
            End If
        
        Loop
    
    End If
    
BuildContinuous_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        mobjPrng.StopProcessing = gblnStopProcessing
        DoEvents
        BuildContinuous = False
    Else
        BuildContinuous = True
    End If
    
    mobjBigFiles.API_CloseFile hFile   ' Close open file
    
    CloseAllFiles            ' Verify all opened files have been closed
    Set mobjPrng = Nothing   ' Free objects from memory when not needed
    Set mobjBigFiles = Nothing
    
    Erase abytData()   ' Always empty arrays when not needed
    Erase abytRand()
    Erase alngRand()
       
    On Error GoTo 0
    Exit Function
    
BuildContinuous_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume BuildContinuous_CleanUp
  
End Function

Private Sub UpdateContinuousFile(ByRef abytData() As Byte, _
                                 ByRef curAmtLeft As Currency, _
                                 ByRef curFilePos As Currency, _
                                 ByVal hFile As Long)
    
    ' Called by BuildContinuous()
    
    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngPointer As Long
    Dim strRecord  As String
    
    lngLength = UBound(abytData)       ' Capture data length
    lngPointer = 1                     ' set to first positionin output string
    strRecord = Space$(lngLength * 2)  ' preload output string with spaces
    
    ' build the output record
    If mblnHexConvert Then
    
        ' Loop thru the byte array and build the output record
        For lngIndex = 0 To lngLength - 1
          
            ' load with the hex values
            Mid$(strRecord, lngPointer, 2) = Right$("00" & Hex$(abytData(lngIndex)), 2)
            lngPointer = lngPointer + 2
        
        Next lngIndex
        
    Else
        strRecord = ByteArrayToString(abytData())
    End If
    
    Erase abytData()
    strRecord = UCase$(Trim$(strRecord))
    strRecord = Left$(strRecord, lngLength)
    abytData() = StringToByteArray(strRecord)
    
    ' Write to target file
    If Not mobjBigFiles.API_WriteFile(hFile, curFilePos, abytData()) Then
        gblnStopProcessing = True
    End If
               
    curFilePos = curFilePos + CCur(UBound(abytData) + 1)  ' Adjust pointers accordingly
    curAmtLeft = curAmtLeft - CCur(UBound(abytData) + 1)  ' calc how much is left
    
End Sub

Private Sub UpdateDisplay(ByVal curByteCount As Currency)
    
    Dim lngPercent As Long
          
    lblProgress(1).Caption = Format$(curByteCount, "#,##0") & " / " & _
                             Format$(mcurNumberOfBytes, "#,##0") & " bytes"
                             
    lngPercent = CalcProgress(curByteCount, mcurNumberOfBytes)
    
    ProgressBar picProgressBar, lngPercent, vbBlue
    
End Sub

Private Sub LockDownCtls()
    
    With frmMain
        .cmdChoice(0).Visible = False
        .cmdChoice(0).Enabled = False
        .cmdChoice(1).Enabled = True
        .cmdChoice(1).Visible = True
        .cmdChoice(2).Enabled = False
        .lblProgress(1).Caption = "0 / " & Format$(mcurNumberOfBytes, "#,##0") & " bytes"
        .fraMain.Enabled = False
    End With
    
End Sub

Private Sub UnlockCtls()
    
    With frmMain
        .fraMain.Enabled = True
        .cmdChoice(0).Enabled = True
        .cmdChoice(0).Visible = True
        .cmdChoice(1).Visible = False
        .cmdChoice(1).Enabled = False
        .cmdChoice(2).Enabled = True
        .lblProgress(1).Caption = "0 / " & Format$(mcurNumberOfBytes, "#,##0") & " bytes"
    End With
    
End Sub

Private Sub ChoosePrefix()
    mstrPrefix = Choose(mlngAlgorithm + 1, "MS", "IC", "KISS", "MWC", "MOA", "MT", "MTA", "MTB", "TT8")
End Sub

Private Sub ResetProgressMsg()
    
    With frmMain
        .lblProgress(0).Caption = PROGRESS_MSG
        .lblProgress(1).Caption = vbNullString
    End With
    
    ResetProgressBar
    
End Sub

Private Sub ResetProgressBar()

    ' Resets progressbar to zero
    ' with all white background
    ProgressBar picProgressBar, 0, vbWhite
    
End Sub

' ***************************************************************************
' Routine:       ProgessBar
'
' Description:   Fill a picturebox as if it were a horizontal progress bar.
'
' Parameters:    objProgBar - name of picture box control
'                lngPercent - Current percentage value
'                lngForeColor - Optional-The progression color. Default = Black.
'                           can use standard VB colors or long Integer
'                           values representing a color.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 14-FEB-2005  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
' 05-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' ***************************************************************************
Private Sub ProgressBar(ByRef objProgBar As PictureBox, _
                        ByVal lngPercent As Long, _
               Optional ByVal lngForeColor As Long = vbBlue)

    Dim strPercent As String
    
    Const MAX_PERCENT As Long = 100
    
    ' Called by ResetProgressBar() routine
    ' to reinitialize progress bar properties.
    ' If forecolor is white then progressbar
    ' is being reset to a starting position.
    If lngForeColor = vbWhite Then
        
        With objProgBar
            .AutoRedraw = True      ' Required to prevent flicker
            .BackColor = &HFFFFFF   ' White
            .DrawMode = 10          ' Not Xor Pen
            .FillStyle = 0          ' Solid fill
            .FontName = "Arial"     ' Name of font
            .FontSize = 11          ' Font point size
            .FontBold = True        ' Font is bold.  Easier to see.
            Exit Sub                ' Exit this routine
        End With
    
    End If
        
    ' If no progress then leave
    If lngPercent < 1 Then
        Exit Sub
    End If
    
    ' Verify flood display has not exceeded 100%
    If lngPercent <= MAX_PERCENT Then

        With objProgBar
        
            ' Error trap in case code attempts to set
            ' scalewidth greater than the max allowable
            If lngPercent > .ScaleWidth Then
                lngPercent = .ScaleWidth
            End If
               
            .Cls                        ' Empty picture box
            .ForeColor = lngForeColor   ' Reset forecolor
         
            ' set picture box ScaleWidth equal to maximum percentage
            .ScaleWidth = MAX_PERCENT
            
            ' format percent into a displayable value (ex: 25%)
            strPercent = Format$(CLng((lngPercent / .ScaleWidth) * 100)) & "%"
            
            ' Calculate X and Y coordinates within
            ' picture box and and center data
            .CurrentX = (.ScaleWidth - .TextWidth(strPercent)) \ 2
            .CurrentY = (.ScaleHeight - .TextHeight(strPercent)) \ 2
                
            objProgBar.Print strPercent   ' print percentage string in picture box
            
            ' Print flood bar up to new percent position in picture box
            objProgBar.Line (0, 0)-(lngPercent, .ScaleHeight), .ForeColor, BF
        
        End With
        DoEvents   ' allow flood to complete drawing
    End If

End Sub

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the record to be written.  The
'                write process has been speeded up by 50% or more by
'                adjusting the record length based on amount of data left
'                to write. Since VB6 normally only allocates 64kb of memory,
'                I have limited the maximum record size to 60kb so the
'                application will not have to perform as much swapping.
'
' Parameters:    curAmtLeft - Amount of data left to be written
'
' Returns:       New record size as a long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-Jul-2007  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Determine the record size to write.
    Select Case curAmtLeft
           Case Is >= KB_4: GetBlockSize = KB_4
           Case Else:       GetBlockSize = CLng(curAmtLeft)
    End Select
    
End Function

