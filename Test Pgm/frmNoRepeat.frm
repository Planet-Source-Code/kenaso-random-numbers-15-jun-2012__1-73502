VERSION 5.00
Begin VB.Form frmNoRepeat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rnd Testing - Non-Repeating Numbers"
   ClientHeight    =   4785
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7170
   Begin VB.Frame Frame1 
      Caption         =   "Test selection (10 loops)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   150
      TabIndex        =   8
      Top             =   2850
      Width           =   3690
      Begin VB.OptionButton optTest 
         Caption         =   "16 out of 16, lowest = 1, step 1, unsorted"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   9
         Top             =   480
         Width           =   3390
      End
      Begin VB.OptionButton optTest 
         Caption         =   "10 out of 95, lowest = 25, step 5, sorted"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   225
         TabIndex        =   4
         Top             =   1500
         Width           =   3390
      End
      Begin VB.OptionButton optTest 
         Caption         =   "4 out of 40, lowest = 10, step 1, sorted"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   225
         TabIndex        =   3
         Top             =   1245
         Width           =   3390
      End
      Begin VB.OptionButton optTest 
         Caption         =   "2 out of 6, lowest = 1, step 1, unsorted"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   225
         TabIndex        =   2
         Top             =   990
         Width           =   3390
      End
      Begin VB.OptionButton optTest 
         Caption         =   "16 out of 16, lowest = 0, step 1, unsorted"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   1
         Top             =   735
         Width           =   3390
      End
      Begin VB.OptionButton optTest 
         Caption         =   "6 out of 54, lowest = 1, step 1, sorted"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   225
         Value           =   -1  'True
         Width           =   3390
      End
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   4425
      TabIndex        =   6
      Top             =   4350
      Width           =   990
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   4425
      TabIndex        =   5
      Top             =   3900
      Width           =   990
   End
   Begin VB.TextBox txtNumbers 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   225
      Width           =   6900
   End
End
Attribute VB_Name = "frmNoRepeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private mintTest As Integer
  Private mobjPrng As Object
  
Private Sub cmdChoice_Click(Index As Integer)

    Dim alngResults() As Long  ' Returning list of numbers
    Dim lngIndex1     As Long
    Dim lngIndex2     As Long
    Dim strTemp       As String
    
    txtNumbers = vbNullString
    strTemp = vbNullString
    DoEvents
    
    ' Determine which button was pressed
    Select Case Index
           Case 0   ' Start the test
                Screen.MousePointer = vbHourglass
                cmdChoice(0).Enabled = False
                
                Select Case mintTest
                       
                       Case 0  ' Return 6 numbers, lowest value = 1, Highest value = 54, increment by 1, sorted
                               ' Texas Lotto example
                            For lngIndex1 = 1 To 10
                                DoEvents
                                alngResults() = mobjPrng.NonRepeatingNbrs(6, 1, 54, , True)
                                
                                If mobjPrng.IsArrayInitialized(alngResults()) Then
                                    For lngIndex2 = 0 To UBound(alngResults) - 1
                                        strTemp = strTemp & Format$(alngResults(lngIndex2), "@@@@")
                                    Next lngIndex2
                                    
                                    strTemp = strTemp & vbNewLine
                                Else
                                    Exit For
                                End If
                            Next lngIndex1
                            
                       Case 1  ' Return 16 numbers, lowest value = 1, Highest value = 16, increment by 1, unsorted
                            For lngIndex1 = 1 To 10
                                DoEvents
                                alngResults() = mobjPrng.NonRepeatingNbrs(16, 1, 16, , False)
                                
                                If mobjPrng.IsArrayInitialized(alngResults()) Then
                                    For lngIndex2 = 0 To UBound(alngResults) - 1
                                        strTemp = strTemp & Format$(alngResults(lngIndex2), "@@@")
                                    Next lngIndex2
                                    
                                    strTemp = strTemp & vbNewLine
                                Else
                                    Exit For
                                End If
                            Next lngIndex1
                       
                       Case 2  ' Return 16 numbers, lowest value = 0, Highest value = 15, increment by 1, unsorted
                            For lngIndex1 = 1 To 10
                                DoEvents
                                alngResults() = mobjPrng.NonRepeatingNbrs(16, 0, 15, , False)

                                If mobjPrng.IsArrayInitialized(alngResults()) Then
                                    For lngIndex2 = 0 To UBound(alngResults) - 1
                                        strTemp = strTemp & Format$(alngResults(lngIndex2), "@@@")
                                    Next lngIndex2
    
                                    strTemp = strTemp & vbNewLine
                                Else
                                    Exit For
                                End If
                            Next lngIndex1
                       
                       Case 3  ' Return 2 numbers, lowest value = 1, Highest value = 6, increment by 1, unsorted
                            For lngIndex1 = 1 To 10
                                DoEvents
                                alngResults() = mobjPrng.NonRepeatingNbrs(2, 1, 6, , False)
                                
                                If mobjPrng.IsArrayInitialized(alngResults()) Then
                                    For lngIndex2 = 0 To UBound(alngResults) - 1
                                        strTemp = strTemp & Format$(alngResults(lngIndex2), "@@@@")
                                    Next lngIndex2
                                    
                                    strTemp = strTemp & vbNewLine
                                Else
                                    Exit For
                                End If
                            Next lngIndex1
                       
                       Case 4  ' Return 4 numbers, lowest value = 10, Highest value = 40, increment by 1, sorted
                            For lngIndex1 = 1 To 10
                                DoEvents
                                alngResults() = mobjPrng.NonRepeatingNbrs(4, 10, 40, 1, True)
                                
                                If mobjPrng.IsArrayInitialized(alngResults()) Then
                                    For lngIndex2 = 0 To UBound(alngResults) - 1
                                        strTemp = strTemp & Format$(alngResults(lngIndex2), "@@@@")
                                    Next lngIndex2
                                    
                                    strTemp = strTemp & vbNewLine
                                Else
                                    Exit For
                                End If
                            Next lngIndex1
                       
                       Case 5  ' Return 10 numbers, lowest value = 25, Highest value = 95, increment by 5, sorted
                            For lngIndex1 = 1 To 10
                                DoEvents
                                alngResults() = mobjPrng.NonRepeatingNbrs(10, 25, 95, 5, True)
                                
                                If mobjPrng.IsArrayInitialized(alngResults()) Then
                                    For lngIndex2 = 0 To UBound(alngResults) - 1
                                        strTemp = strTemp & Format$(alngResults(lngIndex2), "@@@@")
                                    Next lngIndex2
                                    
                                    strTemp = strTemp & vbNewLine
                                Else
                                    Exit For
                                End If
                            Next lngIndex1
                End Select
                
                txtNumbers.Text = strTemp
                Screen.MousePointer = vbDefault
                cmdChoice(0).Enabled = True
                
           Case 1  ' terminate application
                Unload Me
    End Select
    
End Sub

Private Sub Form_Load()

    Set mobjPrng = New kiPrng.cPrng
    optTest_Click 0
    
    With frmNoRepeat
        ' center on the screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless               ' display the screen with no flicker
        .Refresh
    End With
    
    DoEvents
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
    Close                        ' close any open files
    Set mobjPrng = Nothing       ' free object from memory
    Unload frmNoRepeat           ' deactivate this form object
    Set frmNoRepeat = Nothing    ' free form object from memory
    
End Sub

Private Sub optTest_Click(Index As Integer)

    Dim intIndex As Integer
    
    For intIndex = 0 To 5
        If intIndex = Index Then
            optTest(intIndex).Value = True
            mintTest = intIndex
        Else
            optTest(intIndex).Value = False
        End If
    Next
  
End Sub
