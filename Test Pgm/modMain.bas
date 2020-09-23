Attribute VB_Name = "modMain"
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MAXLONG   As Long = &H7FFFFFFF  ' 2147483647
  Private Const MAX_LIMIT As Long = 10000
  Private Const SAMPLE    As Long = 100
  Private Const KB_32     As Long = 32768
  Private Const KB_64     As Long = 65536

' ***************************************************************************
' Module API Declares
' ***************************************************************************
  ' This is a rough translation of the GetTickCount API. The
  ' tick count of a PC is only valid for the first 49.7 days
  ' since the last reboot.  When you capture the tick count,
  ' you are capturing the total number of milliseconds elapsed
  ' since the last reboot.  The elapsed time is stored as a
  ' DWORD value. Therefore, the time will wrap around to zero
  ' if the system is run continuously for 49.7 days.
  Private Declare Function GetTickCount Lib "kernel32" () As Long
  
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long
  
' ***************************************************************************
' Module Variables
' ***************************************************************************
  Private hFile         As Long
  Private mlngAlgorithm As Long
  Private mstrFile      As String
  Private mstrGenerator As String
  Private mobjPrng      As Object
  
  

' ****************************************************************************
' Rename this routine to "Main" and then press F5 to execute.
'
' Testing software:
' Diehard by George Marsaglia
' http://stat.fsu.edu/pub/diehard/
'
' Ent Software
' http://www.fourmilab.ch/random/
' Scroll down and download the file Random.zip
'
' Compiled (binaries) NIST Statisical Test Suite v1.6
' http://www.cs.sunysb.edu/~algorith/implement/rng/distrib/sts-1.6.zip
'
' If anyone knows how to compile the newest version NIST Statistical
' Test Suite 2.1.1 using .NET C#, please upload to the net and let others
' know.  After all, most programmers reading this are Visual Basic 6.0
' users only.  Thank you
' http://csrc.nist.gov/groups/ST/toolkit/rng/documentation_software.html
'
' GeneratorTypes:
'    0 - CryptoAPI   3 - MWC (Multiply With Carry)   6 - MT11232A
'    1 - ISAAC       4 - MOA (Mother-of-All)         7 - MT11232B
'    2 - KISS        5 - MT19937                     8 - TT800
' ****************************************************************************
Public Sub Main()

    Screen.MousePointer = vbHourglass
    DoEvents
    
    Set mobjPrng = Nothing
    
    For mlngAlgorithm = 0 To 8
             
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
        
        Test1   ' Float values
        Test2   ' Long Integer values
        Test3   ' Avg of 10 iterations of float (Test1)
        Test4   ' Avg of single dice throwing
        Test5   ' Create 11mb binary test file (Diehard, ENT, NIST testing)
        
        Set mobjPrng = Nothing
    
    Next mlngAlgorithm
    
    Screen.MousePointer = vbDefault
    DoEvents
    
    MsgBox "Finished"
    
End Sub

' ****************************************************************************
' Test1:  RANDOM NUMBER GENERATOR (Test actual output)
'         Called by "Main"
'
' Only 1 in every 100 generated values will be written to the file.
'
' Sample Output:
'
'    Average:  0.001801584444
'     Lowest: -0.999965964816
'    Highest:  0.999990548939
'
'     0.802474411578  -0.167852585282  ...  -0.140644020466   0.469271864495
'    -0.501796334750   0.238365060295  ...   0.360001724381   0.461231733887
'     0.909769143402  -0.843400150465  ...  -0.376032714441   0.691831045175
' ****************************************************************************
Private Sub Test1()
    
    Dim hFile       As Long
    Dim lngIndex    As Long
    Dim lngLoop     As Long
    Dim lngColumn   As Long
    Dim lngCount    As Long
    Dim adblData()  As Double
    Dim adblTemp()  As Double
    Dim dblTotal    As Double
    Dim dblLow      As Double
    Dim dblHigh     As Double
    Dim strFmt      As String
    Dim strTemp     As String
    Dim strTitle    As String

    Dim lngStart    As Long
    Dim strElapsed  As String
    
    Const FILE_NAME As String = "_Test01.txt"

    Erase adblData()
    Erase adblTemp()
    ReDim adblTemp(MAX_LIMIT)
    
    lngColumn = 0
    lngCount = 0
    dblLow = 1#
    dblHigh = 0#
    dblTotal = 0#
    strFmt = String$(17, "@")

    PrepNames FILE_NAME
    strTitle = "Test 1:" & vbNewLine & Format$(MAX_LIMIT * 10, "#,0") & _
               " random generated numbers" & vbNewLine & _
               "using the " & mstrGenerator & " generator." & vbNewLine & _
               "Displaying every " & CStr(SAMPLE) & "th generated value."
               
    hFile = FreeFile                     ' get first free file handle
    Open mstrFile For Output As #hFile   ' create empty output file

    Print #hFile, strTitle               ' print the title
    Print #hFile, " "
    Print #hFile, "Version: " & mobjPrng.Version ' Save DLL version information
    Print #hFile, " "

    Erase adblData()           ' Empty array
    lngStart = GetTickCount()  ' starting time
    
    For lngLoop = 1 To 10
        
        ' generate random float values
        Select Case mlngAlgorithm
               Case 0: adblData() = mobjPrng.BuildRndData(MAX_LIMIT, ePRNG_DBL_ARRAY, False)
               Case 1: adblData() = mobjPrng.ISAAC_Prng(MAX_LIMIT, True)
               Case 2: adblData() = mobjPrng.KISS_Prng(MAX_LIMIT, True)
               Case 3: adblData() = mobjPrng.MWC_Prng(MAX_LIMIT, True)
               Case 4: adblData() = mobjPrng.MOA_Prng(MAX_LIMIT, True)
               Case 5: adblData() = mobjPrng.MT_Prng(MAX_LIMIT, True)
               Case 6: adblData() = mobjPrng.MTA_Prng(MAX_LIMIT, True)
               Case 7: adblData() = mobjPrng.MTB_Prng(MAX_LIMIT, True)
               Case 8: adblData() = mobjPrng.TT800_Prng(MAX_LIMIT, True)
        End Select
        
        ' Gather the output statistics
        For lngIndex = 0 To MAX_LIMIT - 1
    
            dblTotal = dblTotal + adblData(lngIndex)  ' Accumulate overall total
    
            If dblLow > adblData(lngIndex) Then
                dblLow = adblData(lngIndex)           ' check for lowest value
            ElseIf adblData(lngIndex) > dblHigh Then
                dblHigh = adblData(lngIndex)          ' check for highest value
            End If
    
            ' Just use every 100th value for the output file
            If lngIndex Mod SAMPLE = 0 Then
                adblTemp(lngCount) = adblData(lngIndex)
                lngCount = lngCount + 1
            End If

        Next lngIndex
    Next lngLoop
    
    ReDim Preserve adblTemp(lngCount)
    strElapsed = ElapsedTime(GetTickCount() - lngStart) ' finish time
    Print #hFile, "Elapsed:  " & strElapsed & vbNewLine
    
    ' format and write the statistics to the file
    Print #hFile, "Average: " & Format$(FormatNumber(dblTotal / CDbl(MAX_LIMIT), 14), strFmt)
    Print #hFile, " Lowest: " & Format$(FormatNumber(dblLow, 14), strFmt)
    Print #hFile, "Highest: " & Format$(FormatNumber(dblHigh, 14), strFmt)
    Print #hFile, " "

    ' dump the contents of the array to the output file
    For lngIndex = 0 To UBound(adblTemp) - 1

        ' write to the test file.
        strTemp = Format$(FormatNumber(adblTemp(lngIndex), 14), strFmt)
        Print #hFile, strTemp & Space$(2);   ' write to the file
        lngColumn = lngColumn + 1            ' increment column counter
        
        ' see if we have 5 columns
        If lngColumn = 5 Then
           Print #hFile, ""    ' prints Chr$(13) + Chr$(10) at the end of the line
           lngColumn = 0       ' reset column counter
        End If

    Next lngIndex

    Close #hFile
    Erase adblData()
    Erase adblTemp()

End Sub

' ****************************************************************************
' Test2:  RANDOM NUMBER GENERATOR (Test actual output)
'         Called by "Main"
'
' Use this type scenario to capture an array of long integers
'
' Sample output:
'
'    Average:    -4060298
'     Lowest: -2147234620
'    Highest:  2147451487
'
'     2017985050  -1717396184   ...   1101853862   1721100454    -40232883
'     -652127292  -1047167738   ...   -533075425   -902048183   1740833607
'     1573020271   1987555641   ...   -438649914   -845848487   2041929893
' ****************************************************************************
Private Sub Test2()

    Dim strFmt      As String
    Dim strTemp     As String
    Dim strTitle    As String
    Dim dblTotal    As Double
    Dim dblLow      As Double
    Dim dblHigh     As Double
    Dim hFile       As Long
    Dim lngIndex    As Long
    Dim lngLoop     As Long
    Dim lngColumn   As Long
    Dim lngCount    As Long
    Dim alngData()  As Long
    Dim alngTemp()  As Long

    Dim lngStart    As Long
    Dim strElapsed  As String

    Const FILE_NAME As String = "_Test02.txt"

    Erase alngData()
    Erase alngTemp()
    ReDim alngTemp(MAX_LIMIT)
    
    lngColumn = 0
    lngCount = 0
    dblLow = MAXLONG
    dblHigh = 0#
    dblTotal = 0#
    strFmt = String$(11, "@")

    PrepNames FILE_NAME
    strTitle = "Test 2:" & vbNewLine & Format$(MAX_LIMIT * 10, "#,0") & _
               " random generated numbers" & vbNewLine & _
               "using the " & mstrGenerator & " generator." & vbNewLine & _
               "Displaying every " & CStr(SAMPLE) & "th generated value."
               
    hFile = FreeFile                     ' get first free file handle
    Open mstrFile For Output As #hFile   ' create empty output file

    Print #hFile, strTitle               ' print the title
    Print #hFile, " "
    Print #hFile, "Version: " & mobjPrng.Version  ' Save DLL version information
    Print #hFile, " "

    Erase alngData()           ' Empty array
    lngStart = GetTickCount()  ' starting time

    For lngLoop = 1 To 10
    
        ' generate random float values
        Select Case mlngAlgorithm
               Case 0: alngData() = mobjPrng.BuildRndData(MAX_LIMIT, ePRNG_LONG_ARRAY, False)
               Case 1: alngData() = mobjPrng.ISAAC_Prng(MAX_LIMIT, False)
               Case 2: alngData() = mobjPrng.KISS_Prng(MAX_LIMIT, False)
               Case 3: alngData() = mobjPrng.MWC_Prng(MAX_LIMIT, False)
               Case 4: alngData() = mobjPrng.MOA_Prng(MAX_LIMIT, False)
               Case 5: alngData() = mobjPrng.MT_Prng(MAX_LIMIT, False)
               Case 6: alngData() = mobjPrng.MTA_Prng(MAX_LIMIT, False)
               Case 7: alngData() = mobjPrng.MTB_Prng(MAX_LIMIT, False)
               Case 8: alngData() = mobjPrng.TT800_Prng(MAX_LIMIT, False)
        End Select
     
        ' Gather the output statistics
        For lngIndex = 0 To MAX_LIMIT - 1
    
            dblTotal = dblTotal + alngData(lngIndex)  ' Accumulate overall total
    
            If dblLow > alngData(lngIndex) Then
                dblLow = alngData(lngIndex)           ' check for lowest value
            ElseIf alngData(lngIndex) > dblHigh Then
                dblHigh = alngData(lngIndex)          ' check for highest value
            End If
    
            ' Just use every 100th value for the output file
            If lngIndex Mod SAMPLE = 0 Then
                alngTemp(lngCount) = alngData(lngIndex)
                lngCount = lngCount + 1
            End If

        Next lngIndex
    Next lngLoop
    
    ReDim Preserve alngTemp(lngCount)
    strElapsed = ElapsedTime(GetTickCount() - lngStart) ' finish time
    Print #hFile, "Elapsed:  " & strElapsed & vbNewLine
    
    ' format and write the statistics to the file
    Print #hFile, "Average: " & Format$(Int(dblTotal / CDbl(MAX_LIMIT)), strFmt)
    Print #hFile, " Lowest: " & Format$(dblLow, strFmt)
    Print #hFile, "Highest: " & Format$(dblHigh, strFmt)
    Print #hFile, " "

    ' dump the contents of the array to the output file
    For lngIndex = 0 To UBound(alngTemp) - 1

        
        strTemp = Format$(alngTemp(lngIndex), strFmt)  ' format the output
        Print #hFile, strTemp & Space$(2);             ' write to the test file
        lngColumn = lngColumn + 1                      ' increment column counter
        
        ' see if we have 7 columns
        If lngColumn = 7 Then
           Print #hFile, ""    ' prints Chr$(13) + Chr$(10) at the end of the line
           lngColumn = 0       ' reset column counter
        End If

    Next lngIndex

    Close #hFile
    Erase alngData()
    Erase alngTemp()

End Sub

' ****************************************************************************
' Test3:  RANDOM NUMBER GENERATOR (Check median average of 10,000 values)
'         Called by "Main".  See the difference in time when producing
'         normal strong values versus cryptographic values.
'
' Sample Output:
'
' Median average of 100,000 values (10,000 * 10 iterations)
' using MWC (Multiply With Carry) generator.
'
' Version: kiPrng.dll (tm) v1.6.86 Copyright (c) 2004-2012  Kenneth Ives  kenaso@tx.rr.com
'
'                        *** Strong values ***
'            Low Value          High Value          Median Avg        Time
'  1.  -0.99998023686830    0.99997558537770    0.00102664503395   00:00:00.328
'  2.  -0.99997146380965    0.99998408229921    0.00269017381215   00:00:00.312
'  3.  -0.99998659361054    0.99994413414869    0.00127735815358   00:00:00.328
'  4.  -0.99993940163302    0.99999187467520   -0.00085869086409   00:00:00.329
'  5.  -0.99999650288273    0.99998987605695   -0.00109222915400   00:00:00.343
'  6.  -0.99999852431838    0.99999869475041    0.00066345509596   00:00:00.328
'  7.  -0.99996637878839    0.99999093962733   -0.00241727640134   00:00:00.329
'  8.  -0.99997403286297    0.99998534796659    0.00166836523529   00:00:00.328
'  9.  -0.99999211123113    0.99998557381232    0.00044845501046   00:00:00.328
' 10.  -0.99999328562890    0.99997309083018    0.00135367001980   00:00:00.312
'
'                        *** Cryptographic values ***
'            Low Value          High Value          Median Avg        Time
'  1.  -0.99995659571040    0.99998358357597   -0.00119131353324   00:00:04.906
'  2.  -0.99999067140643    0.99998402874816   -0.00103042013194   00:00:04.922
'  3.  -0.99994818027961    0.99998691864211    0.00357002856765   00:00:04.922
'  4.  -0.99999417923291    0.99993704212728   -0.00249493294190   00:00:04.937
'  5.  -0.99997988855366    0.99997840030018    0.00029878583278   00:00:04.922
'  6.  -0.99999529868264    0.99999914551054   -0.00147286364780   00:00:06.531
'  7.  -0.99998382013191    0.99999797856335    0.00093716583817   00:00:04.938
'  8.  -0.99997112015162    0.99995998525891   -0.00119451336710   00:00:04.922
'  9.  -0.99999710125749    0.99998817732457   -0.00324351706127   00:00:04.921
' 10.  -0.99998654564742    0.99998080404375    0.00023199582987   00:00:04.922
' ****************************************************************************
Private Sub Test3()

    Dim strFmt      As String
    Dim strTemp     As String
    Dim strTitle    As String
    Dim strMsg      As String
    Dim lngLoop     As Long
    Dim lngIndex    As Long
    Dim lngIdx      As Long
    Dim dblTotal    As Double
    Dim dblLow      As Double
    Dim dblHigh     As Double
    Dim adblData()  As Double

    Dim lngStart    As Long
    Dim strElapsed  As String

    Const FILE_NAME As String = "_Test03.txt"

    Erase adblData()
    strFmt = String$(17, "@")
    PrepNames FILE_NAME
    strTitle = "Test 3:" & vbNewLine & "Median average of " & Format$(MAX_LIMIT * 10, "#,0") & _
               " values (" & Format$(MAX_LIMIT, "#,0") & " * 10 iterations)" & vbNewLine & _
               "using " & mstrGenerator & " generator."

    hFile = FreeFile                     ' get first free file handle
    Open mstrFile For Output As #hFile   ' create empty output file

    Print #hFile, strTitle               ' print the title
    Print #hFile, " "
    Print #hFile, "Version: " & mobjPrng.Version  ' Save DLL version information
    Print #hFile, " "

    Print #hFile, Space$(11) & "Low Value" & Space$(10) & _
                  "High Value" & Space$(10) & "Median Avg" & Space$(8) & "Time"

    For lngLoop = 1 To 10

        dblLow = 1#
        dblHigh = 0#
        dblTotal = 0#

        lngStart = GetTickCount()   ' starting time
        
        For lngIdx = 1 To 10
        
            ' generate random float values
            Select Case mlngAlgorithm
                   Case 0: adblData() = mobjPrng.BuildRndData(MAX_LIMIT, ePRNG_DBL_ARRAY, False)
                   Case 1: adblData() = mobjPrng.ISAAC_Prng(MAX_LIMIT, True)
                   Case 2: adblData() = mobjPrng.KISS_Prng(MAX_LIMIT, True)
                   Case 3: adblData() = mobjPrng.MWC_Prng(MAX_LIMIT, True)
                   Case 4: adblData() = mobjPrng.MOA_Prng(MAX_LIMIT, True)
                   Case 5: adblData() = mobjPrng.MT_Prng(MAX_LIMIT, True)
                   Case 6: adblData() = mobjPrng.MTA_Prng(MAX_LIMIT, True)
                   Case 7: adblData() = mobjPrng.MTB_Prng(MAX_LIMIT, True)
                   Case 8: adblData() = mobjPrng.TT800_Prng(MAX_LIMIT, True)
            End Select
        
            ' Gather the output statistics
            For lngIndex = 0 To MAX_LIMIT - 1
        
                dblTotal = dblTotal + adblData(lngIndex)  ' Accumulate overall total
        
                If dblLow > adblData(lngIndex) Then
                    dblLow = adblData(lngIndex)           ' check for lowest value
                ElseIf adblData(lngIndex) > dblHigh Then
                    dblHigh = adblData(lngIndex)          ' check for highest value
                End If
        
            Next lngIndex
        Next lngIdx
        
        strElapsed = ElapsedTime(GetTickCount() - lngStart) ' finish time
        
        ' format and write the stats to the file
        strMsg = Format$(lngLoop, "@@") & ".  " & _
                 Format$(FormatNumber(dblLow, 14), strFmt) & Space$(3) & _
                 Format$(FormatNumber(dblHigh, 14), strFmt) & Space$(3) & _
                 Format$(FormatNumber(dblTotal / CDbl(MAX_LIMIT * 10), 14), strFmt) & Space$(3) & _
                 strElapsed
        
        Print #hFile, strMsg

    Next lngLoop

    Close #hFile
    Erase adblData()

End Sub

' ****************************************************************************
' Test4:  RANDOM NUMBER GENERATOR
'         Called by "Main"
'
' Testing the throw of a single dice for 100,000 times.
' Test results should be close to each other like this:
'
'    Throw a single dice 100,000 times
'    using Microsoft CryptoAPI generator.
'
'    0.  0        ' <- Should always be zero
'    1.  16339
'    2.  16817
'    3.  16777
'    4.  16728
'    5.  16656
'    6.  16684
'    Other.  0    ' <- Should always be zero
'
' ****************************************************************************
Private Sub Test4()

    Dim strFmt      As String
    Dim strTemp     As String
    Dim strTitle    As String
    Dim lngCount    As Long
    Dim lngRange    As Long
    Dim lngValue    As Long
    Dim lngIndex    As Long
    Dim lngMaxAmt   As Long
    Dim lng_0       As Long
    Dim lng_1       As Long
    Dim lng_2       As Long
    Dim lng_3       As Long
    Dim lng_4       As Long
    Dim lng_5       As Long
    Dim lng_6       As Long
    Dim lng_Other   As Long
    Dim dblValue    As Double
    Dim adblData()  As Double

    Dim lngStart    As Long
    Dim strElapsed  As String

    Const UPPER     As Long = 6
    Const LOWER     As Long = 1
    Const FILE_NAME As String = "_Test04.txt"

    PrepNames FILE_NAME
    
    lng_0 = 0
    lng_1 = 0
    lng_2 = 0
    lng_3 = 0
    lng_4 = 0
    lng_5 = 0
    lng_6 = 0
    lng_Other = 0
    lngCount = 0
    lngIndex = 0
    lngRange = 1 + (UPPER - LOWER)
    lngMaxAmt = MAX_LIMIT * 10
    
    strTitle = "Test 4:" & vbNewLine & "Throw a single dice " & Format$(lngMaxAmt, "#,0") & _
               " times" & vbNewLine & "using " & mstrGenerator & " generator."

    Erase adblData()            ' Empty array
    lngStart = GetTickCount()   ' starting time
    
    ' Generate random numbers between 1 and 6
    Do
        If lngCount Mod MAX_LIMIT = 0 Then
            
            lngIndex = 0
            
            ' generate random float values
            Select Case mlngAlgorithm
                   Case 0: adblData() = mobjPrng.BuildRndData(MAX_LIMIT, ePRNG_DBL_ARRAY, False)
                   Case 1: adblData() = mobjPrng.ISAAC_Prng(MAX_LIMIT, True)
                   Case 2: adblData() = mobjPrng.KISS_Prng(MAX_LIMIT, True)
                   Case 3: adblData() = mobjPrng.MWC_Prng(MAX_LIMIT, True)
                   Case 4: adblData() = mobjPrng.MOA_Prng(MAX_LIMIT, True)
                   Case 5: adblData() = mobjPrng.MT_Prng(MAX_LIMIT, True)
                   Case 6: adblData() = mobjPrng.MTA_Prng(MAX_LIMIT, True)
                   Case 7: adblData() = mobjPrng.MTB_Prng(MAX_LIMIT, True)
                   Case 8: adblData() = mobjPrng.TT800_Prng(MAX_LIMIT, True)
            End Select
        
        End If
        
       ' generate a value that falls between two numbers, inclusive
        dblValue = Abs((((UPPER + LOWER) - 1) * adblData(lngIndex)) + LOWER)
        lngValue = Int((dblValue * 6) + 1)
        lngIndex = lngIndex + 1
        
        Do While lngValue > UPPER
            lngValue = lngValue - lngRange
        Loop
        
        Do While LOWER > lngValue
            lngValue = lngValue + lngRange
        Loop
        
        ' Determine the value's worth
        Select Case lngValue
               Case 0:  lng_0 = lng_0 + 1
               Case 1:  lng_1 = lng_1 + 1
               Case 2:  lng_2 = lng_2 + 1
               Case 3:  lng_3 = lng_3 + 1
               Case 4:  lng_4 = lng_4 + 1
               Case 5:  lng_5 = lng_5 + 1
               Case 6:  lng_6 = lng_6 + 1
               Case Else:  lng_Other = lng_Other + 1
        End Select
        
        lngCount = lngCount + 1
        If lngCount > lngMaxAmt Then
            Exit Do
        End If
    Loop

    strElapsed = ElapsedTime(GetTickCount() - lngStart) ' finish time
    
    ' Fill the output file
    hFile = FreeFile                     ' get first free file handle
    Open mstrFile For Output As #hFile   ' create empty output file

    Print #hFile, strTitle               ' print the title
    Print #hFile, " "
    Print #hFile, "Version: " & mobjPrng.Version  ' Save DLL version information
    Print #hFile, " "

    Print #hFile, "Elapsed:  " & strElapsed    ' elapsed time
    Print #hFile, " "
    Print #hFile, "0.  " & lng_0
    Print #hFile, "1.  " & lng_1
    Print #hFile, "2.  " & lng_2
    Print #hFile, "3.  " & lng_3
    Print #hFile, "4.  " & lng_4
    Print #hFile, "5.  " & lng_5
    Print #hFile, "6.  " & lng_6
    Print #hFile, "Other.  " & lng_Other

    Close #hFile
    Erase adblData()

End Sub

' ****************************************************************************
' Test5 - Binary test file
'
' This will build the approximate 11mb (11,468,800 bytes) binary input file
' needed when using Diehard or ENT for randomness testing.  On a 500mhz,
' 256mb RAM PC, this will take less than two minutes to create the test file.
'
' Test file size:  2,867,200 32-bit random integers (11,468,800 bytes)
'
' ---------------------------------
' Randomness testing software
' ---------------------------------
' Diehard by George Marsaglia
' http://stat.fsu.edu/pub/diehard/
'
' ENT Software
' http://www.fourmilab.ch/random/
' Scroll down and download the file Random.zip
'
' ****************************************************************************
Private Sub Test5()
    
    Dim strTarget    As String
    Dim strElapsed   As String
    Dim hFile        As Long
    Dim lngStop      As Long
    Dim lngStart     As Long
    Dim lngAmtLeft   As Long
    Dim lngArraySize As Long
    Dim lngPointer   As Long
    Dim alngData()   As Long
    Dim abytData()   As Byte
    
    Const MAX_AMT As Long = 11468800    ' File size in bytes

    If mlngAlgorithm = 0 Then
        lngAmtLeft = MAX_AMT            ' Number of bytes needed
    Else
        lngAmtLeft = (MAX_AMT \ 4)      ' Number of long integers needed
    End If
    
    Erase alngData()                    ' empty arrays
    Erase abytData()
    lngPointer = 1                      ' init pointer for output file
    PrepNames ".bin"                    ' Prepare output file name

    hFile = FreeFile                    ' capture first free file handle
    Open mstrFile For Output As #hFile  ' Create an empty file
    Close #hFile                        ' close the file

    hFile = FreeFile                                  ' capture first free file handle
    Open mstrFile For Binary Access Write As #hFile   ' re-open file in binary mode

    '-----------------------------------------------------------------------------
    lngStart = GetTickCount()   ' starting time
    
    Do
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
    
        lngAmtLeft = lngAmtLeft - (lngArraySize - 1)
            
        If lngAmtLeft <= 1 Then
            Exit Do
        End If
    Loop
    
    strElapsed = ElapsedTime(GetTickCount() - lngStart)  ' Finish time
    '-----------------------------------------------------------------------------
    
    Close #hFile       ' Close open file
    Erase alngData()   ' Always empty arrays when not needed
    Erase abytData()
    
    Debug.Print strElapsed & "  " & mstrFile
    
End Sub

' ****************************************************************************
' Main_NoRepeat:   RANDOM NUMBER GENERATOR
'
' Non-Repeating numbers using CryptoAPI
' Rename this routine to "Main" and then press F5 to execute.
'
' ****************************************************************************
Private Sub Main_NoRepeat()

    ' Load the non-repeating number form
    Load frmNoRepeat

End Sub

' ****************************************************************************
' Main_Craps:  RANDOM NUMBER GENERATOR
'
' This is the Craps test.
' Rename this routine to "Main" and then press F5 to execute.
'
' Excert from DieHard Tests.txt:
'
'      "This is the CRAPS TEST. It plays multiple games of craps, finds
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
' Rules for Craps (http://www.gambling-systems.om/crapsmodics.html)
'
'                   Craps - Pass line bet
'
' When it is your turn to throw the craps dice, you must determine
' whether to bet the pass line or the don't pass line. Most shooters,
' as well as most of the other craps players at the table, will bet
' the pass line, as it is the modic wager of craps.
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
'
' GeneratorTypes:
'    0 - CryptoAPI   3 - MWC (Multiply With Carry)   6 - MT11232A
'    1 - ISAAC       4 - MOA (Mother-of-All)         7 - MT11232B
'    2 - KISS        5 - MT19937                     8 - TT800
'
' This is the Craps test.
' Rename this routine to "Main" and then press F5 to execute.
'
' ****************************************************************************
Private Sub Main_Craps()

    frmCraps.Algorithm 0   ' <- Algorithm type
    
    ' Load the form and start playing
    Load frmCraps

End Sub

' ****************************************************************************
' Main_Blackjack:  RANDOM NUMBER GENERATOR
'
' This is the Blackjack test.
' Rename this routine to "Main" and then press F5 to execute.
'
' Below is an excert of the rules of blackjack.
'==========================================================================
' Rules for Blackjack (http://www.blackjackinfo.com/)
'
' A blackjack, or natural, is a total of 21 in your first two cards.
'
' The modic premise of the game is that you want to have a hand value that
' is closer to 21 than that of the dealer, without going over 21.
'
' Once all the bets are made, the dealer will deal the cards to the players.
' He'll make two passes around the table starting at his left (your right)
' so that the players and the dealer have two cards each.
'
' The dealer must play his hand in a specific way, with no choices allowed.
' There are two popular rule variations that determine what totals the dealer
' must draw to.   In any given casino, you can tell which rule is in effect
' by looking at the blackjack tabletop.  It should be clearly labeled with
' one of these rules:
'
'    "Dealer stands on all 17s":  This is the most common rule.  In this
'    case, the dealer must continue to take cards ("hit") until his total
'    is 17 or greater.  An Ace in the dealer's hand is always counted as
'    11 if possible without the dealer going over 21.  For example, (Ace,8)
'    would be 19 and the dealer would stop drawing cards ("stand").  Also,
'    (Ace,6) is 17 and again the dealer will stand.  (Ace,5) is only 16, so
'    the dealer would hit.  He will continue to draw cards until the hand's
'    value is 17 or more.  For example, (Ace,5,7) is only 13 so he hits
'    again.  (Ace,5,7,5) makes 18 so he would stop ("stand") at that point.
'
'    "Dealer hits soft 17":  Some casinos use this rule variation instead.
'    This rule is identical except for what happens when the dealer has a
'    soft total of 17.  Hands such as (Ace,6),  (Ace,5,Ace), and (Ace,2,4)
'    are all examples of soft 17.  The dealer hits these hands, and stands
'    on soft 18 or higher, or hard 17 or higher.  When this rule is used,
'    the house advantage against the players is slightly increased.
'
' Again, the dealer has no choices to make in the play of his hand.  He
' cannot split pairs, but must instead simply hit until he reaches at least
' 17 or busts by going over 21.
'
' Of course, this does not take into consideration any "Stupid human tricks".
' Such as, "I know a system." or "I just know that next card is not a 10 or
' a face card.".
'
' GeneratorTypes:
'    0 - CryptoAPI   3 - MWC (Multiply With Carry)   6 - MT11232A
'    1 - ISAAC       4 - MOA (Mother-of-All)         7 - MT11232B
'    2 - KISS        5 - MT19937                     8 - TT800
'
' This is the Blackjack test.
' Rename this routine to "Main" and then press F5 to execute.
'
' ****************************************************************************
Private Sub Main_Blackjack()

    frmBlackjack.Algorithm 5   ' <- Algorithm type
    
    ' Load the form and start playing
    Load frmBlackjack
    
End Sub

Private Sub PrepNames(ByRef strFilename As String)

    Dim strPath As String
    
    strPath = FormatPath()
    
    Select Case mlngAlgorithm
           Case 0
                mstrFile = strPath & "MS" & strFilename
                mstrGenerator = "MS (Microsoft CryptoAPI)"
           Case 1
                mstrFile = strPath & "IC" & strFilename
                mstrGenerator = "IC (ISAAC)"
           Case 2
                mstrFile = strPath & "KISS" & strFilename
                mstrGenerator = "KISS (Keep It Simple Stupid)"
           Case 3
                mstrFile = strPath & "MWC" & strFilename
                mstrGenerator = "MWC (Multiply With Carry)"
           Case 4
                mstrFile = strPath & "MOA" & strFilename
                mstrGenerator = "MOA (Mother-of-All)"
           Case 5
                mstrFile = strPath & "MT" & strFilename
                mstrGenerator = "MT (Mersenne Twister 19937)"
           Case 6
                mstrFile = strPath & "MTA" & strFilename
                mstrGenerator = "MTA (Mersenne Twister 11231A)"
           Case 7
                mstrFile = strPath & "MTB" & strFilename
                mstrGenerator = "MTB (Mersenne Twister 11231B)"
           Case 8
                mstrFile = strPath & "TT8" & strFilename
                mstrGenerator = "TT8 (TT800)"
    End Select

End Sub

' ***************************************************************************
' Routine:       ElapsedTime
'
' Description:   Formats time display
'
' Reference:     Karl E. Peterson, http://vb.mvps.org/
'
' Returns:       Formatted output
'                Ex:  12:34:56.789  <- 12 hours 34 minutes 56 seconds 789 thousandths
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function ElapsedTime(ByVal lngMilliseconds As Long) As String

    Dim lngDays As Long
    
    Const ONE_DAY As Long = 86400000   ' Number of milliseconds in a day
    
    ElapsedTime = vbNullString                 ' Verify output string is empty
    lngDays = Fix(lngMilliseconds / ONE_DAY)   ' Calculate number of days
        
    ' See if one or more days has passed
    If lngDays > 0 Then
        ElapsedTime = CStr(lngDays) & " day(s)  "                 ' Start loading output string
        lngMilliseconds = lngMilliseconds - (ONE_DAY * lngDays)   ' Calculate number of milliseconds left
    End If

    ' Continue formatting output string as HH:MM:SS
    ElapsedTime = ElapsedTime & Format$(DateAdd("s", (lngMilliseconds \ 1000), #12:00:00 AM#), "HH:MM:SS")
    lngMilliseconds = lngMilliseconds - ((lngMilliseconds \ 1000) * 1000)   ' Calc number of milliseconds left
    
    ' Append thousandths to output string
    ElapsedTime = ElapsedTime & "." & Format$(lngMilliseconds, "000")
   
End Function

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Determines whether a path to a file system object such as
'                a file or directory is valid. This function tests the
'                validity of the path. A path specified by Universal Naming
'                Convention (UNC) is limited to a file only; that is,
'                \\server\share\file is permitted. A UNC path to a server
'                or server share is not permitted; that is, \\server or
'                \\server\share. This function returns FALSE if a mounted
'                remote drive is out of service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'
' Reference:     http://msdn.microsoft.com/en-us/library/bb773584(v=vs.85).aspx
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))
   
End Function
 
Public Function FormatPath() As String

    Dim strPath As String
    
    strPath = App.Path
    strPath = Mid$(strPath, 1, InStrRev(strPath, "\"))
    strPath = IIf(Right$(strPath, 1) = "\", strPath, strPath & "\")
    strPath = strPath & "Data_Output"
    strPath = IIf(Right$(strPath, 1) = "\", strPath, strPath & "\")
    
    If Not IsPathValid(strPath) Then
        MkDir strPath
    End If
    
    FormatPath = strPath
    
End Function
