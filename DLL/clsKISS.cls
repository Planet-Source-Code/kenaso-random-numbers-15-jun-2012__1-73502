VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cKISS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsKISS
'
' Description:   The KISS generator, (Keep It Simple Stupid), is designed
'                to combine the two multiply-with-carry generators in MWC
'                with the 3-shift register SHR3 and the congruential
'                generator CONG, using addition and exclusive-or.  Has a
'                period about 2^123.
'
'                The  MWC generator concatenates two 16-bit multiply-
'                with-carry generators, x(n)=36969x(n-1)+carry,
'                y(n)=18000y(n-1)+carry  mod 2^16, has period about
'                2^60 and seems to pass all tests of randomness. A
'                favorite stand-alone generator---faster than KISS,
'                which contains it.
'
'                SHR3 is a congruential generator with period 2^32-1.
'                It uses y(n)=y(n-1)(I+L^17)(I+R^13)(I+L^5), with the y's
'                viewed as binary vectors, L the 32x32 binary matrix that
'                shifts a vector left 1, and R its transpose. SHR3 seems
'                to pass all except those related to the binary rank test,
'                since 32 successive values, as binary vectors, must be
'                linearly independent, while 32 successive truly random
'                32-bit integers, viewed as binary vectors, will be
'                linearly independent only about 29% of the time.  The
'                leading half of its 32 bits seem to pass tests, but bits
'                in the last half are too regular.
'
'                CONG is a congruential generator with the widely used
'                69069 multiplier: x(n)=69069x(n-1)+1234567. It has a
'                period 2^32. The leading half of its 32 bits seem to
'                pass tests, but bits in the last half are too regular.
'
'                The generators MWC and KISS seem to pass all Diehard
'                tests.  By themselves, CONG and SHR3 do not.
'
' References:    KISS Random number generator for Visual Basic
'                by Kenneth Ives  kenaso@tx.rr.com
'
'                George Marsaglia code in C
'                http://www.ciphersbyritter.com/NEWS4/RANDC.HTM
'
'                George Marsaglia geo@stat.fsu.edu
'                http://stat.fsu.edu/pub/diehard/
'
' *****************************************************************************
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE COPYRIGHT
' OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
' SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED
' TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
' PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
' LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
' NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 10-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated seed creation process.  The seed creation process can
'              make or break a random number generator.
' 05-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Force a reseed after approximately one million random
'                values have been genereated.
'              - Speedup process by not accessing bit shifting routine.
'                I tried several routines and all were slow.  Instead I
'                used Pablo Ronchi's method of calculating the bitshift
'                count [Example shift 10 (1024=2^10)].  To shift left, you
'                multiply [ex:  n=uMultiply(n,1024)].  To shift right, you
'                divide.  I could not believe how much faster this process
'                has become.
'                Ex:  n=12345   Bitshift=10
'                     Left shift    12641280
'                     Right shift   12
' 18-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed an oversight where results are stored.  See CalcLong()
'                routine.
'              - Sped up process by calculating multiple random values at a
'                time.  See KISS_Prng() routine.
' 05-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added call to MakeCryptoValue() routine to create crypto
'              quality numbers.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsKISS"
  Private Const MAX_RAND    As Long = 256
  Private Const RND_COUNT   As Long = 1000000          ' One million
  Private Const SH_5        As Long = 32               ' 32 = (2^5)
  Private Const SH_13       As Long = 8192             ' 8192 = (2^13)
  Private Const SH_16       As Long = 65536            ' 65536 = (2^16)
  Private Const SH_17       As Long = 131072           ' 131072 = (2^17)
  Private Const GB_4        As Double = (2# ^ 32)      ' 4294967296
  Private Const MAX_DWORD   As Double = (2# ^ 32) - 1  ' 4294967295
  Private Const DBL_LOW     As Double = 0.000000000001
  Private Const DBL_HIGH    As Double = (1.999999999998 / MAX_DWORD)
  
' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Long Integer)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a lng Seed
' Variable name:     malngSeed
'
' ***************************************************************************
  Private mcurFrequency As Currency
  Private mlngRndCnt    As Long
  Private malngRand()   As Long
  Private malngSeed()   As Long


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Get Version() As String
Attribute Version.VB_Description = "Read only - designates DLL information"
    Version = DLL_NAME & ".dll (tm) v" & CStr(App.Major) & "." & _
              CStr(App.Minor) & "." & CStr(App.Revision) & _
              " Copyright (c) 2004-2012 - Kenneth Ives  kenaso@tx.rr.com"
End Property

Public Property Let StopProcessing(ByVal blnData As Boolean)
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing"
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing"
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents
End Property



' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       KISS_Prng
'
' Description:   Generate one long integer or one double precision value
'                of cryptographic quality.
'
' Parameters:    lngArraySize - [Optional] Number of elements in array
'                     Default return number of 1.
'                blnReturnFloat - [Optional] TRUE (Default) = Return double
'                     precision values in an array.  FALSE = Return long
'                     integer values in an array.
'
' Returns:       An array of random generated values
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Sep-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 18-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              Sped up process by calculating multiple random values at a
'              time.
' ***************************************************************************
Public Function KISS_Prng(Optional ByVal lngArraySize As Long = 1, _
                          Optional ByVal blnReturnFloat As Boolean = True) As Variant

    Dim lngIdx     As Long
    Dim lngCount   As Long
    Dim lngIndex   As Long
    Dim alngData() As Long
    Dim adblData() As Double
    
    Const ROUTINE_NAME As String = "KISS_Prng"
    
    On Error GoTo KISS_Prng_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo KISS_Prng_CleanUp
    End If
            
    lngCount = 0
    
    If blnReturnFloat Then
        ReDim adblData(lngArraySize)
    Else
        ReDim alngData(lngArraySize)
    End If
    
    ' If more than one million random values
    ' have been created then reseed
    If mlngRndCnt > RND_COUNT Then
        CreateSeed   ' Reseed arrays
    End If
    
    ' Load random data into array
    For lngIndex = 0 To (lngArraySize - 1)
    
        CalcLong   ' Load array with random values
        
        ' unload the random data array into the appropriate return array
        For lngIdx = 0 To (MAX_RAND - 1)
        
            If blnReturnFloat Then
                ' returns -0.9999999999999 to 0.9999999999999
                adblData(lngCount) = CalcOneDouble(malngRand(lngIdx))
            Else
                ' returns -2147483648 to 2147483647
                alngData(lngCount) = malngRand(lngIdx)
            End If

            lngCount = lngCount + 1   ' increment current counter
    
            ' If the requested number of elements have
            ' been collected then exit this loop
            If lngCount = lngArraySize Then
                Exit For    ' exit For..Next loop
            End If
            
        Next lngIdx

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

        ' increment reseed counter
        mlngRndCnt = mlngRndCnt + lngCount
        
        ' If more than one million random values
        ' have been created then reseed
        If mlngRndCnt > RND_COUNT Then
            CreateSeed   ' Reseed arrays
        End If
    
        ' If requested number of elements have
        ' been collected then exit this loop
        If lngCount = lngArraySize Then
            Exit For  ' exit For..Next loop
        End If
        
    Next lngIndex

KISS_Prng_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If Not gblnStopProcessing Then

        If blnReturnFloat Then
            KISS_Prng = adblData()
        Else
            KISS_Prng = alngData()
        End If
    End If
    
    Erase alngData()  ' Always empty arrays when no not needed
    Erase adblData()
    
    On Error GoTo 0
    Exit Function

KISS_Prng_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume KISS_Prng_CleanUp

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

Private Sub CalcLong()

    Dim lngIndex As Long
    Dim MWC      As Long
    Dim SHR3     As Long
    Dim CONG     As Long
    
    ReDim malngRand(MAX_RAND)   ' Size random number array

    ' variables z, w, jsr, jcong should be assigned
    ' seed values other than their initial values.
    '    w     = malngSeed(0)
    '    z     = malngSeed(1)
    '    jsr   = malngSeed(2)
    '    jcong = malngSeed(3)
    '
    For lngIndex = 0 To (MAX_RAND - 1)
    
        ' #define wnew   (w=18000*(w&65535)+(w>>16))
        malngSeed(0) = UnsignedMultiply(18000, _
                       UnsignedAdd((malngSeed(0) And 65535), _
                       UnsignedDivide(malngSeed(0), SH_16)))
    
        ' 26-Jun-2011 Changed AND to XoR (z&65535) -> (z^65535)
        ' #define znew   (z=36969*(z^65535)+(z>>16))
        malngSeed(1) = UnsignedMultiply(36969, _
                       UnsignedAdd((malngSeed(1) Xor 65535), _
                       UnsignedDivide(malngSeed(1), SH_16)))
    
        ' 05-Nov-2011 Changed shift ((z<<16)+w) to rotate ((z<<<16)+w)
        ' #define MWC    ((z<<<16)+w)
        malngSeed(1) = UnsignedAdd((RotateLong(malngSeed(1), 16)), malngSeed(0))
        MWC = malngSeed(1)
    
        ' 10-Jun-2012 Changed shift (jsr^=(jsr<<17)) to rotate (jsr^=(jsr<<<17))
        ' #define SHR3   (jsr^=(jsr<<17), jsr^=(jsr>>13), jsr^=(jsr<<5))
        malngSeed(2) = malngSeed(2) Xor RotateLong(malngSeed(2), -17)
        malngSeed(2) = malngSeed(2) Xor UnsignedDivide(malngSeed(2), SH_13)
        malngSeed(2) = malngSeed(2) Xor UnsignedMultiply(malngSeed(2), SH_5)
        SHR3 = malngSeed(2)
    
        ' #define CONG  (jcong=69069*jcong+1234567)
        malngSeed(3) = UnsignedMultiply(malngSeed(3), 69069)
        malngSeed(3) = UnsignedAdd(malngSeed(3), 1234567)
        CONG = malngSeed(3)
        
        ' KISS = ((MWC^SHR3)+CONG)
        ' Store random number in return array
        malngRand(lngIndex) = UnsignedAdd((MWC Xor SHR3), CONG)
    
    Next lngIndex
    
End Sub

' ***************************************************************************
' Routine:       CalcOneDouble
'
' Description:   Convert a long integer to a double precision number.
'                Returns a decimal position of 14 places. This was
'                chosen so that exponentials would not be returned.
'                If you are working with large numbers then your return
'                values will probably be in string format so as to get
'                the full value of your calculations.
'
'                    Ex: 2147483647  ->  -2.32830732662872E-10
'
'                Desired return value is within this range:
'
'                    -0.99999999999999 to 0.99999999999999
'
'                    Ex: 2147483647  ->  -0.00000000023283
'
' Returns:       Double precision value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names
' ***************************************************************************
Private Function CalcOneDouble(ByVal dblRand As Double) As Double

    Dim dblTemp As Double

    If dblRand < 0 Then
        dblTemp = dblRand + GB_4
    Else
        dblTemp = dblRand
    End If

    CalcOneDouble = FormatNumber(DBL_LOW + ((dblTemp * DBL_HIGH) - 1#), 14)

End Function

' ***************************************************************************
' I prefer to calculate the seed values.  Because the new CPU's are so fast,
' I have opted to carry over one of the generated values in a global variable
' just to make sure I have a unique value to start my calculations.  Getting
' just the system time will show that only 1-3 milliseconds may have changed
' since the last pass through this routine. This will generate almost, if not
' the same, values.  Using a global variable, these calculated values will
' always be different.
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Sub CreateSeed()

    Dim lngIndex As Long
    Dim lngLoop  As Long
    Dim curData  As Currency
    
    ' Safety - in case the carryover seed
    '          has not been inititialized
    If glngCarryOver = 0 Then
    
        ' Test for hi-performance timer
        If mcurFrequency > 0 Then
            QueryPerformanceCounter curData                              ' Capture Hi-Performance microseconds
            glngCarryOver = CLng(Right$(CStr(Fix(curData * 10000)), 9))  ' Capture last nine digits
        Else
            glngCarryOver = GetTickCount()                               ' Capture millseconds
        End If
        
        glngCarryOver = (glngCarryOver - (glngCarryOver \ 10))           ' Subtract approx 10%
    End If
    
    ReDim malngSeed(4)  ' Size seed array
    
    For lngLoop = 1 To 2
        For lngIndex = 0 To 3

            glngCarryOver = UnsignedMultiply(glngCarryOver, 29943829) - 1
            malngSeed(lngIndex) = UnsignedMultiply(glngCarryOver, 65536)
            glngCarryOver = UnsignedAdd(glngCarryOver, malngSeed(lngIndex))
        
        Next lngIndex
    Next lngLoop
    
    mlngRndCnt = 0   ' Reset counter to zero

End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objKISS = new cKISS
                
    Erase malngRand()   ' Always start with empty arrays
    Erase malngSeed()
    
    mlngRndCnt = RND_COUNT + 1                ' Force a reseed.  See Isaac_PRNG()
    QueryPerformanceFrequency mcurFrequency   ' Test for Hi-Performance timer

    ' Init property values
    StopProcessing = False   ' Preset property value

End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objKISS = Nothing
    
    mcurFrequency = 0@
    Erase malngRand()    ' Always empty arrays when not needed
    Erase malngSeed()

End Sub

