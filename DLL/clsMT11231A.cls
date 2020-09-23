VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMT11231A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsMT11231A
'
' MT11231A is a variation of the Mersenne Twister algorithm.
' It has a seed value of 2^11231-1.
'
' This will pass all of the Diehard and ENT randomness tests.
' Code has been modified for only two types of output.
'   Long Integer       -2147483648 to 2147483647
'   Double precision   -0.9999999999999 to 0.9999999999999
'
' For the unabridged VBA code, visit the Mersenne Twister Home Page
' http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/emt.html
' Look under Mersenne Twister "Various versions/languages links/codes".
'
' I am using Pablo Ronchi's VBA code to create my version of the
' Mersenne Twister algorithm.
' http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/VERSIONS/BASIC/basic.html
'
' This is one of three additional versions of the Mersenne Twister algorithm
' named MT11231A, MT11231B & TT800.
'
' Reference:  The Mersenne Twister (variations)
'             http://www.quadibloc.com/crypto/co4814.htm
'
' ===========================================================================
' This is the Visual Basic for Applications (VBA) version of the  MT11231A,
' or "MERSENNE TWISTER" algorithm for pseudo random number generation,
' with initialization improved, by  MAKOTO MATSUMOTO  and  TAKUJI NISHIMURA,
' of 2002/1/26.
'
' This original translation to VBA was made and tested by Pablo Mariano Ronchi
' (2005-Sep-12)
'
' Note 1: VBA is the Visual Basic language used in MS Access, MS Excel and,
'         in general, in MS Office, and is called simply "Visual Basic" or
'         VBA, hereinafter.
' Note 2: This same code compiles in Visual Basic (VB) without modifications.
'
' Please read the comments about this VBA version that follow the ones below,
' by the authors of the "MERSENNE TWISTER" algorithm.
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
' 21-Sep-2005  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' 10-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated seed creation process.  The seed creation process can
'              make or break a random number generator.
' 11-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              Corrected logic flaw in CalcOneLong() routine.
' 05-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added call to MakeCryptoValue() routine to create crypto
'                quality numbers.
'              - CreateSeed() routine is only called once during the life of
'                this module.  Seed values are updated at the end of each
'                random value creation.  See CalcOneLong() routine.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME   As String = "clsMT11231A"
  Private Const N             As Long = 351
  Private Const K             As Long = 175
  Private Const SH_1          As Long = &H2             ' 2 = (2^1)
  Private Const SH_7          As Long = &H80            ' 128 = (2^7)
  Private Const SH_11         As Long = &H800&          ' 2048 = (2^11)
  Private Const SH_15         As Long = &H8000&         ' 32768 = (2^15)
  Private Const SH_17         As Long = &H20000         ' 131072 = (2^17)
  Private Const FIRST_13_BITS As Long = &H1000          ' 4096 = (2^12)
  Private Const LAST_19_BITS  As Long = &H40000         ' 262144 = (2^18)
  Private Const GB_4          As Double = (2# ^ 32)     ' 4294967296
  Private Const MAX_DWORD     As Double = (2# ^ 32) - 1 ' 4294967295
  Private Const DBL_LOW       As Double = 0.000000000001
  Private Const DBL_HIGH      As Double = (1.999999999998 / MAX_DWORD)

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
  Private mblnSeeded    As Boolean
  Private mcurFrequency As Currency
  Private mlngCount     As Long
  Private malngSeed()   As Long


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Get Version() As String
    Version = DLL_NAME & ".dll (tm) v" & CStr(App.Major) & "." & _
              CStr(App.Minor) & "." & CStr(App.Revision) & _
              " Copyright (c) 2004-2012 - Kenneth Ives  kenaso@tx.rr.com"
End Property

Public Property Let StopProcessing(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents
End Property



' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       MTA_Prng
'
' Description:   Returns an array of random generated values based on
'                quantity requested.
'
' Parameters:    lngArraySize - [Optional] Number of elements in array
'                     Default return number of 1.
'                blnReturnFloat - [Optional] TRUE (Default) = Return double
'                     precision values in an array.  FALSE = Return long
'                     integer values in an array.
'
' Returns:       an array of random generated values
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Sep-2005  Kenneth Ives  kenaso@tx.rr.com
'              Module created
' 02-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Removed code that created multiple occurances of clsRandom.cls
' ***************************************************************************
Public Function MTA_Prng(Optional ByVal lngArraySize As Long = 1, _
                         Optional ByVal blnReturnFloat As Boolean = True) As Variant

    Dim lngIndex   As Long
    Dim alngData() As Long
    Dim adblData() As Double

    Const ROUTINE_NAME As String = "MTA_Prng"
    
    On Error GoTo MTA_Prng_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo MTA_Prng_CleanUp
    End If
            
    Erase alngData()  ' Always start with empty arrays
    Erase adblData()
    
    If blnReturnFloat Then
        ReDim adblData(lngArraySize)
    Else
        ReDim alngData(lngArraySize)
    End If
    
    ' Load random data into array
    For lngIndex = 0 To (lngArraySize - 1)

        If blnReturnFloat Then
            ' returns -0.9999999999999 to 0.9999999999999
            adblData(lngIndex) = CalcOneDouble()
        Else
            ' returns -2147483648 to 2147483647
            alngData(lngIndex) = CalcOneLong()
        End If

    Next lngIndex

MTA_Prng_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If Not gblnStopProcessing Then

        If blnReturnFloat Then
            MTA_Prng = adblData()
        Else
            MTA_Prng = alngData()
        End If
    End If
    
    Erase alngData()  ' Always empty arrays when not needed
    Erase adblData()
    mlngCount = 0
    
    On Error GoTo 0
    Exit Function

MTA_Prng_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume MTA_Prng_CleanUp

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

Private Function CalcOneLong() As Long

    ' Hex values in this routine are the main difference between MT11231A
    ' and MT11231B
    '
    ' Pablo Ronchi's method of performing bit shifting is fast and reliable.
    ' First calculate bit shift count [ex:  shift bits 10  (1024=2^10)].
    ' To shift left, multiply [ex:  x=uMult(n,1024)]
    ' To shift right, divide  [ex:  x=uDivide(n,1024)]
    '
    ' Ex:  Shift bits     10
    '      Number         12345
    '      Left shift     12641280
    '      Right shift    12

    Dim lngPos   As Long
    Dim lngRand  As Long
    Dim lngIndex As Long
    Dim lngMagic As Long

    If mlngCount > N Then

        ' CreateSeed routine is only called once
        ' during the life of this module.  Seed
        ' values are updated at the end of each
        ' random value creation.
        If Not mblnSeeded Then
            CreateSeed
        End If
        
        For lngIndex = 0 To (N - K)
        
            lngRand = (malngSeed(lngIndex) And FIRST_13_BITS) Or _
                      (malngSeed(lngIndex + 1) Xor LAST_19_BITS)
            
            lngMagic = IIf(lngMagic Mod 2 = 0, &HE4BD75F5, 0)  ' Toggle values
            
            malngSeed(lngIndex) = lngMagic Xor (malngSeed(lngIndex + K) Xor _
                                  UnsignedDivide(lngRand, SH_1))
        Next lngIndex
        
        lngPos = lngIndex   ' Remember last position
        
        For lngIndex = lngPos To (N - 1)
            
            lngRand = (malngSeed(lngIndex) And FIRST_13_BITS) Or _
                      (malngSeed(lngIndex + 1) Xor LAST_19_BITS)
                      
            lngMagic = IIf(lngMagic Mod 2 = 0, &HE4BD75F5, 0)  ' Toggle values
            
            malngSeed(lngIndex) = lngMagic Xor (malngSeed(lngIndex + (K - N)) Xor _
                                  UnsignedDivide(lngRand, SH_1))
        Next lngIndex

        lngRand = (malngSeed(N - 1) And FIRST_13_BITS) Or _
                  (malngSeed(0) Xor LAST_19_BITS)
        
        lngMagic = IIf(lngMagic Mod 2 = 0, &HE4BD75F5, 0)  ' Toggle values
            
        malngSeed(N - 1) = lngMagic Xor (malngSeed(K - 1) Xor _
                           UnsignedDivide(lngRand, SH_1))
        mlngCount = 0
    
    End If

    lngRand = malngSeed(mlngCount)
    lngRand = (lngRand Xor UnsignedDivide(lngRand, SH_11))                    ' Shift right
    lngRand = (lngRand Xor UnsignedMultiply(lngRand, SH_7) And &H655E5280)    ' Shift left
    lngRand = (lngRand Xor UnsignedMultiply(lngRand, SH_15) And &HFFD58000)   ' Shift left
    lngRand = (lngRand Xor UnsignedDivide(lngRand, SH_17))                    ' Shift right
        
    ' 05-May-2012  Update seed values with each random value creation
    malngSeed(mlngCount) = lngRand Or malngSeed(mlngCount)
    malngSeed(mlngCount) = lngRand And malngSeed(mlngCount)
    
    mlngCount = mlngCount + 1   ' increment counter
    CalcOneLong = lngRand       ' return random generated value
    
End Function

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
Private Function CalcOneDouble() As Double

    Dim dblRand As Double
    Dim dblTemp As Double

    dblRand = CDbl(CalcOneLong())  ' generate one long integer

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
' 10-Dec-2007  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 30-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added global seed holding variable because of faster CPU's
' 10-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated seed creation process.  The seed creation process can
'              make or break a random number generator.
' ***************************************************************************
Private Sub CreateSeed()

    Dim lngIndex As Long
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

    ReDim malngSeed(N + 1)    ' Size seed array
    
    For lngIndex = 0 To N
        
        malngSeed(lngIndex) = glngCarryOver Xor &HFFFF0000
        glngCarryOver = UnsignedMultiply(1812433253, glngCarryOver) + 1
        
        malngSeed(lngIndex) = malngSeed(lngIndex) Or _
                              UnsignedDivide((glngCarryOver And &HFFFF0000), 65536)
        glngCarryOver = UnsignedMultiply(1812433253, glngCarryOver) + 1
        
    Next lngIndex

    mblnSeeded = True   ' Designates seed array is intialized

End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objMT11731A = new cMT11731A
                
    Erase malngSeed()
    mlngCount = N + 1
    mblnSeeded = False
    QueryPerformanceFrequency mcurFrequency   ' Test for Hi-Performance timer
        
    ' Init property values
    StopProcessing = False
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objMT11731A = Nothing
    
    mcurFrequency = 0@
    Erase malngSeed()    ' Always empty arrays when not needed

End Sub

