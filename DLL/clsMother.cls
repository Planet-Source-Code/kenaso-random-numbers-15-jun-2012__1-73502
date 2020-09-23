VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMother"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsMother (MOA = Mother of All)
'
' Description:   George Marsaglia's comments:
'
'                Yet another Random Number Generator
'
'                Random number generators are frequently posted on the
'                network; my colleagues and I posted ULTRA in 1992 and,
'                from the number of requests for releases to use it in
'                software packages, it seems to be widely used.
'
'                I have long been interested in Random Number Generator's
'                and several of my early ones are used as system generators
'                or in statistical packages.
'
'                So why another one?  And why here?
'
'                Because I want to describe a generator, or rather,
'                a class of generators, so promising I am inclined to
'                Call it
'
'                   "The Mother of All Random Number Generators"
'
'                and because the generator seems promising enough to
'                justify shortcutting the many months, even years, before
'                new developments are widely known through publication in
'                a journal.
'
'                This new class leads to simple, fast programs that produce
'                sequences with very long periods. They use multiplication,
'                which experience has shown does a better job of mixing
'                bits than do +,- or exclusive-or, and they do it with
'                easily implemented arithmetic modulo a power of 2, unlike
'                arithmetic modulo a prime. The latter, while satisfactory,
'                is difficult to implement. But the arithmetic here modulo
'                2^16 or 2^32 does not suffer the flaws of ordinary
'                congruential generators for those moduli: trailing bit
'                too regular.  On the contrary, all bits of the integers
'                produced by this new method, whether leading or trailing,
'                have passed extensive tests of randomness.
'
'                Here is an idea of how it works, using, say, integers of
'                six decimal digits from which we return random 3-digit
'                integers.  Start with n=123456, the seed.
'
'                Then form a new n=672*456+123=306555 and return 555.
'                Then form a new n=672*555+306=373266 and return 266.
'                Then form a new n=672*266+373=179125 and return 125,
'
'                and so on.  Got it?  This is a multiply-with-carry sequence
'                x(n)=672*x(n-1)+carry mod b=1000, where the carry is the
'                number of b's dropped in the modulus reduction. The resulting
'                sequence of 3-digit x's has period 335,999.  Try it.
'
'                No big deal, but that's just an example to give the idea. Now
'                consider the sequence of 16-bit integers produced by the two
'                C statements:
'
'                     k=30903*(k&65535)+(k>>16); return(k&65535);
'
'                Notice that it is doing just what we did in the example:
'                multiply the bottom half (by 30903, carefully chosen), add the
'                top half and return the new bottom.
'
'                That will produce a sequence of 16-bit integers with a period
'                greater than 2^29, and if we concatenate two such:
'
'                     k=30903*(k&65535)+(k>>16);
'                     j=18000*(j&65535)+(j>>16);
'                     return((k<<16)+j);
'
'                we get a sequence of more than 2^59 32-bit integers before
'                cycling.
'
'                The following segment in a (properly initialized) C procedure
'                will generate more than 2^118 32-bit random integers from six
'                random seed values I , j, k, l, m, n:
'
'                         k=30903*(k&65535)+(k>>16);
'                         j=18000*(j&65535)+(j>>16);
'                         i=29013*(i&65535)+(i>>16);
'                         l=30345*(l&65535)+(l>>16);
'                         m=30903*(m&65535)+(m>>16);
'                         n=31083*(n&65535)+(n>>16);
'                         return((k+i+m)>>16)+j+l+n);
'
'                And it will do it much faster than any of several widely used
'                generators designed to use 16-bit integer arithmetic, such as
'                that of Wichman-Hill that combines congruential sequences for
'                three 15-bit primes (Applied Statistics, v31, p188-190, 1982),
'                period about 2^42.
'
'                I call these multiply-with-carry generators.  Here is an
'                extravagant 16-bit example that is easily implemented in C or
'                Fortran. It does such a thorough job of mixing the bits of the
'                previous eight values that it is difficult to imagine a test
'                of randomness it could not pass:
'
'                   x[n]=12013x[n-8]+1066x[n-7]+1215x[n-6]+1492x[n-5]+1776x[n-4]
'                        +1812x[n-3]+1860x[n-2]+1941x[n-1]+carry mod 2^16.
'
'                The linear combination occupies at most 31 bits of a 32-bit
'                integer. The bottom 16 is the output, the top 15 the next carry.
'                It is probably best to implement with 8 case segments. It takes
'                8 microseconds on my PC. Of course it just provides 16-bit random
'                integers, but awfully good ones. For 32 bits you would have to
'                combine it with another, such as:
'
'                   x[n]=9272x[n-8]+7777x[n-7]+6666x[n-6]+5555x[n-5]+4444x[n-4]
'                        +3333x[n-3]+2222x[n-2]+1111x[n-1]+carry mod 2^16.
'
'                Concatenating those two gives a sequence of 32-bit random integers
'                (from 16 random 16-bit seeds), period about 2^250.  It is so
'                awesome it may merit the Mother of All Random Number Generator's
'                title.
'
'                The coefficients in those two linear combinations suggest that it
'                is easy to get long-period sequences, and that is true.  The result
'                is due to Cemal Kac, who extended the theory we gave for
'                add-with-carry sequences: Choose a base b and give r seed values
'                x[1],...,x[r] and an initial 'carry' c. Then the multiply-with-carry
'                sequence:
'
'                    x[n]=a1*x[n-1]+a2*x[n-2]+...+ar*x[n-r]+carry mod b,
'
'                where the new carry is the number of b's dropped in the modulus
'                reduction, will have period the order of b in the group of
'                residues relatively prime to m=ar*b^r+...+a1b^1-1.  Furthermore,
'                the x's are, in reverse order, the digits in the expansion of
'                k/m to the base b, for some 0<k<m.
'
'                In practice b=2^16 or b=2^32 allows the new integer and the new
'                carry to be the bottom and top half of a 32- or 64-bit linear
'                combination of  16-or 32-bit integers.  And it is easy to find
'                suitable m's if you have a primality test:
'
'                    just search through candidate coefficients until you get an
'                    m that is a safeprime---both m and (m-1)/2 are prime.  Then
'                    the period of the multiply-with-carry sequence will be the
'                    prime (m-1)/2. (It can't be m-1 because b=2^16 or 2^32 is a
'                    square.)
'
'                Here is an interesting simple MWC generator with period> 2^92,
'                for 32-bit arithmetic:
'
'                     x[n]=1111111464*(x[n-1]+x[n-2]) + carry mod 2^32.
'
'                Suppose you have functions, say top() and bot(), that give the
'                top and bottom halves of a 64-bit result.  Then, with initial
'                32-bit x, y and carry c, simple statements such as:
'
'                     y = bot(1111111464 * (x + y) + c)
'                     x = y
'                     c = Top(y)
'
'                will, repeated, give over 2^92 random 32-bit y's.
'
' References:    Mother_Of_All Random number generator for Visual Basic
'                by Kenneth Ives  kenaso@tx.rr.com
'
'                George Marsaglia code in C
'                ftp://ftp.taygeta.com/pub/c/mother.c
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
' 30-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated CalcOneLong() for a better randomness output.
' 10-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated seed creation process.  The seed creation process can
'              make or break a random number generator.
'              Updated code to more closely resemble the C version.
' 09-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Replaced some For..Next loops with API CopyMemory() in
'              routines CalcOneLong() and RandomInit() for speed and easier
'              maintenance.
' 02-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in CalcOneLong() routine that started with
'              the moving of the elements.
' 30-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug created when using API CopyMemory to copy from a
'              variant to a long integer without converting the data.  See
'              RandomInit() routine for loading base arrays.
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
'              - Fixed a bug. Corrected left shift to right shift in
'                RandomInit() routine.
'              - Sped up process by calculating multiple random values at a
'                time.  See MOA_Prng() routine.
' 05-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added call to MakeCryptoValue() routine to create crypto
'              quality numbers.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsMother"
  Private Const MAX_RAND    As Long = 256
  Private Const RND_COUNT   As Long = 1000000          ' One million
  Private Const SH_16       As Long = 65536            ' 65536 = (2^16)
  Private Const M15_MASK    As Long = 32767
  Private Const M16_MASK    As Long = 65535
  Private Const M16_LONG    As Long = 65536            ' 65536 = (2^16)
  Private Const MAX_LONG    As Long = 2147483647
  Private Const GB_4        As Double = (2# ^ 32)      ' 4294967296
  Private Const MAX_DWORD   As Double = (2# ^ 32) - 1  ' 4294967295
  Private Const DBL_LOW     As Double = 0.000000000001
  Private Const DBL_HIGH    As Double = (1.999999999998 / MAX_DWORD)
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Long Integer)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a lng MOA1
' Variable name:     malngMOA1
'
' ***************************************************************************
  Private mcurFrequency As Currency
  Private mlngRndCnt    As Long
  Private malngRand()   As Long
  Private malngMOA1()   As Long
  Private malngMOA2()   As Long
  Private malngBase1()  As Long
  Private malngBase2()  As Long

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
' Routine:       MOA_Prng
'
' Description:   George Marsaglia's The mother of all random number
'                generators producing uniformly distributed pseudo random
'                32 bit values with period about 2^250.
'
'                The arrays mother1 and mother2 store carry values in their
'                first element, and random 16 bit numbers in elements 1 to 8.
'                These random numbers are moved to elements 2 to 9 and a new
'                carry and number are generated and placed in elements 0 and 1.
'                The arrays mother1 and mother2 are filled with random 16 bit
'                values on first call of Mother by another generator.
'
'                A 32 bit random number is obtained by combining the output of
'                the two generators and returned in *pSeed.  It is also scaled
'                by 2^32-1 and returned as a double between 0 and 1.
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
Public Function MOA_Prng(Optional ByVal lngArraySize As Long = 1, _
                         Optional ByVal blnReturnFloat As Boolean = True) As Variant

    Dim lngIdx     As Long
    Dim lngCount   As Long
    Dim lngIndex   As Long
    Dim alngData() As Long
    Dim adblData() As Double
    
    Const ROUTINE_NAME As String = "MOA_Prng"
    
    On Error GoTo MOA_Prng_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo MOA_Prng_CleanUp
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
        RandomInit   ' Reseed arrays
    End If
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo MOA_Prng_CleanUp
    End If
            
    ' Load random data into array
    For lngIndex = 0 To (lngArraySize - 1)
    
        CalcLong   ' Load array with random values
        
        ' unload random data array into appropriate return array
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
            RandomInit   ' Reseed arrays
        End If
    
        ' If requested number of elements have
        ' been collected then exit this loop
        If lngCount = lngArraySize Then
            Exit For  ' exit For..Next loop
        End If
        
    Next lngIndex

MOA_Prng_CleanUp:
    DoEvents
    If Not gblnStopProcessing Then

        If blnReturnFloat Then
            MOA_Prng = adblData()
        Else
            MOA_Prng = alngData()
        End If
    End If
    
    Erase alngData()  ' Always empty arrays when no longer needed
    Erase adblData()
    
    On Error GoTo 0
    Exit Function

MOA_Prng_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume MOA_Prng_CleanUp

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

Private Function CalcLong() As Long
    
    Dim lngIdx     As Long
    Dim lngTemp    As Long
    Dim lngIndex   As Long
    Dim lngNumber1 As Long
    Dim lngNumber2 As Long
    
    ReDim malngRand(MAX_RAND)   ' Size random number array

    For lngIndex = 0 To (MAX_RAND - 1)
    
        ' Move elements 1-8 to 2-9
        ' memmove(mother1+2,mother1+1,8*sizeof(short));
        ' memmove(mother2+2,mother2+1,8*sizeof(short));
        MoveMemory malngMOA1(2), malngMOA1(1), 32&
        MoveMemory malngMOA2(2), malngMOA2(1), 32&
        
        ' Copy carry values to numberi
        ' number1=mother1[0];
        ' number2=mother2[0];
        lngNumber1 = malngMOA1(0)
        lngNumber2 = malngMOA2(0)
        
        ' Zero out element one.  Will
        ' be loaded after this loop.
        malngMOA1(1) = 0
        malngMOA2(1) = 0
        
        For lngIdx = 2 To 9
        
            ' number1+=1941*mother1[2]+1860*mother1[3]+1812*mother1[4]+1776*mother1[5]+
            '          1492*mother1[6]+1215*mother1[7]+1066*mother1[8]+12013*mother1[9];
            lngTemp = UnsignedMultiply(malngMOA1(lngIdx), malngBase1(lngIdx - 2))
            lngNumber1 = UnsignedAdd(lngNumber1, lngTemp)
        
            ' number2+=1111*mother2[2]+2222*mother2[3]+3333*mother2[4]+4444*mother2[5]+
            '          5555*mother2[6]+6666*mother2[7]+7777*mother2[8]+9272*mother2[9];
            lngTemp = UnsignedMultiply(malngMOA2(lngIdx), malngBase2(lngIdx - 2))
            lngNumber2 = UnsignedAdd(lngNumber2, lngTemp)
        
        Next lngIdx
    
        ' Put low bits of numberi into motheri[1]
        ' mother1[1]=m16Mask&number1;
        ' mother2[1]=m16Mask&number2;
        malngMOA1(1) = M16_MASK And lngNumber1
        malngMOA2(1) = M16_MASK And lngNumber2
        
        ' Put high bits of numberi into motheri[0] as new carry
        ' mother1[0]=number1/m16Long;
        ' mother2[0]=number2/m16Long;
        malngMOA1(0) = Int(lngNumber1 / M16_LONG)
        malngMOA2(0) = Int(lngNumber2 / M16_LONG)
                
        lngTemp = RotateLong(malngMOA1(1), 16)                     ' Rotate left
        malngRand(lngIndex) = UnsignedAdd(lngTemp, malngMOA2(1))   ' Insert data into return array
              
    Next lngIndex
    
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
Private Function CalcOneDouble(ByVal dblRand As Double) As Double

    Dim dblTemp As Double

    If dblRand < 0 Then
        dblTemp = dblRand + GB_4
    Else
        dblTemp = dblRand
    End If

    CalcOneDouble = FormatNumber(DBL_LOW + ((dblTemp * DBL_HIGH) - 1#), 14)

End Function

Private Sub RandomInit()

    Dim lngIdx    As Long
    Dim lngTemp   As Long
    Dim lngCount  As Long
    Dim lngNumber As Long
    Dim avntData  As Variant
    
    CreateSeed   ' Update carryover seed
    
    ReDim malngMOA1(10)  ' Size arrays
    ReDim malngMOA2(10)
    
    ' Initialize with 9 random values the first time
    '    sNumber= *pSeed&m16Mask;   /* The low 16 bits
    '    number= *pSeed&m31Mask;    /* Only want 31 bits
    '    p=mother1;
    lngTemp = glngCarryOver And M16_MASK
    lngNumber = glngCarryOver And MAX_LONG
    lngIdx = 0
    
    ' for (n=18;n--;)
    '     number=30903*sNumber+(number>>16);   /* One line multiply-with-carry
    '     *p++=sNumber=number&m16Mask;
    '     if (n==9)
    '         p=mother2;
    For lngCount = 19 To 0 Step -1
                
        lngTemp = UnsignedMultiply(30903, lngTemp)
        lngNumber = UnsignedAdd(lngTemp, UnsignedDivide(lngNumber, SH_16))  ' Shift right then add
        lngTemp = lngNumber And M16_MASK
               
        Select Case lngCount
               Case 0 To 9:   malngMOA2(lngIdx) = lngTemp  ' load Mother2 array
               Case 10 To 19: malngMOA1(lngIdx) = lngTemp  ' load Mother1 array
        End Select
        
        ' if at midway point then reset index
        ' to zero to start loading second array
        ' else increment index for current array
        lngIdx = IIf(lngCount = 10, 0, lngIdx + 1)
        
    Next lngCount

    ' /* make carry 15 bits
    ' mother1[0]&=m15Mask;
    ' mother2[0]&=m15Mask;
    malngMOA1(0) = malngMOA1(0) And M15_MASK
    malngMOA2(0) = malngMOA2(0) And M15_MASK
    
    '=================================================================
    ' Load base arrays
    '=================================================================
    ReDim malngBase1(8)   ' Size base arrays
    ReDim malngBase2(8)
    
    avntData = Empty   ' Verify variant is empty
    avntData = Array(1941, 1860, 1812, 1776, 1492, 1215, 1066, 12013)
    
    ' Copy variant data to an array
    ' while converting to long integer
    For lngIdx = 0 To 7
        malngBase1(lngIdx) = Val(avntData(lngIdx))
    Next lngIdx
    
    avntData = Empty   ' Verify variant is empty
    avntData = Array(1111, 2222, 3333, 4444, 5555, 6666, 7777, 9272)
    
    ' Copy variant data to an array
    ' while converting to long integer
    For lngIdx = 0 To 7
        malngBase2(lngIdx) = Val(avntData(lngIdx))
    Next lngIdx
        
    avntData = Empty   ' Always empty variants when not needed
    
End Sub

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
' 10-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated seed creation process.  The seed creation process can
'              make or break a random number generator.
' ***************************************************************************
Private Sub CreateSeed()

    Dim lngLoop As Long
    Dim lngTemp As Long
    Dim curData As Currency
    
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
        
    For lngLoop = 1 To 2
    
        glngCarryOver = UnsignedMultiply(glngCarryOver, 29943829) - 1
        lngTemp = UnsignedMultiply(glngCarryOver, 65536)
        glngCarryOver = UnsignedAdd(lngTemp, glngCarryOver)
    
    Next lngLoop
        
    mlngRndCnt = 0   ' Reset counter to zero

End Sub

Private Sub ResetVariables()

    Erase malngMOA1()
    Erase malngMOA2()
    Erase malngRand()
    Erase malngBase1()
    Erase malngBase2()
    
End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objMother = new cMother
                
    ResetVariables                            ' Always start with empty arrays
    mlngRndCnt = RND_COUNT + 1                ' Force a reseed
    QueryPerformanceFrequency mcurFrequency   ' Test for Hi-Performance timer

    ' Preset property values
    StopProcessing = False
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objMother = Nothing
    
    mcurFrequency = 0@
    ResetVariables       ' Always start with empty arrays

End Sub


