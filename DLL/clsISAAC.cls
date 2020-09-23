VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIsaac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsISAAC.cls
'
' ISAAC Random number generator for Visual Basic 6.0 and VBA
' by Kenneth Ives  kenaso@tx.rr.com
'
' This code is Public Domain. You may use this code as you like.
' There are no guarantees.
'
' Original C code by Bob Jenkins, March 1996
' http://www.burtleburtle.net/bob/rand/isaacafa.html
' Bob Jenkins  bob_jenkins@burtleburtle.net
'
' ***************************************************************************
'
' ACKNOWLEDGEMENTS:
'
'  Thank you Bob Jenkins for making your ISAAC code available to the public.
'
'  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar translated Mersenne Twister
'  to Visual Basic. I found his math routines invaluable. See uAdd() routine.
'  http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/VERSIONS/BASIC/basic.html
'
' ***************************************************************************
'  ISAAC has an output within these ranges:
'
'      -0.9999999999999, 0.9999999999999    Double Precision
'      -2147483648, 2147483647              Long Integer
'
' The output will pass all the Diehard and Ent randomness tests.
'
' Diehard by George Marsaglia
' http://stat.fsu.edu/pub/diehard/
'
' Ent Software
' http://www.fourmilab.ch/random/
' Scroll down and download the file Random.zip
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
' 21-Oct-2005  Kenneth Ives  kenaso@tx.rr.com
'              Combined calcs for a double, speeded up the process.
'              Removed obsolete variables.
'              Optimized code.  Gained 9-12% increase in speed.
' 10-Jan-2007  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete variables.
'              Optimized code.  Gained 10% increase in speed.
' 25-Oct-2007  Kenneth Ives  kenaso@tx.rr.com
'              Combined a redundant set of calculations in the RandomInit()
'              routine and gained 3/10ths of a second.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated RandomInit() routine by using static data instead of
'              performing calculations.
' 05-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Speeded up the process by changing some variables status and
'                relocating them in the routine.  See RandomInit().
'              - Updated ISAAC_Calc() routine for speed.
'              - Force a reseed after approximately one million random
'                values have been genereated.  See ISAAC_Prng() and
'                RandomInit() routines.
'              - Speedup process by not accessing bit shifting routine.
'                I tried several routines and all were slow.  Instead I
'                used Pablo Ronchi's method of calculating the bitshift
'                count [Example shift 10 (1024=2^10)].
'                To shift left, multiply [ex:  x=uMult(n,1024)]
'                To shift right, divide  [ex:  x=uDivide(n,1024)]
'
'                Ex:  Shift bits     10
'                     Number 12345
'                     Left shift     12641280
'                     Right shift    12
' 05-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added call to MakeCryptoValue() routine to create crypto
'              quality numbers.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsISAAC"
  Private Const MAX_RAND    As Long = 256
  Private Const RND_COUNT   As Long = 1000000          ' One million
  Private Const SH_2        As Long = 4                ' 4 = (2^2)
  Private Const SH_4        As Long = 16               ' 16 = (2^4)
  Private Const SH_6        As Long = 64               ' 64 = (2^6)
  Private Const SH_8        As Long = 256              ' 256 = (2^8)
  Private Const SH_9        As Long = 512              ' 512 = (2^9)
  Private Const SH_10       As Long = 1024             ' 1024 = (2^10)
  Private Const SH_11       As Long = 2048             ' 2048 = (2^11)
  Private Const SH_13       As Long = 8192             ' 8192 = (2^13)
  Private Const SH_16       As Long = 65536            ' 65536 = (2^16)
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
  Private mblnFirstTime As Boolean
  Private mcurFrequency As Currency
  Private mlngRndCnt    As Long
  Private malngMem()    As Long
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
' Routine:       ISAAC_Prng
'
' Description:   A quantity of random values will be generated based on the
'                user request.
'
' Parameters:    lngArraySize - [Optional] Number of elements in return array.
'                     Default return number of 1.
'                blnReturnFloat - [Optional] Return random values in an array.
'                     TRUE - Double precision
'                     FALSE - Long integer
'
' Returns:       An array of random generated values
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Sep-2005  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 21-Oct-2005  Kenneth Ives  kenaso@tx.rr.com
'              Combined calcs for a double, speeded up the process.
' 16-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              Added logic to perform complete reseed after one million or
'              more values have been generated
' ***************************************************************************
Public Function ISAAC_Prng(Optional ByVal lngArraySize As Long = 1, _
                           Optional ByVal blnReturnFloat As Boolean = True) As Variant

    Dim lngIdx     As Long
    Dim lngIndex   As Long
    Dim lngCount   As Long
    Dim alngData() As Long
    Dim adblData() As Double

    Const ROUTINE_NAME As String = "ISAAC_Prng"
    
    On Error GoTo ISAAC_Prng_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo ISAAC_Prng_CleanUp
    End If
            
    lngCount = 0
    
    ' Based on type of data desired, adjust
    ' the appropriate return array.
    If blnReturnFloat Then
        ReDim adblData(lngArraySize)
    Else
        ReDim alngData(lngArraySize)
    End If
    
    ' If more than one million random values
    ' have been created with current seeds
    ' then reload seed values
    If mlngRndCnt > RND_COUNT Then
        
        RandomInit True   ' Full seeding and load random number array
        
        ' If first time into this module then perform a second seeding
        ' because malngRand() array will now have some data to add to
        ' the mix for stronger values
        If mblnFirstTime Then
            mblnFirstTime = False   ' Reset flag to FALSE
            RandomInit True         ' Full seeding and reload random number array
        End If
    
    End If
    
    ' Start creating random data to be returned
    For lngIndex = 0 To (lngArraySize - 1)
        
        ISAAC_Calc   ' load random data array
       
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

        ' Increment random number counter
        mlngRndCnt = mlngRndCnt + lngCount
            
        ' If more than one million random values
        ' have been created then update seeds
        If mlngRndCnt > RND_COUNT Then
            RandomInit True   ' Full seeding and reload random number array
        Else
            RandomInit False  ' Partial mixing and calculations
        End If
        
        ' If requested number of elements have
        ' been collected then exit this loop
        If lngCount = lngArraySize Then
            Exit For  ' exit For..Next loop
        End If
        
    Next lngIndex

ISAAC_Prng_CleanUp:
    ' An error occurred or user opted to STOP processing
    If Not gblnStopProcessing Then

        If blnReturnFloat Then
            ISAAC_Prng = adblData()  ' Return float values
        Else
            ISAAC_Prng = alngData()  ' Return long integer values
        End If
    End If
    
    On Error GoTo 0   ' nullify this error trap
    Exit Function

ISAAC_Prng_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume ISAAC_Prng_CleanUp

End Function



' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

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

' void randinit(flag)
Private Sub RandomInit(ByVal blnUpdateSeed As Boolean)

    ' word flag;
    ' word i;
    ' ub4 a,b,c,d,e,f,g,h;
    Dim lngIndex As Long

    ' 08-Oct-2011 Changed these variables to STATIC to
    '             hold their value between calls
    Static AA As Long
    Static BB As Long
    Static CC As Long
    Static DD As Long
    Static EE As Long
    Static FF As Long
    Static GG As Long
    Static HH As Long
      
    If blnUpdateSeed Then
        
        ' aa=bb=cc=0;
        CreateSeed   ' Use seed values, not zeroes

        '    ' a=b=c=d=e=f=g=h=0x9e3779b9;  /* the golden ratio */
        '    AA = &H9E3779B9
        '    BB = &H9E3779B9
        '    CC = &H9E3779B9
        '    DD = &H9E3779B9
        '    EE = &H9E3779B9
        '    FF = &H9E3779B9
        '    GG = &H9E3779B9
        '    HH = &H9E3779B9
        '
        '    ' for (i=0; i<4; ++i)
        '    '    mix(a,b,c,d,e,f,g,h);
        '    For lngIndex = 1 To 4
        '        Mix AA, BB, CC, DD, EE, FF, GG, HH
        '    Next lngIndex
        '
        ' The above is a repeatable process
        ' with the same results each time
        AA = &H1367DF5A    '  325574490
        BB = &H95D90059    ' -1780940711
        CC = &HC3163E4B    ' -1021952437
        DD = &HF421AD8     '  255990488
        EE = &HD92A4A78    ' -651539848
        FF = &HA51A3C49    ' -1525007287
        GG = &HC4EFEA1B    ' -990909925
        HH = &H30609119    '  811634969
    End If
    
    ' for (i=0; i<256; i+=8)   /* fill in mm[] with messy stuff */
    For lngIndex = 0 To (MAX_RAND - 1) Step 8

        ' if (flag)  /* use all the information in the seed */
        If blnUpdateSeed Then
            ' a+=randrsl[i  ]; b+=randrsl[i+1]; c+=randrsl[i+2]; d+=randrsl[i+3];
            ' e+=randrsl[i+4]; f+=randrsl[i+5]; g+=randrsl[i+6]; h+=randrsl[i+7];
            AA = UnsignedAdd(AA, malngRand(lngIndex))
            BB = UnsignedAdd(BB, malngRand(lngIndex + 1))
            CC = UnsignedAdd(CC, malngRand(lngIndex + 2))
            DD = UnsignedAdd(DD, malngRand(lngIndex + 3))
            EE = UnsignedAdd(EE, malngRand(lngIndex + 4))
            FF = UnsignedAdd(FF, malngRand(lngIndex + 5))
            GG = UnsignedAdd(GG, malngRand(lngIndex + 6))
            HH = UnsignedAdd(HH, malngRand(lngIndex + 7))
        
        End If

        ' mix(a,b,c,d,e,f,g,h);
        Mix AA, BB, CC, DD, EE, FF, GG, HH

        ' mm[i  ]=a; mm[i+1]=b; mm[i+2]=c; mm[i+3]=d;
        ' mm[i+4]=e; mm[i+5]=f; mm[i+6]=g; mm[i+7]=h;
        malngMem(lngIndex) = AA
        malngMem(lngIndex + 1) = BB
        malngMem(lngIndex + 2) = CC
        malngMem(lngIndex + 3) = DD
        malngMem(lngIndex + 4) = EE
        malngMem(lngIndex + 5) = FF
        malngMem(lngIndex + 6) = GG
        malngMem(lngIndex + 7) = HH
        
    Next lngIndex

    ' if (flag)
    If blnUpdateSeed Then
        ' Do a second pass to make all of the seed.
        ' Will affect all of malngMem()
        '
        ' for (i=0; i<256; i+=8)
        For lngIndex = 0 To (MAX_RAND - 1) Step 8
            
            ' a+=mm[i  ]; b+=mm[i+1]; c+=mm[i+2]; d+=mm[i+3];
            ' e+=mm[i+4]; f+=mm[i+5]; g+=mm[i+6]; h+=mm[i+7];
            AA = UnsignedAdd(AA, malngMem(lngIndex))
            BB = UnsignedAdd(BB, malngMem(lngIndex + 1))
            CC = UnsignedAdd(CC, malngMem(lngIndex + 2))
            DD = UnsignedAdd(DD, malngMem(lngIndex + 3))
            EE = UnsignedAdd(EE, malngMem(lngIndex + 4))
            FF = UnsignedAdd(FF, malngMem(lngIndex + 5))
            GG = UnsignedAdd(GG, malngMem(lngIndex + 6))
            HH = UnsignedAdd(HH, malngMem(lngIndex + 7))

            ' mix(a,b,c,d,e,f,g,h);
            Mix AA, BB, CC, DD, EE, FF, GG, HH

            ' mm[i  ]=a; mm[i+1]=b; mm[i+2]=c; mm[i+3]=d;
            ' mm[i+4]=e; mm[i+5]=f; mm[i+6]=g; mm[i+7]=h;
            malngMem(lngIndex) = AA
            malngMem(lngIndex + 1) = BB
            malngMem(lngIndex + 2) = CC
            malngMem(lngIndex + 3) = DD
            malngMem(lngIndex + 4) = EE
            malngMem(lngIndex + 5) = FF
            malngMem(lngIndex + 6) = GG
            malngMem(lngIndex + 7) = HH
        
        Next lngIndex

        ISAAC_Calc   ' Load random data array [malngRand()]
        
    End If

End Sub

' #define mix(a,b,c,d,e,f,g,h) \
Private Sub Mix(ByRef AA As Long, _
                ByRef BB As Long, _
                ByRef CC As Long, _
                ByRef DD As Long, _
                ByRef EE As Long, _
                ByRef FF As Long, _
                ByRef GG As Long, _
                ByRef HH As Long)

    ' a^=b<<11; d+=a; b+=c;
    ' b^=c>>2;  e+=b; c+=d;
    ' c^=d<<8;  f+=c; d+=e;
    ' d^=e>>16; g+=d; e+=f;
    ' e^=f<<10; h+=e; f+=g;
    ' f^=g>>4;  a+=f; g+=h;
    ' g^=h<<8;  b+=g; h+=a;
    ' h^=a>>9;  c+=h; a+=b;
    AA = AA Xor (UnsignedMultiply(BB, SH_11))
    DD = UnsignedAdd(DD, AA)
    BB = UnsignedAdd(BB, CC)
    
    BB = BB Xor (UnsignedDivide(CC, SH_2))
    EE = UnsignedAdd(EE, BB)
    CC = UnsignedAdd(CC, DD)
    
    CC = CC Xor (UnsignedMultiply(DD, SH_8))
    FF = UnsignedAdd(FF, CC)
    DD = UnsignedAdd(DD, EE)
    
    DD = DD Xor (UnsignedDivide(EE, SH_16))
    GG = UnsignedAdd(GG, DD)
    EE = UnsignedAdd(EE, FF)
    
    EE = EE Xor (UnsignedMultiply(FF, SH_10))
    HH = UnsignedAdd(HH, EE)
    FF = UnsignedAdd(FF, GG)
    
    FF = FF Xor (UnsignedDivide(GG, SH_4))
    AA = UnsignedAdd(AA, FF)
    GG = UnsignedAdd(GG, HH)
    
    GG = GG Xor (UnsignedMultiply(HH, SH_8))
    BB = UnsignedAdd(BB, GG)
    HH = UnsignedAdd(HH, AA)
    
    HH = HH Xor (UnsignedDivide(AA, SH_9))
    CC = UnsignedAdd(CC, HH)
    AA = UnsignedAdd(AA, BB)

End Sub

' void isaac()
Private Sub ISAAC_Calc()

    ' register ub4 i;
    Dim lngAA    As Long
    Dim lngIndex As Long

    '*******************
    ' malngSeed(0) = aa
    ' malngSeed(1) = bb
    ' malngSeed(2) = cc
    '*******************
        
    ' cc=cc+1
    ' bb=bb+cc
    malngSeed(2) = UnsignedAdd(malngSeed(2), 1)             ' Just gets incremented once per 256 results
    malngSeed(1) = UnsignedAdd(malngSeed(1), malngSeed(2))  ' Then added to malngSeed(1)

        ' for (i=0; i<256; ++i)
        '
        '   switch (i%4)                   ----------------------------------+
        '      case 0: aa = aa^(aa<<13);     The shift portion is performed  |
        '      case 1: aa = aa^(aa>>6);      in within this routine. The Xor |
        '      case 2: aa = aa^(aa<<2);      portion is performed in the     |
        '      case 3: aa = aa^(aa>>16);     CommonCalcs() routine.          |
        '                                  ----------------------------------+
        '
        '   x = mm[i];                             -----------------------------+
        '   aa = mm[(i+128)%256] + aa;               This portion is performed  |
        '   mm[i] = y = mm[(x>>2)%256] + aa + bb;    in CommonCalcs() routine   |
        '   randrsl[i] = bb = mm[(y>>10)%256] + x;                              |
        '                                          -----------------------------+
    ' I tried doing this in bigger chunks but
    ' found no additional increase in speed
    For lngIndex = 0 To (MAX_RAND - 1) Step 16
    
        lngAA = (UnsignedMultiply(malngSeed(0), SH_13))  ' aa = aa^(aa<<13)
        CommonCalcs lngIndex, lngAA
    
        lngAA = (UnsignedDivide(malngSeed(0), SH_6))     ' aa = aa^(aa>>6)
        CommonCalcs (lngIndex + 1), lngAA
    
        lngAA = (UnsignedMultiply(malngSeed(0), SH_2))   ' aa = aa^(aa<<2)
        CommonCalcs (lngIndex + 2), lngAA
    
        lngAA = (UnsignedDivide(malngSeed(0), SH_16))    ' aa = aa^(aa>>16)
        CommonCalcs (lngIndex + 3), lngAA
    
        lngAA = (UnsignedMultiply(malngSeed(0), SH_13))  ' aa = aa^(aa<<13)
        CommonCalcs (lngIndex + 4), lngAA
    
        lngAA = (UnsignedDivide(malngSeed(0), SH_6))     ' aa = aa^(aa>>6)
        CommonCalcs (lngIndex + 5), lngAA
    
        lngAA = (UnsignedMultiply(malngSeed(0), SH_2))   ' aa = aa^(aa<<2)
        CommonCalcs (lngIndex + 6), lngAA
    
        lngAA = (UnsignedDivide(malngSeed(0), SH_16))    ' aa = aa^(aa>>16)
        CommonCalcs (lngIndex + 7), lngAA
    
        lngAA = (UnsignedMultiply(malngSeed(0), SH_13))  ' aa = aa^(aa<<13)
        CommonCalcs (lngIndex + 8), lngAA
    
        lngAA = (UnsignedDivide(malngSeed(0), SH_6))     ' aa = aa^(aa>>6)
        CommonCalcs (lngIndex + 9), lngAA
    
        lngAA = (UnsignedMultiply(malngSeed(0), SH_2))   ' aa = aa^(aa<<2)
        CommonCalcs (lngIndex + 10), lngAA
    
        lngAA = (UnsignedDivide(malngSeed(0), SH_16))    ' aa = aa^(aa>>16)
        CommonCalcs (lngIndex + 11), lngAA
    
        lngAA = (UnsignedMultiply(malngSeed(0), SH_13))  ' aa = aa^(aa<<13)
        CommonCalcs (lngIndex + 12), lngAA
    
        lngAA = (UnsignedDivide(malngSeed(0), SH_6))     ' aa = aa^(aa>>6)
        CommonCalcs (lngIndex + 13), lngAA
    
        lngAA = (UnsignedMultiply(malngSeed(0), SH_2))   ' aa = aa^(aa<<2)
        CommonCalcs (lngIndex + 14), lngAA
    
        lngAA = (UnsignedDivide(malngSeed(0), SH_16))    ' aa = aa^(aa>>16)
        CommonCalcs (lngIndex + 15), lngAA
    
    Next lngIndex

End Sub

Private Sub CommonCalcs(ByVal lngIndex As Long, _
                        ByVal lngAA As Long)

    ' Called by ISAAC_Calc()
    
    '*******************
    ' malngSeed(0) = aa
    ' malngSeed(1) = bb
    ' malngSeed(2) = cc
    '*******************
    
    '
    ' switch (i%4)                   ----------------------------------+
    '    case 0: aa = aa^(aa<<13);                                     |
    '    case 1: aa = aa^(aa>>6);      The shift portion was performed |
    '    case 2: aa = aa^(aa<<2);      in ISAAC_Calc() routine         |
    '    case 3: aa = aa^(aa>>16);                                     |
    '                                ----------------------------------+
    ' x = mm[i];
    ' aa = mm[(i+128)%256] + aa;
    ' mm[i] = y = mm[(x>>2)%256] + aa + bb;
    ' randrsl[i] = bb = mm[(y>>10)%256] + x;
        
    ' register ub4 x,y;
    Dim lngTmpX As Long
    Dim lngTmpY As Long
        
    ' Xor portion of the switch
    ' statement performed here
    malngSeed(0) = malngSeed(0) Xor lngAA
    
    ' x = mm[i];
    lngTmpX = malngMem(lngIndex)
    
    ' aa = mm[(i+128)%256] + aa;
    malngSeed(0) = UnsignedAdd(malngMem((lngIndex + 128) Mod MAX_RAND), malngSeed(0))
    
    ' mm[i] = y = mm[(x>>2)%256] + aa + bb;    [Shift right]
    lngTmpY = malngMem(Abs(UnsignedDivide(lngTmpX, SH_2)) Mod MAX_RAND)
    lngTmpY = UnsignedAdd(lngTmpY, malngSeed(0))
    lngTmpY = UnsignedAdd(lngTmpY, malngSeed(1))
    malngMem(lngIndex) = lngTmpY
    
    ' randrsl[i] = bb = mm[(y>>10)%256] + x;   [Shift right]
    lngTmpY = malngMem(Abs(UnsignedDivide(lngTmpY, SH_10)) Mod MAX_RAND)
    malngSeed(1) = UnsignedAdd(lngTmpX, lngTmpY)
    
    malngRand(lngIndex) = malngSeed(1)   ' Save one random value

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
' 20-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Sub CreateSeed()

    Dim lngLoop  As Long
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
    
    ReDim malngSeed(3)  ' Size seed array
    
    For lngLoop = 1 To 2
        For lngIndex = 0 To 2

            glngCarryOver = UnsignedMultiply(glngCarryOver, 29943829) - 1
            malngSeed(lngIndex) = UnsignedMultiply(glngCarryOver, 65536)
            glngCarryOver = UnsignedAdd(glngCarryOver, malngSeed(lngIndex))
            
        Next lngIndex
    Next lngLoop
    
    malngSeed(2) = Abs(malngSeed(2))                   ' make seed value positive
    malngSeed(2) = malngSeed(2) - (malngSeed(2) \ 3)   ' Subtract approx 33%
    
    mlngRndCnt = 0   ' Reset counter to zero

End Sub

Private Sub ResetVariables()

    Erase malngMem()   ' Empty module level arrays
    Erase malngRand()
    Erase malngSeed()
    
End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objIsaac = new cIsaac
                
    ResetVariables                            ' Always start with empty arrays
    mblnFirstTime = True                      ' Flag designates first time into this class
    mlngRndCnt = RND_COUNT + 1                ' Force a reseed.  See Isaac_PRNG()
    QueryPerformanceFrequency mcurFrequency   ' Test for Hi-Performance timer
        
    ReDim malngMem(MAX_RAND)     ' Size arrays
    ReDim malngRand(MAX_RAND)

    ' Preset property values
    StopProcessing = False
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objIsaac = Nothing
    
    mcurFrequency = 0@
    ResetVariables

End Sub

