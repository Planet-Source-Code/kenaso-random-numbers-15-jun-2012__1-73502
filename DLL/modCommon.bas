Attribute VB_Name = "modCommon"
' ***************************************************************************
' Routine:       modCommon
'
' Description:   Common routines called by classes.
'
'    Thanks to RDE for his VB addin so I could track what calls what.
'    Callers Add-in [FINAL 2.24 - Sept 12, 2011]
'    http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73617&lngWId=1
'
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MAX_LONG As Long = &H7FFFFFFF      '  2147483647
  Private Const MIN_LONG As Long = &H80000000      ' -2147483648
  Private Const GB_4     As Double = (2# ^ 32)     '  4294967296

' ***************************************************************************
' Global Constants
' ***************************************************************************
  Public Const DLL_NAME As String = "kiPrng"
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' Retrieves the frequency of the high-resolution performance counter,
  ' if one exists. The frequency cannot change while the system is running.
  ' If the function fails, the return value is zero.
  Public Declare Function QueryPerformanceFrequency Lib "kernel32" _
         (curFrequency As Currency) As Long
  
  ' The QueryPerformanceCounter function retrieves the current value of the
  ' high-resolution performance counter.
  Public Declare Function QueryPerformanceCounter Lib "kernel32" _
         (curCounter As Currency) As Boolean

  ' This is a rough translation of the GetTickCount API. The
  ' tick count of a PC is only valid for the first 49.7 days
  ' since the last reboot.  When you capture the tick count,
  ' you are capturing the total number of milliseconds elapsed
  ' since the last reboot.  The elapsed time is stored as a
  ' DWORD value. Therefore, the time will wrap around to zero
  ' if the system is run continuously for 49.7 days.
  Public Declare Function GetTickCount Lib "kernel32" () As Long
  
' ***************************************************************************
' Global Variables
'
'                    +-------------- Global level designator
'                    |  +----------- Data type (Long Integer)
'                    |  |     |----- Variable subname
'                    - --- -------------
' Naming standard:   g lng CarryOver
' Variable name:     glngCarryOver
'
' ***************************************************************************
  Public glngCarryOver As Long
  
  
' ***************************************************************************
' Routine:       RotateLong
'
' Description:   Rotate (sometimes called a circular shift) a Long Integer
'                to the left or right a specified number of bits.
'                Equivalent to "<<<" and ">>>"
'
' Parameters:    lngValue - Number value being manipulated
'                lngBitShift - Number of bits to be manipulated.
'                    If bit value is positive then rotate left
'                    else rotate right.
'
' Returns:       Numeric value after bit manipulation
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-Jun-1999  Francesco Balena
'              Rotate left  - http://www.devx.com/vb2themax/Tip/18955
'              Rotate right - http://www.devx.com/vb2themax/Tip/18957
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function RotateLong(ByVal lngValue As Long, _
                           ByVal lngBitShift As Long) As Long

    ' Called by KISS.CalcLong()
    '           Mother.CalcLong()
    '           MWC.CalcLong()
    
    Dim lngLoop    As Long
    Dim lngSignBit As Long
    
    If lngBitShift = 0 Then
        RotateLong = lngValue   ' Nothing to do
    ElseIf lngBitShift > 31 Then
        RotateLong = 0          ' Excessive positive bit positions
    ElseIf lngBitShift < -31 Then
        RotateLong = 0          ' Excessive negative bit positions

    ' Positive bit value means rotate left
    ElseIf lngBitShift > 0 Then

        For lngLoop = 1 To lngBitShift
        
            ' remember the two most significant bits
            lngSignBit = lngValue And &HC0000000
            
            ' clear the bit and shift left by one position
            lngValue = (lngValue And &H3FFFFFFF) * 2
            
            ' if number was negative, then add 1
            ' if bit 30 was set, then set the sign bit
            lngValue = lngValue Or _
                       ((lngSignBit < 0) And &H1&) Or _
                       (CBool(lngSignBit And &H40000000) And &H80000000)
        Next lngLoop
        
        RotateLong = lngValue
                
    ' Negative bit value means rotate right
    ElseIf lngBitShift < 0 Then
    
        ' make bit value positive
        lngBitShift = Abs(lngBitShift)
        
        For lngLoop = 1 To lngBitShift
        
            ' remember the sign bit and bit 0
            lngSignBit = lngValue And &H80000001
        
            ' clear the bit and shift right by one position
            lngValue = (lngValue And &H7FFFFFFE) \ 2
        
            ' if number was negative, then reinsert the bit
            ' if bit 0 was set, then set the sign bit
            lngValue = lngValue Or _
                       ((lngSignBit < 0) And &H40000000) Or _
                       (CBool(lngSignBit And 1) And &H80000000)
        Next lngLoop
        
        RotateLong = lngValue
    
    End If

End Function

' ***************************************************************************
' Routine:       UnsignedAdd
'
' Description:   Function to add two unsigned numbers together as in C.
'                Overflows are ignored!
'
' Parameters:    dblValue1 - Value of A
'                dblValue2 - Value of B
'
' Returns:       Calculated value
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
Public Function UnsignedAdd(ByVal dblValue1 As Double, _
                            ByVal dblValue2 As Double) As Long
                             
    ' Called by Isaac.RandomInit()
    '           Isaac.Mix()
    '           Isaac.Calc()
    '           Isaac.CommonCalcs()
    '           KISS.CalcLong()
    '           Mother.CalcLong()
    '           Mother.RandomInit()
    '           Mother.CreateSeed()
    '           MWC.CalcLong()
    
    Dim dblTemp As Double

    dblTemp = dblValue1 + dblValue2

    If dblTemp < MIN_LONG Then
        UnsignedAdd = CLng(GB_4 + dblTemp)
    Else
        If dblTemp > MAX_LONG Then
            UnsignedAdd = CLng(dblTemp - GB_4)
        Else
            UnsignedAdd = CLng(dblTemp)
        End If
    End If

End Function

' ***************************************************************************
' Routine:       UnsignedMultiply
'
' Description:   Multiplies the two (signed) Long parameters, treated as
'                unsigned long, and returns the lowest 4 bytes of the 8 bytes
'                result as a (signed) Long result.  Overflows are ignored.
'
'                This function emulates the multiplication of two (unsigned
'                long) numbers, and is needed because the type Double has
'                only a 53 bits mantissa and the result of multiplying 2 Long
'                variables might need 64 bits to accurately represent the
'                result in some cases.
'
'                lngValue1   == ABCD  == A000 + B00 + C0 + D
'                lngValue2   == EFGH  == E000 + F00 + G0 + H
'
'         Note:  In the following, "ae" means the 2 bytes result of A*E,
'                "bg" of B*G, etc:
'
'                lngValue1 * lngValue2 == ae 000 000 +  ' discard, result is too high
'                       af 000  00 +     ' discard, result is too high
'                       ag 000   0 +     ' discard, result is too high
'                       ah 000     +     ' take lowest byte
'
'                       be  00 000 +     ' discard, result is too high
'                       bf  00  00 +     ' discard, result is too high
'                       bg  00   0 +     ' take lowest byte
'                       bh  00     +     ' take both bytes
'
'                       ce   0 000 +     ' discard, result is too high
'                       cf   0  00 +     ' take lowest byte
'                       cg   0   0 +     ' take both bytes
'                       ch   0     +     ' take both bytes
'
'                       de     000 +     ' take lowest byte
'                       df      00 +     ' take both bytes
'                       dg       0 +     ' take both bytes
'                       dh               ' take both bytes
'
' Parameters:   lngValue1 - Number that gets multiplied (Multiplicand)
'               lngValue2 - Number doing the multiplying (Multiplier)
'
' Returns:      New calculated value (Product)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 01-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Renamed variables
' ***************************************************************************
Public Function UnsignedMultiply(ByVal lngValue1 As Long, _
                                 ByVal lngValue2 As Long) As Long

    ' Called by Isaac.ISAAC_Calc()
    '           Isaac.CreateSeed()
    '           Isaac.Mix()
    '           KISS.CalcLong()
    '           Kiss.CreateSeed()
    '           Mother.CalcLong()
    '           Mother.RandomInit()
    '           Mother.CreateSeed()
    '           MT11231A.CalcOneLong()
    '           MT11231A.CreateSeed()
    '           MT11231B.CalcOneLong()
    '           MT11231B.CreateSeed()
    '           MT19937.CalcOneLong()
    '           MT19937.CreateSeed()
    '           TT800.CalcOneLong()
    '           TT800.CreateSeed()
    '           MWC.CalcLong()
    '           MWC.CreateSeed()
    
    Dim BB      As Long
    Dim CC      As Long
    Dim DD      As Long
    Dim FF      As Long
    Dim GG      As Long
    Dim HH      As Long
    Dim R0      As Long
    Dim R1      As Long
    Dim R2      As Long
    Dim R3      As Long
    Dim dblTemp As Double
    
    Const k2_8  As Long = 256
    Const k2_24 As Long = &H1000000   ' 16777216
    Const KB_64 As Long = 65536
   
    'x==ABCD, y==EFGH
    BB = (lngValue1 \ KB_64) Mod k2_8
    CC = (lngValue1 \ k2_8) Mod k2_8
    DD = lngValue1 Mod k2_8
    FF = (lngValue2 \ KB_64) Mod k2_8
    GG = (lngValue2 \ k2_8) Mod k2_8
    HH = lngValue2 Mod k2_8

    ' get the 1st (lowest) byte of the result, R0:
    '       dh             'take both bytes
    R0 = DD * HH

    ' get the 2nd byte of the result, R1, and add carry from R0:
    '       ch   0      +  'take both bytes
    '       dg        0    'take both bytes
    R1 = CC * HH + DD * GG + R0 \ k2_8

    ' get the 3rd byte of the result, R2, and add carry from R1:
    '       bh  00      +  'take both bytes
    '       cg   0    0 +  'take both bytes
    '       df       00    'take both bytes
    R2 = BB * HH + CC * GG + DD * FF + R1 \ k2_8

    ' get the 4th (highest) byte of the result, R3, and add carry from R2:
    '       ah 000      +  'take lowest byte
    '       bg  00    0 +  'take lowest byte
    '       cf   0   00 +  'take lowest byte
    '       de      000    'take lowest byte
    R3 = (((lngValue1 \ k2_24) * HH + BB * GG + CC * FF + DD * (lngValue2 \ k2_24)) Mod k2_8) + R2 \ k2_8
    dblTemp = CDbl(R3 Mod k2_8) * k2_24 + (R2 Mod k2_8) * KB_64 + (R1 Mod k2_8) * k2_8 + (R0 Mod k2_8)

    ' now we have a 32 bits number (dblTemp) that can be processed
    ' without losing precision using the 53 bits mantissa of
    ' the Double type
    If dblTemp < MIN_LONG Then
        UnsignedMultiply = CLng(GB_4 + dblTemp)
    Else
        If dblTemp > MAX_LONG Then
            UnsignedMultiply = CLng(dblTemp - GB_4)
        Else
            UnsignedMultiply = CLng(dblTemp)
        End If
    End If

End Function

' ***************************************************************************
' Routine:       UnsignedDivide
'
' Description:   Divides the two (signed) Long parameters, treated as
'                unsigned long, and returns the result as a (signed)
'                Long result.
'
' Parameters:    lngDividend - Number to be divided (Dividend)
'                lngDivisor - Number performing division (Divisor)
'
' Returns:       New calculated value (Quotient)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 01-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function UnsignedDivide(ByVal lngDividend As Long, _
                               ByVal lngDivisor As Long) As Long
                                
    ' Called by Isaac.ISAAC_Calc()
    '           Isaac.CommonCalcs()
    '           Isaac.CreateSeed()
    '           Isaac.Mix()
    '           KISS.CalcLong()
    '           Kiss.CreateSeed()
    '           Mother.RandomInit()
    '           MT11231A.CalcOneLong()
    '           MT11231A.CreateSeed()
    '           MT11231B.CalcOneLong()
    '           MT11231B.CreateSeed()
    '           MT19937.CalcOneLong()
    '           MT19937.CreateSeed()
    '           TT800.CalcOneLong()
    '           TT800.CreateSeed()
    '           MWC.CalcLong()
    '           MWC.CreateSeed()
    
    If lngDividend < 0 Then
        UnsignedDivide = (Abs(Fix((GB_4 + lngDividend) / lngDivisor)))
    Else
        UnsignedDivide = CLng(Fix(lngDividend / lngDivisor))
    End If

End Function

