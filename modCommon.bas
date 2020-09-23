Attribute VB_Name = "modCommon"
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MAX_LONG  As Long = &H7FFFFFFF      '  2147483647
  Private Const MIN_LONG  As Long = &H80000000      ' -2147483648
  Private Const GB_1      As Double = (2 ^ 30)      '  1073741824
  Private Const MAX_DWORD As Double = (2# ^ 32) - 1 '  4294967295
  Private Const GB_4      As Double = (2# ^ 32)     '  4294967296

' ***************************************************************************
' Routine:       ShiftLong
'
' Description:   Shifts the bits to the right or left the specified number of
'                positions and returns the new value.  Bits "falling off"
'                the edge do not wrap around.  The fill bits are zeroes on
'                the opposite side.  Some common languages like C/C++ or
'                Java have an operator for this job: ">>" or "<<".
'
'                intBitShift is a switch denoting either a left (positive)
'                or right (negative) shift.
'
' Parameters:    lngValue     - Number to be manipulated
'                intBitShift  - number of shift positions
'                               Positive value = left shift
'                               Negative value = right shift
'
' Returns:       Reformatted value
'
'                  Number                Binary
' Original:      123456789   00000111010110111100110100010101
' Left 5:       -344350048   10010100100001100101110101100000
' Right 5:         3858024   00000000001110101101111001101000
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-May-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote rouine
' ***************************************************************************
Public Function ShiftLong(ByVal lngValue As Long, _
                            ByVal intBitShift As Integer) As Long

    Dim lngMask    As Long
    Dim lngSignBit As Long

    ' Test bit shifting ranges
    Select Case intBitShift
           Case 0:        ShiftLong = lngValue   ' no bit shifting requested
           Case Is < -31: ShiftLong = 0          ' return zero if too many shift positions
           Case Is > 31:  ShiftLong = 0          ' return zero if too many shift positions

           ' Shift to the left if a positive bit count
           Case Is > 0
                ShiftLong = (lngValue And (2 ^ (31 - intBitShift) - 1)) * (2 ^ intBitShift)

                If (lngValue And (2 ^ (31 - intBitShift))) = (2 ^ (31 - intBitShift)) Then
                    ShiftLong = (ShiftLong Or &H80000000)
                End If

           ' Shift to the right if a negative bit count
           Case Is < 0
                intBitShift = Abs(intBitShift) ' Make flag positive

                If intBitShift = 31 Then

                    If lngValue < 0 Then
                        ShiftLong = &HFFFFFFFF
                    Else
                        ShiftLong = 0
                    End If

                Else
                    ShiftLong = (lngValue And Not (2 ^ intBitShift - 1)) \ (2 ^ intBitShift)
                End If
    End Select

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
                                
    If lngDividend < 0 Then
        UnsignedDivide = (Abs(Fix((GB_4 + lngDividend) / lngDivisor)))
    Else
        UnsignedDivide = CLng(Fix(lngDividend / lngDivisor))
    End If

End Function


