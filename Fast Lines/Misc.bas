Attribute VB_Name = "Misc"
Option Explicit

Public Const OFFSET_DOUBLE_8_BYTE = 3.59538626972464E+307
Public Const OFFSET_FLOAT_4_BYTE = 6.805646E+38
Public Const OFFSET_INTEGER_4_BYTE = 4294967296#
Public Const OFFSET_INTEGER_2_BYTE = 65536
Public Const OFFSET_INTEGER_1_BYTE = 256

Public Const MAX_DOUBLE_8_BYTE = 1.79769313486232E+307
Public Const MAX_FLOAT_4_BYTE = 3.402823E+38
Public Const MAX_INTEGER_4_BYTE = 2147483648#
Public Const MAX_INTEGER_2_BYTE = 32768
Public Const MAX_INTEGER_1_BYTE = 128

Public Red_Look_Up_Table_15(255) As Long
Public Green_Look_Up_Table_15(255) As Long
Public Blue_Look_Up_Table_15(255) As Long

Public Red_Look_Up_Table_16(255) As Long
Public Green_Look_Up_Table_16(255) As Long
Public Blue_Look_Up_Table_16(255) As Long

Public Red_Look_Up_Table_24(255) As Long
Public Green_Look_Up_Table_24(255) As Long
Public Blue_Look_Up_Table_24(255) As Long

Public Red_Look_Up_Table_32(255) As Long
Public Green_Look_Up_Table_32(255) As Long
Public Blue_Look_Up_Table_32(255) As Long
Public Alpha_Look_Up_Table_32(255)

Public Sub Build_Color_Look_Up_Tables()

    Dim Current_Red As Long
    Dim Current_Green As Long
    Dim Current_Blue As Long
    Dim Current_Alpha As Long
    
    Dim Temp

    For Current_Red = 0 To 255
        
        Temp = 1024
        Red_Look_Up_Table_15(Current_Red) = ((Current_Red And 31) * Temp)
        
        Temp = 2048
        Red_Look_Up_Table_16(Current_Red) = ((Current_Red And 31) * Temp)
        
        Red_Look_Up_Table_24(Current_Red) = (Current_Red * 65536)
        Red_Look_Up_Table_32(Current_Red) = (Current_Red * 65536)
        
    Next Current_Red
    
    For Current_Green = 0 To 255
    
        Green_Look_Up_Table_15(Current_Green) = ((Current_Green And 31) * 32)
        Green_Look_Up_Table_16(Current_Green) = ((Current_Green And 63) * 32)
        Green_Look_Up_Table_24(Current_Green) = (Current_Green * 256)
        Green_Look_Up_Table_32(Current_Green) = (Current_Green * 256)
        
    Next Current_Green
    
    For Current_Blue = 0 To 255
    
        Blue_Look_Up_Table_15(Current_Blue) = (Current_Blue And 31)
        Blue_Look_Up_Table_16(Current_Blue) = (Current_Blue And 31)
        Blue_Look_Up_Table_24(Current_Blue) = (Current_Blue)
        Blue_Look_Up_Table_32(Current_Blue) = (Current_Blue)
        
    Next Current_Blue
    
    For Current_Alpha = 0 To 255
        
        Temp = 16777216
        Alpha_Look_Up_Table_32(Current_Alpha) = (Current_Alpha * Temp)
    
    Next Current_Alpha

End Sub

Public Function Double_Range(ByVal Value) As Double

    '-1.79769313486232E+307 to 1.79769313486232E+307

    Dim Number_Of_Times_Over
    
    Number_Of_Times_Over = Int(Value / MAX_DOUBLE_8_BYTE)

    If (Number_Of_Times_Over Mod 2 = 1) Or (Number_Of_Times_Over Mod 2 = -1) Then
    
        Double_Range = (Value - (OFFSET_DOUBLE_8_BYTE * Number_Of_Times_Over / 2)) - MAX_DOUBLE_8_BYTE
        
    ElseIf (Number_Of_Times_Over Mod 2 = 0) Then
    
        Double_Range = (Value - (OFFSET_DOUBLE_8_BYTE * Number_Of_Times_Over / 2))
    
    End If

End Function

Public Function Single_Range(ByVal Value) As Single

    '-63.402823E+38 to 3.402823E+38

    Dim Number_Of_Times_Over
    
    Number_Of_Times_Over = Int(Value / MAX_FLOAT_4_BYTE)

    If (Number_Of_Times_Over Mod 2 = 1) Or (Number_Of_Times_Over Mod 2 = -1) Then
    
        Single_Range = (Value - (OFFSET_FLOAT_4_BYTE * Number_Of_Times_Over / 2)) - MAX_FLOAT_4_BYTE
        
    ElseIf (Number_Of_Times_Over Mod 2 = 0) Then
    
        Single_Range = (Value - (OFFSET_FLOAT_4_BYTE * Number_Of_Times_Over / 2))
    
    End If

End Function

Public Function Long_Range(ByVal Value) As Long

    '-2147483648 to 2147483648

    Dim Number_Of_Times_Over
    
    Value = Fix(Value)
    
    Number_Of_Times_Over = Int(Value / MAX_INTEGER_4_BYTE)

    If (Number_Of_Times_Over Mod 2 = 1) Or (Number_Of_Times_Over Mod 2 = -1) Then
    
        Long_Range = (Value - (OFFSET_INTEGER_4_BYTE * Number_Of_Times_Over / 2)) - MAX_INTEGER_4_BYTE
        
    ElseIf (Number_Of_Times_Over Mod 2 = 0) Then
    
        Long_Range = (Value - (OFFSET_INTEGER_4_BYTE * Number_Of_Times_Over / 2))
    
    End If

End Function

Public Function Integer_Range(ByVal Value) As Integer
    
    '-32768 to 32767
    
    Dim Number_Of_Times_Over
    
    Value = Fix(Value)
    
    Number_Of_Times_Over = Int(Value / MAX_INTEGER_2_BYTE)

    If (Number_Of_Times_Over Mod 2 = 1) Or (Number_Of_Times_Over Mod 2) = -1 Then
    
        Integer_Range = (Value - (OFFSET_INTEGER_2_BYTE * Number_Of_Times_Over / 2)) - MAX_INTEGER_2_BYTE
        
    ElseIf (Number_Of_Times_Over Mod 2 = 0) Then
    
        Integer_Range = (Value - (OFFSET_INTEGER_2_BYTE * Number_Of_Times_Over / 2))
    
    End If

End Function

Public Function Byte_Range(ByVal Value) As Integer
    
    'C++ Style -128 to 127
    
    Dim Number_Of_Times_Over
    
    Value = Fix(Value)
    
    Number_Of_Times_Over = Int(Value / MAX_INTEGER_1_BYTE)

    If (Number_Of_Times_Over Mod 2 = 1) Or (Number_Of_Times_Over Mod 2 = -1) Then
    
        Byte_Range = (Value - (OFFSET_INTEGER_1_BYTE * Number_Of_Times_Over / 2)) - MAX_INTEGER_1_BYTE
        
    ElseIf (Number_Of_Times_Over Mod 2 = 0) Then
    
        Byte_Range = (Value - (OFFSET_INTEGER_1_BYTE * Number_Of_Times_Over / 2))
    
    End If

End Function

Public Function Unsigned_Double(ByVal Value)

    Value = Double_Range(Value)

    If Value < 0 Then
    
        Unsigned_Double = Value + OFFSET_DOUBLE_8_BYTE
    
    Else
    
        Unsigned_Double = Value
  
    End If

End Function

Public Function Unsigned_Single(ByVal Value)

    Value = Single_Range(Value)

    If Value < 0 Then
    
        Unsigned_Single = Value + OFFSET_FLOAT_4_BYTE
    
    Else
    
        Unsigned_Single = Value
  
    End If

End Function

Public Function Unsigned_Long(ByVal Value)

    Value = Long_Range(Value)

    If Value < 0 Then
    
        Unsigned_Long = Value + OFFSET_INTEGER_4_BYTE
    
    Else
    
        Unsigned_Long = Value
  
    End If

End Function

Public Function Unsigned_Integer(ByVal Value) As Long
    
    Value = Integer_Range(Value)
    
    If Value < 0 Then
    
        Unsigned_Integer = Value + OFFSET_INTEGER_2_BYTE
    
    Else
    
        Unsigned_Integer = Value
    
    End If

End Function

Public Function Unsigned_Byte(ByVal Value) As Byte
    
    Value = Byte_Range(Value)
    
    If Value < 0 Then
    
        Unsigned_Byte = Value + OFFSET_INTEGER_1_BYTE
    
    Else
    
        Unsigned_Byte = Value
    
    End If

End Function

Public Function RGB_16_Bit_555(ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long) As Long
    
    '5 + 5 + 5 = 15 Bits
    
    RGB_16_Bit_555 = Integer_Range(Blue_Look_Up_Table_15(Blue) + Green_Look_Up_Table_15(Green) + Red_Look_Up_Table_15(Red))
    
'//////////////////////////////////////////////////////////

'USHORT RGB16Bit555(int r, int g, int b)
'{
'// this function simply builds a 5.5.5 format 16 bit pixel
'// assumes input is RGB 0-255 each channel
'r>>=3; g>>=3; b>>=3;
'return(_RGB16BIT555((r),(g),(b)));

'} // end RGB16Bit555

'//////////////////////////////////////////////////////////

End Function

Public Function RGB_16_Bit_565(ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long) As Integer
    
    '5 + 6 + 5 = 16 Bits
    
    RGB_16_Bit_565 = Integer_Range(Blue_Look_Up_Table_16(Blue) + Green_Look_Up_Table_16(Green) + Red_Look_Up_Table_16(Red))
    
End Function

Public Function RGB_24_Bit_888(ByVal Red, ByVal Green, ByVal Blue)
    
    '8 + 8 + 8 = 24 Bits

    RGB_24_Bit_888 = Blue_Look_Up_Table_24(Blue) + Green_Look_Up_Table_24(Green) + Red_Look_Up_Table_24(Red)
    
End Function

Public Function RGB_32_Bit_A888(ByVal Red, ByVal Green, ByVal Blue, ByVal Alpha)
    
    '8 + 8 + 8 + 8 = 32 Bits
    
    RGB_32_Bit_A888 = Blue_Look_Up_Table_32(Blue) + Green_Look_Up_Table_32(Green) + Red_Look_Up_Table_32(Red) + Alpha_Look_Up_Table_32(Alpha)

End Function

Public Function RGB_Color(ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, Optional ByVal Alpha As Long)
    
    'Supports almost all color modes.
    
    Select Case Pixel_Format
    
        Case COLOR_DEPTH_8_BIT
        
            'Palletized
        
        Case COLOR_DEPTH_15_BIT
        
            'Note: Never been tested, my card has no support
            '      for 15 bit color mode.
        
            'RGB_Color = RGB_16_Bit_555(Red, Green, Blue)
            
            'inline RGB_16_Bit_555
            RGB_Color = Integer_Range(Blue_Look_Up_Table_15(Blue) + Green_Look_Up_Table_15(Green) + Red_Look_Up_Table_15(Red))
        
        Case COLOR_DEPTH_16_BIT
        
            'RGB_Color = RGB_16_Bit_565(Red, Green, Blue)
            
            'inline RGB_16_Bit_565
            RGB_Color = Integer_Range(Blue_Look_Up_Table_16(Blue) + Green_Look_Up_Table_16(Green) + Red_Look_Up_Table_16(Red))
        
        Case COLOR_DEPTH_24_BIT
            
            'Note: Never been tested, my card has no support
            '      for 24 bit color mode.
            
            'RGB_Color = RGB_24_Bit_888(Red, Green, Blue)
            
            'inline RGB_24_Bit_888
            RGB_Color = Blue_Look_Up_Table_24(Blue) + Green_Look_Up_Table_24(Green) + Red_Look_Up_Table_24(Red)
        
        Case COLOR_DEPTH_32_BIT
        
            'RGB_Color = RGB_32_Bit_A888(Red, Green, Blue, Alpha)
            
            'inline RGB_32_Bit_A888
            RGB_Color = Blue_Look_Up_Table_32(Blue) + Green_Look_Up_Table_32(Green) + Red_Look_Up_Table_32(Red) + Alpha_Look_Up_Table_32(Alpha)
            
     End Select

End Function

Public Function String_Split(Expression As String, Optional ByVal Delimiter As String = " ", Optional ByVal Limit As Long = -1, Optional ByVal Compare As Long = 0) As Variant
    
    'Since I have Visual Basic 5, I have no support for
    'the Split() function which is new for Visual Basic 6.
    'So I created my own and it works. Need it for loading
    '8 bit palettes (text based).
    
    Dim Current_Position As Long

    Dim Space_Position As Long
    
    Dim Number_Of_Splited_Data As Long
    
    Dim Length As Long

    Dim String_Data() As String
    
    Current_Position = 1
    
    'Print "Expression = " & Expression
    'Print
    
    Do While Current_Position > 0 Or Limit = -1
    
        DoEvents
        
        Space_Position = InStr(Current_Position, Expression, Delimiter, Compare)

        If Space_Position Then
        
            Length = Space_Position - Current_Position
        
            'Print "Space_Position = " & Space_Position
            'Print "Current_Position = " & Current_Position
            'Print "Length = " & Length
        
            If Length <> 0 Then
        
                Number_Of_Splited_Data = Number_Of_Splited_Data + 1
        
                ReDim Preserve String_Data(Number_Of_Splited_Data - 1) As String
            
                String_Data(Number_Of_Splited_Data - 1) = Mid(Expression, Current_Position, Length)
            
                'Print String_Data(Number_Of_Splited_Data - 1)
            
            End If
            
            Current_Position = Space_Position
            
            'Print "Current_Temp_Position = " & Current_Position
            
            Current_Position = Current_Position + 1
            
        Else
            
            Length = Len(Mid(Expression, Current_Position, Len(Expression)))
            
            Number_Of_Splited_Data = Number_Of_Splited_Data + 1
        
            ReDim Preserve String_Data(Number_Of_Splited_Data - 1) As String
        
            String_Data(Number_Of_Splited_Data - 1) = Trim(Mid(Expression, Current_Position, Length))
        
            'Print String_Data(Number_Of_Splited_Data - 1)
            
            Exit Do
            
        End If
        
        'Print
        
        If Current_Position > Len(Expression) Then Exit Do
        
    Loop
    
    String_Split = String_Data

End Function
