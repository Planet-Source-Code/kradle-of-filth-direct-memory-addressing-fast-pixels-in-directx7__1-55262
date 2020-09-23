Attribute VB_Name = "DirectX7_Enigne"
Option Explicit

Public Const DX_KEY_ESC As Long = DIK_ESCAPE
Public Const DX_KEY_1 As Long = DIK_1
Public Const DX_KEY_2 As Long = DIK_2
Public Const DX_KEY_3 As Long = DIK_3
Public Const DX_KEY_4 As Long = DIK_4
Public Const DX_KEY_5 As Long = DIK_5
Public Const DX_KEY_6 As Long = DIK_6
Public Const DX_KEY_7 As Long = DIK_7
Public Const DX_KEY_8 As Long = DIK_8
Public Const DX_KEY_9 As Long = DIK_9

Public Const DX_KEY_0 As Long = DIK_0
Public Const DX_KEY_MINUS As Long = DIK_MINUS
Public Const DX_KEY_EQUALS As Long = DIK_EQUALS
Public Const DX_KEY_BACKSPACE As Long = DIK_BACK
Public Const DX_KEY_TAB As Long = DIK_TAB
Public Const DX_KEY_Q As Long = DIK_Q
Public Const DX_KEY_W As Long = DIK_W
Public Const DX_KEY_E As Long = DIK_E
Public Const DX_KEY_R As Long = DIK_R
Public Const DX_KEY_T As Long = DIK_T

Public Const DX_KEY_Y As Long = DIK_Y
Public Const DX_KEY_U As Long = DIK_U
Public Const DX_KEY_I As Long = DIK_I
Public Const DX_KEY_O As Long = DIK_O
Public Const DX_KEY_P As Long = DIK_P
Public Const DX_KEY_LEFT_BRACKET As Long = DIK_LBRACKET
Public Const DX_KEY_RIGHT_BRACKET As Long = DIK_RBRACKET
Public Const DX_KEY_ENTER As Long = DIK_RETURN
Public Const DX_KEY_LEFT_CTRL As Long = DIK_LCONTROL
Public Const DX_KEY_A As Long = DIK_A

Public Const DX_KEY_S As Long = DIK_S
Public Const DX_KEY_D As Long = DIK_D
Public Const DX_KEY_F As Long = DIK_F
Public Const DX_KEY_G As Long = DIK_G
Public Const DX_KEY_H As Long = DIK_H
Public Const DX_KEY_J As Long = DIK_J
Public Const DX_KEY_K As Long = DIK_K
Public Const DX_KEY_L As Long = DIK_L
Public Const DX_KEY_SEMICOLON As Long = DIK_SEMICOLON
Public Const DX_KEY_APOSTROPHE As Long = DIK_APOSTROPHE

Public Const DX_KEY_BACK_SINGLE_QUOTE As Long = DIK_GRAVE
Public Const DX_KEY_LEFT_SHIFT As Long = DIK_LSHIFT
Public Const DX_KEY_BACKSLASH As Long = DIK_BACKSLASH
Public Const DX_KEY_Z As Long = DIK_Z
Public Const DX_KEY_X As Long = DIK_X
Public Const DX_KEY_C As Long = DIK_C
Public Const DX_KEY_V As Long = DIK_V
Public Const DX_KEY_B As Long = DIK_B
Public Const DX_KEY_N As Long = DIK_N
Public Const DX_KEY_M As Long = DIK_M

Public Const DX_KEY_COMMA As Long = DIK_COMMA
Public Const DX_KEY_PERIOD As Long = DIK_PERIOD
Public Const DX_KEY_SLASH As Long = DIK_SLASH
Public Const DX_KEY_RIGHT_SHIFT As Long = DIK_RSHIFT
Public Const DX_KEY_NUMPAD_ASTERISK As Long = DIK_MULTIPLY
Public Const DX_KEY_LEFT_ALT As Long = DIK_LMENU
Public Const DX_KEY_SPACE As Long = DIK_SPACE
Public Const DX_KEY_CAPSLOCK As Long = DIK_CAPITAL
Public Const DX_KEY_F1 As Long = DIK_F1
Public Const DX_KEY_F2 As Long = DIK_F2

Public Const DX_KEY_F3 As Long = DIK_F3
Public Const DX_KEY_F4 As Long = DIK_F4
Public Const DX_KEY_F5 As Long = DIK_F5
Public Const DX_KEY_F6 As Long = DIK_F6
Public Const DX_KEY_F7 As Long = DIK_F7
Public Const DX_KEY_F8 As Long = DIK_F8
Public Const DX_KEY_F9 As Long = DIK_F9
Public Const DX_KEY_F10 As Long = DIK_F10
Public Const DX_KEY_NUMLOCK As Long = DIK_NUMLOCK
Public Const DX_KEY_SCROLLOCK As Long = DIK_SCROLL

Public Const DX_KEY_NUMPAD_7 As Long = DIK_NUMPAD7
Public Const DX_KEY_NUMPAD_8 As Long = DIK_NUMPAD8
Public Const DX_KEY_NUMPAD_9 As Long = DIK_NUMPAD9
Public Const DX_KEY_NUMPAD_DASH As Long = DIK_SUBTRACT
Public Const DX_KEY_NUMPAD_4 As Long = DIK_NUMPAD4
Public Const DX_KEY_NUMPAD_5 As Long = DIK_NUMPAD5
Public Const DX_KEY_NUMPAD_6 As Long = DIK_NUMPAD6
Public Const DX_KEY_NUMPAD_PLUS As Long = DIK_ADD
Public Const DX_KEY_NUMPAD_1 As Long = DIK_NUMPAD1
Public Const DX_KEY_NUMPAD_2 As Long = DIK_NUMPAD2

Public Const DX_KEY_NUMPAD_3 As Long = DIK_NUMPAD3
Public Const DX_KEY_NUMPAD_0 As Long = DIK_NUMPAD0
Public Const DX_KEY_NUMPAD_PERIOD As Long = DIK_DECIMAL
Public Const DX_KEY_F14 As Long = DIK_F14
Public Const DX_KEY_F15 As Long = DIK_F15
Public Const DX_KEY_F13 As Long = DIK_F13
Public Const DX_KEY_F11 As Long = DIK_F11
Public Const DX_KEY_F12 As Long = DIK_F12
Public Const DX_KEY_NUMPAD_COMMA As Long = DIK_NUMPADCOMMA
Public Const DX_KEY_NUMPAD_ENTER As Long = DIK_NUMPADENTER

Public Const DX_KEY_RIGHT_CONTROL As Long = DIK_RCONTROL
Public Const DX_KEY_NUMPAD_SLASH As Long = DIK_DIVIDE
Public Const DX_KEY_SYS_RQ As Long = DIK_SYSRQ
Public Const DX_KEY_RIGHT_ALT As Long = DIK_RMENU
Public Const DX_KEY_PAUSE_BREAK As Long = DIK_PAUSE
Public Const DX_KEY_HOME As Long = DIK_HOME
Public Const DX_KEY_UP As Long = DIK_UP
Public Const DX_KEY_PAGE_UP As Long = DIK_PRIOR
Public Const DX_KEY_LEFT As Long = DIK_LEFT
Public Const DX_KEY_RIGHT As Long = DIK_RIGHT

Public Const DX_KEY_END As Long = DIK_END
Public Const DX_KEY_DOWN As Long = DIK_DOWN
Public Const DX_KEY_PAGE_DOWN As Long = DIK_NEXT
Public Const DX_KEY_INSERT As Long = DIK_INSERT
Public Const DX_KEY_DELETE As Long = DIK_DELETE
Public Const DX_KEY_LEFT_WINDOWS As Long = DIK_LWIN
Public Const DX_KEY_RIGHT_WINDOWS As Long = DIK_RWIN
Public Const DX_KEY_APPS As Long = DIK_APPS

Public Const COLOR_DEPTH_8_BIT As Long = 8
Public Const COLOR_DEPTH_15_BIT As Long = 15
Public Const COLOR_DEPTH_16_BIT As Long = 16
Public Const COLOR_DEPTH_24_BIT As Long = 24
Public Const COLOR_DEPTH_32_BIT As Long = 32

Public Const MAX_COLORS_PALETTE As Long = 256

Public DirectX7 As DirectX7
Public Direct_Draw As DirectDraw7
Public Direct_Draw_Surface_Caps As DDSCAPS2

Public Direct_Draw_Pixel_Format As DDPIXELFORMAT
Public Pixel_Format As Long

Public Direct_Draw_Palette As DirectDrawPalette
Public Palette_8_Bit(255) As PALETTEENTRY

Public Direct_Draw_Clipper As DirectDrawClipper

Public Primary_Surface As DirectDrawSurface7
Public Primary_Surface_Description As DDSURFACEDESC2
Public Primary_Surface_Rect As RECT
Public Primary_Surface_Pitch As Long
Public Primary_Surface_Pointer As Long

Public Backbuffer_Surface As DirectDrawSurface7
Public Backbuffer_Surface_Description As DDSURFACEDESC2
Public Backbuffer_Surface_Rect As RECT
Public Backbuffer_Surface_Pitch As Long
Public Backbuffer_Surface_Pointer As Long

Public Direct_Input As DirectInput

Public Keyboard_Device As DirectInputDevice
Public Keyboard_State As DIKEYBOARDSTATE

Public Mouse_State As DIMOUSESTATE
Public Mouse_Device As DirectInputDevice
Public Mouse_Properties As DIPROPLONG
Public Mouse_Event_Handle As Long
Public Mouse_Cursor_X As Long, Mouse_Cursor_Y As Long
Public Mouse_Cursor_Position As Point_API

Public Sub DirectX7_Setup(hwnd As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color_Bit As Long)
    
    Set DirectX7 = New DirectX7
    
    Set Direct_Draw = DirectX7.DirectDrawCreate("")

    If Fullscreen_Enabled Then
        
        'Initializing fullscreen mode.
        '------------------------------------------------------------
        Direct_Draw.SetCooperativeLevel hwnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
        Direct_Draw.SetDisplayMode Width, Height, Color_Bit, 0, DDSDM_DEFAULT

        'Creating primary surface.
        '------------------------------------------------------------
        Primary_Surface_Description.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        Primary_Surface_Description.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_SYSTEMMEMORY Or DDSCAPS_3DDEVICE
        Primary_Surface_Description.lBackBufferCount = 1
        Set Primary_Surface = Direct_Draw.CreateSurface(Primary_Surface_Description)
        
        'Creating backbuffer surface.
        '------------------------------------------------------------
        Direct_Draw_Surface_Caps.lCaps = DDSCAPS_BACKBUFFER
        Set Backbuffer_Surface = Primary_Surface.GetAttachedSurface(Direct_Draw_Surface_Caps)
        Backbuffer_Surface.GetSurfaceDesc Backbuffer_Surface_Description

        'Get the pixel format.
        '----------------------------------------------------------
        Primary_Surface.GetPixelFormat Direct_Draw_Pixel_Format
        Pixel_Format = Direct_Draw_Pixel_Format.lRGBBitCount
        
    Else
    
        'Initializing windowed mode.
        '------------------------------------------------------------
        Direct_Draw.SetCooperativeLevel 0, DDSCL_NORMAL
                
        'Creating primary surface.
        '------------------------------------------------------------
        Primary_Surface_Description.lFlags = DDSD_CAPS
        Primary_Surface_Description.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        Primary_Surface_Description.lWidth = Width
        Primary_Surface_Description.lHeight = Height
        Set Primary_Surface = Direct_Draw.CreateSurface(Primary_Surface_Description)
        
        Direct_Draw.GetDisplayMode Primary_Surface_Description 'Obtains surface info.
        
        'Creating backbuffer surface.
        '------------------------------------------------------------
        Backbuffer_Surface_Description.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        Backbuffer_Surface_Description.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY Or DDSCAPS_3DDEVICE
        Backbuffer_Surface_Description.lWidth = Width
        Backbuffer_Surface_Description.lHeight = Height
        Backbuffer_Surface_Rect.Right = Backbuffer_Surface_Description.lWidth
        Backbuffer_Surface_Rect.Bottom = Backbuffer_Surface_Description.lHeight
        Set Backbuffer_Surface = Direct_Draw.CreateSurface(Backbuffer_Surface_Description)
        
        Direct_Draw.GetDisplayMode Backbuffer_Surface_Description 'Obtains surface info.
        
        'Creating clipping regions
        '------------------------------------------------------------
        Set Direct_Draw_Clipper = Direct_Draw.CreateClipper(0)
        Direct_Draw_Clipper.SetHWnd hwnd
        Primary_Surface.SetClipper Direct_Draw_Clipper
        
        'Get the pixel format.
        '----------------------------------------------------------
        
        Primary_Surface.GetPixelFormat Direct_Draw_Pixel_Format
        Pixel_Format = Direct_Draw_Pixel_Format.lRGBBitCount

    End If
    
    'Initializing Direct Input for the keyboard.
    '----------------------------------------------------------
    Set Direct_Input = DirectX7.DirectInputCreate
    Set Keyboard_Device = Direct_Input.CreateDevice("GUID_SysKeyboard")
    Keyboard_Device.SetCommonDataFormat DIFORMAT_KEYBOARD
    Keyboard_Device.SetCooperativeLevel hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    Keyboard_Device.Acquire
    Keyboard_Device.GetDeviceStateKeyboard Keyboard_State
    
    'Initializing Direct Input for the mouse.
    '------------------------------------------------------------
    'Set Direct_Input = DirectX7.DirectInputCreate
    'Set Mouse_Device = Direct_Input.CreateDevice("GUID_SYSMOUSE")
    'Mouse_Device.SetCommonDataFormat DIFORMAT_MOUSE
    'Mouse_Device.SetCooperativeLevel hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
    'Mouse_Properties.lHow = DIPH_DEVICE
    'Mouse_Properties.lObj = 0
    'Mouse_Properties.lData = Mouse_Buffer_Size
    'Mouse_Properties.lSize = Len(Mouse_Properties)
    'Mouse_Device.SetProperty "DIPROP_BUFFERSIZE", Mouse_Properties
    'Mouse_Event_Handle = DirectX7.CreateEvent(Main)
    'Mouse_Device.SetEventNotification Mouse_Event_Handle
    'Mouse_Device.Acquire
    
    'GetCursorPos Mouse_Cursor_Position
    'ScreenToClient hWnd, Mouse_Cursor_Position
    'Mouse_Cursor_X = Mouse_Cursor_Position.X
    'Mouse_Cursor_Y = Mouse_Cursor_Position.Y
    
End Sub

Public Function DirectX_Load_Palette(File_Path As String, Palette() As PALETTEENTRY) As Long
    
    Dim Line_Of_Text() As String, New_Line_Of_Text() As String
    
    Dim Current_Line As Long, New_Current_Line As Long, Temperary_Line As Long
    
    Dim Number_Of_Lines As Long
    
    Dim String_Data As Variant

    If File_Path = "" Then GoTo Error_Handler
    
    If Right(File_Path, 4) <> ".pal" And Right(File_Path, 4) <> ".PAL" Then GoTo Error_Handler
    
    Open File_Path For Input As #1
    
        If EOF(1) Then GoTo Error_Handler
    
        While Not EOF(1)
        
            DoEvents
            
            ReDim Preserve Line_Of_Text(Current_Line) As String
            
            ReDim Preserve New_Line_Of_Text(New_Current_Line) As String
            
            Line Input #1, Line_Of_Text(Current_Line)
            
            Line_Of_Text(Current_Line) = Trim(Line_Of_Text(Current_Line))
            
            If Len(Line_Of_Text(Current_Line)) <> 0 And Len(Trim(Line_Of_Text(Current_Line))) <> 0 Then
            '              Null                                       Blank Line
                
                New_Line_Of_Text(New_Current_Line) = Line_Of_Text(Current_Line)
                
                New_Current_Line = New_Current_Line + 1
            
            End If
            
            Current_Line = Current_Line + 1
        
        Wend
        
    Close #1
        
    Number_Of_Lines = New_Current_Line
    
    For Current_Line = 1 To MAX_COLORS_PALETTE - 1
        
        String_Data = String_Split(New_Line_Of_Text(Current_Line))
        
        Palette(Current_Line).Red = Val(String_Data(0))
        Palette(Current_Line).Green = Val(String_Data(1))
        Palette(Current_Line).Blue = Val(String_Data(2))
        Palette(Current_Line).flags = Val(String_Data(3))
        
    Next Current_Line
    
    DirectX_Load_Palette = 1

    Exit Function

Error_Handler:

    MsgBox "Error Loading Palette", vbCritical

End Function

Public Sub DirectX_Clear(Color As Long)

    Backbuffer_Surface.BltColorFill Backbuffer_Surface_Rect, Color

End Sub

Public Sub DirectX_Blit(hwnd As Long)

    'Hybrid blitter of Fullscreen and Windowed mode.

    If Fullscreen_Enabled Then

        Primary_Surface.Flip Nothing, DDFLIP_WAIT

    Else

        DirectX7.GetWindowRect hwnd, Primary_Surface_Rect
        Primary_Surface.Blt Primary_Surface_Rect, Backbuffer_Surface, Backbuffer_Surface_Rect, DDBLT_WAIT
        
    End If

End Sub

Public Sub DirectX_Wait_For_VSync()

    'Keep it within sync of your monitor's
    'refresh rate. Normally it maintains
    '60 frames per second within your loops.
    'Sometimes 86 frames per second depending
    'on what computer is used.

    Direct_Draw.WaitForVerticalBlank DDWAITVB_BLOCKBEGIN, 0

End Sub

Public Function DirectX_Key_State(DirectX_Key_Code As Long) As Long

    Keyboard_Device.GetDeviceStateKeyboard Keyboard_State
    
    DirectX_Key_State = Keyboard_State.Key(DirectX_Key_Code)

End Function

Public Function DirectX_Lock_Surface() As Long
    
    If Backbuffer_Surface_Pointer Then Exit Function

    Backbuffer_Surface.Lock Backbuffer_Surface_Rect, Backbuffer_Surface_Description, DDLOCK_WAIT Or DDLOCK_SURFACEMEMORYPTR, vbNull
    
    'For those who are baffled on what the hell lpSurface is, lpSurface
    'is one of many hidden members in DirectX7 located within the
    'DDSURFACEDESC2, which is one of Microsoft's nifty secrets they didn't
    'want VB programmers to know about. Come to think of it, I don't think
    'Andre Lamothe (programmer/author of many programming books) knows it
    'exists in Visual Basic cause otherwise it would have been within the
    'Microsoft Visual Basic Game Programming With DirectX book I have.
    'All of his other books like Tricks of the 3D Game Programming Gurus
    '(C++) and Tricks of the Windows Game Programming Gurus (C++) have it
    'and explain lpSurface and Memory Pitches in huge detail. Anyways,
    'lpSurface is a pointer to a surface you are working with, in this
    'case, the backbuffer. Once you have the pointer to the surface and
    'the memory pitch, you can use direct memory addressing to plot
    'pixels at warp speed.
    
    '-----------------------------------------------------------------
    
    Backbuffer_Surface_Pointer = Backbuffer_Surface_Description.lpSurface
    
    '-----------------------------------------------------------------
    
    'Equivilant to C++:
        
    '   unsigned int *video_buffer;
    '   video_buffer = *ddsd.lpSurface;
    
    'Where one array will point to another in memory. In this
    'case, we are pointing to Backbuffer_Surface_Description.lpSurface
    'with an array for memory manipulation (a.k.a. super sonic pixels!!!)
    'This will support multiple color modes.
    
    Select Case Pixel_Format
    
        Case COLOR_DEPTH_8_BIT
        
            'There's 1 byte per pixel (8 bits), so we leave the
            'memory pitch the same.
            
            Backbuffer_Surface_Pitch = Backbuffer_Surface_Description.lPitch
        
            With Safe_Array
            
                .Data_Size = 1
                .Dimensions = 1
                .Data_Pointer = Backbuffer_Surface_Pointer
                .Safe_Array(0).lLBound = 0
                .Safe_Array(0).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
                .Safe_Array(1).lLBound = 0
                .Safe_Array(1).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
        
                CopyMemory ByVal VarPtrArray(Video_Buffer_8()), VarPtr(Safe_Array), 4&
            
            End With
        
        Case COLOR_DEPTH_15_BIT
            
            'Red = 5 bits
            'Green = 5 bits
            'Blue = 5 bits
            
            '5 + 5 + 5 = 15 bits
            
            'There's 2 bytes per pixel (16 bits), so we divide the memory pitch
            'by 2.
            
            Backbuffer_Surface_Pitch = Backbuffer_Surface_Description.lPitch / 2
            
            With Safe_Array
            
                .Data_Size = 2
                .Dimensions = 1
                .Data_Pointer = Backbuffer_Surface_Pointer
                .Safe_Array(0).lLBound = 0
                .Safe_Array(0).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
                .Safe_Array(1).lLBound = 0
                .Safe_Array(1).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
        
                CopyMemory ByVal VarPtrArray(Video_Buffer_16()), VarPtr(Safe_Array), 4&
            
            End With
        
        Case COLOR_DEPTH_16_BIT
            
            'Red = 5 bits
            'Green = 6 bits
            'Blue = 5 bits
            
            '5 + 6 + 5 = 16 bits
            
            'I gotta be honest of what happened to me earlier. I
            'declared the Video_Buffer() as long (4 bytes) and was using it
            'on 16 bit color mode. For some reason my pixels skipped a pixel
            'when I flooded the window. Not only that, my backbuffer surface
            'was supposed to be 640x480 but it was like the backbuffer surface was
            '320x240. The reason why it was doing that was because in 16 bit
            'color mode, each pixel represents 2 bytes. My Video_Buffer() array
            'on the other hand was 4 bytes. So I got this as a result:
            
                '@ = colored pixel
                '# = Byte number
            
            '        1  2  3  4  1  2  3  4
            '       [@][@][ ][ ][@][@][ ][ ]    Yikes!!! Skipped pixels.
            '
            
            'Moral of the story is to make sure your Video_Buffer()
            'arrays that's gonna be used for direct memory addressing
            '(plotting pixels) are consistant to what color mode you are
            'on.
            
            '1 byte per pixel
            '(Video_Buffer() As Byte) for 8 bit color mode.
            
            '2 bytes per pixel
            '(Video_Buffer() As Integer) for 15 or 16 bit color mode.
            
            '3 bytes per pixel
            '(Video_Buffer() As Long) for 24 bit color mode.
            
            '4 bytes per pixel
            '(Video_Buffer() As Long) for 32 bit color mode.
            
            'There's 2 bytes per pixel (16 bits), so we divide the memory pitch
            'by 2.
            
            Backbuffer_Surface_Pitch = Backbuffer_Surface_Description.lPitch / 2
            
            With Safe_Array
            
                .Data_Size = 2
                .Dimensions = 1
                .Data_Pointer = Backbuffer_Surface_Pointer
                .Safe_Array(0).lLBound = 0
                .Safe_Array(0).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
                .Safe_Array(1).lLBound = 0
                .Safe_Array(1).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
        
                CopyMemory ByVal VarPtrArray(Video_Buffer_16()), VarPtr(Safe_Array), 4&
            
            End With
        
        Case COLOR_DEPTH_24_BIT
            
            'Red = 8 bits
            'Green = 8 bits
            'Blue = 8 bits
            
            '8 + 8 + 8 = 24 bits
            
            'Note: Never been tested. All of the computers that's in
            '      my house have no support for 24 bit color mode.
            
            'There's 3 bytes per pixel (24 bits), so we divide the memory pitch
            'by 3.
            
            Backbuffer_Surface_Pitch = Backbuffer_Surface_Description.lPitch / 3
            
            With Safe_Array
            
                .Data_Size = 4
                .Dimensions = 1
                .Data_Pointer = Backbuffer_Surface_Pointer
                .Safe_Array(0).lLBound = 0
                .Safe_Array(0).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
                .Safe_Array(1).lLBound = 0
                .Safe_Array(1).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
        
                CopyMemory ByVal VarPtrArray(Video_Buffer_24()), VarPtr(Safe_Array), 4&
            
            End With
        
        Case COLOR_DEPTH_32_BIT
            
            'Red = 8 bits
            'Green = 8 bits
            'Blue = 8 bits
            'Alpha = 8 bits
            
            '8 + 8 + 8 + 8 = 32 bits
            
            'There's 4 bytes per pixel (32 bits), so we divide the memory pitch
            'by 4.
            
            Backbuffer_Surface_Pitch = Backbuffer_Surface_Description.lPitch / 4
        
            With Safe_Array
            
                .Data_Size = 4
                .Dimensions = 1
                .Data_Pointer = Backbuffer_Surface_Pointer
                .Safe_Array(0).lLBound = 0
                .Safe_Array(0).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
                .Safe_Array(1).lLBound = 0
                .Safe_Array(1).cElements = Window_Width + Window_Height * Backbuffer_Surface_Pitch
        
                CopyMemory ByVal VarPtrArray(Video_Buffer_32()), VarPtr(Safe_Array), 4&
            
            End With
        
    End Select
    
    '-----------------------------------------------------------------
    
    DirectX_Lock_Surface = Backbuffer_Surface_Pointer
    
End Function

Public Function DirectX_Unlock_Surface() As Long
    
    If Backbuffer_Surface_Pointer = 0 Then Exit Function

    Backbuffer_Surface.Unlock Backbuffer_Surface_Rect
    
    Backbuffer_Surface_Pointer = 0
    Backbuffer_Surface_Pitch = 0
    
    DirectX_Unlock_Surface = 1
    
End Function

Public Sub DirectX_Draw_Pixel(ByVal X As Long, ByVal Y As Long, ByVal Color As Long)

    Backbuffer_Surface.SetLockedPixel X, Y, Color

End Sub

Public Sub Draw_Pixel(ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
    
    'Supports multiple color formats.
    
    'Note: it's faster to use just one of the arrays within
    '      this function rather than calling the function itself.
    '      Plus If statements and Select Case statements slow
    '      your functions/subs down even more. Surface must be locked
    '      for this to work and unlocked when done.
    
    Select Case Pixel_Format
    
        Case COLOR_DEPTH_8_BIT
        
            Video_Buffer_8(X + Y * Backbuffer_Surface_Pitch) = Color
        
        Case COLOR_DEPTH_15_BIT
        
            Video_Buffer_16(X + Y * Backbuffer_Surface_Pitch) = Color
        
        Case COLOR_DEPTH_16_BIT
        
            Video_Buffer_16(X + Y * Backbuffer_Surface_Pitch) = Color
        
        Case COLOR_DEPTH_24_BIT
        
            Video_Buffer_24(X + Y * Backbuffer_Surface_Pitch) = Color
        
        Case COLOR_DEPTH_32_BIT
        
            Video_Buffer_32(X + Y * Backbuffer_Surface_Pitch) = Color
        
    End Select

End Sub

Public Function Clip_Line(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Long
    
    On Error Resume Next
    
    Const CLIP_CODE_C As Long = 0 '&H0&
    Const CLIP_CODE_N As Long = 8 '&H8&
    Const CLIP_CODE_S As Long = 4 '&H4&
    Const CLIP_CODE_E As Long = 2 '&H2&
    Const CLIP_CODE_W As Long = 1 '&H1&

    Const CLIP_CODE_NE As Long = 10 '&HA&
    Const CLIP_CODE_SE As Long = 6 '&H6&
    Const CLIP_CODE_NW As Long = 9 '&H9&
    Const CLIP_CODE_SW As Long = 5 '&H5&
    
    Dim Clip_X1 As Long
    Dim Clip_Y1 As Long
    Dim Clip_X2 As Long
    Dim Clip_Y2 As Long
    
    Dim Point_A_Clip_Code As Long
    Dim Point_B_Clip_Code As Long
    
    Clip_X1 = X1
    Clip_Y1 = Y1
    Clip_X2 = X2
    Clip_Y2 = Y2
    
    If Y1 < Min_Clip_Y Then
        
        Point_A_Clip_Code = Point_A_Clip_Code Or CLIP_CODE_N
    
    Else
    
        If Y1 > Max_Clip_Y Then
        
            Point_A_Clip_Code = Point_A_Clip_Code Or CLIP_CODE_S
        
        End If
        
    End If

    If X1 < Min_Clip_X Then
    
        Point_A_Clip_Code = Point_A_Clip_Code Or CLIP_CODE_W
    
    Else
    
        If X1 > Max_Clip_X Then
    
            Point_A_Clip_Code = Point_A_Clip_Code Or CLIP_CODE_E
        
        End If
        
    End If
    
    If Y2 < Min_Clip_Y Then
        
        Point_B_Clip_Code = Point_B_Clip_Code Or CLIP_CODE_N
    
    Else
        
        If Y2 > Max_Clip_Y Then
        
            Point_B_Clip_Code = Point_B_Clip_Code Or CLIP_CODE_S
        
        End If
        
    End If

    If X2 < Min_Clip_X Then
    
        Point_B_Clip_Code = Point_B_Clip_Code Or CLIP_CODE_W
    
    Else
        
        If X2 > Max_Clip_X Then
    
            Point_B_Clip_Code = Point_B_Clip_Code Or CLIP_CODE_E

        End If
        
    End If

    If Point_A_Clip_Code And Point_B_Clip_Code Then
        
        Clip_Line = 0
        
        Exit Function
        
    End If

    If Point_A_Clip_Code = 0 And Point_B_Clip_Code = 0 Then
        
        Clip_Line = 1
        
        X1 = X1
        Y1 = Y1
        X2 = X2
        Y2 = Y2

        Exit Function
    
    End If
    
    Select Case Point_A_Clip_Code
    
        Case CLIP_CODE_C
        
            Clip_X1 = X1
            Clip_Y1 = Y1
        
        Case CLIP_CODE_N
        
            Clip_Y1 = Min_Clip_Y
            Clip_X1 = X1 + 0.5 + (Min_Clip_Y - Y1) * (X2 - X1) / (Y2 - Y1)
            
        Case CLIP_CODE_S
        
            Clip_Y1 = Max_Clip_Y
            Clip_X1 = X1 + 0.5 + (Max_Clip_Y - Y1) * (X2 - X1) / (Y2 - Y1)
        
        Case CLIP_CODE_W
        
            Clip_X1 = Min_Clip_X
            Clip_Y1 = Y1 + 0.5 + (Min_Clip_X - X1) * (Y2 - Y1) / (X2 - X1)
        
        Case CLIP_CODE_E
        
            Clip_X1 = Max_Clip_X
            Clip_Y1 = Y1 + 0.5 + (Max_Clip_X - X1) * (Y2 - Y1) / (X2 - X1)
        
        Case CLIP_CODE_NE
        
            Clip_Y1 = Min_Clip_Y
            
            Clip_X1 = X1 + 0.5 + (Min_Clip_Y - Y1) * (X2 - X1) / (Y2 - Y1)
        
            If Clip_X1 < Min_Clip_X Or Clip_X1 > Max_Clip_X Then
            
                Clip_X1 = Max_Clip_X

                Clip_Y1 = Y1 + 0.5 + (Max_Clip_X - X1) * (Y2 - Y1) / (X2 - X1)
            
            End If
        
        Case CLIP_CODE_SE
        
            Clip_Y1 = Max_Clip_Y
            
            Clip_X1 = X1 + 0.5 + (Max_Clip_Y - Y1) * (X2 - X1) / (Y2 - Y1)
        
            If Clip_X1 < Min_Clip_X Or Clip_X1 > Max_Clip_X Then
                
                Clip_X1 = Max_Clip_X
                
                Clip_Y1 = Y1 + 0.5 + (Max_Clip_X - X1) * (Y2 - Y1) / (X2 - X1)
            
            End If
            
        Case CLIP_CODE_NW

            Clip_Y1 = Min_Clip_Y
            
            Clip_X1 = X1 + 0.5 + (Min_Clip_Y - Y1) * (X2 - X1) / (Y2 - Y1)
        
            If Clip_X1 < Min_Clip_X Or Clip_X1 > Max_Clip_X Then
            
                Clip_X1 = Min_Clip_X
                
                Clip_Y1 = Y1 + 0.5 + (Min_Clip_X - X1) * (Y2 - Y1) / (X2 - X1)
            
            End If
        
        Case CLIP_CODE_SW

            Clip_Y1 = Max_Clip_Y
            
            Clip_X1 = X1 + 0.5 + (Max_Clip_Y - Y1) * (X2 - X1) / (Y2 - Y1)
        
            If Clip_X1 < Min_Clip_X Or Clip_X1 > Max_Clip_X Then
            
                Clip_X1 = Min_Clip_X
                
                Clip_Y1 = Y1 + 0.5 + (Min_Clip_X - X1) * (Y2 - Y1) / (X2 - X1)
            
            End If
        
    End Select
    
    Select Case Point_B_Clip_Code
    
        Case CLIP_CODE_C

            Clip_X2 = X2
            Clip_Y2 = Y2

        Case CLIP_CODE_N
        
            Clip_Y2 = Min_Clip_Y
            Clip_X2 = X2 + (Min_Clip_Y - Y2) * (X1 - X2) / (Y1 - Y2)
            
        Case CLIP_CODE_S
        
            Clip_Y2 = Max_Clip_Y
            Clip_X2 = X2 + (Max_Clip_Y - Y2) * (X1 - X2) / (Y1 - Y2)
        
        Case CLIP_CODE_W
        
            Clip_X2 = Min_Clip_X
            Clip_Y2 = Y2 + (Min_Clip_X - X2) * (Y1 - Y2) / (X1 - X2)
        
        Case CLIP_CODE_E
        
            Clip_X2 = Max_Clip_X
            Clip_Y2 = Y2 + (Max_Clip_X - X2) * (Y1 - Y2) / (X1 - X2)
        
        Case CLIP_CODE_NE
        
            Clip_Y2 = Min_Clip_Y
            
            Clip_X2 = X2 + 0.5 + (Min_Clip_Y - Y2) * (X1 - X2) / (Y1 - Y2)
        
            If Clip_X2 < Min_Clip_X Or Clip_X2 > Max_Clip_X Then
                
                Clip_X2 = Max_Clip_X
                
                Clip_Y2 = Y2 + 0.5 + (Max_Clip_X - X2) * (Y1 - Y2) / (X1 - X2)
            
            End If
        
        Case CLIP_CODE_SE

            Clip_Y2 = Max_Clip_Y
            
            Clip_X2 = X2 + 0.5 + (Max_Clip_Y - Y2) * (X1 - X2) / (Y1 - Y2)
        
            If Clip_X2 < Min_Clip_X Or Clip_X2 > Max_Clip_X Then
                
                Clip_X2 = Max_Clip_X
                
                Clip_Y2 = Y2 + 0.5 + (Max_Clip_X - X2) * (Y1 - Y2) / (X1 - X2)
            
            End If
            
        Case CLIP_CODE_NW

            Clip_Y2 = Min_Clip_Y
            
            Clip_X2 = X2 + 0.5 + (Min_Clip_Y - Y2) * (X1 - X2) / (Y1 - Y2)
        
            If Clip_X2 < Min_Clip_X Or Clip_X2 > Max_Clip_X Then
                
                Clip_X2 = Min_Clip_X
                
                Clip_Y2 = Y2 + 0.5 + (Min_Clip_X - X2) * (Y1 - Y2) / (X1 - X2)
            
            End If
        
        Case CLIP_CODE_SW

            Clip_Y2 = Max_Clip_Y
            
            Clip_X2 = X2 + 0.5 + (Max_Clip_Y - Y2) * (X1 - X2) / (Y1 - Y2)
        
            If Clip_X2 < Min_Clip_X Or Clip_X2 > Max_Clip_X Then

                Clip_X2 = Min_Clip_X

                Clip_Y2 = Y2 + 0.5 + (Min_Clip_X - X2) * (Y1 - Y2) / (X1 - X2)
            
            End If
        
    End Select
        
    If ((Clip_X1 < Min_Clip_X) Or (Clip_X1 > Max_Clip_X) Or _
       (Clip_Y1 < Min_Clip_Y) Or (Clip_Y1 > Max_Clip_Y) Or _
       (Clip_X2 < Min_Clip_X) Or (Clip_X2 > Max_Clip_X) Or _
       (Clip_Y2 < Min_Clip_Y) Or (Clip_Y2 > Max_Clip_Y)) Then
        
        Clip_Line = 0
        Exit Function
        
    End If
    
    X1 = Clip_X1
    Y1 = Clip_Y1
    X2 = Clip_X2
    Y2 = Clip_Y2
    
    Clip_Line = 1

End Function

Public Function Draw_Line_Bresenham(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Color As Long) As Long

    Dim Index As Long
    Dim Delta_X As Long
    Dim Delta_Y As Long
    Dim X_Increment As Long
    Dim Y_Increment As Long
  
    Dim X As Long
    Dim Y As Long
    Dim Discriminant As Long


    X = X1
    Y = Y1

    Delta_X = (X2 - X1)
    Delta_Y = (Y2 - Y1)
  
    If (Delta_X >= 0) Then
     
        X_Increment = 1
  
    ElseIf (Delta_X < 0) Then
  
        X_Increment = -1
        Delta_X = -Delta_X
    
    End If
  
    If (Delta_Y >= 0) Then
  
        Y_Increment = 1
  
    ElseIf (Delta_Y < 0) Then
  
        Y_Increment = -1
        Delta_Y = -Delta_Y
    
    End If
    
    If Delta_X > Delta_Y Then
        
        Discriminant = Delta_X / 2
        
        For Index = 0 To Delta_X - 1
        
            Video_Buffer_32(X + Y * Backbuffer_Surface_Pitch) = Color
                
            Discriminant = Discriminant + Delta_Y
            
            If Discriminant > Delta_X - 1 Then
            
                Discriminant = Discriminant - Delta_X
        
                Y = Y + Y_Increment
        
            End If
            
            X = X + X_Increment
            
        Next Index
        
    Else
        
        Discriminant = Delta_Y / 2
        
        For Index = 0 To Delta_Y - 1
            
            Video_Buffer_32(X + Y * Backbuffer_Surface_Pitch) = Color
            
            Discriminant = Discriminant + Delta_X
            
            If Discriminant > Delta_Y - 1 Then
            
                Discriminant = Discriminant - Delta_Y
        
                X = X + X_Increment
        
            End If
            
            Y = Y + Y_Increment
            
        Next Index
    
    End If
    
    Draw_Line_Bresenham = 1

End Function

Public Function Draw_Clip_Line_Bresenham(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Color As Long) As Long

    Dim Clip_X1 As Long
    Dim Clip_Y1 As Long
    Dim Clip_X2 As Long
    Dim Clip_Y2 As Long
    
    Clip_X1 = X1
    Clip_Y1 = Y1
    Clip_X2 = X2
    Clip_Y2 = Y2
    
    If Clip_Line(Clip_X1, Clip_Y1, Clip_X2, Clip_Y2) = 1 Then
        
        Draw_Line_Bresenham Clip_X1, Clip_Y1, Clip_X2, Clip_Y2, Color
        
    End If
    
    Draw_Clip_Line_Bresenham = 1

End Function

Public Sub DirectX_Draw_Line(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    
    Backbuffer_Surface.SetForeColor Color
    Backbuffer_Surface.DrawLine X1, Y1, X2, Y2

End Sub

Public Sub DirectX_Draw_Rectangle(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

    Backbuffer_Surface.SetForeColor Color
    Backbuffer_Surface.DrawBox X1, Y1, X2, Y2

End Sub

Public Sub DirectX_Draw_Circle(ByVal X As Long, ByVal Y As Long, ByVal Radius As Long, ByVal Color As Long)

    Backbuffer_Surface.SetForeColor Color
    Backbuffer_Surface.DrawCircle X, Y, Radius

End Sub

Public Sub DirectX_Draw_Ellipse(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    
    Backbuffer_Surface.SetForeColor Color
    Backbuffer_Surface.DrawEllipse X1, Y1, X2, Y2

End Sub

Public Sub DirectX_Draw_Text(Text As String, ByVal X As Long, ByVal Y As Long, Font_Transparency_Disabled As Boolean, ByVal Color As Long)

    Backbuffer_Surface.SetForeColor Color
    Backbuffer_Surface.DrawText X, Y, Text, Font_Transparency_Disabled

End Sub

Public Sub DirectX_Shutdown(hwnd As Long)
    
    'If it were fullscreen, it would restore your screen resolution
    'back to normal.
    '-----------------------------------------------------------------

    Direct_Draw.RestoreDisplayMode
    Direct_Draw.SetCooperativeLevel 0, DDSCL_NORMAL
    
    '-----------------------------------------------------------------
    
    'Once you close out of the application, this will allocate memory by
    'freeing any DirectX initializations that have been made.
    
    'DirectX7.DestroyEvent Mouse_Event_Handle
    
    Set Backbuffer_Surface = Nothing
    Set Primary_Surface = Nothing
    'Set Mouse_Device = Nothing
    Set Keyboard_Device = Nothing
    Set Direct_Draw = Nothing
    Set DirectX7 = Nothing
    

End Sub
