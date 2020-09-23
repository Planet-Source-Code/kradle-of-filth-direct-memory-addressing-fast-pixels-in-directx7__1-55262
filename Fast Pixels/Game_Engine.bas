Attribute VB_Name = "Game_Engine"
Option Explicit

Public Const Milliseconds_Per_Second As Long = 1000

Public Get_Time As Long
Public Get_Temperary_Time As Long
Public Milliseconds As Long
Public Get_Frames_Per_Second As Long
Public Frame_Count As Long

Public Game_Active As Long
Public Fullscreen_Enabled As Long

Public Safe_Array As Safe_Array_Header

Public Video_Buffer_8() As Byte
Public Video_Buffer_16() As Integer
Public Video_Buffer_24() As Long
Public Video_Buffer_32() As Long

Public Fullscreen_Width As Long
Public Fullscreen_Height As Long
Public Color_Bit As Long
 
Public Window_Width As Long
Public Window_Height As Long

Public Min_Clip_X As Long
Public Min_Clip_Y As Long
Public Max_Clip_X As Long
Public Max_Clip_Y As Long

Dim X As Long, Y As Long 'For pixels.

Dim X1 As Long, X2 As Long 'For Scanlines.

Public Sub DoEvents_Fast()
    
    'This does events only when absolutely necessary and still prevents
    'your program from locking up. The result is a Do loop that is
    'multiple times faster than an ordinary Do/DoEvents/Loop, which is needed for
    'realtime loops. I've experimented with multiple methods I've found on Planet
    'Source Code, and here are my results:
    
    'Note - This all has been done on my AMD Athlon 1.2 Ghz Processor. Results may vary.
        
'--------------------------------------------------------------------------------------
        
    'Highest durations per second
    '--------------------
    'VB - 192136
    'Exe - 296140
    
        'Slow, slugish, and ugly for realtime.
            
        'DoEvents
    
'-------------------------------------------------------------------------------------

    'Highest durations per second
    '--------------------
    'VB - 688950
    'Exe - 735468
    
        'If PeekMessage(Message, 0, 0, 0, PM_NOREMOVE) Then
                
        '    DoEvents
            
        'End If
        
'--------------------------------------------------------------------------------------
    
    'Highest durations per second
    '--------------------
    'VB - 965230
    'Exe - 1113434
    
        'Problem with this is that it's only active when an event
        'has occured. With this I just simply held a key down.
            
        'If GetInputState() Then
                
        '    DoEvents
            
        'End If

'-------------------------------------------------------------------------------------
        
    'Highest durations per second
    '--------------------
    'VB - 947204
    'Exe - 1101420
    
        'This is the fastest and most reliable method so far.
    
        If GetQueueStatus(QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT) Then
            
            DoEvents
            
        End If
    
End Sub

Public Sub Window_Setup(Window As Form, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Caption As String, Auto_Redraw As Boolean)
    
    Window.Caption = Caption
    Window.Left = X
    Window.Top = Y
    Window.Width = Width * Screen.TwipsPerPixelX
    Window.Height = Height * Screen.TwipsPerPixelY
    Window.ScaleMode = 3
    Window.AutoRedraw = Auto_Redraw
    Window.Show
    Window.Refresh
    
End Sub

Public Sub Game_Loop()

    On Error GoTo Error_Handler
    
    DoEvents_Fast
    
    DirectX_Clear 0 'clear the backbuffer with a color.
    
    'Put your code here.
    '---------------------------------------------------------------
    
    DirectX_Lock_Surface
        
    'These for loops alone take up too much speed (30 FPS in 640x480),
    'but this is just an example of my high speed pixel writer flooding
    'the whole screen.
        
    For Y = 0 To Window_Height - 1
    
        For X = 0 To Window_Width - 1
        
            'If running 16 bit color mode, be sure to use Video_Buffer_16
            'If running 32 bit color mode, be sure to use Video_Buffer_32
            'etc etc etc...
            
            'Now the reason why I decided not to use a function/sub
            'routine for plotting pixels is because calls to functions and
            'subs take some speed away. For example,
            'the expression Fix(Value / (2 ^ Bits_To_Shift)) by itself
            'is 3x faster than using it within the function
            'Right_Bit_Shift(). However, using functions and subs
            'help organize code to make it more readable. Avoid them
            'in areas where speed is highly needed and (worse case)
            'avoid (functions/subs) within (functions/subs) within
            '(functions/subs)!!!
    
            'C++ has a thing VB is kinda lacking...Inline. When you inline
            'a function/sub, the compiler makes its best attempt to run
            'the code within your function/sub rather than making the actual
            'call to the function. Huh?!!! What this means is
            'that when you, for example, want to call Right_Bit_Shift(),
            'what the compiler will do is replace that area you put your
            'function at with Fix(Value / (2 ^ Bits_To_Shift)), which is
            'the code you had in the function. To inline in Visual Basic,
            'you have to do that by hand optimizing your code by manually
            'inserting Fix(Value / (2 ^ Bits_To_Shift)) where Right_Bit_Shift()
            'is. The compiler will not do that for you. And its 2 to 3 times
            'faster than the actual (function/sub) with the same code!!!
            
            'This floods the screen with pixels using direct memory
            'addressing. Surface must be locked to do it and unlocked
            'when done.
            
            '60 FPS (320x240x32)
            
                'Video_Buffer_32(X + Y * Backbuffer_Surface_Pitch) = 3000
            
            '30 FPS  (320x240x32)
            '        (RGB is slow cause its a function thats being called.
            '        Rnd is also slow but cutting it out made no difference
            '        cause I still have 30 FPS.)
            
                Video_Buffer_32(X + Y * Backbuffer_Surface_Pitch) = RGB(Rnd * 255, 0, 255)
        
        Next X
        
    Next Y
    
    DirectX_Unlock_Surface
    
    DirectX_Draw_Text "FPS" & " - " & Str(Get_Frames_Per_Second), 50, 50, False, RGB(255, 255, 255)
    
    '---------------------------------------------------------------
    
    Frame_Count = Frame_Count + 1
    
    'If it has been 1 second...
    
    If GetTickCount - Milliseconds >= Milliseconds_Per_Second Then
        
        Get_Frames_Per_Second = Frame_Count
        
        Frame_Count = 0

        Milliseconds = GetTickCount

    End If
    
    DirectX_Blit Main.hwnd
        
    DirectX_Wait_For_VSync 'Keep it within sync of your monitor's
                           'refresh rate. Normally it maintains
                           '60 frames per second within your loops.
                           'Sometimes 86 frames per second depending
                           'on what computer is used.
                           
    If DirectX_Key_State(DX_KEY_ESC) Then
        
        Game_Active = 0

    End If
    
    Exit Sub
    
Error_Handler:
    
    DirectX_Unlock_Surface
    
End Sub

Public Sub Main_Loop()
    
    Game_Active = 1
    
    Fullscreen_Enabled = 1
    
    Window_Width = 320
    
    Window_Height = 240
    
    Color_Bit = 32
    
    Window_Setup Main, 0, 0, Window_Width, Window_Height, "DirectX Fast Pixels", False
    
    DirectX7_Setup Main.hwnd, Window_Width, Window_Height, Color_Bit
    
    Build_Color_Look_Up_Tables
    
    Milliseconds = GetTickCount
    
    'This will help the ordinary DoEvents function work faster
    'than usual.
    
    '------------------------------------------------------
    
    Get_Thread = GetCurrentThread
    Get_Process = GetCurrentProcess
    
    SetThreadPriority Get_Thread, THREAD_PRIORITY_HIGHEST
    SetPriorityClass Get_Process, HIGH_PRIORITY_CLASS
    
    '------------------------------------------------------
    
    While Game_Active
    
        Game_Loop
    
    Wend
    
    DirectX_Shutdown Main.hwnd
    
    End

End Sub



