Attribute VB_Name = "Module2"
'This is a module that allows you to control a Hitachi HD44780 (or compatible) chipset LCD display (20x4)
'which is connected to the LTP (Printer ) port. A Great example of hardware works can be found here
'http://www.overclockers.com.au/techstuff/a_diy_lcd/
'
'The software presented on most sites is written in C / C++ or other non VB language.
'Purpose of this module is to add the ability to control LCD Displays to VB applications
'One VERY important detail is that neihter Windows OSes based on NT (NT, 2000, XP) nor Visual
'Basic itself provides functionality that allows you to directly access hardware ports.
'That's why we'll be using a popular I/O Driver Port95NT also called DLPortIO which gives you access to
'lpt port on win 9x, NT, ME, XP...   Great huh ? :-)
'It can be downloaded and used freely.  Use google or altavista to find the .ZIP
'or try here: http://www.driverlinx.com/
'
'
'To make this software work you will need a Hitachi HD44780 (or compatible) chipset LCD display with 14 0r 16
'connectors, some soldering skills, a LPT (printer) cable and some free time.
'Use drawings attached to attach LCD Display to the cable, (A great guide can be found here
'http://www.overclockers.com.au/techstuff/a_diy_lcd/
'(connectors are the same on 16x2 and 20x4 so it doesn't matter that the article is about 16x2 display))
', install Port95NT and use this module
'to control the LCD Display.
'
'  Good Luck!
'Copyright 2004 Dmitriy Prokopov Malmoe Sweden.
'Source code is written by me, the only thing copied from others are LCD Hex Control Codes.
'Keep this message in this module and you may use it in your applications with no charge.

'*********************************************************************************************************************




Public Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

'u really need only this one
Public Declare Function DlPortReadPortUchar Lib "dlportio.dll" (ByVal Port As Long) As Byte

'these are for reference purposes. are not used in project, just show what else DlPortIO can do.
Public Declare Function DlPortReadPortUshort Lib "dlportio.dll" (ByVal Port As Long) As Integer
Public Declare Function DlPortReadPortUlong Lib "dlportio.dll" (ByVal Port As Long) As Long

Public Declare Sub DlPortReadPortBufferUchar Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortReadPortBufferUshort Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortReadPortBufferUlong Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)

Public Declare Sub DlPortWritePortUchar Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Byte)
Public Declare Sub DlPortWritePortUshort Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Integer)
Public Declare Sub DlPortWritePortUlong Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Long)

Public Declare Sub DlPortWritePortBufferUchar Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortWritePortBufferUshort Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortWritePortBufferUlong Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)




Public Const LPT = &H378              'Printer Port Address (LPT1 in Hex)
Public Const LPT_Data = LPT + 0       'Data port
Public Const LPT_Control = LPT + 2    'Control port


Public Function LCD_Init()


LCD_Control &H38, 1.64   'Set 8 bit Mode
LCD_Control &H38, 0.1
LCD_Control &H38, 1.64
LCD_Control &H6, 0.04
LCD_Control &HC, 0.04    'Display on, cursor off, blinking cursor off
LCD_Control &H1, 1.64    'Clear screen


'LCD_Control &H2, 1.64   'Home
'LCD_Control &HF, 0.04   'Display on, cursor on, blinking cursor on

sleep 1
End Function



Public Sub LCD_Clear()
    op = LCD_Control(&H1, 1.64)    'Clear screen
End Sub


Public Sub LCD_DisplayOn()
    LCD_Control &HC, 2

End Sub


Public Sub LCD_DisplayOff()
    LCD_Control &H8, 2
End Sub



Public Sub LCD_WriteString(data As String, Optional row As Integer, Optional col As Integer)

    If IsMissing(row) Or IsMissing(col) Then    'no position data
    Else
        LCD_SetCursor row, col
    End If
    
    For i = 1 To Len(data)
            LCD_Data Asc(Mid(data, i, 1))
    Next i
End Sub


Public Sub LCD_Write20x4_Screen(screen() As String)
'screen must be build  by "20 characters x 4 array elements" method. string() array  must be zero based
'string(0) will be row 1
'string(3) will be row 4
'Ex.
'data(0)        =       "  This is my first  "
'data(1)        =       "       screen       "
'data(2)        =       "Date: Wed 02/03/2004"
'data(3)        =       "left  center   right"
    LCD_Clear
    For i = 0 To 3
        LCD_WriteString screen(i), i + 1, 1
    Next i
End Sub

Public Sub LCD_SetCursor(row As Integer, col As Integer)
    Select Case row
        Case 1
            LCD_Control &H80 + col - 1, 1
        Case 2
            LCD_Control &HC0 + col - 1, 1
        Case 3
            LCD_Control &H94 + col - 1, 1
        Case 4
            LCD_Control &HD4 + col - 1, 1
    End Select
End Sub

Public Sub LCD_CursorLeft()
    LCD_Control &H10
End Sub

Public Sub LCD_CursorRight()
    LCD_Control &H14
End Sub


Public Sub LCD_CursorOn()
    LCD_Control &HD
End Sub
 
Public Sub LCD_CursorOff()
    LCD_Control &HC
End Sub











'/*LOW LEVEL**********************************************************************************************
Public Function LCD_Control(data, Optional sleep)
    If IsMissing(sleep) Then sleep = 1
    
    DlPortWritePortUchar LPT_Data, data    'Send the data to the data port
    DlPortWritePortUchar LPT_Control, &H2  'Set enable (RS=0, R/W=0, E=1)
    DlPortWritePortUchar LPT_Control, &H3  'Clear enable (RS=0, R/W=0, E=0)
    
    sleep sleep
End Function

Public Sub LCD_Data(data)

    DlPortWritePortUchar LPT_Control, &H7  '// RS=1, R/W=0, E=0

    DlPortWritePortUchar LPT_Data, data

    DlPortWritePortUchar LPT_Control, &H6 '// RS=1, R/W=0, E=1
    DlPortWritePortUchar LPT_Control, &H7 '// RS=1, R/W=0, E=0
    DlPortWritePortUchar LPT_Control, &H5 '// RS=1, R/W=1, E=0

    sleep 1
End Sub















