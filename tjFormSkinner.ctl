VERSION 5.00
Begin VB.UserControl tjFormSkinner 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer timerBUTTON 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   240
   End
End
Attribute VB_Name = "tjFormSkinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************
'     Types
' ************************************************

' These will be used for a future update
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTRIPLE
End Type

' Used in determining the position of the pointer
Private Type POINT
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'*************************************************************
'   Constants
'*************************************************************
Private Const DT_LEFT = &H0
Private Const DT_TOP = &H0
Private Const DT_RIGHT = &H2
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
Private Const DT_NOCLIP = &H100
Private Const DT_CALCRECT = &H400

Private Const ALTERNATE = 1      ' ALTERNATE and WINDING are
Private Const WINDING = 2        ' constants for FillMode.
Private Const BLACKBRUSH = 4     ' Constant for brush type.
Private Const WHITE_BRUSH = 0    ' Constant for brush type.

Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4

Private Const BUTTONSIZE = 16

'*************************************************************
'   Required API Declarations
'*************************************************************
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINT) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

' *************************************************
'     Enurmators for background picture aligment
' *************************************************
Public Enum BackGroundPosition
    TopLeft = 1
    TopRight = 2
    TopCenter = 3
    BottomLeft = 4
    BottomRight = 5
    BottomCenter = 6
    Centered = 7
    Stretched = 8
End Enum


' ********************************************
'     Variables for property storage
' ********************************************
Private uc_ShowBorder As Boolean
Private uc_BorderColor As OLE_COLOR

Private uc_ShowCaptionArea As Boolean
Private uc_CaptionColorTo As OLE_COLOR
Private uc_CaptionColorFrom As OLE_COLOR
Private uc_Caption As String
Private uc_CaptionFont As StdFont
Private uc_CaptionColor As OLE_COLOR
Private uc_Icon As StdPicture
Private uc_IconSize As Integer

Private uc_ColorTo As OLE_COLOR
Private uc_ColorFrom As OLE_COLOR

Private uc_ShowMin As Boolean
Private uc_MinLeft As Long
Private uc_MinColorTo As OLE_COLOR
Private uc_MinColorFrom As OLE_COLOR
Private uc_MinColor As OLE_COLOR
Private uc_MinOverColor As OLE_COLOR

Private uc_ShowMax As Boolean
Private uc_MaxLeft As Long
Private uc_MaxColorTo As OLE_COLOR
Private uc_MaxColorFrom As OLE_COLOR
Private uc_MaxColor As OLE_COLOR
Private uc_MaxOverColor As OLE_COLOR

Private uc_ShowClose As Boolean
Private uc_CloseLeft As Long
Private uc_CloseColorTo As OLE_COLOR
Private uc_CloseColorFrom As OLE_COLOR
Private uc_CloseColor As OLE_COLOR
Private uc_CloseOverColor As OLE_COLOR

Private uc_BackGroundPicture As StdPicture
Private uc_BackGroundPicturePosition As BackGroundPosition

' ************************************************************
'    Other variables, but not used as property storage
' ************************************************************
Private uc_OverButton As Integer
Private uc_CaptionBottom As Long
Private uc_ButtonTop As Long

Private MouseDownForm As Integer
Private MouseDownPoint As POINT

' ************************************************************
'    Events
' ************************************************************
Public Event MinClicked()
Public Event MaxClicked()
Public Event CloseClicked()

'==========================================================================
' Properties
'==========================================================================
Public Property Let Top(ByRef New_Value As Long)
    UserControl.Extender.Top = New_Value
End Property
Public Property Get Top() As Long
    Top = UserControl.Extender.Top
End Property

Public Property Let Left(ByRef New_Value As Long)
    UserControl.Extender.Left = New_Value
End Property
Public Property Get Left() As Long
    Left = UserControl.Extender.Left
End Property

Public Property Let Height(ByRef New_Value As Long)
    UserControl.Extender.Height = New_Value
End Property
Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Let Width(ByRef New_Value As Long)
    UserControl.Width = New_Value
End Property
Public Property Get Width() As Long
    Width = UserControl.Width
End Property

Public Property Let BorderColor(ByRef new_BorderColor As OLE_COLOR)
    uc_BorderColor = new_BorderColor
    PropertyChanged "BorderColor"
    PaintFrame
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = uc_BorderColor
End Property

'------------- Close Button Colors ----------
Public Property Let CloseColor(ByRef new_Color As OLE_COLOR)
    uc_CloseColor = new_Color
    PropertyChanged "CloseColor"
    PaintFrame
End Property

Public Property Get CloseColor() As OLE_COLOR
    CloseColor = uc_CloseColor
End Property

Public Property Let CloseOverColor(ByRef new_Color As OLE_COLOR)
    uc_CloseOverColor = new_Color
    PropertyChanged "CloseOverColor"
    PaintFrame
End Property

Public Property Get CloseOverColor() As OLE_COLOR
    CloseOverColor = uc_CloseOverColor
End Property

Public Property Let CloseColorTo(ByRef new_Color As OLE_COLOR)
    uc_CloseColorTo = new_Color
    PropertyChanged "CloseColorTo"
    PaintFrame
End Property

Public Property Get CloseColorTo() As OLE_COLOR
    CloseColorTo = uc_CloseColorTo
End Property

Public Property Let CloseColorFrom(ByRef new_Color As OLE_COLOR)
    uc_CloseColorFrom = new_Color
    PropertyChanged "CloseColorFrom"
    PaintFrame
End Property

Public Property Get CloseColorFrom() As OLE_COLOR
    CloseColorFrom = uc_CloseColorFrom
End Property
'------------- Close Button Colors ----------

'------------- Max Button Colors ----------
Public Property Let MaxColor(ByRef new_Color As OLE_COLOR)
    uc_MaxColor = new_Color
    PropertyChanged "MaxColor"
    PaintFrame
End Property

Public Property Get MaxColor() As OLE_COLOR
    MaxColor = uc_MaxColor
End Property

Public Property Let MaxOverColor(ByRef new_Color As OLE_COLOR)
    uc_MaxOverColor = new_Color
    PropertyChanged "MaxOverColor"
    PaintFrame
End Property

Public Property Get MaxOverColor() As OLE_COLOR
    MaxOverColor = uc_MaxOverColor
End Property

Public Property Let MaxColorTo(ByRef new_Color As OLE_COLOR)
    uc_MaxColorTo = new_Color
    PropertyChanged "MaxColorTo"
    PaintFrame
End Property

Public Property Get MaxColorTo() As OLE_COLOR
    MaxColorTo = uc_MaxColorTo
End Property

Public Property Let MaxColorFrom(ByRef new_Color As OLE_COLOR)
    uc_MaxColorFrom = new_Color
    PropertyChanged "MaxColorFrom"
    PaintFrame
End Property

Public Property Get MaxColorFrom() As OLE_COLOR
    MaxColorFrom = uc_MaxColorFrom
End Property
'------------- Max Button Colors ----------

'------------- Min Button Colors ----------
Public Property Let MinColor(ByRef new_Color As OLE_COLOR)
    uc_MinColor = new_Color
    PropertyChanged "MinColor"
    PaintFrame
End Property

Public Property Get MinColor() As OLE_COLOR
    MinColor = uc_MinColor
End Property

Public Property Let MinOverColor(ByRef new_Color As OLE_COLOR)
    uc_MinOverColor = new_Color
    PropertyChanged "MinOverColor"
    PaintFrame
End Property

Public Property Get MinOverColor() As OLE_COLOR
    MinOverColor = uc_MinOverColor
End Property

Public Property Let MinColorTo(ByRef new_Color As OLE_COLOR)
    uc_MinColorTo = new_Color
    PropertyChanged "MinColorTo"
    PaintFrame
End Property

Public Property Get MinColorTo() As OLE_COLOR
    MinColorTo = uc_MinColorTo
End Property

Public Property Let MinColorFrom(ByRef new_Color As OLE_COLOR)
    uc_MinColorFrom = new_Color
    PropertyChanged "MinColorFrom"
    PaintFrame
End Property

Public Property Get MinColorFrom() As OLE_COLOR
    MinColorFrom = uc_MinColorFrom
End Property
'------------- Min Button Colors ----------


Public Property Let ShowBorder(ByRef new_ShowBorder As Boolean)
    uc_ShowBorder = new_ShowBorder
    PropertyChanged "ShowBorder"
    PaintFrame
End Property

Public Property Get ShowBorder() As Boolean
    ShowBorder = uc_ShowBorder
End Property

Public Property Let ShowCaptionArea(ByRef new_ShowCaptionArea As Boolean)
    uc_ShowCaptionArea = new_ShowCaptionArea
    PropertyChanged "ShowCaptionArea"
    PaintFrame
End Property

Public Property Get ShowCaptionArea() As Boolean
    ShowCaptionArea = uc_ShowCaptionArea
End Property

Public Property Let CaptionColorTo(ByRef new_Color As OLE_COLOR)
    uc_CaptionColorTo = new_Color
    PropertyChanged "CaptionColorTo"
    PaintFrame
End Property

Public Property Get CaptionColorTo() As OLE_COLOR
    CaptionColorTo = uc_CaptionColorTo
End Property

Public Property Let CaptionColorFrom(ByRef new_Color As OLE_COLOR)
    uc_CaptionColorFrom = new_Color
    PropertyChanged "CaptionColorFrom"
    PaintFrame
End Property

Public Property Get CaptionColorFrom() As OLE_COLOR
    CaptionColorFrom = uc_CaptionColorFrom
End Property

Public Property Let Caption(ByRef new_Caption As String)
    uc_Caption = new_Caption
    PropertyChanged "Caption"
    PaintFrame
End Property

Public Property Get Caption() As String
    Caption = uc_Caption
End Property

Public Property Let CaptionColor(ByRef new_Color As OLE_COLOR)
    uc_CaptionColor = new_Color
    PropertyChanged "CaptionColor"
    PaintFrame
End Property

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = uc_CaptionColor
End Property

Public Property Set Font(ByRef new_font As StdFont)
    SetFont new_font
    PropertyChanged "Font"
    PaintFrame
End Property

Public Property Get Font() As StdFont
    Set Font = uc_CaptionFont
End Property

Public Property Get Icon() As StdPicture
    Set Icon = uc_Icon
End Property

Public Property Set Icon(ByVal New_Picture As StdPicture)
    Set uc_Icon = New_Picture
    PropertyChanged "Icon"
    PaintFrame
End Property


Public Property Get BackGroundPicture() As StdPicture
    Set BackGroundPicture = uc_BackGroundPicture
End Property

Public Property Set BackGroundPicture(ByVal New_Picture As StdPicture)
    Set uc_BackGroundPicture = New_Picture
    PropertyChanged "BackGroundPicture"
    PaintFrame
End Property


Public Property Let BackGroundPicturePosition(ByRef new_Position As BackGroundPosition)
    uc_BackGroundPicturePosition = new_Position
    PropertyChanged "BackgroundPicturePosition"
    PaintFrame
End Property

Public Property Get BackGroundPicturePosition() As BackGroundPosition
    BackGroundPicturePosition = uc_BackGroundPicturePosition
End Property


Public Property Let ColorTo(ByRef new_Color As OLE_COLOR)
    uc_ColorTo = new_Color
    PropertyChanged "ColorTo"
    PaintFrame
End Property

Public Property Get ColorTo() As OLE_COLOR
    ColorTo = uc_ColorTo
End Property

Public Property Let ColorFrom(ByRef new_Color As OLE_COLOR)
    uc_ColorFrom = new_Color
    PropertyChanged "ColorFrom"
    PaintFrame
End Property

Public Property Get ColorFrom() As OLE_COLOR
    ColorFrom = uc_ColorFrom
End Property

Public Property Let ShowMinButton(ByRef new_ShowButton As Boolean)
    uc_ShowMin = new_ShowButton
    PropertyChanged "ShowMinButton"
    PaintFrame
End Property

Public Property Get ShowMinButton() As Boolean
    ShowMinButton = uc_ShowMin
End Property

Public Property Let ShowMaxButton(ByRef new_ShowButton As Boolean)
    uc_ShowMax = new_ShowButton
    PropertyChanged "ShowMaxButton"
    PaintFrame
End Property

Public Property Get ShowMaxButton() As Boolean
    ShowMaxButton = uc_ShowMax
End Property

Public Property Let ShowCloseButton(ByRef new_ShowButton As Boolean)
    uc_ShowClose = new_ShowButton
    PropertyChanged "ShowCloseButton"
    PaintFrame
End Property

Public Property Get ShowCloseButton() As Boolean
    ShowCloseButton = uc_ShowClose
End Property

Public Property Get IconSize() As Integer
    IconSize = uc_IconSize
End Property

Public Property Let IconSize(ByVal New_Value As Integer)
    uc_IconSize = New_Value
    PropertyChanged "IconSize"
    PaintFrame
End Property

Private Sub SetFont(ByRef new_font As StdFont)
    With uc_CaptionFont
        .Bold = new_font.Bold
        .Italic = new_font.Italic
        .Name = new_font.Name
        .SIZE = new_font.SIZE
    End With
    Set UserControl.Font = uc_CaptionFont
End Sub


'==========================================================================
' API Functions and subroutines
'==========================================================================

' full version of APILine
Private Sub APILineEx(lhdcEx As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lcolor As Long)

    'Use the API LineTo for Fast Drawing
    Dim pt As POINT
    Dim hPen As Long, hPenOld As Long
    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, X1, Y1, pt
    LineTo lhdcEx, X2, Y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen
End Sub

Private Function APIRectangle(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal H As Long, Optional lcolor As OLE_COLOR = -1) As Long
    'Draw an api rectangle
    Dim hPen As Long, hPenOld As Long
    Dim R
    Dim pt As POINT
    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(hDC, hPen)
    MoveToEx hDC, x, y, pt
    LineTo hDC, x + w, y
    LineTo hDC, x + w, y + H
    LineTo hDC, x, y + H
    LineTo hDC, x, y
    SelectObject hDC, hPenOld
    DeleteObject hPen
End Function

Private Sub DrawGradientEx(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, ByVal x As Long, ByVal y As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional blnVertical = True)
    
    'Draw a Vertical or horizontal Gradient in the current HDC
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    If blnVertical Then
        dR = (sR - eR) / Y2
        dG = (sG - eG) / Y2
        dB = (sB - eB) / Y2
        For ni = 1 To Y2 - 1
            APILineEx lhdcEx, x, y + ni, X2, y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
    Else
        dR = (sR - eR) / X2
        dG = (sG - eG) / X2
        dB = (sB - eB) / X2
        For ni = 1 To X2 - 1
            APILineEx lhdcEx, x + ni, y, x + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
    End If
End Sub

'Blend two colors
Private Function BlendColors(ByVal lcolor1 As Long, ByVal lcolor2 As Long)
    BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)
End Function

' ******************************************
'    Timer will only be enabled if min, max
'    or close button is shown
'
'    Determines if the button is over a
'    button. If so, the button color is
'    changed.
'
'    The variable uc_OverButton is set to
'    determine which button the mouse is
'    over, this is to prevent repetitive
'    calls.
' ******************************************
Private Sub timerBUTTON_Timer()
    Dim mousePt As POINT

    GetCursorPos mousePt
    If WindowFromPoint(mousePt.x, mousePt.y) <> UserControl.hwnd Then
        ' mouse is not over the control at all
        If uc_OverButton <> 0 Then
            uc_OverButton = 0

            If uc_ShowClose Then PaintCloseButton uc_CloseColor, uc_CloseColorFrom, uc_CloseColorTo
            If uc_ShowMax Then PaintMaxButton uc_MaxColor, uc_MaxColorFrom, uc_MaxColorTo
            If uc_ShowMin Then PaintMinButton uc_MinColor, uc_MinColorFrom, uc_MinColorTo
            UserControl.Refresh
            
        End If
    Else
        ' Mouse is over the control

        ' Adjust position in relation to the control,
        ' easier to determine if the mouse is over a
        ' button
        If UserControl.Parent.BorderStyle = 0 Then
            mousePt.x = mousePt.x - (((UserControl.Parent.Left / Screen.TwipsPerPixelX) + (UserControl.Extender.Left / Screen.TwipsPerPixelX)))
            mousePt.y = mousePt.y - (((UserControl.Parent.Top / Screen.TwipsPerPixelY) + (UserControl.Extender.Top / Screen.TwipsPerPixelY)))
        Else
            mousePt.x = mousePt.x - (((UserControl.Parent.Left / Screen.TwipsPerPixelX) + (UserControl.Extender.Left / Screen.TwipsPerPixelX)) + 4)
            mousePt.y = mousePt.y - (((UserControl.Parent.Top / Screen.TwipsPerPixelY) + (UserControl.Extender.Top / Screen.TwipsPerPixelY)) + 23)
        End If
                
        If mousePt.x >= uc_CloseLeft And mousePt.x <= (uc_CloseLeft + BUTTONSIZE) And mousePt.y >= uc_ButtonTop And mousePt.y <= (uc_ButtonTop + BUTTONSIZE) And uc_ShowClose Then
            ' over close button
            If uc_OverButton <> 1 Then
                uc_OverButton = 1

                If uc_ShowMin Then PaintMinButton uc_MinColor, uc_MinColorFrom, uc_MinColorTo
                If uc_ShowMax Then PaintMaxButton uc_MaxColor, uc_MaxColorFrom, uc_MaxColorTo
                PaintCloseButton uc_CloseOverColor, uc_CloseColorFrom, uc_CloseColorTo
                UserControl.Refresh
            End If
        ElseIf mousePt.x >= uc_MaxLeft And mousePt.x <= (uc_MaxLeft + BUTTONSIZE) And mousePt.y >= uc_ButtonTop And mousePt.y <= (uc_ButtonTop + BUTTONSIZE) And uc_ShowMax Then
            ' over max button
            If uc_OverButton <> 2 Then
                uc_OverButton = 2

                If uc_ShowMin Then PaintMinButton uc_MinColor, uc_MinColorFrom, uc_MinColorTo
                If uc_ShowClose Then PaintCloseButton uc_CloseColor, uc_CloseColorFrom, uc_CloseColorTo
                PaintMaxButton uc_MaxOverColor, uc_MaxColorFrom, uc_MaxColorTo
                UserControl.Refresh
            
            End If
        ElseIf mousePt.x >= uc_MinLeft And mousePt.x <= (uc_MinLeft + BUTTONSIZE) And mousePt.y >= uc_ButtonTop And mousePt.y <= (uc_ButtonTop + BUTTONSIZE) And uc_ShowMin Then
            ' over min button
            If uc_OverButton <> 3 Then
                uc_OverButton = 3

                If uc_ShowClose Then PaintCloseButton uc_CloseColor, uc_CloseColorFrom, uc_CloseColorTo
                If uc_ShowMax Then PaintMaxButton uc_MaxColor, uc_MaxColorFrom, uc_MaxColorTo
                PaintMinButton uc_MinOverColor, uc_MinColorFrom, uc_MinColorTo
                UserControl.Refresh
                
            End If
        ElseIf uc_OverButton <> 0 Then
            ' not over any buttons, but was, but is also over the control
            uc_OverButton = 0

            If uc_ShowClose Then PaintCloseButton uc_CloseColor, uc_CloseColorFrom, uc_CloseColorTo
            If uc_ShowMax Then PaintMaxButton uc_MaxColor, uc_MaxColorFrom, uc_MaxColorTo
            If uc_ShowMin Then PaintMinButton uc_MinColor, uc_MinColorFrom, uc_MinColorTo
            UserControl.Refresh
            
        End If
    End If
    
End Sub

Private Sub UserControl_Click()
    ' used in determining if a button has been pressed.
    
    ' if no buttons shown, it does not matter
    If (uc_ShowMin Or uc_ShowMax Or uc_ShowClose) = False Then Exit Sub
    
    Dim mousePt As POINT

    GetCursorPos mousePt
    If WindowFromPoint(mousePt.x, mousePt.y) <> UserControl.hwnd Then
        ' may need to set
    Else
        ' see if mouse over close button
        If UserControl.Parent.BorderStyle = 0 Then
            mousePt.x = mousePt.x - (((UserControl.Parent.Left / Screen.TwipsPerPixelX) + (UserControl.Extender.Left / Screen.TwipsPerPixelX)))
            mousePt.y = mousePt.y - (((UserControl.Parent.Top / Screen.TwipsPerPixelY) + (UserControl.Extender.Top / Screen.TwipsPerPixelY)))
        Else
            mousePt.x = mousePt.x - (((UserControl.Parent.Left / Screen.TwipsPerPixelX) + (UserControl.Extender.Left / Screen.TwipsPerPixelX)) + 4)
            mousePt.y = mousePt.y - (((UserControl.Parent.Top / Screen.TwipsPerPixelY) + (UserControl.Extender.Top / Screen.TwipsPerPixelY)) + 23)
        End If

        
        If mousePt.x >= uc_CloseLeft And mousePt.x <= (uc_CloseLeft + BUTTONSIZE) And mousePt.y >= uc_ButtonTop And mousePt.y <= (uc_ButtonTop + BUTTONSIZE) And uc_ShowClose Then
          ' over close button
            RaiseEvent CloseClicked
        ElseIf mousePt.x >= uc_MaxLeft And mousePt.x <= (uc_MaxLeft + BUTTONSIZE) And mousePt.y >= uc_ButtonTop And mousePt.y <= (uc_ButtonTop + BUTTONSIZE) And uc_ShowMax Then
            ' over max button
            RaiseEvent MaxClicked
        ElseIf mousePt.x >= uc_MinLeft And mousePt.x <= (uc_MinLeft + BUTTONSIZE) And mousePt.y >= uc_ButtonTop And mousePt.y <= (uc_ButtonTop + BUTTONSIZE) And uc_ShowMin Then
            ' over min button
            RaiseEvent MinClicked
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Set uc_Font = New StdFont
    Set UserControl.Font = uc_Font
    uc_BorderColor = vbBlack
    uc_ColorTo = 1
    uc_ColorFrom = 1
    uc_ShowBorder = True
    uc_IconSize = 16
    uc_BackGroundPicturePosition = TopLeft
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' use in moving the form which holds the control
    ' caption area must be shown
    ' the parent form must also be borderless
    
    Dim tempY As Long
    
    If MouseDownForm = 0 Then
        ' if no caption area, cannot move
        If uc_ShowCaptionArea = False Then Exit Sub
        Call GetCursorPos(MouseDownPoint)

        tempY = MouseDownPoint.y - (((UserControl.Parent.Top / Screen.TwipsPerPixelY) + (UserControl.Extender.Top / Screen.TwipsPerPixelY)))
        
        ' mouse pointer is beyond caption area, cannot move
        If tempY >= uc_CaptionBottom Then Exit Sub
    End If
    MouseDownForm = 1

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Z As POINT
    Dim newX As Long
    Dim newY As Long
    
    On Local Error Resume Next
    
    ' cannot move if mouse button is not down
    '             if no cpation area
    '             if over a button
    '             if parent form has a border
    If MouseDownForm <> 1 Then Exit Sub
    If uc_ShowCaptionArea = False Then Exit Sub
    If uc_OverButton <> 0 Then Exit Sub
    If UserControl.Parent.BorderStyle <> 0 Then Exit Sub
    
    ' moves form by getting current position and moving
    ' it based upon the previous position
    Call GetCursorPos(Z)
    
    newX = (Z.x - MouseDownPoint.x) * Screen.TwipsPerPixelX
    newY = (Z.y - MouseDownPoint.y) * Screen.TwipsPerPixelY
    
    UserControl.Parent.Top = UserControl.Parent.Top + newY
    UserControl.Parent.Left = UserControl.Parent.Left + newX
    
    Call GetCursorPos(MouseDownPoint)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' mouse button no longer pressed down
    MouseDownForm = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        uc_ShowBorder = .ReadProperty("ShowBorder", True)
        uc_BorderColor = .ReadProperty("BorderColor", vbBlack)
        uc_ShowCaptionArea = .ReadProperty("ShowCaptionArea", False)
        uc_CaptionColorTo = .ReadProperty("CaptionColorTo", Ambient.BackColor)
        uc_CaptionColorFrom = .ReadProperty("CaptionColorFrom", Ambient.BackColor)
        uc_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        uc_CaptionColor = .ReadProperty("CaptionColor", vbBlack)
        Set uc_CaptionFont = .ReadProperty("Font", Ambient.Font)
        Set uc_Icon = .ReadProperty("Icon", Nothing)
        Set uc_BackGroundPicture = .ReadProperty("BackGroundPicture", Nothing)
        uc_BackGroundPicturePosition = .ReadProperty("BackGroundPicturePosition", TopLeft)
        uc_IconSize = .ReadProperty("IconSize", 16)

        uc_ColorTo = .ReadProperty("ColorTo", Ambient.BackColor)
        uc_ColorFrom = .ReadProperty("ColorFrom", Ambient.BackColor)
        
        uc_ShowMin = .ReadProperty("ShowMinButton", False)
        uc_MinColorTo = .ReadProperty("MinColorTo", Ambient.BackColor)
        uc_MinColorFrom = .ReadProperty("MinColorFrom", Ambient.BackColor)
        uc_MinColor = .ReadProperty("MinColor", vbBlack)
        uc_MinOverColor = .ReadProperty("MinOverColor", vbRed)
        
        uc_ShowMax = .ReadProperty("ShowMaxButton", False)
        uc_MaxColorTo = .ReadProperty("MaxColorTo", Ambient.BackColor)
        uc_MaxColorFrom = .ReadProperty("MaxColorFrom", Ambient.BackColor)
        uc_MaxColor = .ReadProperty("MaxColor", vbBlack)
        uc_MaxOverColor = .ReadProperty("MaxOverColor", vbRed)
        
        uc_ShowClose = .ReadProperty("ShowCloseButton", False)
        uc_CloseColorTo = .ReadProperty("CloseColorTo", Ambient.BackColor)
        uc_CloseColorFrom = .ReadProperty("CloseColorFrom", Ambient.BackColor)
        uc_CloseColor = .ReadProperty("CloseColor", vbBlack)
        uc_CloseOverColor = .ReadProperty("CloseOverColor", vbRed)
 
    End With
    Set UserControl.Font = uc_CaptionFont
    timerBUTTON.Enabled = uc_ShowMin Or uc_ShowMax Or uc_ShowClose
    PaintFrame
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ShowBorder", uc_ShowBorder, True
        .WriteProperty "BorderColor", uc_BorderColor, vbBlack
        .WriteProperty "ShowCaptionArea", uc_ShowCaptionArea, False
        .WriteProperty "CaptionColorTo", uc_CaptionColorTo, Ambient.BackColor
        .WriteProperty "CaptionColorFrom", uc_CaptionColorFrom, Ambient.BackColor
        .WriteProperty "Caption", uc_Caption, Ambient.DisplayName
        .WriteProperty "CaptionColor", uc_CaptionColor, vbBlack
        .WriteProperty "Icon", uc_Icon, Nothing
        .WriteProperty "IconSize", uc_IconSize, 16
        .WriteProperty "Font", uc_CaptionFont, Ambient.Font
        .WriteProperty "ColorTo", uc_ColorTo, Ambient.BackColor
        .WriteProperty "ColorFrom", uc_ColorFrom, Ambient.BackColor
        
        .WriteProperty "ShowMinButton", uc_ShowMin, False
        .WriteProperty "MinColor", uc_MinColor, vbBlack
        .WriteProperty "MinOverColor", uc_MinOverColor, vbRed
        .WriteProperty "MinColorFrom", uc_MinColorFrom, Ambient.BackColor
        .WriteProperty "MinColorTo", uc_MinColorTo, Ambient.BackColor
        
        .WriteProperty "ShowMaxButton", uc_ShowMax, False
        .WriteProperty "MaxColor", uc_MaxColor, vbBlack
        .WriteProperty "MaxOverColor", uc_MaxOverColor, vbRed
        .WriteProperty "MaxColorFrom", uc_MaxColorFrom, Ambient.BackColor
        .WriteProperty "MaxColorTo", uc_MaxColorTo, Ambient.BackColor
        
        .WriteProperty "ShowCloseButton", uc_ShowClose, False
        .WriteProperty "CloseColor", uc_CloseColor, vbBlack
        .WriteProperty "CloseOverColor", uc_CloseOverColor, vbRed
        .WriteProperty "CloseColorFrom", uc_CloseColorFrom, Ambient.BackColor
        .WriteProperty "CloseColorTo", uc_CloseColorTo, Ambient.BackColor
        
        .WriteProperty "BackGroundPicture", uc_BackGroundPicture, Nothing
        .WriteProperty "BackGroundPicturePosition", uc_BackGroundPicturePosition, TopLeft
    End With
End Sub
Private Sub EraseRegion()
    Dim hRgn As Long
    'Creates second region to fill with color.
    hRgn = CreateRoundRectRgn(0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, 0&)
    SetWindowRgn UserControl.hwnd, hRgn, True
    'delete our elliptical region
    DeleteObject hRgn
    UserControl.FillStyle = 0
End Sub

'==================
' Main drawing sub
'==================
Private Sub PaintFrame()
    Dim RC As RECT
    Dim ucHDC As Long
    Dim CaptionHeight As Long
    Dim CaptionUpper As Long
    Dim CaptionLower As Long
    Dim CaptionIndent As Integer
    
    ' clear region area
    EraseRegion
    
    'Clear user control
    UserControl.Cls
    ucHDC = UserControl.hDC
    
    
    ' get caption area heights
    CaptionHeight = UserControl.TextHeight(uc_Caption) + 10
    If Not (uc_Icon Is Nothing) Then
        If CaptionHeight < uc_IconSize + 10 Then CaptionHeight = uc_IconSize + 10
    End If
    CaptionUpper = CaptionHeight / 2
    CaptionLower = CaptionHeight - CaptionUpper
    

    
    ' paint if caption area is true
    If uc_ShowCaptionArea Then
        If uc_ShowBorder Then
            'all the border
            SetRect RC, 0&, 0&, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
            APIRectangle ucHDC, RC.Top, RC.Left, RC.Right, RC.Bottom, uc_BorderColor
                     
            SetRect RC, 0&, 0&, UserControl.ScaleWidth - 2, CaptionHeight
            APIRectangle ucHDC, RC.Top, RC.Left, RC.Right, RC.Bottom, uc_BorderColor
            
            SetRect RC, 0, 1, UserControl.ScaleWidth - 2, CaptionUpper + 1
            DrawGradientEx ucHDC, BlendColors(uc_CaptionColorTo, vbWhite), BlendColors(uc_CaptionColorFrom, vbWhite), RC.Top, RC.Left, RC.Right, RC.Bottom
            
            SetRect RC, CaptionUpper, 1, UserControl.ScaleWidth - 2, CaptionLower
            DrawGradientEx ucHDC, BlendColors(uc_CaptionColorFrom, vbWhite), BlendColors(uc_CaptionColorTo, vbWhite), RC.Top, RC.Left, RC.Right, RC.Bottom
            
            If uc_BackGroundPicture Is Nothing Or uc_BackGroundPicturePosition <> Stretched Then
                SetRect RC, CaptionHeight, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - (CaptionHeight + 2)
                DrawGradientEx ucHDC, BlendColors(uc_ColorFrom, vbWhite), BlendColors(uc_ColorTo, vbWhite), RC.Top, RC.Left, RC.Right, RC.Bottom
            End If
        Else
            ' caption area
            SetRect RC, 0, 0, UserControl.ScaleWidth, CaptionUpper + 1
            DrawGradientEx ucHDC, BlendColors(uc_CaptionColorTo, vbWhite), BlendColors(uc_CaptionColorFrom, vbWhite), RC.Top, RC.Left, RC.Right, RC.Bottom
            
            SetRect RC, CaptionUpper, 0, UserControl.ScaleWidth, CaptionLower
            DrawGradientEx ucHDC, BlendColors(uc_CaptionColorFrom, vbWhite), BlendColors(uc_CaptionColorTo, vbWhite), RC.Top, RC.Left, RC.Right, RC.Bottom
            
            If uc_BackGroundPicture Is Nothing Or uc_BackGroundPicturePosition <> Stretched Then
                SetRect RC, CaptionHeight, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
                DrawGradientEx ucHDC, BlendColors(uc_ColorFrom, vbWhite), BlendColors(uc_ColorTo, vbWhite), RC.Top, RC.Left, RC.Right, RC.Bottom
            End If
        End If
    Else
        If uc_ShowBorder Then
            ' has a border
            SetRect RC, 0&, 0&, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
            APIRectangle ucHDC, RC.Top, RC.Left, RC.Right, RC.Bottom, uc_BorderColor
            
            If uc_BackGroundPicture Is Nothing Or uc_BackGroundPicturePosition <> Stretched Then
                SetRect RC, 0, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
                DrawGradientEx ucHDC, BlendColors(uc_ColorFrom, vbWhite), BlendColors(uc_ColorTo, vbWhite), RC.Top, RC.Left, RC.Right, RC.Bottom
            End If
        Else
            ' no border
            If uc_BackGroundPicture Is Nothing Or uc_BackGroundPicturePosition <> Stretched Then
                SetRect RC, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
                DrawGradientEx ucHDC, BlendColors(uc_ColorFrom, vbWhite), BlendColors(uc_ColorTo, vbWhite), RC.Top, RC.Left, RC.Right, RC.Bottom
            End If
        End If
    End If
    
    
    ' display background picture if one
    If Not (uc_BackGroundPicture Is Nothing) Then
        Dim bX As Long
        Dim bY As Long
        Dim bW As Long
        Dim bH As Long
        
        Dim pW As Long
        Dim pH As Long
        
        Dim bCH As Long
        
        pW = uc_BackGroundPicture.Width / Screen.TwipsPerPixelX
        pH = uc_BackGroundPicture.Height / Screen.TwipsPerPixelX
        
        bCH = CaptionHeight
        
        If uc_ShowCaptionArea = False Then bCH = 0
        
        Select Case uc_BackGroundPicturePosition
            Case TopLeft
                bX = 0
                bY = bCH
            Case TopRight
                bX = UserControl.ScaleWidth - pW
                bY = bCH
            Case TopCenter
                bX = (UserControl.ScaleWidth - pW) / 2
                bY = bCH
            Case BottomLeft
                bX = 0
                bY = UserControl.ScaleHeight - pH
            Case BottomRight
                bX = UserControl.ScaleWidth - pW
                bY = UserControl.ScaleHeight - pH
            Case BottomCenter
                bX = (UserControl.ScaleWidth - pW) / 2
                bY = UserControl.ScaleHeight - pH
            Case Centered
                bX = (UserControl.ScaleWidth - pW) / 2
                bY = (UserControl.ScaleHeight - pH) / 2
            Case Stretched
                bX = 0
                bY = bCH
        End Select
        bW = pW
        bH = pH
        If uc_BackGroundPicturePosition = Stretched Then
            bW = UserControl.ScaleWidth
            bH = UserControl.ScaleHeight - bCH
        End If
        
        If uc_ShowBorder Then
            Select Case uc_BackGroundPicturePosition
                Case TopLeft
                    bX = bX + 2
                    bY = bY + 2
                Case TopRight
                    bX = bX - 3
                    bY = bY + 2
                Case TopCenter
                    bY = bY + 2
                Case BottomLeft
                    bX = bX + 2
                    bY = bY - 3
                Case BottomRight
                    bX = bX - 3
                    bY = bY - 3
                Case BottomCenter
                    bY = bY - 3
                Case Stretched
                    bX = bX + 1
                    bY = bY + 1
            End Select
        End If
        
        If uc_BackGroundPicturePosition = Stretched Then
            If uc_ShowBorder Then
                bW = bW - 3
                bH = bH - 3
            End If
        End If
        
        UserControl.PaintPicture uc_BackGroundPicture, bX, bY, bW, bH
    End If
    
    ' show icon if set
    CaptionIndent = 5
    If Not (uc_Icon Is Nothing) Then
        UserControl.PaintPicture uc_Icon, CaptionIndent, (CaptionHeight - uc_IconSize) / 2, uc_IconSize, uc_IconSize
        CaptionIndent = uc_IconSize + 10
    End If
    
    ' show caption
    If Len(uc_Caption) <> 0 Then
    
        UserControl.ForeColor = uc_CaptionColor
        
        SetRect RC, CaptionIndent, (CaptionHeight - UserControl.TextHeight(uc_Caption)) / 2, UserControl.ScaleWidth - 20, UserControl.TextHeight(uc_Caption) + 10
        DrawTextEx ucHDC, uc_Caption, Len(uc_Caption), RC, DT_LEFT Or DT_WORDBREAK Or DT_VCENTER, ByVal 0&
    
    End If
    
    ' calculate button position variables
    If uc_ShowCaptionArea Then
        uc_ButtonTop = (CaptionHeight - BUTTONSIZE) / 2
        uc_CaptionBottom = CaptionHeight
    Else
        uc_ButtonTop = 5
    End If
    
    Dim buttonLeft As Long
    
    buttonLeft = UserControl.ScaleWidth - (BUTTONSIZE + 5)
    If uc_ShowClose Then
        uc_CloseLeft = buttonLeft
        buttonLeft = buttonLeft - (BUTTONSIZE + 5)
    End If
    If uc_ShowMax Then
        uc_MaxLeft = buttonLeft
        buttonLeft = buttonLeft - (BUTTONSIZE + 5)
    End If
    If uc_ShowMin Then
        uc_MinLeft = buttonLeft
    End If
    
    ' paint the buttons
    PaintButtons
    
    ' set button timer, if a button is displayed
    timerBUTTON.Enabled = uc_ShowMin Or uc_ShowMax Or uc_ShowClose
    
End Sub

Private Sub PaintButtons()
    If uc_ShowClose Then
        PaintCloseButton uc_CloseColor, uc_CloseColorFrom, uc_CloseColorTo
    End If
    
    If uc_ShowMax Then
        PaintMaxButton uc_MaxColor, uc_MaxColorFrom, uc_MaxColorTo
    End If

    If uc_ShowMin Then
        PaintMinButton uc_MinColor, uc_MinColorFrom, uc_MinColorTo
    End If

End Sub

Private Sub PaintButton(buttonLeft As Long, ColorFrom As Long, ColorTo As Long)
    ' creates the button borders and gradient infill
    
    ' draw border of the button
    ' top line in white
    APILineEx UserControl.hDC, buttonLeft, uc_ButtonTop, buttonLeft + BUTTONSIZE, uc_ButtonTop, vbWhite
    
    ' left line in white
    APILineEx UserControl.hDC, buttonLeft, uc_ButtonTop, buttonLeft, uc_ButtonTop + BUTTONSIZE, vbWhite
    
    ' right line in black
    APILineEx UserControl.hDC, buttonLeft + BUTTONSIZE, uc_ButtonTop, buttonLeft + BUTTONSIZE, uc_ButtonTop + BUTTONSIZE, vbBlack
    
    ' bottom line in black
    APILineEx UserControl.hDC, buttonLeft, uc_ButtonTop + BUTTONSIZE, buttonLeft + BUTTONSIZE + 1, uc_ButtonTop + BUTTONSIZE, vbBlack
    
    ' draw the gradient colors inside the box
    DrawGradientEx UserControl.hDC, BlendColors(ColorFrom, vbWhite), BlendColors(ColorTo, vbWhite), buttonLeft + 1, uc_ButtonTop, (buttonLeft + BUTTONSIZE), BUTTONSIZE ' (uc_ButtonTop + BUTTONSIZE) - 2

End Sub

Private Sub PaintMinButton(buttonColor As Long, ColorFrom As Long, ColorTo As Long)
    ' invert gradient colors if mouse over the button
    If buttonColor = uc_MinColor Then
        PaintButton uc_MinLeft, ColorFrom, ColorTo
    Else
        PaintButton uc_MinLeft, ColorTo, ColorFrom
    End If

    APILineEx UserControl.hDC, uc_MinLeft + 3, (uc_ButtonTop + BUTTONSIZE) - 4, (uc_MinLeft + BUTTONSIZE) - 2, (uc_ButtonTop + BUTTONSIZE) - 4, buttonColor
    APILineEx UserControl.hDC, uc_MinLeft + 3, (uc_ButtonTop + BUTTONSIZE) - 3, (uc_MinLeft + BUTTONSIZE) - 2, (uc_ButtonTop + BUTTONSIZE) - 3, buttonColor
End Sub

Private Sub PaintMaxButton(buttonColor As Long, ColorFrom As Long, ColorTo As Long)
    ' invert gradient colors if mouse over the button
    If buttonColor = uc_MaxColor Then
        PaintButton uc_MaxLeft, ColorFrom, ColorTo
    Else
        PaintButton uc_MaxLeft, ColorTo, ColorFrom
    End If
    APIRectangle UserControl.hDC, uc_MaxLeft + 3, uc_ButtonTop + 3, BUTTONSIZE - 6, BUTTONSIZE - 6, buttonColor
    APILineEx UserControl.hDC, uc_MaxLeft + 3, uc_ButtonTop + 4, (uc_MaxLeft + BUTTONSIZE) - 2, uc_ButtonTop + 4, buttonColor
End Sub

Private Sub PaintCloseButton(buttonColor As Long, ColorFrom As Long, ColorTo As Long)
    ' invert gradient colors if mouse over the button
    If buttonColor = uc_CloseColor Then
        PaintButton uc_CloseLeft, ColorFrom, ColorTo
    Else
        PaintButton uc_CloseLeft, ColorTo, ColorFrom
    End If
    ' top left to bottom right
    APILineEx UserControl.hDC, uc_CloseLeft + 4, uc_ButtonTop + 4, (uc_CloseLeft + BUTTONSIZE) - 3, (uc_ButtonTop + BUTTONSIZE) - 3, buttonColor
    APILineEx UserControl.hDC, uc_CloseLeft + 4, uc_ButtonTop + 5, (uc_CloseLeft + BUTTONSIZE) - 3, (uc_ButtonTop + BUTTONSIZE) - 2, buttonColor
    
    ' top right to bottom left
    APILineEx UserControl.hDC, (uc_CloseLeft + BUTTONSIZE) - 4, uc_ButtonTop + 4, uc_CloseLeft + 3, (uc_ButtonTop + BUTTONSIZE) - 3, buttonColor
    APILineEx UserControl.hDC, (uc_CloseLeft + BUTTONSIZE) - 4, uc_ButtonTop + 5, uc_CloseLeft + 3, (uc_ButtonTop + BUTTONSIZE) - 2, buttonColor
    
End Sub

Private Sub UserControl_Resize()
    PaintFrame
End Sub


' ########## For Future Update - For Enable/Disble #########
'Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False, Optional ByVal XPBlend As Boolean = False)
'
'    If DstW = 0 Or DstH = 0 Then Exit Sub
'
'    Dim B As Long, H As Long, F As Long, I As Long, newW As Long
'    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
'    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
'    Dim Data1() As RGBTRIPLE, Data2() As RGBTRIPLE
'    Dim Info As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
'    Dim hOldOb As Long
'    Dim SrcDC As Long, tObj As Long, ttt As Long
'
'    SrcDC = CreateCompatibleDC(hDC)
'
'    If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
'    If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)
'    If SrcPic.Type = 1 Then 'check if it's an icon or a bitmap
'        tObj = SelectObject(SrcDC, SrcPic)
'    Else
'        Dim hBrush As Long
'        tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
'        hBrush = CreateSolidBrush(TransColor) 'MaskColor)
'        DrawIconEx SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, &H1 Or &H2
'        DeleteObject hBrush
'    End If
'
'    TmpDC = CreateCompatibleDC(SrcDC)
'    Sr2DC = CreateCompatibleDC(SrcDC)
'    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
'    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
'    TmpObj = SelectObject(TmpDC, TmpBmp)
'    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
'    ReDim Data1(DstW * DstH * 3 - 1)
'    ReDim Data2(UBound(Data1))
'    With Info.bmiHeader
'        .biSize = Len(Info.bmiHeader)
'        .biWidth = DstW
'        .biHeight = DstH
'        .biPlanes = 1
'        .biBitCount = 24
'    End With
'
'    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
'    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
'    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
'    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0
'
'    If BrushColor > 0 Then
'        BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
'        BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
'        BrushRGB.rgbRed = BrushColor And &HFF
'    End If
'    useMask = True
'    If Not useMask Then TransColor = -1
'
'    newW = DstW - 1
'
'    For H = 0 To DstH - 1
'        F = H * DstW
'        For B = 0 To newW
'            I = F + B
'            If GetNearestColor(hDC, CLng(Data2(I).rgbRed) + 256& * Data2(I).rgbGreen + 65536 * Data2(I).rgbBlue) <> TransColor Then
'                With Data1(I)
'                    If BrushColor > -1 Then
'                        If MonoMask Then
'                            If (CLng(Data2(I).rgbRed) + Data2(I).rgbGreen + Data2(I).rgbBlue) <= 384 Then Data1(I) = BrushRGB
'                        Else
'                            Data1(I) = BrushRGB
'                        End If
'                    Else
'                        If isGreyscale Then
'                            gCol = CLng(Data2(I).rgbRed * 0.3) + Data2(I).rgbGreen * 0.59 + Data2(I).rgbBlue * 0.11
'                            .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
'                        Else
'                            If XPBlend Then
'                                .rgbRed = (CLng(.rgbRed) + Data2(I).rgbRed * 2) \ 3
'                                .rgbGreen = (CLng(.rgbGreen) + Data2(I).rgbGreen * 2) \ 3
'                                .rgbBlue = (CLng(.rgbBlue) + Data2(I).rgbBlue * 2) \ 3
'                            Else
'                                Data1(I) = Data2(I)
'                            End If
'                        End If
'                    End If
'                End With
'            End If
'        Next B
'    Next H
'
'    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0
'
'    Erase Data1, Data2
'    DeleteObject SelectObject(TmpDC, TmpObj)
'    DeleteObject SelectObject(Sr2DC, Sr2Obj)
'    If SrcPic.Type = 3 Then DeleteObject SelectObject(SrcDC, tObj)
'    DeleteDC TmpDC: DeleteDC Sr2DC
'    DeleteObject tObj: DeleteDC SrcDC
'End Sub
'
