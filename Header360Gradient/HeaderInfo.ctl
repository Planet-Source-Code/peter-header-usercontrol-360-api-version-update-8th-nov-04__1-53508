VERSION 5.00
Begin VB.UserControl HeaderInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000006&
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   255
   ToolboxBitmap   =   "HeaderInfo.ctx":0000
End
Attribute VB_Name = "HeaderInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
 
' Inspiration from Jim K's 'PictureBox as Info Header'
' http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=53222&lngWId=1

' Credit's
' --------
' Keith LaVolpe.
' Image/caption drawing routines, alignment calculations(caption/image)
' imported &/or modified from 'La Volpe Buttons vH.1'.
' Keith's logic of drawing to an offscreenDc then BitBlt to the controlDc to prevent
' flicker has also been adopted (after some much appreciated help from him).

' Bug Fixes Posted 13th May (Memory leaks).
' 360 Degree Gradient posted 8th Nov.

Private Type PointAPI                ' general use. Typically used for cursor location
    X                             As Long
    Y                             As Long
End Type
Private Type PointSng                'Internal Point structure
    X   As Single                    'Uses Singles for more precision.
    Y   As Single
End Type
Private Type RECT                    ' used to set/ref boundaries of a rectangle
    Left                          As Long
    Top                           As Long
    Right                         As Long
    Bottom                        As Long
End Type
Private Type RectAPI    'API Rect structure
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type
Private Type BITMAP                  ' used to determine if an image is a bitmap
    bmType                        As Long
    bmWidth                       As Long
    bmHeight                      As Long
    bmWidthBytes                  As Long
    bmPlanes                      As Integer
    bmBitsPixel                   As Integer
    bmBits                        As Long
End Type
Private Type LOGFONT                 ' used to create fonts
    lfHeight                      As Long
    lfWidth                       As Long
    lfEscapement                  As Long
    lfOrientation                 As Long
    lfWeight                      As Long
    lfItalic                      As Byte
    lfUnderline                   As Byte
    lfStrikeOut                   As Byte
    lfCharSet                     As Byte
    lfOutPrecision                As Byte
    lfClipPrecision               As Byte
    lfQuality                     As Byte
    lfPitchAndFamily              As Byte
    lfFaceName                    As String * 32
End Type
Private Type ICONINFO                ' used to determine if image is an icon
    fIcon                         As Long
    xHotSpot                      As Long
    yHotSpot                      As Long
    hbmMask                       As Long
    hbmColor                      As Long
End Type
Private Type HeaderDCInfo            ' used to manage the drawing DC
    hDC                           As Long       ' the temporary DC handle
    OldBitmap                     As Long       ' the original bitmap of the DC
    OldPen                        As Long       ' the original pen of the DC
    OldBrush                      As Long       ' the original brush of the DC
    OldFont                       As Long       ' the original font of the DC
End Type
Private Type HeaderProperties        'used to store header values
    hAngle                        As Integer    ' gradient Angle
    hCaption                      As String     ' header caption
    hCaptionAlign                 As AlignmentConstants ' caption alignment (3 options)
    hCaptionStyle                 As CaptionEffectCnsts ' raised/sunken/default
    hGradientBackStyle            As BackStyleCnsts
    hCRect                        As RECT       ' cached caption's bounding rectangle
    hUCRect                       As RECT       ' cached control bounding rectangle
    hSgc                          As Long       ' Start Gradient Color
    hEgc                          As Long       ' End Gradient Color
    hBorderClr                    As Long       ' border color
    hBorderVis                    As Boolean    ' border visible
    hCnrSize                      As CrnrSzeCnsts 'Cnr size small/large
    hCnrShape                     As CrnrShpeCnsts 'Cnr shape
    hCapShadow                    As Long
End Type
Private Type ImageProperties         'used to store image values
    Image                         As StdPicture ' button image
    TransImage                    As Long
    TransSize                     As PointAPI
    Align                         As ImagePlacementConstants ' image alignment (6 options)
    Size                          As Integer    ' image size (5 options)
    iRect                         As RECT       ' cached image's bounding rectangle
    SourceSize                    As PointAPI   ' cached source image dimensions
    Type                          As Long       ' cached source image type (bmp/ico)
End Type

' Custom CONSTANTS
' ================================================================================
' DrawText API/Image/Gradient Constants
Private Const CI_BITMAP           As Long = &H0
Private Const CI_ICON             As Long = &H1
Private Const DT_CALCRECT         As Long = &H400
Private Const DT_CENTER           As Long = &H1
Private Const DT_WORDBREAK        As Long = &H10
Private Const PS_SOLID            As Long = 0
Private Const RGN_DIFF            As Integer = 4
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

'  Variables
' ================================================================================
Private HeaderDC                  As HeaderDCInfo       ' menu DC for drawing menu items
Private myImage                   As ImageProperties    ' cached image properties
Private myProps                   As HeaderProperties   ' cached ctrl properties

'  Enumerators
' ================================================================================
'Gradient Finish Style
Public Enum BackStyleCnsts
    d_Opaque = 0
    d_Transparent = 1
End Enum
#If False Then
Private d_Opaque, d_Transparent
#End If
'Used for corner shape
Public Enum CrnrShpeCnsts
   d_Square = 0
   d_Rounded = 1
   d_RoundedTop = 2
End Enum
#If False Then
Private d_Square, d_Rounded, d_RoundedTop
#End If
'Used for Corner Size
Public Enum CrnrSzeCnsts
   d_Small = 3
   d_Large = 5
End Enum
#If False Then
Private d_Small, d_Large
#End If
' Used to set/reset HDC objects
Private Enum ColorObjects
    cObj_Brush = 0
    cObj_Pen = 1
    cObj_Text = 2
End Enum
#If False Then
Private cObj_Brush, cObj_Pen, cObjText
#End If
' caption styles
Public Enum CaptionEffectCnsts
    d_Default = 0
    d_Sunken = 1
    d_Raised = 2
End Enum
#If False Then
Private d_Default, d_Sunken, d_Raised
#End If
' image alignment
Public Enum ImagePlacementConstants
    d_LeftEdge = 0
    d_LeftOfCaption = 1
    d_RightEdge = 2
    d_RightOfCaption = 3
    d_TopCenter = 4
    d_BottomCenter = 5
End Enum
#If False Then
Private d_LeftEdge, d_LeftOfCaption, d_RightEdge, d_RightOfCaption, d_TopCenter, d_BottomCenter
#End If
' image sizes
Public Enum ImageSizeConstants
    d_16x16 = 0
    d_24x24 = 1
    d_32x32 = 2
    d_Fill_Stretch = 3
    d_Fill_ScaleUpDown = 4
End Enum
#If False Then
Private d_16x16, d_24x24, d_32x32, d_Fill_Stretch, d_Fill_ScaleUpDown
#End If

'  API
' ================================================================================
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
                                                 ByVal hSrcRgn1 As Long, _
                                                 ByVal hSrcRgn2 As Long, _
                                                 ByVal nCombineMode As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, _
                                                 ByVal imageType As Long, _
                                                 ByVal newWidth As Long, _
                                                 ByVal newHeight As Long, _
                                                 ByVal lFlags As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
                                                   ByVal nHeight As Long, _
                                                   ByVal nPlanes As Long, _
                                                   ByVal nBitCount As Long, _
                                                   lpBits As Any) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" _
                                                             (lpLogFont As LOGFONT) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, _
                                                    ByVal y1 As Long, _
                                                    ByVal X2 As Long, _
                                                    ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
                                                  ByVal xLeft As Long, _
                                                  ByVal yTop As Long, _
                                                  ByVal hIcon As Long, _
                                                  ByVal cxWidth As Long, _
                                                  ByVal cyWidth As Long, _
                                                  ByVal istepIfAniCur As Long, _
                                                  ByVal hbrFlickerFreeDraw As Long, _
                                                  ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Integer
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, _
                                                   piconinfo As ICONINFO) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               lpPoint As PointAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, _
                                                    ByVal hPalette As Long, _
                                                    ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, _
                                                 ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, _
                                                 ByVal nMapMode As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, _
                                               ByVal x1 As Long, _
                                               ByVal y1 As Long, _
                                               ByVal X2 As Long, _
                                               ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, _
                                                   ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, _
                                                 ByVal X As Long, _
                                                 ByVal Y As Long, _
                                                 ByVal nWidth As Long, _
                                                 ByVal nHeight As Long, _
                                                 ByVal hSrcDC As Long, _
                                                 ByVal xSrc As Long, _
                                                 ByVal ySrc As Long, _
                                                 ByVal nSrcWidth As Long, _
                                                 ByVal nSrcHeight As Long, _
                                                 ByVal dwRop As Long) As Long

Public Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, laRetColors() As Long) As Long
Dim lIdx    As Long
Dim lRed    As Long
Dim lGrn    As Long
Dim lBlu    As Long
Dim fRedStp As Single
Dim fGrnStp As Single
Dim fBluStp As Single
' Creates an array of colors blending from Color1 to Color2 in
' lSteps number of steps. Returns the count and fills the laRetColors() array.

' Stop possible error
If lSteps < 2 Then lSteps = 2
    
' Extract Red, Blue and Green values from the start and end colors.
lRed = (lColor1 And &HFF&)
lGrn = (lColor1 And &HFF00&) / &H100
lBlu = (lColor1 And &HFF0000) / &H10000
    
' Find the amount of change for each color element per color change.
fRedStp = Div(CSng((lColor2 And &HFF&) - lRed), CSng(lSteps))
fGrnStp = Div(CSng(((lColor2 And &HFF00&) / &H100&) - lGrn), CSng(lSteps))
fBluStp = Div(CSng(((lColor2 And &HFF0000) / &H10000) - lBlu), CSng(lSteps))
    
' Create the colors
ReDim laRetColors(lSteps - 1)
laRetColors(0) = lColor1            'First Color
laRetColors(lSteps - 1) = lColor2   'Last Color

For lIdx = 1 To lSteps - 2          'All Colors between
    laRetColors(lIdx) = CLng(lRed + (fRedStp * CSng(lIdx))) + _
        (CLng(lGrn + (fGrnStp * CSng(lIdx))) * &H100&) + _
        (CLng(lBlu + (fBluStp * CSng(lIdx))) * &H10000)
Next lIdx
    
' Return number of colors in array
BlendColors = lSteps
End Function

Private Sub CheckStuff() ' Check Colors and Angle
Dim lIdx As Long

' Colors
If myProps.hGradientBackStyle = d_Transparent Then
   myProps.hEgc = ConvertColor(UserControl.Parent.BackColor)   'Parent.BackColor
End If
If myProps.hSgc < 0 Then
   lIdx = (myProps.hSgc And Not &H80000000)
   If lIdx >= 0 And lIdx <= 24 Then
       myProps.hSgc = GetSysColor(lIdx)
   End If
End If
If myProps.hEgc < 0 Then
   lIdx = (myProps.hEgc And Not &H80000000)
   If lIdx >= 0 And lIdx <= 24 Then
      myProps.hEgc = GetSysColor(lIdx)
   End If
End If

' Angle
' Angles are counter-clockwise and may be
' any Single value from 0 to 359.999999999.
'  135  90 45
'     \ | /
' 180 --o-- 0
'     / | \
'  235 270 315
' Check angle to ensure between 0 and 359.999999999
myProps.hAngle = CDbl(myProps.hAngle) - Int(Int(CDbl(myProps.hAngle) / 360#) * 360#)
End Sub

Private Sub CalculateBoundingRects(bNormalizeImage As Boolean)
' Routine measures and places the rectangles to draw
' the caption and image on the control. The results
' are cached so this routine doesn't need to run
' every time the button is redrawn/painted
Dim cRect As RECT, tRect As RECT, iRect As RECT
Dim imgOffset As RECT, bImgWidthAdj As Boolean, bImgHeightAdj As Boolean
Dim lEdge As Long, adjWidth As Long

adjWidth = myProps.hUCRect.Right

With myImage
    If (.SourceSize.X + .SourceSize.Y) > 0 Then
        ' image in use, calculations for image rectangle
        If .Size < 33 Then
           Select Case .Align
             Case d_LeftEdge, d_LeftOfCaption
                  imgOffset.Left = .Size
                  bImgWidthAdj = True
             Case d_RightEdge, d_RightOfCaption
                  imgOffset.Right = .Size
                  bImgWidthAdj = True
             Case d_TopCenter
                  imgOffset.Top = .Size
                  bImgHeightAdj = True
             Case d_BottomCenter
                  imgOffset.Bottom = .Size
                  bImgHeightAdj = True
           End Select
        End If
    End If
End With

If Len(myProps.hCaption) Then
    Dim sCaption As String  ' note: Replace$ not compatible with VB5
    sCaption = Replace$(myProps.hCaption, "||", vbNewLine)
    ' calculate total available button width available for text
    cRect.Right = adjWidth - 8 - (myImage.Size * Abs(CInt(bImgWidthAdj)))
    cRect.Bottom = ScaleHeight - 8 - (myImage.Size * Abs(CInt(bImgHeightAdj = True And myImage.Align > d_RightOfCaption)))

    ' calculate size of rectangle to hold that text, using multiline flag
    DrawText HeaderDC.hDC, sCaption, Len(sCaption), cRect, DT_CALCRECT Or DT_WORDBREAK
    If myProps.hCaptionStyle Then
       cRect.Right = cRect.Right + 2
       cRect.Bottom = cRect.Bottom + 2
    End If
End If

' now calculate the position of the text rectangle
If Len(myProps.hCaption) Then
   tRect = cRect
   Select Case myProps.hCaptionAlign
     Case vbLeftJustify
          OffsetRect tRect, imgOffset.Left + lEdge + 4 + (Abs(CInt(imgOffset.Left > 0) * 4)), 0
     Case vbRightJustify
          OffsetRect tRect, adjWidth - imgOffset.Right - 4 - cRect.Right - (Abs(CInt(imgOffset.Right > 0) * 4)), 0
     Case vbCenter
          If imgOffset.Left > 0 And myImage.Align = d_LeftOfCaption Then
             OffsetRect tRect, (adjWidth - (imgOffset.Left + cRect.Right + 4)) \ 2 + lEdge + 4 + imgOffset.Left, 0
           Else
             If imgOffset.Right > 0 And myImage.Align = d_RightOfCaption Then
                OffsetRect tRect, (adjWidth - (imgOffset.Right + cRect.Right + 4)) \ 2 + lEdge, 0
              Else
                OffsetRect tRect, ((adjWidth - (imgOffset.Left + imgOffset.Right)) - cRect.Right) \ 2 + lEdge + imgOffset.Left, 0
             End If
          End If
   End Select
End If

If (myImage.SourceSize.X + myImage.SourceSize.Y) > 0 Then
    ' finalize image rectangle position
   Select Case myImage.Align
     Case d_LeftEdge
          iRect.Left = lEdge + 4
     Case d_LeftOfCaption
          If Len(myProps.hCaption) Then
             iRect.Left = tRect.Left - 4 - imgOffset.Left
           Else
             iRect.Left = lEdge + 4
          End If
     Case d_RightOfCaption
          If Len(myProps.hCaption) Then
             iRect.Left = tRect.Right + 4
           Else
             iRect.Left = adjWidth - 4 - imgOffset.Right
          End If
     Case d_RightEdge
          iRect.Left = adjWidth - 4 - imgOffset.Right
     Case d_TopCenter
          iRect.Top = (ScaleHeight - (cRect.Bottom + imgOffset.Top)) \ 2
          OffsetRect tRect, 0, iRect.Top + 2 + imgOffset.Top
     Case d_BottomCenter
          iRect.Top = (ScaleHeight - (cRect.Bottom + imgOffset.Bottom)) \ 2 + cRect.Bottom + 4
          OffsetRect tRect, 0, iRect.Top - 2 - cRect.Bottom
   End Select
   If myImage.Align < d_TopCenter Then
      OffsetRect tRect, 0, (ScaleHeight - cRect.Bottom) \ 2
      iRect.Top = (ScaleHeight - myImage.Size) \ 2
    Else
      iRect.Left = (adjWidth - myImage.Size) \ 2 + lEdge
   End If
   iRect.Right = iRect.Left + myImage.Size
   iRect.Bottom = iRect.Top + myImage.Size
 Else
   OffsetRect tRect, 0, (ScaleHeight - cRect.Bottom) \ 2
End If

' sanity checks
With tRect
    If .Top < 4 Then .Top = 4
    If .Left < 4 + lEdge Then .Left = 4 + lEdge
    If .Right > adjWidth - 4 Then .Right = adjWidth - 4
    If .Bottom > ScaleHeight - 5 Then .Bottom = ScaleHeight - 5
End With
myProps.hCRect = tRect

Select Case myImage.Size
  Case Is < 33
       With iRect
           If .Top < 4 Then .Top = 4
           If .Left < 4 + lEdge Then .Left = 4 + lEdge
           If .Right > adjWidth - 4 Then .Right = adjWidth - 4
           If .Bottom > ScaleHeight - 5 Then .Bottom = ScaleHeight - 5
       End With
  Case 40     ' stretch
       SetRect iRect, 1, 1, ScaleWidth - 1, ScaleHeight - 1
      bNormalizeImage = True
  Case Else   ' scale
       If (myImage.SourceSize.X + myImage.SourceSize.Y) > 0 Then
           With iRect
               ScaleImage adjWidth - 12, ScaleHeight - 12, cRect.Right, cRect.Bottom
                .Left = (adjWidth - cRect.Right) \ 2 + lEdge
                .Top = (ScaleHeight - cRect.Bottom) \ 2
                .Right = .Left + cRect.Right
                .Bottom = .Top + cRect.Bottom
           End With
          bNormalizeImage = True
       End If
End Select
myImage.iRect = iRect
With iRect
    If bNormalizeImage Then NormalizeImage .Right - .Left, .Bottom - .Top
End With
End Sub

Private Function ConvertColor(tColor As Long) As Long
'Converts VB color constants to real color values
If tColor < 0 Then
   ConvertColor = GetSysColor(tColor And &HFF&)
 Else
   ConvertColor = tColor
End If
End Function

Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
' Divides dNumer by dDenom if dDenom <> 0
' Eliminates 'Division By Zero' error.
If dDenom <> 0 Then
   Div = dNumer / dDenom
 Else
   Div = 0
End If
End Function

Private Sub DrawCaption()
Dim tRect As RECT
Dim lColor As Long
Dim sCaption As String
Dim bColor As Long

With myProps
    bColor = .hCapShadow
    ' set the rectangle & may be adjusted a little later
    tRect = .hCRect
    ' note Replace$ not compatible with VB5
    sCaption = Replace$(.hCaption, "||", vbNewLine)
    ' Setting text colors and offsets
    ' get the right forecolor to use
    lColor = ConvertColor(UserControl.ForeColor)
    If (.hCaptionStyle And UserControl.Enabled = True) Then
        ' drawing raised/sunken caption styles
        Dim shadeOffset As Integer
        'Select the caption style
        If .hCaptionStyle = d_Raised Then shadeOffset = 40 Else shadeOffset = -40
        'First shadow
        SetHeaderColors True, HeaderDC.hDC, cObj_Text, ShadeColor(.hCapShadow, shadeOffset, False)
        OffsetRect tRect, -1, 0
        DrawText HeaderDC.hDC, sCaption, Len(sCaption), tRect, DT_WORDBREAK Or Choose(.hCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
        'Second shadow
        SetHeaderColors True, HeaderDC.hDC, cObj_Text, ShadeColor(.hCapShadow, -shadeOffset, False)
        OffsetRect tRect, 2, 2
        DrawText HeaderDC.hDC, sCaption, Len(sCaption), tRect, DT_WORDBREAK Or Choose(.hCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
        OffsetRect tRect, -1, -1
    End If
End With
SetHeaderColors True, HeaderDC.hDC, cObj_Text, lColor
DrawText HeaderDC.hDC, sCaption, Len(sCaption), tRect, DT_WORDBREAK Or Choose(myProps.hCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
End Sub

Private Sub DrawBorder()
Dim r As Single
Dim g As Single
Dim b As Single
Dim ur As Integer
Dim ub As Integer
Dim hBrush As Long
    
If Not myProps.hBorderVis Then Exit Sub

With HeaderDC
     hBrush = CreateSolidBrush(myProps.hBorderClr)
     FrameRect .hDC, myProps.hUCRect, hBrush
     DeleteObject hBrush
     
     With myProps.hUCRect
          ur = .Right
          ub = .Bottom
     End With
     
     If myProps.hCnrShape = d_Square Then GoTo Done
     Select Case myProps.hCnrSize
       Case d_Small
        'Left top corner
         SetPixel .hDC, 1, 2, RGB(r, g, b)
         SetPixel .hDC, 1, 1, RGB(r, g, b)
         SetPixel .hDC, 2, 1, RGB(r, g, b)
        'right top corner
         SetPixel .hDC, ur - 3, 1, RGB(r, g, b)
         SetPixel .hDC, ur - 2, 1, RGB(r, g, b)
         SetPixel .hDC, ur - 2, 2, RGB(r, g, b)
       Case d_Large
        'Left top corner
         SetPixel .hDC, 1, 4, RGB(r, g, b)
         SetPixel .hDC, 1, 3, RGB(r, g, b)
         SetPixel .hDC, 2, 2, RGB(r, g, b)
         SetPixel .hDC, 3, 1, RGB(r, g, b)
         SetPixel .hDC, 4, 1, RGB(r, g, b)
        'right top corner
         SetPixel .hDC, ur - 5, 1, RGB(r, g, b)
         SetPixel .hDC, ur - 4, 1, RGB(r, g, b)
         SetPixel .hDC, ur - 3, 2, RGB(r, g, b)
         SetPixel .hDC, ur - 2, 3, RGB(r, g, b)
         SetPixel .hDC, ur - 2, 4, RGB(r, g, b)
     End Select
     
     If myProps.hCnrShape = d_RoundedTop Then GoTo Done
     Select Case myProps.hCnrSize
       Case d_Small
        'left bottom corner
         SetPixel .hDC, 1, ub - 3, RGB(r, g, b)
         SetPixel .hDC, 1, ub - 2, RGB(r, g, b)
         SetPixel .hDC, 2, ub - 2, RGB(r, g, b)
         SetPixel .hDC, 3, ub - 1, RGB(r, g, b)
        'right bottom corner
         SetPixel .hDC, ur - 2, ub - 3, RGB(r, g, b)
         SetPixel .hDC, ur - 2, ub - 2, RGB(r, g, b)
         SetPixel .hDC, ur - 3, ub - 2, RGB(r, g, b)
         
       Case d_Large
        'left bottom corner
         SetPixel .hDC, 1, ub - 5, RGB(r, g, b)
         SetPixel .hDC, 1, ub - 4, RGB(r, g, b)
         SetPixel .hDC, 2, ub - 3, RGB(r, g, b)
         SetPixel .hDC, 3, ub - 2, RGB(r, g, b)
         SetPixel .hDC, 4, ub - 2, RGB(r, g, b)
        'right bottom corner
         SetPixel .hDC, ur - 2, ub - 5, RGB(r, g, b)
         SetPixel .hDC, ur - 2, ub - 4, RGB(r, g, b)
         SetPixel .hDC, ur - 3, ub - 3, RGB(r, g, b)
         SetPixel .hDC, ur - 4, ub - 2, RGB(r, g, b)
         SetPixel .hDC, ur - 5, ub - 2, RGB(r, g, b)
     End Select
Done:
End With
End Sub

Private Sub DrawIcon(iRect As RECT)
' Routine will draw the button image
Dim imgWidth  As Long
Dim imgHeight As Long
Dim rcImage   As RECT
    
If (myImage.SourceSize.X + myImage.SourceSize.Y) = 0 Then
   Exit Sub
End If
With iRect
    If myImage.TransImage = 0 Then
       NormalizeImage .Right - .Left, .Bottom - .Top
    End If
    imgWidth = .Right - .Left
    imgHeight = .Bottom - .Top
End With
' destination rectangle for drawing on the DC
If myImage.Type = CI_ICON Then
' draw icon directly onto the temporary DC
' for icons, we can draw directly on the destination DC
   DrawIconEx HeaderDC.hDC, iRect.Left, iRect.Top, myImage.Image.Handle, imgWidth, imgHeight, 0, 0, &H3
 Else
' draw transparent bitmap onto the temporary DC
   DrawTransparentBitmap HeaderDC.hDC, iRect, myImage.TransImage, rcImage, , imgWidth, imgHeight
End If
End Sub

Private Sub DrawGradient()
Dim bDone       As Boolean
Dim iIncX       As Integer
Dim iIncY       As Integer
Dim hPen        As Long
Dim hOldPen     As Long
Dim lIdx        As Long
Dim lPointCnt   As Long
Dim laColors()  As Long
Dim lWidth      As Long
Dim lHeight     As Long
Dim lRet        As Long
Dim fMovX       As Single
Dim fMovY       As Single
Dim fDist       As Single
Dim fAngle      As Single
Dim fLongSide   As Single
Dim uTmpPt      As PointAPI
Dim uaPts()     As PointAPI
Dim uaTmpPts()  As PointSng
'On Error GoTo ErrCtrl

' Check Colors and Angle
CheckStuff

lWidth = myProps.hUCRect.Right - 1
lHeight = myProps.hUCRect.Bottom - 1
    
'Start with center of rect
ReDim uaTmpPts(2)
uaTmpPts(2).X = Int(lWidth / 2)
uaTmpPts(2).Y = Int(lHeight / 2)
    
'Calc distance to furthest edge as if rect were square
fLongSide = IIf(lWidth > lHeight, lWidth, lHeight)
fDist = (Sqr((fLongSide ^ 2) + (fLongSide ^ 2)) + 2) / 2
    
'Create points to the left and the right at a 0º angle (horizontal)
uaTmpPts(0).X = uaTmpPts(2).X - fDist
uaTmpPts(0).Y = uaTmpPts(2).Y
uaTmpPts(1).X = uaTmpPts(2).X + fDist
uaTmpPts(1).Y = uaTmpPts(2).Y
    
'Lines will be drawn perpendicular to mfAngle so
'add 90º and correct for 360º wrap
fAngle = CDbl(myProps.hAngle + 90) - Int(Int(CDbl(myProps.hAngle + 90) / 360#) * 360#)
    
'Rotate second and third points to fAngle
Call RotatePoint(uaTmpPts(2), uaTmpPts(0), fAngle)
Call RotatePoint(uaTmpPts(2), uaTmpPts(1), fAngle)
    
'We now have a line that crosses the center and
'two sides of the rect at the correct angle.
    
'Calc the starting quadrant, direction of and amount of first move
'(fMovX, fMovY moves line from center to starting edge)
'and direction of each incremental move (iIncX, iIncY).
Select Case myProps.hAngle
   Case 0 To 90
        'Left Bottom
        If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
            'Move line to left edge; Draw left to right
            fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
            fMovY = 0
            iIncX = 1
            iIncY = 0
         Else
            'Move line to bottom edge; Draw bottom to top
            fMovX = 0
            fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
            iIncX = 0
            iIncY = -1
        End If
   Case 90 To 180
        'Right Bottom
        If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
            'Move line to right edge; Draw right to left
            fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
            fMovY = 0
            iIncX = -1
            iIncY = 0
         Else
            'Move line to bottom edge; Draw bottom to top
            fMovX = 0
            fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
            iIncX = 0
            iIncY = -1
        End If
   Case 180 To 270
        'Right Top
        If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
            'Move line to right edge; Draw right to left
            fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
            fMovY = 0
            iIncX = -1
            iIncY = 0
         Else
            'Move line to top edge; Draw top to bottom
            fMovX = 0
            fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
            iIncX = 0
            iIncY = 1
        End If
   Case Else   '(270 to 360)
        'Left Top
        If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
            'Move line to left edge; Draw left to right
            fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
            fMovY = 0
            iIncX = 1
            iIncY = 0
         Else
            'Move line to top edge; Draw top to bottom
            fMovX = 0
            fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
            iIncX = 0
            iIncY = 1
        End If
End Select
    
'At this point we could calculate where the lines will cross the rect edges, but
'this would slow things down. The picObj clipping region will take care of this.
    
'Start with 1000 points and add more if needed. This increases
'speed by not re-dimming the array in each loop.
ReDim uaPts(999)
    
'Set the first two points in the array
uaPts(0).X = uaTmpPts(0).X + fMovX
uaPts(0).Y = uaTmpPts(0).Y + fMovY
uaPts(1).X = uaTmpPts(1).X + fMovX
uaPts(1).Y = uaTmpPts(1).Y + fMovY
    
lIdx = 2
'Create the rest of the points by incrementing both points
'on each line iIncX, iIncY from the previous line's points.
'Where we stop depends on the direction of travel.
'We'll continue until both points in a set reach the end.
While Not bDone
    uaPts(lIdx).X = uaPts(lIdx - 2).X + iIncX
    uaPts(lIdx).Y = uaPts(lIdx - 2).Y + iIncY
    lIdx = lIdx + 1
    Select Case True
        Case iIncX > 0  'Moving Left to Right
            bDone = uaPts(lIdx - 1).X > lWidth And uaPts(lIdx - 2).X > lWidth
        Case iIncX < 0  'Moving Right to Left
            bDone = uaPts(lIdx - 1).X < 0 And uaPts(lIdx - 2).X < 0
        Case iIncY > 0  'Moving Top to Bottom
            bDone = uaPts(lIdx - 1).Y > lHeight And uaPts(lIdx - 2).Y > lHeight
        Case iIncY < 0  'Moving Bottom to Top
            bDone = uaPts(lIdx - 1).Y < 0 And uaPts(lIdx - 2).Y < 0
    End Select
    If (lIdx Mod 1000) = 0 Then
        ReDim Preserve uaPts(UBound(uaPts) + 1000)
    End If
Wend
    
'Free excess memory (may have 1001 points dimmed to 2000)
ReDim Preserve uaPts(lIdx - 1)
    
'Create the array of colors blending from mlColor1 to mlColor2
lRet = BlendColors(myProps.hSgc, myProps.hEgc, lIdx / 2, laColors)
    
'Now draw each line in it's own color
For lIdx = 0 To UBound(uaPts) - 1 Step 2
    'Move to next point
    lRet = MoveToEx(HeaderDC.hDC, uaPts(lIdx).X, uaPts(lIdx).Y, uTmpPt)
    'Create the colored pen and select it into the DC
    hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx / 2)))
    hOldPen = SelectObject(HeaderDC.hDC, hPen)
    'Draw the line
    lRet = LineTo(HeaderDC.hDC, uaPts(lIdx + 1).X, uaPts(lIdx + 1).Y)
    'Get the pen back out of the DC and destroy it
    lRet = SelectObject(HeaderDC.hDC, hOldPen)
    lRet = DeleteObject(hPen)
Next lIdx
    
ErrCtrl:
'Free the memory
Erase laColors
Erase uaPts
Erase uaTmpPts

On Error GoTo 0
End Sub

Private Sub DrawInfoHeader(Optional b As Boolean = False)
DrawGradient
DrawCaption
DrawBorder
DrawIcon myImage.iRect
MakeRegion
GetSetOffDC False
End Sub

Private Sub DrawTransparentBitmap(ByVal lHDCdest As Long, _
                                  destRect As RECT, _
                                  ByVal lBMPsource As Long, _
                                  bmpRect As RECT, _
                                  Optional lMaskColor As Long = -1, _
                                  Optional lNewBmpCx As Long, _
                                  Optional lNewBmpCy As Long)
Dim lMask2Use      As Long      'COLORREF
Dim lBmMask        As Long
Dim lBmAndMem      As Long
Dim lBmColor       As Long
Dim lBmObjectOld   As Long
Dim lBmMemOld      As Long
Dim lBmColorOld    As Long
Dim lHDCMem        As Long
Dim lHDCscreen     As Long
Dim lHDCsrc        As Long
Dim lHDCMask       As Long
Dim lHDCcolor      As Long
Dim X              As Long
Dim Y              As Long
Dim srcX           As Long
Dim srcY           As Long
Dim lRatio(0 To 1) As Single
Const DSna         As Long = &H220326

' =====================================================================
' A pretty good transparent bitmap maker I use in several projects
' Modified here to remove stuff I wont use (i.e., Flipping/Rotating images)
' =====================================================================
lHDCscreen = GetDC(0&)
lHDCsrc = CreateCompatibleDC(lHDCscreen)     'Create a temporary HDC compatible to the Destination HDC
SelectObject lHDCsrc, lBMPsource             'Select the bitmap
srcX = myImage.TransSize.X ' lNewBmpCx                  'Get width of bitmap
srcY = myImage.TransSize.Y ' lNewBmpCy                'Get height of bitmap
With bmpRect
    If .Right = 0 Then
        .Right = srcX
     Else
       srcX = .Right - .Left
    End If
    If .Bottom = 0 Then
        .Bottom = srcY
     Else
       srcY = .Bottom - .Top
    End If
End With
With destRect
    If (.Right) = 0 Then
       X = lNewBmpCx
     Else
       X = (.Right - .Left)
    End If
    If (.Bottom) = 0 Then
       Y = lNewBmpCy
     Else
       Y = (.Bottom - .Top)
    End If
End With
If lNewBmpCx > X Or lNewBmpCy > Y Then
   lRatio(0) = (X / lNewBmpCx)
   lRatio(1) = (Y / lNewBmpCy)
   If lRatio(1) < lRatio(0) Then
      lRatio(0) = lRatio(1)
   End If
   lNewBmpCx = lRatio(0) * lNewBmpCx
   lNewBmpCy = lRatio(0) * lNewBmpCy
   Erase lRatio
End If

lMask2Use = ConvertColor(GetPixel(lHDCsrc, 0, 0))
'Create some DCs & bitmaps
lHDCMask = CreateCompatibleDC(lHDCscreen)
lHDCMem = CreateCompatibleDC(lHDCscreen)
lHDCcolor = CreateCompatibleDC(lHDCscreen)
lBmColor = CreateCompatibleBitmap(lHDCscreen, srcX, srcY)
lBmAndMem = CreateCompatibleBitmap(lHDCscreen, X, Y)
lBmMask = CreateBitmap(srcX, srcY, 1&, 1&, ByVal 0&)
lBmColorOld = SelectObject(lHDCcolor, lBmColor)
lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
lBmObjectOld = SelectObject(lHDCMask, lBmMask)
ReleaseDC 0&, lHDCscreen
' ====================== Start working here ======================
SetMapMode lHDCMem, GetMapMode(lHDCdest)
SelectPalette lHDCMem, 0, True
RealizePalette lHDCMem
BitBlt lHDCMem, 0&, 0&, X, Y, lHDCdest, destRect.Left, destRect.Top, vbSrcCopy
SelectPalette lHDCcolor, 0, True
RealizePalette lHDCcolor
SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
SetTextColor lHDCcolor, GetTextColor(lHDCsrc)
BitBlt lHDCcolor, 0&, 0&, srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, vbSrcCopy
SetBkColor lHDCcolor, lMask2Use
SetTextColor lHDCcolor, vbWhite
BitBlt lHDCMask, 0&, 0&, srcX, srcY, lHDCcolor, 0&, 0&, vbSrcCopy
SetTextColor lHDCcolor, vbBlack
SetBkColor lHDCcolor, vbWhite
BitBlt lHDCcolor, 0, 0, srcX, srcY, lHDCMask, 0, 0, DSna
StretchBlt lHDCMem, 0, 0, lNewBmpCx, lNewBmpCy, lHDCMask, 0&, 0&, srcX, srcY, vbSrcAnd
StretchBlt lHDCMem, 0&, 0&, lNewBmpCx, lNewBmpCy, lHDCcolor, 0, 0, srcX, srcY, vbSrcPaint
BitBlt lHDCdest, destRect.Left, destRect.Top, X, Y, lHDCMem, 0&, 0&, vbSrcCopy
'Delete memory bitmaps & DCs
DeleteObject SelectObject(lHDCcolor, lBmColorOld)
DeleteObject SelectObject(lHDCMask, lBmObjectOld)
DeleteObject SelectObject(lHDCMem, lBmMemOld)
DeleteDC lHDCMem
DeleteDC lHDCMask
DeleteDC lHDCcolor
DeleteDC lHDCsrc
End Sub

Private Sub GetGDIMetrics(ByVal sObject As String)
' This routine caches information we don't want to keep gathering every time a button is redrawn.
Select Case sObject
Case "UC"
    GetClientRect UserControl.hwnd, myProps.hUCRect
Case "Font"
    ' called when font is changed or control is initialized
    Dim newFont As LOGFONT
    With newFont
        .lfCharSet = 1
        .lfFaceName = UserControl.Font.Name & Chr$(0)
        .lfHeight = (UserControl.Font.Size * -20) / Screen.TwipsPerPixelY
        .lfWeight = UserControl.Font.Weight
        .lfItalic = Abs(CInt(UserControl.Font.Italic))
        .lfStrikeOut = Abs(CInt(UserControl.Font.Strikethrough))
        .lfUnderline = Abs(CInt(UserControl.Font.Underline))
    End With
    If HeaderDC.OldFont Then
        DeleteObject SelectObject(HeaderDC.hDC, CreateFontIndirect(newFont))
    Else
        HeaderDC.OldFont = SelectObject(HeaderDC.hDC, CreateFontIndirect(newFont))
    End If
Case "Picture"
    ' get key image information
    Dim bmpInfo As BITMAP
    Dim icoInfo As ICONINFO
    With myImage
        If .Image Is Nothing Then
            If .TransImage Then DeleteObject .TransImage
            .SourceSize.X = 0
            .SourceSize.Y = 0
        Else
            GetGDIObject .Image.Handle, LenB(bmpInfo), bmpInfo
            If bmpInfo.bmBits = 0 Then
                GetIconInfo .Image.Handle, icoInfo
                If icoInfo.hbmColor <> 0 Then
                    ' downside... API creates 2 bitmaps that we need to destroy since they aren't used in this
                    ' routine & are not destroyed automatically. To prevent memory leak, we destroy them here
                    GetGDIObject icoInfo.hbmColor, LenB(bmpInfo), bmpInfo
                    DeleteObject icoInfo.hbmColor
                    If icoInfo.hbmMask <> 0 Then DeleteObject icoInfo.hbmMask
                     .Type = CI_ICON        ' flag indicating image is an icon
                End If
            Else
                 .Type = CI_BITMAP     ' flag indicating image is a bitmap
            End If
             .SourceSize.X = bmpInfo.bmWidth
             .SourceSize.Y = bmpInfo.bmHeight
        End If
    End With
End Select
End Sub

Private Sub GetSetOffDC(ByVal bSet As Boolean)
' This sets up our off screen DC & pastes results onto our control.
Dim hBmp As Long
With HeaderDC
    If bSet Then
       If .hDC = 0 Then
          .hDC = CreateCompatibleDC(UserControl.hDC)
          SetBkMode .hDC, 3&
    ' by pulling these objects now, we ensure no memory leaks &
    ' changing the objects as needed can be done in 1 line of code
    ' in the SetButtonColors routine
          .OldBrush = SelectObject(.hDC, CreateSolidBrush(0&))
          .OldPen = SelectObject(.hDC, CreatePen(0&, 1&, 0&))
       End If
       If HeaderDC.OldBitmap = 0 Then
          hBmp = CreateCompatibleBitmap(UserControl.hDC, UserControl.ScaleWidth, UserControl.ScaleHeight)
          .OldBitmap = SelectObject(.hDC, hBmp)
       End If
     Else
       BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, .hDC, 0, 0, vbSrcCopy
    End If
End With
End Sub

Private Sub MakeRegion()
Dim X                       As Integer
Dim rgn1                    As Long
Dim rgnMain                 As Long
Dim r As Integer
Dim b As Integer

With myProps
     r = .hUCRect.Right
     b = .hUCRect.Bottom
     X = .hCnrSize
End With

DeleteObject rgnMain
rgnMain = CreateRectRgn(0, 0, r, b)

If myProps.hCnrShape = d_Square Then GoTo SetRgn

'Left top cnr
rgn1 = CreateRectRgn(0, 0, 1, X)
CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
DeleteObject rgn1
rgn1 = CreateRectRgn(0, 0, X, 1)
CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
DeleteObject rgn1

'Right top cnr
rgn1 = CreateRectRgn(r - 1, 0, r, X)
CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
DeleteObject rgn1
rgn1 = CreateRectRgn(r - X, 0, r, 1)
CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
DeleteObject rgn1

If myProps.hCnrSize = d_Large Then
  'Left top cnr
   rgn1 = CreateRectRgn(0, 0, 3, 2)
   CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
   DeleteObject rgn1
   rgn1 = CreateRectRgn(0, 0, 2, 3)
   CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
   DeleteObject rgn1
   'Right top cnr
   rgn1 = CreateRectRgn(r - 3, 0, r, 2)
   CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
   DeleteObject rgn1
   rgn1 = CreateRectRgn(r - 2, 0, r, 3)
   CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
   DeleteObject rgn1
End If

If myProps.hCnrShape = d_RoundedTop Then GoTo SetRgn

'Left bottom cnr
rgn1 = CreateRectRgn(0, b - 1, X, b)
CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
DeleteObject rgn1
rgn1 = CreateRectRgn(0, b - X, 1, b)
CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
DeleteObject rgn1
'Right bottom cnr
rgn1 = CreateRectRgn(r - 1, b - X, r, b)
CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
DeleteObject rgn1
rgn1 = CreateRectRgn(r - X, b - 1, r, b)
CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
DeleteObject rgn1

If myProps.hCnrSize = d_Large Then
   'Left bottom cnr
   rgn1 = CreateRectRgn(0, b - 2, 3, b)
   CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
   DeleteObject rgn1
   rgn1 = CreateRectRgn(0, b - 3, 2, b)
   CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
   DeleteObject rgn1
   'Right bottom cnr
   rgn1 = CreateRectRgn(r - 3, b - 2, r, b)
   CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
   DeleteObject rgn1
   rgn1 = CreateRectRgn(r - 2, b - 3, r, b)
   CombineRgn rgnMain, rgnMain, rgn1, RGN_DIFF
   DeleteObject rgn1
End If

SetRgn:
SetWindowRgn UserControl.hwnd, rgnMain, True
End Sub

Private Sub NormalizeImage(ByVal newSizeX As Long, _
                           ByVal newSizeY As Long)
Dim cTrans     As Long
Dim valGreen   As Long
Dim valRed     As Long
Dim valBlue    As Long
Dim tGreen     As Long
Dim tRed       As Long
Dim tBlue      As Long
Dim X          As Long
Dim Y          As Long
Dim cPixel     As Long
Dim oldBMP     As Long
Dim newDC      As Long
    
If myImage.Image Is Nothing Then
   Exit Sub
End If
If myImage.TransImage Then
   DeleteObject myImage.TransImage
End If
If myImage.Type Then
   Exit Sub
End If
' pain in the tush.
' In order to make a bitmap transparent, we need to decide which color will be the transparent color
' Well, the API CopyImage is used to resize images to fit the buttons. The downside is that this
' API has a habit of changing pixel colors very slightly. Even a single RGB value changed by a
' value of one can prevent the transparency routines from making the image transparent.
' This routine cleans up an image to help ensure it can be made transparent.
' Note: This routine is called each time a button is resized
' So even though this routine can be time consuming, it is called normally during IDE or initial form load.
' Last but not least, the routine also builds the non-rectangular regions for
' custom button shapes & returns the regions back to the CreateButtonRegion routine
' these are used only for creating custom button regions
' can't use ButtonDC.hDC -- need to create another DC 'cause if a clipping
' region is active (shaped/circular buttons), selecting image into DC may fail
    
newDC = CreateCompatibleDC(UserControl.hDC)
' get the image into a DC so we can clean it up
With myImage
     If .Type Then    ' icons
        .TransImage = CreateCompatibleBitmap(UserControl.hDC, newSizeX, newSizeY)
        oldBMP = SelectObject(newDC, .TransImage)
        DrawIconEx newDC, 0, 0, .Image.Handle, newSizeX, newSizeY, 0&, 0&, &H3
      Else    ' bitmaps
        .TransImage = CopyImage(.Image.Handle, .Type, newSizeX, newSizeY, ByVal 0&)
        oldBMP = SelectObject(newDC, .TransImage)
     End If
End With 'myImage
' determine the mask color (top left corner pixel)
cTrans = GetPixel(newDC, 0, 0)
' get the RGB values for that pixel
valRed = (cTrans \ (&H100 ^ 0) And &HFF)
valGreen = (cTrans \ (&H100 ^ 1) And &HFF)
valBlue = (cTrans \ (&H100 ^ 2) And &HFF)
' now loop thru each pixel & clean up any that were changed by the CopyImage API
For Y = 0 To newSizeY
    For X = 0 To newSizeX
        cPixel = GetPixel(newDC, X, Y)                      ' current pixel
        tRed = (cPixel \ (&H100 ^ 0) And &HFF)          ' RGB values for current pixel
        tGreen = (cPixel \ (&H100 ^ 1) And &HFF)
        tBlue = (cPixel \ (&H100 ^ 2) And &HFF)
' Test to see if the current pixel is real close to the transparent color used & change it if so
        If tRed >= valRed - 3 And tRed <= valRed + 3 And tBlue >= valBlue - 3 And tBlue <= valBlue + 3 And tGreen >= valGreen - 3 And tGreen <= valGreen + 3 Then
           SetPixel newDC, X, Y, cTrans
        End If
    Next X
Next Y
' Pull the image out of the DC & use it for all other image routines
SelectObject newDC, oldBMP
DeleteDC newDC
myImage.TransSize.X = newSizeX
myImage.TransSize.Y = newSizeY
End Sub

Private Sub RotatePoint(uAxisPt As PointSng, uRotatePt As PointSng, fDegrees As Single)
Dim fDX         As Single
Dim fDY         As Single
Dim fRadians    As Single

fRadians = fDegrees * RADS
fDX = uRotatePt.X - uAxisPt.X
fDY = uRotatePt.Y - uAxisPt.Y
uRotatePt.X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
uRotatePt.Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
End Sub

Private Sub ScaleImage(ByVal SizeX As Long, _
                       ByVal SizeY As Long, _
                       ImgX As Long, _
                       ImgY As Long)
' helper function for resizing images to scale
Dim Ratio(0 To 1) As Double
Ratio(0) = SizeX / myImage.SourceSize.X
Ratio(1) = SizeY / myImage.SourceSize.Y
If Ratio(1) < Ratio(0) Then
   Ratio(0) = Ratio(1)
End If
ImgX = myImage.SourceSize.X * Ratio(0)
ImgY = myImage.SourceSize.Y * Ratio(0)
Erase Ratio
End Sub

Private Sub SetHeaderColors(ByVal bSet As Boolean, _
                            ByVal m_hDC As Long, _
                            TypeObject As ColorObjects, _
                            ByVal lColor As Long, _
                            Optional bSamePenColor As Boolean = True, _
                            Optional PenWidth As Long = 1, _
                            Optional PenStyle As Long = 0)
' This is the basic routine that sets a DC's pen, brush or font color
' here we store the most recent "sets" so we can reset when needed

If bSet Then    ' changing a DC's setting
   Select Case TypeObject
     Case cObj_Brush         ' brush is being changed
          DeleteObject SelectObject(HeaderDC.hDC, CreateSolidBrush(lColor))
          If bSamePenColor Then   ' if the pen color will be the same
             DeleteObject SelectObject(HeaderDC.hDC, CreatePen(PenStyle, PenWidth, lColor))
          End If
     Case cObj_Pen   ' pen is being changed (mostly for drawing lines)
          DeleteObject SelectObject(HeaderDC.hDC, CreatePen(PenStyle, PenWidth, lColor))
     Case cObj_Text  ' text color is changing
          SetTextColor m_hDC, ConvertColor(lColor)
   End Select
 Else            ' resetting the DC back to the way it was
   With HeaderDC
       DeleteObject SelectObject(.hDC, .OldBrush)
       DeleteObject SelectObject(.hDC, .OldPen)
   End With
End If
End Sub

Private Function ShadeColor(ByVal lColor As Long, _
                            shadeOffset As Integer, _
                            ByVal lessBlue As Boolean, _
                            Optional ByVal bFocusRect As Boolean, _
                            Optional ByVal bInvert As Boolean) As Long
' Basically supply a value between -255 and +255. Positive numbers make
' the passed color lighter and negative numbers make the color darker
Dim valRGB(0 To 2) As Integer
Dim I              As Integer
 
CalcNewColor:
valRGB(0) = (lColor And &HFF) + shadeOffset
valRGB(1) = ((lColor And &HFF00&) / 255&) + shadeOffset

If lessBlue Then
   valRGB(2) = (lColor And &HFF0000) / &HFF00&
   valRGB(2) = valRGB(2) + ((valRGB(2) * CLng(shadeOffset)) \ &HC0)
 Else
   valRGB(2) = (lColor And &HFF0000) / &HFF00& + shadeOffset
End If
For I = 0 To 2
    If valRGB(I) > 255 Then
       valRGB(I) = 255
    End If
    If valRGB(I) < 0 Then
       valRGB(I) = 0
    End If
    If bInvert Then
       valRGB(I) = Abs(255 - valRGB(I))
    End If
Next I
ShadeColor = valRGB(0) + 256& * valRGB(1) + 65536 * valRGB(2)
Erase valRGB
If bFocusRect And (ShadeColor = vbBlack Or ShadeColor = vbWhite) Then
   shadeOffset = -shadeOffset
   If shadeOffset = 0 Then
      shadeOffset = 64
   End If
   GoTo CalcNewColor
End If
End Function




'##########################################################
'=========================================================
'/////////////////////Properties\\\\\\\\\\\\\\\\\\\\\\\\\\\
'=========================================================
'##########################################################


Public Property Get BorderColor() As OLE_COLOR
BorderColor = myProps.hBorderClr
End Property
Public Property Let BorderColor(nColor As OLE_COLOR)
'Finishing Gradient Color
myProps.hBorderClr = nColor
DrawInfoHeader
PropertyChanged "BorderColor"
End Property

Public Property Let BorderShape(nStyle As CrnrShpeCnsts)
myProps.hCnrShape = nStyle
PropertyChanged "BorderShape"
DrawInfoHeader
End Property
Public Property Get BorderShape() As CrnrShpeCnsts
BorderShape = myProps.hCnrShape
End Property

Public Property Get BorderVisible() As Boolean
BorderVisible = myProps.hBorderVis
End Property
Public Property Let BorderVisible(b As Boolean)
'Finishing Gradient Color
myProps.hBorderVis = b
DrawInfoHeader
PropertyChanged "BVisible"
End Property

Public Property Get Caption() As String
Caption = myProps.hCaption
End Property
Public Property Let Caption(ByVal newValue As String)
myProps.hCaption = newValue                     ' cache the caption = newValue
CalculateBoundingRects False
DrawInfoHeader
PropertyChanged "Caption"
End Property

Public Property Let CaptionAlign(nAlign As AlignmentConstants)
' Caption options: Left, Right or Center Justified
If nAlign < vbLeftJustify Or nAlign > vbCenter Then Exit Property
If myImage.Align > d_RightOfCaption And nAlign < vbCenter And (myImage.SourceSize.X + myImage.SourceSize.Y) > 0 Then
    ' also prevent left/right justifying captions when image is centered in caption
    If UserControl.Ambient.UserMode = False Then
        ' if not in user mode, then explain whey it is prevented
        MsgBox "When header images are aligned top/bottom center, " & vbCrLf & "header captions can only be center aligned", vbOKOnly + vbInformation
    End If
    Exit Property
End If
myProps.hCaptionAlign = nAlign
CalculateBoundingRects False              ' recalculate text/image bounding rects
DrawInfoHeader
PropertyChanged "CapAlign"
End Property
Public Property Get CaptionAlign() As AlignmentConstants
CaptionAlign = myProps.hCaptionAlign
End Property

Public Property Let CaptionStyle(nStyle As CaptionEffectCnsts)

' Sets the style, raised/sunken or flat (default)

If nStyle < d_Default Or nStyle > d_Raised Then Exit Property
myProps.hCaptionStyle = nStyle
PropertyChanged "CapStyle"
If Len(myProps.hCaption) Then
    CalculateBoundingRects False
    DrawInfoHeader
End If
End Property
Public Property Get CaptionStyle() As CaptionEffectCnsts
CaptionStyle = myProps.hCaptionStyle
End Property

Public Property Let CornerSize(nStyle As CrnrSzeCnsts)
myProps.hCnrSize = nStyle
PropertyChanged "CornerSize"
DrawInfoHeader
End Property
Public Property Get CornerSize() As CrnrSzeCnsts
CornerSize = myProps.hCnrSize
End Property

Public Property Get Font() As StdFont
Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal newValue As StdFont)
Set UserControl.Font = newValue
GetGDIMetrics "Font"
CalculateBoundingRects False          ' recalculate caption's text/image bounding rects
DrawInfoHeader
PropertyChanged "Font"
End Property

Public Property Get ForeColorShadow() As OLE_COLOR
ForeColorShadow = myProps.hCapShadow
End Property
Public Property Let ForeColorShadow(ByVal newValue As OLE_COLOR)
myProps.hCapShadow = newValue
DrawInfoHeader
PropertyChanged "ForeColorShadow"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal newValue As OLE_COLOR)
UserControl.ForeColor = newValue
DrawInfoHeader
PropertyChanged "ForeColor"
End Property

Public Property Get Angle() As Integer
Angle = myProps.hAngle
End Property
Public Property Let Angle(a As Integer)
myProps.hAngle = a
DrawInfoHeader
PropertyChanged "Angle"
End Property

Public Property Get GradientAngle() As Integer
GradientAngle = myProps.hAngle
End Property
Public Property Let GradientAngle(a As Integer)
'Gradient Direction
myProps.hAngle = a
DrawInfoHeader
PropertyChanged "GradientAngle"
End Property

Public Property Get GradientFinish() As OLE_COLOR
GradientFinish = myProps.hEgc
End Property
Public Property Let GradientFinish(nColor As OLE_COLOR)
'Finishing Gradient Color
myProps.hEgc = nColor
DrawInfoHeader
PropertyChanged "GradientFinish"
End Property

Public Property Get GradientFinishStyle() As BackStyleCnsts
GradientFinishStyle = myProps.hGradientBackStyle
End Property
Public Property Let GradientFinishStyle(Styles As BackStyleCnsts)
'Sets whether or not the finish color is d_Opaque
myProps.hGradientBackStyle = Styles
DrawInfoHeader
PropertyChanged "GradientFinishStyle"
End Property

Public Property Get GradientStart() As OLE_COLOR
GradientStart = myProps.hSgc
End Property
Public Property Let GradientStart(nColor As OLE_COLOR)
'Starting Gradient Color
myProps.hSgc = nColor
DrawInfoHeader
PropertyChanged "GradientStart"
End Property

Public Property Get Picture() As StdPicture
Set Picture = myImage.Image
End Property
Public Property Set Picture(xPic As StdPicture)
' Sets the button image which to display
With myImage
     Set .Image = xPic
     If .Size = 0 Then
        .Size = 16
     End If
End With 'myImage
    
GetGDIMetrics "Picture"
CalculateBoundingRects True              ' recalculate button's text/image bounding rects
DrawInfoHeader
PropertyChanged "Image"
End Property

Public Property Get PictureAlign() As ImagePlacementConstants
PictureAlign = myImage.Align
End Property
Public Property Let PictureAlign(ImgAlign As ImagePlacementConstants)
' Image alignment options for button (6 different positions)
If ImgAlign < d_LeftEdge Or ImgAlign > d_BottomCenter Then
   Exit Property
End If
myImage.Align = ImgAlign
'If ImgAlign = d_BottomCenter Or ImgAlign = d_TopCenter Then CaptionAlign = vbCenter
CalculateBoundingRects False             ' recalculate button's text/image bounding rects
DrawInfoHeader
PropertyChanged "ImgAlign"
End Property

Public Property Get PictureSize() As ImageSizeConstants
If myImage.Size = 0 Then
   myImage.Size = 16
End If
' parameters are 0,1,2,3,4 & 5, but we store them as 16,24,32,40, & 44
PictureSize = Choose(myImage.Size / 8 - 1, d_16x16, d_24x24, d_32x32, d_Fill_Stretch, d_Fill_ScaleUpDown)
End Property
Public Property Let PictureSize(nSize As ImageSizeConstants)
' Sets up to 5 picture sizes
If PictureSize < d_16x16 Or PictureSize > d_Fill_ScaleUpDown Then
   Exit Property
End If
myImage.Size = (nSize + 2) * 8      ' I just want the size as pixel x pixel
CalculateBoundingRects True         ' recalculate text/image bounding rects
DrawInfoHeader
PropertyChanged "ImgSize"
End Property

Private Sub UserControl_InitProperties()
On Error Resume Next
With myProps
    .hBorderClr = vbBlue
    .hBorderVis = True
    .hCaption = Ambient.DisplayName
    .hCapShadow = &H8000000F
    .hCnrSize = d_Large
    .hCnrShape = d_Rounded
    Set UserControl.Font = Parent.Font
    UserControl.ForeColor = vbWhite
    .hAngle = 270
    .hSgc = &HF3855A
    .hEgc = vbWhite
    .hGradientBackStyle = d_Opaque
End With
End Sub

Private Sub Usercontrol_Initialize()
''''''''''
End Sub

Private Sub UserControl_Paint()
DrawInfoHeader
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
     myProps.hBorderClr = .ReadProperty("BorderColor", vbBlue)
     myProps.hBorderVis = .ReadProperty("BorderVisible", True)
     myProps.hCnrShape = .ReadProperty("BorderShape", d_Rounded)
     myProps.hCaption = .ReadProperty("Caption", Ambient.DisplayName)
     myProps.hCaptionStyle = .ReadProperty("CapStyle", 0)
     myProps.hCaptionAlign = .ReadProperty("CapAlign", 0)
     myProps.hCapShadow = .ReadProperty("ForeColorShadow", &H8000000F)
     myProps.hCnrSize = .ReadProperty("CornerSize", d_Large)
     Set UserControl.Font = .ReadProperty("Font", Font)
     UserControl.ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
     myProps.hAngle = .ReadProperty("GradientAngle", 270)
     myProps.hSgc = .ReadProperty("GradientStart", vbBlue)
     myProps.hEgc = .ReadProperty("GradientFinish", vbWhite)
     myProps.hGradientBackStyle = .ReadProperty("GradientFinishStyle", 0)
     Set myImage.Image = .ReadProperty("Image", Nothing)
     myImage.Size = .ReadProperty("ImgSize", 16)
     myImage.Align = .ReadProperty("ImgAlign", 0)
End With
End Sub

Private Sub UserControl_Resize()
' since we are using a separate DC for drawing, we need to resize the
' bitmap in that DC each time the control resizes
With HeaderDC
     If .hDC Then
         DeleteObject SelectObject(.hDC, .OldBitmap)
        .OldBitmap = 0  ' this will force a new bitmap for existing DC
     End If
End With 'HeaderDC

GetSetOffDC True
GetGDIMetrics "Font"
GetGDIMetrics "UC"
GetGDIMetrics "Picture"
CalculateBoundingRects False
DrawInfoHeader True
End Sub

Private Sub UserControl_Terminate()
' Header is ending, let's clean up
With HeaderDC
     If .hDC Then
' get rid of left over pen & brush
        SetHeaderColors False, .hDC, cObj_Pen, 0
' get rid of logical font
        DeleteObject SelectObject(.hDC, .OldFont)
' destroy the separate Bitmap & select original back into DC
        DeleteObject SelectObject(.hDC, .OldBitmap)
' destroy the temporary DC
        DeleteDC .hDC
     End If
End With 'HeaderDC
' kill image used for transparencies when selected button pic is a bitmap
With myImage
    If .TransImage Then DeleteObject .TransImage
    On Error Resume Next
    If .Image Then DeleteObject .Image
    On Error GoTo 0
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
     .WriteProperty "BorderColor", myProps.hBorderClr, vbBlue
     .WriteProperty "BorderVisible", myProps.hBorderVis, True
     .WriteProperty "BorderShape", myProps.hCnrShape, d_Rounded
     .WriteProperty "Caption", myProps.hCaption, Ambient.DisplayName
     .WriteProperty "CapAlign", myProps.hCaptionAlign, 0
     .WriteProperty "CapStyle", myProps.hCaptionStyle, 0
     .WriteProperty "ForeColorShadow", myProps.hCapShadow, &H8000000F
     .WriteProperty "CornerSize", myProps.hCnrSize, d_Large
     .WriteProperty "Font", UserControl.Font, Nothing
     .WriteProperty "ForeColor", UserControl.ForeColor
     .WriteProperty "GradientAngle", myProps.hAngle, 270
     .WriteProperty "GradientStart", myProps.hSgc, vbBlue
     .WriteProperty "GradientFinish", myProps.hEgc, vbWhite
     .WriteProperty "GradientFinishStyle", myProps.hGradientBackStyle, 0
     .WriteProperty "ImgAlign", myImage.Align, 0
     .WriteProperty "Image", myImage.Image, Nothing
     .WriteProperty "ImgSize", myImage.Size, 16
End With
End Sub
