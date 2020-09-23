Attribute VB_Name = "Module2"
Option Explicit

'TextOut TextAligns
Private Const TA_BASELINE = 24
Private Const TA_BOTTOM = 8
Private Const TA_CENTER = 6
Private Const TA_LEFT = 0
Private Const TA_NOUPDATECP = 0
Private Const TA_RIGHT = 2
Private Const TA_TOP = 0
Private Const TA_UPDATECP = 1
Private Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
Private Const TA_RTLREADING = &H100

Private Const PI = 3.14159265358979

Private Const PROOF_QUALITY = 2
Private Const LOGPIXELSY = 90

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type Size
        cx As Long
        cy As Long
End Type

Private Type POINTF
   x As Single
   y As Single
End Type

Private Declare Function TextOut Lib "gdi32" Alias "TextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal Height As Long, ByVal Width As Long, ByVal Escapement As Long, ByVal Orientation As Long, ByVal Weight As Long, ByVal Italic As Long, ByVal Underline As Long, ByVal StrikeOut As Long, ByVal Charset As Long, ByVal OutPrecision As Long, ByVal ClipPrecision As Long, ByVal Quality As Long, ByVal PitchAndFamily As Long, ByVal FaceName As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'rotates a point around a center point
Private Function RotatePoint(ByRef pt As POINTF, ByRef CenterPoint As POINTF, ByVal Angle As Single) As POINTF
    RotatePoint.x = (pt.x - CenterPoint.x) * Cos(Angle * PI / 180) - (pt.y - CenterPoint.y) * Sin(Angle * PI / 180) + CenterPoint.x
    RotatePoint.y = (pt.x - CenterPoint.x) * Sin(Angle * PI / 180) + (pt.y - CenterPoint.y) * Cos(Angle * PI / 180) + CenterPoint.y
End Function

Private Function NewPOINTAPI(ByVal x As Long, ByVal y As Long) As POINTAPI
    NewPOINTAPI.x = x
    NewPOINTAPI.y = y
End Function

Private Function NewPOINTF(ByVal x As Single, ByVal y As Single) As POINTF
    NewPOINTF.x = x
    NewPOINTF.y = y
End Function

Private Function POINTFToPOINTAPI(ByRef pt As POINTF) As POINTAPI
    POINTFToPOINTAPI.x = pt.x
    POINTFToPOINTAPI.y = pt.y
End Function

Private Function POINTAPIToPOINTF(ByRef pt As POINTAPI) As POINTF
    POINTAPIToPOINTF.x = pt.x
    POINTAPIToPOINTF.y = pt.y
End Function

'insteed of Object.TextWidth and Object.TextHeight, to be compatible for all programming languages
Private Function GetTextSize(ByVal str As String, ByVal hFont As Long) As Size

    Dim dc As New DeviceContext
    ' a memory DC for help getting text width and height, bcause GetTextExtentPoint32 needs hDC
    Dim s As Size 'return value for GetTextExtentPoint32
    Dim OldFont As Long
    
    dc.Create 24, 0, 0 'no matter what is DC's size
    
    OldFont = SelectObject(dc.Handle, hFont)
    GetTextExtentPoint32 dc.Handle, str, Len(str), s
    SelectObject dc.Handle, OldFont
    DeleteObject OldFont
    GetTextSize = s
    dc.Dispose
End Function
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            ' _
Draw Rotated Text                                                                                                                                                                                                                                                                                                                                                                                                                   _ _
CreateFont API function draws rotated text but not correctly(it rotates the text around the starting point of                                                                                                                                                                                                                                                                                                                                                                                                                                              _ _
writing),this function draws rotated text around center point of the specefied text.                                                                                                                                                                                                                                                                                                                                                                                                                                    _ _
this function uses CreateFont but corrects the starting point to rotate
Public Sub RotateText(ByVal hDC As Long, ByVal str As String, ByVal Font As StdFont, ByVal FontColor As Long, _
                      ByVal TextAlign As AlignmentConstants, ByVal RightToLeft As Boolean, ByVal x As Long, ByVal y As Long, ByVal Angle As Single)
                      
    Dim pt(1 To 2) As POINTF 'pt(1) is for text x and y, pt(2) is for text size
    Dim dc As New DeviceContext 'for converting font size from pixel to point using GetDeviceCaps API function
    Dim sz As Size 'return value for GetTextSize
    Dim hFont As Long ' font hangle for CreateFont API function
    Dim FontSize As Long
    
    
    dc.Create 24, 1, 1
    FontSize = -MulDiv(Font.Size, GetDeviceCaps(dc.Handle, LOGPIXELSY), 72) 'pixels to points
    
    hFont = CreateFont(FontSize, 0, Angle * 10, 0, Font.Weight, Font.Italic, Font.Underline, Font.Strikethrough, _
    Font.Charset, 0, 0, 6, 0, Font.Name)

    sz = GetTextSize(str, hFont)
    pt(1).x = x ' original x to correct
    pt(1).y = y ' original y to correct
    pt(2).x = sz.cx + pt(1).x 'end point of text rectangle
    pt(2).y = sz.cy + pt(1).y
        
    If TextAlign = vbLeftJustify Then
        'if text align is left, the center point to rotate is starting x and starting y
        'we have to rotate the rectangle of the text ( pt(1) and pt(2) ) around the center point of the text
        'and change the start point of drawing the text to the rotated starting x and y, with the munis of the
        'specefied angle
        pt(1) = RotatePoint(pt(1), _
                            NewPOINTF((pt(1).x + pt(2).x) / 2, (pt(1).y + pt(2).y) / 2), -Angle)
    ElseIf TextAlign = vbRightJustify Then
        'if text align is right, the x of the point to rotate is the x of the end point, and the y is the y
        'of the start point, with minus angle
        pt(1) = RotatePoint(NewPOINTF(pt(2).x, pt(1).y), _
                            NewPOINTF((pt(1).x + pt(2).x) / 2, (pt(1).y + pt(2).y) / 2), -Angle)
    ElseIf TextAlign = vbCenter Then
        'if text text align is center, the x of the rotated point is the centered x of the text rectangle,
        'the y of the rotated point is the start point y, with minus angle
        pt(1) = RotatePoint(NewPOINTF((pt(1).x + pt(2).x) / 2, _
                            pt(1).y), NewPOINTF((pt(1).x + pt(2).x) / 2, (pt(1).y + pt(2).y) / 2), -Angle)
    End If
    
    Dim OldFont As Long
    OldFont = SelectObject(hDC, hFont)
    
    
    PrintText hDC, str, FontColor, TextAlign, RightToLeft, pt(1).x, pt(1).y
    SelectObject hDC, OldFont
    DeleteObject hFont
End Sub

Private Function NewRECT(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As RECT
    NewRECT.Left = Left
    NewRECT.Top = Top
    NewRECT.Right = Right
    NewRECT.Bottom = Bottom
End Function
                                                                                                                                                                                                                                                                                ' _
This function draws text on DC, use this instead of Object.Print method, its supports unicode.                                                                                                                                                                                                                                                                                _ _
This function draws text Using TextOut, not DrawText, because DrawText draws the text in a                                                                                                                                                                                                                                                                                 _ _
specefied rectangle, but TextOut draws the text in the whole DC specifying the point to start.
Private Sub PrintText(ByVal hDC As Long, ByVal str As String, ByVal FontColor As Long, ByVal TextAlignment As AlignmentConstants, ByVal RightToLeft As Boolean, ByVal x As Long, ByVal y As Long)
    Dim Flags As Long
    Dim hFont As Long
    Dim OldFont As Long
    Dim SaveTextAlignment As Long
    Dim SaveTextColor As Long
    
    SaveTextAlignment = GetTextAlign(hDC)
    SaveTextColor = GetTextColor(hDC)
    
    If TextAlignment = vbLeftJustify Then
        Flags = Flags Or TA_LEFT
    ElseIf TextAlignment = vbRightJustify Then
        Flags = Flags Or TA_RIGHT
    ElseIf TextAlignment = vbCenter Then
        Flags = Flags Or TA_CENTER
    End If
    
    If RightToLeft = True Then Flags = Flags Or TA_RTLREADING
    
    SetTextAlign hDC, Flags
    SetTextColor hDC, FontColor
    
    TextOut hDC, x, y, StrPtr(str), Len(str)
        
    SetTextAlign hDC, SaveTextAlignment
    SetTextColor hDC, SaveTextColor
    
End Sub
