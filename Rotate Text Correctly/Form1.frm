VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotate Text Correctly"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PI = 3.14159265358979

Private Sub Form_Load()
    
    Dim Angle As Double
    
    'i translated each language name to it's language, for example: i translated the word
    'Greek' from english to greek, and got its chars codes by microsoft word
    
    Dim English As String: English = "English"
    Dim Arabic As String: Arabic = ChrW(&H627) & ChrW(&H644) & ChrW(&H639) & ChrW(&H631) & ChrW(&H628) & ChrW(&H64A) & ChrW(&H629)
    Dim Hebrew As String: Hebrew = ChrW(&H5E2) & ChrW(&H5D1) & ChrW(&H5E8) & ChrW(&H5D9) & ChrW(&H5EA)
    Dim Greek As String: Greek = ChrW(&H395) & ChrW(&H3BB) & ChrW(&H3BB) & ChrW(&H3B7) & ChrW(&H3BD) & ChrW(&H3B9) & ChrW(&H3BA) & ChrW(&H3AC)
    Dim Hindi As String: Hindi = ChrW(&H939) & ChrW(&H93F) & ChrW(&H928) & ChrW(&H94D) & ChrW(&H926) & ChrW(&H940)
    Dim Japanese As String: Japanese = ChrW(&H65E5) & ChrW(&H672C) & ChrW(&H8A9E)
    Dim Chinese As String: Chinese = ChrW(&H4E2D) & ChrW(&H6587)
    Dim Russian As String: Russian = ChrW(&H420) & ChrW(&H443) & ChrW(&H441) & ChrW(&H441) & ChrW(&H43A) & ChrW(&H438) & ChrW(&H439)
    Dim Thai As String: Thai = ChrW(&HE20) & ChrW(&HE32) & ChrW(&HE29) & ChrW(&HE32) & ChrW(&HE44) & ChrW(&HE17) & ChrW(&HE22)
    
    Me.Show
    Dim a As Double
    On Error Resume Next
    
    Do
        a = a + 1
        Angle = Angle + Sine(a * PI / 180) * PI
        DoEvents       '(Sin(a * PI / 180) * PI
        Cls
        
        RotateText Me.hDC, English, Me.Font, vbRed, vbLeftJustify, False, 50, 50, Angle
        RotateText Me.hDC, Arabic, Me.Font, RGB(0, 192, 0), vbCenter, True, 50, 200, Angle
        RotateText Me.hDC, Hebrew, Me.Font, RGB(0, 192, 192), vbCenter, True, 50, 350, Angle
        RotateText Me.hDC, Greek, Me.Font, vbBlack, vbLeftJustify, False, 200, 50, Angle
        RotateText Me.hDC, Hindi, Me.Font, vbMagenta, vbCenter, False, 200, 200, Angle
        RotateText Me.hDC, Japanese, Me.Font, vbBlue, vbCenter, False, 200, 350, Angle
        RotateText Me.hDC, Chinese, Me.Font, RGB(160, 160, 0), vbCenter, False, 400, 50, Angle
        RotateText Me.hDC, Russian, Me.Font, RGB(0, 0, 128), vbLeftJustify, False, 350, 200, Angle
        RotateText Me.hDC, Thai, Me.Font, RGB(128, 0, 0), vbLeftJustify, False, 350, 350, Angle
        
        Refresh
        If Err <> 0 Then
            Angle = 0
            a = 0
            Err.Clear
        End If
        
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Function Sine(ByVal x As Double) As Double
    Sine = _
    x - x ^ 3 / 6 + x ^ 5 / 120 - x ^ 7 / 5040 + x ^ 9 / 362880 - x ^ 11 / 39916800 + x ^ 13 / 6227020800# - x ^ 15 / 1307674368000# + x ^ 17 / 355687428096000# - x ^ 19 / 1.21645100408832E+17 + x ^ 21 / 5.10909421717094E+19 - x ^ 23 / 2.5852016738885E+22 + x ^ 25 / 1.5511210043331E+25 - x ^ 27 / 1.08888694504184E+28 + x ^ 29 / 8.8417619937397E+30 - x ^ 31 / 8.22283865417792E+33 + x ^ 33 / 8.68331761881189E+36 - x ^ 35 / 1.03331479663861E+40 + x ^ 37 / 1.37637530912263E+43 - x ^ 39 / 2.03978820811974E+46 + x ^ 41 / 3.34525266131638E+49 - x ^ 43 / 6.04152630633738E+52 + x ^ 45 / 1.1962222086548E+56 - x ^ 47 / 2.58623241511168E+59 + x ^ 49 / 6.08281864034268E+62 - x ^ 51 / 1.55111875328738E+66 + x ^ 53 / 4.27488328406003E+69 - x ^ 55 / 1.26964033536583E+73 + x ^ 57 / 4.05269195048772E+76 - x ^ 59 / 1.3868311854569E+80 + x ^ 61 / 5.07580213877225E+83 - x ^ 63 / 1.98260831540444E+87 + x ^ 65 / 8.24765059208247E+90 - x ^ 67 / 3.64711109181887E+94 + x ^ 69 / 1.71122452428141E+98 - x ^ 71 / 8.50478588567862E+101
End Function
