VERSION 5.00
Begin VB.UserControl xVistaForm 
   Alignable       =   -1  'True
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   ScaleHeight     =   570
   ScaleWidth      =   2310
End
Attribute VB_Name = "xVistaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    '****************************************************************
    ' Project:      Creates an Ownerdrawn Vista Style Form control
    ' Programmer:   Alexander Mungall
    ' UserControl:  xVistaForm
    ' Email:        goober_mpc@hotmail.com
    '----------------------------------------------------------------
    ' xVistaForm Copyright© Alexander Mungall, All Rights Reserved
    ' Feel free to use this code for personal use in anyway you see
    ' fit, but please give credit where credit is due...
    ' It's all I ask.
    '****************************************************************
    Option Explicit
    
    ' Booleans
    Private bCloseButton As Boolean
    Private bCloseButtonClicked As Boolean
    Private bCloseButtonHover As Boolean
    Private bDisplayIcon As Boolean
    Private bMaximiseButton As Boolean
    Private bMaximiseButtonClicked As Boolean
    Private bMaximiseButtonHover As Boolean
    Private bMinimiseButton As Boolean
    Private bMinimiseButtonClicked As Boolean
    Private bMinimiseButtonHover As Boolean
    Private bMouseOnForm As Boolean
    Private bPaintForm As Boolean
    Private bTransparency As Boolean
    Private bUnloadForm As Boolean
    
    ' Controls
    Private imgFormPic As Image
    Private myForm As Form
    Private WithEvents lblFormCaption As Label
Attribute lblFormCaption.VB_VarHelpID = -1
    Private WithEvents moveForm As Form
Attribute moveForm.VB_VarHelpID = -1
    Private WithEvents picForm As PictureBox
Attribute picForm.VB_VarHelpID = -1
    Private WithEvents tmrMouseMove As Timer
Attribute tmrMouseMove.VB_VarHelpID = -1

    ' Enums
    Public Enum xVistaStyles
        VistaLite = 0
        VistaDark = 1
    End Enum
    Private xVisualStyles As xVistaStyles

    ' Integers
    Private i As Integer
    Private iHorizontal As Integer
    Private iNumControls As Integer
    Private iTransparency As Integer
    Private iVertical As Integer
    
    ' Longs
    Private Col As Long
    Private lBottomR As Long
    Private lBottomG As Long
    Private lBottomB As Long
    Private lButtonGradientBottom(7) As Long
    Private lButtonGradientBottomClicked(7) As Long
    Private lButtonGradientBottomHover(7) As Long
    Private lButtonGradientTop As Long
    Private lButtonGradientTopClicked As Long
    Private lButtonGradientTopHover As Long
    Private lButtonInnerBorder As Long
    Private lButtonOuterBorder As Long
    Private lCloseButtonGradientBottom(7) As Long
    Private lCloseButtonGradientBottomClicked(7) As Long
    Private lCloseButtonGradientBottomHover(7) As Long
    Private lCloseButtonGradientTop As Long
    Private lCloseButtonGradientTopClicked As Long
    Private lCloseButtonGradientTopHover As Long
    Private lCloseButtonInnerBorder As Long
    Private lCloseButtonOuterBorder As Long
    Private lTopR As Long
    Private lTopG As Long
    Private lTopB As Long
    Private lFormCaptionColor As Long
    Private lFormGradientBottom As Long
    Private lFormGradientTop As Long
    Private lFormMiddleBorder As Long
    Private lFormOuterBorder As Long
    Private lngReturnValue As Long

    ' Move a Titleless Window
    Private Declare Function ReleaseCapture Lib "user32" () As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Const HTCAPTION = 2
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const WM_SYSCOMMAND = &H112
    Private Const SC_MOVE = &HF010
    Private Const WM_POPUPSYSTEMMENU = &H313
    
    ' Make a Semi Transparent Form
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Const G = (-20)
    Private Const LWA_COLORKEY = &H1
    Private Const LWA_ALPHA = &H2
    Private Const ULW_COLORKEY = &H1
    Private Const ULW_ALPHA = &H2
    Private Const ULW_OPAQUE = &H4
    Private Const WS_EX_LAYERED = &H80000
    
    ' Show a Form in the Taskbar
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Const GWL_EXSTYLE = (-20)
    Private Const WS_EX_APPWINDOW = &H40000
    Private Const SW_HIDE = 0
    Private Const SW_NORMAL = 1

    ' Types
    Private Type POINTAPI
        X As Long
        Y As Long
    End Type
    
    ' Functions
    Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

    ' Local constants and variables declarations
    Const BorderPixels = 5
    Private TwipX As Single, TwipY As Single, BorderWidth As Single, BorderHeight As Single

    '****************************************************************
    ' Gradient Code: Written by Mark Gordon (msg555)
    '----------------------------------------------------------------
    ' Copyright© Mark Gordon, All Rights Reserved
    '----------------------------------------------------------------
    Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
    Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    
    Private Const DIB_RGB_COLORS = 0&
    Private Const BI_RGB = 0&
    
    Private Type BITMAPINFOHEADER '40 bytes
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
    
    Private Type RGBQUAD
       rgbBlue As Byte
       rgbGreen As Byte
       rgbRed As Byte
       rgbReserved As Byte
    End Type
    
    Private Type BITMAPINFO
      bmiHeader As BITMAPINFOHEADER
      bmiColors As RGBQUAD
    End Type
    
    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, iPic As StdPicture) As Long
    Private Type PictDesc
        cbSizeofStruct As Long
        picType As Long
        hImage As Long
        xExt As Long
        yExt As Long
    End Type
    Private Type Guid
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
    End Type
    
    Private Enum Blends
        RGBBlend = 0
        HSLBlend = 1
    End Enum

Private Function CreateGradient(Width As Long, Height As Long, LeftToRight As Boolean, LeftTopColor As Long, RightBottomColor As Long, BlendType As Blends) As StdPicture
    Dim hBmp As Long, Bits() As Byte
    Dim RS As Byte, GS As Byte, BS As Byte 'Start RGB
    Dim RE As Byte, GE As Byte, BE As Byte 'End RGB
    Dim HS As Single, SS As Single, LS As Single 'Start HSL
    Dim HE As Single, SE As Single, LE As Single 'End HSL
    Dim Rc As Byte, GC As Byte, BC As Byte 'Current iteration RGB
    Dim X As Long, Y As Long
    ReDim Bits(0 To 3, 0 To Width - 1, 0 To Height - 1)
    
    RgbCol LeftTopColor, RS, GS, BS
    RgbCol RightBottomColor, RE, GE, BE
    
    If BlendType = RGBBlend Then
        If LeftToRight Then
            For X = 0 To Width - 1
                Rc = (1& * RS - RE) * ((Width - 1 - X) / (Width - 1)) + RE
                GC = (1& * GS - GE) * ((Width - 1 - X) / (Width - 1)) + GE
                BC = (1& * BS - BE) * ((Width - 1 - X) / (Width - 1)) + BE
                For Y = 0 To Height - 1
                    Bits(2, X, Y) = Rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        Else
            For Y = 0 To Height - 1
                Rc = (1& * RS - RE) * ((Height - 1 - Y) / (Height - 1)) + RE
                GC = (1& * GS - GE) * ((Height - 1 - Y) / (Height - 1)) + GE
                BC = (1& * BS - BE) * ((Height - 1 - Y) / (Height - 1)) + BE
                For X = 0 To Width - 1
                    Bits(2, X, Y) = Rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        End If
    ElseIf BlendType = HSLBlend Then
        RGBToHSL RS, GS, BS, HS, SS, LS
        RGBToHSL RE, GE, BE, HE, SE, LE
        If LeftToRight Then
            For X = 0 To Width - 1
                HSLToRGB (1& * HS - HE) * ((Width - 1 - X) / (Width - 1)) + HE, _
                        (1& * SS - SE) * ((Width - 1 - X) / (Width - 1)) + SE, _
                        (1& * LS - LE) * ((Width - 1 - X) / (Width - 1)) + LE, _
                        Rc, GC, BC
                For Y = 0 To Height - 1
                    Bits(2, X, Y) = Rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        Else
            For Y = 0 To Height - 1
                HSLToRGB (1& * HS - HE) * ((Height - 1 - Y) / (Height - 1)) + HE, _
                        (1& * SS - SE) * ((Height - 1 - Y) / (Height - 1)) + SE, _
                        (1& * LS - LE) * ((Height - 1 - Y) / (Height - 1)) + LE, _
                        Rc, GC, BC
                For X = 0 To Width - 1
                    Bits(2, X, Y) = Rc
                    Bits(1, X, Y) = GC
                    Bits(0, X, Y) = BC
                Next
            Next
        End If
    End If

    Dim BI As BITMAPINFO
    With BI.bmiHeader
        .biSize = Len(BI.bmiHeader)
        .biWidth = Width
        .biHeight = -Height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = ((((.biWidth * .biBitCount) + 31) \ 32) * 4) * Abs(.biHeight)
    End With
    hBmp = CreateBitmap(Width, Height, 1&, 32&, ByVal 0)
    SetDIBits 0&, hBmp, 0, Abs(BI.bmiHeader.biHeight), Bits(0, 0, 0), BI, DIB_RGB_COLORS

    Dim IGuid As Guid, PicDst As PictDesc
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    With PicDst
        .cbSizeofStruct = Len(PicDst)
        .hImage = hBmp
        .picType = vbPicTypeBitmap
    End With
    OleCreatePictureIndirect PicDst, IGuid, True, CreateGradient
End Function

'Helper Functions
Private Sub RgbCol(Col As Long, ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
    R = Col And &HFF&
    G = (Col And &HFF00&) \ &H100&
    B = (Col And &HFF0000) \ &H10000
End Sub

Private Sub RGBToHSL(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, H As Single, S As Single, L As Single)
    'http://www.vbAccelerator.com
    Dim Max As Single
    Dim Min As Single
    Dim delta As Single
    Dim rR As Single, rG As Single, rB As Single

    rR = R / 255: rG = G / 255: rB = B / 255

    '{Given: rgb each in [0,1].
    ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    L = (Max + Min) / 2    '{This is the lightness}
    '{Next calculate saturation}
    If Max = Min Then
        'begin {Acrhomatic case}
        S = 0
        H = 0
        'end {Acrhomatic case}
    Else
        'begin {Chromatic case}
             '{First calculate the saturation.}
        If L <= 0.5 Then
            S = (Max - Min) / (Max + Min)
        Else
            S = (Max - Min) / (2 - Max - Min)
        End If
        
        '{Next calculate the hue.}
        delta = Max - Min
        If rR = Max Then
            H = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            H = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            H = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
        End If
        'Debug.Print h
        'h = h * 60
        'If h < 0# Then
        '     h = h + 360            '{Make degrees be nonnegative}
        'End If
    'end {Chromatic Case}
    End If
'end {RGB_to_HLS}
End Sub

Private Sub HSLToRGB(ByVal H As Single, ByVal S As Single, ByVal L As Single, R As Byte, G As Byte, B As Byte)
    'http://www.vbAccelerator.com
    Dim rR As Single, rG As Single, rB As Single
    Dim Min As Single, Max As Single
    
    If S = 0 Then
        ' Achromatic case:
        rR = L: rG = L: rB = L
    Else
        ' Chromatic case:
        ' delta = Max-Min
        If L <= 0.5 Then
            'S = (Max - Min) / (Max + Min)
            ' Get Min value:
            Min = L * (1 - S)
        Else
            'S = (Max - Min) / (2 - Max - Min)
            ' Get Min value:
            Min = L - S * (1 - L)
        End If
        ' Get the Max value:
        Max = 2 * L - Min
       
        ' Now depending on sector we can evaluate the h,l,s:
        If (H < 1) Then
            rR = Max
            If (H < 0) Then
                rG = Min
                rB = rG - H * (Max - Min)
            Else
                rB = Min
                rG = H * (Max - Min) + rB
            End If
        ElseIf (H < 3) Then
            rG = Max
            If (H < 2) Then
                rB = Min
                rR = rB - (H - 2) * (Max - Min)
            Else
                rR = Min
                rB = (H - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If (H < 4) Then
                rR = Min
                rG = rR - (H - 4) * (Max - Min)
            Else
                rG = Min
                rR = (H - 4) * (Max - Min) + rG
            End If
        End If
    End If
    R = rR * 255: G = rG * 255: B = rB * 255
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR > rG) Then
        If (rR > rB) Then
            Maximum = rR
        Else
            Maximum = rB
        End If
    Else
        If (rB > rG) Then
            Maximum = rB
        Else
            Maximum = rG
        End If
    End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR < rG) Then
        If (rR < rB) Then
            Minimum = rR
        Else
            Minimum = rB
        End If
    Else
        If (rB < rG) Then
            Minimum = rB
        Else
            Minimum = rG
        End If
    End If
End Function

Public Function MakeSemiTransparent(ByVal hWnd As Long, ByVal Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
     
    Perc = ((100 - Perc) / 100) * 255
    If Perc < 0 Or Perc > 255 Then
        MakeSemiTransparent = 1
    Else
        Msg = GetWindowLong(hWnd, G)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hWnd, G, Msg
        
        ' Set the Form header bottom colour
        Col = myForm.BackColor
        lBottomR = (Col And &HFF&)
        lBottomG = (Col And &HFF00&) / &H100
        lBottomB = (Col And &HFF0000) / &H10000
        
        SetLayeredWindowAttributes hWnd, RGB(lBottomR, lBottomG, lBottomB), Perc, LWA_ALPHA
        MakeSemiTransparent = 0
    End If
    If Err Then
        MakeSemiTransparent = 2
    End If
End Function

Private Sub ShowInTheTaskbar(hWnd As Long, bShow As Boolean)
    Dim lStyle As Long
    ShowWindow hWnd, SW_HIDE
    
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    If bShow = False Then
        If lStyle And WS_EX_APPWINDOW Then
            lStyle = lStyle - WS_EX_APPWINDOW
        End If
    Else
        lStyle = lStyle Or WS_EX_APPWINDOW
    End If
    
    SetWindowLong hWnd, GWL_EXSTYLE, lStyle
    
    App.TaskVisible = bShow
    
    ShowWindow hWnd, SW_NORMAL
End Sub

Public Property Get Caption() As String
    Caption = lblFormCaption.Caption
End Property

Public Property Let Caption(ByVal New_TheCaption As String)
    lblFormCaption.Caption() = New_TheCaption
    PropertyChanged "Caption"
    Call UserControl_Paint
End Property

Public Property Get DisplayIcon() As Boolean
    DisplayIcon = bDisplayIcon
End Property

Public Property Let DisplayIcon(ByVal New_DisplayIcon As Boolean)
    bDisplayIcon = New_DisplayIcon
    PropertyChanged "DisplayIcon"
    Call UserControl_Resize
End Property

Public Property Get Font() As Font
    Set Font = lblFormCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblFormCaption.Font = New_Font
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblFormCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblFormCaption.ForeColor() = New_ForeColor
    lFormCaptionColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Icon() As Picture
    Set Icon = imgFormPic.Picture
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set imgFormPic.Picture = New_Icon
    Set myForm.Icon = New_Icon
    PropertyChanged "Icon"
    bDisplayIcon = True
    Call UserControl_Resize
End Property

Public Property Get ShowCloseButton() As Boolean
    ShowCloseButton = bCloseButton
End Property

Public Property Let ShowCloseButton(ByVal New_ShowCloseButton As Boolean)
    bCloseButton = New_ShowCloseButton
    PropertyChanged "ShowCloseButton"
    Call UserControl_Resize
End Property

Public Property Get ShowMinimiseButton() As Boolean
    ShowMinimiseButton = bMinimiseButton
End Property

Public Property Let ShowMinimiseButton(ByVal New_ShowMinimiseButton As Boolean)
    bMinimiseButton = New_ShowMinimiseButton
    PropertyChanged "ShowMinimiseButton"
    Call UserControl_Resize
End Property

Public Property Get ShowMaximiseButton() As Boolean
    ShowMaximiseButton = bMaximiseButton
End Property

Public Property Let ShowMaximiseButton(ByVal New_ShowMaximiseButton As Boolean)
    bMaximiseButton = New_ShowMaximiseButton
    PropertyChanged "ShowMaximiseButton"
    Call UserControl_Resize
End Property

Public Property Get Style() As xVistaStyles
    Style = xVisualStyles
End Property

Public Property Let Style(val As xVistaStyles)
    ' Determine which color scheme has been selected
    xVisualStyles = val

    ' Set the colour scheme
    Call SelectColorScheme

    lblFormCaption.ForeColor() = lFormCaptionColor
    lFormCaptionColor = lFormCaptionColor
    PropertyChanged "ForeColor"

    ' Repaint the control
    Call UserControl_Paint

    ' Draw the Form Header and Buttons
    bPaintForm = False
    Call picForm_Paint
End Property

Public Property Get Transparency() As Boolean
    Transparency = bTransparency
End Property

Public Property Let Transparency(ByVal New_Transparency As Boolean)
    bTransparency = New_Transparency
    PropertyChanged "Transparency"
    Call UserControl_Resize
End Property

Public Property Get TransparencyLevel() As String
    TransparencyLevel = iTransparency
End Property

Public Property Let TransparencyLevel(ByVal New_TransparencyLevel As String)
    iTransparency = New_TransparencyLevel
    PropertyChanged "TransparencyLevel"
    Call UserControl_Resize
End Property

Private Sub lblFormCaption_DblClick()
    If myForm.WindowState <> vbMaximized Then
        myForm.WindowState = vbMaximized
    Else
        myForm.WindowState = vbNormal
    End If
End Sub

Private Sub SelectColorScheme()
    Select Case xVisualStyles
        Case VistaLite
            lFormCaptionColor = &H0&
            lFormGradientBottom = &HEAD1B9
            lFormGradientTop = &HD0B498
            lFormMiddleBorder = &HE4CF28
            lFormOuterBorder = &H0&
            lButtonGradientBottom(1) = &HCCB198
            lButtonGradientBottom(2) = &HD1B79E
            lButtonGradientBottom(3) = &HD8BEA4
            lButtonGradientBottom(4) = &HDFC5AC
            lButtonGradientBottom(5) = &HDFC5AC
            lButtonGradientBottom(6) = &HE5CBB2
            lButtonGradientBottom(7) = &HE9D0B7
            lButtonGradientBottomClicked(1) = &H523B20
            lButtonGradientBottomClicked(2) = &H5B441F
            lButtonGradientBottomClicked(3) = &H736223
            lButtonGradientBottomClicked(4) = &H918727
            lButtonGradientBottomClicked(5) = &H918727
            lButtonGradientBottomClicked(6) = &HACA82B
            lButtonGradientBottomClicked(7) = &HC8C927
            lButtonGradientBottomHover(1) = &HA3732D
            lButtonGradientBottomHover(2) = &HAF7B2C
            lButtonGradientBottomHover(3) = &HBF892C
            lButtonGradientBottomHover(4) = &HD09A2C
            lButtonGradientBottomHover(5) = &HD09A2C
            lButtonGradientBottomHover(6) = &HDFA929
            lButtonGradientBottomHover(7) = &HEBC624
            lButtonGradientTop = &HE7D3C1
            lButtonGradientTopClicked = &H9C886E
            lButtonGradientTopHover = &HEFCB96
            lButtonInnerBorder = &HF2E7DE
            lButtonOuterBorder = &H886F5D
            lCloseButtonGradientBottom(1) = &H2C43B8
            lCloseButtonGradientBottom(2) = &H3249BA
            lCloseButtonGradientBottom(3) = &H3F54BF
            lCloseButtonGradientBottom(4) = &H4F62C5
            lCloseButtonGradientBottom(5) = &H4F62C5
            lCloseButtonGradientBottom(6) = &H6373CD
            lCloseButtonGradientBottom(7) = &H7685D5
            lCloseButtonGradientBottomClicked(1) = &H1883&
            lCloseButtonGradientBottomClicked(2) = &H1987&
            lCloseButtonGradientBottomClicked(3) = &H12B85
            lCloseButtonGradientBottomClicked(4) = &H124391
            lCloseButtonGradientBottomClicked(5) = &H124391
            lCloseButtonGradientBottomClicked(6) = &H2C68A8
            lCloseButtonGradientBottomClicked(7) = &H4A93C1
            lCloseButtonGradientBottomHover(1) = &H223D2
            lCloseButtonGradientBottomHover(2) = &H223D2
            lCloseButtonGradientBottomHover(3) = &HD33D5
            lCloseButtonGradientBottomHover(4) = &H2151DA
            lCloseButtonGradientBottomHover(5) = &H2151DA
            lCloseButtonGradientBottomHover(6) = &H3974E0
            lCloseButtonGradientBottomHover(7) = &H56A0E8
            lCloseButtonGradientTop = &H929FE4
            lCloseButtonGradientTopClicked = &H768FBF
            lCloseButtonGradientTopHover = &HADB9FC
            lCloseButtonInnerBorder = &HCCD3F4
            lCloseButtonOuterBorder = &H221443
        Case VistaDark
            lFormCaptionColor = &HFFFFFF
            lFormGradientBottom = &H322624
            lFormGradientTop = &H9E9794
            lFormMiddleBorder = &HB9B8B4
            lFormOuterBorder = &H0&
            lButtonGradientBottom(1) = &H433E35
            lButtonGradientBottom(2) = &H464138
            lButtonGradientBottom(3) = &H413E36
            lButtonGradientBottom(4) = &H403D35
            lButtonGradientBottom(5) = &H403C37
            lButtonGradientBottom(6) = &H413D38
            lButtonGradientBottom(7) = &H423E39
            lButtonGradientBottomClicked(1) = &H523B20
            lButtonGradientBottomClicked(2) = &H5B441F
            lButtonGradientBottomClicked(3) = &H736223
            lButtonGradientBottomClicked(4) = &H918727
            lButtonGradientBottomClicked(5) = &H918727
            lButtonGradientBottomClicked(6) = &HACA82B
            lButtonGradientBottomClicked(7) = &HC8C927
            lButtonGradientBottomHover(1) = &HA3732D
            lButtonGradientBottomHover(2) = &HAF7B2C
            lButtonGradientBottomHover(3) = &HBF892C
            lButtonGradientBottomHover(4) = &HD09A2C
            lButtonGradientBottomHover(5) = &HD09A2C
            lButtonGradientBottomHover(6) = &HDFA929
            lButtonGradientBottomHover(7) = &HEBC624
            lButtonGradientTop = &HC0BDB8
            lButtonGradientTopClicked = &H9C886E
            lButtonGradientTopHover = &HEFCB96
            lButtonInnerBorder = &HB9B8B4
            lButtonOuterBorder = &H322624 '&H886F5D
            lCloseButtonGradientBottom(1) = &H2C43B8
            lCloseButtonGradientBottom(2) = &H3249BA
            lCloseButtonGradientBottom(3) = &H3F54BF
            lCloseButtonGradientBottom(4) = &H4F62C5
            lCloseButtonGradientBottom(5) = &H4F62C5
            lCloseButtonGradientBottom(6) = &H6373CD
            lCloseButtonGradientBottom(7) = &H7685D5
            lCloseButtonGradientBottomClicked(1) = &H1883&
            lCloseButtonGradientBottomClicked(2) = &H1987&
            lCloseButtonGradientBottomClicked(3) = &H12B85
            lCloseButtonGradientBottomClicked(4) = &H124391
            lCloseButtonGradientBottomClicked(5) = &H124391
            lCloseButtonGradientBottomClicked(6) = &H2C68A8
            lCloseButtonGradientBottomClicked(7) = &H4A93C1
            lCloseButtonGradientBottomHover(1) = &H223D2
            lCloseButtonGradientBottomHover(2) = &H223D2
            lCloseButtonGradientBottomHover(3) = &HD33D5
            lCloseButtonGradientBottomHover(4) = &H2151DA
            lCloseButtonGradientBottomHover(5) = &H2151DA
            lCloseButtonGradientBottomHover(6) = &H3974E0
            lCloseButtonGradientBottomHover(7) = &H56A0E8
            lCloseButtonGradientTop = &H929FE4
            lCloseButtonGradientTopClicked = &H768FBF
            lCloseButtonGradientTopHover = &HADB9FC
            lCloseButtonInnerBorder = &HCCD3F4
            lCloseButtonOuterBorder = &H221443
    End Select
End Sub

Private Sub UserControlsCreate()
    If iNumControls = 0 Then
        ' Create the controls only once
        iNumControls = 1
        
        ' Add the Frame Header picturebox
        Set picForm = UserControl.Controls.Add("VB.PictureBox", "picForm")
        picForm.AutoRedraw = True
        picForm.BorderStyle = 0
        picForm.Visible = True
        
        ' Add the Frame Timer
        Set tmrMouseMove = UserControl.Controls.Add("VB.Timer", "tmrMouseMove")
        tmrMouseMove.Enabled = False
        tmrMouseMove.Interval = 10
        
        ' Add the Frame Caption
        Set lblFormCaption = Controls.Add("VB.Label", "lblFormCaption", picForm)
        lblFormCaption.BackStyle = 0
        lblFormCaption.Caption = UserControl.Extender.Name
        lblFormCaption.ForeColor = lFormCaptionColor
        lblFormCaption.Visible = True
        
        ' Add the Frame Header Image
        Set imgFormPic = Controls.Add("VB.Image", "imgFormPic", picForm)
        Set imgFormPic.Picture = Nothing
    End If
End Sub

Private Sub lblFormCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveForm_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblFormCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 75 Then Y = 75
    Call picForm_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblFormCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveForm_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub moveForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long

    If Button = vbLeftButton Then
        If bMinimiseButtonHover = True Then
            bCloseButtonHover = False
            bCloseButtonClicked = False
            bMaximiseButtonClicked = False
            bMaximiseButtonHover = False
            bMinimiseButtonClicked = True
            Call DrawMinimiseButton
            picForm.Refresh
        ElseIf bMaximiseButtonHover = True Then
            bCloseButtonHover = False
            bCloseButtonClicked = False
            bMaximiseButtonClicked = True
            bMinimiseButtonClicked = False
            bMinimiseButtonHover = False
            Call DrawMaximiseButton
            picForm.Refresh
        ElseIf bCloseButtonHover = True Then
            bCloseButtonClicked = True
            bMaximiseButtonClicked = False
            bMaximiseButtonHover = False
            bMinimiseButtonClicked = False
            bMinimiseButtonHover = False
            Call DrawCloseButton
            picForm.Refresh
        End If
        
        If X < (UserControl.Width - 1500) And (Y > 30 And Y <= 375) Then
            Call ReleaseCapture
            myForm.MousePointer = vbSizeAll
            lngReturnValue = SendMessage(moveForm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
            myForm.MousePointer = vbNormal
        End If
        
        Dim RLeft As Single
        Dim RTop As Single
        Dim RRight As Boolean
        Dim RBottom As Boolean
        Dim StartX As Single
        Dim StartY As Single
        Dim MX As Single
        Dim MY As Single

        ' Read the mouse pointer screen position on the beginning
        GetMouseXY StartX, StartY
        ' We don't use the X,Y arguments which are incorect when MouseDown event called
        ' from other screen objects (like Label1 and Image1 hereunder)
        X = StartX - myForm.Left
        Y = StartY - myForm.Top

        ' Flags indicating "from where" the form is being resized
        RLeft = IIf(X < BorderWidth, myForm.Width, 0)
        RTop = IIf(Y < BorderHeight, myForm.Height, 0)
        RRight = (X > myForm.Width - BorderWidth)
        RBottom = (Y > myForm.Height - BorderHeight)

        ' Place the mouse pointer on the form border for more accuracy
        If RLeft Then SetMouseXY myForm.Left, myForm.Top + Y
        If RTop Then SetMouseXY myForm.Left + X, myForm.Top
        If RRight Then SetMouseXY myForm.Left + myForm.Width, myForm.Top + Y
        If RBottom Then SetMouseXY myForm.Left + X, myForm.Top + myForm.Height

        ' Save the mouse pointer screen position on the beginning in variables
        GetMouseXY StartX, StartY
        ' While left mouse button is pressed
        While GetAsyncKeyState(vbLeftButton) < 0
            ' Read the actual mouse pointer screen position
            GetMouseXY MX, MY

            If RRight Or RLeft Or RBottom Or RTop Then      ' If the form is resized (not moved)
                If RLeft And RLeft + StartX - MX > BorderWidth * 2 Then myForm.Move MX, myForm.Top, RLeft + StartX - MX
                If RTop And RTop + StartY - MY > BorderHeight * 2 Then myForm.Move myForm.Left, MY, myForm.Width, RTop + StartY - MY
                If RRight And MX - myForm.Left > BorderWidth * 2 Then myForm.Width = MX - myForm.Left
                If RBottom And MY - myForm.Top > BorderHeight * 2 Then myForm.Height = MY - myForm.Top
''                Else                                            ' If the form is moved (not resized)
''                    MousePointer = vbSizeAll                    ' Sets the mouse cursor showing move
''                    myForm.Move MX - X, MY - Y                  ' Actually moves the form on screen
            End If
            DoEvents                                        ' To allow Windows painting events
        Wend

        Call DrawMinimiseButton
        Call DrawMaximiseButton
        Call DrawCloseButton
        myForm.Refresh
    End If
End Sub

Private Sub moveForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Reset the mouse cursor when left mouse button is not pressed
    If GetAsyncKeyState(vbLeftButton) >= 0 Then myForm.MousePointer = 0
    ' Set the correct mouse cursor according to its position on the form
    If (X + BorderWidth) > myForm.Width - BorderWidth Or X < BorderWidth Then myForm.MousePointer = vbSizeWE
    If Y > myForm.Height - BorderHeight Or Y < 30 Then myForm.MousePointer = vbSizeNS
    If Y > 30 And Y <= (UserControl.Height + 15) And (X > 30 And X < UserControl.Width - 30) Then myForm.MousePointer = vbNormal
    If (X + BorderWidth > myForm.Width - BorderWidth And Y > myForm.Height - BorderHeight) Or (X < BorderWidth And Y < BorderHeight) Then myForm.MousePointer = vbSizeNWSE
    If (X + BorderWidth > myForm.Width - BorderWidth And Y < BorderHeight) Or (X < BorderWidth And Y > myForm.Height - BorderHeight) Then myForm.MousePointer = vbSizeNESW
End Sub

Private Sub moveForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    MsgBox bMinimiseButtonClicked & " " & bMaximiseButtonClicked & " " & bCloseButtonClicked
    If bMinimiseButtonHover = True And bMinimiseButton = True Then
        myForm.WindowState = 1
    ElseIf bMaximiseButtonHover = True And bMaximiseButton = True Then
        If myForm.WindowState <> 2 Then
            myForm.WindowState = 2
        Else
            myForm.WindowState = 0
        End If
    ElseIf bCloseButtonHover = True And bCloseButton = True Then
        tmrMouseMove.Enabled = False
        Unload moveForm
        Unload myForm
        bUnloadForm = True
    End If
End Sub

Private Sub moveForm_Resize()
    Call UserControl_Resize
End Sub

' Reads actual mouse pointer screen position and convert it to TWIP scale
Private Sub GetMouseXY(X As Single, Y As Single)
    Dim lpPoint As POINTAPI

    GetCursorPos lpPoint
    X = lpPoint.X * TwipX
    Y = lpPoint.Y * TwipY
End Sub

' Places mouse pointer on given screen position given in TWIP scale
Private Sub SetMouseXY(ByVal X As Single, ByVal Y As Single)
    SetCursorPos X / TwipX, Y / TwipY
End Sub

Private Sub HoverStateNoIcon(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X >= (UserControl.Width - 1570) And X <= (UserControl.Width - 1120)) And (Y >= 75 And Y <= 240) Then
        bCloseButtonHover = False
        bMaximiseButtonHover = False
        bMinimiseButtonHover = True
        If Button = 1 Then bMinimiseButtonClicked = True
    ElseIf X >= (UserControl.Width - 1075) And X <= (UserControl.Width - 645) And (Y >= 75 And Y <= 240) Then
        bCloseButtonHover = False
        bMaximiseButtonHover = True
        bMinimiseButtonHover = False
        If Button = 1 Then bMaximiseButtonClicked = True
    ElseIf X >= (UserControl.Width - 600) And X <= (UserControl.Width - 180) And (Y >= 75 And Y <= 240) Then
        bCloseButtonHover = True
        bMaximiseButtonHover = False
        bMinimiseButtonHover = False
        If Button = 1 Then bCloseButtonClicked = True
    Else
        bCloseButtonClicked = False
        bCloseButtonHover = False
        bMaximiseButtonClicked = False
        bMaximiseButtonHover = False
        bMinimiseButtonClicked = False
        bMinimiseButtonHover = False
    End If
End Sub

Private Sub HoverStateIcon(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X >= (UserControl.Width - 1855) And X <= (UserControl.Width - 1420)) And (Y >= 75 And Y <= 240) Then
        bCloseButtonHover = False
        bMaximiseButtonHover = False
        bMinimiseButtonHover = True
        If Button = 1 Then bMinimiseButtonClicked = True
    ElseIf X >= (UserControl.Width - 1360) And X <= (UserControl.Width - 960) And (Y >= 75 And Y <= 240) Then
        bCloseButtonHover = False
        bMaximiseButtonHover = True
        bMinimiseButtonHover = False
        If Button = 1 Then bMaximiseButtonClicked = True
    ElseIf X >= (UserControl.Width - 885) And X <= (UserControl.Width - 465) And (Y >= 75 And Y <= 240) Then
        bCloseButtonHover = True
        bMaximiseButtonHover = False
        bMinimiseButtonHover = False
        If Button = 1 Then bCloseButtonClicked = True
    Else
        bCloseButtonClicked = False
        bCloseButtonHover = False
        bMaximiseButtonClicked = False
        bMaximiseButtonHover = False
        bMinimiseButtonClicked = False
        bMinimiseButtonHover = False
    End If
End Sub

Private Sub picForm_DblClick()
    Call lblFormCaption_DblClick
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveForm_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Set the mouse to default
    If Y > 30 And (X > 30 Or X < UserControl.Width - 30) Then myForm.MousePointer = vbNormal
    
    ' Determine which button has hover focus
    If bDisplayIcon = False And myForm.WindowState = 0 Then
        Call HoverStateNoIcon(Button, Shift, X, Y)
    ElseIf bDisplayIcon = True And myForm.WindowState = 0 Then
        Call HoverStateIcon(Button, Shift, X, Y)
    ElseIf bDisplayIcon = False And myForm.WindowState = 2 Then
        Call HoverStateNoIcon(Button, Shift, X, Y)
    ElseIf bDisplayIcon = True And myForm.WindowState = 2 Then
        Call HoverStateIcon(Button, Shift, X, Y)
    End If
    
    ' Enable the timer
    tmrMouseMove.Enabled = True
    
    Call moveForm_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub DrawCloseButton()
    On Error GoTo ErrHandler
    If bCloseButton = True Then
        ' Draw the Buttons Outside Border
        picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 75)-(UserControl.Width - 105, UserControl.Extender.Top + 75), lCloseButtonOuterBorder
        picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 315)-(UserControl.Width - 105, UserControl.Extender.Top + 315), lCloseButtonOuterBorder
        picForm.Line (UserControl.Width - 525, UserControl.Extender.Top + 90)-(UserControl.Width - 525, UserControl.Extender.Top + 315), lCloseButtonOuterBorder
        picForm.Line (UserControl.Width - 105, UserControl.Extender.Top + 90)-(UserControl.Width - 105, UserControl.Extender.Top + 315), lCloseButtonOuterBorder

        ' Draw the Buttons Inside Border
        picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 90), lCloseButtonInnerBorder
        picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 300)-(UserControl.Width - 105, UserControl.Extender.Top + 300), lCloseButtonInnerBorder
        picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 90)-(UserControl.Width - 510, UserControl.Extender.Top + 300), lCloseButtonInnerBorder
        picForm.Line (UserControl.Width - 120, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 300), lCloseButtonInnerBorder

        ' Close Button Top Gradient Base Colour
        iVertical = 105
        For i = 1 To 6
            If bCloseButtonHover = False And bCloseButtonClicked = False Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientTop
            ElseIf bCloseButtonHover = True And bCloseButtonClicked = False Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientTopHover
            ElseIf bCloseButtonHover = True And bCloseButtonClicked = True Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientTopClicked
            End If
            iVertical = iVertical + 15
        Next
        
        ' Button Bottom Gradient Base Colour
        For i = 1 To 7
            If bCloseButtonHover = False And bCloseButtonClicked = False Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientBottom(i)
            ElseIf bCloseButtonHover = True And bCloseButtonClicked = False Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientBottomHover(i)
            ElseIf bCloseButtonHover = True And bCloseButtonClicked = True Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientBottomClicked(i)
            End If
            iVertical = iVertical + 15
        Next
        
        ' Draw the Close Button Display
        ' Outside Borders
        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 135)-(UserControl.Width - 330, UserControl.Extender.Top + 135), &H665653
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 135)-(UserControl.Width - 240, UserControl.Extender.Top + 135), &H665653
        
        picForm.Line (UserControl.Width - 390, UserControl.Extender.Top + 150)-(UserControl.Width - 375, UserControl.Extender.Top + 150), &H665653
        picForm.Line (UserControl.Width - 330, UserControl.Extender.Top + 150)-(UserControl.Width - 315, UserControl.Extender.Top + 150), &H665653
        picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 150)-(UserControl.Width - 285, UserControl.Extender.Top + 150), &H665653
        picForm.Line (UserControl.Width - 240, UserControl.Extender.Top + 150)-(UserControl.Width - 225, UserControl.Extender.Top + 150), &H665653
        
        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 165)-(UserControl.Width - 360, UserControl.Extender.Top + 165), &H665653
        picForm.Line (UserControl.Width - 315, UserControl.Extender.Top + 165)-(UserControl.Width - 300, UserControl.Extender.Top + 165), &H665653
        picForm.Line (UserControl.Width - 255, UserControl.Extender.Top + 165)-(UserControl.Width - 240, UserControl.Extender.Top + 165), &H665653
       
        picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 180)-(UserControl.Width - 345, UserControl.Extender.Top + 180), &H665653
        picForm.Line (UserControl.Width - 270, UserControl.Extender.Top + 180)-(UserControl.Width - 255, UserControl.Extender.Top + 180), &H665653

        picForm.Line (UserControl.Width - 345, UserControl.Extender.Top + 195)-(UserControl.Width - 330, UserControl.Extender.Top + 195), &H665653
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 195)-(UserControl.Width - 270, UserControl.Extender.Top + 195), &H665653

        picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 210)-(UserControl.Width - 345, UserControl.Extender.Top + 210), &H665653
        picForm.Line (UserControl.Width - 270, UserControl.Extender.Top + 210)-(UserControl.Width - 255, UserControl.Extender.Top + 210), &H665653

        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 225)-(UserControl.Width - 360, UserControl.Extender.Top + 225), &H665653
        picForm.Line (UserControl.Width - 315, UserControl.Extender.Top + 225)-(UserControl.Width - 300, UserControl.Extender.Top + 225), &H665653
        picForm.Line (UserControl.Width - 255, UserControl.Extender.Top + 225)-(UserControl.Width - 240, UserControl.Extender.Top + 225), &H665653

        picForm.Line (UserControl.Width - 390, UserControl.Extender.Top + 240)-(UserControl.Width - 375, UserControl.Extender.Top + 240), &H665653
        picForm.Line (UserControl.Width - 330, UserControl.Extender.Top + 240)-(UserControl.Width - 315, UserControl.Extender.Top + 240), &H665653
        picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 240)-(UserControl.Width - 285, UserControl.Extender.Top + 240), &H665653
        picForm.Line (UserControl.Width - 240, UserControl.Extender.Top + 240)-(UserControl.Width - 225, UserControl.Extender.Top + 240), &H665653

        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 255)-(UserControl.Width - 330, UserControl.Extender.Top + 255), &H665653
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 255)-(UserControl.Width - 240, UserControl.Extender.Top + 255), &H665653

        ' Inside Button Colours
        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 150)-(UserControl.Width - 330, UserControl.Extender.Top + 150), &HFFFFFF
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 150)-(UserControl.Width - 240, UserControl.Extender.Top + 150), &HFFFFFF
        
        picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 165)-(UserControl.Width - 315, UserControl.Extender.Top + 165), &HFFFFFF
        picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 165)-(UserControl.Width - 255, UserControl.Extender.Top + 165), &HFFFFFF
        
        picForm.Line (UserControl.Width - 345, UserControl.Extender.Top + 180)-(UserControl.Width - 270, UserControl.Extender.Top + 180), &HFFFFFF
        
        picForm.Line (UserControl.Width - 330, UserControl.Extender.Top + 195)-(UserControl.Width - 285, UserControl.Extender.Top + 195), &HE9E9E9
        
        picForm.Line (UserControl.Width - 345, UserControl.Extender.Top + 210)-(UserControl.Width - 270, UserControl.Extender.Top + 210), &HE2E2E2
        
        picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 225)-(UserControl.Width - 315, UserControl.Extender.Top + 225), &HDCDCDC
        picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 225)-(UserControl.Width - 255, UserControl.Extender.Top + 225), &HDCDCDC
        
        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 240)-(UserControl.Width - 330, UserControl.Extender.Top + 240), &HD7D7D7
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 240)-(UserControl.Width - 240, UserControl.Extender.Top + 240), &HD7D7D7
        
        ' Border Single Pixel Colours
        If xVisualStyles = VistaLite Then
            picForm.Line (UserControl.Width - 525, UserControl.Extender.Top + 75)-(UserControl.Width - 510, UserControl.Extender.Top + 75), &HBCA6A0
            picForm.Line (UserControl.Width - 105, UserControl.Extender.Top + 75)-(UserControl.Width - 90, UserControl.Extender.Top + 75), &HBCA6A0
            picForm.Line (UserControl.Width - 525, UserControl.Extender.Top + 315)-(UserControl.Width - 510, UserControl.Extender.Top + 315), &HBCA6A0
            picForm.Line (UserControl.Width - 105, UserControl.Extender.Top + 315)-(UserControl.Width - 90, UserControl.Extender.Top + 315), &HBCA6A0
        End If
    End If
ErrHandler:
End Sub

Private Sub DrawMinimiseButton()
    On Error GoTo ErrHandler
    
    ' Draw the Buttons
    If bMinimiseButton = True Then
        ' Draw the Minimise Buttons Outside Border
        If xVisualStyles = VistaLite Then
            picForm.Line (UserControl.Width - 1485, UserControl.Extender.Top + 75)-(UserControl.Width - 1035, UserControl.Extender.Top + 75), &HC4AA91
            picForm.Line (UserControl.Width - 1485, UserControl.Extender.Top + 315)-(UserControl.Width - 1035, UserControl.Extender.Top + 315), &HC4AA91
        End If
        
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 75)-(UserControl.Width - 1050, UserControl.Extender.Top + 75), lButtonOuterBorder
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 315)-(UserControl.Width - 1050, UserControl.Extender.Top + 315), lButtonOuterBorder
        picForm.Line (UserControl.Width - 1485, UserControl.Extender.Top + 90)-(UserControl.Width - 1485, UserControl.Extender.Top + 315), lButtonOuterBorder
        picForm.Line (UserControl.Width - 1050, UserControl.Extender.Top + 90)-(UserControl.Width - 1050, UserControl.Extender.Top + 315), lButtonOuterBorder

        ' Draw the Buttons Inside Border
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 90)-(UserControl.Width - 1070, UserControl.Extender.Top + 90), lButtonInnerBorder
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 300)-(UserControl.Width - 1055, UserControl.Extender.Top + 300), lButtonInnerBorder
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 90)-(UserControl.Width - 1470, UserControl.Extender.Top + 300), lButtonInnerBorder
        picForm.Line (UserControl.Width - 1070, UserControl.Extender.Top + 90)-(UserControl.Width - 1070, UserControl.Extender.Top + 300), lButtonInnerBorder

        ' Button Top Gradient Base Colour
        iVertical = 105
        For i = 1 To 6
            If bMinimiseButtonHover = False And bMinimiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientTop
            ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientTopHover
            ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = True Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientTopClicked
            End If
            iVertical = iVertical + 15
        Next
        
        ' Button Bottom Gradient Base Colour
        For i = 1 To 7
            If bMinimiseButtonHover = False And bMinimiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientBottom(i)
            ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientBottomHover(i)
            ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = True Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientBottomClicked(i)
            End If
            iVertical = iVertical + 15
        Next
         
        ' Draw the Minimise Button Display
        ' Outside Borders
        If xVisualStyles = VistaLite Then
            picForm.Line (UserControl.Width - 1350, UserControl.Extender.Top + 195)-(UserControl.Width - 1175, UserControl.Extender.Top + 195), &H9E8A7C    ' Top Border
            picForm.Line (UserControl.Width - 1350, UserControl.Extender.Top + 255)-(UserControl.Width - 1175, UserControl.Extender.Top + 255), &H9E8A7C    ' Bottom Border
        End If
        
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 195)-(UserControl.Width - 1190, UserControl.Extender.Top + 195), &H665653    ' Top Border
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 255)-(UserControl.Width - 1190, UserControl.Extender.Top + 255), &H665653    ' Bottom Border
        picForm.Line (UserControl.Width - 1350, UserControl.Extender.Top + 210)-(UserControl.Width - 1350, UserControl.Extender.Top + 255), &H665653    ' Left Border
        picForm.Line (UserControl.Width - 1190, UserControl.Extender.Top + 210)-(UserControl.Width - 1190, UserControl.Extender.Top + 255), &H665653    ' Right Border
    
        ' Inside Button Display
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 210)-(UserControl.Width - 1190, UserControl.Extender.Top + 210), &HFFFFFF    ' Top Border
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 225)-(UserControl.Width - 1190, UserControl.Extender.Top + 225), &HDCDCDC    ' Middle Border
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 240)-(UserControl.Width - 1190, UserControl.Extender.Top + 240), &HD7D7D7    ' Bottom Border
    End If
ErrHandler:
End Sub

Private Sub DrawMaximiseButton()
    On Error GoTo ErrHandler
    If bMaximiseButton = True Then
        ' Draw the Maximise Buttons Outside Border
        If xVisualStyles = VistaLite Then
            picForm.Line (UserControl.Width - 1005, UserControl.Extender.Top + 75)-(UserControl.Width - 555, UserControl.Extender.Top + 75), &HC4AA91
            picForm.Line (UserControl.Width - 1005, UserControl.Extender.Top + 315)-(UserControl.Width - 555, UserControl.Extender.Top + 315), &HC4AA91
        End If
        
        picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 75)-(UserControl.Width - 570, UserControl.Extender.Top + 75), lButtonOuterBorder
        picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 315)-(UserControl.Width - 570, UserControl.Extender.Top + 315), lButtonOuterBorder
        picForm.Line (UserControl.Width - 1005, UserControl.Extender.Top + 90)-(UserControl.Width - 1005, UserControl.Extender.Top + 315), lButtonOuterBorder
        picForm.Line (UserControl.Width - 570, UserControl.Extender.Top + 90)-(UserControl.Width - 570, UserControl.Extender.Top + 315), lButtonOuterBorder

        ' Draw the Buttons Inside Border
        picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + 90)-(UserControl.Width - 590, UserControl.Extender.Top + 90), lButtonInnerBorder
        picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 300)-(UserControl.Width - 575, UserControl.Extender.Top + 300), lButtonInnerBorder
        picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 90)-(UserControl.Width - 990, UserControl.Extender.Top + 300), lButtonInnerBorder
        picForm.Line (UserControl.Width - 590, UserControl.Extender.Top + 90)-(UserControl.Width - 590, UserControl.Extender.Top + 300), lButtonInnerBorder

        ' Button Top Gradient Base Colour
        iVertical = 105
        For i = 1 To 6
            If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientTop
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientTopHover
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientTopClicked
            End If
            iVertical = iVertical + 15
        Next

        ' Button Bottom Gradient Base Colour
        For i = 1 To 7
            If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientBottom(i)
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientBottomHover(i)
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientBottomClicked(i)
            End If
            iVertical = iVertical + 15
        Next
        
        ' Draw the Maximise Button Display
        ' Inside Button Display
        picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 150), &HFFFFFF
        picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 165)-(UserControl.Width - 705, UserControl.Extender.Top + 165), &HFFFFFF
        picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 180)-(UserControl.Width - 705, UserControl.Extender.Top + 180), &HFFFFFF
        picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 705, UserControl.Extender.Top + 195), &HE9E9E9
        picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 210)-(UserControl.Width - 705, UserControl.Extender.Top + 210), &HE2E2E2
        picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 225)-(UserControl.Width - 705, UserControl.Extender.Top + 225), &HDCDCDC
        picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 240)-(UserControl.Width - 705, UserControl.Extender.Top + 240), &HD7D7D7
        
        ' Outside Borders
        If xVisualStyles = VistaLite Then
            picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 135)-(UserControl.Width - 690, UserControl.Extender.Top + 135), &H9E8A7C    ' Top Border
            picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 255)-(UserControl.Width - 690, UserControl.Extender.Top + 255), &H9E8A7C    ' Bottom Border
        End If
        
        picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 135)-(UserControl.Width - 705, UserControl.Extender.Top + 135), &H665653    ' Top Border
        picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 255)-(UserControl.Width - 705, UserControl.Extender.Top + 255), &H665653    ' Bottom Border
        picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 150)-(UserControl.Width - 870, UserControl.Extender.Top + 255), &H665653    ' Left Border
        picForm.Line (UserControl.Width - 705, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 255), &H665653    ' Right Border

        picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 180)-(UserControl.Width - 735, UserControl.Extender.Top + 180), &H665653    ' Top Border
        picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 215)-(UserControl.Width - 735, UserControl.Extender.Top + 215), &H665653    ' Bottom Border
        picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 180)-(UserControl.Width - 825, UserControl.Extender.Top + 215), &H665653    ' Left Border
        picForm.Line (UserControl.Width - 750, UserControl.Extender.Top + 180)-(UserControl.Width - 750, UserControl.Extender.Top + 215), &H665653    ' Right Border
        
        If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
            picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), lButtonGradientBottom(1)
        ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
            picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), lButtonGradientBottomHover(1)
        ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
            picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), lButtonGradientBottomClicked(1)
        End If
    End If
ErrHandler:
End Sub

Private Sub picForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveForm_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picForm_Paint()
    ' Paints the Form header and label
    If bPaintForm = False Then
        ' Set the Form header bottom colour
        Col = lFormGradientBottom
        lBottomR = (Col And &HFF&)
        lBottomG = (Col And &HFF00&) / &H100
        lBottomB = (Col And &HFF0000) / &H10000

        ' Set the Form header top colour
        Col = lFormGradientTop
        lTopR = (Col And &HFF&)
        lTopG = (Col And &HFF00&) / &H100
        lTopB = (Col And &HFF0000) / &H10000

        ' Clear the Form picturebox for drawing and apply the gradient colour
        picForm.Cls
        Set picForm.Picture = CreateGradient(UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY, False, RGB(lTopR, lTopG, lTopB), RGB(lBottomR, lBottomG, lBottomB), RGBBlend)
        
        ' Draw the Buttons
        Call DrawMinimiseButton
        Call DrawMaximiseButton
        Call DrawCloseButton
        
        ' Top Border Line
        picForm.Line (0, 0)-(picForm.Width - 15, 0), lFormOuterBorder
        
        ' Left Border Line
        picForm.Line (0, 0)-(0, picForm.Height), lFormOuterBorder
        picForm.Line (15, 15)-(15, picForm.Height), UserControl.BackColor
        
        ' Right Border Line
        picForm.Line (picForm.Width - 30, 15)-(picForm.Width - 30, picForm.Height), lFormMiddleBorder
        picForm.Line (picForm.Width - 15, 0)-(picForm.Width - 15, picForm.Height), lFormOuterBorder

        bPaintForm = True
    End If
End Sub

Private Sub tmrMouseMove_Timer()
    Dim pt As POINTAPI

    ' See where the cursor is.
    GetCursorPos pt

    ' Translate into window coordinates.
    If WindowFromPointXY(pt.X, pt.Y) <> picForm.hWnd Then
        bCloseButtonHover = False
        bCloseButtonClicked = False
        bMinimiseButtonHover = False
        bMaximiseButtonClicked = False
        bMaximiseButtonHover = False
        bMinimiseButtonClicked = False
        
        If bMouseOnForm = False Then
            ' Draw the Buttons
            Call DrawMinimiseButton
            Call DrawMaximiseButton
            Call DrawCloseButton
            bMouseOnForm = True
        End If
    Else
        bMouseOnForm = False
        
        ' Draw the Buttons
        Call DrawMinimiseButton
        Call DrawMaximiseButton
        Call DrawCloseButton
    End If
End Sub

Private Sub UserControl_Initialize()
    ' Initialise the default values
    bCloseButton = True
    bCloseButtonClicked = False
    bCloseButtonHover = False
    bDisplayIcon = False
    bMaximiseButton = True
    bMaximiseButtonClicked = False
    bMaximiseButtonHover = False
    bMinimiseButton = True
    bMinimiseButtonClicked = False
    bMinimiseButtonHover = False
    bMouseOnForm = False
    bPaintForm = False
    bTransparency = False
    bUnloadForm = False
    iNumControls = 0
    iTransparency = 15
    
    ' TwipX and TwipY used only for easier writing
    TwipX = Screen.TwipsPerPixelX
    TwipY = Screen.TwipsPerPixelY
    
    ' Following variables used to speed-up the process (prevent recalc of BorderPixels * ...)
    BorderWidth = BorderPixels * TwipX
    BorderHeight = BorderPixels * TwipY
    
    xVisualStyles = VistaLite
End Sub

Private Sub UserControl_InitProperties()
    Call UserControlsCreate
    Set myForm = UserControl.Parent
    myForm.BorderStyle = 0
    
    Call SelectColorScheme
    Call UserControl_Paint
    
    Set myForm.Icon = Nothing
    myForm.AutoRedraw = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call moveForm_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    Set myForm = UserControl.Parent
    myForm.BorderStyle = 0
    
    If bTransparency = True Then Call MakeSemiTransparent(UserControl.Parent.hWnd, iTransparency)
    
    ' Draw the Form Border Lines
    myForm.Cls
    
    ' Left Form Border Line
    myForm.Line (0, 0)-(0, myForm.Height), lFormOuterBorder
    myForm.Line (30, 0)-(30, myForm.Height), lFormGradientBottom
    myForm.Line (45, 0)-(45, myForm.Height), lFormGradientBottom
    
    ' Right Border Line
    myForm.Line (myForm.Width - 15, 0)-(myForm.Width - 15, myForm.Height), lFormOuterBorder
    myForm.Line (myForm.Width - 30, 0)-(myForm.Width - 30, myForm.Height), lFormMiddleBorder
    myForm.Line (myForm.Width - 45, 0)-(myForm.Width - 45, myForm.Height), lFormGradientBottom
    myForm.Line (myForm.Width - 60, 0)-(myForm.Width - 60, myForm.Height), lFormGradientBottom
    
    ' Bottom Border Line
    myForm.Line (30, myForm.Height - 60)-(myForm.Width - 30, myForm.Height - 60), lFormGradientBottom
    myForm.Line (30, myForm.Height - 45)-(myForm.Width - 30, myForm.Height - 45), lFormGradientBottom
    myForm.Line (15, myForm.Height - 30)-(myForm.Width - 15, myForm.Height - 30), lFormMiddleBorder
    myForm.Line (0, myForm.Height - 15)-(myForm.Width - 15, myForm.Height - 15), lFormOuterBorder
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set moveForm = UserControl.Parent
    Set myForm = UserControl.Parent
    myForm.AutoRedraw = True
    
    ' Load saved properties
    xVisualStyles = PropBag.ReadProperty("Style", 1)
    Call SelectColorScheme
    Call UserControlsCreate
    lblFormCaption.Caption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
    bDisplayIcon = PropBag.ReadProperty("DisplayIcon", False)
    Set lblFormCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lFormCaptionColor = PropBag.ReadProperty("ForeColor", &H0&)
    imgFormPic.Picture = PropBag.ReadProperty("Icon", Nothing)
    bCloseButton = PropBag.ReadProperty("ShowCloseButton", True)
    bMinimiseButton = PropBag.ReadProperty("ShowMinimiseButton", True)
    bMaximiseButton = PropBag.ReadProperty("ShowMaximiseButton", True)
    xVisualStyles = PropBag.ReadProperty("Style", 0)
    bTransparency = PropBag.ReadProperty("Transparency", False)
    iTransparency = PropBag.ReadProperty("TransparencyLevel", 15)
    
    Call ShowInTheTaskbar(myForm.hWnd, True)
    If bTransparency = True Then Call MakeSemiTransparent(UserControl.Parent.hWnd, iTransparency)
    Call SelectColorScheme
    lblFormCaption.ForeColor = lFormCaptionColor
    Call UserControl_Paint
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    ' Create the required controls
    Call UserControlsCreate
    
    UserControl.Height = 390
    UserControl.Extender.Align = 1

    ' Position the Picture & Form text
    If bDisplayIcon = False Then
        imgFormPic.Visible = False
        lblFormCaption.Move picForm.Left + 75, picForm.Top + 90, picForm.Width - 90, picForm.Height
    Else
        imgFormPic.Move picForm.Left + 60, picForm.Top + 60, 240, 240
        lblFormCaption.Move imgFormPic.Left + imgFormPic.Width + 75, picForm.Top + 90, picForm.Width - imgFormPic.Width - 90, picForm.Height
        imgFormPic.Visible = True
    End If

    ' Position the Form header
    picForm.Move 0, 0, UserControl.Width, UserControl.Height
    picForm.ZOrder 0

    bPaintForm = False
    Call picForm_Paint
    bPaintForm = True
    
    Call UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
    If bUnloadForm = False Then
        myForm.Cls
        myForm.AutoRedraw = False
    Else
        Unload moveForm
        Unload myForm
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Save properties
    Call PropBag.WriteProperty("Caption", lblFormCaption.Caption, UserControl.Extender.Name)
    Call PropBag.WriteProperty("DisplayIcon", bDisplayIcon, False)
    Call PropBag.WriteProperty("Font", lblFormCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lFormCaptionColor, &H0&)
    Call PropBag.WriteProperty("Icon", imgFormPic.Picture, Nothing)
    Call PropBag.WriteProperty("ShowCloseButton", bCloseButton, True)
    Call PropBag.WriteProperty("ShowMinimiseButton", bMinimiseButton, True)
    Call PropBag.WriteProperty("ShowMaximiseButton", bMaximiseButton, True)
    Call PropBag.WriteProperty("Style", xVisualStyles, 0)
    Call PropBag.WriteProperty("Transparency", bTransparency, False)
    Call PropBag.WriteProperty("TransparencyLevel", iTransparency, 15)
    myForm.AutoRedraw = False
End Sub
