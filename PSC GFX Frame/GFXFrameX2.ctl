VERSION 5.00
Begin VB.UserControl GFXFrameX 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1920
   ScaleWidth      =   2310
   ToolboxBitmap   =   "GFXFrameX2.ctx":0000
   Begin VB.Label LblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   210
      Left            =   270
      TabIndex        =   0
      Top             =   885
      Width           =   525
   End
End
Attribute VB_Name = "GFXFrameX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Enum MyFrameStyle
    [Rounded] = 0
    [Basic] = 1
    End Enum

Public Enum ColType
    [FormAmbient] = 0
    [Selected] = 1
    End Enum

Public Enum AlignOptions
    [AlignLeft] = 0
    [AlignCenter] = 1
    [AlignRight] = 2
    End Enum

Private Type Colours
    R As Double
    G As Double
    B As Double
    L As Long
    End Type

Private vFrameStyle         As MyFrameStyle

Private vBackColour         As OLE_COLOR
Private vBackColourIs       As ColType

Private vFillColour         As OLE_COLOR
Private vFillColourIs       As ColType

Private vBorderColour       As OLE_COLOR

'Private vFont               As Font
'Private vCaption            As String
Private vAlignment          As AlignOptions
Private vShowTab            As Boolean
Private vTabWidthOff        As Integer

Private vFrameOpacity       As Integer
Private vFrameColourIs      As ColType
Private vFrameHighlight     As OLE_COLOR
Private vFrameLowLight      As OLE_COLOR
Private vCornerDepth        As Integer

Private vShowShadow         As Boolean
Private vShadowOpacity      As Integer
Private vShadowDepth        As Integer

Const mShadowOpacity = 50
Const mShadowDepth = 5
Const mBorderColour = vbBlack
Const mCornerDepth = 10

Private Sub RedrawFrame()
    
    Dim UseBackColour As Long
    
    LblCaption.BackColor = FillColour
    LblCaption.Top = P2T(1)
    
    Select Case vAlignment
        Case [AlignLeft]
            LblCaption.Left = P2T(5 + (CornerDepth \ 2) + vTabWidthOff)
            
        Case [AlignCenter]
            LblCaption.Left = (UserControl.Width \ 2) - (LblCaption.Width \ 2) - vTabWidthOff
            If vShowShadow = True Then
                LblCaption.Left = LblCaption.Left - P2T(vShadowDepth)
            End If
            
        Case [AlignRight]
            LblCaption.Left = (UserControl.Width - LblCaption.Width) - P2T(5 + (CornerDepth \ 2) + vTabWidthOff)
            If vShowShadow = True Then
                LblCaption.Left = LblCaption.Left - P2T(vShadowDepth)
            End If
            
    End Select
    
    LblCaption.Refresh
    
    
    If BackColourIs = [FormAmbient] Then
        UserControl.BackColor = Ambient.BackColor
        UseBackColour = Ambient.BackColor
    Else
        UserControl.BackColor = vBackColour
        UseBackColour = vBackColour
    End If
        
    UserControl.Cls
    
    Dim LP As Integer
    Dim LineCol As Long
    Dim ShadowStart As Long
    Dim FrameTop As Integer
    Dim CaptionOff As Integer
     
    If Trim(LblCaption.Caption) <> "" Then
        FrameTop = T2P(LblCaption.Height / 2)
    Else
        FrameTop = 1
    End If
    
    CaptionOff = vTabWidthOff
    
    Select Case vFrameStyle
        Case [Rounded]
            
            If ShowShadow = True Then

                ShadowStart = ColourStepEx(vbBlack, UseBackColour, 100, 100 - ShadowOpacity)
                
                UserControl.FillStyle = 1
                
                For LP = 0 To ShadowDepth
                    LineCol = ColourStepEx(ShadowStart, UseBackColour, ShadowDepth, LP)
                    UserControl.ForeColor = LineCol
                    RoundRect UserControl.hdc, ShadowDepth + (ShadowDepth - LP), FrameTop + ShadowDepth + (ShadowDepth - LP), T2P(UserControl.Width) - (ShadowDepth - LP), T2P(UserControl.Height) - (ShadowDepth - LP), CornerDepth, CornerDepth
                Next LP
                
                SmoothCorners FrameTop
                
                UserControl.FillStyle = 0
                UserControl.FillColor = FillColour
                UserControl.ForeColor = BorderColour
                RoundRect UserControl.hdc, 0, FrameTop, T2P(UserControl.Width) - ShadowDepth, T2P(UserControl.Height) - ShadowDepth, CornerDepth, CornerDepth
                
                'SmoothCorners
                
            Else
                UserControl.FillStyle = 0
                UserControl.FillColor = FillColour
                UserControl.ForeColor = BorderColour
                RoundRect UserControl.hdc, 0, FrameTop, T2P(UserControl.Width), T2P(UserControl.Height), CornerDepth, CornerDepth
            End If
        
        Case [Basic]
        
            If ShowShadow = True Then
                
                ShadowStart = ColourStepEx(vbBlack, UseBackColour, 100, 100 - ShadowOpacity)
                
                UserControl.FillStyle = 1
                
                For LP = 0 To ShadowDepth
                    LineCol = ColourStepEx(ShadowStart, UseBackColour, ShadowDepth, LP)
                    UserControl.ForeColor = LineCol
                    Rectangle UserControl.hdc, ShadowDepth + (ShadowDepth - LP), FrameTop + ShadowDepth + (ShadowDepth - LP), T2P(UserControl.Width) - (ShadowDepth - LP), T2P(UserControl.Height) - (ShadowDepth - LP)
                Next LP
                
                UserControl.FillStyle = 0
                UserControl.FillColor = FillColour
                UserControl.ForeColor = BorderColour
                Rectangle UserControl.hdc, 0, FrameTop, T2P(UserControl.Width) - ShadowDepth, T2P(UserControl.Height) - ShadowDepth
            Else
                UserControl.FillStyle = 0
                UserControl.FillColor = FillColour
                UserControl.ForeColor = BorderColour
                Rectangle UserControl.hdc, 0, FrameTop, T2P(UserControl.Width), T2P(UserControl.Height)
            End If
                    
    End Select
    
    UserControl.FillStyle = 0
    UserControl.FillColor = FillColour
    
    
    If ShowTab = True Then
        UserControl.ForeColor = BorderColour
    Else
        UserControl.ForeColor = FillColour
    End If
    
    If Trim(LblCaption.Caption) <> "" Then
        
        LblCaption.Visible = True
        
        If vFrameStyle = [Rounded] Then
            RoundRect UserControl.hdc, T2P(LblCaption.Left) - CaptionOff, 0, T2P(LblCaption.Left + LblCaption.Width) + CaptionOff - 1, T2P(LblCaption.Height), 8, 8
        Else
            Rectangle UserControl.hdc, T2P(LblCaption.Left) - CaptionOff, 0, T2P(LblCaption.Left + LblCaption.Width) + CaptionOff - 1, T2P(LblCaption.Height)
        End If
    
        UserControl.ForeColor = FillColour
        Rectangle UserControl.hdc, T2P(LblCaption.Left) - CaptionOff, T2P(LblCaption.Height \ 2) + 1, T2P(LblCaption.Left + LblCaption.Width) + CaptionOff - 1, T2P(LblCaption.Height)
    Else
        LblCaption.Visible = False
    End If
    
End Sub

Private Sub SmoothCorners(FrameTop As Integer)
    
    Dim Colours(4) As Colours
    Dim LongColour As Long
    Dim X As Integer, Y As Integer
    Dim LP As Integer
    Dim Rep As Integer
    
    Dim FX As Integer, TX As Integer
    Dim FY As Integer, TY As Integer
    
    For Rep = 0 To 2
        
        'Set The Coordinates To Only Process The Corners
        Select Case Rep
            Case 0 'Bottom Right Corner
                FX = T2P(UserControl.ScaleWidth) - (CornerDepth * 2)
                TX = T2P(UserControl.ScaleWidth) - 3
                FY = T2P(UserControl.ScaleHeight) - (CornerDepth * 2)
                TY = T2P(UserControl.ScaleHeight) - 3
                
            Case 1 'Top Right Corner
                FX = T2P(UserControl.ScaleWidth) - (CornerDepth * 2)
                TX = T2P(UserControl.ScaleWidth) - 3
                FY = FrameTop  'T2P(UserControl.ScaleHeight) - (CornerDepth * 2)
                TY = (ShadowDepth * 4) + FrameTop 'T2P(UserControl.ScaleHeight) - 3
                
            Case 2 'Bottom Left Corner
                FX = ShadowDepth
                TX = ShadowDepth * 4
                FY = T2P(UserControl.ScaleHeight) - (CornerDepth * 2)
                TY = T2P(UserControl.ScaleHeight) - 3
                
        End Select
        
        
        For X = FX To TX
            For Y = FY To TY
            
                LongColour = GetPixel(UserControl.hdc, X, Y)
                
                If LongColour = UserControl.BackColor Then
                
                    For LP = 0 To 3
                        Colours(LP).L = GetPColour(X, Y, LP)
                        CRGB Colours(LP).L, Colours(LP).R, Colours(LP).G, Colours(LP).B
                    Next LP
                    
                    'Find Any Missed Spots In Our Corners, Get The Pixel Values Around It Then
                        'Get The Average And Colour It In.
                    Colours(4).R = (Colours(0).R + Colours(1).R + Colours(2).R + Colours(3).R) / 4
                    Colours(4).G = (Colours(0).G + Colours(1).G + Colours(2).G + Colours(3).G) / 4
                    Colours(4).B = (Colours(0).B + Colours(1).B + Colours(2).B + Colours(3).B) / 4
                    
                    SetPixelV UserControl.hdc, X, Y, RGB(Colours(4).R, Colours(4).G, Colours(4).B)
                
                End If
                
            Next Y
        Next X
    Next Rep
    
End Sub

Private Function GetPColour(X As Integer, Y As Integer, Direction As Integer) As Long
    
    Select Case Direction
        Case 0 'UP
            GetPColour = GetPixel(UserControl.hdc, X, Y - 1)
        Case 1 'Right
            GetPColour = GetPixel(UserControl.hdc, X + 1, Y)
        Case 2 'Down
            GetPColour = GetPixel(UserControl.hdc, X, Y + 1)
        Case 3 'Left
            GetPColour = GetPixel(UserControl.hdc, X - 1, Y)
    End Select
    
End Function


Private Sub UserControl_AmbientChanged(PropertyName As String)
    
    Refresh
    
End Sub

Private Sub UserControl_InitProperties()
    
    vShadowOpacity = mShadowOpacity
    vShadowDepth = mShadowDepth
    vFillColour = Ambient.BackColor
    vBackColour = Ambient.BackColor
    vBackColourIs = [FormAmbient]
    vShowShadow = True
    vFrameStyle = [Basic]
    vBorderColour = mBorderColour
    vCornerDepth = mCornerDepth
    LblCaption.Caption = UserControl.Name
    Set LblCaption.Font = Ambient.Font
    vShowTab = True
    vAlignment = [AlignLeft]
    LblCaption.ForeColor = Ambient.ForeColor
    vTabWidthOff = 4
    
End Sub

Public Property Get FrameStyle() As MyFrameStyle
    
    FrameStyle = vFrameStyle
    
End Property

Public Property Let FrameStyle(ByVal vNewValue As MyFrameStyle)
    
    vFrameStyle = vNewValue
    PropertyChanged "FrameStyle"
    
    Refresh
    
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        vFrameStyle = .ReadProperty("FrameStyle", [Basic])
        vShadowOpacity = .ReadProperty("ShadowOpacity", mShadowOpacity)
        vShowShadow = .ReadProperty("ShowShadow", True)
        vShadowDepth = .ReadProperty("ShadowDepth", mShadowDepth)
        vFillColour = .ReadProperty("FillColour", Ambient.BackColor)
        vBackColour = .ReadProperty("BackColour", Ambient.BackColor)
        vBackColourIs = .ReadProperty("BackColourIs", [FormAmbient])
        vBorderColour = .ReadProperty("BorderColour", mBorderColour)
        vCornerDepth = .ReadProperty("CornerDepth", mCornerDepth)
        LblCaption.Caption = .ReadProperty("Caption", UserControl.Name)
        Set LblCaption.Font = .ReadProperty("Font", Ambient.Font)
        vShowTab = .ReadProperty("ShowTab", True)
        vAlignment = .ReadProperty("Alignment", [AlignLeft])
        LblCaption.ForeColor = .ReadProperty("FontColour", Ambient.ForeColor)
        vTabWidthOff = .ReadProperty("TabWidthOff", 2)
    End With
        
    Refresh
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        .WriteProperty "FrameStyle", vFrameStyle, [Basic]
        .WriteProperty "ShadowOpacity", vShadowOpacity, mShadowOpacity
        .WriteProperty "ShowShadow", vShowShadow, True
        .WriteProperty "ShadowDepth", vShadowDepth, mShadowDepth
        .WriteProperty "FillColour", vFillColour, Ambient.BackColor
        .WriteProperty "BackColour", vBackColour, Ambient.BackColor
        .WriteProperty "BackColourIs", vBackColourIs, [FormAmbient]
        .WriteProperty "BorderColour", vBorderColour, mBorderColour
        .WriteProperty "CornerDepth", vCornerDepth, mCornerDepth
        .WriteProperty "Caption", LblCaption.Caption, UserControl.Name
        .WriteProperty "Font", LblCaption.Font, Ambient.Font
        .WriteProperty "ShowTab", vShowTab, True
        .WriteProperty "Alignment", vAlignment, [AlignLeft]
        .WriteProperty "FontColour", LblCaption.ForeColor, Ambient.ForeColor
        .WriteProperty "TabWidthOff", vTabWidthOff, 2
    End With
    
    
    
End Sub

Public Property Get ShadowOpacity() As Integer
    
    ShadowOpacity = vShadowOpacity
    
End Property

'Shadow Transparency
Public Property Let ShadowOpacity(ByVal vNewValue As Integer)
    
    If vNewValue < 0 Then vNewValue = 0
    If vNewValue > 100 Then vNewValue = 100
    
    vShadowOpacity = vNewValue
    PropertyChanged "ShadowOpacity"
    
    Refresh
    
End Property

Public Property Get ShowShadow() As Boolean
    
    ShowShadow = vShowShadow
    
End Property

'Display The Shadow
Public Property Let ShowShadow(ByVal vNewValue As Boolean)
    
    vShowShadow = vNewValue
    PropertyChanged "ShowShadow"
    
    Refresh
    
End Property

Public Property Get ShadowDepth() As Integer
    
    ShadowDepth = vShadowDepth
    
End Property

'Shadow Legth
Public Property Let ShadowDepth(ByVal vNewValue As Integer)
    
    If vNewValue < 2 Then vNewValue = 2
    If vNewValue > 40 Then vNewValue = 50
    
    vShadowDepth = vNewValue
    PropertyChanged "ShadowDepth"
    
    Refresh
    
End Property

Private Sub UserControl_Resize()
    
    Refresh
    
End Sub


Public Sub Refresh()
    
    RedrawFrame
    
End Sub

Private Function P2T(Value As Integer) As Integer
    
    P2T = Value * Screen.TwipsPerPixelX
    
End Function

Private Function T2P(Value As Integer) As Integer
    
    T2P = Value / Screen.TwipsPerPixelX
    
End Function

'My Own Personal Function For Returning Gradient Values
Private Function ColourStepEx(PenColour As Long, CanvasColour As Long, Steps As Integer, Step As Integer) As Long
    
    Dim R1 As Double, R2 As Double, G1 As Double, G2 As Double, B1 As Double, B2 As Double
    Dim RD As Double, GD As Double, BD As Double
    
    CRGB PenColour, R1, G1, B1
    CRGB CanvasColour, R2, G2, B2
    
    RD = R1 - R2
    GD = G1 - G2
    BD = B1 - B2
    
    RD = (RD \ Steps) * Step
    GD = (GD \ Steps) * Step
    BD = (BD \ Steps) * Step
                            
                            
    ColourStepEx = RGB(R1 - RD, G1 - GD, B1 - BD)
    
End Function

'Get RGB Values From A Long Colour
Private Function CRGB(LongColour As Long, Optional Red As Double, Optional Green As Double, Optional Blue As Double) As Long

    Red = LongColour And 255
    Green = (LongColour \ 256) And 255
    Blue = (LongColour \ 65536) And 255

End Function


Public Property Get FillColour() As OLE_COLOR
    
    FillColour = vFillColour
    
End Property

Public Property Let FillColour(ByVal vNewValue As OLE_COLOR)

    vFillColour = vNewValue
    PropertyChanged "FillColour"
    
    Refresh

End Property

Public Property Get BackColour() As OLE_COLOR
    
    BackColour = vBackColour
    
End Property

Public Property Let BackColour(ByVal vNewValue As OLE_COLOR)
    
    vBackColour = vNewValue
    PropertyChanged "BackColour"
    
    Refresh
    
End Property

Public Property Get BackColourIs() As ColType
    
    BackColourIs = vBackColourIs

End Property

Public Property Let BackColourIs(ByVal vNewValue As ColType)
    
    vBackColourIs = vNewValue
    PropertyChanged "BackColourIs"
    
    Refresh
    
End Property

Public Property Get BorderColour() As OLE_COLOR
    
    BorderColour = vBorderColour
    
End Property

Public Property Let BorderColour(ByVal vNewValue As OLE_COLOR)
    
    vBorderColour = vNewValue
    PropertyChanged "BorderColour"
    
    Refresh
    
End Property

Public Property Get CornerDepth() As Integer
    
    CornerDepth = vCornerDepth
    
End Property

Public Property Let CornerDepth(ByVal vNewValue As Integer)
    
    If vNewValue < 0 Then vNewValue = 0
    If vNewValue > 50 Then vNewValue = 50
    
    vCornerDepth = vNewValue
    PropertyChanged "CornerDepth"
    
    Refresh
    
End Property

Public Property Get Caption() As String
    
    Caption = LblCaption.Caption
    
End Property

Public Property Let Caption(ByVal vNewValue As String)
    
    LblCaption.Caption = vNewValue
    LblCaption.Refresh
    Refresh
    
End Property

Public Property Get Font() As Font
    
    Set Font = LblCaption.Font
    
End Property

Public Property Set Font(ByVal vNewValue As Font)
    
    Set LblCaption.Font = vNewValue
    LblCaption.Refresh
    Refresh
    
End Property

Public Property Get ShowTab() As Boolean
    
    ShowTab = vShowTab
    
End Property

Public Property Let ShowTab(ByVal vNewValue As Boolean)
    
    vShowTab = vNewValue
    PropertyChanged "ShowTab"
    
    Refresh
    
End Property

Public Property Get Alignment() As AlignOptions
    
    Alignment = vAlignment
    
End Property

Public Property Let Alignment(ByVal vNewValue As AlignOptions)
    
    vAlignment = vNewValue
    PropertyChanged "Alignment"
    
    Refresh
End Property

Public Property Get FontColour() As OLE_COLOR
    
    FontColour = LblCaption.ForeColor
    
End Property

Public Property Let FontColour(ByVal vNewValue As OLE_COLOR)
    
    LblCaption.ForeColor = vNewValue
    Refresh
    
End Property

Public Property Get TabWidthOff() As Integer
    
    TabWidthOff = vTabWidthOff
    
End Property

Public Property Let TabWidthOff(ByVal vNewValue As Integer)
    
    If vNewValue < 2 Then vNewValue = 2
    
    vTabWidthOff = vNewValue
    PropertyChanged "TabWidthOff"
    
    Refresh
    
End Property
