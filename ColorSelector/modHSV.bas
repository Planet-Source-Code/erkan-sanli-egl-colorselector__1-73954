Attribute VB_Name = "modHSV"
Option Explicit

Public Enum ColorType
    [tRGB]
    [tHSV]
    [tWhi]
    [tHB]
End Enum

Public Const Resolution     As Single = 255
Public Const sng1Div255     As Single = 0.0039215
Public Const ApproachVal    As Single = 0.0000001

Public NewRGB          As COLORRGB
Public NewHSV          As COLORHSV
Public NewWRGB         As COLORRGB
Public NewBRGB         As COLORRGB

Public Whi  As Integer
Public Hue  As Integer
Public Blk  As Integer
Public Min  As Single
Public Max  As Single

Public Function HUEtoRGB(ByVal H As Integer) As COLORRGB
       
    On Error Resume Next
    
    HUEtoRGB.R = IIf(H < 128, 85 - H, H - 170)
    HUEtoRGB.G = IIf(H < 85, H, 170 - H)
    HUEtoRGB.B = IIf(H < 170, H - 85, 255 - H)
    HUEtoRGB = ColorScale(HUEtoRGB, 6)

End Function

Public Function Div(ByVal R1 As Single, ByVal R2 As Single) As Single
    
    If R2 = 0 Then R2 = ApproachVal
    Div = R1 / R2

End Function

Public Sub FindMinMax()
    
    On Error Resume Next
    
    With NewRGB
        Max = IIf(.R > .G, .R, .G)
        If .B > Max Then Max = .B
        Min = IIf(.R < .G, .R, .G)
        If .B < Min Then Min = .B
    End With

End Sub

Public Function WhiToRGB(W As Integer) As COLORRGB
  
    Dim Val As Single

    On Error Resume Next
    
    Val = W * sng1Div255
    WhiToRGB.R = NewWRGB.R + ((Resolution - NewWRGB.R) * Val)
    WhiToRGB.G = NewWRGB.G + ((Resolution - NewWRGB.G) * Val)
    WhiToRGB.B = NewWRGB.B + ((Resolution - NewWRGB.B) * Val)

End Function

Public Function BlktoRGB(B As Integer) As COLORRGB

    Dim RatioR  As Single
    Dim RatioG  As Single
    Dim RatioB  As Single
    
    On Error Resume Next
    
    RatioR = Div(NewBRGB.R - Min, Resolution)
    RatioG = Div(NewBRGB.G - Min, Resolution)
    RatioB = Div(NewBRGB.B - Min, Resolution)
    BlktoRGB.R = ColorLimits(CInt(Min + (255 - B) * RatioR))
    BlktoRGB.G = ColorLimits(CInt(Min + (255 - B) * RatioG))
    BlktoRGB.B = ColorLimits(CInt(Min + (255 - B) * RatioB))

End Function

Public Sub RGBtoWRGB()

    Dim RatioR  As Single
    Dim RatioG  As Single
    Dim RatioB  As Single
    
    On Error Resume Next
    
    If NewRGB.R = 255 And NewRGB.R = 255 And NewRGB.B = 255 Then
        NewWRGB.R = 0
        NewWRGB.G = 0
        NewWRGB.B = 0
    Else
        RatioR = Div(Resolution - NewRGB.R, Resolution - Min)
        RatioG = Div(Resolution - NewRGB.G, Resolution - Min)
        RatioB = Div(Resolution - NewRGB.B, Resolution - Min)
        NewWRGB.R = ColorLimits(CInt(NewRGB.R - Min * RatioR))
        NewWRGB.G = ColorLimits(CInt(NewRGB.G - Min * RatioG))
        NewWRGB.B = ColorLimits(CInt(NewRGB.B - Min * RatioB))
    End If

End Sub

Public Sub RGBtoBRGB()

    Dim RatioR  As Single
    Dim RatioG  As Single
    Dim RatioB  As Single
    
    On Error Resume Next

    RatioR = Div(NewRGB.R - Min, Resolution - Blk)
    RatioG = Div(NewRGB.G - Min, Resolution - Blk)
    RatioB = Div(NewRGB.B - Min, Resolution - Blk)
    
    NewBRGB.R = ColorLimits(CInt(Min + Resolution * RatioR))
    NewBRGB.G = ColorLimits(CInt(Min + Resolution * RatioG))
    NewBRGB.B = ColorLimits(CInt(Min + Resolution * RatioB))

End Sub

Public Sub RGBtoHSV()
    
    On Error Resume Next

    With NewRGB
        Select Case Max
            Case .R: NewHSV.H = CInt(Div((.G - .B) * 42.5, Max - Min))
            Case .G: NewHSV.H = CInt(Div((.B - .R) * 42.5, Max - Min)) + 85
            Case .B: NewHSV.H = CInt(Div((.R - .G) * 42.5, Max - Min)) + 170
        End Select
        If NewHSV.H < 0 Then NewHSV.H = NewHSV.H + Resolution
        NewHSV.S = Resolution - Div(Min * Resolution, Max)
        NewHSV.V = Max
    End With

End Sub

Public Sub HBtoHSV()
        
    On Error Resume Next
    
    NewHSV.H = Hue
    NewHSV.S = Resolution - Div(Min * Resolution, Max)
    NewHSV.V = Max
    
End Sub

Public Function HSVtoRGB(HueSatVal As COLORHSV) As COLORRGB

    Dim iRGB    As COLORRGB
    Dim Val     As Single
    
    On Error Resume Next
    
    iRGB = HUEtoRGB(HueSatVal.H)
    Min = (Resolution - HueSatVal.S) * HueSatVal.V * sng1Div255
    Max = HueSatVal.V
    Val = (Max - Min) * sng1Div255

    HSVtoRGB.R = Min + iRGB.R * Val
    HSVtoRGB.G = Min + iRGB.G * Val
    HSVtoRGB.B = Min + iRGB.B * Val

End Function

Public Sub FindBlk()

    On Error Resume Next
    
    Blk = Resolution - Div(Resolution * (Max - Min), Resolution - Min)

End Sub
