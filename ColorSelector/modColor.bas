Attribute VB_Name = "modColor"
Option Explicit

Public Type COLORRGB
    R               As Integer
    G               As Integer
    B               As Integer
End Type

Public Type COLORHSV
    H               As Integer
    S               As Integer
    V               As Integer
End Type

Public Function ColorSet(iR As Integer, iG As Integer, iB As Integer) As COLORRGB

    ColorSet.R = iR
    ColorSet.G = iG
    ColorSet.B = iB

End Function

Public Function ColorLimits(ByVal iColor As Integer) As Integer

    ColorLimits = iColor
    If ColorLimits < 0 Then ColorLimits = 0
    If ColorLimits > 255 Then ColorLimits = 255
    
End Function

Public Function ColorScale(iC As COLORRGB, S As Single) As COLORRGB

    ColorScale.R = ColorLimits(iC.R * S)
    ColorScale.G = ColorLimits(iC.G * S)
    ColorScale.B = ColorLimits(iC.B * S)

End Function

Public Function ColorLongToRGB(lColor As Long) As COLORRGB

    ColorLongToRGB.R = (lColor And &HFF&)
    ColorLongToRGB.G = (lColor And &HFF00&) / &H100&
    ColorLongToRGB.B = (lColor And &HFF0000) / &H10000

End Function

