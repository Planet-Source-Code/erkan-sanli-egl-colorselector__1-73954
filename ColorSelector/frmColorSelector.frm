VERSION 5.00
Begin VB.Form frmColorDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Selector"
   ClientHeight    =   3255
   ClientLeft      =   765
   ClientTop       =   1065
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   519
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picNew 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4200
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   26
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer tmrPick 
      Left            =   240
      Top             =   3720
   End
   Begin VB.CheckBox chkPick 
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmColorSelector.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin EGL_ColorSelector.UpDown udRGBHSV 
      Height          =   300
      Index           =   5
      Left            =   6960
      TabIndex        =   25
      Top             =   2160
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
   End
   Begin EGL_ColorSelector.UpDown udRGBHSV 
      Height          =   300
      Index           =   4
      Left            =   6960
      TabIndex        =   24
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
   End
   Begin EGL_ColorSelector.UpDown udRGBHSV 
      Height          =   300
      Index           =   3
      Left            =   6960
      TabIndex        =   23
      Top             =   1440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
   End
   Begin EGL_ColorSelector.UpDown udRGBHSV 
      Height          =   300
      Index           =   2
      Left            =   6960
      TabIndex        =   22
      Top             =   1080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
   End
   Begin EGL_ColorSelector.UpDown udRGBHSV 
      Height          =   300
      Index           =   1
      Left            =   6960
      TabIndex        =   21
      Top             =   720
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
   End
   Begin EGL_ColorSelector.UpDown udRGBHSV 
      Height          =   300
      Index           =   0
      Left            =   6960
      TabIndex        =   20
      Top             =   345
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
   End
   Begin VB.PictureBox picW 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2700
      Left            =   3150
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   360
      Width           =   345
      Begin VB.Shape shpW 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         Top             =   600
         Width           =   345
      End
   End
   Begin VB.PictureBox picHB 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2700
      Left            =   360
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   360
      Width           =   2700
      Begin VB.Shape shpHB 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   1320
         Top             =   1200
         Width           =   90
      End
   End
   Begin VB.PictureBox picRGBHSV 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   5
      Left            =   4200
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2700
      Begin VB.Shape shpV 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         Top             =   0
         Width           =   60
      End
   End
   Begin VB.PictureBox picRGBHSV 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   4200
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2700
      Begin VB.Shape shpS 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         Top             =   0
         Width           =   60
      End
   End
   Begin VB.PictureBox picRGBHSV 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   4200
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2700
      Begin VB.Shape shpH 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         Top             =   0
         Width           =   60
      End
   End
   Begin VB.PictureBox picRGBHSV 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   4200
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2700
      Begin VB.Shape shpB 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         Top             =   0
         Width           =   60
      End
   End
   Begin VB.PictureBox picRGBHSV 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   4200
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   2700
      Begin VB.Shape shpG 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         Top             =   0
         Width           =   60
      End
   End
   Begin VB.PictureBox picRGBHSV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   4200
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   2700
      Begin VB.Shape shpR 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   240
         Top             =   0
         Width           =   60
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "B l a c k n e s s "
      Height          =   1935
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   135
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Whiteness"
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hue"
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Value :"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   5
      Left            =   3600
      TabIndex        =   14
      Top             =   2160
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Sat :"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   3600
      TabIndex        =   12
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Hue :"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   3600
      TabIndex        =   10
      Top             =   1440
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Blue :"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   3600
      TabIndex        =   8
      Top             =   1080
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Green :"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   3600
      TabIndex        =   6
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Red :"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   360
      Width           =   570
   End
End
Attribute VB_Name = "frmColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor& Lib "user32" (ByVal bshow As Long)

Private Const Step1  As Single = 1.448864 '255/176
Private Const Step2  As Single = 0.690196 '176/255
Private Const Step3  As Single = 0.705882 '180/255
Private Const Step4  As Single = 1.465517 '255/174
Private Const Step5  As Single = 0.682352 '174/255

Private Sub Form_Load()

    FillHue picRGBHSV(3)
    FillHB picHB
    NewRGB.R = 50
    NewRGB.G = 100
    NewRGB.B = 150
    RefreshDisplay (tRGB)
    FillDisplay
    
End Sub

Private Sub FillHue(picBox As PictureBox)
    
    Dim idx     As Integer
    Dim X       As Single
    Dim C1      As COLORRGB
      
    For idx = 0 To Resolution
        C1 = HUEtoRGB(idx)
        picBox.Line (X, 0)-(X + Step3, picBox.ScaleHeight), RGB(C1.R, C1.G, C1.B), B
        X = X + Step3
    Next
    
End Sub

Private Sub FillWhiteness(picBox As PictureBox)
    
    Dim idx     As Integer
    Dim Y       As Single
    Dim C1      As COLORRGB

    For idx = 0 To Resolution
        C1 = WhiToRGB(idx)
        picBox.Line (0, Y)-(picBox.ScaleWidth, Y + Step3), RGB(C1.R, C1.G, C1.B), B
        Y = Y + Step3
    Next
    
End Sub

Private Sub FillGradient(picBox As PictureBox, C1 As COLORRGB, C2 As COLORRGB)
    
    Dim idx     As Integer
    Dim X       As Single
    Dim R       As Single
    Dim G       As Single
    Dim B       As Single
    Dim stepR   As Single
    Dim stepG   As Single
    Dim stepB   As Single

    stepR = (C2.R - C1.R) * sng1Div255
    stepG = (C2.G - C1.G) * sng1Div255
    stepB = (C2.B - C1.B) * sng1Div255
    R = C1.R
    G = C1.G
    B = C1.B
    For idx = 0 To Resolution
        picBox.Line (X, 0)-(X + Step3, picBox.ScaleHeight), RGB(R, G, B), B
        R = R + stepR
        G = G + stepG
        B = B + stepB
        X = X + Step3
    Next

End Sub

Private Sub FillHB(picBox As PictureBox)

    Dim idx     As Integer
    Dim idy     As Integer
    Dim ihsv    As COLORHSV
    Dim C1      As COLORRGB
    Dim X       As Single
    Dim Y       As Single

    ihsv.S = 255
    For idx = 0 To Resolution
        ihsv.H = idx
        Y = 0
        For idy = 0 To Resolution
            ihsv.V = 255 - idy
            C1 = HSVtoRGB(ihsv)
            picBox.PSet (X, Y), RGB(C1.R, C1.G, C1.B)
            Y = Y + Step3
        Next
        X = X + Step3
    Next

End Sub

Public Sub FillDisplay()
    
    With NewRGB
        FillGradient picRGBHSV(0), ColorSet(0, .G, .B), ColorSet(255, .G, .B)
        FillGradient picRGBHSV(1), ColorSet(.R, 0, .B), ColorSet(.R, 255, .B)
        FillGradient picRGBHSV(2), ColorSet(.R, .G, 0), ColorSet(.R, .G, 255)
    End With
    
    Dim irgbMin As COLORRGB
    Dim irgbMax As COLORRGB
    Dim ihsvMin As COLORHSV
    Dim ihsvMax As COLORHSV

    ihsvMax = NewHSV
    ihsvMax.V = Resolution
    irgbMax = HSVtoRGB(ihsvMax)
    FillGradient picRGBHSV(5), ColorSet(0, 0, 0), irgbMax
    ihsvMax = NewHSV
    ihsvMax.S = Resolution
    irgbMax = HSVtoRGB(ihsvMax)
    ihsvMin = NewHSV
    ihsvMin.S = 0
    irgbMin = HSVtoRGB(ihsvMin)
    FillGradient picRGBHSV(4), irgbMin, irgbMax
    
End Sub

Public Sub RefreshDisplay(cType As ColorType)
    
    Select Case cType
        
        Case tRGB
            FindMinMax
            RGBtoWRGB
            FillWhiteness picW
            Whi = Min
            RGBtoHSV
            Hue = NewHSV.H
            FindBlk
            RGBtoBRGB
        
        Case tHSV
            NewRGB = HSVtoRGB(NewHSV)
            Hue = NewHSV.H
            FindMinMax
            RGBtoWRGB
            FillWhiteness picW
            Whi = Min
            FindBlk
            RGBtoBRGB
            
        Case tWhi
            NewRGB = WhiToRGB(Whi)
            FindMinMax
            RGBtoHSV
            FindBlk
            RGBtoBRGB
       
        Case tHB
            FindMinMax
            NewRGB = BlktoRGB(Blk)
            RGBtoWRGB
            FillWhiteness picW
            Whi = Min
            NewBRGB = HUEtoRGB(Hue)
            HBtoHSV
    
    End Select
    udRGBHSV(0).Value = NewRGB.R
    udRGBHSV(1).Value = NewRGB.G
    udRGBHSV(2).Value = NewRGB.B
    udRGBHSV(3).Value = NewHSV.H
    udRGBHSV(4).Value = NewHSV.S
    udRGBHSV(5).Value = NewHSV.V
    shpR.Left = Step2 * NewRGB.R
    shpG.Left = Step2 * NewRGB.G
    shpB.Left = Step2 * NewRGB.B
    shpH.Left = Step2 * NewHSV.H
    shpS.Left = Step2 * NewHSV.S
    shpV.Left = Step2 * NewHSV.V
    shpW.Top = Step2 * Whi
    shpHB.Top = Step5 * Blk
    shpHB.Left = Step5 * Hue
    picNew.BackColor = RGB(NewRGB.R, NewRGB.G, NewRGB.B)

End Sub

Private Sub GetPosition(Pos As Single, Col As Integer)
    
    Dim Val As Integer
    
    Val = Pos - 2
    If Val < 0 Then Val = 0
    If Val > 176 Then Val = 176
    Col = CInt(Val * Step1)

End Sub

Private Sub HBPosition(X As Single, Y As Single, H As Integer, B As Integer)
        
    Dim Val As Integer
    
    Val = X - 3
    If Val < 0 Then Val = 0
    If Val > 174 Then Val = 174
    H = CInt(Val * Step4)
    
    Val = Y - 3
    If Val < 0 Then Val = 0
    If Val > 174 Then Val = 174
    B = CInt(Val * Step4)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ShowCursor True
    Unload frmPick
    End

End Sub

Private Sub udRGBHSV_Change(Index As Integer)
        
    Select Case Index
        Case 0: NewRGB.R = udRGBHSV(0).Value: RefreshDisplay tRGB
        Case 1: NewRGB.G = udRGBHSV(1).Value: RefreshDisplay tRGB
        Case 2: NewRGB.B = udRGBHSV(2).Value: RefreshDisplay tRGB
        Case 3: NewHSV.H = udRGBHSV(3).Value: RefreshDisplay tHSV
        Case 4: NewHSV.S = udRGBHSV(4).Value: RefreshDisplay tHSV
        Case 5: NewHSV.V = udRGBHSV(5).Value: RefreshDisplay tHSV
    End Select

End Sub

Private Sub picRGBHSV_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    If Button = 1 Then
        Select Case Index
            Case 0: GetPosition X, NewRGB.R: RefreshDisplay tRGB
            Case 1: GetPosition X, NewRGB.G: RefreshDisplay tRGB
            Case 2: GetPosition X, NewRGB.B: RefreshDisplay tRGB
            Case 3: GetPosition X, NewHSV.H: RefreshDisplay tHSV
            Case 4: GetPosition X, NewHSV.S: RefreshDisplay tHSV
            Case 5: GetPosition X, NewHSV.V: RefreshDisplay tHSV
        End Select
    End If

End Sub

Private Sub picW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        GetPosition Y, Whi: RefreshDisplay tWhi
    End If

End Sub

Private Sub picHB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        HBPosition X, Y, Hue, Blk: RefreshDisplay tHB
    End If
    
End Sub

Private Sub picRGBHSV_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    picRGBHSV_MouseMove Index, Button, Shift, X, Y
    ShowCursor False

End Sub

Private Sub picHB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picHB_MouseMove Button, Shift, X, Y
    ShowCursor False

End Sub

Private Sub picW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    picW_MouseMove Button, Shift, X, Y
    ShowCursor False

End Sub

Private Sub picRGBHSV_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    FillDisplay
    ShowCursor True

End Sub

Private Sub udRGBHSV_MouseUp(Index As Integer)
    
    FillDisplay

End Sub

Private Sub picW_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    FillDisplay
    ShowCursor True

End Sub

Private Sub picHB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    FillDisplay
    ShowCursor True

End Sub

Private Sub cmdOK_Click()
    
    Unload Me
    
End Sub

Private Sub chkPick_Click()
    
    frmPick.Show
    RefreshDisplay tRGB
    FillDisplay
    
End Sub

