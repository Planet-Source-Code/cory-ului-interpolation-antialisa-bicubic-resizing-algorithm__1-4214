VERSION 5.00
Begin VB.Form frmAntialisa 
   Caption         =   "Interpolation Antialisa Bicubic Resizing Algorithm"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEdge 
      Caption         =   "NoEdge"
      Height          =   372
      Left            =   3960
      TabIndex        =   7
      Top             =   4080
      Width           =   1092
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "sDrawImage"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.PictureBox pctImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2280
      Picture         =   "frmAntialisa.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Text            =   "256"
      Top             =   3348
      Width           =   735
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Text            =   "256"
      Top             =   3708
      Width           =   735
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      Height          =   192
      Left            =   3960
      TabIndex        =   5
      Top             =   3720
      Width           =   504
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      Height          =   192
      Left            =   3960
      TabIndex        =   3
      Top             =   3360
      Width           =   444
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   192
      Left            =   4560
      TabIndex        =   1
      Top             =   2160
      Width           =   84
   End
End
Attribute VB_Name = "frmAntialisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'** Interpolation Antialisa Bicubic Resizing Algorithm **'
'** Code was writen by Cory Watt(mouak@crosswinds.net)
'** Use as you wish, just never sell, unless compiled in
'** a excuting application/program!
'** Alot of thanx goes to my friend Kim Doo-hyun, Thanx **'
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Sub sDrawImage(SrcHDC As Long, OffsetX As Integer, OffsetY As Integer, srcW As Integer, srcH As Integer, dstW1 As Integer, dstH1 As Integer, dOffsetX As Integer, dOffsetY As Integer, DstHDC As Long, DstEdge As Byte)
Dim dx As Integer, dy As Integer, iX As Integer, iY As Integer, x As Integer, y As Integer
Dim i11 As Long, i12 As Long, i21 As Long, i22 As Long
Dim V1 As Integer, V2 As Integer, V3 As Integer, S1 As Integer, S2 As Integer, S3 As Integer, N1 As Integer, N2 As Integer, N3 As Integer, H1 As Integer, H2 As Integer, H3 As Integer, U1 As Integer, U2 As Integer, U3 As Integer, P1 As Integer, P2 As Integer, P3 As Integer
Dim Color11qRed As Integer, Color11qGreen As Integer, Color11qBlue As Integer, _
Color21qRed As Integer, Color21qGreen As Integer, Color21qBlue As Integer, _
Color22qRed As Integer, Color22qGreen As Integer, Color22qBlue As Integer, _
Color12qRed As Integer, Color12qGreen As Integer, Color12qBlue As Integer
Dim dstW As Integer, dstH As Integer
Dim iRX As Integer, iOrX As Integer, iRY As Integer, iOrY As Integer, dw As Integer, dh As Integer
If DstEdge = 1 Then
    dstW = dstW1 + (dstW1 / srcW)
    dstH = dstH1 + (dstH1 / srcH)
Else
    dstW = dstW1
    dstH = dstH1
End If
For dy = 0 To srcH - 1
    iOrY = iRY
    iRY = ((dstH) / srcH) * (dy + 1)
    For dx = 0 To srcW - 1
        iOrX = iRX
        iRX = ((dstW) / srcW) * (dx + 1)
        
        '(Getting 4 Colors.  Of X, upper-left,
        'upper-right, lower-left, lower-right.)
        i11 = GetPixel(SrcHDC, dx + OffsetX, dy + OffsetY)
        i12 = GetPixel(SrcHDC, dx + 1 + OffsetX, dy + OffsetY)
        i21 = GetPixel(SrcHDC, dx + OffsetX, dy + 1 + OffsetY)
        i22 = GetPixel(SrcHDC, dx + 1 + OffsetX, dy + 1 + OffsetY)

        
        iX = iOrX
        iY = iOrY
        dw = iRX - iOrX
        dh = iRY - iOrY
       
        '(Get the Three Color values, Red, Green,
        'and blue.)
        
        '(upper-left)
        Color11qRed = i11 Mod 256
        Color11qGreen = (i11 \ 256) Mod 256
        Color11qBlue = (i11 \ 65536) Mod 256
      
        '(lower-left)
        Color12qRed = i12 Mod 256
        Color12qGreen = (i12 \ 256) Mod 256
        Color12qBlue = (i12 \ 65536) Mod 256
        
        '(upper-right)
        Color21qRed = i21 Mod 256
        Color21qGreen = (i21 \ 256) Mod 256
        Color21qBlue = (i21 \ 65536) Mod 256
       
        '(lower-right)
        Color22qRed = i22 Mod 256
        Color22qGreen = (i22 \ 256) Mod 256
        Color22qBlue = (i22 \ 65536) Mod 256
        
        '(Red)
        N1 = Color21qRed - Color11qRed
        H1 = Color11qRed
        
        
        '(Green)
        N2 = Color21qGreen - Color11qGreen
        H2 = Color11qGreen
        
        '(Blue)
        N3 = Color21qBlue - Color11qBlue
        H3 = Color11qBlue
            
        '(Cubic!)
        '(Red)
        U1 = Color22qRed - Color12qRed
        P1 = Color12qRed
        
        '(Green)
        U2 = Color22qGreen - Color12qGreen
        P2 = Color12qGreen
        
        '(Blue)
        U3 = Color22qBlue - Color12qBlue
        P3 = Color12qBlue
        
        For y = 0 To dh - 1
            '(Now begins the Interpolation)
            Color11qRed = H1 + ((N1) / dh) * y
            Color11qGreen = H2 + ((N2) / dh) * y
            Color11qBlue = H3 + ((N3) / dh) * y
            
            Color12qRed = P1 + ((U1) / dh) * y
            Color12qGreen = P2 + ((U2) / dh) * y
            Color12qBlue = P3 + ((U3) / dh) * y
            
            
            '(Red)
            V1 = Color12qRed - Color11qRed
            S1 = Color11qRed
            
            '(Green)
            V2 = Color12qGreen - Color11qGreen
            S2 = Color11qGreen
            
            '(Blue)
            V3 = Color12qBlue - Color11qBlue
            S3 = Color11qBlue
            
            
            For x = 0 To dw - 1
                Color11qRed = S1 + ((V1) / dw) * x
                Color11qGreen = S2 + ((V2) / dw) * x
                Color11qBlue = S3 + ((V3) / dw) * x
                
                '(Set a Pixel, may need some changing,
                If DstEdge = 1 Then
                    If x + iX < dstW1 And y + iY < dstH1 Then
                        SetPixel DstHDC, x + iX + dOffsetX, y + iY + dOffsetY, RGB(Color11qRed, Color11qGreen, Color11qBlue)
                    End If
                Else
                    SetPixel DstHDC, x + iX + dOffsetX, y + iY + dOffsetY, RGB(Color11qRed, Color11qGreen, Color11qBlue)
                End If
            Next x
        Next y
    If dx = srcW - 1 Then iRX = 0
    Next dx

    
    '(not need)
    Label1.Caption = dy
    DoEvents
    
    If dy = srcH - 1 Then iRY = 0
Next dy
End Sub

Private Sub cmdDraw_Click()
On Error Resume Next
    If txtWidth < 32 Or txtHeight < 32 Then
        MsgBox "Enlarges an Image only (at the moment!)", vbExclamation, "Enlarging!"
    Else
        sDrawImage pctImage.hdc, 0, 0, 32, 32, txtWidth, txtHeight, 0, 0, Me.hdc, chkEdge.Value
    End If
End Sub

