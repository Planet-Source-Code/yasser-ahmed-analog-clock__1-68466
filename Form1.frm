VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analog Clock"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   DrawWidth       =   3
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   4680
      Width           =   3735
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   0
         Width           =   495
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   4
         Left            =   2030
         Shape           =   3  'Circle
         Top             =   195
         Width           =   135
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   5
         Left            =   2030
         Shape           =   3  'Circle
         Top             =   30
         Width           =   135
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   6
         Left            =   950
         Shape           =   3  'Circle
         Top             =   195
         Width           =   135
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   7
         Left            =   950
         Shape           =   3  'Circle
         Top             =   30
         Width           =   135
      End
      Begin VB.Image Image6 
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   480
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   1080
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   1560
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   2160
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   2640
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2040
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   36
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":15FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":255E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":34C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4426
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4BD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   6
      Visible         =   0   'False
      X1              =   1080
      X2              =   2880
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1080
      X2              =   2760
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Visible         =   0   'False
      X1              =   1080
      X2              =   2760
      Y1              =   3120
      Y2              =   3120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Dim S, H, HH, M, N, K, R As Integer
Dim W, He As Long
Dim A(60), B(60), AA(60), BB(60), AAA(60), BBB(60)

Private Sub Form_Activate()
    'Set the window position to topmost
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub


Private Sub Form_Load()

    Dim Ret As Long
    'Set the window style to 'Layered'
    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    'Set the opacity of the layered window to 200
    SetLayeredWindowAttributes Me.hWnd, 0, 200, LWA_ALPHA
    
   
End Sub


Private Sub Timer1_Timer()
Cls
S = Val(Mid$(Time$, 7, 2))
M = Val(Mid$(Time$, 4, 2))
HH = Val(Mid$(Time$, 1, 2))

If HH >= 12 Then
    HH = HH - 12
End If

H = (HH * 5) + (Int(M / 12))


For N = 0 To 59
    K = -(N / 30) * (22 / 7) + (22 / 7)
    AAA(N) = ((Form1.Width) / 2) + (((Form1.Width) / 3) * Sin(K))
    AA(N) = ((Form1.Width) / 2) + (((Form1.Width) / 3.5) * Sin(K))
    A(N) = ((Form1.Width) / 2) + (((Form1.Width) / 4) * Sin(K))
    BBB(N) = ((Form1.Height) / 2) + (((Form1.Height) / 3) * Cos(K))
    BB(N) = ((Form1.Height) / 2) + (((Form1.Height) / 3.5) * Cos(K))
    B(N) = ((Form1.Height) / 2) + (((Form1.Height) / 4) * Cos(K))
    PSet (AA(N), BB(N)), &H80FF&
Next N



For N = 0 To 59 Step 5
    Line (AA(N), BB(N))-(A(N), B(N)), &H80FF&
    Form1.CurrentX = AAA(N) - (150 * (Form1.Height / Form1.Width))
    Form1.CurrentY = BBB(N) - (100 * (Form1.Height / Form1.Width))
    Form1.FontBold = True
    Form1.FontSize = 12
    Print N / 5
Next N

Circle ((Form1.Width) / 2, (Form1.Height) / 2), 50

Line1.Visible = True: Line2.Visible = True: Line3.Visible = True


Picture1.Left = ((Form1.Width) / 2) - ((Picture1.Width) / 2)
Picture1.Top = ((Form1.Height) / 2) + ((((Form1.Height) / 2.8) * (Cos(-1 * (22 / 7) + (22 / 7)))))

Line1.X1 = A(S): Line1.X2 = ((Form1.Width) / 2)
Line1.Y1 = B(S): Line1.Y2 = ((Form1.Height) / 2)

Line2.X1 = A(M): Line2.X2 = ((Form1.Width) / 2)
Line2.Y1 = B(M): Line2.Y2 = ((Form1.Height) / 2)

Line3.X1 = A(H): Line3.X2 = ((Form1.Width) / 2)
Line3.Y1 = B(H): Line3.Y2 = ((Form1.Height) / 2)


Image1.Picture = ImageList1.ListImages(Val(Mid(Time, Len(Time) - 3, 1)) + 1).Picture
Image2.Picture = ImageList1.ListImages(Val(Mid(Time, Len(Time) - 4, 1)) + 1).Picture
Image3.Picture = ImageList1.ListImages(Val(Mid(Time, Len(Time) - 6, 1)) + 1).Picture
Image4.Picture = ImageList1.ListImages(Val(Mid(Time, Len(Time) - 7, 1)) + 1).Picture
Image5.Picture = ImageList1.ListImages(Val(Mid(Time, Len(Time) - 9, 1)) + 1).Picture
If (Hour(Time) >= 1 And Hour(Time) <= 9) Or (Hour(Time) >= 13 And Hour(Time) <= 21) Then
    Image6.Picture = ImageList1.ListImages(1).Picture
Else
    Image6.Picture = ImageList1.ListImages(Val(Mid(Time, Len(Time) - 10, 1)) + 1).Picture
End If

Label2.Caption = Mid(Time, Len(Time) - 1, 2)
End Sub
