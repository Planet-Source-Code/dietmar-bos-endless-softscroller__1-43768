VERSION 5.00
Begin VB.MDIForm frmMDIScroll 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   -210
   ClientWidth     =   9735
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1965
      Left            =   0
      ScaleHeight     =   1905
      ScaleWidth      =   9675
      TabIndex        =   3
      Top             =   0
      Width           =   9735
      Begin VB.Label Label1 
         Caption         =   $"frmMDIScroll.frx":0000
         Height          =   1365
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   4065
      End
   End
   Begin VB.PictureBox ScrollPicture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   2715
      Width           =   9735
      Begin VB.PictureBox ScrollPicture2 
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   12375
         TabIndex        =   1
         Top             =   -30
         Width           =   12435
         Begin VB.Label ScrollLabel 
            BackColor       =   &H00000000&
            Caption         =   "Click ME!"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   -30
            TabIndex        =   2
            Top             =   0
            Width           =   12345
         End
      End
      Begin VB.Timer ScrollTimer 
         Left            =   1080
         Top             =   0
      End
   End
End
Attribute VB_Name = "frmMDIScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bScroller As Boolean
Private iCharCounter As Integer
Private iTickCounter As Integer
Private cScrollTxt As String
Private Const TICK_WIDTH = 8
Private Const SCROLLDISTANCE = 15
Private Const TIMERDELAY = 30

Private Sub MDIForm_Load()
    Me.Show
End Sub

Private Sub ScrollLabel_Click()
    
    If bScroller Then
        
        bScroller = False
        ScrollTimer.Enabled = False
        ScrollTimer.Interval = TIMERDELAY

    Else
        
        cScrollTxt = "                                                                                         " & _
            "TestScroller .................... 123456789....... more useless Text here. ...... 1234567890 abcde" & _
            "fghijklmnopqrstuvwxyz ............................. and still lots of string to go................" & _
            "..................................mary had a little lamb .................. its more difficult to " & _
            "find something meaningful than to write the code for this one ....................... Ok, enough " & _
            "now, should be sufficient to demonstrate the usage of this thing.  ....over>:"
        
        iTickCounter = 0
        iCharCounter = 0
        ScrollPicture2.Left = 1
        bScroller = True
        ScrollTimer.Enabled = True
        ScrollTimer.Interval = TIMERDELAY

    End If

End Sub

Private Sub scrollTimer_Timer()
    
    ScrollPicture2.Left = ScrollPicture2.Left - SCROLLDISTANCE
    iTickCounter = iTickCounter + 1
    
    If iTickCounter >= TICK_WIDTH Then
        
        iCharCounter = iCharCounter + 1
        If iCharCounter > Len(cScrollTxt) - 80 Then iCharCounter = 1
        iTickCounter = 0
        ScrollLabel.Caption = Mid(cScrollTxt, iCharCounter)
        ScrollPicture2.Left = 1
    
    End If

End Sub


Private Sub MDIForm_Activate()
    
    MDIForm_Resize

End Sub


Private Sub MDIForm_Resize()
    
    If Me.WindowState = vbNormal Then

        If Me.Top <= 0 Then Me.Top = 1
        If Me.Top >= Screen.Height Then Me.Top = 1
        If Me.Left <= 0 Then Me.Left = 1
        If Me.Left >= Screen.Width Then Me.Left = 1
        If Me.Width > Screen.Width Then Me.Width = Screen.Width
        If Me.Height > Screen.Height Then Me.Height = Screen.Height
    
        Me.ScrollPicture2.Width = Me.Width
        Me.ScrollLabel.Width = Me.Width + 10000
        
    End If

End Sub

