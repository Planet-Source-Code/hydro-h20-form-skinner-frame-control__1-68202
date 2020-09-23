VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "This Is A Test"
   ClientHeight    =   5235
   ClientLeft      =   3720
   ClientTop       =   2040
   ClientWidth     =   6420
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin Project1.tjFormSkinner tjFormSkinner1 
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9128
      BorderColor     =   32768
      ShowCaptionArea =   -1  'True
      CaptionColorTo  =   0
      CaptionColorFrom=   8454016
      Caption         =   "This Is A Test (click and hold to move)"
      Icon            =   "Form1.frx":000C
      IconSize        =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorTo         =   16384
      ColorFrom       =   12648384
      ShowMinButton   =   -1  'True
      MinColor        =   32768
      MinColorFrom    =   32768
      MinColorTo      =   8454016
      ShowMaxButton   =   -1  'True
      MaxColor        =   32768
      MaxColorFrom    =   32768
      MaxColorTo      =   8454016
      ShowCloseButton =   -1  'True
      CloseColor      =   32768
      CloseColorFrom  =   32768
      CloseColorTo    =   8454016
      BackGroundPicture=   "Form1.frx":5AF6
      BackGroundPicturePosition=   5
      Begin Project1.tjFormSkinner tjFormSkinner3 
         Height          =   2535
         Left            =   3000
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4471
         CaptionColorTo  =   0
         CaptionColorFrom=   0
         Caption         =   "Can be used as a image holder"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorTo         =   1
         ColorFrom       =   1
         MinOverColor    =   0
         MinColorFrom    =   0
         MinColorTo      =   0
         MaxOverColor    =   0
         MaxColorFrom    =   0
         MaxColorTo      =   0
         CloseOverColor  =   0
         CloseColorFrom  =   0
         CloseColorTo    =   0
         BackGroundPicture=   "Form1.frx":63D0
         BackGroundPicturePosition=   8
      End
      Begin Project1.tjFormSkinner tjFormSkinner2 
         Height          =   1815
         Left            =   120
         Top             =   3240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3201
         BorderColor     =   255
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorTo         =   4194368
         ColorFrom       =   16761087
         BackGroundPicture=   "Form1.frx":189DD
         BackGroundPicturePosition=   5
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form1.frx":1E4C7
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   120
            TabIndex        =   1
            Top             =   120
            Width           =   3135
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":1E56F
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    tjFormSkinner1.Move 0, 0, Me.Width - 2, Me.Height - 2
End Sub

Private Sub Form_Resize()
    tjFormSkinner1.Move 0, 0, Me.Width - 2, Me.Height - 2
End Sub

Private Sub tjFormSkinner1_CloseClicked()
    Unload Me
End Sub

Private Sub tjFormSkinner1_MaxClicked()
    If Me.WindowState = 2 Then
        Me.WindowState = 0
    Else
        Me.WindowState = 2
    End If
End Sub

Private Sub tjFormSkinner1_MinClicked()
    Me.WindowState = 1
End Sub
