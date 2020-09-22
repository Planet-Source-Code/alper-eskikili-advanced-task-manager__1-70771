VERSION 5.00
Begin VB.Form FormAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ::.... About ....::"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.Button Button1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   873
      ButtonStyle     =   3
      ButtonStyleColors=   4
      ButtonTheme     =   5
      CaptionEffect   =   1
      CaptionStyle    =   1
      BackColor       =   181749
      BackColorPressed=   11530238
      BackColorHover  =   14875135
      BorderColor     =   181749
      BorderColorPressed=   -2147483628
      BorderColorHover=   -2147483627
      Caption         =   "Odesa Advanced Task Manager"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000F&
      X1              =   -240
      X2              =   9240
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Image Image3 
      Height          =   1080
      Left            =   5400
      Picture         =   "Form1.frx":0000
      Top             =   3600
      Width           =   2250
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   120
      Picture         =   "Form1.frx":397F
      Top             =   3960
      Width           =   420
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "Programmed And  By : Alper ESKÝKILIÇ E-Mail: odesayazilim@gmail.com "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Odesa Yazýlým Advanced Task Manager"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000F&
      X1              =   4440
      X2              =   4440
      Y1              =   3480
      Y2              =   5640
   End
   Begin VB.Image Image1 
      Height          =   2970
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":4369
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Advanced Task Manager
'By Alper ESKIKILIC
'odesayazilim@gmail.com
'www.odesayazilim.com
Private Sub Button1_Click()

End Sub
