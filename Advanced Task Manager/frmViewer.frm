VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewer 
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Advanced Task Manager 1.0 - www.odesayazilim.com"
   ClientHeight    =   4455
   ClientLeft      =   2745
   ClientTop       =   3555
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   9960
   Begin Project1.Button Button4 
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ButtonStyleColors=   4
      ButtonTheme     =   3
      BackColor       =   11903133
      BackColorPressed=   15525862
      BackColorHover  =   16250356
      BorderColor     =   -2147483627
      BorderColorPressed=   -2147483628
      BorderColorHover=   -2147483627
      Caption         =   "About Program"
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
   Begin Project1.Button Button3 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   3720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ButtonStyle     =   0
      ButtonStyleColors=   4
      ButtonTheme     =   3
      BackColor       =   11903133
      BackColorPressed=   15525862
      BackColorHover  =   16250356
      BorderColor     =   11903133
      BorderColorPressed=   11903133
      BorderColorHover=   11903133
      EffectColor     =   16776960
      Caption         =   "Exit"
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
   Begin Project1.Button Button2 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   3720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      ButtonStyle     =   0
      ButtonStyleColors=   4
      ButtonTheme     =   3
      BackColor       =   11903133
      BackColorPressed=   15525862
      BackColorHover  =   16250356
      BorderColor     =   11903133
      BorderColorPressed=   11903133
      BorderColorHover=   11903133
      EffectColor     =   16776960
      Caption         =   "Kill Working File(s)"
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
   Begin Project1.Button Button1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ButtonStyle     =   0
      ButtonStyleColors=   4
      ButtonTheme     =   3
      BackColor       =   11903133
      BackColorPressed=   15525862
      BackColorHover  =   16250356
      BorderColor     =   11903133
      BorderColorPressed=   11903133
      BorderColorHover=   11903133
      EffectColor     =   16776960
      Caption         =   "Show File(s) Info (dll & ocx)"
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
   Begin MSComctlLib.StatusBar STB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   450
      SimpleText      =   "Odesa System"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6376
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PID"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Path"
         Object.Width           =   5645
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Mem Usage"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Threads"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Creation Time"
         Object.Width           =   2083
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Creation Date"
         Object.Width           =   2134
      EndProperty
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Advanced Task Manager
'By Alper ESKIKILIC
'odesayazilim@gmail.com
'www.odesayazilim.com

Private Sub Button1_Click()
If ListView1.SelectedItem.SubItems(2) <> "System Directory" Then
        frmModules.Show
    End If
End Sub

Private Sub Button2_Click()
Dim lnghProcess As Long

    lnghProcess = OpenProcess(1&, -1&, ListView1.SelectedItem.SubItems(1))
    Call TerminateProcess(lnghProcess, 0&)
        ListView1.ListItems.Clear
    Call ProcessLoad
    STB.Panels(1).Text = "Working Files: " & ListView1.ListItems.Count
End Sub

Private Sub Button3_Click()
 End
End Sub

Private Sub Button4_Click()
FormAbout.Show
End Sub

Private Sub Form_Load()
    Call ProcessLoad
    STB.Panels(1).Text = "Working Files: " & ListView1.ListItems.Count
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu mnuFile
End Sub

Private Sub mnuClose_Click()
Dim sProcess As Long
    sProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, ListView1.SelectedItem.SubItems(1))
    Call CloseHandle(sProcess)
End Sub

Private Sub mnuExit_Click()
   
End Sub

Private Sub mnuKill_Click()

End Sub
