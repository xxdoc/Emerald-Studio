VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "ImageX.ocx"
Begin VB.Form ProjectWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emerald Studio"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9360
   Icon            =   "ProjectWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   StartUpPosition =   2  '屏幕中心
   Begin Emerald_Studio.EButton toolbutton 
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "创建一个新的Emerald工程"
      Align           =   1
   End
   Begin Emerald_Studio.EButton toolbutton 
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   3240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "打开旧的项目"
      Align           =   1
   End
   Begin ImageX.aicAlphaImage Icons 
      Height          =   360
      Index           =   1
      Left            =   360
      Top             =   3240
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Image           =   "ProjectWindow.frx":1BCC2
      Props           =   5
   End
   Begin ImageX.aicAlphaImage Icons 
      Height          =   360
      Index           =   0
      Left            =   360
      Top             =   2640
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      Image           =   "ProjectWindow.frx":1C1A7
      Props           =   5
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   24
      X2              =   592
      Y1              =   280
      Y2              =   280
   End
   Begin VB.Label recentttile 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "近期编辑"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   4320
      Width           =   795
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "让我们开始"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   4185
      TabIndex        =   0
      Top             =   1800
      Width           =   1005
   End
   Begin ImageX.aicAlphaImage LOGO 
      Height          =   960
      Left            =   4200
      Top             =   720
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Image           =   "ProjectWindow.frx":1C6F8
      Props           =   5
   End
End
Attribute VB_Name = "ProjectWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Unload MainWindow
End Sub

Private Sub toolbutton_Click(Index As Integer)
    Select Case Index
        Case 0
            CreateWindow.Show: Me.Hide
    End Select
End Sub
