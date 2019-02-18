VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "ImageX.ocx"
Begin VB.Form CreateWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emerald Studio"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10095
   ControlBox      =   0   'False
   Icon            =   "CreateWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   673
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox protext 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Text            =   "New Project"
      Top             =   5760
      Width           =   6135
   End
   Begin Emerald_Studio.EButton pather 
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   6120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      DefaultColor    =   16382457
      HoverColor      =   15592941
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "..."
      Align           =   0
   End
   Begin VB.TextBox pathtext 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Text            =   "D:\My Doc\Emerald Projects\New Project\"
      Top             =   6120
      Width           =   6135
   End
   Begin Emerald_Studio.EButton okbtn 
      Height          =   375
      Left            =   8400
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      DefaultColor    =   15592941
      HoverColor      =   13556250
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "创建"
      Align           =   0
   End
   Begin Emerald_Studio.EButton cancelbtn 
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      DefaultColor    =   15592941
      HoverColor      =   13556250
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "取消"
      Align           =   0
   End
   Begin VB.Label pathname 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "位置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Top             =   6120
      Width           =   390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   24
      X2              =   640
      Y1              =   376
      Y2              =   376
   End
   Begin VB.Label proname 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "工程名"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   5760
      Width           =   585
   End
   Begin ImageX.aicAlphaImage toolicons 
      Height          =   960
      Index           =   1
      Left            =   600
      Top             =   2760
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Image           =   "CreateWindow.frx":000C
      Props           =   5
   End
   Begin VB.Label tooldes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "提供基本的绘图和界面管理功能。"
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
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   3240
      Width           =   2925
   End
   Begin VB.Label tooltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用于软件项目"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0B000&
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   2880
      Width           =   1365
   End
   Begin ImageX.aicAlphaImage toolicons 
      Height          =   960
      Index           =   0
      Left            =   600
      Top             =   1320
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Image           =   "CreateWindow.frx":0F59
      Props           =   5
   End
   Begin VB.Label tooldes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "包括了存档及存档安全防护功能和音效播放功能。"
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
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   4290
   End
   Begin VB.Label tooltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用于游戏项目"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDA1A&
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Label title 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "创建一个新的Emerald工程"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDA1A&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label toolfocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFDF0&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FCFDF0&
      Height          =   1485
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   10050
   End
   Begin VB.Label toolfocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FCFDF0&
      Height          =   1485
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   10050
   End
End
Attribute VB_Name = "CreateWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancelbtn_Click()
    Me.Hide
    ProjectWindow.Show
End Sub

Private Sub okbtn_Click()
    Me.Hide
    MainWindow.Show
End Sub

Private Sub tooldes_Click(Index As Integer)
    Call toolfocus_Click(Index)
End Sub

Private Sub toolfocus_Click(Index As Integer)
    For i = 0 To toolfocus.UBound
        toolfocus(i).BackColor = RGB(255, 255, 255)
    Next
    
    Select Case Index
        Case 0
            toolfocus(Index).BackColor = &HFCFDF0
        Case 1
            toolfocus(Index).BackColor = &HFFFCF2
    End Select
End Sub

Private Sub toolicons_Click(Index As Integer, ByVal Button As Integer)
    Call toolfocus_Click(Index)
End Sub

Private Sub tooltitle_Click(Index As Integer)
    Call toolfocus_Click(Index)
End Sub
