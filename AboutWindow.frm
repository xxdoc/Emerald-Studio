VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "ImageX.ocx"
Begin VB.Form AboutWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Me"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7905
   Icon            =   "AboutWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   1  '所有者中心
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Github Pages:         https://red-error404.github.io/233"
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
      Left            =   480
      TabIndex        =   7
      Top             =   3600
      Width           =   5025
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "* 此开源项目使用MIT开源协议（包括Emerald绘图框架）"
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
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   4950
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "* 最终代码调试和编译由Microsoft Visual Basic 6.0完成"
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
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Width           =   4860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   16
      X2              =   512
      Y1              =   368
      Y2              =   368
   End
   Begin ImageX.aicAlphaImage LOGO 
      Height          =   960
      Left            =   3480
      Top             =   600
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Image           =   "AboutWindow.frx":000C
      Props           =   5
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "* 这只是一个代码生成器和设计器，不是IDE"
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
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   3780
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "邮箱:                       ris_vb@126.com"
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
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   3345
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "QQ:                        1361778219"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Indev.218"
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
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   7875
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Emerald Studio"
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
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   7875
   End
End
Attribute VB_Name = "AboutWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
