VERSION 5.00
Begin VB.Form SetWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emerald Studio"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9825
   Icon            =   "SetWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   655
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1560
      TabIndex        =   25
      Text            =   "120"
      Top             =   7440
      Width           =   6135
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1560
      TabIndex        =   23
      Text            =   "10"
      Top             =   6720
      Width           =   6135
   End
   Begin VB.CheckBox formatcheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".erd Emerald´æµµÎÄ¼þ"
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   20
      Top             =   5400
      Width           =   5175
   End
   Begin VB.CheckBox formatcheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".ers Emerald StudioÉèÖÃÎÄ¼þ"
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   19
      Top             =   5040
      Width           =   5175
   End
   Begin VB.CheckBox formatcheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ".erp Emerald¹¤³ÌÎÄ¼þ"
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   18
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Text            =   "Î´ÉèÖÃ"
      Top             =   4080
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Text            =   "Î´ÉèÖÃ"
      Top             =   3360
      Width           =   6135
   End
   Begin VB.TextBox pathtext 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "D:\My Doc\Emerald Projects\New Project\"
      Top             =   1680
      Width           =   6135
   End
   Begin Emerald_Studio.EButton pather 
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   1680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      DefaultColor    =   16382457
      HoverColor      =   15592941
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
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
   Begin Emerald_Studio.EButton EButton1 
      Height          =   255
      Left            =   7800
      TabIndex        =   11
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      DefaultColor    =   16382457
      HoverColor      =   15592941
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
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
   Begin Emerald_Studio.EButton EButton2 
      Height          =   255
      Left            =   7800
      TabIndex        =   14
      Top             =   4080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      DefaultColor    =   16382457
      HoverColor      =   15592941
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
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
   Begin Emerald_Studio.EButton EButton3 
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      DefaultColor    =   16382457
      HoverColor      =   15592941
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "»ñÈ¡SDK"
      Align           =   0
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "×î´ó³·Ïú²½Öè"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1560
      TabIndex        =   24
      Top             =   7080
      Width           =   1170
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "¹¤³Ì×Ô¶¯±£´æ¼ä¸ôÊ±¼ä£¨·ÖÖÓ£©"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1560
      TabIndex        =   22
      Top             =   6360
      Width           =   2730
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   88
      X2              =   88
      Y1              =   424
      Y2              =   512
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "»Ö¸´"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   480
      TabIndex        =   21
      Top             =   6840
      Width           =   390
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   88
      X2              =   88
      Y1              =   312
      Y2              =   400
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ÎÄ¼þ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Top             =   5160
      Width           =   390
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Emerald SDK ´¢´æÎ»ÖÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   3720
      Width           =   2025
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Basic 6.0 IDE Î»ÖÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Top             =   3000
      Width           =   2235
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   88
      X2              =   88
      Y1              =   200
      Y2              =   288
   End
   Begin VB.Label UIOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Icelolly UI"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   9
      Top             =   2400
      Width           =   1320
   End
   Begin VB.Label UIOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dark"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   8
      Top             =   2400
      Width           =   1320
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "µ÷ÊÔ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   480
      TabIndex        =   7
      Top             =   3480
      Width           =   390
   End
   Begin VB.Label UIOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDA1A&
      Caption         =   "Light"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   2400
      Width           =   1320
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "UIÖ÷Ìâ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ä¬ÈÏ´¢´æÎ»ÖÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   88
      X2              =   88
      Y1              =   88
      Y2              =   176
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "³£¹æ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label title 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ÉèÖÃ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      Width           =   390
   End
End
Attribute VB_Name = "SetWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
