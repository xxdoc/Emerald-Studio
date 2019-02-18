VERSION 5.00
Begin VB.Form MainWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emerald Studio"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   15000
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   576
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1000
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame TipFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDA1A&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   8280
      Width           =   15015
   End
   Begin VB.Frame proframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Width           =   3615
      Begin Emerald_Studio.EButton proclose 
         Height          =   255
         Left            =   3120
         TabIndex        =   57
         Top             =   240
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
         Content         =   "∨"
         Align           =   0
      End
      Begin VB.PictureBox proframe_h 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   3255
         TabIndex        =   42
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox protext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   34
         Text            =   "test"
         ToolTipText     =   "元素的内容"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox protext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   33
         Text            =   "test"
         ToolTipText     =   "元素的名称，可以为空"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox protext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   32
         Text            =   "16"
         ToolTipText     =   "字体大小，只有在文字元素中有效"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.TextBox protext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   30
         Text            =   "0 - Regular"
         ToolTipText     =   "字体样式，只有在文字元素中有效"
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox protext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   28
         Text            =   "0"
         ToolTipText     =   "元素高度，在图形中禁用"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox protext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   26
         Text            =   "0"
         ToolTipText     =   "元素宽度，在图形中禁用"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox protext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   24
         Text            =   "0"
         ToolTipText     =   "Y坐标"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox protext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   22
         Text            =   "0"
         ToolTipText     =   "X坐标"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox objCombo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label alignflag 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "R"
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
         Index           =   2
         Left            =   2640
         TabIndex        =   41
         ToolTipText     =   "活动元素标记"
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label alignflag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "C"
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
         Left            =   1920
         TabIndex        =   40
         ToolTipText     =   "活动元素标记"
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label alignflag 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0B000&
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "微软雅黑"
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
         Left            =   1200
         TabIndex        =   39
         ToolTipText     =   "活动元素标记"
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Align"
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
         Index           =   8
         Left            =   240
         TabIndex        =   38
         Top             =   4680
         Width           =   465
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Content"
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
         Index           =   4
         Left            =   240
         TabIndex        =   37
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Index           =   7
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label activeflag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "A"
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
         Left            =   3000
         TabIndex        =   35
         ToolTipText     =   "活动元素标记"
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
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
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   4320
         Width           =   360
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Style"
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
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   3960
         Width           =   450
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   240
         X2              =   3240
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   240
         X2              =   3240
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
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
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
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
         Index           =   2
         Left            =   240
         TabIndex        =   25
         Top             =   3000
         Width           =   555
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
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
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   240
         TabIndex        =   21
         Top             =   2280
         Width           =   120
      End
      Begin VB.Label protitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "属性列表"
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
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame colorframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   11400
      TabIndex        =   5
      Top             =   0
      Width           =   3615
      Begin VB.TextBox colortext 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   43
         Text            =   "242"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox colortext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   19
         Text            =   "DEDEDE"
         ToolTipText     =   "粘贴或复制该HEX颜色"
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox colortext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   17
         Text            =   "242"
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox colortext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Text            =   "242"
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox colortext 
         Appearance      =   0  'Flat
         BackColor       =   &H00F9F9F9&
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
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   15
         Text            =   "242"
         Top             =   3240
         Width           =   2055
      End
      Begin VB.PictureBox colorpad 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DEDEDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   2
         Left            =   360
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   11
         ToolTipText     =   "各种颜色调整板"
         Top             =   2640
         Width           =   2895
         Begin VB.Shape colorpoint 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   2
            Left            =   0
            Shape           =   2  'Oval
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.PictureBox colorpad 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DEDEDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   360
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   10
         ToolTipText     =   "颜色透明度调整"
         Top             =   2880
         Width           =   2895
         Begin VB.Shape colorpoint 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   1320
            Shape           =   2  'Oval
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.PictureBox colorpad 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DEDEDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   0
         Left            =   360
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   9
         ToolTipText     =   "颜色调整面板"
         Top             =   720
         Width           =   1695
         Begin VB.Shape colorpoint 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   840
            Shape           =   2  'Oval
            Top             =   720
            Width           =   135
         End
      End
      Begin Emerald_Studio.EButton colorclose 
         Height          =   255
         Left            =   3120
         TabIndex        =   59
         Top             =   240
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
         Content         =   "∨"
         Align           =   0
      End
      Begin VB.Label colormem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   2400
         TabIndex        =   44
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label colormem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   2640
         TabIndex        =   45
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label ctext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Code   #"
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
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   4440
         Width           =   780
      End
      Begin VB.Label ctext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
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
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   3960
         Width           =   390
      End
      Begin VB.Label ctext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
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
         Left            =   360
         TabIndex        =   13
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label ctext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
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
         Left            =   360
         TabIndex        =   12
         Top             =   3240
         Width           =   345
      End
      Begin VB.Label colortitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "调色板"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame expframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3615
      Begin Emerald_Studio.EButton expclose 
         Height          =   255
         Left            =   3120
         TabIndex        =   58
         Top             =   240
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
         Content         =   "∨"
         Align           =   0
      End
      Begin VB.Label exptitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "资源管理器"
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame ToolFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   11400
      TabIndex        =   0
      Top             =   4800
      Width           =   3615
      Begin Emerald_Studio.EButton toolitems 
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         DefaultColor    =   16382457
         HoverColor      =   15592941
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
         Content         =   "图形元素"
         Align           =   1
      End
      Begin Emerald_Studio.EButton toolitems 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   55
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         DefaultColor    =   16382457
         HoverColor      =   15592941
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
         Content         =   "文字元素"
         Align           =   1
      End
      Begin Emerald_Studio.EButton toolitems 
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   56
         Top             =   1560
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         DefaultColor    =   16382457
         HoverColor      =   15592941
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
         Content         =   "形状元素"
         Align           =   1
      End
      Begin Emerald_Studio.EButton toolclose 
         Height          =   255
         Left            =   3120
         TabIndex        =   60
         Top             =   240
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
         Content         =   "∨"
         Align           =   0
      End
      Begin VB.Label tooltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "工具箱"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame designf 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   3600
      TabIndex        =   46
      Top             =   0
      Width           =   7815
      Begin VB.HScrollBar dsgbar0 
         Height          =   255
         Left            =   0
         Max             =   100
         TabIndex        =   48
         Top             =   8040
         Width           =   7815
      End
      Begin VB.VScrollBar dsgbar1 
         Height          =   8055
         Left            =   7560
         Max             =   100
         TabIndex        =   47
         Top             =   0
         Width           =   255
      End
      Begin VB.Frame dsgframe 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   720
         TabIndex        =   49
         Top             =   720
         Width           =   6375
         Begin VB.PictureBox us 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00F2F2F2&
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   0
            Left            =   1680
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   53
            Top             =   5160
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   705
            Left            =   15
            TabIndex        =   50
            Top             =   15
            Width           =   6345
            Begin VB.Label dsgtitle 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "设计器"
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
               Left            =   240
               TabIndex        =   51
               Top             =   240
               Width           =   585
            End
            Begin VB.Label dsgback 
               Appearance      =   0  'Flat
               BackColor       =   &H00F9F9F9&
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
               Height          =   735
               Left            =   0
               TabIndex        =   52
               Top             =   0
               Width           =   6390
            End
         End
         Begin VB.Shape dsgf 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00DEDEDE&
            Height          =   6375
            Left            =   0
            Top             =   0
            Width           =   6375
         End
         Begin VB.Shape prepareFrame 
            BackColor       =   &H00F8FAD8&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00CEDA1A&
            Height          =   2775
            Left            =   1680
            Top             =   1680
            Visible         =   0   'False
            Width           =   3735
         End
      End
   End
   Begin VB.Menu 文件 
      Caption         =   "文件(&F)"
      Begin VB.Menu opencmd 
         Caption         =   "打开(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu newcmd 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu splitline0 
         Caption         =   "-"
      End
      Begin VB.Menu savecmd 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu splitline1 
         Caption         =   "-"
      End
      Begin VB.Menu closecmd 
         Caption         =   "关闭(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu editcmd 
      Caption         =   "编辑(&E)"
      Begin VB.Menu copycmd 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu pastecmd 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu splitline2 
         Caption         =   "-"
      End
      Begin VB.Menu delcmd 
         Caption         =   "移除(&D)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu viewcmd 
      Caption         =   "视图(&V)"
      Begin VB.Menu toolboxcmd 
         Caption         =   "工具箱"
      End
      Begin VB.Menu paintcmd 
         Caption         =   "调色板"
      End
      Begin VB.Menu exporcmd 
         Caption         =   "资源管理器"
      End
      Begin VB.Menu designcmd 
         Caption         =   "设计器"
      End
      Begin VB.Menu procmd 
         Caption         =   "属性表"
      End
   End
   Begin VB.Menu runcmd 
      Caption         =   "生成(&R)"
      Begin VB.Menu openvbcmd 
         Caption         =   "打开 Visual Studio 6.0 工程"
      End
      Begin VB.Menu excutecmd 
         Caption         =   "生成可执行文件"
      End
      Begin VB.Menu packcmd 
         Caption         =   "一键打包安装程序"
      End
   End
   Begin VB.Menu toolcmd 
      Caption         =   "工具(&T)"
      Begin VB.Menu setcmd 
         Caption         =   "设置"
      End
      Begin VB.Menu colorgetcmd 
         Caption         =   "配色采集器"
      End
      Begin VB.Menu designiconcmd 
         Caption         =   "图形设计器"
      End
      Begin VB.Menu prosetcmd 
         Caption         =   "工程设置"
      End
   End
   Begin VB.Menu helpcmd 
      Caption         =   "帮助(&H)"
      Begin VB.Menu helpdoccmd 
         Caption         =   "帮助文档"
      End
      Begin VB.Menu splitline4 
         Caption         =   "-"
      End
      Begin VB.Menu aboutcmd 
         Caption         =   "关于我"
      End
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pad_a As Long, pad_r As Long, pad_g As Long, pad_b As Long, keyc As Boolean

Private Sub aboutcmd_Click()
    AboutWindow.Show
End Sub

Private Sub closecmd_Click()
    Unload Me
End Sub

Private Sub colorpad_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dim c(3) As Byte
        Select Case Index
            Case 0
                If X < 0 Then X = 0
                If X > colorpad(0).ScaleWidth - 1 Then X = colorpad(0).ScaleWidth - 1
                If Y < 0 Then Y = 0
                If Y > colorpad(0).ScaleHeight - 1 Then Y = colorpad(0).ScaleHeight - 1
                CopyMemory c(0), colorpad(0).Point(X, Y), 4
                pad_r = c(0): pad_g = c(1): pad_b = c(2)
                Call setCPadC
                Call UpdatePadPP
                colorpoint(0).Move X - 4.5, Y - 4.5
            Case 1
                If X < 0 Then X = 0
                If X > colorpad(1).ScaleWidth Then X = colorpad(1).ScaleWidth
                pad_a = X / colorpad(1).ScaleWidth * 255
                Call setCPadC
                colorpoint(1).Move X - 4.5
            Case 2
                If X < 0 Then X = 0
                If X > colorpad(2).ScaleWidth - 1 Then X = colorpad(2).ScaleWidth - 1
                CopyMemory c(0), colorpad(2).Point(X, 0), 4
                pad_r = c(0): pad_g = c(1): pad_b = c(2)
                Call setCPadC
                Call setCPad
                Call UpdatePadPP
                colorpoint(2).Move X - 4.5
        End Select
    End If
End Sub

Private Sub colorpad_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call colorpad_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub colortext_Change(Index As Integer)
    If keyc = False Then Exit Sub
    
    keyc = False
    If Index = 3 Then
        Dim c(3) As Byte, l As Long
        l = CLng("&H" & colortext(3).text)
        CopyMemory c(0), l, 4
        pad_r = c(2): pad_g = c(1): pad_b = c(0)
        Call setCPadC
        Call setCPad
        Call UpdatePadPP
        Exit Sub
    End If

    If Val(colortext(Index).text) > 255 Then colortext(Index).text = Right(colortext(Index).text, 2)
    If Val(colortext(Index).text) < 0 Then colortext(Index).text = 0
    
    pad_r = Val(colortext(0).text): pad_g = Val(colortext(1).text): pad_b = Val(colortext(2).text)
    pad_a = Val(colortext(4).text)
    Call setCPadC
    Call setCPad
    Call UpdatePadPP
End Sub

Private Sub colortext_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    keyc = True
End Sub

Private Sub dsgbar0_Change()
    dsgframe.Left = (designf.Width * 15 / 2 - dsgframe.Width / 2) - _
                            dsgbar0.value / dsgbar0.max * designf.Width * 15 * 2 + designf.Width * 15
End Sub
Private Sub dsgbar0_Scroll()
    Call dsgbar0_Change
End Sub
Private Sub dsgbar1_Change()
    dsgframe.top = (designf.Height * 15 / 2 - dsgframe.Height / 2) - _
                            dsgbar1.value / dsgbar1.max * designf.Height * 15 * 2 + designf.Height * 15
End Sub
Private Sub dsgbar1_Scroll()
    Call dsgbar1_Change
End Sub
Private Sub Form_Load()
    proframe_h.Visible = True
    
    '调色板初始化
    pad_a = 255: pad_r = 255: pad_g = 0: pad_b = 0
    Call setCPadC
    Call setCPad
    Call UpdatePadPP
    
    '设计窗口初始化
    dsgbar0.value = 50: dsgbar1.value = 50
End Sub
Public Sub UpdatePadPP()
    colorpoint(0).Move colorpad(0).ScaleWidth / 2 - 4.5, colorpad(0).ScaleHeight / 2 - 4.5
    colorpoint(1).Move pad_a / 255 * (colorpad(1).ScaleWidth - 4.5), 0
End Sub
Public Sub setCPadC()
    If colortext(0).text <> pad_r Then colortext(0).text = pad_r
    If colortext(1).text <> pad_g Then colortext(1).text = pad_g
    If colortext(2).text <> pad_b Then colortext(2).text = pad_b
    Dim strc As String
    strc = Hex(RGB(pad_b, pad_g, pad_r))
    For i = 1 To 6 - Len(strc)
        strc = "0" & strc
    Next
    If colortext(3).text <> strc Then colortext(3).text = strc
    
    If colortext(4).text <> pad_a Then colortext(4).text = pad_a
    colormem(0).BackColor = RGB(pad_r, pad_g, pad_b)
End Sub
Public Sub setCPad()
    Dim r As Long, g As Long, b As Long
    r = pad_r: g = pad_g: b = pad_b
    
    Dim gr As Long, br As Long, p As Long
    Dim c() As Long, best As Long
    GdipCreateFromHDC colorpad(0).hdc, gr

    GdipCreatePath FillModeWinding, p
    GdipAddPathRectangle p, 0, 0, colorpad(0).ScaleWidth, colorpad(0).ScaleHeight

    ReDim c(3)
    c(0) = argb(255, 255, 255, 255): c(2) = argb(255, 0, 0, 0): c(3) = argb(255, 0, 0, 0)
    '饱和颜色
    best = IIf(r > g, r, g): best = IIf(best < b, b, best)
    If best = r Then c(1) = argb(255, 255, g, b)
    If best = g Then c(1) = argb(255, r, 255, b)
    If best = b Then c(1) = argb(255, r, g, 255)
    
    GdipCreatePathGradientFromPath p, br
    GdipSetPathGradientSurroundColorsWithCount br, c(0), UBound(c) + 1
    GdipSetPathGradientCenterColor br, argb(255, r, g, b)
    
    GdipFillPath gr, br, p
    
    GdipDeleteGraphics gr
    GdipDeleteBrush br
    GdipDeletePath p
    
    colorpad(0).Refresh
    
    Dim im As Long
    GdipCreateBitmapFromFile StrPtr(App.path & "\assets\alpha.png"), im
    
    GdipCreateFromHDC colorpad(1).hdc, gr
    GdipCreateLineBrush NewPointF(0, 0), NewPointF(colorpad(1).ScaleWidth, 0), argb(0, 255, 255, 255), argb(255, r, g, b), WrapModeTile, br
    GdipDrawImage gr, im, 0, 0
    GdipFillRectangle gr, br, 0, 0, colorpad(1).ScaleWidth, colorpad(1).ScaleHeight
    
    colorpad(1).Refresh
    
    GdipDeleteGraphics gr
    GdipDeleteBrush br
    GdipDisposeImage im
    
    GdipCreateFromHDC colorpad(2).hdc, gr
    GdipGraphicsClear gr, 0
    
    Dim a() As Single, i As Integer
    
    ReDim c(5), a(5)
    c(0) = argb(255, 255, 0, 0): c(1) = argb(255, 255, 255, 0): c(2) = argb(255, 0, 255, 0)
    c(3) = argb(255, 0, 255, 255): c(4) = argb(255, 0, 0, 255): c(5) = argb(255, 255, 0, 255)
    
    For i = 0 To UBound(a)
        a(i) = Int(i / 5 * 100) / 100
    Next
    
    GdipCreateLineBrush NewPointF(0, 0), NewPointF(colorpad(2).ScaleWidth - 1, 0), 0, 0, WrapModeTileFlipXY, br
    GdipSetLinePresetBlend br, c(0), a(0), UBound(c) + 1
    
    GdipFillRectangle gr, br, 0, 0, colorpad(2).ScaleWidth, colorpad(2).ScaleHeight
    
    GdipDeleteGraphics gr
    GdipDeleteBrush br
    
    colorpad(2).Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndEmerald
    On Error Resume Next
    Unload ProjectWindow
    Unload AboutWindow
    Unload CreateWindow
    Unload SetWindow
End Sub

Private Sub setcmd_Click()
    SetWindow.Show
End Sub
