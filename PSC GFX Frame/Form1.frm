VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   13050
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   8100
      ScaleHeight     =   1545
      ScaleWidth      =   4125
      TabIndex        =   13
      Top             =   435
      Width           =   4125
      Begin Project1.GFXFrameX GFXFrameX12 
         Height          =   330
         Left            =   225
         TabIndex        =   14
         Top             =   495
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         ShowShadow      =   0   'False
         FillColour      =   16777215
         BackColour      =   16777215
         BackColourIs    =   1
         BorderColour    =   11495265
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabWidthOff     =   4
      End
   End
   Begin Project1.GFXFrameX GFXFrameX8 
      Height          =   3630
      Left            =   4950
      TabIndex        =   7
      Top             =   3675
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   6403
      FrameStyle      =   0
      FillColour      =   14584719
      BorderColour    =   8214097
      Caption         =   "Multi Layer Controls"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColour      =   4210752
      TabWidthOff     =   8
      Begin Project1.GFXFrameX GFXFrameX11 
         Height          =   930
         Left            =   315
         TabIndex        =   12
         Top             =   2325
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   1640
         FillColour      =   8214097
         BackColour      =   14584719
         BackColourIs    =   1
         Caption         =   "Tab On For Different Fill Colours"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         TabWidthOff     =   6
      End
      Begin Project1.GFXFrameX GFXFrameX10 
         Height          =   420
         Left            =   300
         TabIndex        =   9
         Top             =   1500
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   741
         FrameStyle      =   0
         ShadowDepth     =   3
         FillColour      =   16777215
         BackColour      =   14584719
         BackColourIs    =   1
         BorderColour    =   11495265
         CornerDepth     =   8
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabWidthOff     =   4
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   60
            TabIndex        =   10
            Text            =   "Jazz Up Existing Controls (Text Box)"
            Top             =   60
            Width           =   3150
         End
      End
      Begin Project1.GFXFrameX GFXFrameX9 
         Height          =   960
         Left            =   120
         TabIndex        =   8
         Top             =   375
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   1693
         FrameStyle      =   0
         FillColour      =   14584719
         BackColour      =   14584719
         BackColourIs    =   1
         BorderColour    =   8214097
         CornerDepth     =   20
         Caption         =   "Big Corners"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowTab         =   0   'False
         FontColour      =   4210752
         TabWidthOff     =   4
      End
   End
   Begin Project1.GFXFrameX GFXFrameX7 
      Height          =   675
      Left            =   480
      TabIndex        =   6
      Top             =   6660
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   1191
      ShadowOpacity   =   30
      ShadowDepth     =   3
      BorderColour    =   49344
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabWidthOff     =   4
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No Caption"
         Height          =   345
         Left            =   210
         TabIndex        =   11
         Top             =   180
         Width           =   2460
      End
   End
   Begin Project1.GFXFrameX GFXFrameX6 
      Height          =   1275
      Left            =   540
      TabIndex        =   5
      Top             =   5220
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2249
      FrameStyle      =   0
      ShadowOpacity   =   30
      ShadowDepth     =   10
      BorderColour    =   49152
      Caption         =   "Higher Shadow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      TabWidthOff     =   10
   End
   Begin Project1.GFXFrameX GFXFrameX5 
      Height          =   1395
      Left            =   4965
      TabIndex        =   4
      Top             =   2130
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   2461
      ShowShadow      =   0   'False
      BorderColour    =   12632256
      Caption         =   "Very Basic Frame - No Tab"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowTab         =   0   'False
      FontColour      =   4210752
      TabWidthOff     =   4
   End
   Begin Project1.GFXFrameX GFXFrameX4 
      Height          =   1455
      Left            =   3975
      TabIndex        =   3
      Top             =   540
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2566
      FrameStyle      =   0
      FillColour      =   12648447
      BorderColour    =   16761024
      Caption         =   "Center - User Colours"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      FontColour      =   8421631
      TabWidthOff     =   10
   End
   Begin Project1.GFXFrameX GFXFrameX3 
      Height          =   1515
      Left            =   480
      TabIndex        =   2
      Top             =   3540
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   2672
      BorderColour    =   255
      Caption         =   "Right Align"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      FontColour      =   33023
      TabWidthOff     =   4
   End
   Begin Project1.GFXFrameX GFXFrameX2 
      Height          =   1275
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2249
      FrameStyle      =   0
      Caption         =   "Big Tab"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      TabWidthOff     =   20
   End
   Begin Project1.GFXFrameX GFXFrameX1 
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   780
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2143
      Caption         =   "Standard Frame"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabWidthOff     =   4
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

