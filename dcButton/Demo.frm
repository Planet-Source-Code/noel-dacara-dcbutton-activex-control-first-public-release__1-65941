VERSION 5.00
Object = "*\AdcButton.vbp"
Begin VB.Form fDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dcButton Demo - *** Should be more faster when compiled ***"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin Dacara_dcButton.dcButton dcXPToolbar 
      Height          =   690
      Index           =   2
      Left            =   5535
      TabIndex        =   48
      Top             =   5925
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1217
      ButtonStyle     =   10
      Caption         =   "UNLOAD ME"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin Dacara_dcButton.dcButton dcXPToolbar 
      Height          =   750
      Index           =   0
      Left            =   5535
      TabIndex        =   46
      Top             =   4140
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1323
      ButtonStyle     =   10
      Caption         =   "BUG FIXES"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin VB.Frame fraCrystal 
      Caption         =   "Crystal Style"
      Height          =   840
      Left            =   3825
      TabIndex        =   63
      Top             =   4050
      Width           =   1575
      Begin Dacara_dcButton.dcButton dcCrystal 
         Height          =   375
         Index           =   0
         Left            =   135
         TabIndex        =   34
         Top             =   300
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         BackColor       =   65280
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcCrystal 
         Height          =   375
         Index           =   2
         Left            =   1020
         TabIndex        =   36
         Top             =   300
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         BackColor       =   255
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcCrystal 
         Height          =   375
         Index           =   1
         Left            =   570
         TabIndex        =   35
         Top             =   300
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         BackColor       =   33023
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
   End
   Begin VB.Frame fraPopupMenu 
      Caption         =   "Popup Menu"
      Height          =   1635
      Left            =   3825
      TabIndex        =   62
      Top             =   4980
      Width           =   1575
      Begin Dacara_dcButton.dcButton dcPopup 
         Height          =   405
         Index           =   0
         Left            =   180
         TabIndex        =   37
         Top             =   285
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BackColor       =   12230304
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":0000
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcPopup 
         Height          =   405
         Index           =   1
         Left            =   585
         TabIndex        =   38
         Top             =   285
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BackColor       =   16751432
         ButtonStyle     =   1
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":00E2
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcPopup 
         Height          =   405
         Index           =   2
         Left            =   990
         TabIndex        =   39
         Top             =   285
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BackColor       =   10591645
         ButtonStyle     =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":01C4
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcPopup 
         Height          =   405
         Index           =   3
         Left            =   180
         TabIndex        =   40
         Top             =   690
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":02A6
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcPopup 
         Height          =   405
         Index           =   4
         Left            =   585
         TabIndex        =   41
         Top             =   690
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         ButtonStyle     =   4
         Caption         =   ""
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":0388
         PicSizeH        =   16
         PicSizeW        =   16
         State           =   3
      End
      Begin Dacara_dcButton.dcButton dcPopup 
         Height          =   405
         Index           =   5
         Left            =   990
         TabIndex        =   42
         Top             =   690
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BackColor       =   13815503
         ButtonStyle     =   5
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":046A
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcPopup 
         Height          =   405
         Index           =   6
         Left            =   180
         TabIndex        =   43
         Top             =   1095
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BackColor       =   14995922
         ButtonStyle     =   9
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":054C
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcPopup 
         Height          =   405
         Index           =   7
         Left            =   585
         TabIndex        =   44
         Top             =   1095
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         ButtonStyle     =   10
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":062E
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcPopup 
         Height          =   405
         Index           =   8
         Left            =   990
         TabIndex        =   45
         Top             =   1095
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
         BackColor       =   1228031
         ButtonStyle     =   11
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":0710
         PicSizeH        =   16
         PicSizeW        =   16
      End
   End
   Begin VB.Frame fraMoreFeatures 
      Caption         =   "More Features"
      Height          =   3045
      Left            =   150
      TabIndex        =   61
      Top             =   4050
      Width           =   3495
      Begin Dacara_dcButton.dcButton dcAlignSample 
         Height          =   885
         Left            =   180
         TabIndex        =   31
         Top             =   1365
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1561
         BackColor       =   15133676
         ButtonStyle     =   7
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         PicNormal       =   "Demo.frx":07F2
         PicSizeH        =   32
         PicSizeW        =   32
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   2595
         Left            =   1605
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Demo.frx":0AD8
         Top             =   270
         Width           =   1785
      End
      Begin Dacara_dcButton.dcButton dcSpecialEffects 
         Height          =   405
         Left            =   180
         TabIndex        =   32
         Top             =   2385
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   714
         BackColor       =   14995922
         ButtonStyle     =   9
         Caption         =   "Special Effects"
         Effects         =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcShape 
         Height          =   390
         Index           =   0
         Left            =   165
         TabIndex        =   25
         Top             =   315
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         BackColor       =   12230304
         ButtonShape     =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcShape 
         Height          =   390
         Index           =   1
         Left            =   570
         TabIndex        =   26
         Top             =   315
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   688
         BackColor       =   12230304
         ButtonShape     =   3
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcShape 
         Height          =   390
         Index           =   2
         Left            =   1005
         TabIndex        =   27
         Top             =   315
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         BackColor       =   12230304
         ButtonShape     =   1
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcShape 
         Height          =   390
         Index           =   3
         Left            =   165
         TabIndex        =   28
         Top             =   840
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         BackColor       =   10591645
         ButtonShape     =   2
         ButtonStyle     =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcShape 
         Height          =   390
         Index           =   4
         Left            =   570
         TabIndex        =   29
         Top             =   840
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   688
         BackColor       =   10591645
         ButtonShape     =   3
         ButtonStyle     =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcShape 
         Height          =   390
         Index           =   5
         Left            =   1005
         TabIndex        =   30
         Top             =   840
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         BackColor       =   10591645
         ButtonShape     =   1
         ButtonStyle     =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicOpacity      =   0
      End
   End
   Begin VB.Frame fraWhatsNew 
      Caption         =   "What's New?"
      Height          =   2565
      Left            =   6390
      TabIndex        =   60
      Top             =   4050
      Width           =   2535
      Begin VB.TextBox txtWhatsNew 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1395
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "Demo.frx":0C01
         Top             =   945
         Width           =   2085
      End
      Begin Dacara_dcButton.dcButton dcWhatsNew 
         Height          =   465
         Left            =   210
         TabIndex        =   49
         Top             =   330
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   820
         BackColor       =   10591645
         ButtonStyle     =   2
         Caption         =   "LET ME TEST IT"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HandPointer     =   -1  'True
         PicOpacity      =   0
      End
   End
   Begin VB.Frame fraOpacity 
      Caption         =   "Picture Opacity"
      Height          =   2115
      Left            =   150
      TabIndex        =   59
      Top             =   1830
      Width           =   1605
      Begin VB.HScrollBar hsbOpacity 
         Height          =   315
         LargeChange     =   10
         Left            =   180
         Max             =   100
         Min             =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1575
         Value           =   55
         Width           =   1230
      End
      Begin Dacara_dcButton.dcButton dcOpacity 
         Height          =   1095
         Left            =   210
         TabIndex        =   11
         Top             =   330
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1931
         BackColor       =   10591645
         ButtonStyle     =   2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         PicAlign        =   0
         PicNormal       =   "Demo.frx":0CBF
         PicOpacity      =   0.55
         PicSizeH        =   71
         PicSizeW        =   68
      End
   End
   Begin VB.Frame fraStandard 
      Caption         =   "Standard Style Button"
      Height          =   1620
      Left            =   4830
      TabIndex        =   58
      Top             =   120
      Width           =   2205
      Begin Dacara_dcButton.dcButton dcStandard 
         Height          =   450
         Index           =   0
         Left            =   210
         TabIndex        =   8
         Top             =   345
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   794
         ButtonStyle     =   6
         Caption         =   "Command Button"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton dcStandard 
         Height          =   450
         Index           =   1
         Left            =   210
         TabIndex        =   9
         Top             =   930
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   794
         ButtonStyle     =   6
         Caption         =   "Checkbox"
         CheckBox        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraOfficeXP 
      Caption         =   "Office XP Style"
      Height          =   780
      Left            =   150
      TabIndex        =   56
      Top             =   960
      Width           =   1600
      Begin Dacara_dcButton.dcButton dcOfficeXP 
         Height          =   405
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   255
         Width           =   450
         _ExtentX        =   397
         _ExtentY        =   397
         ButtonStyle     =   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":1767
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcOfficeXP 
         Height          =   405
         Index           =   1
         Left            =   570
         TabIndex        =   4
         Top             =   255
         Width           =   450
         _ExtentX        =   397
         _ExtentY        =   397
         ButtonStyle     =   4
         Caption         =   ""
         CheckBox        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":18C1
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcOfficeXP 
         Height          =   405
         Index           =   2
         Left            =   1035
         TabIndex        =   5
         Top             =   255
         Width           =   450
         _ExtentX        =   397
         _ExtentY        =   397
         ButtonStyle     =   4
         Caption         =   ""
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":1A1B
         PicSizeH        =   16
         PicSizeW        =   16
         State           =   3
      End
   End
   Begin VB.Frame fraOffice2k3 
      Caption         =   "Office 2003 Style"
      Height          =   780
      Left            =   150
      TabIndex        =   55
      Top             =   120
      Width           =   1600
      Begin Dacara_dcButton.dcButton dcOffice2k3 
         Height          =   405
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   255
         Width           =   450
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   0
         PicNormal       =   "Demo.frx":1B75
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcOffice2k3 
         Height          =   405
         Index           =   1
         Left            =   570
         TabIndex        =   1
         Top             =   255
         Width           =   450
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   ""
         CheckBox        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":1CCF
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin Dacara_dcButton.dcButton dcOffice2k3 
         Height          =   405
         Index           =   2
         Left            =   1035
         TabIndex        =   2
         Top             =   255
         Width           =   450
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   12230304
         ButtonStyle     =   3
         Caption         =   ""
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicNormal       =   "Demo.frx":1E29
         PicSizeH        =   16
         PicSizeW        =   16
         State           =   3
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   7200
      ScaleHeight     =   99
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   111
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   195
      Width           =   1725
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   105
         Top             =   105
      End
      Begin VB.Timer Timer1 
         Interval        =   20
         Left            =   105
         Top             =   105
      End
   End
   Begin VB.Frame fraOpera 
      Caption         =   "Opera Browser Style Button"
      Height          =   2115
      Left            =   6390
      TabIndex        =   54
      Top             =   1830
      Width           =   2535
      Begin Dacara_dcButton.dcButton dcOpera 
         Height          =   420
         Index           =   0
         Left            =   210
         TabIndex        =   22
         Top             =   345
         Width           =   2100
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   13815503
         ButtonStyle     =   5
         Caption         =   "Opera Browser"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton dcOpera 
         Height          =   420
         Index           =   1
         Left            =   210
         TabIndex        =   23
         Top             =   900
         Width           =   2100
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   13815503
         ButtonStyle     =   5
         Caption         =   "Checkbox"
         CheckBox        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Dacara_dcButton.dcButton dcOpera 
         Height          =   420
         Index           =   2
         Left            =   210
         TabIndex        =   24
         Top             =   1455
         Width           =   2100
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   13815503
         ButtonStyle     =   5
         Caption         =   "Disabled"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         State           =   3
      End
   End
   Begin VB.Frame fraXP 
      Caption         =   "XP Style Buttons"
      Height          =   2115
      Left            =   1920
      TabIndex        =   52
      Top             =   1830
      Width           =   4305
      Begin Dacara_dcButton.dcButton dcXPBlue 
         Height          =   420
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   345
         Width           =   1215
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   15133676
         ButtonStyle     =   7
         Caption         =   "Blue"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcXPOliveGreen 
         Height          =   420
         Index           =   0
         Left            =   210
         TabIndex        =   16
         Top             =   900
         Width           =   1215
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   14742518
         ButtonStyle     =   8
         Caption         =   "Olive Green"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin Dacara_dcButton.dcButton dcXPSilver 
         Height          =   420
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   1455
         Width           =   1215
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   14995922
         ButtonStyle     =   9
         Caption         =   "Silver"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin Dacara_dcButton.dcButton dcXPBlue 
         Height          =   420
         Index           =   1
         Left            =   1545
         TabIndex        =   14
         Top             =   345
         Width           =   1215
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   15133676
         ButtonStyle     =   7
         Caption         =   "Checkbox"
         CheckBox        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin Dacara_dcButton.dcButton dcXPOliveGreen 
         Height          =   420
         Index           =   1
         Left            =   1545
         TabIndex        =   17
         Top             =   915
         Width           =   1215
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   14742518
         ButtonStyle     =   8
         Caption         =   "Checkbox"
         CheckBox        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin Dacara_dcButton.dcButton dcXPSilver 
         Height          =   420
         Index           =   1
         Left            =   1545
         TabIndex        =   20
         Top             =   1455
         Width           =   1215
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   14995922
         ButtonStyle     =   9
         Caption         =   "Checkbox"
         CheckBox        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin Dacara_dcButton.dcButton dcXPBlue 
         Height          =   420
         Index           =   2
         Left            =   2865
         TabIndex        =   15
         Top             =   345
         Width           =   1215
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   15133676
         ButtonStyle     =   7
         Caption         =   "Disabled"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         State           =   3
      End
      Begin Dacara_dcButton.dcButton dcXPOliveGreen 
         Height          =   420
         Index           =   2
         Left            =   2865
         TabIndex        =   18
         Top             =   900
         Width           =   1215
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   14742518
         ButtonStyle     =   8
         Caption         =   "Disabled"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         State           =   3
      End
      Begin Dacara_dcButton.dcButton dcXPSilver 
         Height          =   420
         Index           =   2
         Left            =   2865
         TabIndex        =   21
         Top             =   1455
         Width           =   1215
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   14995922
         ButtonStyle     =   9
         Caption         =   "Disabled"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         State           =   3
      End
   End
   Begin VB.Frame fraYahoo 
      Caption         =   "Yahoo Style Button"
      Height          =   1620
      Left            =   1920
      TabIndex        =   51
      Top             =   120
      Width           =   2760
      Begin Dacara_dcButton.dcButton dcYahoo 
         Height          =   450
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Top             =   345
         Width           =   2325
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   16744576
         ButtonStyle     =   11
         Caption         =   "Download Now!"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         PicOpacity      =   0
      End
      Begin Dacara_dcButton.dcButton dcYahoo 
         Height          =   450
         Index           =   1
         Left            =   210
         TabIndex        =   7
         Top             =   930
         Width           =   2325
         _ExtentX        =   397
         _ExtentY        =   397
         BackColor       =   1228031
         ButtonStyle     =   11
         Caption         =   "Sign up for Yahoo!"
         CheckBox        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Dacara_dcButton.dcButton dcXPToolbar 
      Height          =   690
      Index           =   1
      Left            =   5535
      TabIndex        =   47
      Top             =   5070
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1217
      ButtonStyle     =   10
      Caption         =   "UNFIXED BUG"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Made in the Philippines"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3810
      TabIndex        =   53
      Top             =   6705
      Width           =   5115
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Made in the Philippines"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   3780
      TabIndex        =   57
      Top             =   6675
      Width           =   5115
   End
   Begin VB.Menu TestMenu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu TestMenu1 
         Caption         =   "Button should have stayed in down state"
      End
      Begin VB.Menu TestMenu2 
         Caption         =   "As you click again, the menu should close"
      End
      Begin VB.Menu TestMenu3 
         Caption         =   "Then the button returns to normal or hot"
      End
   End
   Begin VB.Menu AlignMenu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu PicAlign 
         Caption         =   "Behind Text"
         Index           =   0
      End
      Begin VB.Menu PicAlign 
         Caption         =   "Bottom Edge"
         Index           =   1
      End
      Begin VB.Menu PicAlign 
         Caption         =   "Bottom of Caption"
         Index           =   2
      End
      Begin VB.Menu PicAlign 
         Caption         =   "Left Edge"
         Index           =   3
      End
      Begin VB.Menu PicAlign 
         Caption         =   "Left of Caption"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu PicAlign 
         Caption         =   "Right Edge"
         Index           =   5
      End
      Begin VB.Menu PicAlign 
         Caption         =   "Right of Caption"
         Index           =   6
      End
      Begin VB.Menu PicAlign 
         Caption         =   "Top Edge"
         Index           =   7
      End
      Begin VB.Menu PicAlign 
         Caption         =   "Top of Caption"
         Index           =   8
      End
   End
End
Attribute VB_Name = "fDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Zoom desktop window display
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
    Private Type POINTAPI
        x As Long
        Y As Long
    End Type
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Private Const SRCCOPY As Long = &HCC0020

' Open bug fixes file
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub dcOpera_Click(Index As Integer)
    Select Case Index
        Case 0
            ShellExecute hwnd, "open", "http://www.opera.com/", "", "", vbNormalFocus
    End Select
    
End Sub

Private Sub dcWhatsNew_Click()
    fTest.Show vbModeless, Me
    Me.Hide ' Only hide this form after finish loading the fTest form
    
End Sub

Private Sub dcXPBlue_Click(Index As Integer)
    Select Case Index
        Case 0: dcXPBlue(0).Enabled = False
        Case 1: dcXPBlue(2).Enabled = dcXPBlue(1).Value
        Case 2: dcXPBlue(0).Enabled = True
    End Select
    
End Sub

Private Sub dcXPOliveGreen_Click(Index As Integer)
    Select Case Index
        Case 0: dcXPOliveGreen(0).Enabled = False
        Case 1: dcXPOliveGreen(2).Enabled = dcXPOliveGreen(1).Value
        Case 2: dcXPOliveGreen(0).Enabled = True
    End Select
    
End Sub

Private Sub dcXPSilver_Click(Index As Integer)
    Select Case Index
        Case 0: dcXPSilver(0).Enabled = False
        Case 1: dcXPSilver(2).Enabled = dcXPSilver(1).Value
        Case 2: dcXPSilver(0).Enabled = True
    End Select
    
End Sub

Private Sub dcXPToolbar_Click(Index As Integer)
    Select Case Index
        Case 0
            ShellExecute hwnd, "open", "Fixed Bugs.htm", "", App.Path, vbNormalFocus
        Case 1
            ShellExecute hwnd, "open", "Unfixed Bug.htm", "", App.Path, vbNormalFocus
        Case 2
            Unload Me
    End Select
    
End Sub

Private Sub dcYahoo_Click(Index As Integer)
    Select Case Index
        Case 1
            ShellExecute hwnd, "open", "http://www.yahoomail.com/", "", "", vbNormalFocus
    End Select
    
End Sub

Private Sub Form_Load()
    Tag = 0
    Call Picture1_Click
    
    ' Change default color used by the control
    dcStandard(0).OverrideColor eucGrayText, vbButtonShadow
    dcStandard(1).OverrideColor eucGrayText, vbGrayText
    
    ' Setup popup menu control
    dcPopup(0).SetPopupMenu TestMenu, emaTopLeft
    dcPopup(1).SetPopupMenu TestMenu, emaTop
    dcPopup(2).SetPopupMenu TestMenu, emaTopRight
    dcPopup(3).SetPopupMenu TestMenu, emaLeft
    dcPopup(5).SetPopupMenu TestMenu, emaRight
    dcPopup(6).SetPopupMenu TestMenu, emaLeftBottom
    dcPopup(7).SetPopupMenu TestMenu, emaBottom
    dcPopup(8).SetPopupMenu TestMenu, emaRightBottom
    
    dcAlignSample.SetPopupMenu AlignMenu
    
End Sub

Private Sub Form_Resize()
    If (Me.WindowState = vbMinimized) Then
        Tag = Timer1
        Timer1 = False
    Else
        Timer1 = Tag
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Timer2.Enabled = False
    
End Sub

Private Sub hsbOpacity_Change()
    dcOpacity.PictureOpacity = hsbOpacity.Value
    
End Sub

Private Sub PicAlign_Click(Index As Integer)
    If (Not Index = dcAlignSample.PictureAlignment) Then
        PicAlign(dcAlignSample.PictureAlignment).Checked = False
        PicAlign(Index).Checked = True
        dcAlignSample.PictureAlignment = Index
    End If
End Sub

Private Sub Picture1_Click()
    Timer1 = Not Timer1
    Timer2 = Not Timer1
    
End Sub

Private Sub Timer1_Timer()
    Static lpPoint As POINTAPI
    
    If (GetCursorPos(lpPoint)) Then
        Call ShowZoom(lpPoint) ' Show zoom from cursor position
    End If
    
End Sub

Private Sub ShowZoom(lpPoint As POINTAPI)
    Dim hDC As Long
    Dim hwnd As Long
    Dim nTmp As Long
    
    Dim x As Long
    Dim Y As Long
    Dim nWidth As Long
    Dim nHeight As Long
    
    Dim xSrc As Long
    Dim ySrc As Long
    Dim nSrcWidth As Long
    Dim nSrcHeight As Long
    
    Const nZoomLevel As Long = 5 ' 1 = 100%, 2 = 200%, 3 = 300% and so on...
    
    hwnd = GetDesktopWindow()
    hDC = GetDC(hwnd)
    
    With Picture1
        .Cls
        x = 0
        Y = 0
        nWidth = .ScaleWidth
        nHeight = .ScaleHeight
    End With
    
    xSrc = (nWidth / 2) / nZoomLevel
    xSrc = lpPoint.x - xSrc
    ySrc = (nHeight / 2) / nZoomLevel
    ySrc = lpPoint.Y - ySrc
    nSrcWidth = nWidth / nZoomLevel
    nSrcHeight = nHeight / nZoomLevel
    
    nTmp = nSrcWidth * nZoomLevel
    
    If (nTmp > nWidth) Then
        nWidth = nTmp
    ElseIf (nTmp < nWidth) Then
        nSrcWidth = nSrcWidth + 1
        nWidth = nTmp + nZoomLevel
    End If
    
    nTmp = nSrcHeight * nZoomLevel
    
    If (nTmp > nHeight) Then
        nHeight = nTmp
    ElseIf (nTmp < nHeight) Then
        nSrcHeight = nSrcHeight + 1
        nHeight = nTmp + nZoomLevel
    End If
    
    Call StretchBlt(Picture1.hDC, _
                    x, _
                    Y, _
                    nWidth, _
                    nHeight, _
                    hDC, _
                    xSrc, _
                    ySrc, _
                    nSrcWidth, _
                    nSrcHeight, _
                    SRCCOPY)
    
    Call ReleaseDC(hwnd, hDC)
    
End Sub

Private Sub Timer2_Timer()
    Picture1.Cls
    If (Len(Picture1.Tag) = 0) Then
        Picture1.CurrentX = 30
        Picture1.CurrentY = 35
        Picture1.Print "CLICK TO"
        Picture1.CurrentX = 15
        Picture1.CurrentY = Picture1.CurrentY + 3
        Picture1.Print "TOGGLE ZOOM"
        Picture1.Tag = 1
    Else
        Picture1.Tag = ""
    End If
    
End Sub
