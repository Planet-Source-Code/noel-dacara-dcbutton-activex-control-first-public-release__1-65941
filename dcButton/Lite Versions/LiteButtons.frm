VERSION 5.00
Begin VB.Form fLiteButtons_Demo 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin dcLiteButtons_Demo.dcButtonCrystal dcButtonCrystal1 
      Height          =   495
      Left            =   1845
      TabIndex        =   3
      Top             =   975
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      BackColor       =   12230304
      Caption         =   "Crystal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin dcLiteButtons_Demo.dcButtonXPSilver dcButtonXPSilver1 
      Height          =   495
      Left            =   1845
      TabIndex        =   1
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      BackColor       =   14995922
      Caption         =   "XP Silver"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin dcLiteButtons_Demo.dcButtonXPBlue dcButtonXPBlue1 
      Height          =   495
      Left            =   195
      TabIndex        =   0
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      BackColor       =   15133676
      Caption         =   "XP Blue"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin dcLiteButtons_Demo.dcButtonMacOSx dcButtonMacOSx1 
      Height          =   495
      Left            =   195
      TabIndex        =   2
      Top             =   975
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      BackColor       =   10591645
      Caption         =   "Mac OS X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin dcLiteButtons_Demo.dcButtonMac dcButtonMac1 
      Height          =   495
      Left            =   1050
      TabIndex        =   4
      Top             =   1695
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      BackColor       =   16751432
      Caption         =   "Mac"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This project ensures that there aren't any declaration conflicts between components."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   180
      TabIndex        =   5
      Top             =   2415
      Width           =   3180
   End
End
Attribute VB_Name = "fLiteButtons_Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
