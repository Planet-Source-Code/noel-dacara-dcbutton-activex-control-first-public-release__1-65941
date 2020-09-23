VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Sample"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   5280
   Begin pSample.uSample uSample2 
      Height          =   420
      Left            =   1725
      TabIndex        =   7
      Top             =   2130
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   741
   End
   Begin pSample.uSample uSample1 
      Height          =   450
      Left            =   180
      TabIndex        =   6
      Top             =   1410
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   794
   End
   Begin pSample.uSample ucSample 
      Height          =   1350
      Left            =   120
      TabIndex        =   3
      Top             =   1305
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2381
   End
   Begin VB.Label Label1 
      Caption         =   $"fMain.frx":0000
      Height          =   2145
      Left            =   2820
      TabIndex        =   5
      Top             =   135
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2790
      Width           =   4980
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Leave"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4020
      TabIndex        =   2
      Top             =   2370
      Width           =   1065
   End
   Begin VB.Label lblEnter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2835
      TabIndex        =   1
      Top             =   2370
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Self-subclassing UserControl sample
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0 Released to PSC................................................................... 20060301
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit

Private Sub Form_Load()
  lbl.Caption = "The UserControl uses subclassing to detect mouse entry and exit. Also, the control subclasses the parent form's size and move messages."
End Sub

Private Sub Form_Resize()
  lblStatus.Move 0, Me.ScaleHeight - lblStatus.Height, Me.ScaleWidth
End Sub

Private Sub ucSample_MouseEnter()
  Me.lblEnter.BackColor = RGB(0, 255, 0)
  Me.lblExit.BackColor = &H8000000F
  lblStatus.Caption = " Mouse enter"
End Sub

Private Sub ucSample_MouseLeave()
  Me.lblEnter.BackColor = &H8000000F
  Me.lblExit.BackColor = RGB(0, 255, 0)
  lblStatus.Caption = " Mouse leave"
End Sub

Private Sub ucSample_Status(ByVal sStatus As String)
  lblStatus.Caption = " " & sStatus
End Sub
