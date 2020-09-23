VERSION 5.00
Object = "*\AdcButton.vbp"
Begin VB.Form fTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dcButton Tester"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Test.frx":0000
      Left            =   180
      List            =   "Test.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Button D"
      Height          =   1290
      Left            =   165
      TabIndex        =   6
      Top             =   3900
      Width           =   1515
      Begin VB.CheckBox Check3 
         Caption         =   "Default"
         Height          =   210
         Left            =   165
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   900
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check Box"
         Height          =   210
         Left            =   165
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enabled"
         Height          =   210
         Left            =   165
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   300
         Width           =   1095
      End
   End
   Begin Dacara_dcButton.dcButton dcButton1 
      Height          =   600
      Left            =   165
      TabIndex        =   1
      Top             =   690
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1058
      BackColor       =   12230304
      Caption         =   "BUTTON A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      MaskColor       =   16777215
      PicNormal       =   "Test.frx":0004
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin Dacara_dcButton.dcButton dcButton2 
      Height          =   600
      Left            =   165
      TabIndex        =   2
      Top             =   1485
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1058
      BackColor       =   12230304
      Caption         =   "BUTTON B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      MaskColor       =   16777215
      MouseIcon       =   "Test.frx":031E
      PicNormal       =   "Test.frx":033A
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin Dacara_dcButton.dcButton dcButton3 
      Height          =   600
      Left            =   165
      TabIndex        =   3
      Top             =   2280
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1058
      BackColor       =   12230304
      Caption         =   "BUTTON C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      MaskColor       =   16777215
      MouseIcon       =   "Test.frx":0654
      PicNormal       =   "Test.frx":0670
      PicSizeH        =   32
      PicSizeW        =   32
   End
   Begin Dacara_dcButton.dcButton dcButton4 
      Default         =   -1  'True
      Height          =   600
      Left            =   165
      TabIndex        =   10
      Top             =   3075
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1058
      BackColor       =   12230304
      Caption         =   "BUTTON D"
      CheckBox        =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      MaskColor       =   16777215
      MouseIcon       =   "Test.frx":098A
      PicNormal       =   "Test.frx":09A6
      PicSizeH        =   32
      PicSizeW        =   32
      State           =   3
   End
   Begin VB.TextBox Text1 
      Height          =   5010
      Left            =   1815
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   180
      Width           =   6535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Test.frx":0CC0
      ForeColor       =   &H0000FF00&
      Height          =   705
      Left            =   150
      TabIndex        =   5
      Top             =   5355
      Width           =   8205
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    dcButton4.Enabled = Check1.Value = vbChecked
End Sub

Private Sub Check2_Click()
    dcButton4.CheckBoxMode = Check2.Value = vbChecked
End Sub

Private Sub Check3_Click()
    dcButton4.Default = Check3.Value = vbChecked
End Sub

Private Sub Combo1_Click()
    Dim dc As Control
    For Each dc In Controls
        If (TypeOf dc Is dcButton) Then
            dc.ButtonStyle = Combo1.ListIndex
            dc.ColorScheme
            If (dc.ButtonStyle = ebsMac) Then
                ' Just a small demonstration about overriding
                ' predefined colors of a specific button style
                dc.OverrideColor eucHoverColor, &H3940EA, True
                dc.OverrideColor eucDownColor, &H4FAB6F, True
            End If
        End If
    Next
End Sub

Private Sub dcButton2_Click()
    MsgBox "Button B should now be in normal state with no focus" & vbCrLf & vbTab & vbTab & "- - - - - - - -" & vbCrLf & vbCrLf & "Move the cursor on a button the press the Spacebar:" & vbCrLf & vbCrLf & vbTab & "The button should be in hot state.", , ""
End Sub

Private Sub dcButton3_Click()
    fDummy.Show vbModeless, Me
End Sub

Private Sub Form_Load()
    Dim f As Integer
        f = FreeFile()
        
    Open App.Path & "\notes.txt" For Binary Access Read Lock Write As f
    Text1 = Input(LOF(f), f)
    Close f
    
    With Combo1
        .AddItem "Crystal Button"
        .AddItem "Mac Theme"
        .AddItem "Mac OSx"
        .AddItem "Office 2003"
        .AddItem "Office XP"
        .AddItem "Opera Browser"
        .AddItem "Standard"
        .AddItem "XP Blue"
        .AddItem "XP Olive Green"
        .AddItem "XP Silver"
        .AddItem "XP Toolbar"
        .AddItem "Yahoo Style"
    End With
    
    Combo1.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fDemo.Show
End Sub
