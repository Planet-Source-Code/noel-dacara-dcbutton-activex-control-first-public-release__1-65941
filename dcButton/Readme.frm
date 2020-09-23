VERSION 5.00
Begin VB.Form fReadme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Something to read..."
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   255
      Top             =   270
   End
   Begin VB.TextBox Text1 
      Height          =   5640
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Readme.frx":0000
      Top             =   120
      Width           =   6765
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK... I understand. Now, show me what you got!"
      Default         =   -1  'True
      Height          =   450
      Left            =   105
      TabIndex        =   0
      Top             =   5880
      Width           =   6765
   End
End
Attribute VB_Name = "fReadme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me               '
End Sub                     '
                            '
Private Sub Form_Load()     '
    Command1.Enabled = 0    '
    Timer1.Interval = 1     ' Start timer
End Sub                     '
                            '
Private Sub Form_Unload(Cancel As Integer)
    fDemo.Show              '
End Sub                     '
                            '
Private Sub Timer1_Timer()  '
    Load fDemo              ' Load demo form
                            ' The next instruction is not executed until loading is done
    Command1.Enabled = True ' Enable the button once the form has done loading
    Command1.SetFocus       ' Set the button to focus
    Timer1.Enabled = False  ' Kill timer
End Sub                     '

