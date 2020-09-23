VERSION 5.00
Begin VB.Form fDummy 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Dummy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To test click me then click this window again."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MouseIcon       =   "Dummy.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2790
      Width           =   3180
   End
   Begin VB.Label Label3 
      Caption         =   "Once a window is activated again, the last control on focus will be restored and all controls will now respond to any events."
      Height          =   600
      Left            =   165
      TabIndex        =   2
      Top             =   2115
      Width           =   3105
   End
   Begin VB.Label Label2 
      Caption         =   $"Dummy.frx":015E
      Height          =   1005
      Left            =   165
      TabIndex        =   1
      Top             =   1035
      Width           =   3105
   End
   Begin VB.Label Label1 
      Caption         =   $"Dummy.frx":0237
      Height          =   795
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   3105
   End
End
Attribute VB_Name = "fDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then Unload Me
End Sub

Private Sub Label4_Click()
    fTest.SetFocus
End Sub
