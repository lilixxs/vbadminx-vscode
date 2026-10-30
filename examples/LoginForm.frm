VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Sign in"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label1
      Caption         =   "User name"
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1200
   End
   Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} TextBox1
      Height          =   300
      Left            =   1500
      TabIndex        =   1
      Top             =   270
      Width           =   2700
   End
   Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} CheckBox1
      Caption         =   "Remember me"
      Height          =   300
      Left            =   1500
      TabIndex        =   2
      Top             =   780
      Width           =   2400
   End
   Begin {6E182020-F460-11CE-9BCD-00AA00608E01} Frame1
      Caption         =   "Options"
      Height          =   900
      Left            =   240
      TabIndex        =   3
      Top             =   1260
      Width           =   3960
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} OptionButton1
         Caption         =   "Local database"
         Height          =   240
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2200
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} OptionButton2
         Caption         =   "Remote server"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   540
         Width           =   2200
      End
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} CommandButton1
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1500
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} CommandButton2
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    If TextBox1.Text = vbNullString Then
        MsgBox "Enter a user name.", vbExclamation, Caption
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub
