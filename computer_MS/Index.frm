VERSION 5.00
Begin VB.Form WelcomeForm 
   BackColor       =   &H8000000D&
   Caption         =   "Welcome "
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton lognbtn 
      BackColor       =   &H8000000D&
      Caption         =   "Login"
      Height          =   615
      Left            =   6840
      MaskColor       =   &H0000FF00&
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton regbtn 
      Caption         =   "Register"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "WELCOME TO ONLINE HARDWARE AND CONSTRUCTION INVENTORY SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1095
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "WelcomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lognbtn_Click()
LoginForm.Show ' Replace LoginForm with the actual name of your login form
End Sub

 Private Sub regbtn_Click()
    RegisterForm.Show ' Replace RegisterForm with the actual name of your registration form
End Sub






