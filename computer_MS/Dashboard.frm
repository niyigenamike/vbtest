VERSION 5.00
Begin VB.Form DashboardForm 
   BackColor       =   &H8000000D&
   Caption         =   $"Dashboard.frx":0000
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11610
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LogoutBtn 
      Caption         =   "Logout"
      Height          =   375
      Left            =   12360
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton CommandSupplier 
      Caption         =   "Supplier"
      Height          =   2295
      Left            =   9720
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton CommandOrder 
      Caption         =   "Orders"
      Height          =   2295
      Left            =   6840
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton CommandCustomer 
      Caption         =   "Customers"
      Height          =   2295
      Left            =   3360
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton CommandPoduct 
      BackColor       =   &H80000012&
      Caption         =   "Products"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "YOU ARE WELCOME TO THE GASABO  COMPUTER SHOP MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   11055
   End
End
Attribute VB_Name = "DashboardForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandCustomer_Click()
CustomerForm.Show
End Sub

Private Sub CommandOrder_Click()
OrderForm.Show
End Sub

Private Sub CommandPoduct_Click()
ProductForm.Show
End Sub

Private Sub CommandSupplier_Click()
SupplierForm.Show
End Sub

Private Sub CommandTransaction_Click()
TransactionForm.Show
End Sub


Private Sub LogoutBtn_Click()
    ' Ask the user if they want to log out
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you really want to log out?", vbYesNo + vbQuestion, "Logout Confirmation")

    ' Check the user's response
    If response = vbYes Then
        ' Show the Welcome form if the user clicked Yes
        WelcomeForm.Show
        ' Optionally, hide or close the current form
        Me.Hide ' or Me.Close if you want to completely close the current form
    Else
        ' If the user clicked No, do nothing or display a message
        MsgBox "You are still logged in.", vbInformation, "Logout Cancelled"
    End If
End Sub

