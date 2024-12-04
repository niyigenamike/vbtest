VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LoginForm 
   BackColor       =   &H8000000D&
   Caption         =   "LOGIN FORM"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1200
      Top             =   4320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Login Form.frx":0000
      OLEDBString     =   $"Login Form.frx":00A3
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "user"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton RegisterBtn 
      Caption         =   "Register"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton loginbtn 
      Caption         =   "Login"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox passwrd 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox uname 
      BackColor       =   &H80000007&
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   15
      Left            =   480
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Don't have an account"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub loginbtn_Click()
    ' Declare variables for the input fields
    Dim username As String, password As String

    ' Retrieve input values from textboxes
    username = uname.Text
    password = passwrd.Text

    ' Validate if both fields are filled
    If uname.Text = "" Or passwrd.Text = "" Then
        MsgBox "Please enter both username and password.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Database connection string
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\MIKE\Downloads\computer_MS\computer_MS\database\computer_ms_db_production.mdb;Persist Security Info=False;"
    conn.Open

    ' Create the SQL command to check for username and password
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "SELECT * FROM [user] WHERE username = ? AND [password] = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("username", 200, 1, 50, username)
        .Parameters.Append .CreateParameter("password", 200, 1, 50, password)
        
        ' Execute the command and get the recordset
        Set rs = .Execute
    End With

    ' Check if any records were returned (i.e., the user exists with the provided credentials)
    If Not rs.EOF Then
        MsgBox "Login successful!", vbInformation, "Login"
        
        ' Open the Dashboard form (Replace "Dashboard" with your actual form name)
        DashboardForm.Show
        
        ' Optionally, close the login form after successful login
       

    Else
        MsgBox "Invalid username or password. Please try again.", vbCritical, "Login Failed"
    End If

    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

    If Not consn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Private Sub CancelBtn_Click()
    ' Clear the username and password fields
    uname.Text = ""
    passwrd.Text = ""

    ' Optionally, you can reset the focus to the username field
    uname.SetFocus
End Sub

Private Sub RegisterBtn_Click()
    ' Open the registration form when Register button is clicked'
    RegisterForm.Show
    ' Optionally, you can close the login form at this point if desired
    ' DoCmd.Close acForm, Me.Name
End Sub

